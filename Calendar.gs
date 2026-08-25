/**
 * Fetches calendar events from all user calendars.
 * Uses CalendarApp for reliable event fetching, then enhances with
 * conference links via Calendar API v3 REST calls.
 *
 * This is the single most expensive operation in the app — it walks every
 * calendar twice — so the result is cached briefly and shared between the
 * calendar widget and the AI context.
 *
 * @param {number} daysAhead - Number of days to look ahead (default: 7)
 * @param {boolean} forceRefresh - Skip the cache and refetch
 * @returns {Array} Array of event objects sorted by start time
 */
function getCalendarEvents(daysAhead, forceRefresh) {
  daysAhead = daysAhead || 7;

  const cacheKey = 'cal_events_' + daysAhead;
  if (!forceRefresh) {
    const cached = cacheGetJson(cacheKey);
    if (cached) return cached;
  }

  const events = fetchCalendarEvents(daysAhead);
  cachePutJson(cacheKey, events, CACHE_TTL_CALENDAR);
  return events;
}

/**
 * Uncached calendar fetch. Use getCalendarEvents() instead.
 */
function fetchCalendarEvents(daysAhead) {
  const now = new Date();
  const end = new Date();
  end.setDate(end.getDate() + daysAhead);
  const start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  
  const allEvents = [];
  const calendars = CalendarApp.getAllCalendars();
  
  // Build a map of conference links per calendar via REST API
  var conferenceMap = {};
  try {
    var token = ScriptApp.getOAuthToken();
    var headers = { 'Authorization': 'Bearer ' + token };
    
    calendars.forEach(function(calendar) {
      try {
        var calId = calendar.getId();
        var eventsUrl = 'https://www.googleapis.com/calendar/v3/calendars/' 
          + encodeURIComponent(calId) + '/events'
          + '?timeMin=' + start.toISOString()
          + '&timeMax=' + end.toISOString()
          + '&singleEvents=true'
          + '&orderBy=startTime'
          + '&conferenceDataVersion=1'
          + '&fields=items(id,iCalUID,hangoutLink,conferenceData)'
          + '&maxResults=250';
        
        var resp = UrlFetchApp.fetch(eventsUrl, { headers: headers, muteHttpExceptions: true });
        var data = JSON.parse(resp.getContentText());
        
        (data.items || []).forEach(function(item) {
          var link = '';
          var name = '';
          
          if (item.conferenceData && item.conferenceData.entryPoints) {
            for (var i = 0; i < item.conferenceData.entryPoints.length; i++) {
              if (item.conferenceData.entryPoints[i].entryPointType === 'video') {
                link = item.conferenceData.entryPoints[i].uri || '';
                name = (item.conferenceData.conferenceSolution && item.conferenceData.conferenceSolution.name) || '';
                break;
              }
            }
          }
          if (!link && item.hangoutLink) {
            link = item.hangoutLink;
            name = 'Google Meet';
          }
          
          if (link && item.iCalUID) {
            conferenceMap[item.iCalUID] = { link: link, name: name };
          }
        });
      } catch(e) {
        // Skip — REST API may not be available for this calendar
      }
    });
  } catch(e) {
    Logger.log('Conference data fetch skipped: ' + e.message);
  }
  
  // Fetch events via CalendarApp (reliable)
  calendars.forEach(function(calendar) {
    try {
      var events = calendar.getEvents(start, end);
      var calName = calendar.getName();
      var calColor = calendar.getColor();
      
      events.forEach(function(event) {
        var rawDesc = event.getDescription() || '';
        var plainDesc = rawDesc.replace(/<[^>]*>/g, ' ').replace(/&nbsp;/g, ' ');
        
        // Match conference link by iCalUID
        var eventId = event.getId();
        var conf = conferenceMap[eventId] || {};
        
        allEvents.push({
          title: event.getTitle(),
          start: event.getStartTime().toISOString(),
          end: event.getEndTime().toISOString(),
          location: event.getLocation() || '',
          description: plainDesc,
          htmlDescription: rawDesc,
          hangoutLink: conf.link || '',
          conferenceName: conf.name || '',
          isAllDay: event.isAllDayEvent(),
          color: calColor || '#007acc',
          calendarName: calName,
          eventId: eventId
        });
      });
    } catch (e) {
      Logger.log('Skipped calendar: ' + calendar.getName() + ' — ' + e.message);
    }
  });
  
  allEvents.sort(function(a, b) {
    return new Date(a.start) - new Date(b.start);
  });
  
  return allEvents;
}

/**
 * Creates a new event on the user's default calendar.
 * @param {Object} eventData - Event data object
 * @returns {Object} Created event info
 */
function createCalendarEvent(eventData) {
  const calendar = CalendarApp.getDefaultCalendar();
  let event;
  
  if (eventData.isAllDay) {
    const startDate = new Date(eventData.startDate);
    if (eventData.endDate) {
      const endDate = new Date(eventData.endDate);
      // Add 1 day to end date for multi-day all-day events
      endDate.setDate(endDate.getDate() + 1);
      event = calendar.createAllDayEvent(
        eventData.title,
        startDate,
        endDate,
        {
          description: eventData.description || '',
          location: eventData.location || ''
        }
      );
    } else {
      event = calendar.createAllDayEvent(
        eventData.title,
        startDate,
        {
          description: eventData.description || '',
          location: eventData.location || ''
        }
      );
    }
  } else {
    const startTime = new Date(eventData.startDate + 'T' + (eventData.startTime || '09:00'));
    const endTime = new Date((eventData.endDate || eventData.startDate) + 'T' + (eventData.endTime || '10:00'));
    
    event = calendar.createEvent(
      eventData.title,
      startTime,
      endTime,
      {
        description: eventData.description || '',
        location: eventData.location || ''
      }
    );
  }
  
  // A new event changes both the widget's data and what the assistant sees
  invalidateCalendarCache();
  invalidateAiContextCache();

  return {
    id: event.getId(),
    title: event.getTitle(),
    start: event.getStartTime().toISOString(),
    end: event.getEndTime().toISOString()
  };
}

/**
 * Keeps a task's linked all-day Calendar event in sync with its deadline.
 * Called from saveTask() on every save; decides whether to create, replace
 * or remove the event based on what actually changed.
 * @param {Object} task - The task being saved, with its new field values
 * @param {Date|string} oldDeadlineRaw - The deadline currently on the sheet row, or falsy for a new task
 * @param {string} oldEventId - The CalendarEventID currently on the sheet row, or falsy
 * @returns {string} The event ID to persist on the row, or '' when there is none
 */
function syncTaskDeadlineEvent(task, oldDeadlineRaw, oldEventId) {
  const newDeadline = task['Дедлайн'];
  const isDone = task['Статус'] === 'Виконано';

  // Done or no deadline — nothing belongs on the calendar
  if (isDone || !newDeadline) {
    if (oldEventId) deleteTaskDeadlineEvent(oldEventId);
    return '';
  }

  const oldKey = oldDeadlineRaw ? formatDateShort(oldDeadlineRaw) : '';
  const newKey = formatDateShort(newDeadline);

  // Same deadline, event already exists — nothing to do
  if (oldEventId && oldKey === newKey) {
    return oldEventId;
  }

  return upsertTaskDeadlineEvent(task['Назва'] || 'Задача', newDeadline, oldEventId);
}

/**
 * Creates the all-day event mirroring a task's deadline, replacing any
 * previous linked event for the same task first.
 * @returns {string} The new event's ID
 */
function upsertTaskDeadlineEvent(title, deadlineValue, existingEventId) {
  if (existingEventId) deleteTaskDeadlineEvent(existingEventId);

  const calendar = CalendarApp.getDefaultCalendar();
  const event = calendar.createAllDayEvent('📋 ' + title, new Date(deadlineValue), {
    description: 'Дедлайн задачі TaskApp. Ця подія оновлюється автоматично разом із задачею.'
  });

  invalidateCalendarCache();
  return event.getId();
}

/**
 * Deletes a task's linked calendar event, tolerating an ID that no longer
 * resolves to anything (e.g. removed by hand in Calendar).
 */
function deleteTaskDeadlineEvent(eventId) {
  if (!eventId) return;
  try {
    const event = CalendarApp.getEventById(eventId);
    if (event) event.deleteEvent();
  } catch (e) {
    Logger.log('Не вдалося видалити подію дедлайну: ' + e.message);
  }
  invalidateCalendarCache();
}
