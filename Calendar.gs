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
          
          // Індексуємо за обома формами ID. Для одноразових подій iCalUID
          // збігається з CalendarApp.getId(), а для примірників повторюваної
          // події — ні, і кнопка Join зникала саме на регулярних зустрічах.
          if (link) {
            var conference = { link: link, name: name };
            if (item.iCalUID) conferenceMap[item.iCalUID] = conference;
            if (item.id) conferenceMap[item.id + '@google.com'] = conference;
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
 * Turns a form or assistant payload into real start/end Date objects.
 *
 * An absent end time means one hour after the start. The prompt promises the
 * assistant exactly that, but the code used to substitute a fixed 10:00 — so
 * "зустріч завтра о 15:00" produced an event ending before it began. Since
 * createEvent applies without a confirmation card, nobody saw it happen.
 * @param {Object} eventData - Needs startDate; startTime/endDate/endTime optional
 * @returns {Object} {start, end}
 */
function resolveEventTimes(eventData) {
  const startDate = String(eventData.startDate || '');
  const startTime = String(eventData.startTime || '09:00');

  if (!/^\d{4}-\d{2}-\d{2}$/.test(startDate)) {
    throw new Error('Некоректна дата початку події: "' + startDate + '".');
  }
  if (!/^\d{2}:\d{2}$/.test(startTime)) {
    throw new Error('Некоректний час початку події: "' + startTime + '".');
  }

  const start = new Date(startDate + 'T' + startTime);
  if (isNaN(start.getTime())) {
    throw new Error('Не вдалося розібрати початок події.');
  }

  const endDate = eventData.endDate ? String(eventData.endDate) : startDate;
  let end;

  if (eventData.endTime) {
    const endTime = String(eventData.endTime);
    if (!/^\d{2}:\d{2}$/.test(endTime)) {
      throw new Error('Некоректний час завершення події: "' + endTime + '".');
    }
    end = new Date(endDate + 'T' + endTime);
  } else {
    end = new Date(start.getTime() + 60 * 60 * 1000);
  }

  // Кінець не пізніше початку — це завжди помилка вводу, а не намір. Година від
  // початку корисніша за виняток, бо подія вже створюється без підтвердження.
  if (isNaN(end.getTime()) || end.getTime() <= start.getTime()) {
    end = new Date(start.getTime() + 60 * 60 * 1000);
  }

  return { start: start, end: end };
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
    const times = resolveEventTimes(eventData);

    event = calendar.createEvent(
      eventData.title,
      times.start,
      times.end,
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
 * Reduces a deadline to a plain yyyy-MM-dd key that two different
 * representations of the same day agree on.
 *
 * The sheet hands back a Date at local midnight, while the client sends
 * "2026-09-10". Running both through toISOString() shifts the Date a day back
 * east of UTC, so the two never matched and every comparison looked like a
 * change. A bare date string is already timezone-free and is taken as written;
 * anything with a time is read in the script's own timezone.
 * @param {Date|string} value
 * @returns {string} yyyy-MM-dd, or '' when there is nothing to compare
 */
function toCalendarDateKey(value) {
  if (!value) return '';

  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value.trim())) {
    return value.trim();
  }

  const date = (value instanceof Date) ? value : new Date(value);
  if (isNaN(date.getTime())) return '';

  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
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

  const oldKey = toCalendarDateKey(oldDeadlineRaw);
  const newKey = toCalendarDateKey(newDeadline);

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
// Prefixed onto every auto-synced deadline event's description so the client
// can tell these apart from real meetings without guessing from the title.
const TASK_DEADLINE_EVENT_MARKER = '[TaskApp:deadline-sync]';

function upsertTaskDeadlineEvent(title, deadlineValue, existingEventId) {
  if (existingEventId) deleteTaskDeadlineEvent(existingEventId);

  const calendar = CalendarApp.getDefaultCalendar();
  const event = calendar.createAllDayEvent('📋 ' + title, new Date(deadlineValue), {
    description: TASK_DEADLINE_EVENT_MARKER + ' Дедлайн задачі TaskApp. Ця подія оновлюється автоматично разом із задачею.'
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

/**
 * Updates an existing calendar event's title/time/location/description.
 * If the requested isAllDay differs from what the event currently is,
 * CalendarApp has no in-place way to convert one into the other — so the
 * event is deleted and recreated instead, and the (new) event's info is
 * returned so the client can pick up its new ID.
 * @param {string} eventId
 * @param {Object} eventData - Same shape as createCalendarEvent() expects
 * @returns {Object} {id, title, start, end}
 */
function updateCalendarEvent(eventId, eventData) {
  const event = CalendarApp.getEventById(eventId);
  if (!event) {
    throw new Error('Подію не знайдено — можливо, її вже видалено.');
  }

  if (event.isAllDayEvent() !== !!eventData.isAllDay) {
    event.deleteEvent();
    invalidateCalendarCache();
    return createCalendarEvent(eventData);
  }

  event.setTitle(eventData.title);
  event.setLocation(eventData.location || '');
  event.setDescription(eventData.description || '');

  if (eventData.isAllDay) {
    const startDate = new Date(eventData.startDate);
    if (eventData.endDate && eventData.endDate !== eventData.startDate) {
      const endDate = new Date(eventData.endDate);
      endDate.setDate(endDate.getDate() + 1);
      event.setAllDayDates(startDate, endDate);
    } else {
      event.setAllDayDate(startDate);
    }
  } else {
    const times = resolveEventTimes(eventData);
    event.setTime(times.start, times.end);
  }

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
 * Calls the Calendar REST API with the script's OAuth token. CalendarApp can
 * neither create conferences nor invite guests with a notification, so those
 * two operations go through REST instead.
 * @param {string} path - Path and query after /calendars/primary/events
 * @param {string} method - 'get' or 'patch'
 * @param {Object} payload - Request body for patch
 * @returns {Object} Parsed API response
 */
function calendarApiFetch(path, method, payload) {
  const options = {
    method: method,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    contentType: 'application/json',
    muteHttpExceptions: true
  };
  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  const url = 'https://www.googleapis.com/calendar/v3/calendars/primary/events' + path;
  const response = UrlFetchApp.fetch(url, options);
  const body = JSON.parse(response.getContentText() || '{}');

  const code = response.getResponseCode();
  if (code === 404) {
    throw new Error('Подію не знайдено в основному календарі — Meet і запрошення працюють лише для подій у ньому.');
  }
  if (code >= 400) {
    throw new Error((body.error && body.error.message) || 'Calendar API помилка');
  }
  return body;
}

/**
 * CalendarApp IDs carry an "@google.com" suffix that the REST API rejects.
 */
function toRestEventId(eventId) {
  return String(eventId).replace(/@google\.com$/, '');
}

/**
 * Pulls the video-conference URL out of an events.get/patch response.
 */
function extractMeetLink(event) {
  const entryPoints = (event.conferenceData && event.conferenceData.entryPoints) || [];
  for (var i = 0; i < entryPoints.length; i++) {
    if (entryPoints[i].entryPointType === 'video' && entryPoints[i].uri) {
      return entryPoints[i].uri;
    }
  }
  return event.hangoutLink || '';
}

/**
 * Attaches a Google Meet conference to an existing event, so the link shows up
 * as a real "Join" button in Calendar and in guests' invitations. Returns the
 * conference already on the event if there is one, rather than creating a second.
 * @param {string} eventId
 * @returns {Object} {link}
 */
function createMeetForEvent(eventId) {
  const restId = toRestEventId(eventId);
  const existing = calendarApiFetch('/' + encodeURIComponent(restId) + '?conferenceDataVersion=1', 'get');

  const existingLink = extractMeetLink(existing);
  if (existingLink) {
    return { link: existingLink };
  }

  const updated = calendarApiFetch(
    '/' + encodeURIComponent(restId) + '?conferenceDataVersion=1',
    'patch',
    {
      conferenceData: {
        createRequest: {
          requestId: Utilities.getUuid(),
          conferenceSolutionKey: { type: 'hangoutsMeet' }
        }
      }
    }
  );

  var link = extractMeetLink(updated);

  // Google provisions the conference asynchronously; on a "pending" create
  // request the link is only there on a follow-up read.
  if (!link) {
    Utilities.sleep(1500);
    const refetched = calendarApiFetch('/' + encodeURIComponent(restId) + '?conferenceDataVersion=1', 'get');
    link = extractMeetLink(refetched);
  }

  if (!link) {
    throw new Error('Google Meet не створився — спробуйте ще раз за кілька секунд.');
  }

  invalidateCalendarCache();
  invalidateAiContextCache();

  return { link: link };
}

/**
 * Writes the conference URL into the event's location, so it is copyable from
 * the "Місце" field instead of only living in conferenceData. The edit modal
 * does this client-side for manual clicks, where the user can still back out
 * before saving; this is for events created without a modal to review.
 * @param {string} eventId
 * @param {string} link
 */
function appendMeetLinkToLocation(eventId, link) {
  const event = CalendarApp.getEventById(eventId);
  if (!event || !link) return;

  const existing = (event.getLocation() || '').trim();
  if (existing.indexOf(link) !== -1) return;

  event.setLocation(existing ? existing + ' · ' + link : link);
  invalidateCalendarCache();
  invalidateAiContextCache();
}

/**
 * Lists who is already invited to an event, for the edit modal.
 * @param {string} eventId
 * @returns {Object} {attendees}
 */
function getEventGuests(eventId) {
  const event = calendarApiFetch('/' + encodeURIComponent(toRestEventId(eventId)), 'get');
  return {
    attendees: (event.attendees || []).map(function (attendee) {
      return { email: attendee.email, status: attendee.responseStatus || 'needsAction' };
    })
  };
}

/**
 * Adds guests to an event and emails them the invitation. Guests already on the
 * event are kept — a patch with an attendees array replaces the whole list.
 * @param {string} eventId
 * @param {string[]} emails
 * @returns {Object} {attendees} - Every guest on the event after the update
 */
function inviteEventAttendees(eventId, emails) {
  const wanted = (emails || [])
    .map(function (email) { return String(email).trim().toLowerCase(); })
    .filter(function (email) { return email.indexOf('@') > 0; });

  if (!wanted.length) {
    throw new Error('Не вказано жодної коректної email-адреси.');
  }

  const restId = toRestEventId(eventId);
  const event = calendarApiFetch('/' + encodeURIComponent(restId), 'get');

  const merged = (event.attendees || []).slice();
  const seen = {};
  merged.forEach(function (attendee) {
    seen[String(attendee.email).toLowerCase()] = true;
  });
  wanted.forEach(function (email) {
    if (!seen[email]) {
      merged.push({ email: email });
      seen[email] = true;
    }
  });

  const updated = calendarApiFetch(
    '/' + encodeURIComponent(restId) + '?sendUpdates=all',
    'patch',
    { attendees: merged }
  );

  invalidateCalendarCache();
  invalidateAiContextCache();

  return {
    attendees: (updated.attendees || []).map(function (attendee) {
      return { email: attendee.email, status: attendee.responseStatus || 'needsAction' };
    })
  };
}

/**
 * Deletes a calendar event by ID, as an explicit user action (unlike
 * deleteTaskDeadlineEvent, failures here are not swallowed — the user
 * clicked Delete and should see if it didn't work).
 */
function deleteCalendarEvent(eventId) {
  const event = CalendarApp.getEventById(eventId);
  if (event) {
    event.deleteEvent();
  }
  invalidateCalendarCache();
  invalidateAiContextCache();
  return true;
}
