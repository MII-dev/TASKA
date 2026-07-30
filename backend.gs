const SPREADSHEET_ID = '1djbxaGjw1qIpFBZb2P14noEY0iErch1hSOBBxlcXP5k';
const SHEET_NAME = 'Tasks';
const PROJECTS_SHEET_NAME = 'Projects';
const DIRECTORY_SPREADSHEET_ID = '1xbaVZaeKf_J_85vJJKpVBNJSu51-HdWo0RVLBTzjo_g';

// ==========================================
// ============ Cache helpers ===============
// ==========================================

// CacheService rejects values over 100 KB per key. Staying under that with a
// margin is cheaper than letting put() throw and break the calling feature.
const CACHE_MAX_VALUE_BYTES = 90000;

const CACHE_TTL_CALENDAR = 180;
const CACHE_TTL_AI_CONTEXT = 60;
const CACHE_TTL_CLICKUP_TASKS = 120;

// Every cache key the AI context is assembled from, so invalidation can be blanket
const AI_CONTEXT_CACHE_KEYS = ['ai_ctx_full', 'ai_ctx_tasks'];

/**
 * Reads a JSON value from the user cache. Returns null on a miss or on a
 * corrupted entry — callers always have a way to recompute.
 */
function cacheGetJson(key) {
  try {
    const raw = CacheService.getUserCache().get(key);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    Logger.log('Кеш: не вдалося прочитати ' + key + ': ' + e.message);
    return null;
  }
}

/**
 * Writes a JSON value to the user cache, silently skipping values that exceed
 * the per-key limit. Caching is an optimization — never let it break a feature.
 */
function cachePutJson(key, value, ttlSeconds) {
  try {
    const serialized = JSON.stringify(value);
    if (serialized.length > CACHE_MAX_VALUE_BYTES) {
      Logger.log('Кеш: ' + key + ' завеликий (' + serialized.length + ' байт), не кешую');
      return;
    }
    CacheService.getUserCache().put(key, serialized, ttlSeconds);
  } catch (e) {
    Logger.log('Кеш: не вдалося записати ' + key + ': ' + e.message);
  }
}

/**
 * Drops the cached AI context. Must run after anything that changes what the
 * assistant would see, otherwise it keeps answering from stale data for a minute.
 */
function invalidateAiContextCache() {
  try {
    CacheService.getUserCache().removeAll(AI_CONTEXT_CACHE_KEYS);
  } catch (e) {
    Logger.log('Кеш: не вдалося скинути контекст ШІ: ' + e.message);
  }
}

/**
 * Drops cached calendar events for every look-ahead window the app uses.
 */
function invalidateCalendarCache() {
  try {
    CacheService.getUserCache().removeAll(['cal_events_1', 'cal_events_7']);
  } catch (e) {
    Logger.log('Кеш: не вдалося скинути календар: ' + e.message);
  }
}

/**
 * Serves the HTML file.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('TaskApp - Управління задачами')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Initializes the spreadsheet with headers if it's empty.
 */
function initSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  
  if (sheet.getLastRow() === 0) {
    const headers = [
      'ID', 'Дата створення', 'Тип', 'Назва', 'Опис', 
      'Замовник', 'Виконавець', 'Дедлайн', 'Статус', 
      'Кроки', 'Посилання на результат', 'Коментарі', 'Дата виконання', 'Пріоритет'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Fetches all tasks from the spreadsheet.
 */
function getTasks() {
  const sheet = initSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      let val = row[i];
      // Convert dates to ISO strings for safer serialization
      if (val instanceof Date) {
        val = val.toISOString();
      }
      obj[header] = val;
    });
    // Parse steps JSON if present
    if (obj['Кроки']) {
      try {
        const stepsVal = obj['Кроки'];
        obj['Кроки'] = typeof stepsVal === 'string' && stepsVal ? JSON.parse(stepsVal) : (Array.isArray(stepsVal) ? stepsVal : []);
      } catch (e) {
        obj['Кроки'] = [];
      }
    } else {
      obj['Кроки'] = [];
    }
    return obj;
  });
}

/**
 * Saves a task (creates new or updates existing).
 */
function saveTask(task) {
  const sheet = initSheet();
  const data = sheet.getDataRange().getValues();
  let headers = data[0];
  
  if (!headers.includes('Дата виконання')) {
    headers.push('Дата виконання');
    sheet.getRange(1, headers.length).setValue('Дата виконання');
  }
  
  if (!headers.includes('Пріоритет')) {
    headers.push('Пріоритет');
    sheet.getRange(1, headers.length).setValue('Пріоритет');
  }

  if (!headers.includes('Проєкт')) {
    headers.push('Проєкт');
    sheet.getRange(1, headers.length).setValue('Проєкт');
  }

  if (!headers.includes('Гілка')) {
    headers.push('Гілка');
    sheet.getRange(1, headers.length).setValue('Гілка');
  }

  // Ensure steps are stringified for storage
  const taskToSave = {...task};
  if (taskToSave['Кроки'] && typeof taskToSave['Кроки'] !== 'string') {
    taskToSave['Кроки'] = JSON.stringify(taskToSave['Кроки']);
  }
  
  let rowIndex = -1;
  if (task.ID) {
    // Find existing task by ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == task.ID) {
        rowIndex = i + 1;
        break;
      }
    }
  } else {
    // Generate new ID
    taskToSave.ID = Utilities.getUuid();
    taskToSave['Дата створення'] = new Date();
  }
  
  const rowValues = headers.map(header => taskToSave[header] || '');
  
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  invalidateAiContextCache();
  return taskToSave.ID;
}

/**
 * Deletes a task by ID.
 */
function deleteTask(id) {
  const sheet = initSheet();
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      invalidateAiContextCache();
      return true;
    }
  }
  return false;
}

// ==========================================
// ============== Projects ==================
// ==========================================

const PROJECT_JSON_FIELDS = ['Гілки', 'Посилання', 'Контакти', 'Дедлайни'];

/**
 * Initializes the Projects sheet with headers if it's empty.
 */
function initProjectsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PROJECTS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(PROJECTS_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    const headers = [
      'ID', 'Дата створення', 'Назва', 'Опис', 'Статус',
      'Гілки', 'Посилання', 'Контакти', 'Дедлайни'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Fetches all projects from the spreadsheet, parsing the JSON-in-cell
 * columns (Гілки/Посилання/Контакти/Дедлайни) the same way getTasks()
 * parses 'Кроки' — corrupted or missing values default to [].
 */
function getProjects() {
  const sheet = initProjectsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((header, i) => {
      let val = row[i];
      if (val instanceof Date) {
        val = val.toISOString();
      }
      obj[header] = val;
    });

    PROJECT_JSON_FIELDS.forEach(field => {
      if (obj[field]) {
        try {
          const fieldVal = obj[field];
          obj[field] = typeof fieldVal === 'string' && fieldVal ? JSON.parse(fieldVal) : (Array.isArray(fieldVal) ? fieldVal : []);
        } catch (e) {
          obj[field] = [];
        }
      } else {
        obj[field] = [];
      }
    });

    return obj;
  });
}

/**
 * Saves a project (creates new or updates existing).
 */
function saveProject(project) {
  const sheet = initProjectsSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const projectToSave = {...project};
  PROJECT_JSON_FIELDS.forEach(field => {
    if (projectToSave[field] && typeof projectToSave[field] !== 'string') {
      projectToSave[field] = JSON.stringify(projectToSave[field]);
    }
  });

  let rowIndex = -1;
  if (project.ID) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == project.ID) {
        rowIndex = i + 1;
        break;
      }
    }
  } else {
    projectToSave.ID = Utilities.getUuid();
    projectToSave['Дата створення'] = new Date();
  }

  const rowValues = headers.map(header => projectToSave[header] || '');

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  invalidateAiContextCache();
  return projectToSave.ID;
}

/**
 * Deletes a project by ID. Tasks that reference it are left with a
 * dangling Проєкт/Гілка ID — the UI simply shows no badge for them.
 */
function deleteProject(id) {
  const sheet = initProjectsSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      invalidateAiContextCache();
      return true;
    }
  }
  return false;
}

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

// ==========================================
// ====== AI Integration (Gemini API) =======
// ==========================================

const AI_CHAT_SYSTEM_INSTRUCTION = `
Ти — ШІ-асистент для додатку TaskApp (система керування задачами з інтегрованими Google Sheets, Google Calendar та ClickUp).
Твоє головне завдання — допомагати користувачу керувати задачами та календарем, надавати аналітику та виконувати дії.

Користувач спілкується з тобою українською мовою. Твої відповіді мають бути чіткими, дружніми та професійними, українською мовою.

У тебе є доступ до поточного контексту користувача:
1. Його поточний час та дата.
2. Поточні задачі TaskApp (активні та нещодавно виконані) — ними ти можеш керувати.
3. Поточні події у календарі на 7 днів.
4. Задачі з ClickUp — якщо інтеграцію підключено. Їх ти бачиш, але ЗМІНЮВАТИ їх не можеш.
   Враховуй їх, коли оцінюєш завантаження чи радиш, за що братися. Якщо користувач просить
   змінити ClickUp-задачу — поясни, що це робиться у вкладці ClickUp вручну.

Зверни увагу: список задач у контексті може бути обрізаний (про це буде відповідна позначка).
Якщо потрібної задачі не видно — попроси користувача уточнити назву, а не вигадуй ID.

Ти можеш не просто відповідати текстом, а й ініціювати дії у системі.
Ти ЗАВЖДИ повертаєш відповідь у форматі JSON:
{
  "text": "Твоє текстове повідомлення користувачу українською мовою. Використовуй markdown для списків чи важливих слів.",
  "actions": [ { "action": "...", "data": { ... } } ]
}

"actions" — масив. Якщо дій немає, поверни порожній масив або не додавай поле взагалі.
Якщо користувач попросив кілька речей за раз — поверни кілька елементів у масиві.

Доступні дії та формат "data":

1. "createTask" — створити задачу в TaskApp.
   Приклад запиту: "Додай задачу підготувати звіт до п'ятниці".
   data:
     - "Назва" (обов'язково): короткий та чіткий заголовок задачі.
     - "Опис" (опціонально): детальний опис.
     - "Дедлайн" (опціонально): дата у форматі YYYY-MM-DD. Відносні дати ("до п'ятниці") вирахуй від поточної дати.
     - "Пріоритет" (опціонально): "Низький" | "Середній" | "Високий".
     - "Тип" (опціонально): категорія/тег задачі.
     - "Виконавець" (опціонально): ім'я виконавця (якщо не вказано, залиш порожнім).

2. "updateTask" — змінити наявну задачу TaskApp.
   data:
     - "ID" (обов'язково): ID задачі з контексту.
     - Поля для зміни: "Назва", "Опис", "Статус" ("Нова" | "В роботі" | "На перевірці" | "Виконано"),
       "Пріоритет" ("Низький" | "Середній" | "Високий"), "Тип", "Дедлайн", "Коментарі",
       "Посилання на результат", "Виконавець".

3. "deleteTask" — видалити задачу TaskApp. Пропонуй лише коли користувач просить про це прямо.
   data: { "ID": "..." }

4. "createEvent" — створити подію в Google Calendar.
   data:
     - "title" (обов'язково)
     - "startDate" (обов'язково): YYYY-MM-DD
     - "startTime" (опціонально): HH:MM, за замовчуванням 09:00
     - "endDate" (опціонально): YYYY-MM-DD, за замовчуванням дорівнює startDate
     - "endTime" (опціонально): HH:MM, за замовчуванням +1 година від початку
     - "isAllDay" (опціонально): true для події на весь день
     - "location" (опціонально), "description" (опціонально)

5. "decomposeTask" — розбити наявну задачу TaskApp на кроки.
   data: { "ID": "..." }

Пам'ятай:
- Дії НЕ виконуються одразу — користувач побачить картку і має підтвердити її вручну.
  Тому у полі "text" пиши в майбутньому часі: "Підготував задачу — перевір і підтверди",
  а не "Задачу створено".
- Завжди звіряй ID задач із контекстом перед тим, як їх оновлювати чи видаляти.
- Будь уважним з датами. Поточна дата надається в контексті.
- Повертай ТІЛЬКИ валідний JSON-об'єкт. Ніякого зайвого тексту поза JSON.
`;

/**
 * Retrieves the Gemini API key from Script Properties.
 * Returns null when it has not been configured — there is deliberately no
 * fallback key in the source.
 */
function getApiKey() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  return (key && key.trim()) ? key.trim() : null;
}

/**
 * Saves the Gemini API key to Script Properties, but only if it actually works.
 * Validates against the live API and leaves the previous key untouched on failure.
 */
function setApiKey(key) {
  const props = PropertiesService.getScriptProperties();

  if (!key || key.trim() === '') {
    props.deleteProperty('GEMINI_API_KEY');
    return true;
  }

  const previousKey = props.getProperty('GEMINI_API_KEY');
  props.setProperty('GEMINI_API_KEY', key.trim());

  try {
    // Cheapest possible round-trip that still proves the key is accepted
    callGemini({ prompt: 'ping', retries: 0 });
    return true;
  } catch (e) {
    if (previousKey) {
      props.setProperty('GEMINI_API_KEY', previousKey);
    } else {
      props.deleteProperty('GEMINI_API_KEY');
    }
    throw new Error('Ключ не пройшов перевірку. ' + e.message);
  }
}

/**
 * Checks if the API key is configured.
 */
function getAiStatus() {
  return {
    hasKey: !!getApiKey()
  };
}

const GEMINI_MODEL = 'gemini-3.1-flash-lite';
const GEMINI_API_BASE = 'https://generativelanguage.googleapis.com/v1beta/models/';
const GEMINI_RETRY_CODES = [429, 500, 502, 503, 504];
const GEMINI_MAX_RETRIES = 2; // 3 attempts total

/**
 * Turns a Gemini HTTP failure into something a user can act on, instead of
 * dumping the raw error JSON into an alert().
 */
function describeGeminiError(code, rawText) {
  var detail = '';
  try {
    var parsed = JSON.parse(rawText);
    if (parsed.error && parsed.error.message) detail = ' (' + parsed.error.message + ')';
  } catch (e) {
    // Not JSON — ignore, the code alone is informative enough
  }

  if (code === 400) return 'Запит відхилено Gemini — можливо, ключ має неправильний формат.' + detail;
  if (code === 401 || code === 403) return 'Ключ Gemini недійсний або не має доступу до моделі ' + GEMINI_MODEL + '.' + detail;
  if (code === 404) return 'Модель ' + GEMINI_MODEL + ' недоступна для цього ключа.' + detail;
  if (code === 429) return 'Перевищено ліміт запитів до Gemini. Спробуйте за хвилину.';
  if (code >= 500) return 'Сервіс Gemini тимчасово перевантажений. Спробуйте ще раз.';
  return 'Помилка Gemini API (HTTP ' + code + ')' + detail;
}

/**
 * Single entry point for every Gemini call in the app.
 *
 * @param {Object} options
 *   - prompt {string}            single-turn text prompt
 *   - chatHistory {Array}        multi-turn history of {role, text}; used instead of prompt
 *   - systemInstruction {string} system prompt
 *   - json {boolean}             ask the model for application/json
 *   - file {Object}              {base64, mimeType} to send inline alongside the prompt
 *   - model {string}             override the default model
 *   - temperature {number}       override sampling temperature
 *   - retries {number}           override retry count (0 disables retrying)
 * @returns {string} The model's text response
 */
function callGemini(options) {
  const opts = options || {};
  const apiKey = getApiKey();

  if (!apiKey) {
    throw new Error('Gemini API ключ не налаштовано. Додайте його в банері налаштування ШІ.');
  }

  const model = opts.model || GEMINI_MODEL;
  const url = GEMINI_API_BASE + model + ':generateContent?key=' + apiKey;

  // Build contents: chat history wins over a single prompt
  let contents;
  if (opts.chatHistory && opts.chatHistory.length > 0) {
    contents = opts.chatHistory.map(function (msg) {
      return {
        role: (msg.role === 'bot' || msg.role === 'model') ? 'model' : 'user',
        parts: [{ text: msg.text }]
      };
    });
  } else {
    const parts = [];
    if (opts.file && opts.file.base64) {
      parts.push({
        inlineData: {
          mimeType: opts.file.mimeType,
          data: opts.file.base64
        }
      });
    }
    parts.push({ text: opts.prompt || '' });
    contents = [{ parts: parts }];
  }

  const payload = { contents: contents };

  if (opts.systemInstruction) {
    payload.systemInstruction = { parts: [{ text: opts.systemInstruction }] };
  }

  const generationConfig = {};
  if (opts.json) generationConfig.responseMimeType = 'application/json';
  if (typeof opts.temperature === 'number') generationConfig.temperature = opts.temperature;
  if (Object.keys(generationConfig).length > 0) payload.generationConfig = generationConfig;

  const fetchOptions = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const maxRetries = (typeof opts.retries === 'number') ? opts.retries : GEMINI_MAX_RETRIES;
  let lastError = null;

  for (var attempt = 0; attempt <= maxRetries; attempt++) {
    if (attempt > 0) {
      // 1s, 2s, 4s — flash-lite hits rate limits often enough to be worth waiting out
      Utilities.sleep(1000 * Math.pow(2, attempt - 1));
    }

    var response = UrlFetchApp.fetch(url, fetchOptions);
    var code = response.getResponseCode();
    var text = response.getContentText();

    if (code === 200) {
      var resObj = JSON.parse(text);
      var candidate = resObj.candidates && resObj.candidates[0];

      if (candidate && candidate.content && candidate.content.parts && candidate.content.parts[0]) {
        return candidate.content.parts[0].text;
      }

      // A blocked or truncated response comes back 200 with no usable parts
      var reason = (candidate && candidate.finishReason) ? candidate.finishReason : 'невідома причина';
      throw new Error('Gemini повернув порожню відповідь (' + reason + ').');
    }

    lastError = new Error(describeGeminiError(code, text));

    if (GEMINI_RETRY_CODES.indexOf(code) === -1) {
      throw lastError; // Permanent failure — retrying will not help
    }

    Logger.log('Gemini ' + code + ', спроба ' + (attempt + 1) + '/' + (maxRetries + 1));
  }

  throw lastError;
}

/**
 * Parses a JSON response from Gemini, tolerating markdown code fences —
 * the model sometimes wraps output in ```json even with responseMimeType set.
 */
function parseGeminiJson(text, contextLabel) {
  if (!text || !text.trim()) {
    throw new Error('Gemini повернув порожню відповідь' + (contextLabel ? ' (' + contextLabel + ')' : '') + '.');
  }

  let cleaned = text.trim();

  // Strip a leading ```json / ``` fence and its closing counterpart
  const fenceMatch = cleaned.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  if (fenceMatch) {
    cleaned = fenceMatch[1].trim();
  }

  try {
    return JSON.parse(cleaned);
  } catch (e) {
    const snippet = cleaned.length > 300 ? cleaned.substring(0, 300) + '…' : cleaned;
    throw new Error(
      'Не вдалося розібрати відповідь ШІ' + (contextLabel ? ' (' + contextLabel + ')' : '') +
      '. Модель повернула не JSON: ' + snippet
    );
  }
}

/**
 * Helper to call Gemini API with standard system instruction.
 */
function callGeminiAI(prompt, systemInstruction, isJson) {
  return callGemini({ prompt: prompt, systemInstruction: systemInstruction, json: isJson });
}

/**
 * Helper to call Gemini API with chat history.
 */
function callGeminiAIChat(chatHistory, systemInstruction, isJson) {
  return callGemini({ chatHistory: chatHistory, systemInstruction: systemInstruction, json: isJson });
}

// How much of each source makes it into the prompt. The whole task list used to
// go in verbatim — including every step's text — which grew without bound.
const AI_CTX_MAX_TASKS = 30;
const AI_CTX_MAX_COMPLETED = 5;
const AI_CTX_MAX_EVENTS = 25;
const AI_CTX_MAX_CLICKUP_TASKS = 25;
const AI_CTX_DESC_LENGTH = 200;
const AI_CTX_NOTES_LENGTH = 150;
const AI_CTX_EVENT_DESC_LENGTH = 120;
const AI_CTX_MAX_PROJECTS = 15;
const AI_CTX_PROJECT_NOTES_LENGTH = 150;

/**
 * Shortens text for the prompt without cutting mid-nonsense.
 */
function truncateForContext(text, maxLength) {
  if (!text) return '';
  const clean = String(text).replace(/\s+/g, ' ').trim();
  if (clean.length <= maxLength) return clean;
  return clean.substring(0, maxLength) + '…';
}

/**
 * Formats a clean textual list of active tasks for the AI context.
 * Trimmed on purpose: only the top tasks in full, the rest rolled up into counts.
 * @param {number} limit - How many active tasks to spell out (default AI_CTX_MAX_TASKS)
 */
function getTasksContext(limit) {
  const maxTasks = limit || AI_CTX_MAX_TASKS;
  const tasks = getTasks();

  if (tasks.length === 0) {
    return 'Немає створених задач.';
  }

  const activeTasks = tasks.filter(t => t.Статус !== 'Виконано');
  const completedTasks = tasks.filter(t => t.Статус === 'Виконано');

  const pMap = { 'Високий': 1, 'Середній': 2, 'Низький': 3 };
  activeTasks.sort((a, b) => {
    const pDiff = (pMap[a.Пріоритет] || 4) - (pMap[b.Пріоритет] || 4);
    if (pDiff !== 0) return pDiff;
    if (!a.Дедлайн) return 1;
    if (!b.Дедлайн) return -1;
    return new Date(a.Дедлайн) - new Date(b.Дедлайн);
  });

  // Aggregates first — this is what "what should I do next" questions actually need,
  // and it stays accurate even for the tasks that get cut from the list below.
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let overdueCount = 0;
  let dueTodayCount = 0;
  const byPriority = { 'Високий': 0, 'Середній': 0, 'Низький': 0, 'Без пріоритету': 0 };

  activeTasks.forEach(t => {
    byPriority[t.Пріоритет in byPriority ? t.Пріоритет : 'Без пріоритету']++;
    if (!t.Дедлайн) return;
    const d = new Date(t.Дедлайн);
    if (isNaN(d)) return;
    d.setHours(0, 0, 0, 0);
    if (d < today) overdueCount++;
    else if (d.getTime() === today.getTime()) dueTodayCount++;
  });

  let context = 'ПОТОЧНИЙ СТАТУС ЗАДАЧ:\n';
  context += `Всього активних: ${activeTasks.length} | Виконаних: ${completedTasks.length}\n`;
  context += `Прострочено: ${overdueCount} | Дедлайн сьогодні: ${dueTodayCount}\n`;
  context += `За пріоритетом — Високий: ${byPriority['Високий']}, Середній: ${byPriority['Середній']}, `;
  context += `Низький: ${byPriority['Низький']}, без пріоритету: ${byPriority['Без пріоритету']}\n\n`;

  const shown = activeTasks.slice(0, maxTasks);
  context += `АКТИВНІ ЗАДАЧІ (${shown.length} з ${activeTasks.length}, від найважливіших):\n`;

  shown.forEach((t, index) => {
    context += `${index + 1}. [${t.Статус}] "${t.Назва}" | ID: ${t.ID}\n`;

    const attrs = [];
    if (t.Пріоритет) attrs.push('пріоритет: ' + t.Пріоритет);
    if (t.Дедлайн) attrs.push('дедлайн: ' + formatDateShort(t.Дедлайн));
    if (t.Виконавець) attrs.push('виконавець: ' + t.Виконавець);
    if (t.Тип) attrs.push('категорія: ' + t.Тип);
    // Step progress only — the full step text was the single biggest cost here
    if (t.Кроки && t.Кроки.length > 0) {
      attrs.push('кроки: ' + t.Кроки.filter(s => s.done).length + '/' + t.Кроки.length);
    }
    if (attrs.length > 0) context += '   ' + attrs.join(' | ') + '\n';

    if (t.Опис) context += '   Опис: ' + truncateForContext(t.Опис, AI_CTX_DESC_LENGTH) + '\n';
    if (t.Коментарі) context += '   Нотатки: ' + truncateForContext(t.Коментарі, AI_CTX_NOTES_LENGTH) + '\n';
  });

  if (activeTasks.length > shown.length) {
    context += `…та ще ${activeTasks.length - shown.length} активних задач (не показано — попроси користувача уточнити, якщо потрібні саме вони).\n`;
  }

  const recentCompleted = completedTasks
    .sort((a, b) => new Date(b['Дата виконання'] || 0) - new Date(a['Дата виконання'] || 0))
    .slice(0, AI_CTX_MAX_COMPLETED);

  if (recentCompleted.length > 0) {
    context += `\nНЕЩОДАВНО ВИКОНАНІ (останні ${recentCompleted.length}):\n`;
    recentCompleted.forEach((t, index) => {
      context += `${index + 1}. "${t.Назва}" (${formatDateShort(t['Дата виконання'])})\n`;
    });
  }

  return context;
}

function formatDateShort(dateStr) {
  if (!dateStr) return '';
  try {
    const d = new Date(dateStr);
    if (isNaN(d)) return dateStr;
    return d.toISOString().split('T')[0];
  } catch (e) {
    return dateStr;
  }
}

/**
 * Formats a clean textual list of projects (with branch progress and
 * upcoming milestones) for the AI context.
 */
function getProjectsContext() {
  const projects = getProjects();
  if (projects.length === 0) {
    return '';
  }

  const tasks = getTasks();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const shown = projects.slice(0, AI_CTX_MAX_PROJECTS);
  let context = `ПРОЄКТИ (${shown.length} з ${projects.length}):\n`;

  shown.forEach((p, index) => {
    const projectTasks = tasks.filter(t => t['Проєкт'] === p.ID);
    const activeCount = projectTasks.filter(t => t.Статус !== 'Виконано').length;
    const overdueCount = projectTasks.filter(t => {
      if (!t.Дедлайн || t.Статус === 'Виконано') return false;
      const d = new Date(t.Дедлайн);
      if (isNaN(d)) return false;
      d.setHours(0, 0, 0, 0);
      return d < today;
    }).length;

    context += `${index + 1}. [${p.Статус || 'Без статусу'}] "${p.Назва}" | ID: ${p.ID}\n`;
    context += `   Активних задач: ${activeCount} | Прострочено: ${overdueCount}\n`;

    const branches = Array.isArray(p.Гілки) ? p.Гілки : [];
    if (branches.length > 0) {
      const branchSummary = branches.map(b => {
        const branchTasks = projectTasks.filter(t => t['Гілка'] === b.id);
        const done = branchTasks.filter(t => t.Статус === 'Виконано').length;
        return `${b.name} ${done}/${branchTasks.length}`;
      }).join(', ');
      context += `   Гілки: ${branchSummary}\n`;
    }

    const milestones = Array.isArray(p.Дедлайни) ? p.Дедлайни : [];
    const nextMilestone = milestones
      .filter(m => !m.done && m.date)
      .sort((a, b) => new Date(a.date) - new Date(b.date))[0];
    if (nextMilestone) {
      context += `   Найближча віха: "${nextMilestone.title}" (${formatDateShort(nextMilestone.date)})\n`;
    }

    if (p.Опис) context += '   Нотатки: ' + truncateForContext(p.Опис, AI_CTX_PROJECT_NOTES_LENGTH) + '\n';
  });

  if (projects.length > shown.length) {
    context += `…та ще ${projects.length - shown.length} проєктів.\n`;
  }

  return context;
}

/**
 * Formats a clean textual list of calendar events for the AI context.
 */
function getCalendarContext() {
  try {
    const events = getCalendarEvents(7);
    if (events.length === 0) {
      return 'Календар порожній на найближчі 7 днів.';
    }

    const shown = events.slice(0, AI_CTX_MAX_EVENTS);
    let context = `РОЗКЛАД НА НАЙБЛИЖЧІ 7 ДНІВ (${shown.length} з ${events.length}):\n`;

    shown.forEach(ev => {
      const startD = new Date(ev.start);
      const dateStr = startD.toLocaleDateString('uk-UA', { weekday: 'short', day: 'numeric', month: 'short' });

      let timeStr;
      if (ev.isAllDay) {
        timeStr = 'Весь день';
      } else {
        const startT = startD.toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit' });
        const endT = new Date(ev.end).toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit' });
        timeStr = startT + ' - ' + endT;
      }

      context += `- ${dateStr} [${timeStr}]: "${ev.title}"`;
      if (ev.location) context += ' | Місце: ' + truncateForContext(ev.location, 80);
      if (ev.description && ev.description.trim()) {
        context += ' | Опис: ' + truncateForContext(ev.description, AI_CTX_EVENT_DESC_LENGTH);
      }
      context += '\n';
    });

    if (events.length > shown.length) {
      context += `…та ще ${events.length - shown.length} подій.\n`;
    }

    return context;
  } catch (e) {
    return `Не вдалося отримати події календаря: ${e.message}`;
  }
}

/**
 * Assembles the full reference context handed to the assistant: current time,
 * tasks, calendar and ClickUp.
 *
 * Rebuilding this used to happen on every single AI call and dominated the
 * response time, so the assembled string is cached briefly. Anything that
 * changes the underlying data must call invalidateAiContextCache().
 *
 * @param {boolean} forceRefresh - Skip the cache
 * @returns {string} The reference context block
 */
function getAiContext(forceRefresh) {
  if (!forceRefresh) {
    const cached = cacheGetJson('ai_ctx_full');
    if (cached) {
      // The clock is the one part that must never come from cache
      return buildAiDateContext() + cached;
    }
  }

  const parts = [getTasksContext()];

  const projectsContext = getProjectsContext();
  if (projectsContext) parts.push(projectsContext);

  parts.push(getCalendarContext());

  const clickUpContext = getClickUpContext();
  if (clickUpContext) parts.push(clickUpContext);

  const body = parts.join('\n');
  cachePutJson('ai_ctx_full', body, CACHE_TTL_AI_CONTEXT);

  return buildAiDateContext() + body;
}

/**
 * The current date/time header — always computed fresh.
 */
function buildAiDateContext() {
  const now = new Date();
  return `Поточний час: ${now.toLocaleTimeString('uk-UA')}, ` +
    `Поточна дата: ${now.toLocaleDateString('uk-UA')} (${now.toISOString()}).\n\n`;
}

// Fields the assistant is allowed to write. ID and creation date are deliberately
// absent — the model must never be able to rewrite a task's identity or history.
const AI_EDITABLE_TASK_FIELDS = [
  'Назва', 'Опис', 'Статус', 'Пріоритет', 'Тип',
  'Дедлайн', 'Коментарі', 'Посилання на результат', 'Виконавець', 'Замовник'
];
const AI_VALID_STATUSES = ['Нова', 'В роботі', 'На перевірці', 'Виконано'];
const AI_VALID_PRIORITIES = ['Низький', 'Середній', 'Високий'];
const AI_KNOWN_ACTIONS = ['message', 'createTask', 'updateTask', 'deleteTask', 'createEvent', 'decomposeTask'];

/**
 * Strips anything the assistant is not allowed to set and validates enum values.
 * Returns only the fields that survived.
 */
function sanitizeAiTaskFields(data) {
  const clean = {};
  if (!data) return clean;

  AI_EDITABLE_TASK_FIELDS.forEach(function (field) {
    if (!data.hasOwnProperty(field)) return;

    var value = data[field];
    if (value === null || value === undefined) return;
    value = String(value).trim();
    if (value === '') return;

    if (field === 'Статус' && AI_VALID_STATUSES.indexOf(value) === -1) {
      Logger.log('ШІ повернув невідомий статус: ' + value);
      return;
    }
    if (field === 'Пріоритет' && AI_VALID_PRIORITIES.indexOf(value) === -1) {
      Logger.log('ШІ повернув невідомий пріоритет: ' + value);
      return;
    }

    clean[field] = value;
  });

  return clean;
}

/**
 * Normalizes what the model returned into a stable shape for the client:
 * {text, actions: [{action, data}]}.
 *
 * Accepts both the array form and the older single-action form, and drops
 * actions that are unknown or missing required data.
 */
function normalizeAiChatResponse(raw) {
  const response = raw || {};
  const result = { text: response.text || '', actions: [] };

  var candidates = [];
  if (Array.isArray(response.actions)) {
    candidates = response.actions;
  } else if (response.action) {
    candidates = [{ action: response.action, data: response.data }];
  }

  candidates.forEach(function (item) {
    if (!item || !item.action) return;

    const action = String(item.action);
    if (AI_KNOWN_ACTIONS.indexOf(action) === -1) {
      Logger.log('ШІ повернув невідому дію: ' + action);
      return;
    }
    if (action === 'message') return; // Text-only, nothing to execute

    const data = item.data || {};

    if (action === 'createTask') {
      const fields = sanitizeAiTaskFields(data);
      if (!fields['Назва']) return; // Nothing usable without a title
      result.actions.push({ action: action, data: fields });
      return;
    }

    if (action === 'updateTask') {
      if (!data.ID) return;
      const fields = sanitizeAiTaskFields(data);
      if (Object.keys(fields).length === 0) return;
      fields.ID = String(data.ID);
      result.actions.push({ action: action, data: fields });
      return;
    }

    if (action === 'deleteTask' || action === 'decomposeTask') {
      if (!data.ID) return;
      result.actions.push({ action: action, data: { ID: String(data.ID) } });
      return;
    }

    if (action === 'createEvent') {
      if (!data.title || !data.startDate) return;
      result.actions.push({
        action: action,
        data: {
          title: String(data.title),
          startDate: String(data.startDate),
          startTime: data.startTime ? String(data.startTime) : '09:00',
          endDate: data.endDate ? String(data.endDate) : String(data.startDate),
          endTime: data.endTime ? String(data.endTime) : '10:00',
          isAllDay: !!data.isAllDay,
          location: data.location ? String(data.location) : '',
          description: data.description ? String(data.description) : ''
        }
      });
    }
  });

  return result;
}

/**
 * Applies an assistant action after the user confirmed it in the UI.
 * Re-validates everything — the client is not trusted to have kept the
 * payload intact between the proposal and the confirmation.
 * @param {string} action - One of AI_KNOWN_ACTIONS
 * @param {Object} data - Action payload
 * @returns {Object} {ok, message}
 */
function applyAiAction(action, data) {
  const payload = data || {};

  if (action === 'createTask') {
    const fields = sanitizeAiTaskFields(payload);
    if (!fields['Назва']) {
      throw new Error('Задача без назви — нема чого створювати.');
    }
    if (!fields['Статус']) fields['Статус'] = 'Нова';
    const id = saveTask(fields);
    return { ok: true, message: 'Задачу створено.', id: id };
  }

  if (action === 'updateTask') {
    if (!payload.ID) throw new Error('Не вказано ID задачі.');

    const tasks = getTasks();
    const existing = tasks.find(function (t) { return t.ID === payload.ID; });
    if (!existing) throw new Error('Задачу не знайдено — можливо, її вже видалено.');

    const fields = sanitizeAiTaskFields(payload);
    if (Object.keys(fields).length === 0) {
      throw new Error('Немає допустимих полів для оновлення.');
    }

    // Merge onto the stored task so untouched columns survive the write
    const merged = {};
    Object.keys(existing).forEach(function (key) { merged[key] = existing[key]; });
    Object.keys(fields).forEach(function (key) { merged[key] = fields[key]; });

    if (fields['Статус'] === 'Виконано' && !merged['Дата виконання']) {
      merged['Дата виконання'] = new Date().toISOString();
    }

    saveTask(merged);
    return { ok: true, message: 'Задачу оновлено.' };
  }

  if (action === 'deleteTask') {
    if (!payload.ID) throw new Error('Не вказано ID задачі.');
    const deleted = deleteTask(payload.ID);
    if (!deleted) throw new Error('Задачу не знайдено.');
    return { ok: true, message: 'Задачу видалено.' };
  }

  if (action === 'decomposeTask') {
    if (!payload.ID) throw new Error('Не вказано ID задачі.');
    const steps = aiDecomposeTask(payload.ID);
    return { ok: true, message: 'Задачу розбито на ' + steps.length + ' кроків.', steps: steps };
  }

  if (action === 'createEvent') {
    if (!payload.title || !payload.startDate) {
      throw new Error('Для події потрібні назва та дата початку.');
    }
    const event = createCalendarEvent(payload);
    return { ok: true, message: 'Подію створено.', event: event };
  }

  throw new Error('Невідома дія: ' + action);
}

/**
 * Processes chat messages in multi-turn mode, with task/calendar context.
 */
function aiChat(message, chatHistory) {
  const fullHistory = chatHistory || [];
  fullHistory.push({ role: 'user', text: message });

  const systemInstruction = AI_CHAT_SYSTEM_INSTRUCTION +
    '\n\nДОВІДКОВИЙ КОНТЕКСТ:\n' + getAiContext();

  const responseText = callGeminiAIChat(fullHistory, systemInstruction, true);
  return normalizeAiChatResponse(parseGeminiJson(responseText, 'чат'));
}

/**
 * Parses natural text query into task properties.
 */
function aiCreateTask(naturalText) {
  const prompt = `Проаналізуй наступний опис задачі та розбий його на структуровані поля: "${naturalText}"`;
  
  const systemInstruction = `Ти — асистент TaskApp. Твоє завдання — перетворити текст у структуровані поля для створення задачі.
Поточна дата: ${new Date().toISOString().split('T')[0]}.
Поверни ТІЛЬКИ JSON об'єкт з наступними полями:
- "Назва" (обов'язково): заголовок задачі.
- "Опис" (опціонально): детальний опис.
- "Дедлайн" (опціонально): дата у форматі YYYY-MM-DD. Якщо в тексті є відносні дати (наприклад, "завтра", "до п'ятниці"), вирахуй їх на основі поточної дати.
- "Пріоритет" (опціонально): "Низький", "Середній" або "Високий".
- "Тип" (опціонально): категорія задачі (напр. Розробка, Дизайн, Особисте, Зустріч).
- "Виконавець" (опціонально): ім'я виконавця (якщо не вказано, залиш порожнім).
Не додавай ніякого іншого тексту, окрім валідного JSON.`;

  const responseText = callGeminiAI(prompt, systemInstruction, true);
  return parseGeminiJson(responseText, 'створення задачі з тексту');
}

/**
 * Decomposes task into sub-steps. Saves them directly and returns them.
 */
function aiDecomposeTask(taskId) {
  const tasks = getTasks();
  const task = tasks.find(t => t.ID === taskId);
  if (!task) {
    throw new Error('Задача не знайдена');
  }
  
  const prompt = `Назва задачі: "${task.Назва}". Опис: "${task.Опис || 'Немає'}". Нотатки: "${task.Коментарі || 'Немає'}".`;
  
  const systemInstruction = `Ти — ШІ-асистент TaskApp. Розбий цю задачу на конкретні кроки для її виконання.
Згенеруй від 3 до 7 логічних, послідовних та чітких кроків.
Поверни результат ТІЛЬКИ у форматі JSON-масиву об'єктів такого виду:
[
  {"text": "Назва кроку 1", "done": false},
  {"text": "Назва кроку 2", "done": false}
]
Не пиши ніяких вступів, висновків чи іншого тексту, окрім JSON-масиву.`;

  const responseText = callGeminiAI(prompt, systemInstruction, true);
  const steps = parseGeminiJson(responseText, 'розбиття на кроки');

  task['Кроки'] = steps;
  saveTask(task);
  
  return steps;
}

/**
 * Generates an analytical report summary for today, week, or month.
 */
function aiSummary(period) {
  const now = new Date();
  let periodText = 'сьогодні';
  if (period === 'week') periodText = 'цей тиждень';
  if (period === 'month') periodText = 'цей місяць';
  
  const prompt = `Зроби звіт про статус задач та календар на період: ${periodText}.`;
  
  const systemInstruction = `Ти — аналітик та мотиваційний асистент TaskApp.
Твоє завдання — проаналізувати задачі та календар користувача та зробити надихаючий, професійний звіт про прогрес за вказаний період.
Ось поточний час/дата: ${now.toString()}.
Ось контекст задач та календаря користувача:
${getAiContext()}

Звіт має містити:
1. Кількість виконаних задач та значущі досягнення.
2. Активні задачі з найвищим пріоритетом та найближчими дедлайнами, на які варто звернути увагу.
3. Короткий огляд розкладу на майбутнє.
4. Мотиваційне напуття для підвищення продуктивності користувача.
Звіт повинен бути оформлений у форматі Markdown.`;

  return callGeminiAI(prompt, systemInstruction);
}

/**
 * Generates a meeting preparation brief using AI.
 * @param {string} eventTitle - The title of the calendar event
 * @param {string} eventStart - The start time of the event (ISO string)
 * @returns {string} Meeting brief in Markdown format
 */
function aiMeetingPrep(eventTitle, eventStart) {
  const now = new Date();
  const eventDate = eventStart ? new Date(eventStart) : null;
  const dateStr = eventDate 
    ? eventDate.toLocaleDateString('uk-UA', { weekday: 'long', day: 'numeric', month: 'long', hour: '2-digit', minute: '2-digit' }) 
    : 'невідомо';
  
  const prompt = `Підготуй бриф для зустрічі: "${eventTitle}" (запланована на ${dateStr}).`;
  
  const systemInstruction = `Ти — професійний асистент TaskApp для підготовки до зустрічей.
Поточний час/дата: ${now.toString()}.

Ось контекст задач та календаря користувача:
${getAiContext()}

Твоє завдання — підготувати короткий, але корисний бриф для зустрічі "${eventTitle}".

Бриф має містити:
1. **Тема зустрічі** — коротко сформулюй, про що може бути зустріч на основі її назви.
2. **Пов'язані задачі** — перевір, чи є активні задачі, які можуть стосуватися цієї зустрічі. Якщо є — перелічи їх зі статусами.
3. **Питання для обговорення** — запропонуй 3-5 ключових питань або тем, які варто обговорити.
4. **Підготовка** — коротко порадь, що варто підготувати або переглянути перед зустріччю.

Відповідай українською мовою у форматі Markdown. Будь лаконічним та практичним.`;

  return callGeminiAI(prompt, systemInstruction);
}

/**
 * Auto-fills task properties (description, priority, category, days needed) based on a task title using Gemini AI.
 * @param {string} title - The title of the task
 * @returns {object} Object containing description, priority, category, and daysNeeded
 */
function aiAutofill(title) {
  const prompt = `Проаналізуй наступну назву задачі та запропонуй додаткові деталі: "${title}"`;
  
  const systemInstruction = `Ти — розумний помічник для керування задачами TaskApp.
Твоє завдання — на основі назви задачі згенерувати деталі у форматі JSON.
Поверни ТІЛЬКИ JSON об'єкт з наступними полями:
- "description": короткий опис задачі, що деталізує кроки чи суть (обов'язково українською мовою).
- "priority": одне зі значень: "Високий", "Середній" або "Низький". Спробуй оцінити пріоритет за назвою (напр. критичні баги, термінові звіти - Високий, повсякденні речі - Середній, читання чи довгострокове навчання - Низький).
- "category": коротка категорія або тег (наприклад, "Розробка", "Дизайн", "Маркетинг", "Фінанси", "Особисте", "Адміністрація" тощо).
- "daysNeeded": ціле число (кількість днів від сьогодні, необхідних для виконання задачі, наприклад: 1, 3, 5, 7). Оціни реалістичний термін для такої задачі.

Не додавай ніякого іншого тексту, крім валідного JSON.`;

  const responseText = callGeminiAI(prompt, systemInstruction, true);
  return parseGeminiJson(responseText, 'автозаповнення');
}

/**
 * Saves/updates the steps (sub-tasks) for a specific task.
 * @param {string} taskId - The ID of the task
 * @param {Array} steps - The array of steps to save
 */
function saveTaskSteps(taskId, steps) {
  const tasks = getTasks();
  const task = tasks.find(t => t.ID === taskId);
  if (!task) {
    throw new Error('Задача не знайдена');
  }
  task['Кроки'] = steps;
  saveTask(task);
}

/**
 * Analyzes the user's tasks and calendar to suggest what to do next.
 * @returns {string} Recommendation in Markdown format
 */
function aiPrioritize() {
  const now = new Date();
  const prompt = `Визнач пріоритети та порекомендуй, що робити далі.`;
  
  const systemInstruction = `Ти — розумний помічник-координатор TaskApp.
Твоє завдання — проаналізувати поточні задачі та календар користувача, визначити найважливіші та найтерміновіші справи та надати чіткі, конкретні рекомендації щодо наступних кроків.
Ось поточний час/дата: ${now.toString()}.
Ось контекст задач та календаря користувача:
${getAiContext()}

Рекомендація має містити:
1. **Головний фокус** — 1-2 найважливіші або найтерміновіші завдання, які треба зробити прямо зараз або найближчим часом.
2. **Швидкі перемоги** — завдання, які можна виконати швидко (якщо такі є).
3. **Наступна зустріч** — нагадування про найближчу подію в календарі та чи треба до неї готуватися.
4. **Порада** — одна практична порада для продуктивності на основі поточного завантаження.

Відповідай українською мовою у форматі Markdown. Будь лаконічним, чітким та практичним.`;

  return callGeminiAI(prompt, systemInstruction);
}

/**
 * Decomposes a task (by title and description) into sub-steps without saving.
 * @param {string} title - The task title
 * @param {string} description - The task description
 * @returns {Array} Array of step objects
 */
function aiDecomposeTaskText(title, description) {
  const prompt = `Назва задачі: "${title}". Опис: "${description || 'Немає'}".`;
  
  const systemInstruction = `Ти — ШІ-асистент TaskApp. Розбий цю задачу на конкретні кроки для її виконання.
Згенеруй від 3 до 7 логічних, послідовних та чітких кроків.
Поверни результат ТІЛЬКИ у форматі JSON-масиву об'єктів такого виду:
[
  {"text": "Назва кроку 1", "done": false},
  {"text": "Назва кроку 2", "done": false}
]
Не пиши ніяких вступів, висновків чи іншого тексту, окрім JSON-масиву.`;

  const responseText = callGeminiAI(prompt, systemInstruction, true);
  return parseGeminiJson(responseText, 'розбиття на кроки');
}

// ==== Document-to-Task (PDF / Word) =======
// ==========================================

/**
 * Extracts text content from a PDF or Word file uploaded as base64.
 * Uses Google Drive conversion to Google Docs for reliable text extraction.
 * Supports both Advanced Drive Service and direct REST API fallback with helpful error messaging.
 * @param {string} base64Data - The file content encoded as base64
 * @param {string} fileName - Original file name (e.g. "letter.pdf")
 * @param {string} mimeType - MIME type of the file
 * @returns {string} Extracted plain text
 */
function extractTextFromFile(base64Data, fileName, mimeType) {
  let convertedDocId = null;
  
  try {
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, mimeType, fileName);
    
    // Check if it's a DOCX file and attempt local unzip text extraction first
    if (fileName.toLowerCase().endsWith('.docx') || mimeType === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
      try {
        const docxText = extractTextFromDocx(blob);
        if (docxText && docxText.trim().length > 0) {
          Logger.log('Successfully extracted text directly from DOCX without Google Drive');
          return docxText;
        }
      } catch (docxErr) {
        Logger.log('Direct DOCX extraction failed, falling back to Drive: ' + docxErr.message);
      }
    }
    
    // Attempt using Advanced Drive Service if enabled (this automatically links and enables Drive API)
    if (typeof Drive !== 'undefined') {
      const resource = {
        mimeType: 'application/vnd.google-apps.document'
      };
      let file;
      if (typeof Drive.Files.create === 'function') {
        // Drive API v3
        resource.name = 'TaskApp_temp_' + fileName;
        file = Drive.Files.create(resource, blob);
      } else if (typeof Drive.Files.insert === 'function') {
        // Drive API v2
        resource.title = 'TaskApp_temp_' + fileName;
        file = Drive.Files.insert(resource, blob);
      } else {
        throw new Error('Не підтримується версія Drive API');
      }
      convertedDocId = file.id;
    } else {
      // Fallback to direct UrlFetchApp API request
      const resource = {
        title: 'TaskApp_temp_' + fileName,
        mimeType: 'application/vnd.google-apps.document'
      };
      
      const uploadUrl = 'https://www.googleapis.com/upload/drive/v2/files?uploadType=multipart&convert=true';
      const boundary = '---TaskAppBoundary' + Utilities.getUuid();
      const metadata = JSON.stringify(resource);
      
      const payload = Utilities.newBlob(
        '--' + boundary + '\r\n' +
        'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
        metadata + '\r\n' +
        '--' + boundary + '\r\n' +
        'Content-Type: ' + mimeType + '\r\n' +
        'Content-Transfer-Encoding: base64\r\n\r\n' +
        base64Data + '\r\n' +
        '--' + boundary + '--'
      ).getBytes();
      
      const options = {
        method: 'post',
        contentType: 'multipart/related; boundary=' + boundary,
        payload: payload,
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()
        },
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(uploadUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 403) {
        throw new Error('Код 403 (Forbidden). Для конвертації Word-файлів необхідно увімкнути "Drive API" у вашому Apps Script.\n\n' +
                        'Як це зробити:\n' +
                        '1. У лівій панелі редактора Google Apps Script натисніть "+" поруч із вкладкою "Сервіси" (Services).\n' +
                        '2. Знайдіть у списку "Drive API", виберіть його та натисніть "Додати" (Add).');
      }
      
      if (responseCode !== 200) {
        throw new Error('Не вдалося завантажити файл у Google Drive (HTTP ' + responseCode + '): ' + responseText);
      }
      
      const fileData = JSON.parse(responseText);
      convertedDocId = fileData.id;
    }
    
    // Extract text from the converted Google Doc
    const doc = DocumentApp.openById(convertedDocId);
    const text = doc.getBody().getText();
    
    return text;
  } catch (e) {
    throw new Error('Помилка витягнення тексту з файлу: ' + e.message);
  } finally {
    // Clean up temporary files from Drive
    try {
      if (convertedDocId) {
        DriveApp.getFileById(convertedDocId).setTrashed(true);
      }
    } catch (cleanupErr) {
      Logger.log('Cleanup warning: ' + cleanupErr.message);
    }
  }
}

/**
 * Local extraction of plain text from a .docx file without external API calls or Drive dependency.
 * @param {Blob} blob - The docx file blob
 * @returns {string} Extracted text
 */
function extractTextFromDocx(blob) {
  const unzipped = Utilities.unzip(blob);
  let docXml = '';
  for (let i = 0; i < unzipped.length; i++) {
    if (unzipped[i].getName() === 'word/document.xml') {
      docXml = unzipped[i].getDataAsString();
      break;
    }
  }
  if (!docXml) return '';
  
  // Extract text from <w:t> tags
  const matches = docXml.match(/<w:t[^>]*>([^<]*)<\/w:t>/g);
  if (!matches) return '';
  
  return matches.map(match => {
    // Strip XML tags
    const text = match.replace(/<[^>]+>/g, '');
    // Decode common XML entities
    return text.replace(/&amp;/g, '&')
               .replace(/&lt;/g, '<')
               .replace(/&gt;/g, '>')
               .replace(/&quot;/g, '"')
               .replace(/&apos;/g, "'");
  }).join(' ');
}

/**
 * Helper to call Gemini API directly with a file payload (e.g. PDF) via inlineData.
 */
function callGeminiAIWithFile(prompt, systemInstruction, base64Data, mimeType, isJson) {
  return callGemini({
    prompt: prompt,
    systemInstruction: systemInstruction,
    file: { base64: base64Data, mimeType: mimeType },
    json: isJson
  });
}

/**
 * Direct PDF analysis with Gemini AI, bypassing Google Drive.
 */
function aiCreateTaskFromPdfDirect(base64Data, fileName) {
  const prompt = `Проаналізуй наданий PDF документ "${fileName}" та сформуй на його основі структуровану задачу для виконання.`;
  const currentDate = new Date().toISOString().split('T')[0];
  
  const systemInstruction = `Ти — ШІ-асистент TaskApp. Твоє завдання — проаналізувати наданий PDF документ та створити на його основі структуровану задачу для виконання.
Поточна дата: ${currentDate}.

Ти маєш визначити:
1. Що потрібно зробити (суть задачі).
2. Хто є замовником або відправником (якщо зрозуміло з тексту).
3. Чи є конкретний дедлайн або терміновість.
4. Який пріоритет задачі (на основі тону листа, термінів, важливості).
5. Конкретні кроки для виконання задачі (від 3 до 7 кроків).

Поверни ТІЛЬКИ валідний JSON-об'єкт з такими полями:
{
  "Назва": "Коротка та чітка назва задачі (до 80 символів)",
  "Опис": "Детальний опис задачі на основі документу. Включи ключові деталі, вимоги та контекст з документу.",
  "Дедлайн": "YYYY-MM-DD або порожній рядок, якщо дедлайн не визначений. Якщо у тексті є відносні дати (напр. 'протягом тижня', 'до кінця місяця'), вирахуй точну дату на основі поточної дати.",
  "Пріоритет": "Високий" | "Середній" | "Низький",
  "Тип": "Категорія задачі одним-двома словами (напр. Лист, Клієнт, Запит, Звіт, Документи, Розробка тощо)",
  "Замовник": "Ім'я відправника або замовника з документу (якщо є) або порожній рядок",
  "Кроки": [
    {"text": "Крок 1: конкретна дія", "done": false},
    {"text": "Крок 2: конкретна дія", "done": false}
  ]
}

Правила:
- Назва має бути дієвою (починатися з дієслова: "Підготувати...", "Відповісти на...", "Розглянути...").
- Кроки мають бути конкретними та послідовними.
- Не додавай ніякого тексту поза JSON.`;

  const responseText = callGeminiAIWithFile(prompt, systemInstruction, base64Data, 'application/pdf', true);
  return parseGeminiJson(responseText, 'аналіз PDF');
}

/**
 * Processes an uploaded document (PDF/Word) with AI to create a structured task.
 * Extracts text from the document and sends it to Gemini for analysis.
 * @param {string} base64Data - The file content encoded as base64
 * @param {string} fileName - Original file name
 * @param {string} mimeType - MIME type of the file
 * @returns {Object} Task object with Назва, Опис, Дедлайн, Пріоритет, Тип, Кроки
 */
function aiCreateTaskFromDocument(base64Data, fileName, mimeType) {
  // If it's a PDF, bypass Drive completely and send directly to Gemini API
  if (mimeType === 'application/pdf' || fileName.toLowerCase().endsWith('.pdf')) {
    try {
      return aiCreateTaskFromPdfDirect(base64Data, fileName);
    } catch (pdfError) {
      throw new Error('Помилка прямого аналізу PDF через Gemini API: ' + pdfError.message);
    }
  }

  // Step 1: Extract text from file (for non-PDFs like Word docs)
  const documentText = extractTextFromFile(base64Data, fileName, mimeType);
  
  if (!documentText || documentText.trim().length === 0) {
    throw new Error('Не вдалося витягнути текст з документу. Файл може бути порожнім або містити лише зображення.');
  }
  
  // Step 2: Truncate very long documents (keep first ~8000 chars for API limits)
  const maxLength = 8000;
  const truncatedText = documentText.length > maxLength 
    ? documentText.substring(0, maxLength) + '\n\n[...документ обрізано через великий розмір...]' 
    : documentText;
  
  // Step 3: Send to Gemini AI for task creation
  const prompt = `Проаналізуй наступний документ (лист, завдання, запит тощо) та сформуй на його основі структуровану задачу для виконання:\n\n---\nФайл: "${fileName}"\n\nТекст документу:\n${truncatedText}\n---`;
  
  const currentDate = new Date().toISOString().split('T')[0];
  
  const systemInstruction = `Ти — ШІ-асистент TaskApp. Твоє завдання — проаналізувати наданий документ (лист, email, запит клієнта, ТЗ, доручення тощо) та створити на його основі структуровану задачу для виконання.

Поточна дата: ${currentDate}.

Ти маєш визначити:
1. Що потрібно зробити (суть задачі).
2. Хто є замовником або відправником (якщо зрозуміло з тексту).
3. Чи є конкретний дедлайн або терміновість.
4. Який пріоритет задачі (на основі тону листа, термінів, важливості).
5. Конкретні кроки для виконання задачі (від 3 до 7 кроків).

Поверни ТІЛЬКИ валідний JSON-об'єкт з такими полями:
{
  "Назва": "Коротка та чітка назва задачі (до 80 символів)",
  "Опис": "Детальний опис задачі на основі документу. Включи ключові деталі, вимоги та контекст з документу.",
  "Дедлайн": "YYYY-MM-DD або порожній рядок, якщо дедлайн не визначений. Якщо у тексті є відносні дати (напр. 'протягом тижня', 'до кінця місяця'), вирахуй точну дату на основі поточної дати.",
  "Пріоритет": "Високий" | "Середній" | "Низький",
  "Тип": "Категорія задачі одним-двома словами (напр. Лист, Клієнт, Запит, Звіт, Документи, Розробка тощо)",
  "Замовник": "Ім'я відправника або замовника з документу (якщо є) або порожній рядок",
  "Кроки": [
    {"text": "Крок 1: конкретна дія", "done": false},
    {"text": "Крок 2: конкретна дія", "done": false}
  ]
}

Правила:
- Назва має бути дієвою (починатися з дієслова: "Підготувати...", "Відповісти на...", "Розглянути...").
- Кроки мають бути конкретними та послідовними.
- Не додавай ніякого тексту поза JSON.`;

  const responseText = callGeminiAI(prompt, systemInstruction, true);
  return parseGeminiJson(responseText, 'аналіз документа');
}

/**
 * Запустіть цю функцію ОДИН раз із редактора Google Apps Script,
 * щоб з'явилося вікно авторизації та ви надали дозволи на роботу з Google Drive API.
 */
function testDrive() {
  Logger.log('Drive API status: ' + (typeof Drive !== 'undefined' ? 'Enabled' : 'Disabled'));
  if (typeof Drive !== 'undefined') {
    try {
      if (typeof Drive.Files.list === 'function') {
        Drive.Files.list({maxResults: 1});
      } else {
        Drive.Files.list();
      }
      Logger.log('Drive API авторизовано успішно!');
    } catch(e) {
      Logger.log('Помилка авторизації Drive API: ' + e.message);
    }
  }
}

/**
 * Завантажує та парсить телефонний довідник з окремої таблиці.
 */
function getPhoneDirectory() {
  try {
    const ss = SpreadsheetApp.openById(DIRECTORY_SPREADSHEET_ID);
    const sheet = ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();
    const fontWeights = sheet.getDataRange().getFontWeights();
    
    let currentDepartment = "Загальні контакти";
    let contacts = [];
    
    // Рядок 0 - Заголовки, Рядок 1 - нумерація колонок, починаємо з 2
    for (let i = 2; i < data.length; i++) {
      let position = String(data[i][0] || '').trim();
      let name = String(data[i][1] || '').trim();
      let personalPhone = String(data[i][2] || '').trim();
      let corpPhone = String(data[i][3] || '').trim();
      
      if (!position && !name) continue; // Порожній рядок
      
      let isBold = fontWeights[i][0] === 'bold';
      
      // Визначаємо, чи це заголовок відділу
      if ((isBold && !name) || (!name && position === position.toUpperCase() && position.length > 3)) {
        // Очищаємо від зайвих пробілів на початку (напр. "  2.1. Відділ радників")
        currentDepartment = position.replace(/^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g, '');
        continue;
      }
      
      // Якщо це контакт (або вакантна посада без імені)
      contacts.push({
        department: currentDepartment,
        position: position,
        name: name,
        personalPhone: personalPhone,
        corpPhone: corpPhone
      });
    }
    
    return contacts;
  } catch (e) {
    Logger.log("Помилка завантаження довідника: " + e.message);
    throw new Error("Не вдалося завантажити довідник. " + e.message);
  }
}

// ==========================================
// ====== ClickUp API Integration ===========
// ==========================================

const CLICKUP_BASE_URL = 'https://api.clickup.com/api/v2';
const CLICKUP_DESC_PREVIEW_LENGTH = 200;
const CLICKUP_MAX_PAGES = 20;
const CLICKUP_TOKEN_PROPERTY = 'CLICKUP_TOKEN';

/**
 * Reads the stored ClickUp token.
 * Kept server-side so AI functions can reach ClickUp without the client
 * having to hand them a token on every call.
 */
function getClickUpToken() {
  const token = PropertiesService.getUserProperties().getProperty(CLICKUP_TOKEN_PROPERTY);
  return (token && token.trim()) ? token.trim() : null;
}

/**
 * Stores the ClickUp token after verifying it works.
 * @param {string} token - ClickUp Personal API Token
 * @returns {Object} {hasToken, user}
 */
function setClickUpToken(token) {
  if (!token || token.trim() === '') {
    throw new Error('ClickUp API токен не надано.');
  }

  // Validate before storing — /user is the cheapest authenticated endpoint
  const userData = clickUpFetch('/user', token.trim());

  PropertiesService.getUserProperties().setProperty(CLICKUP_TOKEN_PROPERTY, token.trim());
  invalidateClickUpCache();
  invalidateAiContextCache();

  return { hasToken: true, user: mapClickUpUser(userData.user) };
}

/**
 * Forgets the stored ClickUp token.
 */
function clearClickUpTokenServer() {
  PropertiesService.getUserProperties().deleteProperty(CLICKUP_TOKEN_PROPERTY);
  invalidateClickUpCache();
  invalidateAiContextCache();
  return { hasToken: false };
}

/**
 * Whether ClickUp is connected — lets the client render the right view
 * without ever holding the token itself.
 */
function getClickUpStatus() {
  return { hasToken: !!getClickUpToken() };
}

/**
 * Resolves the token to use: an explicitly passed one wins, otherwise the stored one.
 */
function resolveClickUpToken(token) {
  const resolved = (token && token.trim()) ? token.trim() : getClickUpToken();
  if (!resolved) {
    throw new Error('ClickUp не підключено. Введіть Personal API Token у вкладці ClickUp.');
  }
  return resolved;
}

function invalidateClickUpCache() {
  try {
    CacheService.getUserCache().removeAll(['cu_tasks_open', 'cu_tasks_all']);
  } catch (e) {
    Logger.log('Кеш: не вдалося скинути ClickUp: ' + e.message);
  }
}

/**
 * Low-level request to the ClickUp API v2.
 * @param {string} endpoint - API endpoint path (e.g. "/user", "/task/abc123")
 * @param {string} token - ClickUp Personal API Token
 * @param {string} method - HTTP method ('get' by default)
 * @param {Object} payload - Optional JSON body for post/put requests
 * @returns {Object} Parsed JSON response
 */
function clickUpFetch(endpoint, token, method, payload) {
  if (!token || token.trim() === '') {
    throw new Error('ClickUp API токен не надано.');
  }

  const options = {
    method: method || 'get',
    headers: {
      'Authorization': token.trim(),
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  if (payload) {
    options.payload = JSON.stringify(payload);
  }

  const response = UrlFetchApp.fetch(CLICKUP_BASE_URL + endpoint, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  if (code === 401) {
    throw new Error('Невалідний ClickUp API токен. Перевірте токен і спробуйте ще раз.');
  }

  if (code < 200 || code >= 300) {
    throw new Error('ClickUp API помилка (HTTP ' + code + '): ' + text);
  }

  return text ? JSON.parse(text) : {};
}

/**
 * Makes an HTTP GET request to the ClickUp API v2.
 */
function clickUpApiRequest(endpoint, token) {
  return clickUpFetch(endpoint, token, 'get');
}

/**
 * Builds the "Space → Folder → List" breadcrumb for a task.
 */
function buildClickUpPath(task) {
  const listName = (task.list || {}).name || '';
  const folderName = (task.folder || {}).name || '';
  const spaceName = (task.space || {}).name || '';
  // If folder is hidden (same as list), don't duplicate
  const folderHidden = (task.folder || {}).hidden || false;

  let path = spaceName;
  if (folderName && !folderHidden) {
    path += ' → ' + folderName;
  }
  if (listName) {
    path += ' → ' + listName;
  }
  return path;
}

/**
 * Normalizes a ClickUp user object into the shape the UI expects.
 */
function mapClickUpUser(user) {
  if (!user) return null;
  const name = user.username || user.email || '';
  const initials = user.initials || name.split(' ').map(function (p) {
    return p.charAt(0).toUpperCase();
  }).join('').substring(0, 2);

  return {
    id: user.id,
    name: name,
    initials: initials,
    color: user.color || '#7c4dff'
  };
}

/**
 * Normalizes a raw ClickUp task into the shape the UI expects.
 * The description is truncated here — the full text is served by getClickUpTaskDetails().
 */
function mapClickUpTask(task, teamName) {
  const statusObj = task.status || {};
  const fullDesc = task.text_content || task.description || '';

  return {
    id: task.id,
    name: task.name,
    description: fullDesc.substring(0, CLICKUP_DESC_PREVIEW_LENGTH),
    hasMoreDescription: fullDesc.length > CLICKUP_DESC_PREVIEW_LENGTH,
    status: {
      name: statusObj.status || 'unknown',
      color: statusObj.color || '#808080',
      type: statusObj.type || ''
    },
    priority: (task.priority && task.priority.id) ? parseInt(task.priority.id, 10) : null,
    dueDate: task.due_date ? parseInt(task.due_date, 10) : null,
    startDate: task.start_date ? parseInt(task.start_date, 10) : null,
    timeEstimate: task.time_estimate ? parseInt(task.time_estimate, 10) : null,
    timeSpent: task.time_spent ? parseInt(task.time_spent, 10) : null,
    parent: task.parent || null,
    listId: (task.list || {}).id || null,
    listName: (task.list || {}).name || '',
    url: task.url || ('https://app.clickup.com/t/' + task.id),
    path: buildClickUpPath(task),
    workspace: teamName || '',
    tags: (task.tags || []).map(function (tag) { return tag.name; }),
    creator: mapClickUpUser(task.creator),
    assignees: (task.assignees || []).map(mapClickUpUser).filter(Boolean)
  };
}

const CLICKUP_STATUS_CACHE_SECONDS = 21600; // 6 h — list statuses change rarely
const CLICKUP_FETCHALL_CHUNK = 40;

/**
 * Normalizes the statuses array of a ClickUp list resource.
 */
function mapClickUpStatuses(listResource) {
  return ((listResource || {}).statuses || []).map(function (s) {
    return {
      name: s.status,
      color: s.color || '#808080',
      type: s.type || '',
      orderindex: s.orderindex
    };
  });
}

/**
 * Fetches the status sets of several lists at once, so the UI can offer a status
 * dropdown per task. Cached results are reused; the rest are fetched in parallel.
 * @param {Array} listIds - Unique ClickUp list IDs
 * @param {string} token - ClickUp Personal API Token
 * @returns {Object} Map of listId -> array of {name, color, type, orderindex}
 */
function fetchClickUpStatusesForLists(listIds, token) {
  const result = {};
  if (!listIds || listIds.length === 0) return result;

  const cache = CacheService.getUserCache();
  const cacheKeys = listIds.map(function (id) { return 'cu_statuses_' + id; });
  const cached = cache.getAll(cacheKeys) || {};

  const missing = [];
  listIds.forEach(function (id) {
    const raw = cached['cu_statuses_' + id];
    if (raw) {
      try {
        result[id] = JSON.parse(raw);
        return;
      } catch (e) {
        // Corrupted cache entry — just refetch it
      }
    }
    missing.push(id);
  });

  if (missing.length === 0) return result;

  const toCache = {};

  // Fetch the cache misses in parallel batches instead of one request at a time
  for (var offset = 0; offset < missing.length; offset += CLICKUP_FETCHALL_CHUNK) {
    var chunk = missing.slice(offset, offset + CLICKUP_FETCHALL_CHUNK);

    var requests = chunk.map(function (listId) {
      return {
        url: CLICKUP_BASE_URL + '/list/' + listId,
        method: 'get',
        headers: {
          'Authorization': token.trim(),
          'Content-Type': 'application/json'
        },
        muteHttpExceptions: true
      };
    });

    var responses;
    try {
      responses = UrlFetchApp.fetchAll(requests);
    } catch (e) {
      Logger.log('Не вдалося завантажити статуси списків: ' + e.message);
      continue;
    }

    responses.forEach(function (response, i) {
      var listId = chunk[i];
      if (response.getResponseCode() !== 200) {
        Logger.log('Статуси списку ' + listId + ' недоступні (HTTP ' + response.getResponseCode() + ')');
        return;
      }
      try {
        var statuses = mapClickUpStatuses(JSON.parse(response.getContentText()));
        result[listId] = statuses;
        toCache['cu_statuses_' + listId] = JSON.stringify(statuses);
      } catch (e) {
        Logger.log('Не вдалося розібрати статуси списку ' + listId + ': ' + e.message);
      }
    });
  }

  if (Object.keys(toCache).length > 0) {
    cache.putAll(toCache, CLICKUP_STATUS_CACHE_SECONDS);
  }

  return result;
}

/**
 * Fetches the set of statuses available for a single list.
 * @param {string} listId - The ClickUp list ID
 * @param {string} token - ClickUp Personal API Token
 * @returns {Array} Array of {name, color, type, orderindex}
 */
function getClickUpListStatuses(listId, token) {
  const map = fetchClickUpStatusesForLists([listId], token);
  return map[listId] || [];
}

/**
 * Fetches tasks assigned to the current user from all ClickUp workspaces,
 * without the per-list status sets. Walks every page — ClickUp returns 100 per page.
 *
 * Split out from getClickUpTasks so the AI context can reuse the task list
 * without paying for the status lookups it does not need.
 *
 * @param {string} token - ClickUp Personal API Token
 * @param {boolean} includeClosed - Whether to include closed/done tasks
 * @returns {Object} {tasks, truncated, currentUser}
 */
function fetchClickUpTasksRaw(token, includeClosed) {
  const cacheKey = includeClosed ? 'cu_tasks_all' : 'cu_tasks_open';
  const cached = cacheGetJson(cacheKey);
  if (cached) return cached;

  // Step 1: Get current user info
  const userData = clickUpFetch('/user', token);
  const userId = userData.user.id;
  const currentUser = mapClickUpUser(userData.user);

  // Step 2: Get all workspaces (teams)
  const teamsData = clickUpFetch('/team', token);
  const teams = teamsData.teams || [];

  if (teams.length === 0) {
    return { tasks: [], truncated: false, currentUser: currentUser };
  }

  // Step 3: Fetch every page of tasks assigned to this user from each workspace
  var allTasks = [];
  var truncated = false;

  teams.forEach(function (team) {
    try {
      var page = 0;
      var lastPage = false;

      while (!lastPage && page < CLICKUP_MAX_PAGES) {
        var endpoint = '/team/' + team.id + '/task'
          + '?assignees[]=' + userId
          + '&subtasks=true'
          + '&include_closed=' + (includeClosed ? 'true' : 'false')
          + '&order_by=due_date'
          + '&reverse=false'
          + '&page=' + page;

        var tasksData = clickUpFetch(endpoint, token);
        var tasks = tasksData.tasks || [];

        tasks.forEach(function (task) {
          allTasks.push(mapClickUpTask(task, team.name));
        });

        lastPage = (tasksData.last_page === true) || tasks.length === 0;
        page++;
      }

      // Hit the page cap with more results still waiting — tell the UI instead of silently dropping them
      if (!lastPage) {
        truncated = true;
        Logger.log('Досягнуто ліміт сторінок для workspace ' + team.name);
      }
    } catch (e) {
      Logger.log('Помилка завантаження задач з workspace ' + team.name + ': ' + e.message);
    }
  });

  // Sort: tasks with due dates first (ascending), then tasks without
  allTasks.sort(function (a, b) {
    if (a.dueDate && b.dueDate) return a.dueDate - b.dueDate;
    if (a.dueDate) return -1;
    if (b.dueDate) return 1;
    return 0;
  });

  const result = { tasks: allTasks, truncated: truncated, currentUser: currentUser };
  cachePutJson(cacheKey, result, CACHE_TTL_CLICKUP_TASKS);
  return result;
}

/**
 * Fetches tasks assigned to the current user, plus the status set of every list
 * involved so the UI can offer an inline status dropdown.
 * @param {string} token - ClickUp Personal API Token (optional if stored)
 * @param {boolean} includeClosed - Whether to include closed/done tasks
 * @returns {Object} {tasks, statusesByList, truncated, currentUser}
 */
function getClickUpTasks(token, includeClosed) {
  const authToken = resolveClickUpToken(token);
  const raw = fetchClickUpTasksRaw(authToken, includeClosed);

  var listIdSet = {};
  raw.tasks.forEach(function (t) {
    if (t.listId) listIdSet[t.listId] = true;
  });

  var statusesByList = {};
  try {
    statusesByList = fetchClickUpStatusesForLists(Object.keys(listIdSet), authToken);
  } catch (e) {
    // Without statuses the UI falls back to read-only badges — not worth failing the whole load
    Logger.log('Не вдалося завантажити статуси списків: ' + e.message);
  }

  return {
    tasks: raw.tasks,
    statusesByList: statusesByList,
    truncated: raw.truncated,
    currentUser: raw.currentUser
  };
}

/**
 * Formats ClickUp tasks for the AI context. Returns an empty string when
 * ClickUp is not connected, so the section simply does not appear.
 */
function getClickUpContext() {
  const token = getClickUpToken();
  if (!token) return '';

  try {
    const raw = fetchClickUpTasksRaw(token, false);
    if (!raw.tasks || raw.tasks.length === 0) {
      return 'CLICKUP: активних задач немає.\n';
    }

    const shown = raw.tasks.slice(0, AI_CTX_MAX_CLICKUP_TASKS);
    let context = `ЗАДАЧІ З CLICKUP (${shown.length} з ${raw.tasks.length}, від найближчого дедлайну):\n`;

    const priorityNames = { 1: 'Urgent', 2: 'High', 3: 'Normal', 4: 'Low' };

    shown.forEach(function (t, index) {
      context += `${index + 1}. [${t.status.name}] "${t.name}" | ClickUp ID: ${t.id}\n`;

      const attrs = [];
      if (t.dueDate) attrs.push('дедлайн: ' + formatDateShort(new Date(t.dueDate).toISOString()));
      if (t.priority && priorityNames[t.priority]) attrs.push('пріоритет: ' + priorityNames[t.priority]);
      if (t.path) attrs.push('список: ' + t.path);
      if (t.creator && t.creator.name) attrs.push('створив: ' + t.creator.name);
      if (attrs.length > 0) context += '   ' + attrs.join(' | ') + '\n';
    });

    if (raw.tasks.length > shown.length) {
      context += `…та ще ${raw.tasks.length - shown.length} задач у ClickUp.\n`;
    }

    return context;
  } catch (e) {
    Logger.log('Не вдалося отримати контекст ClickUp: ' + e.message);
    return '';
  }
}

/**
 * Fetches full details for a single task — full description, assignees,
 * time tracking, checklists and subtasks.
 * @param {string} taskId - The ClickUp task ID
 * @param {string} token - ClickUp Personal API Token
 * @returns {Object} Detailed task object
 */
function getClickUpTaskDetails(taskId, token) {
  if (!taskId) {
    throw new Error('ID задачі не надано.');
  }

  const authToken = resolveClickUpToken(token);
  const task = clickUpFetch('/task/' + taskId + '?include_subtasks=true', authToken);
  const details = mapClickUpTask(task, '');

  // Full, untruncated description
  details.description = task.text_content || task.description || '';
  details.hasMoreDescription = false;

  details.dateCreated = task.date_created ? parseInt(task.date_created, 10) : null;
  details.dateUpdated = task.date_updated ? parseInt(task.date_updated, 10) : null;

  details.checklists = (task.checklists || []).map(function (cl) {
    return {
      name: cl.name || '',
      items: (cl.items || []).map(function (item) {
        return { name: item.name || '', resolved: !!item.resolved };
      })
    };
  });

  details.subtasks = (task.subtasks || []).map(function (st) {
    const stStatus = st.status || {};
    return {
      id: st.id,
      name: st.name,
      status: {
        name: stStatus.status || '',
        color: stStatus.color || '#808080',
        type: stStatus.type || ''
      }
    };
  });

  return details;
}

/**
 * Updates a task's status, priority and/or due date.
 * @param {string} taskId - The ClickUp task ID
 * @param {string} token - ClickUp Personal API Token
 * @param {Object} updates - {status?: string, priority?: number|null, dueDate?: number|null}
 * @returns {Object} The updated field values as ClickUp stored them
 */
function updateClickUpTask(taskId, token, updates) {
  if (!taskId) {
    throw new Error('ID задачі не надано.');
  }
  if (!updates) {
    throw new Error('Немає полів для оновлення.');
  }

  const payload = {};

  if (updates.status) {
    payload.status = updates.status;
  }

  // priority: 1-4, or null to clear it
  if (updates.hasOwnProperty('priority')) {
    payload.priority = (updates.priority === null || updates.priority === '')
      ? null
      : parseInt(updates.priority, 10);
  }

  // dueDate: unix milliseconds, or null to clear it
  if (updates.hasOwnProperty('dueDate')) {
    payload.due_date = (updates.dueDate === null || updates.dueDate === '')
      ? null
      : parseInt(updates.dueDate, 10);
    payload.due_date_time = false;
  }

  if (Object.keys(payload).length === 0) {
    throw new Error('Немає полів для оновлення.');
  }

  const task = clickUpFetch('/task/' + taskId, resolveClickUpToken(token), 'put', payload);
  const statusObj = task.status || {};

  // The cached task list and AI context both just went stale
  invalidateClickUpCache();
  invalidateAiContextCache();

  return {
    id: task.id,
    status: {
      name: statusObj.status || '',
      color: statusObj.color || '#808080',
      type: statusObj.type || ''
    },
    priority: (task.priority && task.priority.id) ? parseInt(task.priority.id, 10) : null,
    dueDate: task.due_date ? parseInt(task.due_date, 10) : null
  };
}

/**
 * Makes an HTTP POST request to the ClickUp API v2.
 * @param {string} endpoint - API endpoint path (e.g. "/task/{id}/comment")
 * @param {string} token - ClickUp Personal API Token
 * @param {Object} payload - JSON payload to send
 * @returns {Object} Parsed JSON response
 */
function clickUpApiPost(endpoint, token, payload) {
  return clickUpFetch(endpoint, token, 'post', payload);
}

/**
 * Fetches comments for a specific ClickUp task.
 * @param {string} taskId - The ClickUp task ID
 * @param {string} token - ClickUp Personal API Token
 * @returns {Array} Array of formatted comment objects
 */
function getClickUpTaskComments(taskId, token) {
  if (!taskId) {
    throw new Error('ID задачі не надано.');
  }

  const data = clickUpApiRequest('/task/' + taskId + '/comment', resolveClickUpToken(token));
  var comments = data.comments || [];
  
  return comments.map(function(comment) {
    // Extract plain text from comment_text array
    var textParts = [];
    if (comment.comment_text) {
      textParts.push(comment.comment_text);
    } else if (comment.comment && Array.isArray(comment.comment)) {
      comment.comment.forEach(function(part) {
        if (part.text) {
          textParts.push(part.text);
        }
      });
    }
    
    var userName = '';
    if (comment.user) {
      userName = comment.user.username || comment.user.email || '';
    }
    
    var userInitials = '';
    if (userName) {
      var parts = userName.split(' ');
      userInitials = parts.map(function(p) { return p.charAt(0).toUpperCase(); }).join('').substring(0, 2);
    }
    
    var userColor = comment.user && comment.user.color ? comment.user.color : '#7c4dff';
    
    return {
      id: comment.id,
      text: textParts.join(''),
      date: comment.date ? parseInt(comment.date) : null,
      userName: userName,
      userInitials: userInitials,
      userColor: userColor,
      userId: comment.user ? comment.user.id : null
    };
  });
}

/**
 * Posts a new comment to a ClickUp task.
 * @param {string} taskId - The ClickUp task ID
 * @param {string} token - ClickUp Personal API Token
 * @param {string} commentText - The comment text to post
 * @returns {Object} The created comment data
 */
function postClickUpTaskComment(taskId, token, commentText) {
  if (!taskId) {
    throw new Error('ID задачі не надано.');
  }
  if (!commentText || commentText.trim() === '') {
    throw new Error('Текст коментаря не може бути порожнім.');
  }

  var payload = {
    comment_text: commentText.trim(),
    notify_all: true
  };

  return clickUpApiPost('/task/' + taskId + '/comment', resolveClickUpToken(token), payload);
}

/**
 * Summarizes a ClickUp task together with its whole comment thread.
 * Most useful on tasks where the actual state lives in 30 comments rather
 * than in the description.
 * @param {string} taskId - The ClickUp task ID
 * @returns {string} Brief in Markdown format
 */
function aiClickUpTaskBrief(taskId) {
  if (!taskId) {
    throw new Error('ID задачі не надано.');
  }

  const token = resolveClickUpToken(null);
  const task = getClickUpTaskDetails(taskId, token);

  var comments = [];
  try {
    comments = getClickUpTaskComments(taskId, token);
  } catch (e) {
    Logger.log('Не вдалося завантажити коментарі для брифу: ' + e.message);
  }

  const priorityNames = { 1: 'Urgent', 2: 'High', 3: 'Normal', 4: 'Low' };

  let taskBlock = `Назва: "${task.name}"\n`;
  taskBlock += `Статус: ${task.status.name}\n`;
  if (task.priority && priorityNames[task.priority]) taskBlock += `Пріоритет: ${priorityNames[task.priority]}\n`;
  if (task.dueDate) taskBlock += `Дедлайн: ${formatDateShort(new Date(task.dueDate).toISOString())}\n`;
  if (task.path) taskBlock += `Розташування: ${task.path}\n`;
  if (task.creator && task.creator.name) taskBlock += `Створив: ${task.creator.name}\n`;
  if (task.assignees && task.assignees.length > 0) {
    taskBlock += `Виконавці: ${task.assignees.map(function (a) { return a.name; }).join(', ')}\n`;
  }
  if (task.timeEstimate) taskBlock += `Оцінка: ${Math.round(task.timeEstimate / 60000)} хв\n`;
  if (task.timeSpent) taskBlock += `Витрачено: ${Math.round(task.timeSpent / 60000)} хв\n`;
  taskBlock += `\nОпис:\n${truncateForContext(task.description, 2000) || 'Немає'}\n`;

  if (task.checklists && task.checklists.length > 0) {
    taskBlock += '\nЧеклісти:\n';
    task.checklists.forEach(function (cl) {
      const done = cl.items.filter(function (i) { return i.resolved; }).length;
      taskBlock += `- ${cl.name} [${done}/${cl.items.length}]\n`;
      cl.items.forEach(function (item) {
        taskBlock += `  ${item.resolved ? '☑' : '☐'} ${item.name}\n`;
      });
    });
  }

  if (task.subtasks && task.subtasks.length > 0) {
    taskBlock += '\nПідзадачі:\n';
    task.subtasks.forEach(function (st) {
      taskBlock += `- [${st.status.name}] ${st.name}\n`;
    });
  }

  // ClickUp returns comments newest-first; reverse so the model reads the thread chronologically
  let commentsBlock = 'Коментарів немає.';
  if (comments.length > 0) {
    const ordered = comments.slice().reverse();
    commentsBlock = ordered.map(function (c) {
      const when = c.date ? new Date(c.date).toLocaleString('uk-UA') : '';
      return `[${when}] ${c.userName}: ${truncateForContext(c.text, 600)}`;
    }).join('\n');
  }

  const prompt = `ЗАДАЧА:\n${taskBlock}\n\nОБГОВОРЕННЯ (від найстарішого до найновішого):\n${commentsBlock}`;

  const systemInstruction = `Ти — асистент TaskApp. Тобі дано задачу з ClickUp разом із усім обговоренням у коментарях.
Поточна дата: ${new Date().toISOString().split('T')[0]}.

Зроби стислий бриф, щоб людина за 30 секунд зрозуміла стан справ, не читаючи всю переписку.

Бриф має містити:
1. **Суть** — чого від нас хочуть, одним-двома реченнями.
2. **Що вже зроблено** — на основі коментарів та чеклістів.
3. **Що блокує** — відкриті питання, очікування на когось, ризики. Якщо нічого не блокує — так і напиши.
4. **Наступний крок** — одна конкретна дія, яку варто зробити далі.

Правила:
- Відповідай українською у форматі Markdown.
- Будь конкретним: імена, дати, цифри з обговорення.
- Не вигадуй того, чого немає в наданих даних.
- Якщо коментарів немає — скажи це прямо і будуй бриф лише з опису.`;

  return callGemini({ prompt: prompt, systemInstruction: systemInstruction });
}

