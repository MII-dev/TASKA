// Колонки, які мають лишатися справжніми датами в таблиці. Спільні для аркушів
// Tasks, Projects і Specs — див. buildSheetRow().
const SHEET_DATE_COLUMNS = [
  'Дата створення', 'Дата оновлення', 'Дедлайн', 'Дата виконання'
];

/**
 * Parses a value headed for a date column into a real Date.
 * A bare "2026-09-10" is built in local terms on purpose — `new Date(string)`
 * would read it as UTC midnight and land on the previous day west of Greenwich.
 * @returns {Date|null} null when the value is not a date at all
 */
function parseSheetDate(value) {
  if (value instanceof Date) return isNaN(value.getTime()) ? null : value;
  if (!value) return null;

  const text = String(value).trim();
  const bare = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (bare) {
    return new Date(Number(bare[1]), Number(bare[2]) - 1, Number(bare[3]));
  }

  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Lays a record out in header order for a sheet write.
 *
 * Replaces the `|| ''` pattern, which had two faults: a `false` or `0` was
 * blanked, and a date column that round-tripped through the client as an ISO
 * string went back in as text. That second one was permanent — after the first
 * edit "Дедлайн" stopped being a date and sorting and filters broke in the
 * sheet itself.
 * @param {string[]} headers - The sheet's header row
 * @param {Object} record - Values keyed by header name
 * @returns {Array} Row values ready for setValues()/appendRow()
 */
function buildSheetRow(headers, record) {
  return headers.map(function (header) {
    const value = record[header];
    if (value === undefined || value === null) return '';

    if (SHEET_DATE_COLUMNS.indexOf(header) !== -1) {
      const asDate = parseSheetDate(value);
      if (asDate) return asDate;
    }
    return value;
  });
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

  if (!headers.includes('CalendarEventID')) {
    headers.push('CalendarEventID');
    sheet.getRange(1, headers.length).setValue('CalendarEventID');
  }

  if (!headers.includes('Дата оновлення')) {
    headers.push('Дата оновлення');
    sheet.getRange(1, headers.length).setValue('Дата оновлення');
  }

  // Ensure steps are stringified for storage
  const taskToSave = {...task};
  if (taskToSave['Кроки'] && typeof taskToSave['Кроки'] !== 'string') {
    taskToSave['Кроки'] = JSON.stringify(taskToSave['Кроки']);
  }

  // Every save counts as touching the task — mirrors how Projects tracks staleness.
  taskToSave['Дата оновлення'] = new Date();

  let rowIndex = -1;
  let existingRow = null;
  if (task.ID) {
    // Find existing task by ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == task.ID) {
        rowIndex = i + 1;
        existingRow = data[i];
        break;
      }
    }
  } else {
    // Generate new ID
    taskToSave.ID = Utilities.getUuid();
    taskToSave['Дата створення'] = new Date();
  }

  // The row snapshot predates this call's header migrations, so a task that
  // never had a deadline/event column simply reads back undefined here —
  // treated the same as "no event yet".
  const deadlineIdx = headers.indexOf('Дедлайн');
  const eventIdIdx = headers.indexOf('CalendarEventID');
  const oldDeadline = existingRow ? existingRow[deadlineIdx] : null;
  const oldEventId = existingRow ? existingRow[eventIdIdx] : null;

  // Always recomputed here rather than trusted from the caller — the sheet
  // row, not the incoming payload, is the source of truth for what's linked.
  taskToSave['CalendarEventID'] = syncTaskDeadlineEvent(taskToSave, oldDeadline, oldEventId);

  const rowValues = buildSheetRow(headers, taskToSave);
  
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  invalidateAiContextCache();
  return taskToSave.ID;
}

/**
 * Serializes a sheet mutation across concurrent calls.
 *
 * Deletes read the whole sheet, work out a row index, then call deleteRow().
 * With two calls in flight the second index is stale and removes the wrong
 * record. The UI fires overlapping google.script.run calls, so a double-tap is
 * enough to reach it.
 * @param {Function} fn - The mutation to run under the lock
 */
function withSheetLock(fn) {
  const lock = LockService.getUserLock();
  lock.waitLock(15000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Deletes a task by ID.
 */
function deleteTask(id) {
  return withSheetLock(function () {
    const sheet = initSheet();
    const data = sheet.getDataRange().getValues();
    const eventIdIdx = data[0].indexOf('CalendarEventID');

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        if (eventIdIdx !== -1 && data[i][eventIdIdx]) {
          deleteTaskDeadlineEvent(data[i][eventIdIdx]);
        }
        sheet.deleteRow(i + 1);
        invalidateAiContextCache();
        return true;
      }
    }
    return false;
  });
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
