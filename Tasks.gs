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
