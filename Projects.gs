// ==========================================
// ============== Projects ==================
// ==========================================

const PROJECT_JSON_FIELDS = ['Гілки', 'Контакти'];

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
      'ID', 'Дата створення', 'Дата оновлення', 'Назва', 'Опис', 'Статус',
      'Гілки', 'Контакти'
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

  if (!headers.includes('Дата оновлення')) {
    headers.push('Дата оновлення');
    sheet.getRange(1, headers.length).setValue('Дата оновлення');
  }

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

  // Every save — create or edit — counts as touching the project, so this
  // always reflects when the project's info was last actually looked at.
  projectToSave['Дата оновлення'] = new Date();

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
