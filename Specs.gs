/**
 * Initializes the Specs sheet with headers if it's empty.
 */
function initSpecsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SPECS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SPECS_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    const headers = [
      'ID', 'Дата створення', 'Дата оновлення', 'Проєкт', 'Тип',
      'ПоходитьВід', 'Назва', 'Статус', 'Зміст', 'Джерело', 'Версія'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Fetches all specs (ТВ and ТЗ) from the spreadsheet.
 */
function getSpecs() {
  const sheet = initSpecsSheet();
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
    return obj;
  });
}

// Sheets caps a cell at 50,000 characters — stay under that with a margin
// so saveSpec() fails loudly instead of letting content silently truncate.
const SPEC_CONTENT_MAX_CHARS = 49000;

/**
 * Saves a spec (creates new or updates existing). Auto-increments Версія on
 * every update to an existing row; starts at 1 on create.
 */
function saveSpec(spec) {
  const sheet = initSpecsSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const content = spec['Зміст'] || '';
  if (content.length > SPEC_CONTENT_MAX_CHARS) {
    throw new Error(
      'Текст задовгий (' + content.length + ' символів, максимум ' + SPEC_CONTENT_MAX_CHARS +
      '). Розбийте на кілька документів.'
    );
  }

  const specToSave = { ...spec };

  let rowIndex = -1;
  let existingRow = null;
  if (spec.ID) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == spec.ID) {
        rowIndex = i + 1;
        existingRow = data[i];
        break;
      }
    }
  } else {
    specToSave.ID = Utilities.getUuid();
    specToSave['Дата створення'] = new Date();
  }

  if (!specToSave['Статус']) specToSave['Статус'] = 'Чернетка';
  if (!specToSave['Джерело']) specToSave['Джерело'] = 'Вручну';

  const versionIdx = headers.indexOf('Версія');
  const oldVersion = existingRow ? parseInt(existingRow[versionIdx], 10) : 0;
  specToSave['Версія'] = (oldVersion > 0 ? oldVersion : 0) + 1;

  specToSave['Дата оновлення'] = new Date();

  const rowValues = headers.map(header => specToSave[header] !== undefined ? specToSave[header] : '');

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  return specToSave.ID;
}

/**
 * Deletes a spec by ID.
 */
function deleteSpec(id) {
  const sheet = initSpecsSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}
