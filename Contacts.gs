// ==========================================
// ============== Contacts ==================
// ==========================================

/**
 * Initializes the Contacts sheet with headers if it's empty.
 */
function initContactsSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONTACTS_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONTACTS_SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    const headers = [
      'ID', 'Дата створення', 'Дата оновлення', "Ім'я", 'Посада', 'Компанія',
      'Телефон', 'Email', 'Telegram', 'Категорія', 'Обраний', 'Нотатки'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Fetches all contacts from the spreadsheet.
 */
function getContacts() {
  const sheet = initContactsSheet();
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
    // Normalize the favorite flag — sheet may hold a real boolean or a stray 'TRUE' string
    obj['Обраний'] = obj['Обраний'] === true || obj['Обраний'] === 'TRUE';
    return obj;
  });
}

/**
 * Turns what the user said into invitable email addresses: "Олег" is looked up
 * in the sheet, a plain address is taken as given. Names the sheet cannot settle
 * are reported rather than guessed — inviting the wrong person is worse than
 * inviting nobody.
 * @param {string[]} guests - Names, emails, or a mix
 * @returns {Object} {emails, unresolved: [{name, reason}]}
 */
function resolveGuestEmails(guests) {
  const contacts = getContacts().filter(function (contact) {
    return contact.Email && String(contact.Email).indexOf('@') > 0;
  });

  const emails = [];
  const unresolved = [];
  const seen = {};

  function addEmail(email) {
    const normalized = String(email).trim();
    const key = normalized.toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      emails.push(normalized);
    }
  }

  (guests || []).forEach(function (rawGuest) {
    const guest = String(rawGuest || '').trim();
    if (!guest) return;

    if (guest.indexOf('@') > 0) {
      addEmail(guest);
      return;
    }

    const needle = guest.toLowerCase();
    const exact = contacts.filter(function (contact) {
      return String(contact["Ім'я"] || '').trim().toLowerCase() === needle;
    });
    const matches = exact.length ? exact : contacts.filter(function (contact) {
      return String(contact["Ім'я"] || '').toLowerCase().indexOf(needle) !== -1;
    });

    if (matches.length === 1) {
      addEmail(matches[0].Email);
    } else {
      unresolved.push({ name: guest, reason: matches.length ? 'ambiguous' : 'not_found' });
    }
  });

  return { emails: emails, unresolved: unresolved };
}

/**
 * Saves a contact (creates new or updates existing).
 */
function saveContact(contact) {
  const sheet = initContactsSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const contactToSave = {...contact};

  let rowIndex = -1;
  if (contact.ID) {
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == contact.ID) {
        rowIndex = i + 1;
        break;
      }
    }
  } else {
    contactToSave.ID = Utilities.getUuid();
    contactToSave['Дата створення'] = new Date();
  }

  contactToSave['Дата оновлення'] = new Date();

  // Falls back to '' only for genuinely missing fields — a `false` favorite
  // flag must survive, unlike the `|| ''` pattern used for text-only rows.
  const rowValues = headers.map(header => contactToSave[header] !== undefined ? contactToSave[header] : '');

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }

  // Контакти входять у контекст асистента — інакше він не побачить нову людину
  invalidateAiContextCache();

  return contactToSave.ID;
}

/**
 * Deletes a contact by ID.
 */
function deleteContact(id) {
  return withSheetLock(function () {
    const sheet = initContactsSheet();
    const data = sheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == id) {
        sheet.deleteRow(i + 1);
        invalidateAiContextCache();
        return true;
      }
    }
    return false;
  });
}
