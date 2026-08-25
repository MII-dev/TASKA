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
