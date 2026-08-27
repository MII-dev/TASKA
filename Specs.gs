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

// ==========================================
// ======= AI: ТВ/ТЗ authoring ==============
// ==========================================

/**
 * Builds a lean project-context block for spec drafting — just the
 * project's own name/description/branch names, not the full getAiContext()
 * (tasks/calendar/ClickUp) that chat/summary use, since drafting a document
 * doesn't need that noise.
 */
function getProjectContextForSpec(projectId) {
  if (!projectId) return '';
  const project = getProjects().find(p => p.ID === projectId);
  if (!project) return '';

  let context = `Проєкт: "${project.Назва}"\n`;
  if (project.Опис) context += `Опис проєкту: ${project.Опис}\n`;

  const branches = Array.isArray(project.Гілки) ? project.Гілки : [];
  if (branches.length > 0) {
    context += `Існуючі напрямки роботи (гілки): ${branches.map(b => b.name).join(', ')}\n`;
  }

  return context;
}

/**
 * Drafts a ТВ (технічні вимоги) — the loose, exploratory first stage.
 * Deliberately allowed to be incomplete: an "Відкриті питання" section is
 * part of the point of a ТВ, not a flaw in the output.
 * @param {Object} input - {title, goal, projectId}
 * @returns {string} Markdown document
 */
function aiDraftRequirements(input) {
  const opts = input || {};
  const projectContext = getProjectContextForSpec(opts.projectId);

  const prompt = `Назва документа: "${opts.title || ''}".\nМета/ідея: ${opts.goal || ''}`;

  const systemInstruction = `Ти — бізнес-аналітик, що допомагає власнику проєкту сформулювати технічні вимоги (ТВ) — перший, чорновий етап перед формальним технічним завданням (ТЗ).

${projectContext ? 'Контекст проєкту:\n' + projectContext + '\n' : ''}
На основі короткого опису мети/ідеї сформуй документ технічних вимог у форматі Markdown з розділами:

## Проблема / потреба
## Цільові користувачі
## Бажаний результат
## Відомі обмеження
## Відкриті питання

"Відкриті питання" — це нормально й очікувано: перелічи конкретні питання, відповіді на які потрібні, щоб перейти від ТВ до повноцінного ТЗ. Не вигадуй відповіді на них.

Відповідай українською мовою. Без вступів на кшталт "Ось документ" — одразу починай з "## Проблема / потреба".`;

  return callGeminiAI(prompt, systemInstruction, false);
}

/**
 * Formalizes a ТВ into a ТЗ (технічне завдання) — the precise, structured
 * document that task generation reads from.
 * @param {string} requirementsContent - The source ТВ's full markdown
 * @param {string} projectId
 * @returns {string} Markdown document
 */
function aiFormalizeSpec(requirementsContent, projectId) {
  const projectContext = getProjectContextForSpec(projectId);

  const prompt = `Технічні вимоги (ТВ) для формалізації у ТЗ:\n\n${requirementsContent}`;

  const systemInstruction = `Ти — технічний письменник, що перетворює чорнові технічні вимоги (ТВ) на формальне технічне завдання (ТЗ), готове для планування розробки.

${projectContext ? 'Контекст проєкту:\n' + projectContext + '\n' : ''}
На основі наданого ТВ сформуй ТЗ у форматі Markdown з розділами:

## Загальна інформація
## Мета та цілі
## Опис функціональності
## Вимоги до інтерфейсу
## Нефункціональні вимоги
## Критерії приймання
## Що поза межами

Правила:
- Розділ "Відкриті питання" з ТВ потрібно перетворити на конкретні рішення там, де вхідних даних достатньо для обґрунтованого висновку.
- Де інформації справді не вистачає — прямо зазнач залишену невизначеність у відповідному розділі (наприклад, у "Опис функціональності"), а не вигадуй деталі, яких немає в ТВ.
- Це формальний документ — уникай розмовного тону, будь конкретним і структурованим.

Відповідай українською мовою. Без вступів — одразу починай з "## Загальна інформація".`;

  return callGeminiAI(prompt, systemInstruction, false);
}

/**
 * Imports an existing ТВ/ТЗ document from a PDF or Word file.
 * DOCX goes through plain text extraction only — no Gemini call — since
 * that's already a faithful transcript and re-running it through the model
 * would only add cost/latency/paraphrasing risk for no benefit. PDF has no
 * plain-text-extraction path anywhere in this app, so it goes through
 * Gemini's file understanding instead, which makes it a reconstruction
 * rather than a verbatim transcript — callers should tag the result
 * accordingly (see 'Джерело' handling on the client).
 * @param {string} base64Data
 * @param {string} fileName
 * @param {string} mimeType
 * @returns {string} Extracted/reconstructed document text
 */
function aiImportSpecFromDocument(base64Data, fileName, mimeType) {
  const isPdf = mimeType === 'application/pdf' || fileName.toLowerCase().endsWith('.pdf');

  if (!isPdf) {
    const text = extractTextFromFile(base64Data, fileName, mimeType);
    if (!text || text.trim().length === 0) {
      throw new Error('Не вдалося витягнути текст з документу. Файл може бути порожнім.');
    }
    return text;
  }

  const prompt = `Перетвори наданий PDF документ "${fileName}" на чистий текст у форматі Markdown, зберігаючи структуру розділів документа.`;
  const systemInstruction = `Ти розпізнаєш вміст PDF документа (технічні вимоги або технічне завдання) та переносиш його в чистий Markdown.
Зберігай оригінальну структуру розділів та формулювання настільки точно, наскільки можливо — це не переказ, а транскрипція.
Не додавай власних коментарів, оцінок чи розділів, яких немає в оригіналі.
Поверни лише текст документа, без вступів.`;

  const text = callGeminiAIWithFile(prompt, systemInstruction, base64Data, 'application/pdf', false);
  if (!text || text.trim().length === 0) {
    throw new Error('Не вдалося розпізнати вміст PDF документу.');
  }
  return text;
}

const SPEC_TASK_GEN_MAX_CHARS = 20000; // higher than the ~8000 used for ancillary letters — here the document IS the primary content
const SPEC_TASK_GEN_MAX_ITEMS = 40;    // hard cap regardless of what the model returns, even if it ignores the "≤25" instruction
const SPEC_TASK_VALID_PRIORITIES = ['Високий', 'Середній', 'Низький'];

/**
 * Analyzes a ТВ/ТЗ document and proposes a task breakdown for implementing
 * it — a proposal only, nothing is written here. The caller (client) shows
 * these for review/edit/selection before any saveTask() call.
 * @param {string} specContent - The document's markdown content
 * @param {string} projectId - Used only to pull light project context into the prompt
 * @returns {Array} Array of {Назва, Опис, Тип, Пріоритет, ГілкаНазва}
 */
function aiGenerateTasksFromSpec(specContent, projectId) {
  const projectContext = getProjectContextForSpec(projectId);

  const truncated = specContent.length > SPEC_TASK_GEN_MAX_CHARS
    ? specContent.substring(0, SPEC_TASK_GEN_MAX_CHARS) + '\n\n[...документ обрізано через великий розмір...]'
    : specContent;

  const prompt = `Документ (ТВ/ТЗ) для аналізу:\n\n${truncated}`;

  const systemInstruction = `Ти — технічний керівник проєкту, що розбиває технічний документ (ТВ або ТЗ) на конкретні задачі для команди розробки.

${projectContext ? 'Контекст проєкту:\n' + projectContext + '\n' : ''}
Проаналізуй наданий документ і поверни ТІЛЬКИ JSON-масив, не більше 25 елементів — найважливіші, конкретні, самодостатні одиниці роботи (кожна повинна бути виконувана окремо, без обов'язкової залежності від порядку виконання інших).

Кожен елемент масиву:
{
  "Назва": "Дієва, конкретна назва задачі (до 80 символів, починається з дієслова)",
  "Опис": "Детальний опис того, що саме треба зробити, на основі документу",
  "Тип": "Категорія задачі одним-двома словами (напр. Backend, Frontend, Дизайн, Тестування, Документація)",
  "Пріоритет": "Високий" | "Середній" | "Низький",
  "ГілкаНазва": "Назва функціональної області/фічі, до якої належить задача (для групування, напр. 'Авторизація', 'Календар')"
}

Не додавай жодного тексту поза JSON-масивом.`;

  const responseText = callGeminiAI(prompt, systemInstruction, true);
  const parsed = parseGeminiJson(responseText, 'генерація задач з ТЗ');

  if (!Array.isArray(parsed)) {
    throw new Error('ШІ повернув неочікуваний формат відповіді. Спробуйте ще раз або скоротіть документ.');
  }

  return parsed
    .filter(item => item && typeof item['Назва'] === 'string' && item['Назва'].trim() !== '')
    .slice(0, SPEC_TASK_GEN_MAX_ITEMS)
    .map(item => ({
      'Назва': String(item['Назва']).trim(),
      'Опис': item['Опис'] ? String(item['Опис']).trim() : '',
      'Тип': item['Тип'] ? String(item['Тип']).trim() : '',
      'Пріоритет': SPEC_TASK_VALID_PRIORITIES.indexOf(item['Пріоритет']) !== -1 ? item['Пріоритет'] : 'Середній',
      'ГілкаНазва': item['ГілкаНазва'] ? String(item['ГілкаНазва']).trim() : 'Без групи'
    }));
}

/**
 * Revises a spec's full content in response to a free-form instruction.
 * Works on either ТВ or ТЗ content — type-agnostic, operates on whatever
 * markdown is passed in. Returns the FULL revised document, not a diff.
 * @param {string} specContent - Current markdown content
 * @param {string} instruction - Free-form revision instruction
 * @returns {string} Revised markdown document
 */
function aiReviseSpec(specContent, instruction) {
  const prompt = `Поточний документ:\n\n${specContent}\n\n---\n\nІнструкція для редагування: ${instruction}`;

  const systemInstruction = `Ти — технічний редактор. Тобі надано документ (технічні вимоги або технічне завдання) та інструкцію, як його змінити.

Застосуй інструкцію та поверни ПОВНИЙ оновлений текст документа (не тільки змінену частину, не diff) у тому ж форматі Markdown, зі збереженням структури розділів, якщо інструкція прямо не просить її змінити.

Відповідай лише текстом документа, українською мовою, без коментарів про те, що було змінено.`;

  return callGeminiAI(prompt, systemInstruction, false);
}
