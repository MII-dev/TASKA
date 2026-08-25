# TASKA

Персональний таск-менеджер на Google Apps Script: Google Sheets як база даних,
Google Calendar та ClickUp як інтеграції, Gemini як AI-асистент.

## Стек

- **Backend**: Google Apps Script (`.gs`), дані зберігаються в Google Sheets
  (аркуші `Tasks`, `Projects`, `Contacts`).
- **Frontend**: один HTML-документ, який Apps Script збирає на льоту з
  партиалів через `HtmlService`-темплейтинг (`<?!= include('Name') ?>`).
- **Деплой**: [`clasp`](https://github.com/google/clasp) — CLI Google для
  Apps Script, замість ручного копіювання коду у веб-редактор.

## Структура

### Backend (`.gs`)

Apps Script склеює всі `.gs`-файли проєкту в одну спільну область
видимості під час виконання — поділ на файли суто організаційний,
на поведінку коду не впливає.

| Файл | Відповідальність |
|---|---|
| `Code.gs` | Точка входу (`doGet`), глобальні константи, хелпер `include()` |
| `Cache.gs` | Обгортки над `CacheService`, інвалідація AI-контексту |
| `Tasks.gs` | CRUD задач (аркуш `Tasks`) |
| `Projects.gs` | CRUD проєктів і гілок (аркуш `Projects`) |
| `Contacts.gs` | CRUD контактів (аркуш `Contacts`) |
| `Calendar.gs` | Читання/створення подій Google Calendar |
| `AI.gs` | Інтеграція з Gemini: чат, дії ШІ, побудова контексту |
| `Documents.gs` | Створення задач з PDF/DOCX через Gemini |
| `ClickUp.gs` | Інтеграція з ClickUp API v2 |

### Frontend (`.html`)

`index.html` — тонкий скелет (~100 рядків), що збирає сторінку з партиалів:

| Файл | Вміст |
|---|---|
| `index.html` | `<head>`, шапка, навігація по вкладках, виклики `include()` |
| `Styles.html` | Весь CSS (design tokens, компоненти) |
| `TasksView.html` | Розмітка вкладки «Задачі» |
| `ProjectsView.html` | Розмітка вкладки «Проєкти» |
| `ContactsView.html` | Розмітка вкладки «Телефони» |
| `ClickUpView.html` | Розмітка вкладки «ClickUp» |
| `Modals.html` | Усі модальні вікна (задача, проєкт, контакт, файл тощо) |
| `Script.html` | Весь клієнтський JS |
| `AiChatPanel.html` | Плаваюча кнопка та панель AI-чату |

Партиали **не** обробляються темплейтинг-рушієм самі по собі — лише
`index.html` сканується на `<? ?>`-скриплети; вміст інших файлів
вставляється як є через `include()`.

## Розробка

```bash
npm install       # встановлює clasp
npm run login     # одноразова авторизація в Google-акаунті
npm run push      # відправити локальні зміни в Apps Script (HEAD)
npm run pull      # підтягнути стан з Apps Script локально
npm run open      # відкрити проєкт в онлайн-редакторі
```

`.clasp.json` містить `scriptId` існуючого проєкту — вже налаштований,
нічого створювати заново не треба.

### Деплой змін на робоче посилання

`push` оновлює лише HEAD-версію скрипту. Щоб зміни зʼявились на
робочому (версійному) посиланні, яким користуються щодня, потрібно
створити нову версію і прив'язати її до існуючого деплойменту:

```bash
npx clasp deployments               # знайти deploymentId робочого деплою
npx clasp deploy -i <deploymentId> -d "опис зміни"
```

Перед оновленням робочого посилання варто перевірити HEAD окремо —
в онлайн-редакторі: **Deploy → Test deployments** дає посилання, що
завжди показує останній запушений код, без впливу на робочий деплой.

### Секрети

`GEMINI_API_KEY` і `CLICKUP_TOKEN` зберігаються в
`PropertiesService` (Script/User Properties) на боці Apps Script,
а не в коді — налаштовуються через UI застосунку (банер налаштувань
ШІ, вкладка ClickUp).
