const SPREADSHEET_ID = '1djbxaGjw1qIpFBZb2P14noEY0iErch1hSOBBxlcXP5k';
const SHEET_NAME = 'Tasks';
const PROJECTS_SHEET_NAME = 'Projects';
const CONTACTS_SHEET_NAME = 'Contacts';

/**
 * Serves the HTML file. Uses the templating engine (not createHtmlOutputFromFile)
 * because index.html assembles itself out of partials via include().
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index').evaluate()
    .setTitle('TaskApp - Управління задачами')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Inlines another HTML file's raw content at a <?!= include('Name') ?> call site.
 * The included file's own content is not re-evaluated as a template — it is only
 * index.html itself that gets scanned for <? ?> scriptlets.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
