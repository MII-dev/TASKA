const SPREADSHEET_ID = '1djbxaGjw1qIpFBZb2P14noEY0iErch1hSOBBxlcXP5k';
const SHEET_NAME = 'Tasks';
const PROJECTS_SHEET_NAME = 'Projects';
const CONTACTS_SHEET_NAME = 'Contacts';

/**
 * Serves the HTML file.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('TaskApp - Управління задачами')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
