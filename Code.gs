/**
 * ----------------------------------------------------------------------------------------
 * Main Code File for Google Apps Script
 * 
 * Contains server-side logic for the Proctoring Tool application, including the main
 * web app entry point (doGet) and general API functions.
 * ----------------------------------------------------------------------------------------
 */

// --- Global Variable --- //
const g = {};

/**
 * ----------------------------------------------------------------------------------------
 * Web App & Add-on UI Functions
 * ----------------------------------------------------------------------------------------
 */

// Main entry point for the web application
function doGet(e) {
  // Use createTemplateFromFile and evaluate() to process the scriptlets (<?!= ... ?>)
  const html = HtmlService.createTemplateFromFile('Index.html').evaluate();
  return html
      .setTitle('Proctoring Tool')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Allows embedding in sites
}

// Entry point for launching as a Google Sheets Add-on
function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('Launch Proctoring Tool', 'showSidebar')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  // Also correct the sidebar to use the templating engine
  const ui = HtmlService.createTemplateFromFile('Index.html').evaluate().setTitle('Proctoring Tool');
  SpreadsheetApp.getUi().showSidebar(ui);
}

/**
 * Includes the content of another file in the current HTML template.
 * This is a standard Apps Script templating feature.
 * @param {string} filename The name of the file to include.
 * @return {string} The content of the included file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
