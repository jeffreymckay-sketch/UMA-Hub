/**
 * ----------------------------------------------------------------------------------------
 * Main Code File for Google Apps Script
 * ----------------------------------------------------------------------------------------
 */

// Global container (Server-side)
const g = {};

function doGet(e) {
  const html = HtmlService.createTemplateFromFile('Index.html').evaluate();
  // FIX: Updated Title
  return html
      .setTitle('Dept. Management App') 
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('Launch App', 'showSidebar')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function showSidebar() {
  const ui = HtmlService.createTemplateFromFile('Index.html').evaluate().setTitle('Dept. Management App');
  SpreadsheetApp.getUi().showSidebar(ui);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}