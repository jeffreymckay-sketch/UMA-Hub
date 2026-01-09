/**
 * @OnlyCurrentDoc
 */

/**
 * Serves the main HTML page of the web app.
 * @returns {HtmlOutput} The HTML page to be displayed.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('University Dept. Management')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Includes the content of another HTML file.
 * This is a common pattern in Google Apps Script web apps.
 * @param {string} filename The name of the file to include.
 * @returns {string} The content of the file.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Fetches the current user's information.
 * This is called by the client-side JavaScript on page load.
 * @returns {object} An object containing the user's email and a success flag.
 */
function api_getUserInfo() {
  try {
    return {
      success: true,
      data: {
        email: Session.getEffectiveUser().getEmail(),
        photoUrl: '' // Placeholder for a profile photo URL
      }
    };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
