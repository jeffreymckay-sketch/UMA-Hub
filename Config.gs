/**
 * @file Config.gs
 * @description This file contains the functions for managing application settings
 * using PropertiesService. It is the single source of truth for all configuration.
 */

/**
 * Retrieves all application settings from PropertiesService.
 *
 * @returns {object} The settings object. Returns a blank object if no settings are found.
 */
function getSettings() {
  try {
    const properties = PropertiesService.getScriptProperties().getProperties();
    return properties;
  } catch (e) {
    console.error('Error retrieving settings from PropertiesService: ' + e.message);
    // Return an empty object to prevent downstream errors
    return {};
  }
}

/**
 * Saves a settings object to PropertiesService.
 * This will overwrite all existing properties.
 *
 * @param {object} settings The settings object to save.
 * @returns {void}
 */
function saveSettings(settings) {
  try {
    PropertiesService.getScriptProperties().setProperties(settings, true); // true to delete other properties
  } catch (e) {
    console.error('Error saving settings to PropertiesService: ' + e.message);
    throw new Error('Failed to save settings. ' + e.message);
  }
}

/**
 * Saves a single setting (key-value pair) to PropertiesService.
 *
 * @param {string} key The key for the setting.
 * @param {string} value The value for the setting.
 * @returns {void}
 */
function saveSetting(key, value) {
  try {
    PropertiesService.getScriptProperties().setProperty(key, value);
  } catch (e) {
    console.error(`Error saving setting '${key}' to PropertiesService: ${e.message}`);
    throw new Error(`Failed to save setting '${key}'. ${e.message}`);
  }
}
