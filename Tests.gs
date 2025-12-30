/**
 * ==================================================================
 * TEST SUITE FOR REFACTORING
 * ==================================================================
 * To run these tests:
 * 1. Open the Apps Script editor.
 * 2. Select the "runRefactoringTests" function from the dropdown menu.
 * 3. Click the "Run" button.
 * 4. View the logs by going to "View" > "Logs" (or Ctrl+Enter).
 *
 * These tests verify that the constants in the Config.gs file match
 * the original hard-coded values that were replaced during refactoring.
 * This helps ensure that the refactoring did not introduce any bugs.
 */

function runRefactoringTests() {
  const testRunner = new TestRunner();

  // Test Suite: Config.gs values
  testRunner.run('CONFIG.SETTINGS_KEYS should be correct', () => {
    testRunner.assertEqual(CONFIG.SETTINGS_KEYS.NURSING, 'nursingExamSettings', 'NURSING settings key');
    testRunner.assertEqual(CONFIG.SETTINGS_KEYS.MLT, 'mltExamSettings', 'MLT settings key');
    testRunner.assertEqual(CONFIG.SETTINGS_KEYS.TECH_HUB, 'techHubSettings', 'TECH_HUB settings key (for scheduling)');
  });

  testRunner.run('CONFIG.NURSING should be correct', () => {
    testRunner.assertEqual(CONFIG.NURSING.ROSTER_KEYWORD, 'roster', 'Nursing roster keyword');
    testRunner.assertEqual(CONFIG.NURSING.URLS.RED_FLAG_REPORT, 'https://docs.google.com/forms/d/e/1FAIpQLSfORKCKol8SsRldNKfvsDy3ILNs9HcFv3gKb8TuxrNrlqxijw/viewform', 'Red Flag Report URL');
    testRunner.assertEqual(CONFIG.NURSING.URLS.PROTOCOL_DOC, 'https://docs.google.com/document/d/1TgKtmoDFqXLK0lBFPNirOAz_TW4S3E_BFhS934VcjOo/edit', 'Protocol Doc URL');
  });
  
  testRunner.run('CONFIG.MLT should be correct', () => {
    testRunner.assertEqual(CONFIG.MLT.DEFAULTS.ROSTER_KEYWORD, 'roster', 'MLT roster keyword');
  });

  testRunner.run('CONFIG.ASSIGNMENT_TYPES should be correct', () => {
    testRunner.assertEqual(CONFIG.ASSIGNMENT_TYPES.TECH_HUB, 'Tech Hub', 'Tech Hub assignment type');
    testRunner.assertEqual(CONFIG.ASSIGNMENT_TYPES.COURSE, 'Course', 'Course assignment type');
  });

  testRunner.run('CONFIG.STATUS should be correct', () => {
    testRunner.assertEqual(CONFIG.STATUS.PLANNED, 'Planned', 'Planned status');
  });
  
  testRunner.run('CONFIG.COLUMN_KEYS should be correct', () => {
    testRunner.assertEqual(CONFIG.COLUMN_KEYS.STAFF_ID, 'StaffID', 'StaffID column key');
    testRunner.assertEqual(CONFIG.COLUMN_KEYS.SHIFT_ID, 'ShiftID', 'ShiftID column key');
    testRunner.assertEqual(CONFIG.COLUMN_KEYS.COURSE_ID, 'CourseID', 'CourseID column key');
  });

  // --- End of Tests ---
  testRunner.report();
}


/**
 * A simple helper class for running tests and reporting results.
 */
class TestRunner {
  constructor() {
    this.results = [];
    this.currentTest = null;
    this.assertions = 0;
  }

  run(testName, testFn) {
    this.currentTest = testName;
    this.assertions = 0;
    try {
      testFn();
      this.results.push({ name: this.currentTest, status: 'PASSED', message: `(${this.assertions} assertions)` });
    } catch (e) {
      this.results.push({ name: this.currentTest, status: 'FAILED', message: e.message });
    }
  }

  assertEqual(actual, expected, message) {
    this.assertions++;
    if (actual !== expected) {
      throw new Error(`Assertion failed: ${message}. Expected "${expected}", but got "${actual}".`);
    }
  }

  report() {
    Logger.log('--- TEST RESULTS ---');
    let failures = 0;
    this.results.forEach(result => {
      if (result.status === 'FAILED') {
        Logger.log(`❌ ${result.status}: ${result.name}`);
        Logger.log(`   └> ${result.message}`);
        failures++;
      } else {
        Logger.log(`✅ ${result.status}: ${result.name} ${result.message}`);
      }
    });
    Logger.log('--------------------');
    if (failures > 0) {
      Logger.log(`🚨 SUMMARY: ${failures} out of ${this.results.length} tests failed.`);
    } else {
      Logger.log(`👍 SUMMARY: All ${this.results.length} tests passed!`);
    }
  }
}
