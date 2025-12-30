const CONFIG = {
  APP_NAME: "University Staff Hub",
  MASTER_SHEET_ID: "1234567890", // Replace with your actual Master Sheet ID

  TABS: {
    STAFF_LIST: "Staff_List",
    STAFF_ASSIGNMENTS: "Staff_Assignments",
    STAFF_AVAILABILITY: "Staff_Availability",
    STAFF_PREFERENCES: "Staff_Preferences",
    COURSE_SCHEDULE: "Course_Schedule",
    TECH_HUB_SHIFTS: "Tech_Hub_Shifts",
    EVENT_TYPES: "Event_Types",
    SETTINGS: "Permissions_Matrix",
    LOGS: "Logs",
    REPORT_DATA: "Report_Data"
  },

  HEADERS: {
    REPORTING: ["Date", "StaffID", "StaffName", "Role", "AssignmentType", "ReferenceID", "Description", "StartTime", "EndTime", "DurationHours", "Status", "PayableHours"]
  },
  
  SETTINGS_KEYS: {
      CALENDAR: 'calendarSettings',
      NURSING: 'nursingExamSettings',
      MLT: 'mltExamSettings',
      TECH_HUB: 'techHubSettings'
  },

  PAGES: [
    { id: 'page-dashboard', title: 'My Dashboard' },
    { id: 'page-scheduling', title: 'Master Schedule' },
    { id: 'page-proctoring-nursing', title: 'Nursing Proctoring' },
    { id: 'page-proctoring-mlt', title: 'MLT Proctoring' },
    { id: 'page-classrooms', title: 'Classroom Support' },
    { id: 'page-analytics', title: 'Analytics' },
    { id: 'page-settings', title: 'Admin Settings' },
  ],
  
  NURSING: {
      ROSTER_KEYWORD: 'roster',
      URLS: {
          RED_FLAG_REPORT: 'https://docs.google.com/forms/d/e/1FAIpQLSfORKCKol8SsRldNKfvsDy3ILNs9HcFv3gKb8TuxrNrlqxijw/viewform',
          PROTOCOL_DOC: 'https://docs.google.com/document/d/1TgKtmoDFqXLK0lBFPNirOAz_TW4S3E_BFhS934VcjOo/edit'
      }
  },

  MLT: {
      DEFAULTS: {
          ROSTER_KEYWORD: 'roster',
          KEYWORDS: {
              EXAM: 'exam',
              DATE: 'date',
              START_TIME: 'start time',
              DURATION: 'duration',
              ROOM: 'room',
              PASSWORD: 'password',
              START_SITE: 'start site' // Legacy
          }
      }
  },

  ASSIGNMENT_TYPES: {
    TECH_HUB: 'Tech Hub',
    COURSE: 'Course'
  },

  STATUS: {
    PLANNED: 'Planned'
  },

  DATE_FORMATS: {
    ISO: 'yyyy-MM-dd'
  },

  WEEKDAYS: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],

  COLUMN_KEYS: {
      STAFF_ID: 'StaffID',
      SHIFT_ID: 'ShiftID',
      COURSE_ID: 'CourseID'
  }
};