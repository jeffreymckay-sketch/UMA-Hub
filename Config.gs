/**
 * -------------------------------------------------------------------
 * CONFIGURATION & CONSTANTS
 * -------------------------------------------------------------------
 */

const CONFIG = {
  ROLES: ['Admin', 'Lead', 'Staff'],
  
  PAGES: [
    'page-dashboard',
    'page-tech-hub',
    'page-mst',
    'page-classroom',
    'page-exams',
    'page-mlt-proctoring',
    'page-analytics',
    'page-inspector',
    'page-guide',
    'page-settings'
  ],
  
  TABS: {
    STAFF_LIST: 'Staff_List',
    CLASSROOM_LIST: 'Classroom_List',
    COURSE_SCHEDULE: 'Course_Schedule',
    TECH_HUB_SHIFTS: 'TechHub_Shifts',
    STAFF_ASSIGNMENTS: 'Staff_Assignments',
    EXAM_APPOINTMENTS: 'Exam_Appointments',
    REPORT_DATA: 'Report_Data',
    STAFF_AVAILABILITY: 'Staff_Availability',
    STAFF_PREFERENCES: 'Staff_Preferences',
    PERMISSIONS_MATRIX: 'Permissions_Matrix',
    EVENT_TYPES: 'Event_Types',
    SYSTEM_LOGS: 'System_Logs'
  },

  HEADERS: {
    STAFF: ['FullName', 'StaffID', 'Roles', 'IsActive', 'Notes'],
    ASSIGNMENTS: ['AssignmentID', 'StaffID', 'AssignmentType', 'ReferenceID', 'StartDate', 'EndDate', 'StartTime', 'EndTime'],
    AVAILABILITY: ['AvailabilityID', 'StaffID', 'DayOfWeek', 'StartTime', 'EndTime'],
    PREFERENCES: ['StaffID', 'TimeBlock', 'Preference'],
    SHIFTS: ['ShiftID', 'Description', 'DayOfWeek', 'StartTime', 'EndTime', 'Zoom'],
    REPORTING: ['Date', 'StaffID', 'FullName', 'PrimaryRole', 'AssignmentType', 'ReferenceID', 'AssignmentDescription', 'PlannedStart', 'PlannedEnd', 'PlannedDurationHours', 'EventStatus', 'ActualDurationHours']
  },

  // NEW: Specific Configuration for MST Course Schedule
  MST: {
    HEADERS: {
      // Exact headers provided by user
      ASSIGNED_STAFF: ['mst assigned by email'], 
      COURSE_UNIQUE_ID: ['eventid'], // The unique ID for saving/logic
      COURSE_CODE: ['course'],       // The display name (e.g. HUS 236)
      FACULTY: ['faculty'],
      DAY: ['day'],
      TIME: ['run time'],            // Specifically "Run Time"
      LOCATION: ['bx location']      // Specifically "BX Location"
    }
  },

  NURSING: {
    KEYWORDS: {
      EXAM: 'exam',
      DATE: 'date',
      START_SITE: 'site', 
      START_ZOOM: 'zoom', 
      DURATION: 'duration',
      ROOM: 'room',
      PASSWORD: 'password'
    },
    ROSTER_KEYWORD: 'augusta',
    TRIGGER_FUNCTION: 'nursing_automatedEmailTrigger'
  },

  MLT: {
    DEFAULTS: {
      KEYWORDS: {
        EXAM: 'exam',
        DATE: 'date',
        START_TIME: 'start time', 
        DURATION: 'duration',
        ROOM: 'room',
        PASSWORD: 'password'
      },
      ROSTER_KEYWORD: 'student',
      SUB_TABLE_STOP: 'end roster'
    }
  }
};