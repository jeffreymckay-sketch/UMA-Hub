
/**
 * ------------------------------------------------------------------
 * DATA MODELS
 * ------------------------------------------------------------------
 * This file defines the core data structures for the application.
 * These models provide a consistent and predictable structure for
 * staff, events (courses and shifts), and assignments.
 * ------------------------------------------------------------------
 */

/**
 * Represents a staff member.
 * @class
 * @param {string} id - The unique identifier for the staff member (email address).
 * @param {string} name - The full name of the staff member.
 * @param {string} role - The role or roles of the staff member.
 * @param {boolean} isActive - Whether the staff member is currently active.
 */
class Staff {
    constructor(id, name, role, isActive) {
        this.id = id;
        this.name = name;
        this.role = role;
        this.isActive = isActive;
    }
}

/**
 * Represents a schedulable event. This is a base class.
 * @class
 * @param {string} id - The unique identifier for the event.
 * @param {string} type - The type of the event (e.g., 'COURSE', 'SHIFT').
 * @param {string} name - The name or title of the event.
 * @param {Date} startDate - The start date and time of the event.
 * @param {Date} endDate - The end date and time of the event.
 */
class SchedulableEvent {
    constructor(id, type, name, startDate, endDate) {
        if (this.constructor === SchedulableEvent) {
            throw new Error("SchedulableEvent is an abstract class and cannot be instantiated directly.");
        }
        this.id = id;
        this.type = type;
        this.name = name;
        this.startDate = startDate;
        this.endDate = endDate;
    }
}

/**
 * Represents a course, extending SchedulableEvent.
 * @class
 * @extends SchedulableEvent
 * @param {string} id - The unique course ID.
 * @param {string} name - The name of the course.
 * @param {Date} startDate - The start date and time of the course.
 * @param {Date} endDate - The end date and time of the course.
 * @param {string} faculty - The faculty member teaching the course.
 * @param {string} location - The location of the course.
 * @param {string} zoomLink - The Zoom link for the course.
 * @param {string[]} daysOfWeek - An array of days when the course occurs.
 */
class Course extends SchedulableEvent {
    constructor(id, name, startDate, endDate, faculty, location, zoomLink, daysOfWeek) {
        super(id, 'COURSE', name, startDate, endDate);
        this.faculty = faculty;
        this.location = location;
        this.zoomLink = zoomLink;
        this.daysOfWeek = daysOfWeek;
    }
}

/**
 * Represents a work shift, extending SchedulableEvent.
 * @class
 * @extends SchedulableEvent
 * @param {string} id - The unique shift ID.
 * @param {string} name - The description of the shift.
 * @param {Date} startDate - The start date and time of the shift.
 * @param {Date} endDate - The end date and time of the shift.
 * @param {string} location - The location of the shift (e.g., 'Tech Hub Desk').
 */
class Shift extends SchedulableEvent {
    constructor(id, name, startDate, endDate, location) {
        super(id, 'SHIFT', name, startDate, endDate);
        this.location = location;
    }
}

/**
 * Represents the assignment of a staff member to an event.
 * @class
 * @param {string} id - The unique identifier for the assignment record.
 * @param {string} staffId - The ID of the assigned staff member.
 * @param {string} eventId - The ID of the assigned event.
 * @param {string} eventType - The type of the assigned event ('COURSE' or 'SHIFT').
 */
class Assignment {
    constructor(id, staffId, eventId, eventType) {
        this.id = id;
        this.staffId = staffId;
        this.eventId = eventId;
        this.eventType = eventType;
    }
}
