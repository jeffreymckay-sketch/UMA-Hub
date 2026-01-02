/**
 * -------------------------------------------------------------------
 * CONTROLLER: MST SCHEDULING (View Only)
 * -------------------------------------------------------------------
 */

// --- VIEW MODEL GENERATOR ---

/**
 * Main entry point for the MST frontend to get all scheduling data.
 * This function now delegates entirely to the Scheduling Controller logic
 * to ensure data consistency across all views.
 * @returns {object} A response object with success status and data or an error.
 */
function getMSTViewData() {
    // This simply acts as a dedicated endpoint for the MST view, 
    // but reuses the exact same logic as the main scheduling view to ensure consistency.
    return getSchedulingData();
}

// NOTE: All Write functions (saveNewAssignment, etc.) have been removed 
// and consolidated into Controller_Scheduling.gs to prevent namespace collisions.
