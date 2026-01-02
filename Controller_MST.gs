/**
 * -------------------------------------------------------------------
 * CONTROLLER: MST SCHEDULING (View Only)
 * -------------------------------------------------------------------
 */

// --- VIEW MODEL GENERATOR ---

function getMSTViewData(master) {
    // NOTE: This function now delegates to the Scheduling Controller logic
    // to ensure consistency between the "Tech Hub" view and "MST" view.
    // It exists to maintain the contract with the frontend if needed,
    // but relies on the Relational Data Model.
    
    const courseItems = [];
    const courseAssignments = [];
    let mstStaffList = [];
    let debugMsg = "";

    try {
        // 1. Build MST Staff List
        const allStaffObjects = Object.values(master.staffMap);
        mstStaffList = allStaffObjects
            .filter(s => (s.Roles || '').toLowerCase().includes('mst') && s.IsActive !== 'FALSE')
            .map(s => ({ id: s.StaffID, name: s.FullName }))
            .sort((a, b) => a.name.localeCompare(b.name));

        // 2. Fetch Data via Scheduling Logic
        // We reuse the logic from Controller_Scheduling to avoid duplication
        const rosterData = getSchedulingRosterData_refactored();
        
        if (rosterData.success) {
            // Map the data to the format expected by the MST View
            rosterData.data.courseItems.forEach(item => {
                courseItems.push({ id: item.id, name: item.name, type: 'Course' });
            });
            
            // The courseAssignments array from Scheduling is already in the correct format
            // (id, staffName, itemName, etc.)
            rosterData.data.courseAssignments.forEach(assign => {
                courseAssignments.push(assign);
            });
        } else {
            debugMsg = rosterData.error || "Failed to load roster data.";
        }

    } catch (e) {
        debugMsg = "Error in MST View: " + e.message;
    }

    return { courseItems, courseAssignments, mstStaffList, debugMsg };
}

// NOTE: All Write functions (saveNewAssignment, etc.) have been removed 
// and consolidated into Controller_Scheduling.gs to prevent namespace collisions.