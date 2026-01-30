/**
 * -------------------------------------------------------------------
 * STAFF MANAGEMENT CONTROLLER
 * Handles: Fetching staff, adding new staff, and updating roles/status
 * -------------------------------------------------------------------
 */

/**
 * Fetches the full staff list for the admin management table.
 */
function api_getFullStaffList() {
  try {
    const sheet = getSheet('Staff_List');
    if (!sheet) throw new Error("Staff_List sheet not found.");

    const data = sheet.getDataRange().getValues();
    const headers = getColumnMap(data[0]);
    
    const staffList = data.slice(1).map((row, index) => {
      return {
        rowNumber: index + 2,
        fullName: row[headers.fullname] || "",
        email: row[headers.staffid] || "",
        roles: row[headers.roles] || "",
        isActive: String(row[headers.isactive]).toUpperCase() === 'TRUE',
        notes: row[headers.notes] || ""
      };
    }).filter(s => s.email !== ""); 

    return { success: true, data: staffList };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Adds a new staff member to the sheet.
 */
function api_addStaffMember(staffObj) {
  try {
    const sheet = getSheet('Staff_List');
    sheet.appendRow([
      staffObj.fullName,
      staffObj.email.trim().toLowerCase(),
      staffObj.roles,
      "TRUE", 
      staffObj.notes || "Added via Staff Hub"
    ]);
    return { success: true, message: "Staff member added successfully!" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Updates an existing staff member's roles or active status.
 */
function api_updateStaffMember(staffObj) {
  try {
    const sheet = getSheet('Staff_List');
    const data = sheet.getDataRange().getValues();
    const headers = getColumnMap(data[0]);
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][headers.staffid]).toLowerCase() === staffObj.email.toLowerCase()) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) throw new Error("Staff member not found.");

    sheet.getRange(rowIndex, headers.fullname + 1).setValue(staffObj.fullName);
    sheet.getRange(rowIndex, headers.roles + 1).setValue(staffObj.roles);
    sheet.getRange(rowIndex, headers.isactive + 1).setValue(staffObj.isActive ? "TRUE" : "FALSE");
    
    return { success: true, message: "Staff member updated." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Updates multiple staff members at once (Bulk Update).
 */
function api_bulkUpdateStaff(staffArray) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // Wait up to 30 seconds for other processes to finish
    
    const sheet = getSheet('Staff_List');
    const data = sheet.getDataRange().getValues();
    const headers = getColumnMap(data[0]);
    
    // Create a map of updates for fast lookup
    const updateMap = {};
    staffArray.forEach(s => {
      updateMap[s.email.toLowerCase()] = s;
    });

    const newData = [data[0]]; // Start with headers

    for (let i = 1; i < data.length; i++) {
      const email = String(data[i][headers.staffid]).toLowerCase();
      const row = data[i];
      
      if (updateMap[email]) {
        const update = updateMap[email];
        row[headers.fullname] = update.fullName;
        row[headers.roles] = update.roles;
        row[headers.isactive] = update.isActive ? "TRUE" : "FALSE";
      }
      newData.push(row);
    }

    sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
    return { success: true, message: `Successfully updated ${staffArray.length} staff members.` };

  } catch (e) {
    return { success: false, message: e.message };
  } finally {
    lock.releaseLock();
  }
}