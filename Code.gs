/**
 * -------------------------------------------------------------------
 * ENTRY POINT & ACCESS CONTROL
 * -------------------------------------------------------------------
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('University Staff Hub')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * -------------------------------------------------------------------
 * ACCESS CONTROL (RBAC)
 * -------------------------------------------------------------------
 */

function api_getAccessControlData() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    const ss = getMasterDataHub(); 
    
    // 1. Get Permissions Matrix
    const permSheet = ss.getSheetByName("Permissions_Matrix");
    const matrix = {};
    
    if (permSheet) {
      const data = permSheet.getDataRange().getValues();
      // Skip header, read existing rows
      for (let i = 1; i < data.length; i++) {
        const page = data[i][0];
        if (page) {
          matrix[page] = {
            Admin: true, // Admin always has access
            Lead: data[i][2] === true,
            Staff: data[i][3] === true
          };
        }
      }
    }

    // 2. Backfill Missing Pages (Auto-discovery from Config)
    // This ensures new pages like 'page-mlt-proctoring' appear in the UI even if not in the sheet yet.
    if (CONFIG.PAGES && Array.isArray(CONFIG.PAGES)) {
      CONFIG.PAGES.forEach(pageId => {
        if (!matrix[pageId]) {
          // Default to OPEN if not defined in sheet
          matrix[pageId] = { Admin: true, Lead: true, Staff: true };
        }
      });
    }

    // 3. Get User Role
    const staffSheet = ss.getSheetByName("Staff_List");
    let userRole = "Staff"; // Default
    
    if (staffSheet) {
      const data = staffSheet.getDataRange().getValues();
      const headers = data[0].map(h => h.toString().toLowerCase());
      const emailIdx = headers.findIndex(h => h.includes("id") || h.includes("email"));
      const roleIdx = headers.findIndex(h => h.includes("role"));

      if (emailIdx > -1 && roleIdx > -1) {
        for (let i = 1; i < data.length; i++) {
          if (data[i][emailIdx].toString().toLowerCase() === userEmail.toLowerCase()) {
            const roleStr = data[i][roleIdx].toString();
            if (roleStr.includes("Admin")) userRole = "Admin";
            else if (roleStr.includes("Lead")) userRole = "Lead";
            break;
          }
        }
      }
    }

    return { 
      userRole: userRole, 
      matrix: matrix,
      email: userEmail
    };

  } catch (e) {
    return { error: e.message };
  }
}

function api_savePermissionsMatrix(newMatrix) {
  try {
    const ss = getMasterDataHub();
    let sheet = ss.getSheetByName("Permissions_Matrix");
    if (!sheet) {
      sheet = ss.insertSheet("Permissions_Matrix");
      sheet.appendRow(["Page Section", "Admin (Locked)", "Lead", "Staff"]);
    }
    
    // Clear existing data (except header)
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();
    }

    const rows = [];
    // Save ALL entries from the incoming matrix (which includes our new defaults)
    for (const [pageId, perms] of Object.entries(newMatrix)) {
      rows.push([pageId, "TRUE", perms.Lead, perms.Staff]);
    }

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * -------------------------------------------------------------------
 * STAFF MANAGEMENT API
 * -------------------------------------------------------------------
 */

function api_getStaffForRoleAssignment() {
  try {
    const ss = getMasterDataHub();
    const sheet = ss.getSheetByName("Staff_List");
    if (!sheet) return { error: "Staff_List tab not found." };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    
    const nameIdx = headers.findIndex(h => h.includes("name"));
    const idIdx = headers.findIndex(h => h.includes("id") || h.includes("email"));
    const roleIdx = headers.findIndex(h => h.includes("role"));

    if (nameIdx === -1 || idIdx === -1 || roleIdx === -1) {
      return { error: "Headers (Name, ID, Role) not found in Staff_List." };
    }

    const staff = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx]) {
        staff.push({
          row: i + 1, // 1-based index
          name: data[i][nameIdx],
          id: data[i][idIdx],
          currentRole: data[i][roleIdx] || "Staff"
        });
      }
    }
    
    return { data: staff, validRoles: ["Admin", "Lead", "Staff"] };
  } catch (e) {
    return { error: e.message };
  }
}

function api_assignStaffRole(staffId, newRole, row) {
  return api_editStaffMember(staffId, null, staffId, row, newRole); 
}

function api_editStaffMember(oldStaffId, newFullName, newStaffId, row, newRoleString) {
  try {
    if (!row) return { success: false, message: "Missing row index." };
    const ss = getMasterDataHub(); 
    const sheet = ss.getSheetByName("Staff_List"); 
    if (!sheet) return { success: false, message: "Staff_List tab not found." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const getColIndex = (search) => {
      const idx = headers.findIndex(h => h.toString().toLowerCase().includes(search.toLowerCase()));
      return idx >= 0 ? idx + 1 : -1;
    };

    const nameCol = getColIndex("Name");
    const idCol = getColIndex("ID"); 
    const roleCol = getColIndex("Role");

    if (nameCol === -1 || idCol === -1 || roleCol === -1) {
      return { success: false, message: "Could not find Name, ID, or Role columns." };
    }

    if (newFullName) sheet.getRange(row, nameCol).setValue(newFullName);
    if (newStaffId) sheet.getRange(row, idCol).setValue(newStaffId);
    
    if (newRoleString !== undefined && newRoleString !== null) {
      sheet.getRange(row, roleCol).setValue(newRoleString);
    }

    return { success: true, message: "Staff updated." };

  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  }
}

function api_staff_create(fullName, staffId, primaryRole) {
  try {
    const ss = getMasterDataHub();
    const sheet = ss.getSheetByName("Staff_List");
    if (!sheet) return { success: false, message: "Staff_List tab missing." };
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase());
    const idIdx = headers.findIndex(h => h.includes("id") || h.includes("email"));
    
    if (idIdx > -1) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][idIdx] === staffId) {
          return { success: false, message: "Staff ID already exists." };
        }
      }
    }
    
    sheet.appendRow([fullName, staffId, primaryRole]);
    return { success: true, message: "Staff member added." };
    
  } catch (e) {
    return { success: false, message: e.message };
  }
}