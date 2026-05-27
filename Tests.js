function DIAGNOSE_runMe() {
  // This function is safe - it only READS data, never changes anything.
  // To run it: click the dropdown at the top that says "Select function" 
  // and choose DIAGNOSE_runMe, then click the Run button (▶)
  
  var results = [];

  // --- Check 1: Can we read the course schedule? ---
  try {
    var mergedData = getMergedMstData();
    results.push("✅ CHECK 1 PASSED: Course schedule loaded. Row count: " + mergedData.length);
    
    // --- Check 2: Are there any problem values hiding in the data? ---
    var problemsFound = 0;
    mergedData.forEach(function(row, rowIndex) {
      for (var key in row) {
        var val = row[key];
        if (val === undefined) {
          results.push("⚠️ PROBLEM FOUND in row " + rowIndex + ", column: " + key + " — value is undefined");
          problemsFound++;
        }
        if (val !== null && typeof val === 'object' && !(val instanceof Date) && typeof val.toISOString !== 'function') {
          results.push("⚠️ PROBLEM FOUND in row " + rowIndex + ", column: " + key + " — unexpected object type: " + typeof val);
          problemsFound++;
        }
      }
    });
    if (problemsFound === 0) results.push("✅ CHECK 2 PASSED: No undefined or bad values found in course data.");

  } catch(e) {
    results.push("❌ CHECK 1 FAILED: Could not read course schedule. Error: " + e.message);
  }

  // --- Check 3: Can we read the Staff List? ---
  try {
    var staffData = getSheet('Staff_List').getDataRange().getValues();
    results.push("✅ CHECK 3 PASSED: Staff list loaded. Row count: " + staffData.length);
  } catch(e) {
    results.push("❌ CHECK 3 FAILED: Could not read Staff_List. Error: " + e.message);
  }

  // --- Check 4: Can we read Staff Assignments? ---
  try {
    var assignData = getSheet('Staff_Assignments').getDataRange().getValues();
    results.push("✅ CHECK 4 PASSED: Staff assignments loaded. Row count: " + assignData.length);
  } catch(e) {
    results.push("❌ CHECK 4 FAILED: Could not read Staff_Assignments. Error: " + e.message);
  }

  // --- Print everything to the log so you can read it ---
  results.forEach(function(line) { console.log(line); });
  console.log("--- DIAGNOSIS COMPLETE ---");
}

function DIAGNOSE_runMe2() {
  // Tests the actual data-building process inside api_getMstViewData
  // Safe - read only, changes nothing.

  var results = [];

  try {
    // Step 1: Build the staff map the same way the real function does
    var staffData = getSheet('Staff_List').getDataRange().getValues();
    var staffHeaders = getColumnMap(staffData[0]);
    results.push("Staff headers found: " + JSON.stringify(staffHeaders));

    var idIdx = staffHeaders['staffid'] !== undefined ? staffHeaders['staffid'] : staffHeaders['id'];
    var nameIdx = staffHeaders['fullname'] !== undefined ? staffHeaders['fullname'] : staffHeaders['name'];
    var roleIdx = staffHeaders['roles'];
    var actIdx = staffHeaders['isactive'];

    results.push("Column indexes — ID: " + idIdx + ", Name: " + nameIdx + ", Role: " + roleIdx + ", Active: " + actIdx);

    if (idIdx === undefined) results.push("❌ PROBLEM: Cannot find a 'StaffID' or 'ID' column in Staff_List");
    if (nameIdx === undefined) results.push("❌ PROBLEM: Cannot find a 'FullName' or 'Name' column in Staff_List");
    if (roleIdx === undefined) results.push("⚠️ WARNING: Cannot find a 'Roles' column in Staff_List");
    if (actIdx === undefined) results.push("⚠️ WARNING: Cannot find an 'IsActive' column in Staff_List");

  } catch(e) {
    results.push("❌ FAILED building staff map: " + e.message);
  }

  try {
    // Step 2: Build the assignment map the same way the real function does
    var assignData = getSheet('Staff_Assignments').getDataRange().getValues();
    var assignHeaders = getColumnMap(assignData[0]);
    results.push("Assignment headers found: " + JSON.stringify(assignHeaders));

    var typeCol = assignHeaders['assignmenttype'];
    var refCol = assignHeaders['referenceid'];
    var staffCol = assignHeaders['staffid'];

    results.push("Assignment column indexes — Type: " + typeCol + ", ReferenceID: " + refCol + ", StaffID: " + staffCol);

    if (typeCol === undefined) results.push("❌ PROBLEM: Cannot find 'AssignmentType' column in Staff_Assignments");
    if (refCol === undefined) results.push("❌ PROBLEM: Cannot find 'ReferenceID' column in Staff_Assignments");
    if (staffCol === undefined) results.push("❌ PROBLEM: Cannot find 'StaffID' column in Staff_Assignments");

  } catch(e) {
    results.push("❌ FAILED building assignment map: " + e.message);
  }

  try {
    // Step 3: Try building the final view objects and check each one for bad values
    var mergedData = getMergedMstData();
    var problemsFound = 0;

    mergedData.forEach(function(courseObj, i) {
      var startD = combineDateAndTime(courseObj.startdate, courseObj.starttime);
      var endD = combineDateAndTime(courseObj.startdate, courseObj.endtime);

      // Check the computed fields that go into the final object
      var testObj = {
        id: String(courseObj.eventid),
        itemName: courseObj.course || "Untitled",
        courseFaculty: courseObj.faculty || "",
        courseDay: courseObj.day || "",
        location: courseObj.bxlocation || "",
        zoomLink: courseObj.zoomlink || "",
        startDateRaw: courseObj.startdate,
        endDateRaw: courseObj.enddate,
        startTimeRaw: courseObj.starttime,
        endTimeRaw: courseObj.endtime,
        startDComputed: startD ? startD.toISOString() : "null",
        endDComputed: endD ? endD.toISOString() : "null"
      };

      for (var key in testObj) {
        if (testObj[key] === undefined) {
          results.push("⚠️ PROBLEM in row " + i + " (" + courseObj.course + "), field '" + key + "' is undefined");
          problemsFound++;
        }
      }
    });

    if (problemsFound === 0) {
      results.push("✅ All " + mergedData.length + " course view objects look clean");
    }

  } catch(e) {
    results.push("❌ FAILED building course view objects: " + e.message);
  }

  results.forEach(function(line) { console.log(line); });
  console.log("--- DIAGNOSIS 2 COMPLETE ---");
}

function DIAGNOSE_runMe3() {
  // Tests the FULL assembly - closest thing to running the real function
  // Safe - read only, changes nothing.

  try {
    var mergedData = getMergedMstData();
    var staffData = getSheet('Staff_List').getDataRange().getValues();
    var assignData = getSheet('Staff_Assignments').getDataRange().getValues();
    var staffHeaders = getColumnMap(staffData[0]);
    var assignHeaders = getColumnMap(assignData[0]);

    var idIdx = staffHeaders['staffid'] !== undefined ? staffHeaders['staffid'] : staffHeaders['id'];
    var nameIdx = staffHeaders['fullname'] !== undefined ? staffHeaders['fullname'] : staffHeaders['name'];
    var roleIdx = staffHeaders['roles'];
    var actIdx = staffHeaders['isactive'];

    // Build staff list exactly as the real function does
    var allStaff = staffData.slice(1).map(function(row) {
      return {
        id: row[idIdx], 
        name: row[nameIdx], 
        role: row[roleIdx] || '', 
        isActive: (String(row[actIdx]).toUpperCase() === 'TRUE' || row[actIdx] === true)
      };
    }).filter(function(s) { return s && s.isActive; });

    console.log("Active staff count: " + allStaff.length);
    allStaff.forEach(function(s, i) {
      console.log("Staff " + i + ": id=" + s.id + " | name=" + s.name + " | role=" + s.role);
    });

    // Build assignment map exactly as the real function does
    var assignmentMap = new Map();
    for (var i = 1; i < assignData.length; i++) {
      if (assignData[i][assignHeaders.assignmenttype] === 'Course') {
        assignmentMap.set(String(assignData[i][assignHeaders.referenceid]), String(assignData[i][assignHeaders.staffid]));
      }
    }
    console.log("Assignments mapped: " + assignmentMap.size);

    // Build staff map
    var staffMap = new Map(allStaff.map(function(s) { return [String(s.id).toLowerCase(), s]; }));

    // Build MST staff list exactly as the real function does
    var mstStaffList = allStaff
      .filter(function(s) { return s.role && s.role.toLowerCase().includes('mst'); })
      .map(function(s) { return { id: s.id, name: s.name }; });

    console.log("MST staff count: " + mstStaffList.length);

    // Build the full course view objects exactly as the real function does
    var courseAssignmentsView = mergedData.map(function(courseObj) {
      var id = String(courseObj.eventid);
      var staffId = assignmentMap.get(id);
      var staff = staffId ? staffMap.get(staffId.toLowerCase()) : null;

      var startD = combineDateAndTime(courseObj.startdate, courseObj.starttime);
      var endD = combineDateAndTime(courseObj.startdate, courseObj.endtime);
      var seriesEndD = new Date(courseObj.enddate);

      var timeDisplay = "TBD";
      if (startD && endD && !isNaN(startD.getTime()) && !isNaN(endD.getTime())) {
        timeDisplay = Utilities.formatDate(startD, Session.getScriptTimeZone(), 'h:mm a') + ' - ' + Utilities.formatDate(endD, Session.getScriptTimeZone(), 'h:mm a');
      }

      var sStr = "?"; if (startD && !isNaN(startD.getTime())) sStr = Utilities.formatDate(startD, Session.getScriptTimeZone(), 'M/d/yy');
      var eStr = "?"; if (seriesEndD && !isNaN(seriesEndD.getTime())) eStr = Utilities.formatDate(seriesEndD, Session.getScriptTimeZone(), 'M/d/yy');

      var safeRaw = {};
      for (var key in courseObj) {
        var val = courseObj[key];
        if (val instanceof Date) val = val.toISOString();
        safeRaw[key] = val;
      }

      return {
        id: id,
        itemName: courseObj.course || "Untitled",
        courseFaculty: courseObj.faculty || "",
        courseDay: courseObj.day || "",
        courseTime: timeDisplay,
        startDateStr: sStr,
        endDateStr: eStr,
        location: courseObj.bxlocation || "",
        zoomLink: courseObj.zoomlink || "",
        staffName: staff ? staff.name : "Unassigned",
        staffId: staff ? staff.id : null,
        raw: safeRaw
      };
    });

    // Now try to JSON stringify the FULL response - this is what Apps Script does when sending to browser
    var fullResponse = { success: true, data: { courseAssignments: courseAssignmentsView, mstStaffList: mstStaffList } };
    
    try {
      var jsonString = JSON.stringify(fullResponse);
      console.log("✅ JSON serialization PASSED. Total characters: " + jsonString.length);
    } catch(jsonErr) {
      console.log("❌ JSON serialization FAILED: " + jsonErr.message);
    }

    console.log("--- DIAGNOSIS 3 COMPLETE ---");

  } catch(e) {
    console.log("❌ ASSEMBLY FAILED: " + e.message);
    console.log("Stack: " + e.stack);
  }
}