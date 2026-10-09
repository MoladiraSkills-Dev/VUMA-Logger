const LOCATIONIQ_API_KEY = "pk.1be07ee2080691339d8fc4f1712dbc95";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('VUMA Logger - Workforce Attendance Portal')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ==========================================================================
   DATABASE INITIALIZATION & TAB GUARD ENGINE
   ========================================================================== */

function getOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (headers && headers.length > 0) {
      sheet.appendRow(headers);
    }
  }
  return sheet;
}

/* ==========================================================================
   AUTHENTICATION & SECURITY CORE
   ========================================================================== */

function generateSalt(length = 16) {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789./';
  let salt = '';
  for (let i = 0; i < length; i++) {
    salt += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return salt;
}

function computeSHA256(password, salt) {
  const combined = password + salt;
  const signature = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combined, Utilities.Charset.UTF_8);
  let hashStr = '';
  for (let i = 0; i < signature.length; i++) {
    let byteVal = signature[i];
    if (byteVal < 0) byteVal += 256;
    let byteString = byteVal.toString(16);
    if (byteString.length == 1) byteString = '0' + byteString;
    hashStr += byteString;
  }
  return hashStr;
}

function registerUser(email, firstName, lastName, password) {
  try {
    const sheet = getOrCreateSheet("Users", ["Email", "First Name", "Last Name", "Password Hash", "Salt", "Account Type", "Status", "Verification Token"]);
    const data = sheet.getDataRange().getValues();
    const cleanedEmail = email.trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === cleanedEmail) {
        return { status: "error", message: "Email is already registered." };
      }
    }
    
    const salt = generateSalt();
    const passwordHash = computeSHA256(password, salt);
    const accountType = "Employee";
    const status = "Verified";
    const token = Utilities.getUuid();
    
    sheet.appendRow([cleanedEmail, firstName.trim(), lastName.trim(), passwordHash, salt, accountType, status, token]);
    return { status: "success", message: "Registration successful! You can now log in." };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

function loginUser(email, password) {
  try {
    const sheet = getOrCreateSheet("Users", ["Email", "First Name", "Last Name", "Password Hash", "Salt", "Account Type", "Status", "Verification Token"]);
    const data = sheet.getDataRange().getValues();
    const cleanedEmail = email.trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === cleanedEmail) {
        const firstName = data[i][1];
        const lastName = data[i][2];
        const savedHash = data[i][3];
        const salt = data[i][4];
        const accountType = data[i][5];
        const status = data[i][6];
        
        if (status === "Blocked") {
          return { status: "error", message: "Your account has been blocked. Please contact an admin." };
        }
        
        const computedHash = computeSHA256(password, salt);
        if (computedHash === savedHash) {
          return {
            status: "success",
            accountType: accountType,
            user: { email: cleanedEmail, firstName: firstName, lastName: lastName }
          };
        } else {
          return { status: "error", message: "Invalid password credentials." };
        }
      }
    }
    return { status: "error", message: "Account email not found." };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

/* ==========================================================================
   TIME & SHIFT CALCULATOR ENGINE (Standardized Floor Override)
   ========================================================================== */

function parseTimeToMinutes(timeStr) {
  if (!timeStr || timeStr === "--" || String(timeStr).includes("1899")) return null;
  const cleanStr = String(timeStr).trim().toUpperCase();
  const isPM = cleanStr.includes("PM");
  const isAM = cleanStr.includes("AM");
  
  const matches = cleanStr.match(/\d+/g);
  if (!matches || matches.length < 2) return null;
  
  let hours = parseInt(matches[0], 10);
  let minutes = parseInt(matches[1], 10);
  let seconds = matches.length >= 3 ? parseInt(matches[2], 10) : 0;
  
  if (isNaN(hours) || isNaN(minutes)) return null;
  if (isPM && hours < 12) hours += 12;
  if (isAM && hours === 12) hours = 0;
  
  return (hours * 60) + minutes + (seconds / 60);
}

function formatMinutesToTimeString(totalMinutes) {
  const hrs = Math.floor(totalMinutes / 60);
  const mins = Math.floor(totalMinutes % 60);
  const secs = Math.round((totalMinutes % 1) * 60);
  return `${String(hrs).padStart(2, '0')}:${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
}

function parseDurationStringToMinutes(durationStr) {
  if (!durationStr || durationStr === "N/A" || durationStr === "--" || durationStr === "In Progress" || durationStr === "No Clock In") return 0;
  let totalMins = 0;
  const hrMatch = String(durationStr).match(/(\d+)\s*hr/i);
  const minMatch = String(durationStr).match(/(\d+)\s*min/i);
  if (hrMatch) totalMins += parseInt(hrMatch[1], 10) * 60;
  if (minMatch) totalMins += parseInt(minMatch[1], 10);
  return totalMins;
}

function formatMinutesToDurationString(totalMinutes) {
  if (totalMinutes <= 0) return "0 mins";
  const hrs = Math.floor(totalMinutes / 60);
  const mins = Math.round(totalMinutes % 60);
  if (hrs === 0) return `${mins} mins`;
  return `${hrs} hr${hrs > 1 ? 's' : ''} ${mins} min${mins > 1 ? 's' : ''}`;
}

function formatMinutesToDecimalHours(totalMinutes) {
  if (totalMinutes <= 0) return "0";
  return (totalMinutes / 60).toFixed(2);
}

/**
 * Enhanced calculation engine. 
 * If NOT admin edit: locks to 08:00 and 17:00.
 * If Admin edit: respects the true custom times and precisely splits OT from Regular.
 */
function calculateHoursAndOvertime(startTimeStr, endTimeStr, dateStr, isAdminEdit = false) {
  try {
    const rawStartMinutes = parseTimeToMinutes(startTimeStr);
    const endMinutes = parseTimeToMinutes(endTimeStr); 
    
    if (rawStartMinutes === null || endMinutes === null) {
      return { floorIn: "--", floorOut: "--", totalHours: "N/A", overtime: "0 mins" };
    }
    
    const dateObj = new Date(dateStr);
    const isSaturday = dateObj.getDay() === 6;
    
    const FLOOR_START_MINS = 480; // 08:00 AM Override
    let FLOOR_END_MINS = 1020;    // 17:00 PM Override
    let LUNCH_DEDUCTION_MINS = 60;   
    
    if (isSaturday) {
      FLOOR_END_MINS = 780;       // 13:00 PM Override
      LUNCH_DEDUCTION_MINS = 0;     
    }
    
    // Dynamic floors based on Admin Edit status
    const effectiveStartMinutes = isAdminEdit ? rawStartMinutes : FLOOR_START_MINS;
    const floorInStr = formatMinutesToTimeString(effectiveStartMinutes);
    
    const effectiveEndMinutes = isAdminEdit ? endMinutes : FLOOR_END_MINS;
    const floorOutStr = formatMinutesToTimeString(effectiveEndMinutes);
    
    let standardWorkMinutes = 0;
    let otMinutes = 0;

    if (isAdminEdit) {
      if (endMinutes > FLOOR_END_MINS) {
         otMinutes = endMinutes - FLOOR_END_MINS;
         standardWorkMinutes = FLOOR_END_MINS - effectiveStartMinutes;
      } else {
         otMinutes = 0;
         standardWorkMinutes = endMinutes - effectiveStartMinutes;
      }
    } else {
      standardWorkMinutes = FLOOR_END_MINS - FLOOR_START_MINS;
      otMinutes = 0;
    }
    
    let netWorkMinutes = standardWorkMinutes > LUNCH_DEDUCTION_MINS ? standardWorkMinutes - LUNCH_DEDUCTION_MINS : standardWorkMinutes;
    if (netWorkMinutes < 0) netWorkMinutes = 0; 
    
    const totalHoursStr = formatMinutesToDurationString(netWorkMinutes);
    const overtimeStr = formatMinutesToDurationString(otMinutes);
    
    return { floorIn: floorInStr, floorOut: floorOutStr, totalHours: totalHoursStr, overtime: overtimeStr };
  } catch(e) {
    return { floorIn: "--", floorOut: "--", totalHours: "N/A", overtime: "0 mins" };
  }
}

/* ==========================================================================
   AUTO-LOGOUT & GEOLOCATION TRACKING ENGINE
   ========================================================================== */

function autoClosePreviousShifts(email, firstName, lastName, currentDateStr) {
  const summarySheet = getOrCreateSheet("Daily_Summaries", ["First Name", "Last Name", "Summary ID", "Date", "Email", "Clock In Time", "Clock Out Time", "Clock-In Floor Time", "Clock-Out Floor Time", "Total Hours worked", "Status", "Worked Overtime"]);
  const data = summarySheet.getDataRange().getDisplayValues();
  const logSheet = getOrCreateSheet("Activity_Logs", ["Log ID", "Email", "First Name", "Last Name", "Date", "Time", "Action", "Latitude", "Longitude", "Resolved Location Name", "Edited By"]);
  
  const targetEmail = email.trim().toLowerCase();
  
  for (let i = 1; i < data.length; i++) {
    const rowDateStr = String(data[i][3]).trim();
    const rowEmail = String(data[i][4]).trim().toLowerCase();
    const rowStatus = String(data[i][10]).trim(); 
    
    if (rowEmail === targetEmail && rowStatus === "In Progress" && rowDateStr !== "" && rowDateStr !== currentDateStr) {
      const isSat = new Date(rowDateStr).getDay() === 6;
      const autoOutTime = isSat ? "13:00:00" : "17:00:00";
      
      const logId = "LOG-" + Math.floor(100000 + Math.random() * 900000);
      logSheet.appendRow([logId, targetEmail, firstName, lastName, rowDateStr, autoOutTime, "Clock Out", "0", "0", "System Auto-Logout", "System Automation"]);
      
      const wasUpdated = updateDailySummaryTable(targetEmail, firstName, lastName, rowDateStr, autoOutTime, "Clock Out");
      
      if (wasUpdated) {
        const weekKey = getISOWeekKey(new Date(rowDateStr));
        syncWeeklySummariesForUser(targetEmail, weekKey);
        syncMonthlySummariesForUser(targetEmail, rowDateStr.substring(0, 7));
      }
    }
  }
}

function processActivityLog(userEmail, firstName, lastName, action, lat, lon) {
  try {
    const logSheet = getOrCreateSheet("Activity_Logs", ["Log ID", "Email", "First Name", "Last Name", "Date", "Time", "Action", "Latitude", "Longitude", "Resolved Location Name", "Edited By"]);
    
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "HH:mm:ss");
    
    const logData = logSheet.getDataRange().getDisplayValues(); 
    const targetEmail = userEmail.trim().toLowerCase();
    
    for (let i = 1; i < logData.length; i++) {
      const rowEmail = String(logData[i][1]).trim().toLowerCase();
      const rowDateStr = String(logData[i][4]).trim();
      const rowAction = String(logData[i][6]).trim();
      
      if (rowEmail === targetEmail && rowDateStr === dateStr) {
        if (rowAction === action) {
          return { status: "error", message: `Action Blocked: You have already completed your "${action}" for today (${dateStr}).` };
        }
      }
    }

    if (action === "Clock In") {
      autoClosePreviousShifts(targetEmail, firstName, lastName, dateStr);
    }
    
    let locationName = "Unknown/Unresolved Location";
    if (lat && lon) {
      const url = `https://us1.locationiq.com/v1/reverse?key=${LOCATIONIQ_API_KEY}&lat=${lat}&lon=${lon}&format=json`;
      try {
        const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        if (response.getResponseCode() === 200) {
          const parsed = JSON.parse(response.getContentText());
          locationName = parsed.display_name || locationName;
        }
      } catch(apiErr) {
        locationName = "Location IQ API Error / Limit reached";
      }
    }
    
    const logId = "LOG-" + Math.floor(100000 + Math.random() * 900000);
    logSheet.appendRow([logId, targetEmail, firstName, lastName, dateStr, timeStr, action, lat, lon, locationName, ""]);
    
    const wasUpdated = updateDailySummaryTable(targetEmail, firstName, lastName, dateStr, timeStr, action);
    if(wasUpdated) {
      const weekKey = getISOWeekKey(new Date(dateStr));
      syncWeeklySummariesForUser(targetEmail, weekKey);
      syncMonthlySummariesForUser(targetEmail, dateStr.substring(0, 7));
    }
    
    return { status: "success", action: action, location: locationName, time: timeStr };
  } catch(e) {
    return { status: "error", message: "Server Error: " + e.toString() };
  }
}

function updateDailySummaryTable(email, firstName, lastName, dateStr, timeStr, action) {
  const summarySheet = getOrCreateSheet("Daily_Summaries", ["First Name", "Last Name", "Summary ID", "Date", "Email", "Clock In Time", "Clock Out Time", "Clock-In Floor Time", "Clock-Out Floor Time", "Total Hours worked", "Status", "Worked Overtime"]);
  const data = summarySheet.getDataRange().getDisplayValues();
  let existingRow = -1;
  let currentStatus = "";
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim() === dateStr && String(data[i][4]).trim().toLowerCase() === email) {
      existingRow = i + 1;
      currentStatus = String(data[i][10]).trim(); 
      break;
    }
  }
  
  if (currentStatus === "Completed") return false; 
  
  if (action === "Clock In") {
    if (existingRow !== -1) {
      summarySheet.getRange(existingRow, 6).setValue(timeStr);
      summarySheet.getRange(existingRow, 8).setValue("08:00:00");
      summarySheet.getRange(existingRow, 11).setValue("In Progress");
    } else {
      const summaryId = "SUM-" + Math.floor(100000 + Math.random() * 900000);
      summarySheet.appendRow([firstName, lastName, summaryId, dateStr, email, timeStr, "--", "08:00:00", "--", "In Progress", "In Progress", "0 mins"]);
    }
    return false; 
  } else if (action === "Clock Out") {
    if (existingRow !== -1) {
      const clockInTime = summarySheet.getRange(existingRow, 6).getValue();
      let floorIn = "--"; let floorOut = "--";
      let totalHours = "N/A"; let overtime = "0 mins";
      
      if (clockInTime && clockInTime !== "--") {
        const calculated = calculateHoursAndOvertime(String(clockInTime), timeStr, dateStr, false);
        floorIn = calculated.floorIn;
        floorOut = calculated.floorOut;
        totalHours = calculated.totalHours;
        overtime = calculated.overtime;
      }
      
      summarySheet.getRange(existingRow, 7).setValue(timeStr);
      summarySheet.getRange(existingRow, 8).setValue(floorIn);
      summarySheet.getRange(existingRow, 9).setValue(floorOut);
      summarySheet.getRange(existingRow, 10).setValue(totalHours);
      summarySheet.getRange(existingRow, 11).setValue("Completed");
      summarySheet.getRange(existingRow, 12).setValue(overtime);
    } else {
      const summaryId = "SUM-" + Math.floor(100000 + Math.random() * 900000);
      summarySheet.appendRow([firstName, lastName, summaryId, dateStr, email, "--", timeStr, "--", "--", "No Clock In", "Completed", "N/A"]);
    }
    return true; 
  }
  return false;
}

/* ==========================================================================
   ADMIN DAILY SUMMARY EDITOR & DELETION ENGINE
   ========================================================================== */

function adminUpdateDailySummary(summaryId, newClockIn, newClockOut, adminEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = ss.getSheetByName("Daily_Summaries");
    if (!summarySheet) return { status: "error", message: "Daily_Summaries sheet missing." };
    
    const data = summarySheet.getDataRange().getDisplayValues();
    let targetRow = -1;
    let email = "";
    let dateStr = "";
    let firstName = "";
    let lastName = "";
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][2]).trim() === String(summaryId).trim()) {
        targetRow = i + 1;
        firstName = String(data[i][0]).trim();
        lastName = String(data[i][1]).trim();
        dateStr = String(data[i][3]).trim();
        email = String(data[i][4]).trim().toLowerCase();
        break;
      }
    }
    if (targetRow === -1) return { status: "error", message: "Summary record not found." };
    
    const logSheet = getOrCreateSheet("Activity_Logs", ["Log ID", "Email", "First Name", "Last Name", "Date", "Time", "Action", "Latitude", "Longitude", "Resolved Location Name", "Edited By"]);
    const logData = logSheet.getDataRange().getDisplayValues();
    let foundOutLog = false;
    
    for (let j = 1; j < logData.length; j++) {
      const logEmail = String(logData[j][1]).trim().toLowerCase();
      const logDate = String(logData[j][4]).trim();
      const logAction = String(logData[j][6]).trim();
      
      if (logEmail === email && logDate === dateStr) {
        if (logAction === "Clock In" && newClockIn) {
          logSheet.getRange(j + 1, 6).setValue(newClockIn);
          logSheet.getRange(j + 1, 11).setValue(adminEmail);
        }
        if (logAction === "Clock Out" && newClockOut) {
          logSheet.getRange(j + 1, 6).setValue(newClockOut);
          logSheet.getRange(j + 1, 11).setValue(adminEmail); 
          foundOutLog = true;
        }
      }
    }
    
    if (newClockOut && newClockOut !== "--" && !foundOutLog) {
      const logId = "LOG-" + Math.floor(100000 + Math.random() * 900000);
      logSheet.appendRow([logId, email, firstName, lastName, dateStr, newClockOut, "Clock Out", "0", "0", "Manual Admin Edit", adminEmail]);
    }

    const finalIn = newClockIn && newClockIn.trim() !== "" ? newClockIn : "--";
    const finalOut = newClockOut && newClockOut.trim() !== "" ? newClockOut : "--";
    
    let floorIn = "--"; let floorOut = "--";
    let totalHours = "N/A"; let overtime = "0 mins";
    let status = "In Progress";
    
    if (finalIn !== "--" && finalOut !== "--") {
       const calc = calculateHoursAndOvertime(finalIn, finalOut, dateStr, true);
       floorIn = calc.floorIn;
       floorOut = calc.floorOut;
       totalHours = calc.totalHours;
       overtime = calc.overtime;
       status = "Completed";
    } else if (finalIn === "--" && finalOut !== "--") {
       totalHours = "No Clock In";
       status = "Completed";
    } else if (finalIn !== "--") {
       floorIn = finalIn;
    }
    
    summarySheet.getRange(targetRow, 6).setValue(finalIn);
    summarySheet.getRange(targetRow, 7).setValue(finalOut);
    summarySheet.getRange(targetRow, 8).setValue(floorIn);
    summarySheet.getRange(targetRow, 9).setValue(floorOut);
    summarySheet.getRange(targetRow, 10).setValue(totalHours);
    summarySheet.getRange(targetRow, 11).setValue(status);
    summarySheet.getRange(targetRow, 12).setValue(overtime);
    
    if (email && dateStr) {
      const weekKey = getISOWeekKey(new Date(dateStr));
      syncWeeklySummariesForUser(email, weekKey);
      syncMonthlySummariesForUser(email, dateStr.substring(0, 7));
    }
    return { status: "success", message: "Daily summary updated. Week & Month totals automatically recalculated." };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

function adminDeleteDailySummary(summaryId, dateStr, email, adminEmail) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = ss.getSheetByName("Daily_Summaries");
    const logSheet = ss.getSheetByName("Activity_Logs");

    const sumData = summarySheet.getDataRange().getDisplayValues();
    let sumRowDeleted = false;
    
    for (let i = sumData.length - 1; i >= 1; i--) {
      if (String(sumData[i][2]).trim() === String(summaryId).trim()) {
        summarySheet.deleteRow(i + 1);
        sumRowDeleted = true;
        break;
      }
    }

    if (!sumRowDeleted) return { status: "error", message: "Summary record could not be found." };

    const targetEmail = email.trim().toLowerCase();
    const logData = logSheet.getDataRange().getDisplayValues();
    
    for (let j = logData.length - 1; j >= 1; j--) {
      const lEmail = String(logData[j][1]).trim().toLowerCase();
      const lDate = String(logData[j][4]).trim();
      if (lEmail === targetEmail && lDate === dateStr) {
        logSheet.deleteRow(j + 1);
      }
    }

    if (targetEmail && dateStr) {
      const weekKey = getISOWeekKey(new Date(dateStr));
      syncWeeklySummariesForUser(targetEmail, weekKey);
      syncMonthlySummariesForUser(targetEmail, dateStr.substring(0, 7));
    }

    return { status: "success", message: "Log explicitly deleted." };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

/* ==========================================================================
   WEEKLY & MONTHLY SUMMARIES ENGINE
   ========================================================================== */

function getISOWeekKey(d) {
  const date = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dayNum = date.getUTCDay() || 7;
  date.setUTCDate(date.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(date.getUTCFullYear(), 0, 1));
  const weekNo = Math.ceil((((date - yearStart) / 86400000) + 1) / 7);
  return `${date.getUTCFullYear()}-W${String(weekNo).padStart(2, '0')}`;
}

function syncWeeklySummariesForUser(userEmail, targetWeekKey) {
  const weeklySheet = getOrCreateSheet("Weekly_Summaries", ["First Name", "Last Name", "Email", "Week Identifier", "Total Hours Worked", "Total Overtime", "Days Worked", "Compliance Status", "Last Updated"]);
  const dailySheet = getOrCreateSheet("Daily_Summaries", ["First Name", "Last Name", "Summary ID", "Date", "Email", "Clock In Time", "Clock Out Time", "Clock-In Floor Time", "Clock-Out Floor Time", "Total Hours worked", "Status", "Worked Overtime"]);
  
  const dailyData = dailySheet.getDataRange().getDisplayValues();
  let totalRegularMins = 0; let totalOvertimeMins = 0; let daysWorkedCount = 0;
  let firstName = ""; let lastName = "";
  const targetEmail = userEmail.trim().toLowerCase();
  
  for (let i = 1; i < dailyData.length; i++) {
    const rowFirstName = String(dailyData[i][0] || "").trim();
    const rowLastName = String(dailyData[i][1] || "").trim();
    const rowDateStr = String(dailyData[i][3] || "").trim();
    const rowEmail = String(dailyData[i][4] || "").trim().toLowerCase();
    const rowTotalHoursStr = String(dailyData[i][9] || "").trim(); 
    const rowOvertimeStr = String(dailyData[i][11] || "").trim();  
    
    if (rowEmail === targetEmail && rowDateStr) {
      const rowWeekKey = getISOWeekKey(new Date(rowDateStr));
      if (rowWeekKey === targetWeekKey) {
        if (!firstName && rowFirstName) firstName = rowFirstName;
        if (!lastName && rowLastName) lastName = rowLastName;
        const regMins = parseDurationStringToMinutes(rowTotalHoursStr);
        const otMins = parseDurationStringToMinutes(rowOvertimeStr);
        if (regMins > 0 || otMins > 0 || dailyData[i][10] === "Completed") { 
          totalRegularMins += regMins; totalOvertimeMins += otMins; daysWorkedCount++;
        }
      }
    }
  }
  
  if (!firstName || !lastName) {
    const userSheet = getOrCreateSheet("Users", ["Email", "First Name", "Last Name", "Password Hash", "Salt", "Account Type", "Status", "Verification Token"]);
    const uData = userSheet.getDataRange().getDisplayValues();
    for (let u = 1; u < uData.length; u++) {
      if (String(uData[u][0]).trim().toLowerCase() === targetEmail) {
        firstName = uData[u][1];
        lastName = uData[u][2];
        break;
      }
    }
  }
  
  const totalCombinedMins = totalRegularMins + totalOvertimeMins;
  const totalHoursFloat = totalCombinedMins / 60;
  
  let complianceStatus = "Under Target (<45 hrs)";
  if (totalHoursFloat >= 55.0) complianceStatus = "🔴 MAX CAP ALERT (55+ hrs)";
  else if (totalHoursFloat >= 45.0) complianceStatus = "Target Met (Compliant)";
  
  const weeklyData = weeklySheet.getDataRange().getDisplayValues();
  let existingRow = -1;
  for (let w = 1; w < weeklyData.length; w++) {
    if (String(weeklyData[w][2]).trim().toLowerCase() === targetEmail && String(weeklyData[w][3]).trim() === targetWeekKey) {
      existingRow = w + 1; break;
    }
  }
  
  const formattedHoursWorked = formatMinutesToDurationString(totalRegularMins);
  const formattedOvertime = formatMinutesToDurationString(totalOvertimeMins);
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  
  if (existingRow !== -1) {
    weeklySheet.getRange(existingRow, 5, 1, 5).setValues([[formattedHoursWorked, formattedOvertime, daysWorkedCount, complianceStatus, nowStr]]);
  } else {
    weeklySheet.appendRow([firstName, lastName, targetEmail, targetWeekKey, formattedHoursWorked, formattedOvertime, daysWorkedCount, complianceStatus, nowStr]);
  }
}

function syncMonthlySummariesForUser(userEmail, monthStr) {
  const monthlySheet = getOrCreateSheet("Monthly_Summaries", ["First Name", "Last Name", "Email", "Month", "Total Hours Worked", "Total Overtime", "Days Worked", "Last Updated"]);
  const dailySheet = getOrCreateSheet("Daily_Summaries", ["First Name", "Last Name", "Summary ID", "Date", "Email", "Clock In Time", "Clock Out Time", "Clock-In Floor Time", "Clock-Out Floor Time", "Total Hours worked", "Status", "Worked Overtime"]);
  
  const dailyData = dailySheet.getDataRange().getDisplayValues();
  let totalRegularMins = 0; let totalOvertimeMins = 0; let daysWorkedCount = 0;
  let firstName = ""; let lastName = "";
  const targetEmail = userEmail.trim().toLowerCase();
  
  for (let i = 1; i < dailyData.length; i++) {
    const rowFirstName = String(dailyData[i][0] || "").trim();
    const rowLastName = String(dailyData[i][1] || "").trim();
    const rowDate = String(dailyData[i][3] || "").trim();
    const rowEmail = String(dailyData[i][4] || "").trim().toLowerCase();
    const rowTotalHoursStr = String(dailyData[i][9] || "").trim(); 
    const rowOvertimeStr = String(dailyData[i][11] || "").trim(); 
    
    if (rowEmail === targetEmail && rowDate.startsWith(monthStr)) {
      if (!firstName && rowFirstName) firstName = rowFirstName;
      if (!lastName && rowLastName) lastName = rowLastName;
      const regMins = parseDurationStringToMinutes(rowTotalHoursStr);
      const otMins = parseDurationStringToMinutes(rowOvertimeStr);
      if (regMins > 0 || otMins > 0 || dailyData[i][10] === "Completed") {
        totalRegularMins += regMins; totalOvertimeMins += otMins; daysWorkedCount++;
      }
    }
  }
  
  if (!firstName || !lastName) return;
  
  const monthlyData = monthlySheet.getDataRange().getDisplayValues();
  let existingRow = -1;
  for (let m = 1; m < monthlyData.length; m++) {
    if (String(monthlyData[m][2]).trim().toLowerCase() === targetEmail && String(monthlyData[m][3]).trim() === monthStr) {
      existingRow = m + 1; break;
    }
  }
  
  const formattedHoursWorked = formatMinutesToDurationString(totalRegularMins);
  const formattedOvertime = formatMinutesToDurationString(totalOvertimeMins);
  const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  
  if (existingRow !== -1) {
    monthlySheet.getRange(existingRow, 5, 1, 4).setValues([[formattedHoursWorked, formattedOvertime, daysWorkedCount, nowStr]]);
  } else {
    monthlySheet.appendRow([firstName, lastName, targetEmail, monthStr, formattedHoursWorked, formattedOvertime, daysWorkedCount, nowStr]);
  }
}

/* ==========================================================================
   LEAVE REQUEST ENGINE
   ========================================================================== */

function submitLeaveRequest(email, firstName, lastName, leaveType, startDate, endDate) {
  try {
    const leaveSheet = getOrCreateSheet("Leave_Requests", ["Request ID", "Email", "First Name", "Last Name", "Leave Type", "Start Date", "End Date", "Status", "Submitted Date"]);
    const reqId = "LR-" + Math.floor(100000 + Math.random() * 900000);
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    leaveSheet.appendRow([reqId, email.trim().toLowerCase(), firstName.trim(), lastName.trim(), leaveType, startDate, endDate, "Pending", nowStr]);
    
    return { status: "success", message: "Leave request submitted successfully. Awaiting admin approval." };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

function adminUpdateLeaveStatus(reqId, newStatus) {
  try {
    const leaveSheet = getOrCreateSheet("Leave_Requests", ["Request ID", "Email", "First Name", "Last Name", "Leave Type", "Start Date", "End Date", "Status", "Submitted Date"]);
    const data = leaveSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(reqId).trim()) {
        leaveSheet.getRange(i + 1, 8).setValue(newStatus);
        return { status: "success", message: `Leave request ${newStatus.toLowerCase()}.` };
      }
    }
    return { status: "error", message: "Leave request not found." };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

/* ==========================================================================
   TIMESHEET EXPORT ENGINE (PDF & CSV)
   ========================================================================== */

function fetchTimesheetData(email, startDateStr, endDateStr) {
  const targetEmail = email.trim().toLowerCase();
  
  const userSheet = getOrCreateSheet("Users", []);
  const userData = userSheet.getDataRange().getDisplayValues();
  let employeeName = "Unknown Employee";
  for (let i = 1; i < userData.length; i++) {
    if (String(userData[i][0]).trim().toLowerCase() === targetEmail) {
      employeeName = `${String(userData[i][1]).trim()} ${String(userData[i][2]).trim()}`;
      break;
    }
  }

  const logSheet = getOrCreateSheet("Activity_Logs", []);
  const logData = logSheet.getDataRange().getDisplayValues();
  const adminEditedDates = new Set();
  
  for (let i = 1; i < logData.length; i++) {
    const rowEmail = String(logData[i][1]).trim().toLowerCase();
    const rowDate = String(logData[i][4]).trim();
    const editedBy = String(logData[i][10]).trim(); 
    
    if (rowEmail === targetEmail && editedBy !== "" && editedBy !== "System Automation") {
      adminEditedDates.add(rowDate);
    }
  }

  const startDt = new Date(startDateStr);
  const endDt = new Date(endDateStr);
  const dateMap = {};
  let currentDt = new Date(startDt);
  
  while (currentDt <= endDt) {
    const isoDate = Utilities.formatDate(currentDt, Session.getScriptTimeZone(), "yyyy-MM-dd");
    const dayName = Utilities.formatDate(currentDt, Session.getScriptTimeZone(), "EEEE");
    dateMap[isoDate] = {
      date: isoDate, day: dayName, start: "", lunch: "", end: "", regMins: 0, otMins: 0, sick: "", annual: "", family: "", other: "", isLeave: false
    };
    currentDt.setDate(currentDt.getDate() + 1);
  }

  const leaveSheet = getOrCreateSheet("Leave_Requests", []);
  const leaveData = leaveSheet.getDataRange().getDisplayValues();
  for (let i = 1; i < leaveData.length; i++) {
    const lEmail = String(leaveData[i][1]).trim().toLowerCase();
    const lType = String(leaveData[i][4]).trim();
    const lStart = new Date(String(leaveData[i][5]).trim());
    const lEnd = new Date(String(leaveData[i][6]).trim());
    const lStatus = String(leaveData[i][7]).trim();

    if (lEmail === targetEmail && lStatus === "Approved") {
      let lCurr = new Date(lStart);
      while (lCurr <= lEnd) {
        const lIso = Utilities.formatDate(lCurr, Session.getScriptTimeZone(), "yyyy-MM-dd");
        if (dateMap[lIso]) {
          dateMap[lIso].isLeave = true; dateMap[lIso].regMins = 0; dateMap[lIso].otMins = 0;
          if (lType === "Sick") dateMap[lIso].sick = "SL";
          else if (lType === "Annual") dateMap[lIso].annual = "AL";
          else if (lType === "Family Responsibility") dateMap[lIso].family = "FL";
          else dateMap[lIso].other = lType;
        }
        lCurr.setDate(lCurr.getDate() + 1);
      }
    }
  }

  const formatTimeForPDF = (t) => {
    if (!t || t === "--" || t === "N/A") return "";
    const p = t.split(":");
    return p.length >= 2 ? p[0].padStart(2, '0') + ":" + p[1].padStart(2, '0') : t.substring(0,5);
  };

  const summarySheet = getOrCreateSheet("Daily_Summaries", []);
  const sumData = summarySheet.getDataRange().getDisplayValues();
  for (let i = 1; i < sumData.length; i++) {
    const sDate = String(sumData[i][3]).trim();
    const sEmail = String(sumData[i][4]).trim().toLowerCase();
    
    if (sEmail === targetEmail && dateMap[sDate] && !dateMap[sDate].isLeave) {
      const sFloorIn = String(sumData[i][7]).trim(); 
      const sFloorOut = String(sumData[i][8]).trim(); 
      const sReg = String(sumData[i][9]).trim();
      const sOt = String(sumData[i][11]).trim();
      
      if (adminEditedDates.has(sDate)) {
        dateMap[sDate].start = formatTimeForPDF(sFloorIn);
        dateMap[sDate].end = formatTimeForPDF(sFloorOut);
        dateMap[sDate].otMins = parseDurationStringToMinutes(sOt);
      } else {
        const dObj = new Date(sDate);
        const isSat = dObj.getDay() === 6;
        dateMap[sDate].start = "08:00";
        dateMap[sDate].end = isSat ? "13:00" : "17:00";
        dateMap[sDate].otMins = 0;
      }
      
      if (dateMap[sDate].start && dateMap[sDate].end) {
        const dObj = new Date(sDate);
        dateMap[sDate].lunch = (dObj.getDay() === 6 || dObj.getDay() === 0) ? "" : "1 hour";
      }
      
      dateMap[sDate].regMins = parseDurationStringToMinutes(sReg);
    }
  }

  const rows = Object.values(dateMap).sort((a, b) => new Date(a.date) - new Date(b.date));
  let finalRows = [];
  let currentWeekKey = "";
  let weekRegMins = 0;
  
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    const wKey = getISOWeekKey(new Date(r.date));
    
    if (currentWeekKey !== "" && wKey !== currentWeekKey) {
       finalRows.push({ isTotal: true, label: "Total Weekly Hours", val: formatMinutesToDecimalHours(weekRegMins) + " Hours" });
       weekRegMins = 0;
    }
    
    currentWeekKey = wKey;
    weekRegMins += r.regMins;
    const dStr = Utilities.formatDate(new Date(r.date), Session.getScriptTimeZone(), "dd MMM yyyy");
    
    let roundedOtHours = Math.round(r.otMins / 60);
    
    finalRows.push({
      isTotal: false, date: dStr, day: r.day, start: r.start, lunch: r.lunch, end: r.end,
      regStr: r.regMins > 0 ? formatMinutesToDecimalHours(r.regMins) + " hrs" : "",
      otStr: roundedOtHours > 0 ? roundedOtHours + " hrs" : "",
      sick: r.sick, annual: r.annual, family: r.family, other: r.other
    });
    
    if (i === rows.length - 1) {
       finalRows.push({ isTotal: true, label: "Total Weekly Hours", val: formatMinutesToDecimalHours(weekRegMins) + " Hours" });
    }
  }
  
  const grandTotalMins = finalRows.filter(r => !r.isTotal).reduce((acc, r) => acc + (r.regStr ? parseFloat(r.regStr)*60 : 0), 0);

  return {
    employeeName: employeeName,
    startDateStr: Utilities.formatDate(startDt, Session.getScriptTimeZone(), "MMM yyyy"),
    rows: finalRows,
    grandTotalStr: (grandTotalMins / 60).toFixed(2) + " Hours"
  };
}

function generateTimesheetHtml(data) {
  let html = `
    <html>
    <head>
      <style>
        body { font-family: Arial, sans-serif; font-size: 10px; color: #333; margin: 20px; }
        .header-container { width: 100%; border-bottom: 2px solid #D01C85; padding-bottom: 10px; margin-bottom: 20px; }
        .title { font-size: 24px; font-weight: bold; color: #D01C85; margin: 0; }
        .subtitle { font-size: 16px; color: #555; margin: 0; }
        .info { margin-top: 10px; font-size: 12px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid #ccc; padding: 6px; text-align: center; }
        th { background-color: #f1f5f9; font-weight: bold; font-size: 9px; text-transform: uppercase; color: #555; }
        .week-total td { background-color: #fdf2f8; font-weight: bold; color: #D01C85; border-top: 2px solid #D01C85; }
        .grand-total td { background-color: #D01C85; color: white; font-weight: bold; font-size: 12px; }
        .footer-table { width: 100%; margin-top: 40px; text-align: left; }
        .footer-table td { border: none; padding: 10px; vertical-align: bottom; }
        .sign-line { border-bottom: 1px solid #000; width: 80%; margin-top: 30px; margin-bottom: 5px; }
        .company-info { margin-top: 30px; font-size: 9px; color: #777; text-align: center; border-top: 1px solid #eee; padding-top: 10px; }
      </style>
    </head>
    <body>
      <div class="header-container">
        <h1 class="title">VUMA Logger</h1>
        <h2 class="subtitle">MONTHLY TIMESHEET</h2>
        <div class="info">
          <strong>Employee:</strong> ${data.employeeName}<br>
          <strong>Start Date of First Week:</strong> ${data.startDateStr}
        </div>
      </div>
      
      <table>
        <thead>
          <tr>
            <th>Date</th><th>Day</th><th>Start Time</th><th>Lunch</th><th>End Time</th>
            <th>Regular Hrs</th><th>Overtime Hrs</th><th>Sick Leave</th><th>Annual Leave</th>
            <th>Family Leave</th><th>Other</th>
          </tr>
        </thead>
        <tbody>
  `;
  
  data.rows.forEach(r => {
    if (r.isTotal) {
      html += `<tr class="week-total"><td colspan="5" style="text-align:left; padding-left:10px;">${r.label}</td><td>${r.val}</td><td colspan="5"></td></tr>`;
    } else {
      html += `<tr>
        <td>${r.date}</td><td>${r.day}</td><td>${r.start}</td><td>${r.lunch}</td><td>${r.end}</td>
        <td>${r.regStr}</td><td>${r.otStr}</td><td>${r.sick}</td><td>${r.annual}</td>
        <td>${r.family}</td><td>${r.other}</td>
      </tr>`;
    }
  });
  
  html += `<tr class="grand-total"><td colspan="5" style="text-align:left; padding-left:10px;">TOTAL HOURS</td><td>${data.grandTotalStr}</td><td colspan="5"></td></tr>`;
  html += `</tbody></table>
  
      <table class="footer-table">
        <tr>
          <td><div class="sign-line"></div>Employee Signature:<br>Date:</td>
          <td><div class="sign-line"></div>Manager Supervisor:<br>Date:</td>
        </tr>
      </table>
      
      <div class="company-info">
        <strong>Powered by</strong> VUMA Logger — Workforce Attendance Portal<br>
        vumatel.co.za
      </div>
    </body>
    </html>
  `;
  return html;
}

function exportTimesheetData(email, startDate, endDate, format) {
  try {
    const data = fetchTimesheetData(email, startDate, endDate);
    
    if (format === 'csv') {
      let csvContent = "DATE,DAY,START TIME,LUNCH,END TIME,REGULAR HRS,OVERTIME HRS,SICK LEAVE,ANNUAL LEAVE,FAMILY LEAVE,OTHER\n";
      data.rows.forEach(r => {
        if (r.isTotal) {
          csvContent += `Total Weekly Hours,,,,,,${r.val},,,,\n`;
        } else {
          csvContent += `"${r.date}","${r.day}","${r.start}","${r.lunch}","${r.end}","${r.regStr}","${r.otStr}","${r.sick}","${r.annual}","${r.family}","${r.other}"\n`;
        }
      });
      csvContent += `TOTAL HOURS,,,,,,${data.grandTotalStr},,,,\n`;
      
      const blob = Utilities.newBlob(csvContent, "text/csv", `${data.employeeName.replace(/\s+/g,'_')}_Timesheet.csv`);
      return { status: "success", base64: Utilities.base64Encode(blob.getBytes()), filename: blob.getName(), mimeType: "text/csv" };
    } 
    
    if (format === 'pdf') {
      const html = generateTimesheetHtml(data);
      const blob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
      blob.setName(`${data.employeeName.replace(/\s+/g,'_')}_Timesheet.pdf`);
      return { status: "success", base64: Utilities.base64Encode(blob.getBytes()), filename: blob.getName(), mimeType: "application/pdf" };
    }
    
    return { status: "error", message: "Invalid format requested." };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

function generateTimesheetsZip(email, startDateStr, endDateStr) {
  try {
    const userSheet = getOrCreateSheet("Users", []);
    const userData = userSheet.getDataRange().getDisplayValues();
    
    let targetEmails = [];
    if (email === "ALL") {
      const seenEmails = new Set();
      for (let i = 1; i < userData.length; i++) {
        const e = String(userData[i][0]).trim().toLowerCase();
        if (e && !seenEmails.has(e)) {
          seenEmails.add(e);
          targetEmails.push(e);
        }
      }
    } else {
      targetEmails.push(email.trim().toLowerCase());
    }

    if (targetEmails.length === 0) {
      return { status: "error", message: "No employees found to generate timesheets." };
    }

    const pdfBlobs = [];
    const processedNames = new Set(); 

    targetEmails.forEach(empEmail => {
      const data = fetchTimesheetData(empEmail, startDateStr, endDateStr);
      const rawName = data.employeeName.trim();
      
      if (!processedNames.has(rawName)) {
        processedNames.add(rawName);
        const html = generateTimesheetHtml(data);
        const blob = HtmlService.createHtmlOutput(html).getAs('application/pdf');
        blob.setName(`${rawName.replace(/\s+/g,'_')}_Timesheet.pdf`);
        pdfBlobs.push(blob);
      }
    });

    if (pdfBlobs.length === 0) {
      return { status: "error", message: "No timesheets could be generated." };
    }

    const dateStrZip = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT", "yyyyMMdd");
    const zipName = `VUMA_Timesheets_${dateStrZip}.zip`;
    
    const zipBlob = Utilities.zip(pdfBlobs, zipName);
    const base64Data = Utilities.base64Encode(zipBlob.getBytes());

    return {
      status: "success",
      filename: zipName,
      mimeType: zipBlob.getContentType(),
      base64: base64Data
    };

  } catch (error) {
    return { status: "error", message: error.toString() };
  }
}

/* ==========================================================================
   EXTRACTION / RE-CALCULATION UTILITY (SMART SYNC)
   ========================================================================== */

function extractExistingLogsToDailySummaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = getOrCreateSheet("Activity_Logs", ["Log ID", "Email", "First Name", "Last Name", "Date", "Time", "Action", "Latitude", "Longitude", "Resolved Location Name", "Edited By"]);
  const summarySheet = getOrCreateSheet("Daily_Summaries", ["First Name", "Last Name", "Summary ID", "Date", "Email", "Clock In Time", "Clock Out Time", "Clock-In Floor Time", "Clock-Out Floor Time", "Total Hours worked", "Status", "Worked Overtime"]);
  
  const logData = logSheet.getDataRange().getDisplayValues();
  if (logData.length <= 1) return "No activity logs found to process.";

  const summaryData = summarySheet.getDataRange().getDisplayValues();
  const existingSummariesMap = {};

  for (let i = 1; i < summaryData.length; i++) {
    const dateStr = String(summaryData[i][3]).trim();
    const email = String(summaryData[i][4]).trim().toLowerCase();
    const status = String(summaryData[i][10]).trim();
    
    if (email && dateStr) {
      const key = `${email}_${dateStr}`;
      existingSummariesMap[key] = {
        rowIndex: i + 1, 
        status: status
      };
    }
  }

  const shiftMap = {};

  for (let i = 1; i < logData.length; i++) {
    const email = String(logData[i][1] || "").trim().toLowerCase();
    const firstName = String(logData[i][2] || "").trim();
    const lastName = String(logData[i][3] || "").trim();
    let dateStr = String(logData[i][4] || "").trim();
    if (dateStr.includes("1899")) dateStr = "";
    const timeStr = String(logData[i][5] || "").trim();
    const action = String(logData[i][6] || "").trim();

    if (!email || !dateStr) continue;
    const key = `${email}_${dateStr}`;

    if (!shiftMap[key]) {
      shiftMap[key] = {
        firstName: firstName,
        lastName: lastName,
        email: email,
        date: dateStr,
        clockIn: "--",
        clockOut: "--"
      };
    }

    if (action === "Clock In") shiftMap[key].clockIn = timeStr;
    else if (action === "Clock Out") shiftMap[key].clockOut = timeStr;
  }

  let newlyAdded = 0;
  let updatedInProgress = 0;
  let skippedCompleted = 0;
  
  const userWeeksToSync = new Set();
  const userMonthsToSync = new Set();

  for (const key in shiftMap) {
    const shift = shiftMap[key];

    if (existingSummariesMap[key] && existingSummariesMap[key].status === "Completed") {
      skippedCompleted++;
      continue; 
    }

    let floorIn = "--"; 
    let floorOut = "--";
    let totalHours = "N/A"; 
    let overtime = "0 mins";
    let status = "In Progress";

    if (shift.clockIn !== "--" && shift.clockOut !== "--") {
      const calculated = calculateHoursAndOvertime(shift.clockIn, shift.clockOut, shift.date, false);
      floorIn = calculated.floorIn;
      floorOut = calculated.floorOut;
      totalHours = calculated.totalHours;
      overtime = calculated.overtime;
      status = "Completed";
    } else if (shift.clockIn !== "--") {
      const startMins = parseTimeToMinutes(shift.clockIn);
      floorIn = (startMins !== null && startMins < 480) ? "08:00:00" : shift.clockIn;
    } else if (shift.clockIn === "--" && shift.clockOut !== "--") {
      totalHours = "No Clock In";
      overtime = "N/A";
      status = "Completed";
    }

    if (existingSummariesMap[key]) {
      const row = existingSummariesMap[key].rowIndex;
      summarySheet.getRange(row, 6).setValue(shift.clockIn);
      summarySheet.getRange(row, 7).setValue(shift.clockOut);
      summarySheet.getRange(row, 8).setValue(floorIn);
      summarySheet.getRange(row, 9).setValue(floorOut);
      summarySheet.getRange(row, 10).setValue(totalHours);
      summarySheet.getRange(row, 11).setValue(status);
      summarySheet.getRange(row, 12).setValue(overtime);
      updatedInProgress++;
    } else {
      const summaryId = "SUM-" + Math.floor(100000 + Math.random() * 900000);
      summarySheet.appendRow([shift.firstName, shift.lastName, summaryId, shift.date, shift.email, shift.clockIn, shift.clockOut, floorIn, floorOut, totalHours, status, overtime]);
      newlyAdded++;
    }
    
    if (shift.date) {
      const weekKey = getISOWeekKey(new Date(shift.date));
      userWeeksToSync.add(`${shift.email}|${weekKey}`);
      userMonthsToSync.add(`${shift.email}|${shift.date.substring(0, 7)}`);
    }
  }

  userWeeksToSync.forEach(item => {
    const parts = item.split("|");
    syncWeeklySummariesForUser(parts[0], parts[1]);
  });

  userMonthsToSync.forEach(item => {
    const parts = item.split("|");
    syncMonthlySummariesForUser(parts[0], parts[1]);
  });

  return `Smart Sync Complete: Added ${newlyAdded} new shifts, Updated ${updatedInProgress} pending shifts, and Safely Skipped ${skippedCompleted} completed/edited shifts.`;
}

/* ==========================================================================
   ADMIN DASHBOARD DATA ENGINE
   ========================================================================== */

function getAdminDashboardData() {
  try {
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    const todayMs = new Date(todayStr).getTime();
    
    const logSheet = getOrCreateSheet("Activity_Logs", ["Log ID", "Email", "First Name", "Last Name", "Date", "Time", "Action", "Latitude", "Longitude", "Resolved Location Name", "Edited By"]);
    const logDisplayData = logSheet.getDataRange().getDisplayValues();
    const logs = [];
    let totalLoginsToday = 0; let totalLogoutsToday = 0;

    for (let i = 1; i < logDisplayData.length; i++) {
      let rawDateStr = String(logDisplayData[i][4] || "").trim();
      if (rawDateStr.includes("1899")) rawDateStr = "";
      const actionStr = String(logDisplayData[i][6] || "").trim();
      
      if (rawDateStr === todayStr) {
        if (actionStr === "Clock In") totalLoginsToday++;
        if (actionStr === "Clock Out") totalLogoutsToday++;
      }

      logs.push({
        id: String(logDisplayData[i][0] || ""),
        email: String(logDisplayData[i][1] || ""),
        firstName: String(logDisplayData[i][2] || ""),
        lastName: String(logDisplayData[i][3] || ""),
        date: rawDateStr, 
        time: String(logDisplayData[i][5] || ""),
        action: actionStr,
        lat: String(logDisplayData[i][7] || ""),
        lon: String(logDisplayData[i][8] || ""),
        location: String(logDisplayData[i][9] || "Unknown Location"),
        editedBy: String(logDisplayData[i][10] || "")
      });
    }
    
    const summarySheet = getOrCreateSheet("Daily_Summaries", ["First Name", "Last Name", "Summary ID", "Date", "Email", "Clock In Time", "Clock Out Time", "Clock-In Floor Time", "Clock-Out Floor Time", "Total Hours worked", "Status", "Worked Overtime"]);
    const summaryData = summarySheet.getDataRange().getDisplayValues();
    const summaries = [];
    for (let i = 1; i < summaryData.length; i++) {
      summaries.push({
        firstName: String(summaryData[i][0] || ""),
        lastName: String(summaryData[i][1] || ""),
        summaryId: String(summaryData[i][2] || ""),
        date: String(summaryData[i][3] || ""),
        email: String(summaryData[i][4] || ""),
        clockIn: String(summaryData[i][5] || "--"),
        clockOut: String(summaryData[i][6] || "--"),
        floorIn: String(summaryData[i][7] || "--"),
        floorOut: String(summaryData[i][8] || "--"),
        totalHours: String(summaryData[i][9] || "N/A"),
        status: String(summaryData[i][10] || "In Progress"),
        overtime: String(summaryData[i][11] || "0 mins")
      });
    }
    
    const weeklySheet = getOrCreateSheet("Weekly_Summaries", ["First Name", "Last Name", "Email", "Week Identifier", "Total Hours Worked", "Total Overtime", "Days Worked", "Compliance Status", "Last Updated"]);
    const weeklyData = weeklySheet.getDataRange().getDisplayValues();
    const weeklySummaries = [];
    for (let i = 1; i < weeklyData.length; i++) {
      weeklySummaries.push({
        firstName: String(weeklyData[i][0] || ""),
        lastName: String(weeklyData[i][1] || ""),
        email: String(weeklyData[i][2] || ""),
        weekKey: String(weeklyData[i][3] || ""),
        totalHours: String(weeklyData[i][4] || "0 mins"),
        totalOvertime: String(weeklyData[i][5] || "0 mins"),
        daysWorked: String(weeklyData[i][6] || "0"),
        complianceStatus: String(weeklyData[i][7] || "Under Target (<45 hrs)"),
        lastUpdated: String(weeklyData[i][8] || "")
      });
    }
    
    const monthlySheet = getOrCreateSheet("Monthly_Summaries", ["First Name", "Last Name", "Email", "Month", "Total Hours Worked", "Total Overtime", "Days Worked", "Last Updated"]);
    const monthlyData = monthlySheet.getDataRange().getDisplayValues();
    const monthlySummaries = [];
    for (let i = 1; i < monthlyData.length; i++) {
      monthlySummaries.push({
        firstName: String(monthlyData[i][0] || ""),
        lastName: String(monthlyData[i][1] || ""),
        email: String(monthlyData[i][2] || ""),
        month: String(monthlyData[i][3] || ""),
        totalHours: String(monthlyData[i][4] || "0 mins"),
        totalOvertime: String(monthlyData[i][5] || "0 mins"),
        daysWorked: String(monthlyData[i][6] || "0"),
        lastUpdated: String(monthlyData[i][7] || "")
      });
    }
    
    const leaveRequests = [];
    const onLeaveEmails = new Set();
    const leaveSheet = getOrCreateSheet("Leave_Requests", ["Request ID", "Email", "First Name", "Last Name", "Leave Type", "Start Date", "End Date", "Status", "Submitted Date"]);
    const leaveData = leaveSheet.getDataRange().getDisplayValues();
    
    for (let i = 1; i < leaveData.length; i++) {
      const reqId = String(leaveData[i][0] || "").trim();
      const empEmail = String(leaveData[i][1] || "").trim().toLowerCase();
      const fName = String(leaveData[i][2] || "").trim();
      const lName = String(leaveData[i][3] || "").trim();
      const lType = String(leaveData[i][4] || "").trim();
      const startDateStr = String(leaveData[i][5] || "").trim();
      const endDateStr = String(leaveData[i][6] || "").trim();
      const status = String(leaveData[i][7] || "").trim();
      const submittedDate = String(leaveData[i][8] || "").trim();

      leaveRequests.push({
        reqId: reqId, email: empEmail, firstName: fName, lastName: lName, type: lType, startDate: startDateStr, endDate: endDateStr, status: status, submittedDate: submittedDate
      });

      if (status === "Approved") {
        const startMs = new Date(startDateStr).getTime();
        const endMs = new Date(endDateStr).getTime();
        if (todayMs >= startMs && todayMs <= endMs) {
          onLeaveEmails.add(empEmail);
        }
      }
    }
    
    const userSheet = getOrCreateSheet("Users", ["Email", "First Name", "Last Name", "Password Hash", "Salt", "Account Type", "Status", "Verification Token"]);
    const userData = userSheet.getDataRange().getDisplayValues();
    const employees = [];
    for (let i = 1; i < userData.length; i++) {
      const empEmail = String(userData[i][0] || "").trim().toLowerCase();
      const standing = String(userData[i][6] || "Verified").trim();
      
      employees.push({
        email: empEmail,
        firstName: String(userData[i][1] || ""),
        lastName: String(userData[i][2] || ""),
        accountType: String(userData[i][5] || "Employee"),
        accountStanding: standing,
        workStatus: onLeaveEmails.has(empEmail) ? "On Leave" : "Active"
      });
    }
    
    return { 
      status: "success", 
      logs: logs.reverse(), 
      summaries: summaries.reverse(), 
      weeklySummaries: weeklySummaries.reverse(),
      monthlySummaries: monthlySummaries.reverse(),
      employees: employees,
      leaveRequests: leaveRequests.reverse(),
      dailyStats: { loginsToday: totalLoginsToday, logoutsToday: totalLogoutsToday }
    };
  } catch(e) {
    return { status: "error", message: "Database Error: " + e.toString() };
  }
}

function adminUpdateUser(targetEmail, updatedFirst, updatedLast, updatedStatus, newPassword) {
  try {
    const sheet = getOrCreateSheet("Users", ["Email", "First Name", "Last Name", "Password Hash", "Salt", "Account Type", "Status", "Verification Token"]);
    const data = sheet.getDataRange().getValues();
    const targetClean = targetEmail.trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === targetClean) {
        const row = i + 1;
        sheet.getRange(row, 2).setValue(updatedFirst.trim());
        sheet.getRange(row, 3).setValue(updatedLast.trim());
        sheet.getRange(row, 7).setValue(updatedStatus);
        
        if (newPassword && newPassword.trim() !== "") {
          const newSalt = generateSalt();
          const newHash = computeSHA256(newPassword.trim(), newSalt);
          sheet.getRange(row, 4).setValue(newHash);
          sheet.getRange(row, 5).setValue(newSalt);
        }
        return { status: "success", message: "Employee profile updated successfully." };
      }
    }
    return { status: "error", message: "User targeted not found." };
  } catch(e) {
    return { status: "error", message: e.toString() };
  }
}

/* ==========================================================================
   HTTP API ENDPOINT (For GitHub Pages integration)
   ========================================================================== */
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const method = payload.method;
    const args = payload.args || [];
    let result = {};

    switch (method) {
      case "registerUser": result = registerUser(...args); break;
      case "loginUser": result = loginUser(...args); break;
      case "processActivityLog": result = processActivityLog(...args); break;
      case "getAdminDashboardData": result = getAdminDashboardData(); break;
      case "submitLeaveRequest": result = submitLeaveRequest(...args); break;
      case "adminUpdateLeaveStatus": result = adminUpdateLeaveStatus(...args); break;
      case "adminUpdateUser": result = adminUpdateUser(...args); break;
      case "adminDeleteDailySummary": result = adminDeleteDailySummary(...args); break;
      case "adminUpdateDailySummary": result = adminUpdateDailySummary(...args); break;
      case "generateTimesheetsZip": result = generateTimesheetsZip(...args); break;
      case "exportTimesheetData": result = exportTimesheetData(...args); break;
      default: result = { status: "error", message: "Unknown method: " + method };
    }

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}
