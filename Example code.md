FNB code
**Code.gs**
const LOCATIONIQ_API_KEY = "pk.1be07ee2080691339d8fc4f1712dbc95";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('FNB Logsheet System - Corporate Access Portal')
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
      // If admin explicitly stretches shift past 17:00 -> separate OT from Regular standard bounds
      if (endMinutes > FLOOR_END_MINS) {
         otMinutes = endMinutes - FLOOR_END_MINS;
         standardWorkMinutes = FLOOR_END_MINS - effectiveStartMinutes;
      } else {
         // If admin enters an early leave (e.g. 15:00) -> 0 OT, actual worked hours
         otMinutes = 0;
         standardWorkMinutes = endMinutes - effectiveStartMinutes;
      }
    } else {
      // Non-Admin: Strict 08:00 to 17:00 calculations regardless of true telemetry
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
       floorIn = finalIn; // Respect admin edit directly even if incomplete
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

/**
 * Permanently deletes a specific daily shift and its underlying raw logs to prevent ghost syncing.
 */
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
        // If admin edited, accurately show exactly what the admin set the floors to.
        dateMap[sDate].start = formatTimeForPDF(sFloorIn);
        dateMap[sDate].end = formatTimeForPDF(sFloorOut);
        dateMap[sDate].otMins = parseDurationStringToMinutes(sOt);
      } else {
        // Non-admin shift strictly defaults to 08:00 AM / 17:00 PM for PDF view and 0 Overtime
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
    
    // Feature: Mathematically round overtime to nearest integer hour
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
        .header-container { width: 100%; border-bottom: 2px solid #00a896; padding-bottom: 10px; margin-bottom: 20px; }
        .title { font-size: 24px; font-weight: bold; color: #00a896; margin: 0; }
        .subtitle { font-size: 16px; color: #555; margin: 0; }
        .info { margin-top: 10px; font-size: 12px; }
        table { width: 100%; border-collapse: collapse; margin-bottom: 20px; }
        th, td { border: 1px solid #ccc; padding: 6px; text-align: center; }
        th { background-color: #f1f5f9; font-weight: bold; font-size: 9px; text-transform: uppercase; color: #555; }
        .week-total td { background-color: #e6fffa; font-weight: bold; color: #00a896; border-top: 2px solid #00a896; }
        .grand-total td { background-color: #00a896; color: white; font-weight: bold; font-size: 12px; }
        .footer-table { width: 100%; margin-top: 40px; text-align: left; }
        .footer-table td { border: none; padding: 10px; vertical-align: bottom; }
        .sign-line { border-bottom: 1px solid #000; width: 80%; margin-top: 30px; margin-bottom: 5px; }
        .company-info { margin-top: 30px; font-size: 9px; color: #777; text-align: center; border-top: 1px solid #eee; padding-top: 10px; }
      </style>
    </head>
    <body>
      <div class="header-container">
        <h1 class="title">MolāDira SKILLS</h1>
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
        <strong>Address:</strong> White Thorn Office Park, 606 Kudu Street Allen's Nek, Roodepoort<br>
        <strong>BEE Status:</strong> BBBEE Level 1<br>
        www.moladiraskills.co.za
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
    const zipName = `Moladira_Timesheets_${dateStrZip}.zip`;
    
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

function forceUpdateOldLogsToShowcaseFeatures() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Daily_Summaries");
  const logSheet = ss.getSheetByName("Activity_Logs");
  
  if (!summarySheet || !logSheet) return "Error: Required sheets not found.";

  const logData = logSheet.getDataRange().getDisplayValues();
  const adminEditedDates = new Set();
  
  for (let j = 1; j < logData.length; j++) {
    const rowEmail = String(logData[j][1]).trim().toLowerCase();
    const rowDate = String(logData[j][4]).trim();
    const editedBy = String(logData[j][10]).trim();
    if (editedBy !== "" && editedBy !== "System Automation") {
      adminEditedDates.add(`${rowEmail}_${rowDate}`);
    }
  }

  const data = summarySheet.getDataRange().getDisplayValues();
  let updatedCount = 0;
  
  const userWeeksToSync = new Set();
  const userMonthsToSync = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const dateStr = String(data[i][3]).trim();
    const email = String(data[i][4]).trim().toLowerCase();
    const clockIn = String(data[i][5]).trim();
    const clockOut = String(data[i][6]).trim();
    const status = String(data[i][10]).trim(); 
    
    if (clockIn && clockIn !== "--" && dateStr) {
      
      let floorIn = "--"; 
      let floorOut = "--";
      let totalHours = "N/A"; 
      let overtime = "0 mins";
      let newStatus = status;

      if (clockIn !== "--" && clockOut !== "--") {
        const isAdminEdit = adminEditedDates.has(`${email}_${dateStr}`);
        const calc = calculateHoursAndOvertime(clockIn, clockOut, dateStr, isAdminEdit);
        floorIn = calc.floorIn;
        floorOut = calc.floorOut;
        totalHours = calc.totalHours;
        overtime = calc.overtime;
        newStatus = "Completed";
      } else if (clockIn !== "--") {
        const startMins = parseTimeToMinutes(clockIn);
        floorIn = (startMins !== null && startMins < 480) ? "08:00:00" : clockIn;
      }

      summarySheet.getRange(row, 8).setValue(floorIn);       
      summarySheet.getRange(row, 9).setValue(floorOut);      
      summarySheet.getRange(row, 10).setValue(totalHours);   
      summarySheet.getRange(row, 11).setValue(newStatus);    
      summarySheet.getRange(row, 12).setValue(overtime);     
      
      updatedCount++;

      const weekKey = getISOWeekKey(new Date(dateStr));
      userWeeksToSync.add(`${email}|${weekKey}`);
      userMonthsToSync.add(`${email}|${dateStr.substring(0, 7)}`);
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

  return `Success! Upgraded ${updatedCount} historical shifts. Weekly & Monthly totals have been resynced.`;
}

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

function resetClockOutFloorTimes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summarySheet = ss.getSheetByName("Daily_Summaries");
  const logSheet = ss.getSheetByName("Activity_Logs");
  
  if (!summarySheet || !logSheet) return "Error: Required sheets not found.";

  const logData = logSheet.getDataRange().getDisplayValues();
  const adminEditedDates = new Set();
  
  for (let j = 1; j < logData.length; j++) {
    const rowEmail = String(logData[j][1]).trim().toLowerCase();
    const rowDate = String(logData[j][4]).trim();
    const editedBy = String(logData[j][10]).trim();
    if (editedBy !== "" && editedBy !== "System Automation") {
      adminEditedDates.add(`${rowEmail}_${rowDate}`);
    }
  }

  const data = summarySheet.getDataRange().getDisplayValues();
  let updatedCount = 0;
  
  const userWeeksToSync = new Set();
  const userMonthsToSync = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = i + 1;
    const dateStr = String(data[i][3]).trim();
    const email = String(data[i][4]).trim().toLowerCase();
    const clockIn = String(data[i][5]).trim();
    const clockOut = String(data[i][6]).trim();
    const status = String(data[i][10]).trim(); 
    
    if (dateStr && clockIn !== "--" && email) {
      
      let newFloorIn = "08:00:00"; 
      let newFloorOut = "--";
      let totalHours = "N/A"; 
      let overtime = "0 mins";
      let newStatus = status;

      const dateObj = new Date(dateStr);
      const isSaturday = dateObj.getDay() === 6;
      const baselineFloorOut = isSaturday ? "13:00:00" : "17:00:00";

      if (clockIn !== "--" && clockOut !== "--") {
        const isAdminEdit = adminEditedDates.has(`${email}_${dateStr}`);
        const calc = calculateHoursAndOvertime(clockIn, clockOut, dateStr, isAdminEdit);
        
        newFloorIn = calc.floorIn;
        newFloorOut = isAdminEdit ? calc.floorOut : baselineFloorOut; 
        
        totalHours = calc.totalHours;
        overtime = calc.overtime;
        newStatus = "Completed";
      } else {
        newFloorOut = baselineFloorOut;
      }

      summarySheet.getRange(row, 8).setValue(newFloorIn);        
      summarySheet.getRange(row, 9).setValue(newFloorOut);       
      summarySheet.getRange(row, 10).setValue(totalHours);       
      summarySheet.getRange(row, 11).setValue(newStatus);        
      summarySheet.getRange(row, 12).setValue(overtime);         
      
      updatedCount++;

      const weekKey = getISOWeekKey(new Date(dateStr));
      userWeeksToSync.add(`${email}|${weekKey}`);
      userMonthsToSync.add(`${email}|${dateStr.substring(0, 7)}`);
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

  return `Success! Reset Clock-Out Floor Times for ${updatedCount} historical shifts. Weekly & Monthly totals have been resynced.`;
}


**HTML**
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>FNB Logsheet System - Corporate Access Portal</title>
  <script src="https://cdn.jsdelivr.net/npm/@tailwindcss/browser@4"></script>
  <style>
    .loader { border-top-color: #0d9488; animation: spinner 1.5s linear infinite; }
    @keyframes spinner { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
    .smooth-scroll { -webkit-overflow-scrolling: touch; }
    
    .bg-fnb-teal { background-color: #00a896; }
    .bg-fnb-teal-gradient { background: linear-gradient(135deg, #00a896 0%, #0d9488 100%); }
    .text-fnb-teal { color: #00a896; }
  </style>
</head>
<body class="bg-slate-100 min-h-screen text-slate-800 font-sans antialiased smooth-scroll pb-12">

  <div id="globalSpinner" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex flex-col justify-center items-center z-50">
    <div class="loader ease-linear rounded-full border-4 border-t-4 border-slate-200 h-14 w-14 mb-4"></div>
    <p class="text-slate-800 font-bold bg-white px-6 py-3 rounded-xl shadow-xl text-sm tracking-wide border border-slate-200" id="spinnerMessage">Processing request...</p>
  </div>

  <main id="mainContainer" class="max-w-md mx-auto p-4 md:p-8 transition-all duration-300">
    
    <!-- AUTHENTICATION VIEW -->
    <section id="authView" class="block bg-white p-6 md:p-8 rounded-3xl shadow-xl border border-slate-200 mt-6 md:mt-12 max-w-md mx-auto">
      <div class="text-center mb-8">
        <div class="w-16 h-16 bg-teal-50 rounded-2xl flex items-center justify-center mx-auto mb-4 border border-teal-100 shadow-sm">
          <span class="text-2xl font-black text-teal-600 tracking-tighter">FNB</span>
        </div>
        <h1 class="text-2xl font-black tracking-tight text-slate-900" id="authTitle">FNB <span class="font-light text-teal-600">Logsheet System</span></h1>
        <p class="text-xs text-slate-500 mt-1.5" id="authSubtitle">Enter credentials to authenticate workspace actions.</p>
      </div>

      <div id="authAlert" class="hidden mb-5 p-4 rounded-2xl text-xs md:text-sm font-medium"></div>

      <form id="loginForm" class="space-y-4" onsubmit="handleLoginSubmit(event)">
        <div>
          <label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Work Email</label>
          <input type="email" id="loginEmail" required class="w-full px-4 py-3 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-500 bg-slate-50 text-slate-900 placeholder-slate-400 text-base">
        </div>
        <div>
          <label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Password</label>
          <input type="password" id="loginPassword" required class="w-full px-4 py-3 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-500 bg-slate-50 text-slate-900 placeholder-slate-400 text-base">
        </div>
        <button type="submit" class="w-full bg-fnb-teal-gradient hover:opacity-95 text-white font-bold py-3.5 rounded-xl transition shadow-md text-base active:scale-95 duration-150">Authenticate Login</button>
        <p class="text-xs text-center text-slate-500 pt-4">New employee? <a href="#" onclick="toggleAuthMode(false)" class="text-teal-600 hover:underline font-bold">Create access account</a></p>
      </form>

      <form id="registerForm" class="space-y-4 hidden" onsubmit="handleRegisterSubmit(event)">
        <div class="grid grid-cols-2 gap-3">
          <div><label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">First Name</label><input type="text" id="regFirst" required class="w-full px-4 py-3 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-500 bg-slate-50 text-slate-900 text-base"></div>
          <div><label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Last Name</label><input type="text" id="regLast" required class="w-full px-4 py-3 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-500 bg-slate-50 text-slate-900 text-base"></div>
        </div>
        <div><label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Email Address</label><input type="email" id="regEmail" required class="w-full px-4 py-3 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-500 bg-slate-50 text-slate-900 text-base"></div>
        <div><label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Create Password</label><input type="password" id="regPassword" required class="w-full px-4 py-3 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-500 bg-slate-50 text-slate-900 text-base"></div>
        <button type="submit" class="w-full bg-fnb-teal-gradient hover:opacity-95 text-white font-bold py-3.5 rounded-xl transition shadow-md text-base active:scale-95 duration-150">Submit Registration</button>
        <p class="text-xs text-center text-slate-500 pt-4">Already have access profile? <a href="#" onclick="toggleAuthMode(true)" class="text-teal-600 hover:underline font-bold">Log in instead</a></p>
      </form>
    </section>

    <!-- EMPLOYEE PORTAL VIEW -->
    <section id="employeeView" class="hidden bg-white rounded-3xl shadow-xl p-6 border border-slate-200 mt-6 max-w-md mx-auto">
      <div class="flex flex-col items-center border-b border-slate-100 pb-5 mb-6 text-center">
        <span class="text-[10px] font-extrabold text-teal-700 tracking-wider uppercase bg-teal-50 border border-teal-200 px-3 py-1 rounded-full mb-3">Session Active</span>
        <h2 class="text-xl font-bold text-slate-900 mb-1" id="employeeGreeting">Hello World and Welcome!</h2>
        <button onclick="logout()" class="text-xs text-slate-500 hover:text-slate-900 bg-slate-100 hover:bg-slate-200 border border-slate-200 px-4 py-2 rounded-xl font-semibold transition mt-3">Logout Session</button>
      </div>

      <div id="logAlert" class="hidden mb-6 p-4 rounded-2xl text-xs border"></div>

      <p class="text-xs text-slate-500 mb-6 text-center leading-relaxed">Your location settings are critical for login logging. Ensure browser location permission is authorized.</p>
      
      <div class="grid grid-cols-1 gap-4 mb-4">
        <button onclick="triggerClockAction('Clock In')" class="flex items-center justify-between p-5 bg-fnb-teal-gradient hover:opacity-95 active:scale-95 text-white rounded-2xl transition-all shadow-md duration-150">
          <div class="flex items-center gap-4"><span class="text-3xl">⏰</span><span class="font-bold text-lg">Clock In</span></div><span class="text-xl">➔</span>
        </button>
        <button onclick="triggerClockAction('Clock Out')" class="flex items-center justify-between p-5 bg-orange-500 hover:bg-orange-600 active:scale-95 text-white rounded-2xl transition-all shadow-md duration-150">
          <div class="flex items-center gap-4"><span class="text-3xl">🚪</span><span class="font-bold text-lg">Clock Out</span></div><span class="text-xl">➔</span>
        </button>
      </div>

      <button onclick="openLeaveModal()" class="w-full flex items-center justify-center gap-2 p-4 bg-slate-50 hover:bg-slate-100 border border-slate-200 active:scale-95 text-slate-700 font-bold rounded-2xl transition shadow-sm text-sm">
        <span>📅</span> Request Leave
      </button>
    </section>

    <!-- ADMIN DASHBOARD VIEW -->
    <section id="adminView" class="hidden space-y-6 mt-4 w-full">
      <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center bg-white p-5 rounded-2xl border border-slate-200 shadow-sm gap-4">
        <div>
          <span class="text-[9px] font-bold text-teal-700 tracking-wider uppercase bg-teal-50 px-2.5 py-0.5 rounded-full border border-teal-200">FNB Admin Space</span>
          <h2 class="text-lg md:text-xl font-extrabold text-slate-900 mt-1" id="adminGreeting">Manager Dashboard</h2>
          <button onclick="toggleAdminClockPanel()" class="mt-2.5 inline-flex items-center gap-1.5 text-xs font-semibold bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 px-3 py-1.5 rounded-xl transition active:scale-95">
            <span>⏰</span><span id="adminClockToggleText">Go to My Clock In / Out</span>
          </button>
        </div>
        <button onclick="logout()" class="bg-rose-50 hover:bg-rose-100 border border-rose-200 text-rose-600 text-xs font-bold px-4 py-2 rounded-xl transition shadow-sm active:scale-95">Exit</button>
      </div>

      <div id="adminSelfClockPanel" class="hidden bg-white rounded-3xl shadow-lg p-6 border border-slate-200 max-w-lg mx-auto">
        <div class="text-center mb-4">
          <h3 class="text-lg font-bold text-slate-900">Admin Work Attendance Logging</h3>
          <p class="text-xs text-slate-500 mt-1">Record your daily attendance directly as an administrator.</p>
        </div>
        <div id="adminLogAlert" class="hidden mb-4 p-4 rounded-2xl text-xs border"></div>
        <div class="grid grid-cols-2 gap-3 mb-4">
          <button onclick="triggerAdminClockAction('Clock In')" class="flex items-center justify-center gap-2 p-4 bg-fnb-teal-gradient hover:opacity-95 active:scale-95 text-white font-bold rounded-2xl transition shadow-md text-sm"><span>⏰</span> Clock In</button>
          <button onclick="triggerAdminClockAction('Clock Out')" class="flex items-center justify-center gap-2 p-4 bg-orange-500 hover:bg-orange-600 active:scale-95 text-white font-bold rounded-2xl transition shadow-md text-sm"><span>🚪</span> Clock Out</button>
        </div>
        <button onclick="openLeaveModal()" class="w-full flex items-center justify-center gap-2 p-3 bg-slate-50 border border-slate-200 hover:bg-slate-100 active:scale-95 text-slate-700 font-bold rounded-xl transition shadow-sm text-xs"><span>📅</span> Register Leave</button>
      </div>

      <!-- NAVIGATION TABS -->
      <div class="flex w-full bg-white p-1.5 rounded-xl border border-slate-200 shadow-sm gap-1 overflow-x-auto smooth-scroll">
        <button onclick="switchAdminTab('logsTab')" id="btn-logsTab" class="flex-1 py-2.5 px-3 rounded-xl font-bold text-xs whitespace-nowrap transition-all bg-fnb-teal text-white shadow-sm">Activity Logs</button>
        <button onclick="switchAdminTab('summariesTab')" id="btn-summariesTab" class="flex-1 py-2.5 px-3 rounded-xl font-bold text-xs whitespace-nowrap transition-all text-slate-500 hover:text-slate-900 bg-transparent">Daily Summaries</button>
        <button onclick="switchAdminTab('weeklyTab')" id="btn-weeklyTab" class="flex-1 py-2.5 px-3 rounded-xl font-bold text-xs whitespace-nowrap transition-all text-slate-500 hover:text-slate-900 bg-transparent">Weekly Summaries</button>
        <button onclick="switchAdminTab('monthlyTab')" id="btn-monthlyTab" class="flex-1 py-2.5 px-3 rounded-xl font-bold text-xs whitespace-nowrap transition-all text-slate-500 hover:text-slate-900 bg-transparent">Monthly Summaries</button>
        <button onclick="switchAdminTab('leaveTab')" id="btn-leaveTab" class="flex-1 py-2.5 px-3 rounded-xl font-bold text-xs whitespace-nowrap transition-all text-slate-500 hover:text-slate-900 bg-transparent">Leave Requests</button>
        <button onclick="switchAdminTab('usersTab')" id="btn-usersTab" class="flex-1 py-2.5 px-3 rounded-xl font-bold text-xs whitespace-nowrap transition-all text-slate-500 hover:text-slate-900 bg-transparent">Staff Profiles</button>
        <button onclick="switchAdminTab('exportTab')" id="btn-exportTab" class="flex-1 py-2.5 px-3 rounded-xl font-bold text-xs whitespace-nowrap transition-all text-slate-500 hover:text-slate-900 bg-transparent">Export Timesheet</button>
      </div>

      <!-- TAB 1: ACTIVITY LOGS & INLINE DAILY STATS WIDGET -->
      <div id="content-logsTab" class="block bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden w-full">
        <div class="p-4 md:p-5 border-b border-slate-200 bg-slate-50 flex flex-col xl:flex-row justify-between items-start xl:items-center gap-4">
          <div class="flex flex-wrap items-center gap-3 md:gap-5 w-full xl:w-auto">
            <h3 class="font-bold text-slate-900 text-sm md:text-base">Telemetry Activity Logs</h3>
            <div class="flex items-center gap-3 bg-white border border-slate-200 px-3 py-1.5 rounded-lg shadow-sm">
              <div class="flex items-center gap-1.5"><span class="flex h-2 w-2 relative"><span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-teal-400 opacity-75"></span><span class="relative inline-flex rounded-full h-2 w-2 bg-teal-500"></span></span><span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Logins:</span><span class="text-sm font-black text-teal-700" id="statLogins">0</span></div>
              <div class="w-px h-4 bg-slate-200"></div>
              <div class="flex items-center gap-1.5"><span class="flex h-2 w-2 relative"><span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-orange-400 opacity-75"></span><span class="relative inline-flex rounded-full h-2 w-2 bg-orange-500"></span></span><span class="text-[10px] font-bold text-slate-500 uppercase tracking-wider">Logouts:</span><span class="text-sm font-black text-orange-600" id="statLogouts">0</span></div>
              <span class="text-[9px] text-slate-400 font-mono ml-2 hidden sm:inline-block" id="statDateDisplay">Today</span>
            </div>
          </div>
          <div class="flex flex-col sm:flex-row gap-2 w-full xl:w-auto">
            <input type="date" id="filterDate" onchange="applyAdminLogFilters()" class="px-3 py-2 border border-slate-200 rounded-xl bg-white text-slate-800 text-xs focus:outline-none focus:ring-2 focus:ring-teal-500">
            <input type="text" id="filterSearch" oninput="applyAdminLogFilters()" placeholder="Search staff name/email..." class="px-3 py-2 border border-slate-200 rounded-xl bg-white text-slate-800 text-xs placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-teal-500 min-w-[200px]">
            <button onclick="resetAdminLogFilters()" class="text-xs text-slate-600 hover:text-slate-900 bg-slate-100 px-3 py-2 rounded-xl transition border border-slate-300 font-bold">Reset</button>
          </div>
        </div>
        <div class="overflow-x-auto smooth-scroll w-full">
          <table class="w-full text-left border-collapse text-xs md:text-sm">
            <thead>
              <tr class="bg-slate-100 text-slate-500 uppercase text-[10px] md:text-xs font-bold tracking-wider border-b border-slate-200">
                <th class="p-3.5">Staff</th><th class="p-3.5">Time</th><th class="p-3.5">Action</th><th class="p-3.5 min-w-[250px]">Resolved Location & Coordinates</th><th class="p-3.5">Edited By</th>
              </tr>
            </thead>
            <tbody id="logsTableBody" class="divide-y divide-slate-200 text-slate-700"></tbody>
          </table>
        </div>
      </div>

      <!-- TAB 2: DAILY SUMMARIES -->
      <div id="content-summariesTab" class="hidden bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden w-full">
        <div class="p-4 md:p-5 border-b border-slate-200 bg-slate-50 flex flex-col md:flex-row justify-between items-start md:items-center gap-3">
          <h3 class="font-bold text-slate-900 text-sm md:text-base">Daily Shift Summaries</h3>
          <div class="flex flex-col sm:flex-row gap-2 w-full md:w-auto">
            <input type="date" id="filterSummaryDate" onchange="applyAdminSummaryFilters()" class="px-3 py-2 border border-slate-200 rounded-xl bg-white text-slate-800 text-xs focus:outline-none focus:ring-2 focus:ring-teal-500">
            <input type="text" id="filterSummarySearch" oninput="applyAdminSummaryFilters()" placeholder="Search staff name/email..." class="px-3 py-2 border border-slate-200 rounded-xl bg-white text-slate-800 text-xs placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-teal-500 min-w-[200px]">
            <button onclick="resetAdminSummaryFilters()" class="text-xs text-slate-600 hover:text-slate-900 bg-slate-100 px-3 py-2 rounded-xl transition border border-slate-300 font-bold">Reset</button>
          </div>
        </div>
        <div class="overflow-x-auto smooth-scroll w-full">
          <table class="w-full text-left border-collapse text-xs md:text-sm">
            <thead>
              <tr class="bg-slate-100 text-slate-500 uppercase text-[10px] md:text-xs font-bold tracking-wider border-b border-slate-200">
                <th class="p-3.5">Employee</th><th class="p-3.5">Date</th><th class="p-3.5">Actual Times (In / Out)</th><th class="p-3.5">Floor Times (In / Out)</th><th class="p-3.5">Net Hours Worked (-1hr Lunch)</th><th class="p-3.5">Overtime (> 17:00 / Sat > 13:00)</th><th class="p-3.5">Status</th><th class="p-3.5 text-right">Action</th>
              </tr>
            </thead>
            <tbody id="summariesTableBody" class="divide-y divide-slate-200 text-slate-700"></tbody>
          </table>
        </div>
      </div>

      <!-- TAB 3: WEEKLY SUMMARIES -->
      <div id="content-weeklyTab" class="hidden bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden w-full">
        <div class="p-4 md:p-5 border-b border-slate-200 bg-slate-50 flex flex-col md:flex-row justify-between items-start md:items-center gap-3">
          <div>
            <h3 class="font-bold text-slate-900 text-sm md:text-base">Weekly Work Summaries & Legal Hour Limits</h3>
            <span class="text-[10px] font-bold text-slate-600 bg-slate-200 border border-slate-300 px-2.5 py-0.5 rounded-full inline-block mt-1">Min Target: 45h | Legal Max: 55h</span>
          </div>
          <div class="flex flex-col sm:flex-row gap-2 w-full md:w-auto">
            <input type="text" id="filterWeeklyKey" oninput="applyAdminWeeklyFilters()" placeholder="Week e.g. 2026-W30" class="px-3 py-2 border border-slate-200 rounded-xl bg-white text-slate-800 text-xs placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-teal-500">
            <input type="text" id="filterWeeklySearch" oninput="applyAdminWeeklyFilters()" placeholder="Search staff name/email..." class="px-3 py-2 border border-slate-200 rounded-xl bg-white text-slate-800 text-xs placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-teal-500 min-w-[200px]">
            <button onclick="resetAdminWeeklyFilters()" class="text-xs text-slate-600 hover:text-slate-900 bg-slate-100 px-3 py-2 rounded-xl transition border border-slate-300 font-bold">Reset</button>
          </div>
        </div>
        <div class="overflow-x-auto smooth-scroll w-full">
          <table class="w-full text-left border-collapse text-xs md:text-sm">
            <thead>
              <tr class="bg-slate-100 text-slate-500 uppercase text-[10px] md:text-xs font-bold tracking-wider border-b border-slate-200">
                <th class="p-3.5">Employee</th><th class="p-3.5">Week</th><th class="p-3.5">Total Hours Worked</th><th class="p-3.5">Total Overtime</th><th class="p-3.5">Days Worked</th><th class="p-3.5">Legal Compliance Status</th>
              </tr>
            </thead>
            <tbody id="weeklyTableBody" class="divide-y divide-slate-200 text-slate-700"></tbody>
          </table>
        </div>
      </div>

      <!-- TAB 4: MONTHLY SUMMARIES -->
      <div id="content-monthlyTab" class="hidden bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden w-full">
        <div class="p-4 md:p-5 border-b border-slate-200 bg-slate-50 flex flex-col md:flex-row justify-between items-start md:items-center gap-3">
          <h3 class="font-bold text-slate-900 text-sm md:text-base">Month-to-Date Work Summaries</h3>
          <div class="flex flex-col sm:flex-row gap-2 w-full md:w-auto">
            <input type="month" id="filterMonthlyMonth" onchange="applyAdminMonthlyFilters()" class="px-3 py-2 border border-slate-200 rounded-xl bg-white text-slate-800 text-xs focus:outline-none focus:ring-2 focus:ring-teal-500">
            <input type="text" id="filterMonthlySearch" oninput="applyAdminMonthlyFilters()" placeholder="Search staff name/email..." class="px-3 py-2 border border-slate-200 rounded-xl bg-white text-slate-800 text-xs placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-teal-500 min-w-[200px]">
            <button onclick="resetAdminMonthlyFilters()" class="text-xs text-slate-600 hover:text-slate-900 bg-slate-100 px-3 py-2 rounded-xl transition border border-slate-300 font-bold">Reset</button>
          </div>
        </div>
        <div class="overflow-x-auto smooth-scroll w-full">
          <table class="w-full text-left border-collapse text-xs md:text-sm">
            <thead>
              <tr class="bg-slate-100 text-slate-500 uppercase text-[10px] md:text-xs font-bold tracking-wider border-b border-slate-200">
                <th class="p-3.5">Employee</th><th class="p-3.5">Month</th><th class="p-3.5">Total Worked Hours</th><th class="p-3.5">Total Overtime</th><th class="p-3.5">Days Worked</th><th class="p-3.5">Last Sync</th>
              </tr>
            </thead>
            <tbody id="monthlyTableBody" class="divide-y divide-slate-200 text-slate-700"></tbody>
          </table>
        </div>
      </div>

      <!-- TAB 5: LEAVE REQUESTS APPROVAL TABLE -->
      <div id="content-leaveTab" class="hidden bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden w-full">
        <div class="p-4 md:p-5 border-b border-slate-200 bg-slate-50">
          <h3 class="font-bold text-slate-900 text-sm md:text-base">Employee Leave Applications</h3>
        </div>
        <div class="overflow-x-auto smooth-scroll w-full">
          <table class="w-full text-left border-collapse text-xs md:text-sm">
            <thead>
              <tr class="bg-slate-100 text-slate-500 uppercase text-[10px] md:text-xs font-bold tracking-wider border-b border-slate-200">
                <th class="p-3.5">Employee</th><th class="p-3.5">Leave Details</th><th class="p-3.5">Dates</th><th class="p-3.5">Status</th><th class="p-3.5 text-right">Approval Actions</th>
              </tr>
            </thead>
            <tbody id="leaveTableBody" class="divide-y divide-slate-200 text-slate-700"></tbody>
          </table>
        </div>
      </div>

      <!-- TAB 6: STAFF PROFILES -->
      <div id="content-usersTab" class="hidden bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden w-full">
        <div class="p-4 md:p-5 border-b border-slate-200 bg-slate-50">
          <h3 class="font-bold text-slate-900 text-sm md:text-base">Employee Directory Profiles</h3>
        </div>
        <div class="overflow-x-auto smooth-scroll w-full">
          <table class="w-full text-left border-collapse text-xs md:text-sm">
            <thead>
              <tr class="bg-slate-100 text-slate-500 uppercase text-[10px] md:text-xs font-bold tracking-wider border-b border-slate-200">
                <th class="p-3.5">Employee</th><th class="p-3.5">Work Status</th><th class="p-3.5">Account Standing</th><th class="p-3.5 text-right">Action</th>
              </tr>
            </thead>
            <tbody id="usersTableBody" class="divide-y divide-slate-200 text-slate-700"></tbody>
          </table>
        </div>
      </div>
      
      <!-- TAB 7: TIMESHEET EXPORT -->
      <div id="content-exportTab" class="hidden bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden w-full">
        <div class="p-4 md:p-5 border-b border-slate-200 bg-slate-50 flex flex-col justify-between gap-2">
          <h3 class="font-bold text-slate-900 text-sm md:text-base">Timesheet Export Engine</h3>
          <p class="text-xs text-slate-500">Generate and download standardized timesheet reports for payroll and compliance.</p>
        </div>
        <div class="p-6">
          <form onsubmit="return false;" class="space-y-4 max-w-3xl">
            <div class="grid grid-cols-1 sm:grid-cols-3 gap-4">
              <div>
                <label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Select Employee</label>
                <select id="exportEmail" class="w-full px-3 py-2.5 border border-slate-200 rounded-xl bg-slate-50 text-slate-900 text-sm focus:ring-2 focus:ring-teal-500"></select>
              </div>
              <div>
                <label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Start Date</label>
                <input type="date" id="exportStart" required class="w-full px-3 py-2.5 border border-slate-200 rounded-xl bg-slate-50 text-slate-900 text-sm focus:ring-2 focus:ring-teal-500">
              </div>
              <div>
                <label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">End Date</label>
                <input type="date" id="exportEnd" required class="w-full px-3 py-2.5 border border-slate-200 rounded-xl bg-slate-50 text-slate-900 text-sm focus:ring-2 focus:ring-teal-500">
              </div>
            </div>
            
            <div id="exportAlert" class="hidden p-3 rounded-xl text-xs font-semibold"></div>
            
            <div class="flex flex-col sm:flex-row gap-3 pt-3">
              <button type="button" onclick="handleExport('pdf')" class="px-5 py-3 bg-fnb-teal-gradient hover:opacity-95 text-white rounded-xl text-sm font-bold shadow-md active:scale-95 transition-all">📄 Export Timesheet (PDF)</button>
              <button type="button" onclick="handleExport('csv')" class="px-5 py-3 bg-slate-800 hover:bg-slate-700 text-white rounded-xl text-sm font-bold shadow-md active:scale-95 transition-all">📊 Export Timesheet (CSV)</button>
              <button type="button" id="btnDownloadZip" onclick="requestZipExport()" class="px-5 py-3 bg-purple-600 hover:bg-purple-700 text-white rounded-xl text-sm font-bold shadow-md active:scale-95 transition-all flex items-center justify-center gap-2">
                <span id="spinnerZip" class="loader border-2 border-slate-400 border-t-white rounded-full w-4 h-4 hidden"></span>
                <span id="textZip">📦 Export Timesheets (.zip)</span>
              </button>
            </div>
          </form>
        </div>
      </div>
    </section>

    <!-- EDIT SUMMARY TIME MODAL WINDOW -->
    <div id="editSummaryModal" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex justify-center items-end p-0 md:items-center md:p-4 z-40">
      <div class="bg-white rounded-t-3xl md:rounded-2xl p-6 w-full max-w-sm border-t md:border border-slate-200 shadow-2xl">
        <div class="flex justify-between items-center mb-4">
          <h3 class="text-base font-bold text-slate-900 flex items-center gap-2"><span>⏱️</span> Edit Shift Times</h3>
          <button onclick="closeSummaryEditModal()" class="text-slate-400 hover:text-slate-800 text-lg font-bold">&times;</button>
        </div>
        
        <form onsubmit="handleSummaryUpdateSave(event)" class="space-y-4">
          <input type="hidden" id="editSumId">
          
          <div class="bg-slate-50 p-3 rounded-xl border border-slate-200 text-xs text-slate-500 mb-4">
            <span class="block font-bold text-slate-800" id="editSumName"></span>
            <span class="block" id="editSumDate"></span>
            <p class="mt-2 text-[10px] leading-snug">Editing timestamps will automatically recalculate Net Hours, Overtime, and aggregate into Weekly & Monthly summaries.</p>
          </div>

          <div class="grid grid-cols-2 gap-3">
            <div>
              <label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Clock In</label>
              <input type="time" step="1" id="editSumIn" class="w-full px-3 py-2.5 border border-slate-200 rounded-xl bg-white text-slate-900 text-sm focus:ring-2 focus:ring-teal-500">
            </div>
            <div>
              <label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Clock Out</label>
              <input type="time" step="1" id="editSumOut" class="w-full px-3 py-2.5 border border-slate-200 rounded-xl bg-white text-slate-900 text-sm focus:ring-2 focus:ring-teal-500">
            </div>
          </div>
          
          <div class="grid grid-cols-2 gap-3 pt-3">
            <button type="button" onclick="closeSummaryEditModal()" class="py-3 border border-slate-200 rounded-xl text-sm font-semibold hover:bg-slate-100 text-slate-700 active:scale-95 duration-150">Cancel</button>
            <button type="submit" class="py-3 bg-fnb-teal-gradient hover:opacity-95 text-white rounded-xl text-sm font-semibold shadow-md active:scale-95 duration-150">Save Changes</button>
          </div>
        </form>
      </div>
    </div>

    <!-- LEAVE REGISTRATION MODAL WINDOW -->
    <div id="leaveModal" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex justify-center items-end p-0 md:items-center md:p-4 z-40">
      <div class="bg-white rounded-t-3xl md:rounded-2xl p-6 w-full max-w-sm border-t md:border border-slate-200 shadow-2xl">
        <div class="flex justify-between items-center mb-4">
          <h3 class="text-base font-bold text-slate-900 flex items-center gap-2"><span>📅</span> Register Leave Request</h3>
          <button onclick="closeLeaveModal()" class="text-slate-400 hover:text-slate-800 text-lg font-bold">&times;</button>
        </div>
        
        <form onsubmit="handleLeaveSubmit(event)" class="space-y-4">
          <div>
            <label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Leave Type</label>
            <select id="leaveType" required class="w-full px-4 py-3 border border-slate-200 rounded-xl bg-slate-50 text-slate-900 text-sm focus:ring-2 focus:ring-teal-500">
              <option value="Annual">Annual Leave</option>
              <option value="Sick">Sick Leave</option>
              <option value="Family Responsibility">Family Responsibility Leave</option>
              <option value="Unpaid">Unpaid Leave</option>
            </select>
          </div>
          <div><label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Start Date</label><input type="date" id="leaveStartDate" required class="w-full px-4 py-3 border border-slate-200 rounded-xl bg-slate-50 text-slate-900 text-sm focus:ring-2 focus:ring-teal-500"></div>
          <div><label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">End Date</label><input type="date" id="leaveEndDate" required class="w-full px-4 py-3 border border-slate-200 rounded-xl bg-slate-50 text-slate-900 text-sm focus:ring-2 focus:ring-teal-500"></div>
          <div class="grid grid-cols-2 gap-3 pt-3">
            <button type="button" onclick="closeLeaveModal()" class="py-3 border border-slate-200 rounded-xl text-sm font-semibold hover:bg-slate-100 text-slate-700 active:scale-95 duration-150">Cancel</button>
            <button type="submit" class="py-3 bg-fnb-teal-gradient hover:opacity-95 text-white rounded-xl text-sm font-semibold shadow-md active:scale-95 duration-150">Submit Request</button>
          </div>
        </form>
      </div>
    </div>

    <!-- EDIT EMPLOYEE MODAL WINDOW -->
    <div id="editModal" class="hidden fixed inset-0 bg-slate-900/60 backdrop-blur-sm flex justify-center items-end p-0 md:items-center md:p-4 z-40">
      <div class="bg-white rounded-t-3xl md:rounded-2xl p-6 w-full max-w-sm border-t md:border border-slate-200 shadow-2xl">
        <h3 class="text-base font-bold text-slate-900 mb-4">Modify Employee Schema</h3>
        <form onsubmit="handleAdminUpdateSave(event)" class="space-y-4">
          <input type="hidden" id="editEmail">
          <div><label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">First Name</label><input type="text" id="editFirst" required class="w-full px-4 py-3 border border-slate-200 rounded-xl bg-slate-50 text-slate-900 text-sm focus:ring-2 focus:ring-teal-500"></div>
          <div><label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Last Name</label><input type="text" id="editLast" required class="w-full px-4 py-3 border border-slate-200 rounded-xl bg-slate-50 text-slate-900 text-sm focus:ring-2 focus:ring-teal-500"></div>
          <div>
            <label class="block text-[10px] font-bold uppercase tracking-wider text-slate-500 mb-1.5">Account Standing</label>
            <select id="editStatus" class="w-full px-4 py-3 border border-slate-200 rounded-xl bg-slate-50 text-slate-900 text-sm focus:ring-2 focus:ring-teal-500">
              <option value="Verified">Verified</option>
              <option value="Blocked">Blocked</option>
            </select>
          </div>
          <div><label class="block text-[10px] font-bold uppercase tracking-wider text-teal-600 mb-1.5">Reset Password (Leave blank to keep current)</label><input type="password" id="editPassword" placeholder="Enter new password" class="w-full px-4 py-3 border border-slate-200 rounded-xl bg-slate-50 text-slate-900 text-sm placeholder-slate-400 focus:ring-2 focus:ring-teal-500"></div>
          <div class="grid grid-cols-2 gap-3 pt-3">
            <button type="button" onclick="closeEditModal()" class="py-3 border border-slate-200 rounded-xl text-sm font-semibold hover:bg-slate-100 text-slate-700 active:scale-95 duration-150">Cancel</button>
            <button type="submit" class="py-3 bg-fnb-teal-gradient hover:opacity-95 text-white rounded-xl text-sm font-semibold shadow-md active:scale-95 duration-150">Save</button>
          </div>
        </form>
      </div>
    </div>
  </main>

  <script>
    let ACTIVE_SESSION_USER = null;
    let CACHED_LOGS_ARRAY = [];
    let CACHED_SUMMARIES_ARRAY = [];
    let CACHED_WEEKLY_ARRAY = [];
    let CACHED_MONTHLY_ARRAY = [];
    let CACHED_LEAVE_ARRAY = [];

    function showLoader(show, msg = "Processing...") {
      const sp = document.getElementById("globalSpinner");
      document.getElementById("spinnerMessage").innerText = msg;
      sp.classList.toggle("hidden", !show);
    }

    function toggleAuthMode(toLogin) {
      document.getElementById("authAlert").classList.add("hidden");
      document.getElementById("authTitle").innerText = toLogin ? "FNB Logsheet System" : "Create Access Account";
      document.getElementById("authSubtitle").innerText = toLogin ? "Enter credentials to authenticate workspace actions." : "Register to initiate system identity logging.";
      document.getElementById("loginForm").classList.toggle("hidden", !toLogin);
      document.getElementById("registerForm").classList.toggle("hidden", toLogin);
    }

    function handleRegisterSubmit(e) {
      e.preventDefault();
      showLoader(true, "Registering profile records...");
      const email = document.getElementById("regEmail").value;
      const first = document.getElementById("regFirst").value;
      const last = document.getElementById("regLast").value;
      const pass = document.getElementById("regPassword").value;

      google.script.run
        .withSuccessHandler(function(res) {
          showLoader(false);
          if(res.status === "success") {
            toggleAuthMode(true);
            const box = document.getElementById("authAlert");
            box.innerText = res.message;
            box.className = "mb-5 p-4 rounded-xl text-xs md:text-sm font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200 block";
            document.getElementById("registerForm").reset();
          } else {
            const box = document.getElementById("authAlert");
            box.innerText = res.message;
            box.className = "mb-5 p-4 rounded-xl text-xs md:text-sm font-semibold bg-rose-50 text-rose-700 border border-rose-200 block";
          }
        })
        .registerUser(email, first, last, pass);
    }

    function handleLoginSubmit(e) {
      e.preventDefault();
      showLoader(true, "Verifying credentials...");
      const email = document.getElementById("loginEmail").value;
      const pass = document.getElementById("loginPassword").value;

      google.script.run
        .withSuccessHandler(function(res) {
          showLoader(false);
          if(res.status === "success") {
            ACTIVE_SESSION_USER = res.user;
            document.getElementById("authView").classList.add("hidden");
            document.getElementById("loginForm").reset();
            const mainContainer = document.getElementById("mainContainer");
            if(res.accountType === "Admin") {
              mainContainer.classList.remove('max-w-md');
              mainContainer.classList.add('max-w-7xl');
              loadAdminDashboard();
            } else {
              mainContainer.classList.remove('max-w-7xl');
              mainContainer.classList.add('max-w-md');
              document.getElementById("employeeGreeting").innerText = `Hello World and Welcome, ${res.user.firstName}!`;
              document.getElementById("employeeView").classList.remove("hidden");
            }
          } else {
            const box = document.getElementById("authAlert");
            box.innerText = res.message;
            box.className = "mb-5 p-4 rounded-xl text-xs md:text-sm font-semibold bg-rose-50 text-rose-700 border border-rose-200 block";
          }
        })
        .loginUser(email, pass);
    }

    function logout() {
      ACTIVE_SESSION_USER = null;
      document.getElementById("adminView").classList.add("hidden");
      document.getElementById("employeeView").classList.add("hidden");
      document.getElementById("authView").classList.remove("hidden");
      const mainContainer = document.getElementById("mainContainer");
      mainContainer.classList.remove('max-w-7xl');
      mainContainer.classList.add('max-w-md');
      toggleAuthMode(true);
    }

    function triggerClockAction(actionType) {
      document.getElementById("logAlert").classList.add("hidden");
      showLoader(true, "Querying device location...");

      if (!navigator.geolocation) {
        showLoader(false);
        const box = document.getElementById("logAlert");
        box.innerText = "GPS services are not supported by this browser.";
        box.className = "mb-6 p-4 rounded-xl text-xs font-semibold bg-rose-50 text-rose-700 border border-rose-200 block";
        return;
      }

      navigator.geolocation.getCurrentPosition(
        function(pos) {
          showLoader(true, "Verifying telemetry coordinates...");
          google.script.run
            .withSuccessHandler(function(res) {
              showLoader(false);
              const box = document.getElementById("logAlert");
              if(res.status === "success") {
                box.innerText = `${res.action} recorded at ${res.time} at ${res.location}`;
                box.className = "mb-6 p-4 rounded-xl text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200 block";
              } else {
                box.innerText = "Engine Error: " + res.message;
                box.className = "mb-6 p-4 rounded-xl text-xs font-semibold bg-rose-50 text-rose-700 border border-rose-200 block";
              }
            })
            .processActivityLog(ACTIVE_SESSION_USER.email, ACTIVE_SESSION_USER.firstName, ACTIVE_SESSION_USER.lastName, actionType, pos.coords.latitude, pos.coords.longitude);
        },
        function(err) {
          showLoader(false);
          const box = document.getElementById("logAlert");
          box.innerText = "GPS blocked: Please enable location settings on your phone browser.";
          box.className = "mb-6 p-4 rounded-xl text-xs font-semibold bg-rose-50 text-rose-700 border border-rose-200 block";
        },
        { enableHighAccuracy: true, timeout: 10000 }
      );
    }

    function toggleAdminClockPanel() {
      const clockPanel = document.getElementById("adminSelfClockPanel");
      const toggleText = document.getElementById("adminClockToggleText");
      if (clockPanel.classList.contains("hidden")) {
        clockPanel.classList.remove("hidden");
        toggleText.innerText = "Return to Dashboard View";
      } else {
        clockPanel.classList.add("hidden");
        toggleText.innerText = "Go to My Clock In / Out";
      }
    }

    function triggerAdminClockAction(actionType) {
      document.getElementById("adminLogAlert").classList.add("hidden");
      showLoader(true, "Querying device location...");

      if (!navigator.geolocation) {
        showLoader(false);
        const box = document.getElementById("adminLogAlert");
        box.innerText = "GPS services are not supported by this browser.";
        box.className = "mb-4 p-4 rounded-xl text-xs font-semibold bg-rose-50 text-rose-700 border border-rose-200 block";
        return;
      }

      navigator.geolocation.getCurrentPosition(
        function(pos) {
          showLoader(true, "Verifying telemetry coordinates...");
          google.script.run
            .withSuccessHandler(function(res) {
              showLoader(false);
              const box = document.getElementById("adminLogAlert");
              if (res.status === "success") {
                box.innerText = `${res.action} recorded at ${res.time} at ${res.location}`;
                box.className = "mb-4 p-4 rounded-xl text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200 block";
                loadAdminDashboard();
              } else {
                box.innerText = "Engine Error: " + res.message;
                box.className = "mb-4 p-4 rounded-xl text-xs font-semibold bg-rose-50 text-rose-700 border border-rose-200 block";
              }
            })
            .processActivityLog(ACTIVE_SESSION_USER.email, ACTIVE_SESSION_USER.firstName, ACTIVE_SESSION_USER.lastName, actionType, pos.coords.latitude, pos.coords.longitude);
        },
        function(err) {
          showLoader(false);
          const box = document.getElementById("adminLogAlert");
          box.innerText = "GPS blocked: Please enable location settings on your browser.";
          box.className = "mb-4 p-4 rounded-xl text-xs font-semibold bg-rose-50 text-rose-700 border border-rose-200 block";
        },
        { enableHighAccuracy: true, timeout: 10000 }
      );
    }

    function openLeaveModal() {
      const todayStr = new Date().toISOString().split('T')[0];
      document.getElementById("leaveStartDate").value = todayStr;
      document.getElementById("leaveEndDate").value = todayStr;
      document.getElementById("leaveModal").classList.remove("hidden");
    }

    function closeLeaveModal() {
      document.getElementById("leaveModal").classList.add("hidden");
    }

    function handleLeaveSubmit(e) {
      e.preventDefault();
      showLoader(true, "Registering leave request & dispatching email...");
      const leaveType = document.getElementById("leaveType").value;
      const startDate = document.getElementById("leaveStartDate").value;
      const endDate = document.getElementById("leaveEndDate").value;

      google.script.run
        .withSuccessHandler(function(res) {
          showLoader(false);
          closeLeaveModal();
          if (res.status === "success") {
            alert(res.message);
            if (ACTIVE_SESSION_USER && document.getElementById("adminView").classList.contains("hidden") === false) {
              loadAdminDashboard();
            }
          } else {
            alert("Error: " + res.message);
          }
        })
        .submitLeaveRequest(ACTIVE_SESSION_USER.email, ACTIVE_SESSION_USER.firstName, ACTIVE_SESSION_USER.lastName, leaveType, startDate, endDate);
    }

    function loadAdminDashboard() {
      showLoader(true, "Loading Dashboard database...");
      google.script.run
        .withFailureHandler(function(err) {
          showLoader(false);
          alert("System Connection Error: " + err.message);
        })
        .withSuccessHandler(function(res) {
          showLoader(false);
          if(res.status === "success") {
            document.getElementById("adminGreeting").innerText = `${ACTIVE_SESSION_USER.firstName}'s Portal`;
            document.getElementById("adminView").classList.remove("hidden");
            
            CACHED_LOGS_ARRAY = res.logs || [];
            CACHED_SUMMARIES_ARRAY = res.summaries || [];
            CACHED_WEEKLY_ARRAY = res.weeklySummaries || [];
            CACHED_MONTHLY_ARRAY = res.monthlySummaries || [];
            CACHED_LEAVE_ARRAY = res.leaveRequests || []; 
            
            const todayStr = new Date().toISOString().split('T')[0];
            const currentMonthStr = todayStr.substring(0, 7);
            
            document.getElementById("filterDate").value = todayStr;
            document.getElementById("filterSummaryDate").value = todayStr;
            document.getElementById("filterMonthlyMonth").value = currentMonthStr;
            
            renderStaffProfilesTable(res.employees);
            renderLeaveRequestsTable(CACHED_LEAVE_ARRAY); 
            
            const exportEmailSelect = document.getElementById("exportEmail");
            exportEmailSelect.innerHTML = "";
            const optAll = document.createElement("option");
            optAll.value = "ALL";
            optAll.text = "All Employees (Zip Archive Only)";
            exportEmailSelect.appendChild(optAll);

            res.employees.forEach(emp => {
               const opt = document.createElement("option");
               opt.value = emp.email;
               opt.text = `${emp.firstName} ${emp.lastName} (${emp.email})`;
               exportEmailSelect.appendChild(opt);
            });
            
            applyAdminLogFilters(); 
            applyAdminSummaryFilters();
            applyAdminWeeklyFilters();
            applyAdminMonthlyFilters();
          } else {
            alert("Database Error: " + res.message);
          }
        })
        .getAdminDashboardData();
    }

    function switchAdminTab(tabName) {
      const tabs = ['logsTab', 'summariesTab', 'weeklyTab', 'monthlyTab', 'leaveTab', 'usersTab', 'exportTab'];
      tabs.forEach(t => {
        const btn = document.getElementById(`btn-${t}`);
        const content = document.getElementById(`content-${t}`);
        if(t === tabName) {
          btn.className = "flex-1 py-2.5 px-3 rounded-xl font-bold text-xs whitespace-nowrap bg-fnb-teal text-white shadow-sm transition-all";
          content.classList.remove("hidden");
          content.classList.add("block");
        } else {
          btn.className = "flex-1 py-2.5 px-3 rounded-xl font-bold text-xs whitespace-nowrap text-slate-500 hover:text-slate-900 bg-transparent transition-all";
          content.classList.remove("block");
          content.classList.add("hidden");
        }
      });
    }

    function requestZipExport() {
      const email = document.getElementById("exportEmail").value;
      const startDate = document.getElementById("exportStart").value;
      const endDate = document.getElementById("exportEnd").value;
      const alertBox = document.getElementById("exportAlert");
      
      const btn = document.getElementById("btnDownloadZip");
      const spinner = document.getElementById("spinnerZip");
      const text = document.getElementById("textZip");
      
      if (!startDate || !endDate) {
        showAlertUI(alertBox, "Please select both a Start Date and an End Date.", true);
        return;
      }
      
      if (new Date(startDate) > new Date(endDate)) {
        showAlertUI(alertBox, "Start Date cannot be after End Date.", true);
        return;
      }
      
      alertBox.classList.add("hidden");
      btn.classList.add("opacity-60", "cursor-not-allowed");
      btn.style.pointerEvents = "none";
      spinner.classList.remove("hidden");
      text.innerText = "Compiling Archive...";

      google.script.run
        .withFailureHandler(function(err) {
          resetZipButton();
          showAlertUI(alertBox, "Export Error: " + err.message, true);
        })
        .withSuccessHandler(function(res) {
          resetZipButton();
          if (res.status === "success") {
            triggerNativeDownload(res.base64, res.mimeType, res.filename);
            showAlertUI(alertBox, "ZIP Archive generated and downloaded successfully.", false);
          } else {
            showAlertUI(alertBox, "Export Error: " + res.message, true);
          }
        })
        .generateTimesheetsZip(email, startDate, endDate);
    }

    function resetZipButton() {
      const btn = document.getElementById("btnDownloadZip");
      const spinner = document.getElementById("spinnerZip");
      const text = document.getElementById("textZip");
      
      btn.classList.remove("opacity-60", "cursor-not-allowed");
      btn.style.pointerEvents = "auto";
      spinner.classList.add("hidden");
      text.innerText = "📦 Export Timesheets (.zip)";
    }

    function handleExport(format) {
      const email = document.getElementById("exportEmail").value;
      const startDate = document.getElementById("exportStart").value;
      const endDate = document.getElementById("exportEnd").value;
      const alertBox = document.getElementById("exportAlert");
      
      if (email === "ALL") {
        showAlertUI(alertBox, "PDF and CSV timesheets require a specific employee. Please select a single employee or use the ZIP export.", true);
        return;
      }
      
      if (!startDate || !endDate) {
        showAlertUI(alertBox, "Please select both a Start Date and an End Date.", true);
        return;
      }
      
      if (new Date(startDate) > new Date(endDate)) {
        showAlertUI(alertBox, "Start Date cannot be after End Date.", true);
        return;
      }
      
      alertBox.classList.add("hidden");
      showLoader(true, `Generating ${format.toUpperCase()} timesheet document...`);
      
      google.script.run
        .withFailureHandler(function(err) {
          showLoader(false);
          showAlertUI(alertBox, "Export Error: " + err.message, true);
        })
        .withSuccessHandler(function(res) {
          showLoader(false);
          if (res.status === "success") {
            triggerNativeDownload(res.base64, res.mimeType, res.filename);
            showAlertUI(alertBox, `${format.toUpperCase()} Export generated and downloaded successfully!`, false);
          } else {
            showAlertUI(alertBox, "Export Error: " + res.message, true);
          }
        })
        .exportTimesheetData(email, startDate, endDate, format);
    }

    function showAlertUI(el, text, isError) {
      el.innerText = text;
      el.className = `mb-4 p-3 rounded-xl text-xs font-semibold block ${isError ? 'bg-rose-50 text-rose-700 border border-rose-200' : 'bg-emerald-50 text-emerald-700 border border-emerald-200'}`;
      el.classList.remove("hidden");
    }

    function triggerNativeDownload(base64Data, contentType, filename) {
      const byteCharacters = atob(base64Data);
      const byteArrays = [];
      for (let offset = 0; offset < byteCharacters.length; offset += 512) {
        const slice = byteCharacters.slice(offset, offset + 512);
        const byteNumbers = new Array(slice.length);
        for (let i = 0; i < slice.length; i++) {
          byteNumbers[i] = slice.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        byteArrays.push(byteArray);
      }
      const blob = new Blob(byteArrays, { type: contentType });
      const link = document.createElement("a");
      const url = URL.createObjectURL(blob);
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      setTimeout(() => {
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
      }, 100);
    }

    function applyAdminLogFilters() {
      const selectedDate = document.getElementById("filterDate").value.trim();
      const searchQuery = document.getElementById("filterSearch").value.trim().toLowerCase();
      
      let totalLogins = 0; let totalLogouts = 0;

      const filteredLogs = CACHED_LOGS_ARRAY.filter(log => {
        const matchesDate = (selectedDate === "") || (log.date === selectedDate);
        const fullName = `${log.firstName} ${log.lastName}`.toLowerCase();
        const matchesSearch = (searchQuery === "") || fullName.includes(searchQuery) || log.email.toLowerCase().includes(searchQuery);
                              
        if(matchesDate && matchesSearch) {
          if (log.action === "Clock In") totalLogins++;
          if (log.action === "Clock Out") totalLogouts++;
          return true;
        }
        return false;
      });
      
      document.getElementById("statLogins").innerText = totalLogins;
      document.getElementById("statLogouts").innerText = totalLogouts;
      document.getElementById("statDateDisplay").innerText = selectedDate ? selectedDate : "All Time";
      
      const logsBody = document.getElementById("logsTableBody");
      logsBody.innerHTML = filteredLogs.length === 0 ? `<tr><td colspan="5" class="p-4 text-center text-slate-500">No matching logs found for active filter configurations.</td></tr>` : "";
      
      filteredLogs.forEach(log => {
        const row = document.createElement("tr");
        row.className = "hover:bg-slate-50 transition-colors";
        
        const coordsDisplay = (log.lat && log.lon) 
          ? `<span class="inline-block mt-1 font-mono text-[9px] text-teal-700 bg-teal-50 px-1.5 py-0.5 rounded border border-teal-200">GPS: ${log.lat}, ${log.lon}</span>`
          : `<span class="inline-block mt-1 font-mono text-[9px] text-slate-400">GPS: N/A</span>`;

        const editedByTag = log.editedBy 
          ? `<span class="inline-block mt-1 font-semibold text-[9px] text-purple-600 bg-purple-50 px-1.5 py-0.5 rounded border border-purple-200">Edited by: ${log.editedBy}</span>` 
          : "";

        row.innerHTML = `
          <td class="p-3.5 align-top"><span class="font-bold text-slate-900 block">${log.firstName} ${log.lastName}</span><span class="text-[10px] text-slate-500 block">${log.email}</span></td>
          <td class="p-3.5 align-top"><span class="text-slate-800 block font-medium">${log.date}</span><span class="text-[10px] text-slate-500 block">${log.time}</span></td>
          <td class="p-3.5 align-top"><span class="px-2.5 py-1 rounded-md text-[10px] font-bold ${log.action === 'Clock In' ? 'bg-teal-50 text-teal-700 border border-teal-200' : 'bg-orange-50 text-orange-700 border border-orange-200'}">${log.action}</span></td>
          <td class="p-3.5 align-top text-xs text-slate-600 leading-normal whitespace-normal break-words"><div class="font-medium text-slate-800">${log.location}</div>${coordsDisplay}</td>
          <td class="p-3.5 align-top">${editedByTag}</td>
        `;
        logsBody.appendChild(row);
      });
    }

    function resetAdminLogFilters() {
      const todayStr = new Date().toISOString().split('T')[0];
      document.getElementById("filterDate").value = todayStr;
      document.getElementById("filterSearch").value = "";
      applyAdminLogFilters();
    }

    function formatTimeForInput(timeStr) {
      if (!timeStr || timeStr === "--" || timeStr === "N/A") return "";
      const parts = timeStr.split(":");
      if (parts.length >= 2) {
        return `${parts[0].padStart(2, '0')}:${parts[1].padStart(2, '0')}:${parts.length > 2 ? parts[2].padStart(2, '0') : "00"}`;
      }
      return "";
    }

    function openSummaryEditModal(sumId, fullName, dateStr, clockIn, clockOut) {
      document.getElementById("editSumId").value = sumId;
      document.getElementById("editSumName").innerText = fullName;
      document.getElementById("editSumDate").innerText = "Date: " + dateStr;
      document.getElementById("editSumIn").value = formatTimeForInput(clockIn);
      document.getElementById("editSumOut").value = formatTimeForInput(clockOut);
      document.getElementById("editSummaryModal").classList.remove("hidden");
    }

    function closeSummaryEditModal() {
      document.getElementById("editSummaryModal").classList.add("hidden");
    }

    function handleSummaryUpdateSave(e) {
      e.preventDefault();
      showLoader(true, "Updating timestamps & syncing summaries...");
      const sumId = document.getElementById("editSumId").value;
      const newIn = document.getElementById("editSumIn").value;
      const newOut = document.getElementById("editSumOut").value;
      const adminEmail = ACTIVE_SESSION_USER.email;
      
      google.script.run
        .withSuccessHandler(function(res) {
          showLoader(false);
          closeSummaryEditModal();
          if (res.status === "success") {
            loadAdminDashboard();
          } else {
            alert("Error updating summary: " + res.message);
          }
        })
        .adminUpdateDailySummary(sumId, newIn, newOut, adminEmail);
    }
    
    // NEW FRONTEND FUNCTION: Delete a Shift
    function deleteDailySummary(summaryId, dateStr, email) {
      if(!confirm("Are you sure you want to completely delete this shift record? This will permanently scrub the underlying activity logs for this day as well.")) return;
      showLoader(true, "Deleting shift record...");
      google.script.run
        .withFailureHandler(function(err) {
            showLoader(false);
            alert("Delete Error: " + err.message);
        })
        .withSuccessHandler(function(res) {
          showLoader(false);
          if (res.status === "success") {
            loadAdminDashboard();
          } else {
            alert("Error: " + res.message);
          }
        })
        .adminDeleteDailySummary(summaryId, dateStr, email, ACTIVE_SESSION_USER.email);
    }

    function applyAdminSummaryFilters() {
      const selectedDate = document.getElementById("filterSummaryDate").value.trim();
      const searchQuery = document.getElementById("filterSummarySearch").value.trim().toLowerCase();
      
      const filteredSummaries = CACHED_SUMMARIES_ARRAY.filter(sum => {
        const matchesDate = (selectedDate === "") || (sum.date === selectedDate);
        const fullName = `${sum.firstName} ${sum.lastName}`.toLowerCase();
        const matchesSearch = (searchQuery === "") || fullName.includes(searchQuery) || sum.email.toLowerCase().includes(searchQuery);
        return matchesDate && matchesSearch;
      });
      renderSummariesTable(filteredSummaries);
    }

    function resetAdminSummaryFilters() {
      document.getElementById("filterSummaryDate").value = "";
      document.getElementById("filterSummarySearch").value = "";
      applyAdminSummaryFilters();
    }

    function renderSummariesTable(summaries) {
      const summaryBody = document.getElementById("summariesTableBody");
      summaryBody.innerHTML = summaries.length === 0 ? `<tr><td colspan="8" class="p-4 text-center text-slate-500">No matching daily summaries found.</td></tr>` : "";
      
      summaries.forEach(sum => {
        const row = document.createElement("tr");
        row.className = "hover:bg-slate-50 transition-colors";
        
        const statusBadge = sum.status === "Completed"
          ? `<span class="px-2.5 py-1 rounded-md text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200">Completed</span>`
          : `<span class="px-2.5 py-1 rounded-md text-[10px] font-bold bg-amber-50 text-amber-700 border border-amber-200">In Progress</span>`;

        const isOvertime = sum.overtime && sum.overtime !== "0 mins" && sum.overtime !== "N/A";
        const overtimeBadge = isOvertime
          ? `<span class="px-2.5 py-1 rounded-md text-[10px] font-bold bg-purple-50 text-purple-700 border border-purple-200">${sum.overtime}</span>`
          : `<span class="text-xs text-slate-500">${sum.overtime}</span>`;

        // Added the new Delete Button logic inline here
        row.innerHTML = `
          <td class="p-3.5"><span class="font-bold text-slate-900 block">${sum.firstName} ${sum.lastName}</span><span class="text-[10px] text-slate-500 block">${sum.email}</span></td>
          <td class="p-3.5 text-slate-700 font-medium">${sum.date}</td>
          <td class="p-3.5"><span class="text-xs text-slate-800 block">In: ${sum.clockIn}</span><span class="text-xs text-slate-500 block">Out: ${sum.clockOut}</span></td>
          <td class="p-3.5"><span class="font-mono text-xs text-teal-700 font-semibold block">In: ${sum.floorIn}</span><span class="font-mono text-xs text-orange-600 font-semibold block">Out: ${sum.floorOut}</span></td>
          <td class="p-3.5 font-semibold text-slate-900">${sum.totalHours}</td>
          <td class="p-3.5">${overtimeBadge}</td>
          <td class="p-3.5">${statusBadge}</td>
          <td class="p-3.5 text-right whitespace-nowrap">
            <button onclick="openSummaryEditModal('${sum.summaryId}', '${sum.firstName} ${sum.lastName}', '${sum.date}', '${sum.clockIn}', '${sum.clockOut}')" class="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold transition border border-slate-300 active:scale-95 shadow-sm">Edit</button>
            <button onclick="deleteDailySummary('${sum.summaryId}', '${sum.date}', '${sum.email}')" class="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 rounded-lg text-xs font-semibold transition border border-rose-200 active:scale-95 shadow-sm ml-1">Delete</button>
          </td>
        `;
        summaryBody.appendChild(row);
      });
    }

    function applyAdminWeeklyFilters() {
      const selectedWeekKey = document.getElementById("filterWeeklyKey").value.trim().toUpperCase();
      const searchQuery = document.getElementById("filterWeeklySearch").value.trim().toLowerCase();
      
      const filteredWeekly = CACHED_WEEKLY_ARRAY.filter(w => {
        const matchesWeek = (selectedWeekKey === "") || (w.weekKey.toUpperCase().includes(selectedWeekKey));
        const fullName = `${w.firstName} ${w.lastName}`.toLowerCase();
        const matchesSearch = (searchQuery === "") || fullName.includes(searchQuery) || w.email.toLowerCase().includes(searchQuery);
        return matchesWeek && matchesSearch;
      });
      renderWeeklyTable(filteredWeekly);
    }

    function resetAdminWeeklyFilters() {
      document.getElementById("filterWeeklyKey").value = "";
      document.getElementById("filterWeeklySearch").value = "";
      applyAdminWeeklyFilters();
    }

    function renderWeeklyTable(weeklyData) {
      const wBody = document.getElementById("weeklyTableBody");
      wBody.innerHTML = weeklyData.length === 0 ? `<tr><td colspan="6" class="p-4 text-center text-slate-500">No matching weekly summaries found.</td></tr>` : "";
      
      weeklyData.forEach(w => {
        const row = document.createElement("tr");
        row.className = "hover:bg-slate-50 transition-colors";

        let complianceBadge = `<span class="px-2.5 py-1 rounded-md text-[10px] font-bold bg-amber-50 text-amber-700 border border-amber-200">Under Target (<45 hrs)</span>`;
        if (w.complianceStatus.includes("MAX CAP ALERT")) complianceBadge = `<span class="px-2.5 py-1 rounded-md text-[10px] font-extrabold bg-rose-100 text-rose-700 border border-rose-300 shadow-sm animate-pulse">🔴 MAX CAP ALERT (55+ hrs)</span>`;
        else if (w.complianceStatus.includes("Target Met")) complianceBadge = `<span class="px-2.5 py-1 rounded-md text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200">Target Met (45-54 hrs)</span>`;

        const isOvertime = w.totalOvertime && w.totalOvertime !== "0 mins" && w.totalOvertime !== "N/A";
        const overtimeBadge = isOvertime
          ? `<span class="px-2.5 py-1 rounded-md text-[10px] font-bold bg-purple-50 text-purple-700 border border-purple-200">${w.totalOvertime}</span>`
          : `<span class="text-xs text-slate-500">${w.totalOvertime}</span>`;

        row.innerHTML = `
          <td class="p-3.5"><span class="font-bold text-slate-900 block">${w.firstName} ${w.lastName}</span><span class="text-[10px] text-slate-500 block">${w.email}</span></td>
          <td class="p-3.5 font-semibold text-teal-700 font-mono">${w.weekKey}</td>
          <td class="p-3.5 font-bold text-slate-900">${w.totalHours}</td>
          <td class="p-3.5">${overtimeBadge}</td>
          <td class="p-3.5 font-semibold text-slate-700">${w.daysWorked} days</td>
          <td class="p-3.5">${complianceBadge}</td>
        `;
        wBody.appendChild(row);
      });
    }

    function applyAdminMonthlyFilters() {
      const selectedMonth = document.getElementById("filterMonthlyMonth").value.trim();
      const searchQuery = document.getElementById("filterMonthlySearch").value.trim().toLowerCase();
      
      const filteredMonthly = CACHED_MONTHLY_ARRAY.filter(m => {
        const matchesMonth = (selectedMonth === "") || (m.month === selectedMonth);
        const fullName = `${m.firstName} ${m.lastName}`.toLowerCase();
        const matchesSearch = (searchQuery === "") || fullName.includes(searchQuery) || m.email.toLowerCase().includes(searchQuery);
        return matchesMonth && matchesSearch;
      });
      renderMonthlyTable(filteredMonthly);
    }

    function resetAdminMonthlyFilters() {
      document.getElementById("filterMonthlyMonth").value = "";
      document.getElementById("filterMonthlySearch").value = "";
      applyAdminMonthlyFilters();
    }

    function renderMonthlyTable(monthlyData) {
      const mBody = document.getElementById("monthlyTableBody");
      mBody.innerHTML = monthlyData.length === 0 ? `<tr><td colspan="6" class="p-4 text-center text-slate-500">No matching monthly summaries found.</td></tr>` : "";
      
      monthlyData.forEach(m => {
        const row = document.createElement("tr");
        row.className = "hover:bg-slate-50 transition-colors";

        const isOvertime = m.totalOvertime && m.totalOvertime !== "0 mins" && m.totalOvertime !== "N/A";
        const overtimeBadge = isOvertime
          ? `<span class="px-2.5 py-1 rounded-md text-[10px] font-bold bg-purple-50 text-purple-700 border border-purple-200">${m.totalOvertime}</span>`
          : `<span class="text-xs text-slate-500">${m.totalOvertime}</span>`;

        row.innerHTML = `
          <td class="p-3.5"><span class="font-bold text-slate-900 block">${m.firstName} ${m.lastName}</span><span class="text-[10px] text-slate-500 block">${m.email}</span></td>
          <td class="p-3.5 font-semibold text-teal-700">${m.month}</td>
          <td class="p-3.5 font-bold text-slate-900">${m.totalHours}</td>
          <td class="p-3.5">${overtimeBadge}</td>
          <td class="p-3.5 font-semibold text-emerald-700">${m.daysWorked} days</td>
          <td class="p-3.5 text-xs text-slate-400 font-mono">${m.lastUpdated}</td>
        `;
        mBody.appendChild(row);
      });
    }

    function renderLeaveRequestsTable(requests) {
      const leaveBody = document.getElementById("leaveTableBody");
      leaveBody.innerHTML = requests.length === 0 ? `<tr><td colspan="5" class="p-4 text-center text-slate-500">No leave applications currently on record.</td></tr>` : "";

      requests.forEach(req => {
        const row = document.createElement("tr");
        row.className = "hover:bg-slate-50 transition-colors";

        let statusClass = "bg-amber-50 text-amber-700 border border-amber-200";
        if (req.status === "Approved") statusClass = "bg-emerald-50 text-emerald-700 border border-emerald-200";
        if (req.status === "Rejected") statusClass = "bg-rose-50 text-rose-700 border border-rose-200";

        const actions = req.status === "Pending" ? `
          <div class="flex justify-end gap-2">
            <button onclick="handleLeaveAction('${req.reqId}', 'Approved')" class="px-3 py-1 bg-emerald-50 hover:bg-emerald-100 border border-emerald-200 text-emerald-700 rounded-lg text-xs font-semibold transition active:scale-95 shadow-sm">Approve</button>
            <button onclick="handleLeaveAction('${req.reqId}', 'Rejected')" class="px-3 py-1 bg-rose-50 hover:bg-rose-100 border border-rose-200 text-rose-700 rounded-lg text-xs font-semibold transition active:scale-95 shadow-sm">Reject</button>
          </div>
        ` : `<span class="text-xs text-slate-500 italic block text-right">Resolved</span>`;

        row.innerHTML = `
          <td class="p-3.5">
            <span class="font-bold text-slate-900 block">${req.firstName} ${req.lastName}</span>
            <span class="text-[10px] text-slate-500 block">${req.email}</span>
          </td>
          <td class="p-3.5 font-semibold text-slate-800">${req.type} Leave</td>
          <td class="p-3.5 text-xs text-slate-700">
            <span class="block font-medium">${req.startDate} to</span>
            <span class="block text-slate-500">${req.endDate}</span>
          </td>
          <td class="p-3.5">
            <span class="px-2.5 py-1 rounded-md text-[10px] font-bold ${statusClass}">${req.status}</span>
          </td>
          <td class="p-3.5 text-right">${actions}</td>
        `;
        leaveBody.appendChild(row);
      });
    }

    function handleLeaveAction(reqId, newStatus) {
      showLoader(true, `Updating leave application & sending decision email...`);
      google.script.run
        .withSuccessHandler(function(res) {
          showLoader(false);
          if (res.status === "success") {
            loadAdminDashboard();
          } else {
            alert("Error updating leave: " + res.message);
          }
        })
        .adminUpdateLeaveStatus(reqId, newStatus);
    }

    function renderStaffProfilesTable(employees) {
      const usersBody = document.getElementById("usersTableBody");
      usersBody.innerHTML = employees.length === 0 ? `<tr><td colspan="4" class="p-4 text-center text-slate-500">No registered employee profiles found.</td></tr>` : "";

      employees.forEach(emp => {
        const row = document.createElement("tr");
        row.className = "hover:bg-slate-50 transition-colors";

        const workStatusBadge = emp.workStatus === "On Leave"
          ? `<span class="px-2.5 py-1 rounded-md text-[10px] font-bold bg-amber-50 text-amber-700 border border-amber-200">On Leave</span>`
          : `<span class="px-2.5 py-1 rounded-md text-[10px] font-bold bg-emerald-50 text-emerald-700 border border-emerald-200">Active</span>`;

        const standingBadge = emp.accountStanding === "Verified"
          ? `<span class="px-2.5 py-1 rounded-md text-[10px] font-bold bg-teal-50 text-teal-700 border border-teal-200">Verified</span>`
          : `<span class="px-2.5 py-1 rounded-md text-[10px] font-bold bg-rose-50 text-rose-700 border border-rose-200">Blocked</span>`;

        row.innerHTML = `
          <td class="p-3.5"><span class="font-bold text-slate-900 block">${emp.firstName} ${emp.lastName}</span><span class="text-[10px] text-slate-500 block">${emp.email}</span></td>
          <td class="p-3.5">${workStatusBadge}</td>
          <td class="p-3.5">${standingBadge}</td>
          <td class="p-3.5 text-right"><button onclick="openEditModal('${emp.email}', '${emp.firstName}', '${emp.lastName}', '${emp.accountStanding}')" class="px-3 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold transition border border-slate-300 active:scale-95 shadow-sm">Edit</button></td>
        `;
        usersBody.appendChild(row);
      });
    }

    function openEditModal(email, first, last, status) {
      document.getElementById("editEmail").value = email;
      document.getElementById("editFirst").value = first;
      document.getElementById("editLast").value = last;
      document.getElementById("editStatus").value = status;
      document.getElementById("editPassword").value = "";
      document.getElementById("editModal").classList.remove("hidden");
    }

    function closeEditModal() {
      document.getElementById("editModal").classList.add("hidden");
    }

    function handleAdminUpdateSave(e) {
      e.preventDefault();
      showLoader(true, "Updating employee profile...");
      const email = document.getElementById("editEmail").value;
      const first = document.getElementById("editFirst").value;
      const last = document.getElementById("editLast").value;
      const status = document.getElementById("editStatus").value;
      const newPassword = document.getElementById("editPassword").value;

      google.script.run
        .withSuccessHandler(function(res) {
          showLoader(false);
          closeEditModal();
          if (res.status === "success") {
            loadAdminDashboard();
          } else {
            alert("Error updating profile: " + res.message);
          }
        })
        .adminUpdateUser(email, first, last, status, newPassword);
    }
  </script>
</body>
</html>