System Role & Objective
You are an expert full-stack developer specializing in Google Apps Script (Backend/API) and modern vanilla JavaScript + Tailwind CSS (Frontend). I need you to build a comprehensive Geolocation Attendance, Leave, and Device Management System.

Architecture Constraint (Decoupled Deployment for Security)

Frontend: A single-page application (HTML/JS/Tailwind) that will be hosted on GitHub Pages. It must communicate with the backend via asynchronous fetch() API calls using POST requests.

Backend: A Google Apps Script (Code.gs) deployed as a Web App executing as the script owner. It must use a doPost(e) function to route incoming JSON requests and return JSON responses.

Security: All API keys (e.g., LocationIQ API), Google Sheet references, hashing logic, and database operations MUST remain hidden in Code.gs. The frontend should only know the Web App URL.

Step 1: Provide the Database Initialization Script
Before writing the core logic, provide a standalone Apps Script function named setupDatabase() that I can run once to automatically create all the necessary Google Sheets and columns. The schema must be exactly:

Users: [Email, First Name, Last Name, Password Hash, Salt, Account Type, Status, Verification Token]

Activity_Logs: [Log ID, Email, First Name, Last Name, Date, Time, Action, Latitude, Longitude, Resolved Location Name, Edited By]

Daily_Summaries: [First Name, Last Name, Summary ID, Date, Email, Clock In Time, Clock Out Time, Clock-In Floor Time, Clock-Out Floor Time, Total Hours worked, Status, Worked Overtime]

Weekly_Summaries: [First Name, Last Name, Email, Week Identifier, Total Hours Worked, Total Overtime, Days Worked, Compliance Status, Last Updated]

Monthly_Summaries: [First Name, Last Name, Email, Month, Total Hours Worked, Total Overtime, Days Worked, Last Updated]

Leave_Requests: [Request ID, Email, First Name, Last Name, Leave Type, Start Date, End Date, Status, Submitted Date]

Device_Requests: [Request ID, Email, First Name, Last Name, Device Details, Start Date, End Date, Reason, Status, Submitted Date]

Step 2: Core Business Logic Specifications (Backend)
Write the Code.gs backend incorporating the following strict business rules:

Custom Authentication: Validate users against the Users sheet using SHA-256 password hashing with unique salts.

Geolocation Logging: On Clock In/Out, accept GPS coordinates from the frontend, reverse-geocode them using LocationIQ, and save the physical address.

Standardized Floor Times (Crucial):

Regardless of actual Clock-In time, the math engine must strictly calculate hours starting from 08:00 AM.

The standard Clock-Out floor is 17:00 PM (Weekdays) and 13:00 PM (Saturdays).

Deduct exactly 1 hour for lunch if standard work minutes exceed 5 hours.

Overtime Lockdown: Normal clock-outs must lock overtime to "0 mins". Overtime can only be awarded if an Admin explicitly overrides the shift via the dashboard (flagging isAdminEdit = true).

Auto-Failsafes:

Forgot to Clock Out: If a shift is "In Progress" from a previous day, auto-log them out at 17:00 (13:00 Sat) with location "System Auto-Logout".

Forgot to Clock In: If a user clocks out without a clock-in record for today, auto-log them in at 08:00 AM before processing the clock-out.

Admin Dashboard Features: Admins must be able to edit specific shift times (which permanently updates the Edited By audit log and awards accurate Overtime), delete shifts, and approve/reject Leave and Device requests.

Timesheet Export Engine: The backend must be able to generate Timesheets in CSV and PDF formats (masking overtime if the shift wasn't admin-edited). It must also feature a batch exporter that zips all employee PDFs into a single .zip file using Utilities.zip().

Step 3: Frontend Requirements (GitHub Pages)
Write the Index.html file assuming it will be hosted externally:

Use Tailwind CSS via CDN. Implement a dark mode aesthetic (bg-slate-950).

Create a robust UI router using JavaScript to switch between the Login/Register view, the standard Employee Portal, and the Admin Dashboard.

Use navigator.geolocation.getCurrentPosition before firing the Clock In/Out payload.

Handle CORS correctly by ensuring all requests sent to the Apps Script Web App use method: "POST", mode: "no-cors" or properly format JSON payloads so Apps Script can read e.postData.contents.

🛠️ The Database Initialization Script
(You can use this right now to instantly set up your blank Google Sheets for any future logger! Just paste this into a blank Code.gs file and hit Run).

JavaScript
/**
 * Run this function ONCE in a new Google Sheet to instantly 
 * build the entire Database Architecture for the Logger.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const schemas = {
    "Users": ["Email", "First Name", "Last Name", "Password Hash", "Salt", "Account Type", "Status", "Verification Token"],
    "Activity_Logs": ["Log ID", "Email", "First Name", "Last Name", "Date", "Time", "Action", "Latitude", "Longitude", "Resolved Location Name", "Edited By"],
    "Daily_Summaries": ["First Name", "Last Name", "Summary ID", "Date", "Email", "Clock In Time", "Clock Out Time", "Clock-In Floor Time", "Clock-Out Floor Time", "Total Hours worked", "Status", "Worked Overtime"],
    "Weekly_Summaries": ["First Name", "Last Name", "Email", "Week Identifier", "Total Hours Worked", "Total Overtime", "Days Worked", "Compliance Status", "Last Updated"],
    "Monthly_Summaries": ["First Name", "Last Name", "Email", "Month", "Total Hours Worked", "Total Overtime", "Days Worked", "Last Updated"],
    "Leave_Requests": ["Request ID", "Email", "First Name", "Last Name", "Leave Type", "Start Date", "End Date", "Status", "Submitted Date"],
    "Device_Requests": ["Request ID", "Email", "First Name", "Last Name", "Device Details", "Start Date", "End Date", "Reason", "Status", "Submitted Date"],
    "Admin_Settings": ["Admin Email", "Role Description", "Created Date"]
  };

  for (const [sheetName, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(sheetName);
    
    // If the sheet doesn't exist, create it
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      // If it exists but might be dirty, clear it completely
      sheet.clear();
    }
    
    // Apply exact headers
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Apply styling to headers
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f8fafc");
    headerRange.setBorder(true, true, true, true, false, false);
    
    // Freeze the top row so headers follow as you scroll
    sheet.setFrozenRows(1);
    
    // Auto-resize columns for better visibility
    sheet.autoResizeColumns(1, headers.length);
  }

  // Delete the default "Sheet1" if it exists to keep things clean
  const sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }

  return "Database Setup Complete! All 8 tables have been generated with formatting.";
}