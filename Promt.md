# System Role & Objective
You are an expert full-stack developer specializing in Google Apps Script (Backend/API) and modern vanilla JavaScript + CSS (Frontend). I need you to build a comprehensive Geolocation Attendance, Leave, and Device Management System.

## Architecture Constraint (Decoupled Deployment for Security)

**Frontend:** A static web application (HTML/JS/CSS) that will be hosted on GitHub Pages. It must communicate with the backend via asynchronous `fetch()` API calls using `POST` requests.
**Backend:** A Google Apps Script (`Code.gs`) deployed as a Web App executing as the script owner (Who has access: "Anyone"). It must use a `doPost(e)` function to route incoming API requests and return JSON responses.
**Security:** All API keys, Google Sheet references, hashing logic, and database operations MUST remain hidden in `Code.gs`. The frontend should only know the Web App URL.

## Step 1: Provide the Database Initialization Script
Before writing the core logic, provide a standalone Apps Script function named `setupDatabase()` that I can run once to automatically create all the necessary Google Sheets and columns. The schema must be exactly:

- **Users:** [Email, First Name, Last Name, Password Hash, Salt, Account Type, Status, Verification Token]
- **Activity_Logs:** [Log ID, Email, First Name, Last Name, Date, Time, Action, Latitude, Longitude, Resolved Location Name, Edited By]
- **Daily_Summaries:** [First Name, Last Name, Summary ID, Date, Email, Clock In Time, Clock Out Time, Clock-In Floor Time, Clock-Out Floor Time, Total Hours worked, Status, Worked Overtime]
- **Weekly_Summaries:** [First Name, Last Name, Email, Week Identifier, Total Hours Worked, Total Overtime, Days Worked, Compliance Status, Last Updated]
- **Monthly_Summaries:** [First Name, Last Name, Email, Month, Total Hours Worked, Total Overtime, Days Worked, Last Updated]
- **Leave_Requests:** [Request ID, Email, First Name, Last Name, Leave Type, Start Date, End Date, Status, Submitted Date]

## Step 2: Core Business Logic Specifications (Backend)
Write the `Code.gs` backend incorporating the following strict business rules:

1. **Custom Authentication:** Validate users against the Users sheet using SHA-256 password hashing with unique salts.
2. **Geolocation Logging:** On Clock In/Out, accept GPS coordinates from the frontend, reverse-geocode them, and save the physical address.
3. **Standardized Floor Times (Crucial):**
   - Regardless of actual Clock-In time, calculate hours starting from standard shifts (e.g., 08:00 AM).
   - Deduct exactly 1 hour for lunch if standard work minutes exceed 5 hours.
4. **Overtime Lockdown:** Normal clock-outs must lock overtime to "0 mins". Overtime can only be awarded if an Admin explicitly overrides the shift via the dashboard (flagging `isAdminEdit = true`).
5. **Auto-Failsafes:** If a shift is "In Progress" from a previous day, auto-log them out at the standard end time. If a user clocks out without a clock-in record for today, auto-log them in before processing the clock-out.
6. **API Routing:** Use a centralized `handleApiRequest` function. **CRITICAL:** Handle CORS correctly by parsing data from `e.parameter.payload` inside `doPost(e)`. Do NOT rely on `e.postData.contents`. Return responses using `ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON)`.

## Step 3: Frontend Requirements (GitHub Pages)
Write the `index.html` and `styles.css` files assuming they will be hosted externally on GitHub Pages:

1. **Design System:** Use pure vanilla CSS (`styles.css`). Create a rich, premium, modern design aesthetic using CSS variables for a specific brand's color palette. Do NOT use Tailwind CSS.
2. **Mobile Responsiveness:** Ensure the HTML has the `<meta name="viewport" content="width=device-width, initial-scale=1.0">` tag. Ensure all CSS grid/flexbox layouts stack appropriately on mobile devices.
3. **Routing:** Create a robust UI router using JavaScript to switch between the Login/Register view, the Employee Portal, and the Admin Dashboard seamlessly without reloading the page.
4. **API Communication:** Use the `fetch()` API with `method: "POST"`, `headers: { "Content-Type": "application/x-www-form-urlencoded" }`, and pass a `URLSearchParams` object containing the `payload`. This guarantees the request bypasses strict CORS blocks and ORB blockages while securing credentials. DO NOT use `google.script.run`.

---

### 🛠️ The Database Initialization Script
*(You can use this right now to instantly set up your blank Google Sheets for any future logger! Just paste this into a blank Code.gs file and hit Run).*

```javascript
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
    "Leave_Requests": ["Request ID", "Email", "First Name", "Last Name", "Leave Type", "Start Date", "End Date", "Status", "Submitted Date"]
  };

  for (const [sheetName, headers] of Object.entries(schemas)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    } else {
      sheet.clear();
    }
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f8fafc");
    headerRange.setBorder(true, true, true, true, false, false);
    
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  const sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }

  return "Database Setup Complete! All tables have been generated with formatting.";
}
```