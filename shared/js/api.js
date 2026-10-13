// ============================================================================
// SHARED: GOOGLE APPS SCRIPT API CLIENT + SESSION HELPERS
// Used by: login (index.html), employee/, admin/
// ============================================================================

// ⚠️ IMPORTANT: Replace this with your Google Apps Script Web App URL!
const APPS_SCRIPT_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwWruO5BhX6LGfHK44Rfd8f9_861rHmnE52dJmuvq8VYAjVPGFZKPOy0skvI9qcKhW8hg/exec";

const ApiClient = {
  execute(methodName, args) {
    return new Promise((resolve, reject) => {
      if (APPS_SCRIPT_WEB_APP_URL.includes("YOUR_DEPLOYMENT_ID")) {
        return reject(new Error("You must set the APPS_SCRIPT_WEB_APP_URL in shared/js/api.js to your actual Web App URL."));
      }

      const payload = JSON.stringify({ method: methodName, args: args });
      const formData = new URLSearchParams();
      formData.append("payload", payload);

      // We use POST with x-www-form-urlencoded to prevent credentials from appearing in the URL.
      // This avoids the CORS preflight (OPTIONS) request, allowing Google Apps Script to process it natively.
      fetch(APPS_SCRIPT_WEB_APP_URL, {
        method: "POST",
        headers: {
          "Content-Type": "application/x-www-form-urlencoded"
        },
        body: formData.toString()
      })
        .then(response => {
          if (!response.ok) throw new Error("Network request failed");
          return response.json();
        })
        .then(data => resolve(data))
        .catch(err => reject(err));
    });
  }
};

// ============================================================================
// MOCK RUNNER OBJECT TO INTERCEPT CALLS
// ============================================================================
const MockRunner = {
  _successFn: null,
  _failFn: null,
  withSuccessHandler(fn) { this._successFn = fn; return this; },
  withFailureHandler(fn) { this._failFn = fn; return this; },
  async _execute(methodName, args) {
    try {
      const result = await ApiClient.execute(methodName, args);
      if (this._successFn) this._successFn(result);
    } catch (err) {
      if (this._failFn) this._failFn(err);
    }
  }
};

// List of backend functions we want to support
const backendMethods = [
  "registerUser", "loginUser", "processActivityLog", "getAdminDashboardData",
  "submitLeaveRequest", "adminUpdateLeaveStatus", "adminUpdateUser",
  "adminDeleteDailySummary", "adminUpdateDailySummary", "exportTimesheetData",
  "generateTimesheetsZip"
];

// Dynamically bind ApiClient methods to MockRunner
backendMethods.forEach(method => {
  MockRunner[method] = function (...args) {
    this._execute(method, args);
    return this;
  };
});

// Provide the global google object
const google = {
  script: {
    get run() {
      // Return a fresh runner instance so callbacks don't overwrite each other
      return Object.assign({}, MockRunner);
    }
  }
};

// ============================================================================
// SESSION (persists across the login -> employee/admin page navigation)
// ============================================================================
const SESSION_KEY = "vumaLoggerSession";
let ACTIVE_SESSION_USER = null;

function saveSession(user, accountType) {
  sessionStorage.setItem(SESSION_KEY, JSON.stringify({ user: user, accountType: accountType }));
}

function getSession() {
  try { return JSON.parse(sessionStorage.getItem(SESSION_KEY)); }
  catch (e) { return null; }
}

/**
 * Page guard. Call at the top of a protected page.
 * Redirects to the login page when there is no session or the role does not match.
 * Returns the logged-in user object when access is allowed.
 */
function requireSession(requiredRole, loginUrl) {
  const session = getSession();
  const isAdmin = session && session.accountType === "Admin";
  const allowed = session && (requiredRole === "Admin" ? isAdmin : !isAdmin);
  if (!allowed) {
    window.location.replace(loginUrl);
    return null;
  }
  ACTIVE_SESSION_USER = session.user;
  return session.user;
}

function logout(loginUrl) {
  ACTIVE_SESSION_USER = null;
  sessionStorage.removeItem(SESSION_KEY);
  window.location.href = loginUrl;
}

// ============================================================================
// SHARED UI HELPERS
// ============================================================================
function showLoader(show, msg = "Processing...") {
  const sp = document.getElementById("globalSpinner");
  document.getElementById("spinnerMessage").innerText = msg;
  if (show) sp.classList.add("active");
  else sp.classList.remove("active");
}

function showAlert(id, msg, isError) {
  const el = document.getElementById(id);
  el.innerText = msg;
  el.className = `alert visible ${isError ? 'alert-error' : 'alert-success'}`;
}
