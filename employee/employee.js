// ============================================================================
// EMPLOYEE PORTAL LOGIC
// Requires shared/js/api.js and shared/js/leave.js
// ============================================================================

const LOGIN_URL = "../index.html";

(function initEmployeePage() {
  const user = requireSession("Employee", LOGIN_URL);
  if (!user) return;
  document.getElementById("employeeGreeting").innerText = `Welcome, ${user.firstName}!`;
})();

function triggerClockAction(actionType) {
  document.getElementById("logAlert").className = "alert";
  showLoader(true, "Querying device location...");
  if (!navigator.geolocation) {
    showLoader(false);
    showAlert("logAlert", "GPS services are not supported by this browser.", true);
    return;
  }
  navigator.geolocation.getCurrentPosition(
    function (pos) {
      showLoader(true, "Verifying telemetry coordinates...");
      google.script.run
        .withSuccessHandler(function (res) {
          showLoader(false);
          if (res.status === "success") {
            showAlert("logAlert", `${res.action} recorded at ${res.time} at ${res.location}`, false);
          } else {
            showAlert("logAlert", res.message, true);
          }
        })
        .processActivityLog(ACTIVE_SESSION_USER.email, ACTIVE_SESSION_USER.firstName, ACTIVE_SESSION_USER.lastName, actionType, pos.coords.latitude, pos.coords.longitude);
    },
    function () {
      showLoader(false);
      showAlert("logAlert", "GPS blocked: Please enable location settings.", true);
    },
    { enableHighAccuracy: true, timeout: 10000 }
  );
}
