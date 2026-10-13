// ============================================================================
// SHARED: LEAVE REQUEST MODAL (used by both employee and admin pages)
// Requires shared/js/api.js
// ============================================================================

function openLeaveModal() {
  const today = new Date().toISOString().split('T')[0];
  document.getElementById("leaveStartDate").value = today;
  document.getElementById("leaveEndDate").value = today;
  document.getElementById("leaveModal").classList.add("active");
}

function closeLeaveModal() { document.getElementById("leaveModal").classList.remove("active"); }

function handleLeaveSubmit(e) {
  e.preventDefault();
  showLoader(true, "Submitting leave request...");
  google.script.run
    .withSuccessHandler(function (res) {
      showLoader(false); closeLeaveModal();
      if (res.status === "success") {
        alert(res.message);
        // Admin page defines loadAdminDashboard; the employee page does not.
        if (typeof loadAdminDashboard === "function") loadAdminDashboard();
      } else { alert("Error: " + res.message); }
    })
    .submitLeaveRequest(ACTIVE_SESSION_USER.email, ACTIVE_SESSION_USER.firstName, ACTIVE_SESSION_USER.lastName,
      document.getElementById("leaveType").value, document.getElementById("leaveStartDate").value, document.getElementById("leaveEndDate").value);
}
