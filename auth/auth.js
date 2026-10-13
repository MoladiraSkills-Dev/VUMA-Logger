// ============================================================================
// LOGIN / REGISTER PAGE LOGIC
// Requires shared/js/api.js
// After a successful login the user is routed to employee/ or admin/.
// ============================================================================

const EMPLOYEE_PAGE = "employee/index.html";
const ADMIN_PAGE = "admin/index.html";

// If a session already exists, skip the login screen.
(function redirectIfLoggedIn() {
  const session = getSession();
  if (session) {
    window.location.replace(session.accountType === "Admin" ? ADMIN_PAGE : EMPLOYEE_PAGE);
  }
})();

function toggleAuthMode(toLogin) {
  document.getElementById("authAlert").classList.remove("visible");
  document.getElementById("authTitle").innerHTML = toLogin
    ? 'VUMA <span style="font-weight:300;color:var(--vuma-magenta);">Logger</span>'
    : 'Create <span style="font-weight:300;color:var(--vuma-magenta);">Account</span>';
  document.getElementById("authSubtitle").innerText = toLogin
    ? "Enter credentials to access the workforce portal."
    : "Register to initiate attendance logging.";
  document.getElementById("loginForm").classList.toggle("hidden", !toLogin);
  document.getElementById("registerForm").classList.toggle("hidden", toLogin);
}

function handleRegisterSubmit(e) {
  e.preventDefault();
  showLoader(true, "Registering profile...");
  google.script.run
    .withSuccessHandler(function (res) {
      showLoader(false);
      if (res.status === "success") {
        toggleAuthMode(true);
        showAlert("authAlert", res.message, false);
        document.getElementById("registerForm").reset();
      } else {
        showAlert("authAlert", res.message, true);
      }
    })
    .registerUser(
      document.getElementById("regEmail").value,
      document.getElementById("regFirst").value,
      document.getElementById("regLast").value,
      document.getElementById("regPassword").value
    );
}

function handleLoginSubmit(e) {
  e.preventDefault();
  showLoader(true, "Verifying credentials...");
  google.script.run
    .withSuccessHandler(function (res) {
      if (res.status === "success") {
        saveSession(res.user, res.accountType);
        window.location.href = res.accountType === "Admin" ? ADMIN_PAGE : EMPLOYEE_PAGE;
      } else {
        showLoader(false);
        showAlert("authAlert", res.message, true);
      }
    })
    .loginUser(document.getElementById("loginEmail").value, document.getElementById("loginPassword").value);
}
