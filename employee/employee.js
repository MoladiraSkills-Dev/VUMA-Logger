// ============================================================================
// EMPLOYEE PORTAL LOGIC
// Requires shared/js/api.js and shared/js/leave.js
// ============================================================================

const LOGIN_URL = "../index.html";
let CURRENT_TODAY_SHIFT = null;

(function initEmployeePage() {
  const user = requireSession("Employee", LOGIN_URL);
  if (!user) return;
  document.getElementById("employeeGreeting").innerText = `Welcome, ${user.firstName}!`;

  // Hydrate instantly from stored session if present
  const session = getSession();
  if (session && session.todayShift) {
    CURRENT_TODAY_SHIFT = session.todayShift;
    updateEmployeeTactileStates(session.todayShift);
  }
  if (session && session.latestLeave) {
    updateEmployeeLeaveBadge(session.latestLeave);
  }

  // Always refresh live status from server
  refreshEmployeeShiftStatus();
})();

function updateEmployeeLeaveBadge(leave) {
  const badge = document.getElementById("employeeLeaveBadge");
  const card = document.getElementById("employeeLeaveStatusCard");
  const cardText = document.getElementById("employeeLeaveStatusText");
  if (!badge) return;

  if (!leave || !leave.status) {
    badge.classList.add("hidden");
    if (card) card.classList.add("hidden");
    return;
  }

  const status = String(leave.status).trim();
  badge.classList.remove("hidden");

  if (status.toLowerCase() === "pending") {
    badge.className = "badge badge-warning";
    badge.innerText = "Application Pending";
    if (card && cardText) {
      card.classList.remove("hidden");
      cardText.innerHTML = `⏳ Your <strong>${leave.type} Leave</strong> application (${leave.startDate} to ${leave.endDate}) is awaiting admin approval.`;
    }
  } else if (status.toLowerCase() === "approved") {
    badge.className = "badge badge-success";
    badge.innerText = "Leave Approved";
    if (card && cardText) {
      card.classList.remove("hidden");
      cardText.innerHTML = `✅ <strong>${leave.type} Leave</strong> approved from ${leave.startDate} to ${leave.endDate}.`;
    }
  } else if (status.toLowerCase() === "rejected") {
    badge.className = "badge badge-danger";
    badge.innerText = "Application Declined";
    if (card && cardText) {
      card.classList.remove("hidden");
      cardText.innerHTML = `❌ <strong>${leave.type} Leave</strong> (${leave.startDate} to ${leave.endDate}) was declined by admin.`;
    }
  } else {
    badge.classList.add("hidden");
    if (card) card.classList.add("hidden");
  }
}

function getEffectiveFloorInTime(timeStr) {
  if (!timeStr || timeStr === "--" || timeStr === "N/A") return "--";
  const cleanStr = String(timeStr).trim();
  const isPM = cleanStr.toUpperCase().includes("PM");
  const isAM = cleanStr.toUpperCase().includes("AM");
  const matches = cleanStr.match(/\d+/g);
  if (matches && matches.length >= 2) {
    let hours = parseInt(matches[0], 10);
    const mins = parseInt(matches[1], 10);
    if (isPM && hours < 12) hours += 12;
    if (isAM && hours === 12) hours = 0;
    if (hours < 8 || (hours === 8 && mins === 0)) {
      return "08:00";
    }
    return `${String(hours).padStart(2, '0')}:${String(mins).padStart(2, '0')}`;
  }
  return cleanStr;
}

function updateEmployeeTactileStates(shift) {
  const btnIn = document.getElementById("btnClockIn");
  const btnOut = document.getElementById("btnClockOut");
  const inSub = document.getElementById("clockInSubtext");
  const outSub = document.getElementById("clockOutSubtext");
  const inBadge = document.getElementById("clockInBadge");
  const outBadge = document.getElementById("clockOutBadge");

  if (!btnIn || !btnOut) return;

  const hasIn = shift && (shift.hasClockedIn || (shift.clockIn && shift.clockIn !== "--"));
  const hasOut = shift && (shift.hasClockedOut || (shift.clockOut && shift.clockOut !== "--"));
  const isCompleted = shift && (shift.status === "Completed" || hasOut);

  if (isCompleted) {
    btnIn.disabled = true;
    btnIn.title = "Shift finalized for today";
    if (inSub) inSub.innerText = "Shift finalized for today";
    if (inBadge) {
      inBadge.classList.remove("hidden");
      inBadge.innerText = `✓ Clocked In (${shift.clockIn})`;
    }

    btnOut.disabled = true;
    btnOut.title = "Shift finalized for today";
    if (outSub) outSub.innerText = `Shift finalized (${shift.clockOut})`;
    if (outBadge) {
      outBadge.classList.remove("hidden");
      outBadge.innerText = "✓ Completed";
    }
  } else if (hasIn) {
    btnIn.disabled = true;
    btnIn.title = `Already clocked in at ${shift.clockIn}`;
    const effectiveFloor = (shift.floorIn && shift.floorIn !== "--") ? shift.floorIn : getEffectiveFloorInTime(shift.clockIn);
    if (inSub) inSub.innerText = `Shift active since ${effectiveFloor}`;
    if (inBadge) {
      inBadge.classList.remove("hidden");
      inBadge.innerText = `✓ In at ${shift.clockIn}`;
    }

    btnOut.disabled = false;
    btnOut.title = "Conclude your shift for today";
    if (outSub) outSub.innerText = "Conclude daily shift & GPS";
    if (outBadge) {
      outBadge.classList.add("hidden");
    }
  } else {
    // Not clocked in yet: Clock Out strictly disabled
    btnIn.disabled = false;
    btnIn.title = "Clock in to start your shift";
    if (inSub) inSub.innerText = "Log morning shift start with GPS";
    if (inBadge) inBadge.classList.add("hidden");

    btnOut.disabled = true;
    btnOut.title = "You cannot clock out until you have clocked in for today";
    if (outSub) outSub.innerText = "Clock in first to enable clock out";
    if (outBadge) {
      outBadge.classList.remove("hidden");
      outBadge.innerText = "Clock-In Required";
    }
  }
}

function refreshEmployeeShiftStatus() {
  if (!ACTIVE_SESSION_USER || !ACTIVE_SESSION_USER.email) return;
  google.script.run
    .withSuccessHandler(function (res) {
      if (res && res.status === "success") {
        if (res.todayShift) {
          CURRENT_TODAY_SHIFT = res.todayShift;
          updateEmployeeTactileStates(res.todayShift);
        }
        if (res.latestLeave !== undefined) {
          updateEmployeeLeaveBadge(res.latestLeave);
        }
        const session = getSession();
        if (session) {
          if (res.todayShift) session.todayShift = res.todayShift;
          if (res.latestLeave !== undefined) session.latestLeave = res.latestLeave;
          sessionStorage.setItem(SESSION_KEY, JSON.stringify(session));
        }
      }
    })
    .getTodayShiftStatus(ACTIVE_SESSION_USER.email);
}

function triggerClockAction(actionType) {
  document.getElementById("logAlert").className = "alert";

  // Strict client-side guard: block Clock Out immediately if not clocked in today
  if (actionType === "Clock Out") {
    const hasIn = CURRENT_TODAY_SHIFT && (CURRENT_TODAY_SHIFT.hasClockedIn || (CURRENT_TODAY_SHIFT.clockIn && CURRENT_TODAY_SHIFT.clockIn !== "--"));
    if (!hasIn) {
      showAlert("logAlert", "Action Blocked: You cannot clock out because you have not clocked in for today yet. Please clock in first.", true);
      return;
    }
  }

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
            refreshEmployeeShiftStatus();
          } else {
            showAlert("logAlert", res.message, true);
          }
        })
        .withFailureHandler(function (err) {
          showLoader(false);
          showAlert("logAlert", "Server communication error: " + (err.message || err), true);
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
