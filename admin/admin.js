// ============================================================================
// ADMIN DASHBOARD LOGIC
// Requires shared/js/api.js and shared/js/leave.js
// ============================================================================

const LOGIN_URL = "../index.html";

let CACHED_LOGS_ARRAY = [];
let CACHED_SUMMARIES_ARRAY = [];
let CACHED_WEEKLY_ARRAY = [];
let CACHED_MONTHLY_ARRAY = [];
let CACHED_LEAVE_ARRAY = [];

(function initAdminPage() {
  if (!requireSession("Admin", LOGIN_URL)) return;
  loadAdminDashboard();
})();

function toggleAdminClockPanel() {
  const panel = document.getElementById("adminSelfClockPanel");
  const txt = document.getElementById("adminClockToggleText");
  if (panel.classList.contains("hidden")) { panel.classList.remove("hidden"); txt.innerText = "Return to Dashboard View"; }
  else { panel.classList.add("hidden"); txt.innerText = "Go to My Clock In / Out"; }
}

function triggerAdminClockAction(actionType) {
  document.getElementById("adminLogAlert").className = "alert";
  showLoader(true, "Querying device location...");
  if (!navigator.geolocation) { showLoader(false); showAlert("adminLogAlert", "GPS not supported.", true); return; }
  navigator.geolocation.getCurrentPosition(
    function (pos) {
      showLoader(true, "Verifying coordinates...");
      google.script.run
        .withSuccessHandler(function (res) {
          showLoader(false);
          if (res.status === "success") { showAlert("adminLogAlert", `${res.action} recorded at ${res.time} at ${res.location}`, false); loadAdminDashboard(); }
          else { showAlert("adminLogAlert", res.message, true); }
        })
        .processActivityLog(ACTIVE_SESSION_USER.email, ACTIVE_SESSION_USER.firstName, ACTIVE_SESSION_USER.lastName, actionType, pos.coords.latitude, pos.coords.longitude);
    },
    function () { showLoader(false); showAlert("adminLogAlert", "GPS blocked.", true); },
    { enableHighAccuracy: true, timeout: 10000 }
  );
}

function loadAdminDashboard() {
  showLoader(true, "Loading dashboard...");
  google.script.run
    .withFailureHandler(function (err) { showLoader(false); alert("Error: " + err.message); })
    .withSuccessHandler(function (res) {
      showLoader(false);
      if (res.status === "success") {
        document.getElementById("adminGreeting").innerText = `${ACTIVE_SESSION_USER.firstName}'s Portal`;
        document.getElementById("adminView").classList.remove("hidden");

        CACHED_LOGS_ARRAY = res.logs || [];
        CACHED_SUMMARIES_ARRAY = res.summaries || [];
        CACHED_WEEKLY_ARRAY = res.weeklySummaries || [];
        CACHED_MONTHLY_ARRAY = res.monthlySummaries || [];
        CACHED_LEAVE_ARRAY = res.leaveRequests || [];

        const today = new Date().toISOString().split('T')[0];
        document.getElementById("filterDate").value = today;
        document.getElementById("filterSummaryDate").value = today;
        document.getElementById("filterMonthlyMonth").value = today.substring(0, 7);

        renderStaffProfilesTable(res.employees);
        renderLeaveRequestsTable(CACHED_LEAVE_ARRAY);

        const sel = document.getElementById("exportEmail");
        sel.innerHTML = "";
        const optAll = document.createElement("option"); optAll.value = "ALL"; optAll.text = "All Employees (ZIP Only)"; sel.appendChild(optAll);
        res.employees.forEach(emp => {
          const opt = document.createElement("option"); opt.value = emp.email; opt.text = `${emp.firstName} ${emp.lastName} (${emp.email})`; sel.appendChild(opt);
        });

        applyAdminLogFilters();
        applyAdminSummaryFilters();
        applyAdminWeeklyFilters();
        applyAdminMonthlyFilters();
      } else { alert("Error: " + res.message); }
    })
    .getAdminDashboardData();
}

function switchAdminTab(tabName) {
  const tabs = ['logsTab', 'summariesTab', 'weeklyTab', 'monthlyTab', 'leaveTab', 'usersTab', 'exportTab'];
  tabs.forEach(t => {
    const btn = document.getElementById(`btn-${t}`);
    const content = document.getElementById(`content-${t}`);
    if (t === tabName) { btn.classList.add("active"); content.classList.remove("hidden"); }
    else { btn.classList.remove("active"); content.classList.add("hidden"); }
  });
}

// ---- Export functions ----
function requestZipExport() {
  const email = document.getElementById("exportEmail").value;
  const startDate = document.getElementById("exportStart").value;
  const endDate = document.getElementById("exportEnd").value;
  const alertBox = document.getElementById("exportAlert");
  if (!startDate || !endDate) { showExportAlert(alertBox, "Please select both dates.", true); return; }
  if (new Date(startDate) > new Date(endDate)) { showExportAlert(alertBox, "Start cannot be after End.", true); return; }
  alertBox.className = "alert";
  document.getElementById("btnDownloadZip").style.opacity = "0.6";
  document.getElementById("spinnerZip").classList.remove("hidden");
  document.getElementById("textZip").innerText = "Compiling...";
  google.script.run
    .withFailureHandler(function (err) { resetZipBtn(); showExportAlert(alertBox, err.message, true); })
    .withSuccessHandler(function (res) {
      resetZipBtn();
      if (res.status === "success") { triggerDownload(res.base64, res.mimeType, res.filename); showExportAlert(alertBox, "ZIP downloaded.", false); }
      else { showExportAlert(alertBox, res.message, true); }
    })
    .generateTimesheetsZip(email, startDate, endDate);
}

function resetZipBtn() {
  document.getElementById("btnDownloadZip").style.opacity = "1";
  document.getElementById("spinnerZip").classList.add("hidden");
  document.getElementById("textZip").innerText = "📦 Export ZIP";
}

function handleExport(format) {
  const email = document.getElementById("exportEmail").value;
  const startDate = document.getElementById("exportStart").value;
  const endDate = document.getElementById("exportEnd").value;
  const alertBox = document.getElementById("exportAlert");
  if (email === "ALL") { showExportAlert(alertBox, "Select a specific employee for PDF/CSV.", true); return; }
  if (!startDate || !endDate) { showExportAlert(alertBox, "Please select both dates.", true); return; }
  if (new Date(startDate) > new Date(endDate)) { showExportAlert(alertBox, "Start cannot be after End.", true); return; }
  alertBox.className = "alert";
  showLoader(true, `Generating ${format.toUpperCase()}...`);
  google.script.run
    .withFailureHandler(function (err) { showLoader(false); showExportAlert(alertBox, err.message, true); })
    .withSuccessHandler(function (res) {
      showLoader(false);
      if (res.status === "success") { triggerDownload(res.base64, res.mimeType, res.filename); showExportAlert(alertBox, `${format.toUpperCase()} exported!`, false); }
      else { showExportAlert(alertBox, res.message, true); }
    })
    .exportTimesheetData(email, startDate, endDate, format);
}

function showExportAlert(el, text, isError) {
  el.innerText = text;
  el.className = `alert visible ${isError ? 'alert-error' : 'alert-success'}`;
}

function triggerDownload(base64Data, contentType, filename) {
  const byteChars = atob(base64Data);
  const byteArrays = [];
  for (let i = 0; i < byteChars.length; i += 512) {
    const slice = byteChars.slice(i, i + 512);
    const nums = new Array(slice.length);
    for (let j = 0; j < slice.length; j++) nums[j] = slice.charCodeAt(j);
    byteArrays.push(new Uint8Array(nums));
  }
  const blob = new Blob(byteArrays, { type: contentType });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  setTimeout(() => { document.body.removeChild(link); URL.revokeObjectURL(link.href); }, 100);
}

// ---- Logs ----
function applyAdminLogFilters() {
  const selectedDate = document.getElementById("filterDate").value.trim();
  const searchQuery = document.getElementById("filterSearch").value.trim().toLowerCase();
  let logins = 0, logouts = 0;
  const filtered = CACHED_LOGS_ARRAY.filter(log => {
    const matchDate = !selectedDate || log.date === selectedDate;
    const fullName = `${log.firstName} ${log.lastName}`.toLowerCase();
    const matchSearch = !searchQuery || fullName.includes(searchQuery) || log.email.toLowerCase().includes(searchQuery);
    if (matchDate && matchSearch) {
      if (log.action === "Clock In") logins++;
      if (log.action === "Clock Out") logouts++;
      return true;
    }
    return false;
  });
  document.getElementById("statLogins").innerText = logins;
  document.getElementById("statLogouts").innerText = logouts;
  document.getElementById("statDateDisplay").innerText = selectedDate || "All Time";
  const body = document.getElementById("logsTableBody");
  body.innerHTML = filtered.length === 0 ? '<tr><td colspan="5" style="padding:1rem;text-align:center;color:var(--slate-500);">No matching logs.</td></tr>' : "";
  filtered.forEach(log => {
    const row = document.createElement("tr");
    const coords = (log.lat && log.lon) ? `<span class="font-mono badge badge-magenta mt-1" style="font-size:0.5625rem;">GPS: ${log.lat}, ${log.lon}</span>` : `<span class="font-mono" style="font-size:0.5625rem;color:var(--slate-400);">GPS: N/A</span>`;
    const editedTag = log.editedBy ? `<span class="badge badge-purple mt-1" style="font-size:0.5625rem;">Edited: ${log.editedBy}</span>` : "";
    row.innerHTML = `
      <td><span style="font-weight:700;color:var(--slate-900);display:block;">${log.firstName} ${log.lastName}</span><span style="font-size:0.625rem;color:var(--slate-500);display:block;">${log.email}</span></td>
      <td><span style="font-weight:500;color:var(--slate-800);display:block;">${log.date}</span><span style="font-size:0.625rem;color:var(--slate-500);display:block;">${log.time}</span></td>
      <td><span class="badge ${log.action === 'Clock In' ? 'badge-magenta' : 'badge-orange'}">${log.action}</span></td>
      <td style="font-size:0.75rem;line-height:1.5;word-break:break-word;"><div style="font-weight:500;color:var(--slate-800);">${log.location}</div>${coords}</td>
      <td>${editedTag}</td>`;
    body.appendChild(row);
  });
}
function resetAdminLogFilters() { document.getElementById("filterDate").value = new Date().toISOString().split('T')[0]; document.getElementById("filterSearch").value = ""; applyAdminLogFilters(); }

// ---- Daily Summaries ----
function applyAdminSummaryFilters() {
  const selectedDate = document.getElementById("filterSummaryDate").value.trim();
  const searchQuery = document.getElementById("filterSummarySearch").value.trim().toLowerCase();
  const filtered = CACHED_SUMMARIES_ARRAY.filter(s => {
    const matchDate = !selectedDate || s.date === selectedDate;
    const fullName = `${s.firstName} ${s.lastName}`.toLowerCase();
    const matchSearch = !searchQuery || fullName.includes(searchQuery) || s.email.toLowerCase().includes(searchQuery);
    return matchDate && matchSearch;
  });
  renderSummariesTable(filtered);
}
function resetAdminSummaryFilters() { document.getElementById("filterSummaryDate").value = new Date().toISOString().split('T')[0]; document.getElementById("filterSummarySearch").value = ""; applyAdminSummaryFilters(); }

function renderSummariesTable(data) {
  const body = document.getElementById("summariesTableBody");
  body.innerHTML = data.length === 0 ? '<tr><td colspan="8" style="padding:1rem;text-align:center;color:var(--slate-500);">No matching summaries.</td></tr>' : "";
  data.forEach(sum => {
    const row = document.createElement("tr");
    const statusBadge = sum.status === "Completed" ? `<span class="badge badge-success">Completed</span>` : `<span class="badge badge-warning">In Progress</span>`;
    const isOT = sum.overtime && sum.overtime !== "0 mins" && sum.overtime !== "N/A";
    const otBadge = isOT ? `<span class="badge badge-purple">${sum.overtime}</span>` : `<span style="font-size:0.75rem;color:var(--slate-500);">${sum.overtime}</span>`;
    row.innerHTML = `
      <td><span style="font-weight:700;color:var(--slate-900);display:block;">${sum.firstName} ${sum.lastName}</span><span style="font-size:0.625rem;color:var(--slate-500);display:block;">${sum.email}</span></td>
      <td style="font-weight:500;color:var(--slate-700);">${sum.date}</td>
      <td><span style="font-size:0.75rem;color:var(--slate-800);display:block;">In: ${sum.clockIn}</span><span style="font-size:0.75rem;color:var(--slate-500);display:block;">Out: ${sum.clockOut}</span></td>
      <td><span class="font-mono" style="font-size:0.75rem;color:var(--vuma-magenta);font-weight:600;display:block;">In: ${formatTimeNoSec(sum.floorIn)}</span><span class="font-mono" style="font-size:0.75rem;color:#EA580C;font-weight:600;display:block;">Out: ${sum.floorOut}</span></td>
      <td style="font-weight:600;color:var(--slate-900);">${sum.totalHours}</td>
      <td>${otBadge}</td>
      <td>${statusBadge}</td>
      <td class="text-right" style="white-space:nowrap;">
        <button onclick="openSummaryEditModal('${sum.summaryId}','${sum.firstName} ${sum.lastName}','${sum.date}','${sum.clockIn}','${sum.clockOut}')" class="btn btn-ghost btn-sm">Edit</button>
        <button onclick="deleteDailySummary('${sum.summaryId}','${sum.date}','${sum.email}')" class="btn btn-danger btn-sm" style="margin-left:0.25rem;">Delete</button>
      </td>`;
    body.appendChild(row);
  });
}

function formatTimeNoSec(t) {
  if (!t || t === "--" || t === "N/A") return "--";
  const p = String(t).trim().split(":");
  return p.length >= 2 ? `${p[0].padStart(2, '0')}:${p[1].padStart(2, '0')}` : t;
}

function formatTimeForInput(t) {
  if (!t || t === "--" || t === "N/A") return "";
  const p = t.split(":");
  return p.length >= 2 ? `${p[0].padStart(2, '0')}:${p[1].padStart(2, '0')}:${p.length > 2 ? p[2].padStart(2, '0') : "00"}` : "";
}

function openSummaryEditModal(sumId, name, date, clockIn, clockOut) {
  document.getElementById("editSumId").value = sumId;
  document.getElementById("editSumName").innerText = name;
  document.getElementById("editSumDate").innerText = "Date: " + date;
  document.getElementById("editSumIn").value = formatTimeForInput(clockIn);
  document.getElementById("editSumOut").value = formatTimeForInput(clockOut);
  document.getElementById("editSummaryModal").classList.add("active");
}
function closeSummaryEditModal() { document.getElementById("editSummaryModal").classList.remove("active"); }

function handleSummaryUpdateSave(e) {
  e.preventDefault();
  showLoader(true, "Updating & syncing...");
  google.script.run
    .withSuccessHandler(function (res) {
      showLoader(false); closeSummaryEditModal();
      if (res.status === "success") loadAdminDashboard();
      else alert("Error: " + res.message);
    })
    .adminUpdateDailySummary(document.getElementById("editSumId").value, document.getElementById("editSumIn").value, document.getElementById("editSumOut").value, ACTIVE_SESSION_USER.email);
}

function deleteDailySummary(summaryId, dateStr, email) {
  if (!confirm("Delete this shift permanently? Activity logs for this day will also be removed.")) return;
  showLoader(true, "Deleting shift...");
  google.script.run
    .withFailureHandler(function (err) { showLoader(false); alert("Error: " + err.message); })
    .withSuccessHandler(function (res) {
      showLoader(false);
      if (res.status === "success") loadAdminDashboard();
      else alert("Error: " + res.message);
    })
    .adminDeleteDailySummary(summaryId, dateStr, email, ACTIVE_SESSION_USER.email);
}

// ---- Weekly ----
function applyAdminWeeklyFilters() {
  const wKey = document.getElementById("filterWeeklyKey").value.trim().toUpperCase();
  const search = document.getElementById("filterWeeklySearch").value.trim().toLowerCase();
  const filtered = CACHED_WEEKLY_ARRAY.filter(w => {
    const matchWeek = !wKey || w.weekKey.toUpperCase().includes(wKey);
    const fullName = `${w.firstName} ${w.lastName}`.toLowerCase();
    const matchSearch = !search || fullName.includes(search) || w.email.toLowerCase().includes(search);
    return matchWeek && matchSearch;
  });
  renderWeeklyTable(filtered);
}
function resetAdminWeeklyFilters() { document.getElementById("filterWeeklyKey").value = ""; document.getElementById("filterWeeklySearch").value = ""; applyAdminWeeklyFilters(); }

function renderWeeklyTable(data) {
  const body = document.getElementById("weeklyTableBody");
  body.innerHTML = data.length === 0 ? '<tr><td colspan="6" style="padding:1rem;text-align:center;color:var(--slate-500);">No matching weekly summaries.</td></tr>' : "";
  data.forEach(w => {
    const row = document.createElement("tr");
    let compBadge = `<span class="badge badge-warning">Under Target</span>`;
    if (w.complianceStatus.includes("MAX CAP")) compBadge = `<span class="badge badge-danger animate-pulse">🔴 MAX CAP (55+h)</span>`;
    else if (w.complianceStatus.includes("Target Met")) compBadge = `<span class="badge badge-success">Target Met</span>`;
    const isOT = w.totalOvertime && w.totalOvertime !== "0 mins" && w.totalOvertime !== "N/A";
    const otBadge = isOT ? `<span class="badge badge-purple">${w.totalOvertime}</span>` : `<span style="font-size:0.75rem;color:var(--slate-500);">${w.totalOvertime}</span>`;
    row.innerHTML = `
      <td><span style="font-weight:700;color:var(--slate-900);display:block;">${w.firstName} ${w.lastName}</span><span style="font-size:0.625rem;color:var(--slate-500);display:block;">${w.email}</span></td>
      <td class="font-mono" style="font-weight:600;color:var(--vuma-magenta);">${w.weekKey}</td>
      <td style="font-weight:700;color:var(--slate-900);">${w.totalHours}</td>
      <td>${otBadge}</td>
      <td style="font-weight:600;color:var(--slate-700);">${w.daysWorked} days</td>
      <td>${compBadge}</td>`;
    body.appendChild(row);
  });
}

// ---- Monthly ----
function applyAdminMonthlyFilters() {
  const month = document.getElementById("filterMonthlyMonth").value.trim();
  const search = document.getElementById("filterMonthlySearch").value.trim().toLowerCase();
  const filtered = CACHED_MONTHLY_ARRAY.filter(m => {
    const matchMonth = !month || m.month === month;
    const fullName = `${m.firstName} ${m.lastName}`.toLowerCase();
    const matchSearch = !search || fullName.includes(search) || m.email.toLowerCase().includes(search);
    return matchMonth && matchSearch;
  });
  renderMonthlyTable(filtered);
}
function resetAdminMonthlyFilters() { document.getElementById("filterMonthlyMonth").value = ""; document.getElementById("filterMonthlySearch").value = ""; applyAdminMonthlyFilters(); }

function renderMonthlyTable(data) {
  const body = document.getElementById("monthlyTableBody");
  body.innerHTML = data.length === 0 ? '<tr><td colspan="6" style="padding:1rem;text-align:center;color:var(--slate-500);">No matching monthly summaries.</td></tr>' : "";
  data.forEach(m => {
    const row = document.createElement("tr");
    const isOT = m.totalOvertime && m.totalOvertime !== "0 mins" && m.totalOvertime !== "N/A";
    const otBadge = isOT ? `<span class="badge badge-purple">${m.totalOvertime}</span>` : `<span style="font-size:0.75rem;color:var(--slate-500);">${m.totalOvertime}</span>`;
    row.innerHTML = `
      <td><span style="font-weight:700;color:var(--slate-900);display:block;">${m.firstName} ${m.lastName}</span><span style="font-size:0.625rem;color:var(--slate-500);display:block;">${m.email}</span></td>
      <td style="font-weight:600;color:var(--vuma-magenta);">${m.month}</td>
      <td style="font-weight:700;color:var(--slate-900);">${m.totalHours}</td>
      <td>${otBadge}</td>
      <td style="font-weight:600;color:#059669;">${m.daysWorked} days</td>
      <td class="font-mono" style="font-size:0.75rem;color:var(--slate-400);">${m.lastUpdated}</td>`;
    body.appendChild(row);
  });
}

// ---- Leave ----
function renderLeaveRequestsTable(requests) {
  const body = document.getElementById("leaveTableBody");
  body.innerHTML = requests.length === 0 ? '<tr><td colspan="5" style="padding:1rem;text-align:center;color:var(--slate-500);">No leave applications on record.</td></tr>' : "";
  requests.forEach(req => {
    const row = document.createElement("tr");
    let statusClass = "badge-warning";
    if (req.status === "Approved") statusClass = "badge-success";
    if (req.status === "Rejected") statusClass = "badge-danger";
    const actions = req.status === "Pending" ? `
      <div style="display:flex;justify-content:flex-end;gap:0.5rem;">
        <button onclick="handleLeaveAction('${req.reqId}','Approved')" class="btn btn-success btn-sm">Approve</button>
        <button onclick="handleLeaveAction('${req.reqId}','Rejected')" class="btn btn-danger btn-sm">Reject</button>
      </div>
    ` : `<span style="font-size:0.75rem;color:var(--slate-500);font-style:italic;display:block;text-align:right;">Resolved</span>`;
    row.innerHTML = `
      <td><span style="font-weight:700;color:var(--slate-900);display:block;">${req.firstName} ${req.lastName}</span><span style="font-size:0.625rem;color:var(--slate-500);display:block;">${req.email}</span></td>
      <td style="font-weight:600;color:var(--slate-800);">${req.type} Leave</td>
      <td style="font-size:0.75rem;color:var(--slate-700);"><span style="display:block;font-weight:500;">${req.startDate} to</span><span style="display:block;color:var(--slate-500);">${req.endDate}</span></td>
      <td><span class="badge ${statusClass}">${req.status}</span></td>
      <td class="text-right">${actions}</td>`;
    body.appendChild(row);
  });
}

function handleLeaveAction(reqId, newStatus) {
  showLoader(true, "Updating leave status...");
  google.script.run
    .withSuccessHandler(function (res) {
      showLoader(false);
      if (res.status === "success") loadAdminDashboard();
      else alert("Error: " + res.message);
    })
    .adminUpdateLeaveStatus(reqId, newStatus);
}

// ---- Staff Profiles ----
function renderStaffProfilesTable(employees) {
  const body = document.getElementById("usersTableBody");
  body.innerHTML = employees.length === 0 ? '<tr><td colspan="4" style="padding:1rem;text-align:center;color:var(--slate-500);">No employee profiles found.</td></tr>' : "";
  employees.forEach(emp => {
    const row = document.createElement("tr");
    const workBadge = emp.workStatus === "On Leave" ? `<span class="badge badge-warning">On Leave</span>` : `<span class="badge badge-success">Active</span>`;
    const standBadge = emp.accountStanding === "Verified" ? `<span class="badge badge-magenta">Verified</span>` : `<span class="badge badge-danger">Blocked</span>`;
    row.innerHTML = `
      <td><span style="font-weight:700;color:var(--slate-900);display:block;">${emp.firstName} ${emp.lastName}</span><span style="font-size:0.625rem;color:var(--slate-500);display:block;">${emp.email}</span></td>
      <td>${workBadge}</td>
      <td>${standBadge}</td>
      <td class="text-right"><button onclick="openEditModal('${emp.email}','${emp.firstName}','${emp.lastName}','${emp.accountStanding}')" class="btn btn-ghost btn-sm">Edit</button></td>`;
    body.appendChild(row);
  });
}

function openEditModal(email, first, last, status) {
  document.getElementById("editEmail").value = email;
  document.getElementById("editFirst").value = first;
  document.getElementById("editLast").value = last;
  document.getElementById("editStatus").value = status;
  document.getElementById("editPassword").value = "";
  document.getElementById("editModal").classList.add("active");
}
function closeEditModal() { document.getElementById("editModal").classList.remove("active"); }

function handleAdminUpdateSave(e) {
  e.preventDefault();
  showLoader(true, "Updating profile...");
  google.script.run
    .withSuccessHandler(function (res) {
      showLoader(false); closeEditModal();
      if (res.status === "success") loadAdminDashboard();
      else alert("Error: " + res.message);
    })
    .adminUpdateUser(document.getElementById("editEmail").value, document.getElementById("editFirst").value, document.getElementById("editLast").value, document.getElementById("editStatus").value, document.getElementById("editPassword").value);
}
