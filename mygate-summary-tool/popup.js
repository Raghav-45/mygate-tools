/**
 * MyGate Summary Tool - Frontend Controller
 * Reactive UI that syncs with background Service Worker scan state.
 * Allows closing popup during scans and indicates automatic Excel downloads.
 */

document.addEventListener("DOMContentLoaded", async () => {
  // DOM Elements
  const reportMonthInput = document.getElementById("reportMonth");
  const settingsCollapse = document.getElementById("settingsCollapse");
  const settingsToggleBtn = document.getElementById("settingsToggleBtn");
  const requestDelayInput = document.getElementById("requestDelayInput");
  const delayValText = document.getElementById("delayValText");

  const generateBtn = document.getElementById("generateBtn");
  const abortBtn = document.getElementById("abortBtn");

  const alertBanner = document.getElementById("alertBanner");
  const progressContainer = document.getElementById("progressContainer");
  const progressStepText = document.getElementById("progressStepText");
  const progressPctText = document.getElementById("progressPctText");
  const progressFill = document.getElementById("progressFill");

  const resultsSection = document.getElementById("resultsSection");
  const kpiReceived = document.getElementById("kpiReceived");
  const kpiClosed = document.getElementById("kpiClosed");
  const kpiPending = document.getElementById("kpiPending");
  const tableBody = document.getElementById("tableBody");
  const autoDlBanner = document.getElementById("autoDlBanner");

  const aboutIconBtn = document.getElementById("aboutIconBtn");
  const aboutModal = document.getElementById("aboutModal");
  const closeModalBtn = document.getElementById("closeModalBtn");

  // State
  let latestRows = [];
  let currentYear = 2026;
  let currentMonth = 6;

  // Set default month to current
  const now = new Date();
  currentYear = now.getFullYear();
  currentMonth = now.getMonth() + 1;
  reportMonthInput.value = `${currentYear}-${String(currentMonth).padStart(2, '0')}`;

  // Load stored settings & active scan state
  const stored = await chrome.storage.local.get(["requestDelay", "summaryScanState"]);
  if (stored.requestDelay) {
    requestDelayInput.value = stored.requestDelay;
    delayValText.textContent = `${stored.requestDelay}s`;
  }

  // Restore live scan state if active or completed recently
  const scan = stored.summaryScanState;
  if (scan) {
    if (scan.rows) {
      latestRows = scan.rows;
      renderTable();
      updateKpis();
    }
    if (scan.isScanning) {
      generateBtn.classList.add("hidden");
      abortBtn.classList.remove("hidden");
      progressContainer.classList.remove("hidden");
      resultsSection.classList.remove("hidden");
      progressStepText.textContent = scan.stepText || "Scanning...";
      progressPctText.textContent = `${scan.pct || 0}%`;
      progressFill.style.width = `${scan.pct || 0}%`;
    } else if (scan.isDone) {
      resultsSection.classList.remove("hidden");
      if (autoDlBanner) autoDlBanner.classList.remove("hidden");
    }
  }

  // Toggle handlers
  settingsToggleBtn.addEventListener("click", () => {
    settingsCollapse.classList.toggle("hidden");
  });

  function toggleModal() {
    aboutModal.classList.toggle("hidden");
  }
  if (aboutIconBtn) aboutIconBtn.addEventListener("click", toggleModal);
  if (closeModalBtn) closeModalBtn.addEventListener("click", toggleModal);
  if (aboutModal) {
    aboutModal.addEventListener("click", (e) => {
      if (e.target === aboutModal) aboutModal.classList.add("hidden");
    });
  }

  requestDelayInput.addEventListener("input", (e) => {
    const val = Number(e.target.value).toFixed(1);
    delayValText.textContent = `${val}s`;
    chrome.storage.local.set({ requestDelay: Number(val) });
  });

  function showAlert(msg, isErr = true) {
    alertBanner.textContent = msg;
    alertBanner.className = `alert-box ${isErr ? 'error' : 'success'}`;
    alertBanner.classList.remove("hidden");
  }
  function hideAlert() { alertBanner.classList.add("hidden"); }

  // Generate Report
  generateBtn.addEventListener("click", () => {
    hideAlert();
    if (autoDlBanner) autoDlBanner.classList.add("hidden");
    if (!reportMonthInput.value) {
      showAlert("Please select a report month.");
      return;
    }

    const [yyyy, mm] = reportMonthInput.value.split('-').map(Number);
    currentYear = yyyy;
    currentMonth = mm;
    latestRows = [];
    tableBody.innerHTML = "";
    kpiReceived.textContent = "0";
    kpiClosed.textContent = "0";
    kpiPending.textContent = "0";

    generateBtn.classList.add("hidden");
    abortBtn.classList.remove("hidden");
    progressContainer.classList.remove("hidden");
    resultsSection.classList.remove("hidden");
    progressStepText.textContent = "Connecting MyGate API...";
    progressPctText.textContent = "0%";
    progressFill.style.width = "0%";

    const delay = Number(requestDelayInput.value);
    chrome.runtime.sendMessage({
      type: "START_SUMMARY_SCAN",
      year: yyyy,
      month: mm,
      requestDelay: delay
    });
  });

  abortBtn.addEventListener("click", () => {
    chrome.runtime.sendMessage({ type: "ABORT_SUMMARY_SCAN" });
    abortBtn.disabled = true;
    abortBtn.textContent = "Stopping...";
  });

  function parseDateToTime(dStr) {
    const [dd, mmm, yyyy] = dStr.split('-');
    const mIdx = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"].indexOf(mmm);
    return new Date(yyyy, mIdx, Number(dd)).getTime();
  }

  function sortRowsChronologically(rows) {
    return [...rows].sort((a, b) => parseDateToTime(a.date) - parseDateToTime(b.date));
  }

  // Message listener from Background
  chrome.runtime.onMessage.addListener((msg) => {
    if (msg.type === "SUMMARY_PROGRESS") {
      if (msg.stepText) progressStepText.textContent = msg.stepText;
      if (msg.pct !== undefined) {
        progressPctText.textContent = `${msg.pct}%`;
        progressFill.style.width = `${msg.pct}%`;
      }
      if (msg.row) {
        latestRows.push(msg.row);
        latestRows = sortRowsChronologically(latestRows);
        renderTable();
        updateKpis();
      }
    }

    if (msg.type === "SUMMARY_DONE" || msg.type === "SUMMARY_ABORTED") {
      generateBtn.classList.remove("hidden");
      abortBtn.classList.add("hidden");
      abortBtn.disabled = false;
      abortBtn.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect></svg> Stop Live Scan`;
      progressContainer.classList.add("hidden");
      
      if (msg.type === "SUMMARY_ABORTED") {
        showAlert("Live scan stopped by user.", true);
      } else {
        if (autoDlBanner) autoDlBanner.classList.remove("hidden");
        showAlert("Report downloaded automatically to your Downloads folder!", false);
      }
    }

    if (msg.type === "SUMMARY_ERROR") {
      generateBtn.classList.remove("hidden");
      abortBtn.classList.add("hidden");
      progressContainer.classList.add("hidden");
      showAlert(`Scan Failed: ${msg.error}`, true);
    }
  });

  function renderTable() {
    tableBody.innerHTML = "";
    latestRows.forEach(r => {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${r.date}</td>
        <td class="num">${r.prevOpen.toLocaleString()}</td>
        <td class="num">${r.received.toLocaleString()}</td>
        <td class="num">${r.closed.toLocaleString()}</td>
      `;
      tableBody.appendChild(tr);
    });
    // Auto scroll bottom
    const scroll = document.querySelector(".table-scroll");
    if (scroll) scroll.scrollTop = scroll.scrollHeight;
  }

  function updateKpis() {
    const totRec = latestRows.reduce((acc, r) => acc + r.received, 0);
    const totClosed = latestRows.reduce((acc, r) => acc + r.closed, 0);
    const lastPend = latestRows.length ? latestRows[latestRows.length - 1].pending : 0;

    kpiReceived.textContent = totRec.toLocaleString();
    kpiClosed.textContent = totClosed.toLocaleString();
    kpiPending.textContent = lastPend.toLocaleString();
  }
});
