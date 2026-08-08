/**
 * MyGate Report Tool - Frontend Controller Logic
 * Replicates the Python openpyxl Excel generation byte-for-byte using ExcelJS.
 */

document.addEventListener("DOMContentLoaded", async () => {
  // DOM Elements
  const fromDateInput = document.getElementById("fromDate");
  const toDateInput = document.getElementById("toDate");
  const categoriesContainer = document.getElementById("categoriesContainer");
  const toggleAllBtn = document.getElementById("toggleAllBtn");

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
  const kpiTotal = document.getElementById("kpiTotal");
  const kpiResolved = document.getElementById("kpiResolved");
  const kpiOpen = document.getElementById("kpiOpen");
  const tableBody = document.getElementById("tableBody");
  
  const aboutIconBtn = document.getElementById("aboutIconBtn");
  const aboutModal = document.getElementById("aboutModal");
  const closeModalBtn = document.getElementById("closeModalBtn");

  // State
  let allCategories = [];
  let latestResults = [];
  let latestSummary = { total: 0, resolved: 0, open: 0 };
  let isScanning = false;

  // Set default dates
  fromDateInput.value = "2024-01-01";
  toDateInput.value = new Date().toISOString().split('T')[0];

  // Load saved settings & active scan state
  const storedConfig = await chrome.storage.local.get(["requestDelay", "selectedCatIds", "ticketsScanState"]);
  if (storedConfig.requestDelay) {
    requestDelayInput.value = storedConfig.requestDelay;
    delayValText.textContent = `${storedConfig.requestDelay}s`;
  }

  const scan = storedConfig.ticketsScanState;
  if (scan) {
    if (scan.results && scan.results.length) {
      latestResults = scan.results;
      latestSummary = scan.summary || { total: 0, resolved: 0, open: 0 };
      renderResults(latestResults, latestSummary);
      resultsSection.classList.remove("hidden");
    }
    if (scan.isScanning) {
      isScanning = true;
      generateBtn.classList.add("hidden");
      abortBtn.classList.remove("hidden");
      progressContainer.classList.remove("hidden");
      progressStepText.textContent = scan.statusText || "Scanning...";
      progressPctText.textContent = `${scan.pct || 0}%`;
      progressFill.style.width = `${scan.pct || 0}%`;
    } else if (scan.isDone) {
      const autoDlBanner = document.getElementById("autoDlBanner");
      if (autoDlBanner) autoDlBanner.classList.remove("hidden");
    }
  }

  // Fetch category list from Background script
  chrome.runtime.sendMessage({ type: "GET_CATEGORIES_LIST" }, (res) => {
    if (res && res.categories) {
      allCategories = res.categories;
      renderCategoryPills(storedConfig.selectedCatIds);
    }
  });

  function renderCategoryPills(selectedIds = null) {
    categoriesContainer.innerHTML = "";
    allCategories.forEach(([name, id]) => {
      const pill = document.createElement("div");
      pill.className = "category-pill selected";
      if (selectedIds && !selectedIds.includes(id)) {
        pill.classList.remove("selected");
      }
      pill.textContent = name;
      pill.dataset.id = id;
      pill.addEventListener("click", () => {
        pill.classList.toggle("selected");
        saveCategorySelection();
      });
      categoriesContainer.appendChild(pill);
    });
  }

  function getSelectedCategoryIds() {
    const pills = categoriesContainer.querySelectorAll(".category-pill.selected");
    return Array.from(pills).map(p => Number(p.dataset.id));
  }

  function saveCategorySelection() {
    chrome.storage.local.set({ selectedCatIds: getSelectedCategoryIds() });
  }

  // UI Event Handlers
  settingsToggleBtn.addEventListener("click", () => {
    settingsCollapse.classList.toggle("hidden");
  });

  function toggleAboutModal() {
    aboutModal.classList.toggle("hidden");
  }

  if (aboutIconBtn) aboutIconBtn.addEventListener("click", toggleAboutModal);
  if (closeModalBtn) closeModalBtn.addEventListener("click", toggleAboutModal);
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

  toggleAllBtn.addEventListener("click", () => {
    const pills = categoriesContainer.querySelectorAll(".category-pill");
    const allSel = Array.from(pills).every(p => p.classList.contains("selected"));
    pills.forEach(p => {
      if (allSel) p.classList.remove("selected");
      else p.classList.add("selected");
    });
    saveCategorySelection();
  });

  function showAlert(msg, isError = true) {
    alertBanner.textContent = msg;
    alertBanner.className = `alert-box ${isError ? 'error' : 'success'}`;
    alertBanner.classList.remove("hidden");
  }

  // Scan Execution
  generateBtn.addEventListener("click", () => {
    const autoDlBanner = document.getElementById("autoDlBanner");
    if (autoDlBanner) autoDlBanner.classList.add("hidden");

    const fromDate = fromDateInput.value;
    const toDate = toDateInput.value;

    if (!fromDate || !toDate) {
      showAlert("Please enter valid Report From and Report Till dates.");
      return;
    }
    if (new Date(fromDate) > new Date(toDate)) {
      showAlert("Report From date cannot be after Till date.");
      return;
    }

    const selectedCategoryIds = getSelectedCategoryIds();
    if (selectedCategoryIds.length === 0) {
      showAlert("Please select at least one Ticket Category.");
      return;
    }

    chrome.storage.local.set({
      requestDelay: Number(requestDelayInput.value)
    });

    isScanning = true;
    latestResults = [];
    latestSummary = { total: 0, resolved: 0, open: 0 };
    alertBanner.classList.add("hidden");
    resultsSection.classList.add("hidden");

    generateBtn.classList.add("hidden");
    abortBtn.classList.remove("hidden");
    progressContainer.classList.remove("hidden");

    progressStepText.textContent = "Connecting to MyGate API...";
    progressPctText.textContent = "0%";
    progressFill.style.width = "0%";

    chrome.runtime.sendMessage({
      type: "START_REPORT_SCAN",
      params: {
        fromDate,
        toDate,
        requestDelayMs: Number(requestDelayInput.value) * 1000,
        selectedCategoryIds
      }
    });
  });

  abortBtn.addEventListener("click", () => {
    chrome.runtime.sendMessage({ type: "ABORT_REPORT_SCAN" });
    abortBtn.disabled = true;
    abortBtn.textContent = "Stopping...";
  });

  // Background Worker Messages
  chrome.runtime.onMessage.addListener((msg) => {
    if (msg.type === "SCAN_PROGRESS_UPDATE") {
      const { stepIndex, totalSteps, statusText } = msg;
      const pct = Math.round((stepIndex / totalSteps) * 100);
      progressStepText.textContent = statusText;
      progressPctText.textContent = `${pct}%`;
      progressFill.style.width = `${pct}%`;
    }
    else if (msg.type === "CATEGORY_COMPLETED") {
      latestResults.push(msg.data);
      latestSummary = msg.summary;
      renderResults(latestResults, latestSummary);
      resultsSection.classList.remove("hidden");
    }
    else if (msg.type === "SCAN_FINISHED" || msg.type === "SCAN_ABORTED") {
      generateBtn.classList.remove("hidden");
      abortBtn.classList.add("hidden");
      abortBtn.disabled = false;
      abortBtn.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect></svg> Stop Live Scan`;
      progressContainer.classList.add("hidden");

      if (msg.type === "SCAN_FINISHED") {
        const autoDlBanner = document.getElementById("autoDlBanner");
        if (autoDlBanner) autoDlBanner.classList.remove("hidden");
        showAlert("Report downloaded automatically to your Downloads folder!", false);
      } else {
        showAlert("Scan stopped by user.", true);
      }
    }
    else if (msg.type === "SCAN_ERROR") {
      generateBtn.classList.remove("hidden");
      abortBtn.classList.add("hidden");
      progressContainer.classList.add("hidden");
      showAlert(msg.errorMessage || `Error querying category: ${msg.category}`);
    }
  });

  function renderResults(results, summary) {
    kpiTotal.textContent = summary.total.toLocaleString();
    kpiResolved.textContent = summary.resolved.toLocaleString();
    kpiOpen.textContent = summary.open.toLocaleString();

    let html = "";
    results.forEach(r => {
      html += `
        <tr>
          <td>${r.name}</td>
          <td class="num">${r.total.toLocaleString()}</td>
          <td class="num">${r.resolved.toLocaleString()}</td>
          <td class="num">${r.open.toLocaleString()}</td>
        </tr>`;
    });

    html += `
      <tr style="background: #FEF08A; font-weight: 700;">
        <td>Total</td>
        <td class="num">${summary.total.toLocaleString()}</td>
        <td class="num">${summary.resolved.toLocaleString()}</td>
        <td class="num">${summary.open.toLocaleString()}</td>
      </tr>`;

    tableBody.innerHTML = html;
  }
});
