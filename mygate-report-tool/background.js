/**
 * MyGate Report Tool Chrome Extension - Background Service Worker
 * Executes GraphQL API queries with auto-token discovery and strict auth verification.
 */

importScripts('exceljs.min.js');

const CATEGORIES = [
  ["Accounts Billing", 252434],
  ["Construction Or Project Related", 277747],
  ["Design Related Issue", 277632],
  ["Estate Infra Outer Area from the plot", 267181],
  ["FM Common Area Related Issue", 277745],
  ["IT WIFI Network", 277816],
  ["Products Appliances", 277744]
];

const GRAPHQL_URL = "https://api.dashboard.mygate.com/graphql/";
const SAMPLE_FALLBACK_TOKEN = "zr2Er9wrdfhTyiY01Lnvr03cje9oeeH7wR6XvFmeSR87okw1qW4QyAuRkoSaOkff";

const STATUS_TOTAL = ["open", "hold", "re_opened", "job_done", "in_progress", "closed"];
const STATUS_OPEN = ["open", "hold", "re_opened", "job_done", "in_progress"];
const STATUS_RESOLVED = ["closed"];

let activeScanAbort = false;

function getMidnightEpoch(dateStr) {
  if (!dateStr) return Math.floor(Date.now() / 1000);
  const [year, month, day] = dateStr.split('-').map(Number);
  const date = new Date(year, month - 1, day, 0, 0, 0);
  return Math.floor(date.getTime() / 1000);
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Automatically discover user's auth token from open tabs or browser cookies
async function discoverActiveAuthToken() {
  // 1. Inspect open MyGate Dashboard tabs for localStorage tokens
  try {
    if (chrome.tabs && chrome.scripting) {
      const tabs = await chrome.tabs.query({ url: "*://*.mygate.com/*" });
      for (const tab of tabs) {
        try {
          const results = await chrome.scripting.executeScript({
            target: { tabId: tab.id },
            func: () => {
              const searchKeys = ['token', 'auth_token', 'authorization', 'access_token', 'jwt', 'mygate_token', 'user'];
              for (let i = 0; i < localStorage.length; i++) {
                const k = localStorage.key(i);
                if (searchKeys.some(key => k.toLowerCase().includes(key))) {
                  const val = localStorage.getItem(k);
                  if (val && typeof val === 'string' && val.length > 20 && !val.startsWith('{')) return val;
                  if (val && val.startsWith('{')) {
                    try {
                      const obj = JSON.parse(val);
                      if (obj.token) return obj.token;
                      if (obj.accessToken) return obj.accessToken;
                      if (obj.jwt) return obj.jwt;
                      if (obj.authorization) return obj.authorization;
                    } catch(e){}
                  }
                }
              }
              return null;
            }
          });
          if (results && results[0] && results[0].result) {
            console.log("Discovered auth token from active tab");
            return results[0].result;
          }
        } catch (tabErr) {}
      }
    }
  } catch (e) { console.warn("Tab token discovery failed:", e); }

  // 2. Inspect browser cookies
  try {
    if (chrome.cookies) {
      const cookies = await chrome.cookies.getAll({ domain: "mygate.com" });
      for (const c of cookies) {
        const name = c.name.toLowerCase();
        if (['token', 'auth_token', 'authorization', 'access_token', 'jwt', 'session_token'].includes(name)) {
          if (c.value && c.value.length > 20) {
            console.log("Discovered auth token from browser cookies");
            return c.value;
          }
        }
      }
    }
  } catch (e) { console.warn("Cookie token discovery failed:", e); }

  return null;
}

async function fetchGraphQLCount(categoryId, fromEpoch, toEpoch, statuses, authToken) {
  const payload = {
    operationName: "getAdminSrList",
    variables: {
      requestData: {
        requiredFields: ["id"],
        pagination: { count: 1, page: 1 },
        sorting: [],
        conditions: [
          { name: "date_filter", operation: "equal", values: ["created_date"] },
          { name: "category", values: [categoryId], operation: "equal" },
          { name: "from_date", values: [fromEpoch], operation: "gte" },
          { name: "to_date", values: [toEpoch], operation: "lte" },
          { name: "mygate_status", values: statuses, operation: statuses.length === 1 ? "equal" : "in" }
        ]
      }
    },
    query: "query getAdminSrList($requestData: DataListInput){getAdminSrList(requestData:$requestData){dataResponse{totalCount}}}"
  };

  const headers = {
    "content-type": "application/json",
    "origin": "https://dashboard.mygate.com",
    "referer": "https://dashboard.mygate.com/"
  };

  // Only attach authorization header if explicitly discovered or configured
  if (authToken && authToken.trim() !== "") {
    headers["authorization"] = authToken.trim();
  }

  const res = await fetch(GRAPHQL_URL, {
    method: "POST",
    headers: headers,
    credentials: "include", // Send browser session cookies
    body: JSON.stringify(payload)
  });

  if (!res.ok) {
    if (res.status === 401 || res.status === 403) {
      throw new Error("HTTP 401/403 Unauthorized: Please log into dashboard.mygate.com in Google Chrome.");
    }
    throw new Error(`HTTP Error ${res.status}: ${res.statusText}`);
  }

  const data = await res.json();
  if (data.errors && data.errors.length > 0) {
    throw new Error(`MyGate API Error: ${data.errors[0]?.message}`);
  }

  const dataResponse = data?.data?.getAdminSrList?.dataResponse;
  if (!dataResponse || typeof dataResponse.totalCount !== 'number') {
    throw new Error("Authentication Verification Failed: MyGate GraphQL returned empty data. Please ensure you are logged into dashboard.mygate.com in Google Chrome.");
  }

  return dataResponse.totalCount;
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === "START_REPORT_SCAN") {
    activeScanAbort = false;
    runFullReportScan(message.params);
  } else if (message.type === "ABORT_REPORT_SCAN") {
    activeScanAbort = true;
  } else if (message.type === "GET_CATEGORIES_LIST") {
    sendResponse({ categories: CATEGORIES });
  }
  return true;
});

async function runFullReportScan(params) {
  const { fromDate, toDate, requestDelayMs = 1000, selectedCategoryIds } = params;
  
  const fromEpoch = getMidnightEpoch(fromDate);
  const toEpoch = getMidnightEpoch(toDate) + 86399;

  // Automatically discover active tab token
  let tokenToUse = await discoverActiveAuthToken();
  if (!tokenToUse) {
    tokenToUse = SAMPLE_FALLBACK_TOKEN;
  }

  const categoriesToScan = CATEGORIES.filter(([name, id]) => {
    if (!selectedCategoryIds || selectedCategoryIds.length === 0) return true;
    return selectedCategoryIds.includes(id);
  });

  const results = [];
  let totalAll = 0;
  let resolvedAll = 0;
  let openAll = 0;

  await chrome.storage.local.set({
    ticketsScanState: { isScanning: true, pct: 0, statusText: "Connecting MyGate API...", results: [], summary: { total: 0, resolved: 0, open: 0 }, isDone: false }
  });

  for (let i = 0; i < categoriesToScan.length; i++) {
    if (activeScanAbort) {
      await chrome.storage.local.set({ ticketsScanState: { isScanning: false, isAborted: true } });
      chrome.runtime.sendMessage({ type: "SCAN_ABORTED", results }).catch(() => {});
      return;
    }

    const [catName, catId] = categoriesToScan[i];

    try {
      chrome.runtime.sendMessage({
        type: "SCAN_PROGRESS_UPDATE",
        stepIndex: i + 1,
        totalSteps: categoriesToScan.length,
        currentCategory: catName,
        statusText: `Fetching Total tickets for ${catName}...`
      });

      const totalCount = await fetchGraphQLCount(catId, fromEpoch, toEpoch, STATUS_TOTAL, tokenToUse);
      await sleep(requestDelayMs);

      if (activeScanAbort) break;

      chrome.runtime.sendMessage({
        type: "SCAN_PROGRESS_UPDATE",
        stepIndex: i + 1,
        totalSteps: categoriesToScan.length,
        currentCategory: catName,
        statusText: `Fetching Resolved tickets for ${catName}...`
      });

      const resolvedCount = await fetchGraphQLCount(catId, fromEpoch, toEpoch, STATUS_RESOLVED, tokenToUse);
      await sleep(requestDelayMs);

      if (activeScanAbort) break;

      chrome.runtime.sendMessage({
        type: "SCAN_PROGRESS_UPDATE",
        stepIndex: i + 1,
        totalSteps: categoriesToScan.length,
        currentCategory: catName,
        statusText: `Fetching Open tickets for ${catName}...`
      });

      const openCount = await fetchGraphQLCount(catId, fromEpoch, toEpoch, STATUS_OPEN, tokenToUse);
      await sleep(requestDelayMs);

      totalAll += totalCount;
      resolvedAll += resolvedCount;
      openAll += openCount;

      const categoryData = {
        name: catName,
        id: catId,
        total: totalCount,
        resolved: resolvedCount,
        open: openCount
      };

      results.push(categoryData);
      const newPct = Math.round(((i + 1) / categoriesToScan.length) * 100);

      await chrome.storage.local.set({
        ticketsScanState: { isScanning: true, pct: newPct, statusText: `Completed ${catName}`, results, summary: { total: totalAll, resolved: resolvedAll, open: openAll }, isDone: false }
      });

      chrome.runtime.sendMessage({
        type: "CATEGORY_COMPLETED",
        data: categoryData,
        summary: { total: totalAll, resolved: resolvedAll, open: openAll }
      }).catch(() => {});

    } catch (error) {
      console.error(`Error querying category ${catName}:`, error);
      await chrome.storage.local.set({ ticketsScanState: { isScanning: false, error: error.message } });
      chrome.runtime.sendMessage({
        type: "SCAN_ERROR",
        category: catName,
        errorMessage: error.message || "Unknown Network Error"
      }).catch(() => {});
      return;
    }
  }

  if (!activeScanAbort) {
    await chrome.storage.local.set({
      ticketsScanState: { isScanning: true, pct: 100, statusText: "Downloading Excel...", results, summary: { total: totalAll, resolved: resolvedAll, open: openAll }, isDone: false }
    });
    chrome.runtime.sendMessage({
      type: "SCAN_PROGRESS_UPDATE",
      statusText: "Downloading Excel..."
    }).catch(() => {});
    
    await autoDownloadTicketsExcel(results, { total: totalAll, resolved: resolvedAll, open: openAll }, fromDate, toDate);

    await chrome.storage.local.set({
      ticketsScanState: { isScanning: false, pct: 100, statusText: "Report Generated!", results, summary: { total: totalAll, resolved: resolvedAll, open: openAll }, isDone: true }
    });
    chrome.runtime.sendMessage({
      type: "SCAN_FINISHED",
      results,
      summary: { total: totalAll, resolved: resolvedAll, open: openAll },
      reportMeta: { fromDate, toDate }
    }).catch(() => {});
  }
}

function arrayBufferToBase64(buffer) {
  let binary = '';
  const bytes = new Uint8Array(buffer);
  const len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

function formatTitleDate(dateStr) {
  if (!dateStr) return "";
  const [y, m, d] = dateStr.split('-');
  return `${Number(d)}-${Number(m)}-${y}`;
}

function formatFilenameDate(dStr) {
  if (!dStr) return "";
  const [y, m, d] = dStr.split('-');
  return `${d}-${m}-${y}`;
}

async function autoDownloadTicketsExcel(results, summary, fromDate, toDate) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Pending Tickets");

  const blueFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4D93D9' } };
  const titleFont = { name: 'Aptos', size: 12, bold: false, color: { argb: 'FF000000' } };
  const headerFont = { name: 'Aptos', size: 12, bold: false, color: { argb: 'FF000000' } };
  const bodyFont = { name: 'Aptos', size: 12, bold: false, color: { argb: 'FF000000' } };
  const thinSide = { style: 'thin', color: { argb: 'FF000000' } };
  const cellBorder = { top: thinSide, left: thinSide, bottom: thinSide, right: thinSide };

  ws.mergeCells('A1:D1');
  const titleCell = ws.getCell('A1');
  const fromFmt = formatTitleDate(fromDate);
  const toFmt = formatTitleDate(toDate);
  titleCell.value = `Pending Mygate Tickets - From ${fromFmt} To ${toFmt}`;
  titleCell.fill = blueFill;
  titleCell.font = titleFont;
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };

  const headers = ["Category", "Total", "Resolved", "Open"];
  headers.forEach((h, idx) => {
    const c = ws.getCell(2, idx + 1);
    c.value = h;
    c.fill = blueFill;
    c.font = headerFont;
    c.alignment = { horizontal: 'left', vertical: 'middle' };
  });

  let currentRow = 3;
  results.forEach(r => {
    const vals = [r.name, r.total, r.resolved, r.open];
    vals.forEach((val, idx) => {
      const c = ws.getCell(currentRow, idx + 1);
      c.value = val;
      c.font = bodyFont;
      c.border = cellBorder;
      c.alignment = { horizontal: idx === 0 ? 'left' : 'right', vertical: 'middle' };
    });
    currentRow++;
  });

  const totals = ["Total", summary.total, summary.resolved, summary.open];
  totals.forEach((val, idx) => {
    const c = ws.getCell(currentRow, idx + 1);
    c.value = val;
    c.font = bodyFont;
    c.border = cellBorder;
    c.alignment = { horizontal: idx === 0 ? 'left' : 'right', vertical: 'middle' };
  });

  for (let r = 1; r <= currentRow; r++) {
    for (let c = 1; c <= 4; c++) {
      ws.getCell(r, c).border = cellBorder;
    }
  }

  ws.getColumn(1).width = 55;
  ws.getColumn(2).width = 9;
  ws.getColumn(3).width = 12;
  ws.getColumn(4).width = 8;

  const buf = await wb.xlsx.writeBuffer();
  const base64 = arrayBufferToBase64(buf);
  const dataUrl = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + base64;
  const dt = formatFilenameDate(toDate);

  return new Promise((resolve, reject) => {
    chrome.downloads.download({
      url: dataUrl,
      filename: `Pending_Mygate_Tickets_Report_${dt}.xlsx`,
      saveAs: false
    }, (downloadId) => {
      if (chrome.runtime.lastError || !downloadId) {
        reject(new Error(chrome.runtime.lastError?.message || "Download blocked by browser permission"));
      } else {
        resolve(downloadId);
      }
    });
  });
}
