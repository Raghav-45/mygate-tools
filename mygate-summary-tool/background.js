/**
 * MyGate Summary Tool - Background Service Worker
 * Replicates Python Complaint Summary Sheet GraphQL scanning logic.
 * Persists scan state across popup closes and automatically downloads standardized Excel binary upon completion.
 */

importScripts('exceljs.min.js');

const GRAPHQL_URL = "https://api.dashboard.mygate.com/graphql/";
let currentScanAbort = false;

async function getAuthToken() {
  const tabs = await chrome.tabs.query({ url: "*://dashboard.mygate.com/*" });
  for (const tab of tabs) {
    try {
      const res = await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: () => {
          for (let i = 0; i < localStorage.length; i++) {
            const key = localStorage.key(i);
            const val = localStorage.getItem(key);
            if (val && val.length > 40 && !val.includes("{") && !val.includes(" ")) {
              return val;
            }
          }
          return null;
        }
      });
      if (res && res[0] && res[0].result) return res[0].result;
    } catch (e) {}
  }
  
  const cookies = await chrome.cookies.getAll({ domain: "mygate.com" });
  for (const c of cookies) {
    if ((c.name.toLowerCase().includes("token") || c.name.toLowerCase().includes("auth")) && c.value.length > 30) {
      return c.value;
    }
  }
  return null;
}

function epoch(dateStr) {
  const [dd, mm, yyyy] = dateStr.split('-').map(Number);
  return Math.floor(new Date(yyyy, mm - 1, dd, 0, 0, 0).getTime() / 1000);
}

async function getCount(fromDate, toDate, statuses, token, delayMs) {
  const payload = {
    operationName: "getAdminSrList",
    variables: {
      requestData: {
        requiredFields: ["id"],
        pagination: { count: 1, page: 1 },
        sorting: [],
        conditions: [
          { name: "date_filter", operation: "equal", values: ["created_date"] },
          { name: "from_date", values: [epoch(fromDate)], operation: "gte" },
          { name: "to_date", values: [epoch(toDate) + 86399], operation: "lte" },
          { name: "mygate_status", values: statuses, operation: "in" }
        ]
      }
    },
    query: `query getAdminSrList($requestData: DataListInput) {
      getAdminSrList(requestData: $requestData) {
        dataResponse { totalCount }
      }
    }`
  };

  const res = await fetch(GRAPHQL_URL, {
    method: "POST",
    headers: {
      "authorization": token,
      "content-type": "application/json",
      "origin": "https://dashboard.mygate.com",
      "referer": "https://dashboard.mygate.com/"
    },
    body: JSON.stringify(payload)
  });

  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const data = await res.json();
  if (data.errors || (data.data && !data.data.getAdminSrList)) {
    throw new Error("Session expired! Please refresh dashboard.mygate.com and log in.");
  }
  
  await new Promise(r => setTimeout(r, delayMs));
  
  return data?.data?.getAdminSrList?.dataResponse?.totalCount || 0;
}

function formatDDMMYYYY(date) {
  const dd = String(date.getDate()).padStart(2, '0');
  const mm = String(date.getMonth() + 1).padStart(2, '0');
  const yyyy = date.getFullYear();
  return `${dd}-${mm}-${yyyy}`;
}

const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
function formatDDMMMYYYY(date) {
  const dd = String(date.getDate()).padStart(2, '0');
  const mmm = MONTH_NAMES[date.getMonth()];
  const yyyy = date.getFullYear();
  return `${dd}-${mmm}-${yyyy}`;
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

function parseDateToTime(dStr) {
  const [dd, mmm, yyyy] = dStr.split('-');
  const mIdx = MONTH_NAMES.indexOf(mmm);
  return new Date(yyyy, mIdx, Number(dd)).getTime();
}

function parseDateToUTCNoon(dStr) {
  const [dd, mmm, yyyy] = dStr.split('-');
  const mIdx = MONTH_NAMES.indexOf(mmm);
  return new Date(Date.UTC(Number(yyyy), mIdx, Number(dd), 12, 0, 0));
}

async function autoDownloadExcel(rows, year, month) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");
  ws.views = [{ showGridLines: true }];

  ws.columns = [
    { width: 27.14 },
    { width: 14.14 },
    { width: 13.86 },
    { width: 11.57 },
    { width: 12.00 }
  ];

  ws.mergeCells("A1:E1");
  const r1 = ws.getRow(1);
  r1.height = 18;
  const tCell = r1.getCell(1);
  tCell.value = "Complaint Summary Sheet";
  tCell.font = { name: "Calibri", size: 14, bold: true, color: { argb: "FF000000" } };
  tCell.alignment = { horizontal: "center", vertical: "center" };

  for (let c = 1; c <= 5; c++) {
    const cell = r1.getCell(c);
    cell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } }
    };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
  }

  const headers = ["Date", "Previous day Open Complaints", "Today Received Complaints", "Today Closed Complaints", "Pending"];
  const r2 = ws.getRow(2);
  r2.values = headers;
  r2.height = 45;

  for (let c = 1; c <= 5; c++) {
    const cell = r2.getCell(c);
    cell.font = { name: "Calibri", size: 11, color: { argb: "FF000000" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
    cell.alignment = { horizontal: "center", vertical: "center", wrapText: true };
    cell.border = {
      top: { style: "thin", color: { argb: "FF000000" } },
      bottom: { style: "thin", color: { argb: "FF000000" } },
      left: { style: "thin", color: { argb: "FF000000" } },
      right: { style: "thin", color: { argb: "FF000000" } }
    };
  }

  const sorted = [...rows].sort((a, b) => parseDateToTime(a.date) - parseDateToTime(b.date));
  sorted.forEach((r, idx) => {
    const rowNum = idx + 3;
    const row = ws.getRow(rowNum);
    const dateObj = parseDateToUTCNoon(r.date);

    row.getCell(1).value = dateObj;
    row.getCell(2).value = r.prevOpen;
    row.getCell(3).value = r.received;
    row.getCell(4).value = r.closed;
    row.getCell(5).value = { formula: `(B${rowNum}+C${rowNum})-D${rowNum}`, result: r.pending };

    for (let c = 1; c <= 5; c++) {
      const cell = row.getCell(c);
      cell.font = { name: "Calibri", size: 11, color: { argb: "FF000000" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFFFF" } };
      cell.alignment = { horizontal: "center", vertical: "center" };
      cell.border = {
        top: { style: "thin", color: { argb: "FF000000" } },
        bottom: { style: "thin", color: { argb: "FF000000" } },
        left: { style: "thin", color: { argb: "FF000000" } },
        right: { style: "thin", color: { argb: "FF000000" } }
      };
      if (c === 1) cell.numFmt = "dd-mm-yyyy";
    }
  });

  const buf = await wb.xlsx.writeBuffer();
  const base64 = arrayBufferToBase64(buf);
  const dataUrl = 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + base64;
  const monthShort = MONTH_NAMES[month - 1];

  return new Promise((resolve, reject) => {
    chrome.downloads.download({
      url: dataUrl,
      filename: `Complaint Summary Sheet - ${monthShort}-${year}.xlsx`,
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

async function runSummaryScan(year, month, delaySec) {
  const token = await getAuthToken();
  if (!token) {
    throw new Error("No MyGate token found! Please open dashboard.mygate.com and log in.");
  }
  
  const delayMs = Math.round(delaySec * 1000);
  const startDate = new Date(year, month - 1, 1);
  const lastDay = new Date(year, month, 0).getDate();
  let endDate = new Date(year, month - 1, lastDay);
  
  const now = new Date();
  if (endDate > now) endDate = new Date(now);
  endDate.setHours(23, 59, 59, 999);
  
  const totalDays = Math.floor((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
  let current = new Date(startDate);
  let dayIdx = 0;
  let rows = [];
  
  await chrome.storage.local.set({
    summaryScanState: { isScanning: true, year, month, pct: 0, stepText: "Connecting MyGate API...", rows: [], isDone: false }
  });

  while (current <= endDate) {
    if (currentScanAbort) {
      await chrome.storage.local.set({ summaryScanState: { isScanning: false, isAborted: true } });
      chrome.runtime.sendMessage({ type: "SUMMARY_ABORTED" }).catch(() => {});
      return;
    }
    
    const dStr = formatDDMMYYYY(current);
    const pct = Math.round((dayIdx / totalDays) * 100);
    
    await chrome.storage.local.set({
      summaryScanState: { isScanning: true, year, month, pct, stepText: `Processing ${dStr}...`, rows, isDone: false }
    });
    chrome.runtime.sendMessage({ type: "SUMMARY_PROGRESS", stepText: `Processing ${dStr}...`, pct }).catch(() => {});
    
    const prevOpen = await getCount("01-01-2024", dStr, ["open", "hold", "re_opened", "in_progress", "job_done"], token, delayMs);
    if (currentScanAbort) return;
    
    const received = await getCount(dStr, dStr, ["open", "hold", "re_opened", "in_progress", "job_done", "closed"], token, delayMs);
    if (currentScanAbort) return;
    
    const closed = await getCount(dStr, dStr, ["closed"], token, delayMs);
    if (currentScanAbort) return;
    
    const pending = prevOpen + received - closed;
    const row = {
      date: formatDDMMMYYYY(current),
      prevOpen,
      received,
      closed,
      pending
    };
    rows.push(row);
    rows.sort((a, b) => parseDateToTime(a.date) - parseDateToTime(b.date));
    
    dayIdx++;
    const newPct = Math.round((dayIdx / totalDays) * 100);
    
    await chrome.storage.local.set({
      summaryScanState: { isScanning: true, year, month, pct: newPct, stepText: `Completed ${dStr}`, rows, isDone: false }
    });
    chrome.runtime.sendMessage({ type: "SUMMARY_PROGRESS", stepText: `Completed ${dStr}`, pct: newPct, row }).catch(() => {});
    
    current.setDate(current.getDate() + 1);
  }
  
  // Automatically generate & download Excel
  await chrome.storage.local.set({
    summaryScanState: { isScanning: true, year, month, pct: 100, stepText: "Downloading Excel...", rows, isDone: false }
  });
  chrome.runtime.sendMessage({ type: "SUMMARY_PROGRESS", stepText: "Downloading Excel...", pct: 100 }).catch(() => {});
  
  await autoDownloadExcel(rows, year, month);

  await chrome.storage.local.set({
    summaryScanState: { isScanning: false, year, month, pct: 100, stepText: "Report Generated!", rows, isDone: true }
  });
  chrome.runtime.sendMessage({ type: "SUMMARY_DONE" }).catch(() => {});
}

chrome.runtime.onMessage.addListener((req, sender, sendResponse) => {
  if (req.type === "CHECK_AUTH_STATUS") {
    getAuthToken().then(t => sendResponse({ hasToken: !!t }));
    return true;
  }
  if (req.type === "ABORT_SUMMARY_SCAN") {
    currentScanAbort = true;
    sendResponse({ aborted: true });
    return true;
  }
  if (req.type === "START_SUMMARY_SCAN") {
    currentScanAbort = false;
    runSummaryScan(req.year, req.month, req.requestDelay || 1.5).catch(err => {
      chrome.storage.local.set({ summaryScanState: { isScanning: false, error: err.message } });
      chrome.runtime.sendMessage({ type: "SUMMARY_ERROR", error: err.message }).catch(() => {});
    });
    sendResponse({ started: true });
    return true;
  }
});
