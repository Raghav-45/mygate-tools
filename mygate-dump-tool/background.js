/**
 * MyGate Dump Tool - Background Service Worker
 * Autonomous Multi-Year Ticket Dump Exporter.
 * Slices date ranges > 1 year into 1-year chunks, requests cloud exports, polls for completion, merges and downloads.
 */

importScripts('exceljs.min.js')

const SAMPLE_FALLBACK_TOKEN =
  'zbdzHQCrz1uOVa3Z9QrFabEIA600Udb6lZPrM2SFIkC597iOyCDKllxwR9ZD7Jqa'
const URL_GRAPHQL = 'https://api.dashboard.mygate.com/graphql/'

let activeScanAbort = false

// Discover active authorization token from open MyGate dashboard tab
async function discoverActiveAuthToken() {
  try {
    const tabs = await chrome.tabs.query({ url: '*://*.mygate.com/*' })
    for (const tab of tabs) {
      try {
        const results = await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: () => {
            for (let i = 0; i < localStorage.length; i++) {
              const key = localStorage.key(i)
              if (key && key.toLowerCase().includes('token')) {
                const val = localStorage.getItem(key)
                if (val && val.length > 20 && !val.includes('{')) return val
              }
            }
            return null
          },
        })
        if (results && results[0] && results[0].result) {
          return results[0].result
        }
      } catch (e) {}
    }
  } catch (err) {}
  return null
}

// Convert YYYY-MM-DD to midnight epoch seconds
function getMidnightEpoch(dateStr) {
  const [y, m, d] = dateStr.split('-')
  return Math.floor(
    new Date(Number(y), Number(m) - 1, Number(d), 0, 0, 0).getTime() / 1000,
  )
}

// Format YYYY-MM-DD to MyGate filter format d:m:yyyy (no leading zeros)
function formatFilterDate(dateStr) {
  const [y, m, d] = dateStr.split('-')
  return `${Number(d)}:${Number(m)}:${Number(y)}`
}

// Format YYYY-MM-DD to display header format d-m-yyyy (no leading zeros, hyphenated)
function formatHeaderDate(dateStr) {
  const [y, m, d] = dateStr.split('-')
  return `${Number(d)}-${Number(m)}-${Number(y)}`
}

// Slice multi-year date range into 1-year windows (max 365 days each)
function sliceIntoYears(fromDateStr, toDateStr) {
  const chunks = []
  let currStart = new Date(fromDateStr)
  const finalEnd = new Date(toDateStr)

  while (currStart <= finalEnd) {
    let currEnd = new Date(currStart)
    currEnd.setFullYear(currEnd.getFullYear() + 1)
    currEnd.setDate(currEnd.getDate() - 1) // 1 year minus 1 day

    if (currEnd > finalEnd) {
      currEnd = new Date(finalEnd)
    }

    const sStr = currStart.toISOString().split('T')[0]
    const eStr = currEnd.toISOString().split('T')[0]
    chunks.push({ fromDate: sStr, toDate: eStr })

    currStart = new Date(currEnd)
    currStart.setDate(currStart.getDate() + 1)
  }
  return chunks.reverse()
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms))
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message.type === 'START_DUMP_EXPORT') {
    activeScanAbort = false
    runMasterDumpExport(message.params)
  } else if (message.type === 'ABORT_DUMP_EXPORT') {
    activeScanAbort = true
  }
  return true
})

async function runMasterDumpExport(params) {
  const { fromDate, toDate, requestDelayMs = 2000 } = params

  let tokenToUse = await discoverActiveAuthToken()
  if (!tokenToUse) tokenToUse = SAMPLE_FALLBACK_TOKEN

  const chunks = sliceIntoYears(fromDate, toDate)
  const chunkResults = []
  let totalRowsMerged = 0

  await chrome.storage.local.set({
    dumpScanState: {
      isScanning: true,
      pct: 0,
      statusText: `Sliced into ${chunks.length} yearly chunks. Initiating...`,
      chunks: [],
      totalRows: 0,
      isDone: false,
    },
  })

  const headers = {
    authorization: tokenToUse,
    'content-type': 'application/json',
    origin: 'https://dashboard.mygate.com',
    referer: 'https://dashboard.mygate.com/',
  }

  const masterWb = new ExcelJS.Workbook()
  const masterWs = masterWb.addWorksheet('Master Dump')
  let headerRowValues = null
  const allMergedRows = []
  const usedReportLinks = new Set()

  // Process chunks strictly sequentially to prevent MyGate cloud job collisions
  for (let i = 0; i < chunks.length; i++) {
    if (activeScanAbort) {
      await handleAbort()
      return
    }

    const chunk = chunks[i]
    const filterFrom = formatFilterDate(chunk.fromDate)
    const filterTo = formatFilterDate(chunk.toDate)
    const cr = {
      rangeStr: `${chunk.fromDate} to ${chunk.toDate}`,
      filterFrom,
      filterTo,
      rowsFound: 'Requesting...',
      status: 'Initiating...',
      downloadUrl: null,
    }
    chunkResults.push(cr)

    const basePct = Math.round((i / chunks.length) * 85)
    await updateState(basePct, `Processing ${cr.rangeStr}...`, chunkResults, totalRowsMerged)

    const payload = {
      operationName: 'getAdminSrList',
      variables: {
        requestData: {
          requiredFields: [
            'id',
            'number',
            'subject',
            'category',
            'sub_category',
            'house',
            'assignee',
            'mygate_status',
            'escalated_group',
            'defaulter',
            'updated_date',
            'urgent',
            'highlight_ticket',
            'ageing',
          ],
          pagination: { count: 25, page: 1 },
          sorting: [],
          conditions: [
            { name: 'date_filter', operation: 'equal', values: ['created_date'] },
            { name: 'mygate_status', values: ['open', 're_opened', 'in_progress', 'job_done', 'hold'], operation: 'in' },
            { name: 'from_date', values: [getMidnightEpoch(chunk.fromDate)], operation: 'gte' },
            { name: 'to_date', values: [getMidnightEpoch(chunk.toDate) + 86399], operation: 'lte' },
          ],
          isDownload: true,
          downloadFilters: {
            Status: ['New ', 'Reopened ', 'In Progress ', 'Job Done ', 'On Hold '],
            From: filterFrom,
            To: filterTo,
            'Date To': filterTo,
            'Date From': filterFrom,
          },
        },
      },
      query:
        'query getAdminSrList($requestData: DataListInput) {\n  getAdminSrList(requestData: $requestData) {\n    dataResponse {\n      data\n      filterType\n      totalCount\n      __typename\n    }\n    message\n    statusCode\n    success\n    __typename\n  }\n}\n',
    }

    try {
      const res = await fetch(URL_GRAPHQL, {
        method: 'POST',
        headers,
        body: JSON.stringify(payload),
      })
      const data = await res.json()
      if (!res.ok || data.errors || !data?.data?.getAdminSrList) {
        throw new Error('Session expired! Please open dashboard.mygate.com and log in.')
      }
      cr.status = 'Generating in Cloud...'
      cr.rowsFound = 'Waiting for Cloud...'
      await updateState(basePct + 5, `Generating report for ${cr.rangeStr}...`, chunkResults, totalRowsMerged)
    } catch (e) {
      console.error('Error requesting export:', e)
      cr.rowsFound = 'Auth Error'
      cr.status = 'Failed (Login Required)'
      await updateState(0, 'Error: Please open dashboard.mygate.com and log in!', chunkResults, totalRowsMerged)
      return
    }

    // Poll until this chunk report is ready
    let chunkReady = false
    let pollAttempts = 0
    const pollPayload = {
      operationName: 'getDownloadReportList',
      variables: {
        requestData: {
          requiredFields: [
            'report_link',
            'report_name',
            'status',
            'download_request_time',
            'download_filters',
          ],
          pagination: { count: 20, page: 1 },
          sorting: [],
          conditions: [],
        },
      },
      query:
        'query getDownloadReportList($requestData: DataListInput) {\n  getDownloadReportList(requestData: $requestData) {\n    dataResponse {\n      data\n      filterType\n      totalCount\n      __typename\n    }\n    message\n    statusCode\n    success\n    __typename\n  }\n}\n',
    }

    while (!chunkReady && pollAttempts < 45) {
      if (activeScanAbort) {
        await handleAbort()
        return
      }
      pollAttempts++
      await sleep(requestDelayMs)

      try {
        const pollRes = await fetch(URL_GRAPHQL, {
          method: 'POST',
          headers,
          body: JSON.stringify(pollPayload),
        })
        const pollData = await pollRes.json()
        const reports = pollData?.data?.getDownloadReportList?.dataResponse?.data || []

        const match = reports.find((r) => {
          if (r.report_name !== 'Helpdesk Report' || r.status !== 'Success' || !r.report_link || usedReportLinks.has(r.report_link))
            return false
          const df = r.download_filters || {}
          const matchFrom = df['From'] === filterFrom || df['Date From'] === filterFrom
          const matchTo = df['To'] === filterTo || df['Date To'] === filterTo
          return matchFrom && matchTo
        })

        if (match && match.report_link) {
          usedReportLinks.add(match.report_link)
          cr.downloadUrl = match.report_link
          chunkReady = true
        }
      } catch (e) {
        console.error('Polling error:', e)
      }
    }

    if (!cr.downloadUrl) {
      cr.status = 'Timed out'
      cr.rowsFound = '0'
      continue
    }

    cr.status = 'Downloading & Parsing...'
    await updateState(basePct + 15, `Merging ${cr.rangeStr}...`, chunkResults, totalRowsMerged)

    try {
      const fileRes = await fetch(cr.downloadUrl)
      const arrayBuf = await fileRes.arrayBuffer()
      const chunkWb = new ExcelJS.Workbook()
      await chunkWb.xlsx.load(arrayBuf)
      const ws = chunkWb.worksheets[0]

      if (!headerRowValues && ws.rowCount >= 3) {
        headerRowValues = ws.getRow(3).values
      }

      let colMap = { id: 1, createdDate: 2, category: 4, subCategory: 5, flat: 7, subject: 9, status: 10 }
      if (headerRowValues) {
        headerRowValues.forEach((val, idx) => {
          if (!val) return
          const s = String(val).toLowerCase().trim()
          if (s === 'id' || s === 'ticket id' || s === 'i.d') colMap.id = idx
          else if (s === 'created date' || s === 'date') colMap.createdDate = idx
          else if (s === 'category') colMap.category = idx
          else if (s === 'sub category' || s === 'subcategory') colMap.subCategory = idx
          else if (s === 'flat' || s === 'house') colMap.flat = idx
          else if (s === 'subject' || s === 'description') colMap.subject = idx
          else if (s === 'status' || s === 'mygate status') colMap.status = idx
        })
      }

      let rowsInChunk = 0
      for (let r = 4; r <= ws.rowCount; r++) {
        const rVals = ws.getRow(r).values
        if (rVals && rVals.length > 1 && rVals[colMap.id] !== undefined && rVals[colMap.id] !== '') {
          const getVal = (idx) => {
            let v = rVals[idx]
            if (v && v instanceof Date) {
              return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate(), 12, 0, 0))
            }
            return v !== undefined && v !== null ? v : ''
          }
          allMergedRows.push({
            id: getVal(colMap.id),
            createdDate: getVal(colMap.createdDate),
            category: getVal(colMap.category),
            subCategory: getVal(colMap.subCategory),
            flat: getVal(colMap.flat),
            subject: getVal(colMap.subject),
            status: getVal(colMap.status),
          })
          rowsInChunk++
        }
      }

      cr.rowsFound = rowsInChunk.toLocaleString()
      cr.status = 'Merged ✅'
      totalRowsMerged += rowsInChunk
      await updateState(basePct + 25, `Merged ${rowsInChunk} rows from ${cr.rangeStr}`, chunkResults, totalRowsMerged)
    } catch (e) {
      console.error('Error downloading chunk:', e)
      cr.status = 'Merge Error'
    }
  }

  // Step 4: Build Standardized Master Spreadsheet and Auto-Download
  await updateState(
    96,
    'Compiling standardized master spreadsheet...',
    chunkResults,
    totalRowsMerged,
  )

  const blueFill = {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: 'FF4D93D9' },
  }
  const titleFont = {
    name: 'Aptos',
    size: 12,
    bold: false,
    color: { argb: 'FF000000' },
  }
  const headerFont = {
    name: 'Aptos',
    size: 12,
    bold: false,
    color: { argb: 'FF000000' },
  }
  const bodyFont = {
    name: 'Aptos',
    size: 12,
    bold: false,
    color: { argb: 'FF000000' },
  }
  const thinSide = { style: 'thin', color: { argb: 'FF000000' } }
  const cellBorder = {
    top: thinSide,
    left: thinSide,
    bottom: thinSide,
    right: thinSide,
  }

  // Row 1: Title
  masterWs.mergeCells('A1:H1')
  for (let c = 1; c <= 8; c++) {
    const cell = masterWs.getCell(1, c)
    cell.fill = blueFill
    cell.border = cellBorder
  }
  const titleCell = masterWs.getCell('A1')
  titleCell.value = 'DLF Independent Floors'
  titleCell.font = titleFont
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' }
  masterWs.getRow(1).height = 24

  // Row 2: Subtitle
  masterWs.mergeCells('A2:H2')
  for (let c = 1; c <= 8; c++) {
    const cell = masterWs.getCell(2, c)
    cell.fill = blueFill
    cell.border = cellBorder
  }
  const subCell = masterWs.getCell('A2')
  subCell.value = `Help Desk Report: From ${formatHeaderDate(fromDate)} To ${formatHeaderDate(toDate)}`
  subCell.font = titleFont
  subCell.alignment = { horizontal: 'center', vertical: 'middle' }
  masterWs.getRow(2).height = 20

  // Row 3: Headers
  const colHeaders = [
    'Sr No.',
    'I.D',
    'Created Date',
    'Category',
    'Sub Category',
    'Flat',
    'Subject',
    'Status',
  ]
  const hRow = masterWs.getRow(3)
  hRow.height = 24
  colHeaders.forEach((h, idx) => {
    const cell = hRow.getCell(idx + 1)
    cell.value = h
    cell.fill = blueFill
    cell.font = headerFont
    cell.border = cellBorder
    cell.alignment = { horizontal: 'center', vertical: 'middle' }
  })

  // Rows 4+: Data
  let currRowIdx = 4
  allMergedRows.forEach((item, idx) => {
    const row = masterWs.getRow(currRowIdx)
    const vals = [
      idx + 1, // Sr No. (1-indexed all the way)
      item.id,
      item.createdDate,
      item.category,
      item.subCategory,
      item.flat,
      item.subject,
      item.status,
    ]
    vals.forEach((val, cIdx) => {
      const cell = row.getCell(cIdx + 1)
      cell.value = val
      cell.font = bodyFont
      cell.border = cellBorder
      // Center Sr No., I.D, Created Date, Sub Category, Flat, Status; left align Category and Subject
      const isLeft = cIdx === 3 || cIdx === 6
      cell.alignment = {
        horizontal: isLeft ? 'left' : 'center',
        vertical: 'middle',
        wrapText: cIdx === 6,
      }
    })
    currRowIdx++
  })

  // Set column widths
  masterWs.getColumn(1).width = 10 // Sr No.
  masterWs.getColumn(2).width = 12 // I.D
  masterWs.getColumn(3).width = 22 // Created Date
  masterWs.getColumn(4).width = 28 // Category
  masterWs.getColumn(5).width = 20 // Sub Category
  masterWs.getColumn(6).width = 16 // Flat
  masterWs.getColumn(7).width = 45 // Subject
  masterWs.getColumn(8).width = 15 // Status

  const buf = await masterWb.xlsx.writeBuffer()
  const base64 = arrayBufferToBase64(buf)
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' +
    base64
  const filename = `MyGate_Master_Helpdesk_Dump_${fromDate}_to_${toDate}.xlsx`

  await new Promise((resolve, reject) => {
    chrome.downloads.download(
      {
        url: dataUrl,
        filename: filename,
        saveAs: false,
      },
      (downloadId) => {
        if (chrome.runtime.lastError || !downloadId) {
          reject(
            new Error(
              chrome.runtime.lastError?.message ||
                'Download blocked by browser permission',
            ),
          )
        } else {
          resolve(downloadId)
        }
      },
    )
  })

  await chrome.storage.local.set({
    dumpScanState: {
      isScanning: false,
      pct: 100,
      statusText: 'Master Dump Downloaded!',
      chunks: chunkResults,
      totalRows: totalRowsMerged,
      isDone: true,
    },
  })

  chrome.runtime
    .sendMessage({
      type: 'DUMP_FINISHED',
      chunks: chunkResults,
      totalRows: totalRowsMerged,
    })
    .catch(() => {})
}

function arrayBufferToBase64(buffer) {
  let binary = ''
  const bytes = new Uint8Array(buffer)
  const len = bytes.byteLength
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i])
  }
  return btoa(binary)
}

async function updateState(pct, statusText, chunks, totalRows) {
  await chrome.storage.local.set({
    dumpScanState: {
      isScanning: true,
      pct,
      statusText,
      chunks,
      totalRows,
      isDone: false,
    },
  })
  chrome.runtime
    .sendMessage({
      type: 'DUMP_PROGRESS_UPDATE',
      pct,
      statusText,
      chunks,
      totalRows,
    })
    .catch(() => {})
}

async function handleAbort() {
  await chrome.storage.local.set({
    dumpScanState: { isScanning: false, isAborted: true },
  })
  chrome.runtime.sendMessage({ type: 'DUMP_ABORTED' }).catch(() => {})
}
