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

// A history row identifies a job; its URL alone does not. MyGate can reuse a
// cloud file URL across jobs, and historical successes remain on this page.
function reportIdentity(report) {
  const filters = report.download_filters || {}
  return JSON.stringify([report.report_name, report.download_request_time,
    Object.keys(filters).sort().map(key => [key, Array.isArray(filters[key])
      ? filters[key].map(v => String(v).trim()).sort() : filters[key]])])
}

function cloudFileKey(link) {
  return String(link || '').split(/[?#]/)[0]
}

function matchesExportFilters(filters, from, to) {
  const matchesDates = (keys, expected) => {
    const values = keys.filter(key => filters[key] != null).map(key => filters[key])
    return values.length > 0 && values.every(value => value === expected)
  }
  const statuses = Array.isArray(filters.Status) ? filters.Status.map(v => String(v).trim().toLowerCase()).sort() : []
  return matchesDates(['From', 'Date From'], from) && matchesDates(['To', 'Date To'], to)
    && JSON.stringify(statuses) === JSON.stringify(['new', 'reopened', 'in progress', 'job done', 'on hold'].sort())
}

function createdDateKey(value) {
  let year, month, day
  if (Object.prototype.toString.call(value) === '[object Date]') {
    year = value.getUTCFullYear(); month = value.getUTCMonth() + 1; day = value.getUTCDate()
  } else {
    const text = String(value ?? '').trim()
    let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:$|[T ,])/)
    if (match) [, year, month, day] = match.map(Number)
    else {
      match = text.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4}|\d{2})(?:$|[ ,])/)
      if (!match) return null
      day = Number(match[1]); month = Number(match[2]); year = Number(match[3])
      if (year < 100) year += 2000
    }
  }
  const date = new Date(Date.UTC(year, month - 1, day))
  if (date.getUTCFullYear() !== year || date.getUTCMonth() + 1 !== month || date.getUTCDate() !== day) return null
  return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`
}

function categoryKey(value) {
  return String(value).trim().replace(/\s+/g, ' ').toLowerCase()
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
  const { fromDate, toDate, requestDelayMs = 2000, expectedOpenByCategory } = params

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
  const allMergedRows = []
  const usedReportLinks = new Set()
  const seenTicketIds = new Set()
  const categoryCounts = new Map()
  const fail = async (cr, message) => {
    if (cr) cr.status = 'Stopped'
    await chrome.storage.local.set({ dumpScanState: {
      isScanning: false, isFailed: true, isDone: false, pct: 0,
      statusText: `${message} No export retry was made.`, error: message,
      chunks: chunkResults, totalRows: totalRowsMerged,
    } })
    chrome.runtime.sendMessage({ type: 'DUMP_FINISHED', error: message, chunks: chunkResults, totalRows: totalRowsMerged }).catch(() => {})
  }
  const pollPayload = {
    operationName: 'getDownloadReportList',
    variables: { requestData: {
      requiredFields: ['report_link', 'report_name', 'status', 'download_request_time', 'download_filters'],
      pagination: { count: 20, page: 1 }, sorting: [], conditions: [],
    } },
    query: 'query getDownloadReportList($requestData: DataListInput) { getDownloadReportList(requestData: $requestData) { dataResponse { data totalCount } message statusCode success } }',
  }
  const readReports = async () => {
    const res = await fetch(URL_GRAPHQL, { method: 'POST', headers, cache: 'no-store', body: JSON.stringify(pollPayload) })
    const data = await res.json()
    const result = data?.data?.getDownloadReportList
    if (!res.ok || data.errors || result?.success === false || !Array.isArray(result?.dataResponse?.data)) {
      throw new Error('Could not read MyGate export history. Stopped before accepting a cloud file.')
    }
    return result.dataResponse.data
  }

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

    let baselineIds, baselineLinks, newestBaselineTime, expectedRows
    try {
      if (i > 0) await sleep(requestDelayMs)
      const baseline = await readReports()
      baselineIds = new Set(baseline.map(reportIdentity))
      baselineLinks = new Set(baseline.filter(r => r.report_link).map(r => cloudFileKey(r.report_link)))
      newestBaselineTime = Math.max(0, ...baseline.filter(r => r.report_name === 'Helpdesk Report')
        .map(r => Date.parse(r.download_request_time)).filter(Number.isFinite))

      // A read-only count verifies completeness. This never requests a cloud
      // export. Exactly one isDownload:true request is allowed per chunk.
      await sleep(requestDelayMs)
      if (activeScanAbort) { await handleAbort(); return }
      const countRequest = { ...payload.variables.requestData, isDownload: false,
        requiredFields: ['id'], pagination: { count: 1, page: 1 } }
      delete countRequest.downloadFilters
      const countRes = await fetch(URL_GRAPHQL, { method: 'POST', headers, cache: 'no-store',
        body: JSON.stringify({ ...payload, variables: { requestData: countRequest } }) })
      const countData = await countRes.json()
      const countResult = countData?.data?.getAdminSrList
      const count = countResult?.dataResponse?.totalCount
      expectedRows = Number(count)
      if (!countRes.ok || countData.errors || countResult?.success === false || count == null
          || count === '' || !Number.isSafeInteger(expectedRows) || expectedRows < 0) {
        throw new Error('Could not verify the expected open-ticket count. Export was not requested.')
      }
      cr.expectedRows = expectedRows
      if (activeScanAbort) { await handleAbort(); return }
      await sleep(requestDelayMs)
      if (activeScanAbort) { await handleAbort(); return }
      const res = await fetch(URL_GRAPHQL, { method: 'POST', headers, body: JSON.stringify(payload) })
      const data = await res.json()
      const result = data?.data?.getAdminSrList
      if (!res.ok || data.errors || !result || result.success === false) {
        throw new Error('MyGate rejected the export request. Check the dashboard session.')
      }
      cr.status = 'Generating in Cloud...'
      cr.rowsFound = 'Waiting for Cloud...'
      await updateState(basePct + 5, `Generating report for ${cr.rangeStr}...`, chunkResults, totalRowsMerged)
    } catch (e) {
      await fail(cr, e.message)
      return
    }

    // Poll only the status of this one request; never resubmit an export.
    let chunkReady = false
    let pollAttempts = 0
    while (!chunkReady && pollAttempts < 45) {
      if (activeScanAbort) {
        await handleAbort()
        return
      }
      pollAttempts++
      await sleep(requestDelayMs)
      if (activeScanAbort) { await handleAbort(); return }

      try {
        const reports = await readReports()
        const match = reports.find(r => r.report_name === 'Helpdesk Report' && r.status === 'Success'
          && r.report_link && !baselineIds.has(reportIdentity(r))
          && Number.isFinite(Date.parse(r.download_request_time))
          && Date.parse(r.download_request_time) >= newestBaselineTime
          && matchesExportFilters(r.download_filters || {}, filterFrom, filterTo))
        if (match) {
          const key = cloudFileKey(match.report_link)
          if (baselineLinks.has(key) || usedReportLinks.has(key)
              || reports.some(r => reportIdentity(r) !== reportIdentity(match) && cloudFileKey(r.report_link) === key)) {
            throw new Error('MyGate reused a cloud file for different export requests. Dump stopped to prevent incorrect counts.')
          }
          usedReportLinks.add(key)
          cr.downloadUrl = match.report_link
          chunkReady = true
        }
      } catch (e) {
        await fail(cr, e.message)
        return
      }
    }

    if (!cr.downloadUrl) {
      await fail(cr, 'The new MyGate export did not become ready in time. No partial dump was saved.')
      return
    }

    if (activeScanAbort) { await handleAbort(); return }
    cr.status = 'Downloading & Parsing...'
    await updateState(basePct + 15, `Merging ${cr.rangeStr}...`, chunkResults, totalRowsMerged)

    try {
      const fileRes = await fetch(cr.downloadUrl, { cache: 'no-store' })
      if (!fileRes.ok) throw new Error('Cloud file download failed. No partial dump was saved.')
      const arrayBuf = await fileRes.arrayBuffer()
      const chunkWb = new ExcelJS.Workbook()
      await chunkWb.xlsx.load(arrayBuf)
      const ws = chunkWb.worksheets[0]

      if (!ws || ws.rowCount < 3) throw new Error('Cloud export has no valid ticket headers.')
      const colMap = {}
      ws.getRow(3).eachCell((cell, idx) => {
        const name = cell.text.toLowerCase().trim()
        if (['id', 'ticket id', 'i.d'].includes(name)) colMap.id = idx
        else if (['created date', 'date'].includes(name)) colMap.createdDate = idx
        else if (name === 'category') colMap.category = idx
        else if (['sub category', 'subcategory'].includes(name)) colMap.subCategory = idx
        else if (['flat', 'house'].includes(name)) colMap.flat = idx
        else if (['subject', 'description'].includes(name)) colMap.subject = idx
        else if (['status', 'mygate status'].includes(name)) colMap.status = idx
      })
      if (['id', 'createdDate', 'category', 'status'].some(key => !colMap[key])) {
        throw new Error('Cloud export is missing required ticket columns.')
      }
      const chunkRows = []
      const chunkIds = new Set()
      for (let r = 4; r <= ws.rowCount; r++) {
        const row = ws.getRow(r)
        const id = row.getCell(colMap.id).text.trim()
        if (!id) {
          if (row.actualCellCount) throw new Error('Cloud export contains a row without a ticket ID.')
          continue
        }
        if (seenTicketIds.has(id) || chunkIds.has(id)) throw new Error('Duplicate ticket IDs found in the cloud export.')
        const date = createdDateKey(row.getCell(colMap.createdDate).value)
        if (!date || date < chunk.fromDate || date > chunk.toDate) {
          throw new Error(`Cloud export contains a created date outside the requested date range (${cr.rangeStr}).`)
        }
        const status = row.getCell(colMap.status).text.trim().toLowerCase().replace(/[_\s]+/g, ' ')
        if (!['new', 'open', 'reopened', 're opened', 'in progress', 'job done', 'on hold', 'hold'].includes(status)) {
          throw new Error('Cloud export contains a ticket status outside the requested open statuses.')
        }
        const item = {}
        for (const [name, col] of Object.entries(colMap)) item[name] = row.getCell(col).value ?? ''
        chunkRows.push(item)
        chunkIds.add(id)
      }
      const rowsInChunk = chunkRows.length
      if (rowsInChunk !== expectedRows) {
        throw new Error(`Open-ticket count mismatch for ${cr.rangeStr}: expected ${expectedRows}, received ${rowsInChunk}. No partial dump was saved.`)
      }
      chunkRows.forEach(item => {
        const key = categoryKey(item.category)
        categoryCounts.set(key, (categoryCounts.get(key) || 0) + 1)
      })
      chunkIds.forEach(id => seenTicketIds.add(id))
      allMergedRows.push(...chunkRows)

      cr.rowsFound = rowsInChunk.toLocaleString()
      cr.status = 'Merged ✅'
      totalRowsMerged += rowsInChunk
      await updateState(basePct + 25, `Merged ${rowsInChunk} rows from ${cr.rangeStr}`, chunkResults, totalRowsMerged)
    } catch (e) {
      console.error('Error downloading chunk:', e)
      await fail(cr, e.message)
      return
    }
  }

  if (expectedOpenByCategory) {
    const expected = new Map(expectedOpenByCategory.map(row => [categoryKey(row.category), Number(row.open)]))
    const keys = new Set([...expected.keys(), ...categoryCounts.keys()])
    if ([...keys].some(key => (expected.get(key) || 0) !== (categoryCounts.get(key) || 0))) {
      await fail(null, 'Open-ticket counts differ between Pending Report and Master Dump. The data changed or MyGate returned a mismatched export; no dump or analysis was saved.')
      return
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
