/**
 * MyGate Dashboard Tool — Popup logic
 *
 * User picks a Master Helpdesk Dump (.xlsx); the file's bytes go to the
 * background service worker, which builds the standardized dashboard workbook
 * (byte-identical to the reference Python tool) and auto-downloads it.
 */

const buildBtn = document.getElementById('buildBtn')
const dumpFileInput = document.getElementById('dumpFile')
const fileNameBox = document.getElementById('fileNameBox')
const asOnDateInput = document.getElementById('asOnDate')
const alertBanner = document.getElementById('alertBanner')
const progressContainer = document.getElementById('progressContainer')
const progressFill = document.getElementById('progressFill')
const progressPctText = document.getElementById('progressPctText')
const progressStepText = document.getElementById('progressStepText')
const resultsSection = document.getElementById('resultsSection')
const autoDlBanner = document.getElementById('autoDlBanner')
const kpiTotal = document.getElementById('kpiTotal')
const kpiInProgress = document.getElementById('kpiInProgress')
const kpiMonths = document.getElementById('kpiMonths')

const aboutIconBtn = document.getElementById('aboutIconBtn')
const aboutModal = document.getElementById('aboutModal')
const closeModalBtn = document.getElementById('closeModalBtn')

// ---- About modal ----------------------------------------------------------
// 'Escape' must close the load-bearing About panel first, then fall through.
aboutIconBtn.addEventListener('click', () => aboutModal.classList.remove('hidden'))
closeModalBtn.addEventListener('click', () => aboutModal.classList.add('hidden'))
aboutModal.addEventListener('click', (e) => {
  if (e.target === aboutModal) aboutModal.classList.add('hidden')
})

// ---- File selection -------------------------------------------------------
dumpFileInput.addEventListener('change', () => {
  const f = dumpFileInput.files && dumpFileInput.files[0]
  if (f) {
    fileNameBox.textContent = f.name
    fileNameBox.classList.remove('hidden')
    buildBtn.disabled = false
  } else {
    fileNameBox.classList.add('hidden')
    fileNameBox.textContent = ''
    buildBtn.disabled = true
  }
})

// ---- Build flow -----------------------------------------------------------
buildBtn.addEventListener('click', () => {
  const f = dumpFileInput.files && dumpFileInput.files[0]
  if (!f) return

  const reader = new FileReader()
  reader.onload = () => {
    const bytes = new Uint8Array(reader.result)
    let binary = ''
    const CHUNK = 0x8000
    for (let i = 0; i < bytes.length; i += CHUNK) {
      binary += String.fromCharCode.apply(
        null,
        bytes.subarray(i, i + CHUNK),
      )
    }
    const dumpBase64 = btoa(binary)

    const params = { dumpBase64, dumpFileName: f.name }
    if (asOnDateInput.value) {
      const [y, m, d] = asOnDateInput.value.split('-')
      params.asOn = { y: Number(y), m: Number(m), d: Number(d) }
    }

    chrome.runtime.sendMessage({ type: 'BUILD_DASHBOARD', params })
    setBuildUI(true)
    setProgress(10, 'Converting...', '')
  }
  reader.readAsArrayBuffer(f)
})

function setBuildUI(isBuilding) {
  buildBtn.disabled = isBuilding || !(dumpFileInput.files && dumpFileInput.files[0])
  progressContainer.classList.toggle('hidden', !isBuilding || buildDoneShown())
}

function buildDoneShown() {
  return !resultsSection.classList.contains('hidden')
}

function setProgress(pct, statusText, err) {
  progressContainer.classList.remove('hidden')
  progressFill.style.width = pct + '%'
  progressPctText.textContent = pct + '%'
  progressStepText.textContent = statusText
  if (err) {
    alertBanner.textContent = '⚠️ ' + err
    alertBanner.classList.remove('hidden')
  }
}

// ---- Live progress --------------------------------------------------------
chrome.runtime.onMessage.addListener((message) => {
  if (message.type === 'DASHBOARD_DONE') {
    complete(message.info)
  } else if (message.type === 'DASHBOARD_FAILED') {
    setProgress(0, 'Build failed', message.error)
    chrome.storage.local.set({
      dashboardScanState: {
        isBuilding: false,
        pct: 0,
        statusText: 'Error: ' + message.error,
        isFailed: true,
      },
    })
  }
})

;(() => {
  // Restore state if the popup reopens mid-build or after a build.
  chrome.storage.local.get('dashboardScanState', ({ dashboardScanState }) => {
    if (!dashboardScanState) return
    if (dashboardScanState.isDone && dashboardScanState.info) {
      complete(dashboardScanState.info)
    } else if (dashboardScanState.isBuilding) {
      setBuildUI(true)
      setProgress(dashboardScanState.pct || 5, dashboardScanState.statusText)
    } else if (dashboardScanState.isFailed && dashboardScanState.statusText) {
      alertBanner.textContent = '⚠️ ' + dashboardScanState.statusText
      alertBanner.classList.remove('hidden')
    }
  })
})()

function complete(info) {
  setBuildUI(false)
  progressContainer.classList.remove('hidden')
  progressFill.style.width = '100%'
  progressPctText.textContent = '100%'
  progressStepText.textContent = 'Dashboard Downloaded!'
  resultsSection.classList.remove('hidden')
  autoDlBanner.classList.remove('hidden')
  kpiTotal.textContent = Number(info.total || 0).toLocaleString()
  kpiInProgress.textContent = Number(
    (info.counts && info.counts['In Progress']) || 0,
  ).toLocaleString()
  kpiMonths.textContent = info.months || 0

  // reset for the next run
  chrome.storage.local.set({
    dashboardScanState: { isBuilding: false, isDone: true, info, pct: 100 },
  })
}
