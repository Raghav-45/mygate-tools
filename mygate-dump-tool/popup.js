/**
 * MyGate Dump Tool - Frontend Controller Logic
 * Handles date pickers, live chunk progress updates, and autonomous storage state sync.
 */

document.addEventListener('DOMContentLoaded', async () => {
  const fromDateInput = document.getElementById('fromDate')
  const toDateInput = document.getElementById('toDate')
  const settingsCollapse = document.getElementById('settingsCollapse')
  const settingsToggleBtn = document.getElementById('settingsToggleBtn')
  const requestDelayInput = document.getElementById('requestDelayInput')
  const delayValText = document.getElementById('delayValText')

  const generateBtn = document.getElementById('generateBtn')
  const abortBtn = document.getElementById('abortBtn')

  const alertBanner = document.getElementById('alertBanner')
  const progressContainer = document.getElementById('progressContainer')
  const progressStepText = document.getElementById('progressStepText')
  const progressPctText = document.getElementById('progressPctText')
  const progressFill = document.getElementById('progressFill')

  const resultsSection = document.getElementById('resultsSection')
  const kpiChunks = document.getElementById('kpiChunks')
  const kpiRows = document.getElementById('kpiRows')
  const kpiStatus = document.getElementById('kpiStatus')
  const tableBody = document.getElementById('tableBody')
  const autoDlBanner = document.getElementById('autoDlBanner')

  const aboutIconBtn = document.getElementById('aboutIconBtn')
  const aboutModal = document.getElementById('aboutModal')
  const closeModalBtn = document.getElementById('closeModalBtn')

  let isScanning = false

  // Set default multi-year dates
  fromDateInput.value = '2024-01-01'
  toDateInput.value = new Date().toISOString().split('T')[0]

  // Load saved config & active scan state
  const stored = await chrome.storage.local.get([
    'requestDelay',
    'dumpScanState',
  ])
  if (stored.requestDelay) {
    requestDelayInput.value = stored.requestDelay
    delayValText.textContent = `${stored.requestDelay.toFixed(1)}s`
  }

  const scan = stored.dumpScanState
  if (scan) {
    if (scan.chunks && scan.chunks.length) {
      renderChunksTable(scan.chunks, scan.totalRows)
      resultsSection.classList.remove('hidden')
    }
    if (scan.isScanning) {
      isScanning = true
      generateBtn.classList.add('hidden')
      abortBtn.classList.remove('hidden')
      progressContainer.classList.remove('hidden')
      progressStepText.textContent = scan.statusText || 'Exporting...'
      progressPctText.textContent = `${scan.pct || 0}%`
      progressFill.style.width = `${scan.pct || 0}%`
      kpiStatus.textContent = 'Active'
    } else if (scan.isDone) {
      if (autoDlBanner) autoDlBanner.classList.remove('hidden')
      kpiStatus.textContent = 'Completed'
    }
  }

  // Toggle Settings
  settingsToggleBtn.addEventListener('click', () => {
    settingsCollapse.classList.toggle('hidden')
  })

  requestDelayInput.addEventListener('input', (e) => {
    const val = Number(e.target.value)
    delayValText.textContent = `${val.toFixed(1)}s`
    chrome.storage.local.set({ requestDelay: val * 1000 })
  })

  function showAlert(msg, isError = true) {
    alertBanner.textContent = msg
    alertBanner.className = `alert-box ${isError ? 'error' : 'success'}`
    alertBanner.classList.remove('hidden')
  }

  generateBtn.addEventListener('click', () => {
    if (autoDlBanner) autoDlBanner.classList.add('hidden')
    const fromDate = fromDateInput.value
    const toDate = toDateInput.value

    if (!fromDate || !toDate) {
      showAlert('Please enter valid Dump From and Dump Till dates.')
      return
    }

    if (new Date(fromDate) > new Date(toDate)) {
      showAlert('Dump From date cannot be after Dump Till date.')
      return
    }

    alertBanner.classList.add('hidden')
    isScanning = true
    generateBtn.classList.add('hidden')
    abortBtn.classList.remove('hidden')
    progressContainer.classList.remove('hidden')
    resultsSection.classList.remove('hidden')

    progressStepText.textContent = 'Initiating multi-year slice...'
    progressPctText.textContent = '0%'
    progressFill.style.width = '0%'
    kpiStatus.textContent = 'Starting...'

    chrome.runtime.sendMessage({
      type: 'START_DUMP_EXPORT',
      params: {
        fromDate,
        toDate,
        requestDelayMs: Number(requestDelayInput.value) * 1000,
      },
    })
  })

  abortBtn.addEventListener('click', () => {
    chrome.runtime.sendMessage({ type: 'ABORT_DUMP_EXPORT' })
    abortBtn.disabled = true
    abortBtn.textContent = 'Stopping...'
  })

  chrome.runtime.onMessage.addListener((msg) => {
    if (msg.type === 'DUMP_PROGRESS_UPDATE') {
      progressStepText.textContent = msg.statusText
      progressPctText.textContent = `${msg.pct}%`
      progressFill.style.width = `${msg.pct}%`
      kpiStatus.textContent = 'Exporting...'
      if (msg.chunks) renderChunksTable(msg.chunks, msg.totalRows)
    } else if (msg.type === 'DUMP_FINISHED' || msg.type === 'DUMP_ABORTED') {
      isScanning = false
      generateBtn.classList.remove('hidden')
      abortBtn.classList.add('hidden')
      abortBtn.disabled = false
      abortBtn.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2" ry="2"></rect></svg> Stop Live Export`
      progressContainer.classList.add('hidden')

      if (msg.type === 'DUMP_FINISHED') {
        kpiStatus.textContent = 'Completed'
        if (autoDlBanner) autoDlBanner.classList.remove('hidden')
        if (msg.chunks) renderChunksTable(msg.chunks, msg.totalRows)
        showAlert('Master Dump exported and downloaded automatically!', false)
      } else {
        kpiStatus.textContent = 'Stopped'
        showAlert('Export stopped by user.', true)
      }
    }
  })

  function renderChunksTable(chunks, totalRows = 0) {
    const readyCount = chunks.filter(
      (c) => c.status && c.status.includes('Merged'),
    ).length
    kpiChunks.textContent = `${readyCount} / ${chunks.length}`
    kpiRows.textContent = totalRows.toLocaleString()

    let html = ''
    chunks.forEach((c) => {
      let statusStyle = 'color: #D97706; font-weight: 600;'
      if (c.status.includes('Merged'))
        statusStyle = 'color: #10B981; font-weight: 700;'
      else if (c.status.includes('Error') || c.status.includes('Failed'))
        statusStyle = 'color: #EF4444; font-weight: 700;'

      html += `
        <tr>
          <td style="font-weight: 600;">${c.rangeStr}</td>
          <td class="num">${c.rowsFound || '-'}</td>
          <td class="num" style="${statusStyle}">${c.status}</td>
        </tr>`
    })
    tableBody.innerHTML = html
  }

  // About Modal
  aboutIconBtn.addEventListener('click', () =>
    aboutModal.classList.remove('hidden'),
  )
  closeModalBtn.addEventListener('click', () =>
    aboutModal.classList.add('hidden'),
  )
  aboutModal.addEventListener('click', (e) => {
    if (e.target === aboutModal) aboutModal.classList.add('hidden')
  })
})
