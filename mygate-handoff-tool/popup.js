document.addEventListener('DOMContentLoaded', async () => {
  const $ = id => document.getElementById(id)
  const labels = { dump: 'Master Dump', report: 'Pending Report', summary: 'Summary Sheet', analysis: 'Analysis Dashboard' }
  const slots = ['report', 'dump', 'analysis', 'summary']
  const emailEnabled = self.HandoffRelease.emailEnabled
  $('emailCard').classList.toggle('email-disabled', !emailEnabled)
  $('emailOverlay').classList.toggle('hidden', emailEnabled)
  $('emailFields').disabled = !emailEnabled
  let saving
  let starting = false
  let initialized = false
  let state = null
  const call = async (type, data = {}) => {
    const result = await chrome.runtime.sendMessage({ type, ...data })
    if (!result?.ok) throw new Error(result?.error || 'Extension background is unavailable. Please try again.')
    return result
  }
  const selectedSlots = () => slots.filter(slot => $(`${slot}Toggle`).checked)
  function updateControls() {
    const busy = !initialized || starting || ['generating', 'emailing'].includes(state?.phase)
    const count = selectedSlots().length
    $('generateBtn').disabled = busy || !count
    $('generateBtn').textContent = `Generate & Download ${count} ${count === 1 ? 'File' : 'Files'}`
    for (const slot of slots) $(`${slot}Toggle`).disabled = busy
    $('selectionHint').textContent = count ? '' : 'Select at least one report.'
    $('selectionHint').classList.toggle('hidden', count > 0)
  }
  function collectCfg() {
    return { fromDate: $('fromDate').value, toDate: $('toDate').value,
      slots: selectedSlots(),
      subject: $('subjectInput').value.trim(),
      includeTable: $('includeTableChk').checked }
  }
  function alert(message, error = true) {
    $('alertBanner').textContent = message
    $('alertBanner').className = `alert-box ${error ? 'error' : 'success'}`
  }
  function render(next) {
    state = next
    const busy = !initialized || starting || ['generating', 'emailing'].includes(state.phase)
    updateControls()
    $('draftBtn').disabled = !emailEnabled || busy || !state.generationReady
    $('draftBtn').textContent = state.generationReady && state.step === 2 && ['failed', 'aborted'].includes(state.phase)
      ? 'Retry Reply Draft' : 'Prepare Reply Draft'
    $('progressContainer').classList.toggle('hidden', !busy)
    $('progressStepText').textContent = state.stepText || 'Starting…'
    const pct = Math.round(state.pct || 0)
    $('progressPctText').textContent = `${pct}%`
    $('progressFill').style.width = `${pct}%`
    const count = (state.files || []).filter(f => f.status === 'downloaded').length
    $('generationStatus').textContent = state.generationReady
      ? `✓ ${count} selected ${count === 1 ? 'file' : 'files'} downloaded. You can also email them manually from Downloads.`
      : state.phase === 'generating' ? 'Generating fresh reports from MyGate…'
      : state.step === 1 && ['failed', 'aborted'].includes(state.phase)
        ? 'Step 1 did not finish. Files already downloaded remain available; generate again to complete the selected files.'
        : 'Sign in to MyGate, then generate the reports.'
    $('emailStatus').textContent = !emailEnabled ? 'Email is disabled until a future update.'
      : state.phase === 'done' ? '✓ Reply draft ready. Review it in Outlook and send.'
      : state.phase === 'emailing' ? 'Preparing the Outlook reply…'
      : state.generationReady ? 'Open the previous report in Outlook Sent Items, then prepare the reply.'
      : 'Complete step 1 to enable email drafting.'
    if (state.error) alert((state.step === 2 ? 'Email step: ' : '') + state.error +
      (state.generationReady ? ' Your downloads are still available.' : ''))
    else if (state.phase === 'ready') alert(`Step 1 complete — ${count} selected ${count === 1 ? 'file' : 'files'} downloaded.`, false)
    else if (state.phase === 'done') alert('Step 2 complete — review the draft in Outlook.', false)
    else $('alertBanner').classList.add('hidden')
    $('filesSection').classList.toggle('hidden', !state.files?.length)
    $('filesBody').replaceChildren()
    for (const file of state.files || []) {
      const tr = document.createElement('tr')
      const name = document.createElement('td')
      name.textContent = labels[file.slot]
      if (file.name) {
        const detail = document.createElement('div')
        detail.className = 'file-name'
        detail.textContent = file.name
        name.append(detail)
      }
      const status = document.createElement('td')
      status.textContent = file.status
      tr.append(name, status)
      $('filesBody').append(tr)
    }
    $('logSection').classList.toggle('hidden', !state.log?.length)
    $('logBox').replaceChildren()
    for (const line of (state.log || []).slice(-60)) {
      const div = document.createElement('div')
      div.className = line.kind
      div.textContent = line.msg
      $('logBox').append(div)
    }
    $('logBox').scrollTop = $('logBox').scrollHeight
  }
  try {
    const initial = await call('HANDOFF_GET_STATE')
    const cfg = initial.cfg
    $('fromDate').value = cfg.fromDate
    $('toDate').value = cfg.toDate
    $('subjectInput').value = cfg.subject
    $('includeTableChk').checked = cfg.includeTable !== false
    for (const slot of slots) $(`${slot}Toggle`).checked = (cfg.slots || ['report', 'dump', 'analysis']).includes(slot)
    initialized = true
    render(initial.state)
  } catch (err) { alert(err.message) }
  for (const id of ['fromDate', 'toDate', 'subjectInput', 'includeTableChk', ...slots.map(s => `${s}Toggle`)]) {
    $(id).addEventListener('input', () => {
      if (!initialized) return
      if (id === 'dumpToggle' && !$('dumpToggle').checked) $('analysisToggle').checked = false
      if (id === 'analysisToggle' && $('analysisToggle').checked) $('dumpToggle').checked = true
      updateControls()
      clearTimeout(saving)
      const cfg = collectCfg()
      const save = () => call('HANDOFF_SAVE_CFG', { cfg }).catch(err => alert(err.message))
      if (id.endsWith('Toggle')) save()
      else saving = setTimeout(save, 400)
    })
  }
  async function start(type) {
    if (!initialized || starting) return
    if (type === 'HANDOFF_DRAFT' && !emailEnabled) return
    // Snapshot the visible selection before any await or control changes.
    const cfg = collectCfg()
    starting = true
    clearTimeout(saving)
    $('generateBtn').disabled = true
    for (const slot of slots) $(`${slot}Toggle`).disabled = true
    $('draftBtn').disabled = true
    $('abortBtn').disabled = false
    $('abortBtn').textContent = '◼ Stop current step'
    try {
      await call(type, { cfg })
      const latest = await call('HANDOFF_GET_STATE')
      starting = false
      render(latest.state)
    } catch (err) {
      starting = false
      if (state) render(state)
      alert(err.message)
    }
  }
  $('generateBtn').addEventListener('click', () => start('HANDOFF_GENERATE'))
  $('draftBtn').addEventListener('click', () => start('HANDOFF_DRAFT'))
  $('abortBtn').addEventListener('click', async () => {
    try {
      await call('HANDOFF_ABORT')
      $('abortBtn').disabled = true
      $('abortBtn').textContent = 'Stopping…'
    } catch (err) { alert(err.message) }
  })
  $('aboutIconBtn').addEventListener('click', () => $('aboutModal').classList.remove('hidden'))
  $('closeModalBtn').addEventListener('click', () => $('aboutModal').classList.add('hidden'))
  $('aboutModal').addEventListener('click', e => { if (e.target === $('aboutModal')) $('aboutModal').classList.add('hidden') })
  chrome.runtime.onMessage.addListener(msg => { if (msg.type === 'HANDOFF_UPDATE') render(msg.state) })
})
