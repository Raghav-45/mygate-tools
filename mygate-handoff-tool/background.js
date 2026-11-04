/** Two independent steps around the original, unmodified generators. */
importScripts('release-features.js', 'exceljs.min.js', 'constants.js', 'engine.js', 'email-template.js',
  'generators/dump.js', 'generators/report.js', 'generators/summary.js', 'generators/analysis.js',
  'outlook-bridge.js')

const SLOTS = ['report', 'dump', 'analysis', 'summary']
const DEFAULT_SLOTS = ['report', 'dump', 'analysis']
// Match the standalone tools' defaults; saved shared-slider values never override these.
const REQUEST_DELAYS_MS = Object.freeze({ report: 1000, dump: 2000, summary: 1500 })
const LABELS = { dump: 'Master Dump', report: 'Pending Report', summary: 'Summary Sheet', analysis: 'Analysis Dashboard' }
const OUTLOOK_TAB_URLS = ['*://outlook.office.com/*', '*://outlook.office365.com/*', '*://outlook.live.com/*',
  '*://outlook.cloud.microsoft/*', '*://outlook.m365.cloud.microsoft/*']
const BASE_STATE = { phase: 'idle', step: null, pct: 0, stepText: '', files: [], log: [], generationReady: false, email: null, error: null }
let active = null
let aborted = false
let currentGenerator = null

function sleep(ms) { return new Promise(resolve => setTimeout(resolve, ms)) }
function todayStr() {
  const d = new Date()
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}
function displayDate(s) { const [y, m, d] = s.split('-'); return `${Number(d)}-${Number(m)}-${y}` }
async function getState() { return (await chrome.storage.local.get('handoffState')).handoffState || structuredClone(BASE_STATE) }
async function aioSetState(patch) {
  const state = { ...await getState(), ...patch }
  await chrome.storage.local.set({ handoffState: state })
  chrome.runtime.sendMessage({ type: 'HANDOFF_UPDATE', state }).catch(() => {})
  return state
}
// The original Outlook bridge calls this logger.
async function aioLog(msg, kind = 'info') {
  const state = await getState()
  await aioSetState({ log: [...state.log, { msg, kind }].slice(-160) })
}
function checkAbort() { if (aborted) throw new Error('__ABORT__') }
function removeSpeedSettings(cfg) {
  delete cfg.requestDelayMs
  delete cfg.settingsVersion
  return cfg
}
function migrateRecipientSettings(cfg) {
  // Upgrade the original saved defaults when an existing installation updates.
  const legacyAddresses = {
    'pawar-pankat@dlf.in': 'pawar-pankaj@dlf.in',
    'ashok kumar': 'tech.if3@sgvgroup.in',
    'estate if3': 'estate.if3@sgvgroup.in',
    'kirshan chander': 'apm2.if3@sgvgroup.in',
    'anurag yadav': 'anurag.yadav@sgvgroup.in',
    'anurag sharma': 'anurag.sharma@sgvgroup.in',
  }
  const key = value => String(value).trim().toLowerCase()
  if (Array.isArray(cfg.to) && cfg.to.length === 2 &&
      cfg.to.some(value => key(value) === 'gulati-divya@dlf.in') &&
      cfg.to.some(value => key(value) === 'pawar-pankat@dlf.in')) {
    cfg.to = [...self.EmailTemplate.DEFAULTS.to]
  }
  for (const field of ['to', 'cc']) {
    if (Array.isArray(cfg[field])) cfg[field] = cfg[field].map(value => legacyAddresses[key(value)] || value)
  }
  return cfg
}
async function getCfg() {
  const { handoffCfg } = await chrome.storage.local.get('handoffCfg')
  return migrateRecipientSettings(removeSpeedSettings({ fromDate: '2024-01-01', toDate: todayStr(), slots: DEFAULT_SLOTS,
    subject: self.EmailTemplate.DEFAULTS.subject, to: self.EmailTemplate.DEFAULTS.to,
    cc: self.EmailTemplate.DEFAULTS.cc, includeTable: true, ...handoffCfg }))
}
function validateCfg(cfg) {
  removeSpeedSettings(cfg)
  const validDate = s => /^\d{4}-\d{2}-\d{2}$/.test(s) && new Date(`${s}T00:00:00`).toString() !== 'Invalid Date'
  if (!validDate(cfg.fromDate) || !validDate(cfg.toDate) || cfg.fromDate > cfg.toDate) throw new Error('Please select valid From and Till dates, in order.')
  if (cfg.toDate > todayStr()) throw new Error('Report Till cannot be in the future.')
  cfg.slots = SLOTS.filter(s => (Array.isArray(cfg.slots) ? cfg.slots : DEFAULT_SLOTS).includes(s))
  if (!cfg.slots.length) throw new Error('Turn on at least one report to generate.')
  if (cfg.slots.includes('analysis') && !cfg.slots.includes('dump')) throw new Error('Analysis Dashboard requires Master Dump. Turn on Master Dump too.')
}

async function runOriginal(slot, params, index, total) {
  let scan = {}
  let captured = null
  let failure = null
  // The original sources keep their own state and download their own workbook.
  // Intercept only those boundaries; all original queries and builders run intact.
  const adapter = {
    tabs: chrome.tabs, scripting: chrome.scripting, cookies: chrome.cookies,
    runtime: { onMessage: { addListener() {} }, lastError: null,
      sendMessage: async msg => { if (msg.error || msg.errorMessage) failure = msg.error || msg.errorMessage } },
    storage: { local: { set: async obj => {
      checkAbort()
      scan = Object.values(obj)[0]
      if (scan.error || scan.isFailed) failure = scan.error || scan.statusText
      await aioSetState({ pct: Math.round((index + (scan.pct || 0) / 100) / total * 100),
        stepText: `${LABELS[slot]} · ${scan.statusText || scan.stepText || 'Working…'}` })
    } } },
    downloads: { download: (options, callback) => {
      captured = { name: options.filename, base64: options.url.split(',')[1] }
      callback(1)
    } },
  }
  currentGenerator = self.HandoffGenerators[slot](adapter, () => {})
  try {
    await currentGenerator.run(params)
    checkAbort()
    if (failure || !scan.isDone || !captured?.base64) throw new Error(failure || scan.statusText || `${LABELS[slot]} did not finish.`)
    // The original dump permits partial exports. Never label an incomplete
    // multi-year batch as a successful handoff or feed it into analysis.
    if (slot === 'dump' && scan.chunks.some(c => !String(c.status).startsWith('Merged'))) {
      throw new Error('Some dump chunks timed out or could not be merged. Please generate again.')
    }
    return { ...captured, scan }
  } finally { currentGenerator = null }
}

async function downloadFile(file) {
  checkAbort()
  const id = await new Promise((resolve, reject) => chrome.downloads.download({
    url: 'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,' + file.base64,
    filename: file.name, saveAs: false,
  }, id => chrome.runtime.lastError || id == null
    ? reject(new Error(chrome.runtime.lastError?.message || 'Could not start download.')) : resolve(id)))
  // A download ID means "started", not "saved". Only unlock step 2 when
  // Chrome confirms all selected downloads completed successfully.
  for (let attempt = 0; attempt < 240; attempt++) {
    checkAbort()
    const [item] = await chrome.downloads.search({ id })
    if (item?.state === 'complete') return id
    if (!item || item.state === 'interrupted') throw new Error(`Download failed: ${item?.error || file.name}`)
    await sleep(500)
  }
  throw new Error(`Download is still pending: ${file.name}. Check Chrome Downloads.`)
}

async function generate(cfg) {
  validateCfg(cfg)
  await chrome.storage.local.remove('handoffBatch')
  await aioSetState({ ...structuredClone(BASE_STATE), phase: 'generating', step: 1,
    stepText: 'Starting MyGate scans…', files: SLOTS.map(slot => ({ slot, status: cfg.slots.includes(slot) ? 'waiting' : 'off' })) })
  const batch = { files: {}, slots: cfg.slots, counts: null, generatedAt: Date.now() }
  const errors = []
  for (const [index, slot] of cfg.slots.entries()) {
    checkAbort()
    const state = await getState()
    const files = state.files
    const entry = files.find(f => f.slot === slot)
    try {
      entry.status = 'generating'
      await aioSetState({ files, stepText: `Generating ${LABELS[slot]}…` })
      let params = { fromDate: cfg.fromDate, toDate: cfg.toDate, requestDelayMs: REQUEST_DELAYS_MS[slot] }
      if (slot === 'summary') {
        const [year, month] = cfg.toDate.split('-').map(Number)
        params = { ...params, year, month }
      }
      if (slot === 'analysis') {
        const dump = batch.files.dump
        if (!dump) throw new Error('Analysis needs a successfully generated dump from this run.')
        params = { dumpBase64: dump.base64, dumpFileName: dump.name }
      }
      const result = await runOriginal(slot, params, index, cfg.slots.length)
      if (slot === 'report') batch.counts = { rows: result.scan.results, summary: result.scan.summary,
        fromFmt: displayDate(cfg.fromDate), toFmt: displayDate(cfg.toDate) }
      const file = { name: result.name, base64: result.base64 }
      entry.name = file.name
      entry.status = 'downloading'
      await aioSetState({ files })
      const downloadId = await downloadFile(file)
      batch.files[slot] = file
      entry.status = 'downloaded'
      entry.downloadId = downloadId
      await aioSetState({ files })
      await aioLog(`✓ ${LABELS[slot]} downloaded: ${file.name}`, 'ok')
    } catch (err) {
      checkAbort()
      entry.status = 'failed'
      errors.push(`${LABELS[slot]}: ${err.message}`)
      await aioSetState({ files })
      await aioLog(errors.at(-1), 'err')
    }
  }
  if (errors.length) throw new Error(errors.join(' · '))
  checkAbort()
  // This is only the current generated batch for step 2, never an import or
  // saved-file generation mode. Cleared on the next run and browser startup.
  await chrome.storage.local.set({ handoffBatch: batch })
  await aioSetState({ phase: 'ready', generationReady: true, pct: 100,
    stepText: `All ${cfg.slots.length} selected files downloaded.` +
      (self.HandoffRelease.emailEnabled ? ' Step 2 is ready.' : ' Attach them to your email manually.'), error: null })
}

async function draft(cfg) {
  if (!self.HandoffRelease.emailEnabled) throw new Error('Email is disabled in this version. It will be available in a future update.')
  const state = await getState()
  const { handoffBatch: batch } = await chrome.storage.local.get('handoffBatch')
  const batchSlots = batch?.slots || Object.keys(batch?.files || {})
  if (!state.generationReady || !batchSlots.length || batchSlots.some(s => !batch.files[s]?.base64)) throw new Error('Complete step 1 before drafting the email.')
  if (!Array.isArray(cfg.to) || !cfg.to.length) throw new Error('Enter at least one To recipient.')
  await aioSetState({ phase: 'emailing', step: 2, pct: 5, error: null, email: null, stepText: 'Opening the Outlook reply…' })
  checkAbort()
  const tabId = await ensureOutlookTab()
  const command = async (op, payload, timeout) => {
    checkAbort()
    const result = await ocCmd(tabId, op, payload, timeout)
    checkAbort()
    if (!result?.ok) throw new Error(result?.error || `Outlook could not complete: ${op}`)
    return result
  }
  await command('ping', {}, 15000)
  // Only this batch's own reply can be retried. Keep Outlook's native reply
  // subject and conversation headers so the daily reports remain in one thread.
  const draftKey = JSON.stringify([batch.generatedAt, cfg.to, cfg.cc || [], cfg.subject])
  const reply = await openReply(tabId, { draftKey, subject: cfg.subject })
  await aioSetState({ pct: 25, stepText: 'Checking reply recipients…' })
  await command('recipients', { to: cfg.to, cc: cfg.cc || [] }, 180000)
  const et = self.EmailTemplate
  const c = batch.counts
  const tableHtml = cfg.includeTable !== false && c
    ? et.buildPendingTableHtml(c.rows, c.summary, c.fromFmt, c.toFmt) : null
  await command('body', { html: et.buildBodyHtml({ tableHtml }),
    fallbackSignatureHtml: et.fallbackSignatureHtml(), keepSignature: true })
  // Preserve the existing email format: dump + dashboard attachments,
  // with Pending Report counts in the body when that report was generated.
  const attachments = ['dump', 'analysis'].filter(slot => batch.files[slot]).map(slot => ({ slot, ...batch.files[slot] }))
  if (attachments.length) {
    await aioSetState({ pct: 60, stepText: 'Attaching generated files…' })
    const attached = await command('attach', { files: attachments }, 600000)
    if (attachments.some(f => !attached.results?.some(r => r.name === f.name && r.ok))) {
      throw new Error('Some files need manual attachment. The generated workbooks are in Downloads.')
    }
  }
  await aioLog('Outlook reply draft prepared. Review it in Outlook and send when ready.', 'ok')
  await aioSetState({ phase: 'done', pct: 100, email: { status: 'drafted', sourceUrl: reply.sourceUrl },
    stepText: 'Reply draft ready for review.', error: null })
}

function launch(step, cfg) {
  if (active) throw new Error('A step is already running.')
  active = step
  aborted = false
  // Regular extension API activity keeps long MyGate scans alive even if the
  // toolbar popup closes. Buffers never pass through the popup message channel.
  const keepAlive = setInterval(() => chrome.runtime.getPlatformInfo().catch(() => {}), 20000)
  const task = (async () => {
    try {
      await chrome.storage.local.set({ handoffCfg: cfg })
      if (step === 1) await generate(cfg)
      else await draft(cfg)
    } catch (err) {
      const stopped = aborted || err.message === '__ABORT__'
      await aioSetState({ phase: stopped ? 'aborted' : 'failed', step,
        error: stopped ? 'Step stopped.' : err.message,
        stepText: step === 2 ? 'Email step stopped. Step 1 downloads are still available.' : 'Report generation did not finish.' })
      await aioLog(stopped ? 'Step stopped.' : err.message, stopped ? 'warn' : 'err')
    } finally {
      clearInterval(keepAlive)
      active = null
    }
  })()
  return task
}

chrome.runtime.onMessage.addListener((message, sender, respond) => {
  if (!message?.type?.startsWith('HANDOFF_')) return
  ;(async () => {
    switch (message.type) {
      case 'HANDOFF_GET_STATE': {
        let state = await getState()
        if (!active && ['generating', 'emailing'].includes(state.phase)) {
          state = await aioSetState({ phase: 'failed', error: 'The previous step was interrupted. Please retry that step.' })
        }
        respond({ ok: true, state, cfg: await getCfg() })
        break
      }
      case 'HANDOFF_SAVE_CFG':
        await chrome.storage.local.set({ handoffCfg: removeSpeedSettings({ ...message.cfg }) })
        respond({ ok: true })
        break
      case 'HANDOFF_GENERATE':
      case 'HANDOFF_DRAFT': {
        if (message.type === 'HANDOFF_DRAFT' && !self.HandoffRelease.emailEnabled) throw new Error('Email is disabled in this version. It will be available in a future update.')
        if (active) throw new Error('A step is already running.')
        const cfg = { ...await getCfg(), ...message.cfg }
        if (active) throw new Error('A step is already running.')
        const step = message.type === 'HANDOFF_GENERATE' ? 1 : 2
        if (step === 1) {
          // Never substitute saved/default tools for a missing run selection.
          if (!Array.isArray(message.cfg?.slots)) throw new Error('Report selection was not loaded. Reopen the extension and select the reports to generate.')
          cfg.slots = [...message.cfg.slots]
          validateCfg(cfg)
        }
        launch(step, cfg)
        respond({ ok: true })
        break
      }
      case 'HANDOFF_ABORT':
        aborted = true
        currentGenerator?.abort()
        respond({ ok: true })
        break
      default: respond({ ok: false, error: 'Unknown action.' })
    }
  })().catch(err => respond({ ok: false, error: err.message }))
  return true
})

chrome.runtime.onStartup.addListener(async () => {
  await chrome.storage.local.remove('handoffBatch')
  await aioSetState({ ...structuredClone(BASE_STATE), stepText: 'Generate a fresh batch to start.' })
})
