/** Integration tests: real original generators, synthetic MyGate responses,
 * actual workbook bytes, and simulated Chrome/Outlook boundaries. No live accounts. */
import assert from 'node:assert/strict'
import fs from 'node:fs'
import path from 'node:path'
import vm from 'node:vm'
import os from 'node:os'
import { spawnSync } from 'node:child_process'
import { createRequire } from 'node:module'
import { fileURLToPath } from 'node:url'
const require = createRequire(import.meta.url)
const dir = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const root = path.dirname(dir)
const ExcelJS = require(path.join(dir, 'exceljs.min.js'))
const params = { fromDate: '2024-01-01', toDate: '2024-01-02',
  slots: ['report', 'dump', 'analysis', 'summary'],
  to: ['test@example.invalid'], cc: [], subject: 'Test', includeTable: true }
const cloud = new ExcelJS.Workbook()
const sheet = cloud.addWorksheet('Export')
sheet.addRow(['DLF Independent Floors'])
sheet.addRow(['Helpdesk Report'])
sheet.addRow(['I.D', 'Created Date', 'Unused', 'Category', 'Sub Category', 'Unused 2', 'Flat', 'Unused 3', 'Subject', 'Status'])
const testCategories = ['Accounts Billing', 'Construction Or Project Related', 'Design Related Issue',
  'Estate Infra Outer Area from the plot', 'FM Common Area Related Issue', 'IT WIFI Network', 'Products Appliances']
for (const [categoryIndex, category] of testCategories.entries()) {
  for (let i = 0; i < 7; i++) sheet.addRow([123 + categoryIndex * 7 + i, '01-01-2024', '', category, 'Test', '', 'Test flat', '', 'Test complaint', 'New'])
}
const cloudBytes = await cloud.xlsx.writeBuffer()
const clone = value => structuredClone(value)
const delay = () => new Promise(resolve => setTimeout(resolve, 1))

function harness({ enableEmail } = {}) {
  const store = {}
  const downloads = []
  const network = []
  const requestedDelays = []
  const outlook = []
  let router
  let startup
  let exportFilters
  let exportSequence = 0
  const control = { noCloudReport: false, failDownload: false, badReport: false, failEmail: false, partialAttach: false, reportOpen: 7 }
  const chrome = {
    storage: { local: {
      get: async key => ({ [key]: clone(store[key]) }),
      set: async obj => { Object.assign(store, clone(obj)) },
      remove: async key => { delete store[key] },
    } },
    runtime: { onMessage: { addListener: fn => { router = fn } }, onStartup: { addListener: fn => { startup = fn } },
      sendMessage: async () => {}, getPlatformInfo: async () => ({}), getURL: p => 'extension:///' + p },
    tabs: { query: async () => [{ id: 1 }], get: async () => ({ id: 1, status: 'complete' }) },
    scripting: { executeScript: async () => [{ result: 'synthetic-token-for-offline-test' }] },
    cookies: { getAll: async () => [] },
    downloads: {
      download: (file, cb) => { downloads.push(file); cb(downloads.length) },
      search: async ({ id }) => [{ id, state: control.failDownload ? 'interrupted' : 'complete', error: 'FILE_FAILED' }],
    },
  }
  const context = vm.createContext({ chrome, location: { href: 'extension:///background.js' }, console: { log() {}, warn() {}, error() {} }, ExcelJS,
    Uint8Array, ArrayBuffer, TextEncoder, TextDecoder, Blob, Response, DecompressionStream, CompressionStream,
    Buffer, atob, btoa, structuredClone,
    // Ignore zero-delay yields used by the local workbook engine.
    setTimeout: (fn, ms) => { if (ms > 0) requestedDelays.push(ms); return setTimeout(fn, 0) }, clearTimeout, setInterval, clearInterval,
    fetch: async (url, options) => {
      if (url.startsWith('extension:///')) return { text: async () => fs.readFileSync(path.join(dir, url.slice(13)), 'utf8') }
      if (url.startsWith('https://synthetic.invalid/')) return { ok: true, arrayBuffer: async () => cloudBytes }
      const body = JSON.parse(options.body)
      network.push(body)
      if (body.operationName === 'getDownloadReportList') return { ok: true, json: async () => ({ data: {
        getDownloadReportList: { dataResponse: { data: control.noCloudReport || !exportFilters ? [] : [{ report_name: 'Helpdesk Report',
          status: 'Success', report_link: `https://synthetic.invalid/dump${exportSequence}.xlsx`,
          download_request_time: new Date(Date.UTC(2026, 10, 4, 10, 0, exportSequence)).toISOString(), download_filters: exportFilters }] } } } }) }
      const req = body.variables.requestData
      if (req.isDownload) {
        exportFilters = req.downloadFilters
        exportSequence++
        return { ok: true, json: async () => ({ data: { getAdminSrList: {} } }) }
      }
      if (control.badReport && req.conditions.some(c => c.name === 'category')) return { ok: false, status: 401 }
      const statuses = req.conditions.find(c => c.name === 'mygate_status').values
      const totalCount = !req.conditions.some(c => c.name === 'category') ? 49
        : statuses.length === 1 ? 3 : statuses.includes('closed') ? 10 : control.reportOpen
      return { ok: true, json: async () => ({ data: { getAdminSrList: { dataResponse: { totalCount } } } }) }
    },
  })
  context.self = context
  context.globalThis = context
  context.importScripts = (...files) => files.forEach(file => vm.runInContext(fs.readFileSync(path.join(dir, file), 'utf8'), context, { filename: file }))
  context.importScripts('background.js')
  // Most checks exercise the shipped flag; the kill-switch check overrides it.
  if (typeof enableEmail === 'boolean') context.HandoffRelease = Object.freeze({ emailEnabled: enableEmail })
  vm.runInContext(`ensureOutlookTab = async () => 1;`, context)
  context.openReply = async (_tab, payload) => { outlook.push({ op: 'replyAll', payload }); return { ok: true, sourceUrl: 'https://outlook.cloud.microsoft/mail/sentitems/id/previous-report' } }
  context.ocCmd = async (_tab, op, payload) => {
    outlook.push({ op, payload })
    if (control.failEmail) throw new Error('Synthetic Outlook UI change')
    if (op === 'attach') return { ok: true, results: payload.files.map((f, i) => ({ name: f.name, ok: !control.partialAttach || i === 0 })) }
    return { ok: true }
  }
  return { store, downloads, network, requestedDelays, outlook, control, context, startup: () => startup(),
    request: message => new Promise(resolve => router(message, {}, resolve)),
    run: async (step, cfg = params) => {
      const response = await new Promise(resolve => router({ type: step === 1 ? 'HANDOFF_GENERATE' : 'HANDOFF_DRAFT', cfg }, {}, resolve))
      assert.equal(response.ok, true, response.error)
      for (let i = 0; i < 10000; i++) {
        if (vm.runInContext('active === null', context)) return store.handoffState
        await delay()
      }
      throw new Error('Test timed out')
    } }
}

// Embedded code must remain byte-for-byte intact, including all formatting.
for (const slot of ['dump', 'report', 'summary', 'analysis']) {
  assert.ok(fs.readFileSync(path.join(dir, `generators/${slot}.js`), 'utf8')
    .includes(fs.readFileSync(path.join(root, `mygate-${slot}-tool/background.js`), 'utf8')))
}
for (const f of ['engine.js', 'constants.js', 'popup.css']) assert.deepEqual(fs.readFileSync(path.join(dir, f)), fs.readFileSync(path.join(root, 'mygate-analysis-tool', f)))
console.log('PASS: original generator bodies, analysis assets, and shared design are unchanged')

// A reply must use the user's selected Outlook tab, never a newly opened email
// or another Outlook window. Even a failed reply cannot navigate to compose.
const replyBridgeCalls = []
const replyBridgeControl = { tab: { id: 9 }, fail: false }
const replyBridge = vm.createContext({
  OUTLOOK_TAB_URLS: ['*://outlook.cloud.microsoft/*'],
  chrome: {
    tabs: {
      query: async query => { replyBridgeCalls.push({ query }); return replyBridgeControl.tab ? [replyBridgeControl.tab] : [] },
      sendMessage: async (tabId, message, target) => {
        replyBridgeCalls.push({ tabId, message, target })
        return replyBridgeControl.fail
          ? { ok: false, error: 'No report is open', deeplinks: ['https://outlook.cloud.microsoft/mail/deeplink/compose'] }
          : { ok: true, sourceUrl: 'https://outlook.cloud.microsoft/mail/sentitems/id/previous-report' }
      },
    },
    scripting: { executeScript: async () => [] },
  },
  sleep: async () => {}, aioLog: async () => {}, setTimeout, clearTimeout,
})
vm.runInContext(fs.readFileSync(path.join(dir, 'outlook-bridge.js'), 'utf8'), replyBridge)
assert.equal(await vm.runInContext('ensureOutlookTab()', replyBridge), 9)
assert.equal(replyBridgeCalls[0].query.active, true)
assert.equal(replyBridgeCalls[0].query.currentWindow, true)
await vm.runInContext('openReply(9, {draftKey:"batch-1",subject:"Mygate Complaint Summary."})', replyBridge)
assert.equal(replyBridgeCalls.at(-1).message.aio, 'replyAll')
assert.equal(replyBridgeCalls.at(-1).target.frameId, 0)
replyBridgeControl.tab = null
await assert.rejects(vm.runInContext('ensureOutlookTab()', replyBridge), /previous report email/)
replyBridgeControl.fail = true
await assert.rejects(vm.runInContext('openReply(9, {draftKey:"batch-1"})', replyBridge), /No report is open/)
assert.equal(replyBridgeCalls.filter(call => call.message).length, 2)
console.log('PASS: replies use the active Outlook tab and main frame; failures never open or navigate to New mail')

const defaults = harness()
assert.equal(defaults.context.HandoffRelease.emailEnabled, true)
const initial = await defaults.request({ type: 'HANDOFF_GET_STATE' })
assert.deepEqual(Array.from(initial.cfg.slots), ['report', 'dump', 'analysis'])
assert.equal(initial.cfg.requestDelayMs, undefined)
assert.equal(initial.cfg.to, undefined)
assert.equal(initial.cfg.cc, undefined)
defaults.store.handoffCfg = { to: ['legacy@example.invalid'], cc: ['Ashok Kumar'] }
const legacyRecipients = (await defaults.request({ type: 'HANDOFF_GET_STATE' })).cfg
assert.equal(legacyRecipients.to, undefined)
assert.equal(legacyRecipients.cc, undefined)
console.log('PASS: obsolete recipient settings are ignored; Outlook owns the reply recipients')
defaults.store.handoffCfg = { requestDelayMs: 5000, settingsVersion: 2 }
const migratedCfg = (await defaults.request({ type: 'HANDOFF_GET_STATE' })).cfg
assert.equal(migratedCfg.requestDelayMs, undefined)
assert.equal(migratedCfg.settingsVersion, undefined)
await defaults.request({ type: 'HANDOFF_SAVE_CFG', cfg: { ...params, requestDelayMs: 5000, settingsVersion: 2 } })
assert.equal(defaults.store.handoffCfg.requestDelayMs, undefined)
assert.equal(defaults.store.handoffCfg.settingsVersion, undefined)
let defaultState = await defaults.run(1, { ...params, slots: initial.cfg.slots, requestDelayMs: 1300 })
assert.equal(defaultState.phase, 'ready', defaultState.error)
assert.equal(defaults.downloads.length, 3)
assert.equal(defaults.store.handoffCfg.requestDelayMs, undefined)
assert.deepEqual([...new Set(defaults.requestedDelays)].sort((a, b) => a - b), [1000, 2000])
assert.equal(defaultState.files.find(f => f.slot === 'summary').status, 'off')
assert.deepEqual(defaults.downloads.map(f => f.filename.split('_')[0]), ['Pending', 'MyGate', 'Mygate Complaints Dashboard as on 02.01.24.xlsx'])
console.log('PASS: default selection downloads Pending Tickets, Dump, Analysis in order; Summary is off and obsolete shared delays are ignored')

const h = harness()
let state = await h.run(1)
assert.equal(state.phase, 'ready', JSON.stringify(state))
assert.equal(state.generationReady, true)
assert.equal(h.downloads.length, 4)
assert.equal(h.outlook.length, 0)
assert.equal(h.network.filter(r => r.variables.requestData?.isDownload).length, 1)
assert.equal(state.files.every(f => f.status === 'downloaded'), true)
assert.deepEqual(h.downloads.map(f => f.filename), [
  'Pending_Mygate_Tickets_Report_02-01-2024.xlsx',
  'MyGate_Master_Helpdesk_Dump_2024-01-01_to_2024-01-02.xlsx',
  'Mygate Complaints Dashboard as on 02.01.24.xlsx',
  'Complaint Summary Sheet - Jan-2024.xlsx',
])
const batch = h.store.handoffBatch
for (const slot of ['dump', 'report', 'summary', 'analysis']) assert.ok(batch.files[slot].base64.length > 100)
const report = new ExcelJS.Workbook()
await report.xlsx.load(Buffer.from(batch.files.report.base64, 'base64'))
assert.equal(report.worksheets[0].getCell('B10').value, 70)
assert.equal(report.worksheets[0].getCell('D10').value, 49)
const dump = new ExcelJS.Workbook()
await dump.xlsx.load(Buffer.from(batch.files.dump.base64, 'base64'))
assert.equal(dump.worksheets[0].getCell('B4').value, 123)
const analysis = new ExcelJS.Workbook()
await analysis.xlsx.load(Buffer.from(batch.files.analysis.base64, 'base64'))
assert.ok(analysis.getWorksheet('Data'))
assert.equal(analysis.getWorksheet('Data').getCell('B2').value, 123)
console.log('PASS: one click generates four actual workbooks, analysis contains the new dump ticket, no Outlook access')

const networkCount = h.network.length
h.control.failEmail = true
state = await h.run(2)
assert.equal(state.phase, 'failed')
assert.equal(state.generationReady, true)
assert.deepEqual(h.store.handoffBatch, batch)
assert.equal(h.downloads.length, 4)
h.control.failEmail = false
state = await h.run(2)
assert.equal(state.phase, 'done', state.error)
assert.equal(h.network.length, networkCount)
assert.equal(h.downloads.length, 4)
assert.ok(h.outlook.find(c => c.op === 'body').payload.html.includes('Pending Mygate Tickets'))
assert.equal(h.outlook.find(c => c.op === 'replyAll').payload.draftKey, JSON.stringify([batch.generatedAt, params.subject]))
assert.equal(h.outlook.find(c => c.op === 'replyAll').payload.subject, params.subject)
assert.ok(!h.outlook.some(c => ['newMail', 'recipients', 'subject', 'send'].includes(c.op)), 'Reply drafting must retain native recipients, thread and subject and never send')
assert.equal(state.email.sourceUrl, 'https://outlook.cloud.microsoft/mail/sentitems/id/previous-report')
assert.ok(h.outlook.find(c => c.op === 'body').payload.fallbackSignatureHtml.includes('Help desk'))
assert.equal(h.outlook.find(c => c.op === 'attach').payload.files.length, 2)
assert.equal(h.outlook.some(c => c.op === 'send'), false)
console.log('PASS: Outlook failure preserves step 1; retry uses same batch without scans or downloads; draft is never sent')
state = await h.run(2, { ...params, to: [], cc: [] })
assert.equal(state.phase, 'done', state.error)
assert.equal(h.outlook.filter(c => c.op === 'replyAll').at(-1).payload.draftKey, JSON.stringify([batch.generatedAt, params.subject]))
assert.equal(h.outlook.some(c => c.op === 'recipients'), false)
console.log('PASS: reply recipients are inherited untouched; empty or changed legacy recipient settings cannot change retry ownership')

h.control.partialAttach = true
state = await h.run(2)
assert.equal(state.phase, 'failed')
assert.equal(state.generationReady, true)
console.log('PASS: incomplete attachments are reported as email failure, with downloads preserved')

// A different selection for the next run must not change this batch's draft.
defaults.network.length = 0
defaultState = await defaults.run(2, { ...params, slots: ['summary'] })
assert.equal(defaultState.phase, 'done', defaultState.error)
assert.equal(defaults.outlook.find(c => c.op === 'attach').payload.files.length, 2)
assert.equal(defaults.network.length, 0)

for (const selection of [['report'], ['dump'], ['summary']]) {
  const one = harness()
  const single = await one.run(1, { ...params, slots: selection, requestDelayMs: 200 })
  assert.equal(single.phase, 'ready', single.error)
  assert.deepEqual(Object.keys(one.store.handoffBatch.files), selection)
  assert.equal(one.downloads.length, 1)
  const expectedDelay = { report: 1000, dump: 2000, summary: 1500 }[selection[0]]
  assert.ok(one.requestedDelays.length > 0)
  assert.ok(one.requestedDelays.every(ms => ms === expectedDelay), `${selection[0]} did not use its original delay`)
  if (selection.includes('report')) {
    assert.equal(one.network.some(r => r.operationName === 'getDownloadReportList' || r.variables.requestData.isDownload), false)
    assert.equal(one.downloads[0].filename, 'Pending_Mygate_Tickets_Report_02-01-2024.xlsx')
  }
  const drafted = await one.run(2, { ...params, slots: selection })
  assert.equal(drafted.phase, 'done', drafted.error)
  const attach = one.outlook.find(c => c.op === 'attach')
  assert.equal(attach?.payload.files.length || 0, selection.includes('dump') ? 1 : 0)
  const body = one.outlook.find(c => c.op === 'body').payload.html
  assert.equal(body.includes('Pending Mygate Tickets'), selection.includes('report'))
}
const invalid = harness()
for (const selection of [[], ['analysis'], ['report', 'analysis']]) {
  const response = await invalid.request({ type: 'HANDOFF_GENERATE', cfg: { ...params, slots: selection } })
  assert.equal(response.ok, false)
}
assert.equal(invalid.downloads.length, 0)
assert.equal(invalid.network.length, 0)
const missingSelection = { ...params }
delete missingSelection.slots
assert.equal((await invalid.request({ type: 'HANDOFF_GENERATE', cfg: missingSelection })).ok, false)
assert.equal(invalid.network.length, 0)
console.log('PASS: individual tools use their original delays; invalid dependencies/all-off are rejected; email uses only the completed batch')

const disabledEmail = harness({ enableEmail: false })
await disabledEmail.run(1, { ...params, slots: ['report'] })
const readyBeforeDraft = clone(disabledEmail.store.handoffState)
const blocked = await disabledEmail.request({ type: 'HANDOFF_DRAFT', cfg: { ...params, emailEnabled: true } })
assert.equal(blocked.ok, false)
assert.match(blocked.error, /disabled/)
assert.equal(disabledEmail.outlook.length, 0)
assert.deepEqual(disabledEmail.store.handoffState, readyBeforeDraft)
console.log('PASS: the release kill switch blocks email requests and ignores saved email-enable flags')

// New generation must never fall back to the previous batch after a failure.
h.control.badReport = true
state = await h.run(1)
assert.equal(state.phase, 'failed')
assert.equal(state.generationReady, false)
assert.equal(h.store.handoffBatch, undefined)
assert.equal(state.files.find(f => f.slot === 'report').status, 'failed')
assert.equal(state.files.find(f => f.slot === 'summary').status, 'downloaded')
console.log('PASS: report failure leaves other successful downloads available and never reuses stale files')

const timeout = harness()
timeout.control.noCloudReport = true
state = await timeout.run(1)
assert.equal(state.phase, 'failed')
assert.equal(state.files.find(f => f.slot === 'analysis').status, 'failed')
assert.equal(timeout.downloads.length, 2)
console.log('PASS: incomplete dump blocks dependent analysis, while Report and Summary still download')

const mismatchedCounts = harness()
mismatchedCounts.control.reportOpen = 8
state = await mismatchedCounts.run(1)
assert.equal(state.phase, 'failed')
assert.equal(state.generationReady, false)
assert.equal(mismatchedCounts.store.handoffBatch, undefined)
assert.deepEqual(mismatchedCounts.downloads.map(f => f.filename.split('_')[0]), ['Pending', 'Complaint Summary Sheet - Jan-2024.xlsx'])
assert.match(state.log.map(log => log.msg).join(' '), /counts differ|count/i)
assert.equal(mismatchedCounts.network.filter(r => r.variables.requestData?.isDownload).length, 1)
console.log('PASS: inconsistent report and dump counts block dump, analysis and email without another export')

const interrupted = harness()
interrupted.control.failDownload = true
state = await interrupted.run(1)
assert.equal(state.generationReady, false)
assert.equal(state.phase, 'failed')
assert.equal(interrupted.store.handoffBatch, undefined)
console.log('PASS: interrupted downloads cannot unlock step 2')

const abort = harness()
const first = await abort.request({ type: 'HANDOFF_GENERATE', cfg: params })
assert.equal(first.ok, true)
const duplicate = await abort.request({ type: 'HANDOFF_GENERATE', cfg: params })
assert.equal(duplicate.ok, false)
await abort.request({ type: 'HANDOFF_ABORT' })
for (let i = 0; i < 1000 && vm.runInContext('active !== null', abort.context); i++) await delay()
assert.equal(abort.store.handoffState.phase, 'aborted')
assert.equal(abort.store.handoffState.generationReady, false)
console.log('PASS: concurrent clicks are rejected and stopping generation leaves a retryable state')

await h.startup()
assert.equal(h.store.handoffBatch, undefined)
assert.equal(h.store.handoffState.generationReady, false)
console.log('PASS: browser restart clears the temporary batch, preserving no saved-file mode')

// Reproduce the popup race in an actual browser. A delayed initial response used
// to overwrite edits made while loading, causing an unintended Dump export.
const chromePath = ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/usr/bin/google-chrome', '/usr/bin/chromium'].find(p => fs.existsSync(p))
if (chromePath) {
  const preview = fs.mkdtempSync(path.join(os.tmpdir(), 'handoff-popup-regression-'))
  for (const file of ['popup.js', 'popup.css', 'handoff.css', 'release-features.js']) fs.copyFileSync(path.join(dir, file), path.join(preview, file))
  const stub = `<script>
    const sent = [];
    const cfg = ${JSON.stringify(params)};
    cfg.slots = ['report', 'dump', 'analysis'];
    let releaseInitialSettings;
    const initialSettingsPending = new Promise(resolve => { releaseInitialSettings = resolve; });
    let initialSettingsReturned = false;
    const chrome = { runtime: {
      sendMessage: async msg => {
        sent.push(msg);
        if (msg.type === 'HANDOFF_GET_STATE' && !initialSettingsReturned) {
          await initialSettingsPending;
          initialSettingsReturned = true;
        }
        return {ok:true,cfg,state:{phase:sent.some(m=>m.type==='HANDOFF_GENERATE')?'ready':'idle',generationReady:sent.some(m=>m.type==='HANDOFF_GENERATE'),files:[{slot:'report',status:'downloaded'}],log:[]}};
      }, onMessage:{addListener(){}}
    }};
    let failure = null;
    function check(ok, text) { if (!ok) throw new Error(text); }
    function verify(fn) { try { fn(); } catch (e) { failure = e.message; } }
    window.addEventListener('load', () => {
      setTimeout(() => { verify(() => {
        check(document.getElementById('generateBtn').disabled, 'Generate unlocked before settings loaded');
        check(document.getElementById('draftBtn').disabled, 'Email unlocked before generation');
        for (const slot of ['report','dump','analysis','summary']) check(document.getElementById(slot+'Toggle').disabled, 'Selection editable before settings loaded');
        document.getElementById('generateBtn').click();
        check(!sent.some(m=>m.type==='HANDOFF_GENERATE'), 'Generation started while settings were pending');
      }); releaseInitialSettings(); }, 20);
      setTimeout(() => verify(() => {
        const dump = document.getElementById('dumpToggle');
        check(!dump.disabled, 'Selection stayed locked after loading');
        dump.checked = false; dump.dispatchEvent(new Event('input'));
        check(!document.getElementById('analysisToggle').checked, 'Analysis dependency not cleared');
        check(sent.some(m=>m.type==='HANDOFF_SAVE_CFG' && m.cfg.slots.join(',')==='report'), 'Selection not saved immediately');
        document.getElementById('generateBtn').click();
      }), 600);
      setTimeout(() => verify(() => {
        check(sent.find(m=>m.type==='HANDOFF_GENERATE')?.cfg.slots.join(',')==='report', 'Run did not use report-only selection');
        check(!document.getElementById('requestDelayInput'), 'Shared speed slider is still visible');
        check(!('requestDelayMs' in sent.find(m=>m.type==='HANDOFF_GENERATE').cfg), 'Popup still supplies a shared delay');
        check(!document.getElementById('draftBtn').disabled, 'Email stayed locked after generation');
        check(!document.getElementById('emailFields').disabled, 'Email form stayed disabled');
        check(document.getElementById('emailOverlay').classList.contains('hidden'), 'Email overlay is blocking the form');
        check(!document.getElementById('toInput') && !document.getElementById('ccInput'), 'Reply still offers recipient fields that can duplicate native recipients');
        document.getElementById('draftBtn').click();
        check(sent.some(m=>m.type==='HANDOFF_DRAFT' && m.cfg.slots.join(',')==='report'), 'Email button did not request a draft');
        check(!('to' in sent.find(m=>m.type==='HANDOFF_DRAFT').cfg) && !('cc' in sent.find(m=>m.type==='HANDOFF_DRAFT').cfg), 'Popup still passes recipients for re-entry');
      }), 1250);
      setTimeout(() => { document.body.dataset.testResult = failure || 'PASS'; }, 1500);
    });
  </script>`
  const html = fs.readFileSync(path.join(dir, 'popup.html'), 'utf8')
    .replace(/<link\b[^>]*\bhref=["']https?:\/\/[^>]*>/gi, '')
    .replace(/(<img\b[^>]*\bsrc=)["']https?:\/\/[^"']*["']/gi, '$1"data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"')
    .replace('<script src="release-features.js"></script>', stub + '<script src="release-features.js"></script>')
  fs.writeFileSync(path.join(preview, 'popup.html'), html)
  const browser = spawnSync(chromePath, ['--headless', '--no-sandbox', '--disable-gpu', '--no-first-run',
    `--user-data-dir=${path.join(preview, 'profile')}`, '--virtual-time-budget=2500', '--dump-dom', `file://${path.join(preview, 'popup.html')}`], { encoding: 'utf8', timeout: 30000, maxBuffer: 1024 * 1024 })
  assert.equal(browser.status, 0, browser.stderr)
  assert.match(browser.stdout, /data-test-result="PASS"/, browser.stdout.match(/data-test-result="[^"]*"/)?.[0] || 'Popup test did not finish')
  console.log('PASS: real Chrome popup locks loading controls, saves report-only selection, and enables email drafting after generation')
} else console.log('SKIP: popup browser regression requires Chrome or Chromium; worker checks passed')
await import('./outlook_dom_tests.mjs')
await import('./outlook_attachment_tests.mjs')
console.log('All handoff integration checks passed.')

async function runDumpRegressions() {
const source = fs.readFileSync(path.join(root, 'mygate-dump-tool/background.js'), 'utf8')
const time = '2026-11-04 03:46:44 PM'
const oldTime = '2026-11-04 03:27:34 PM'
const filters = (from = '1:1:2026', to = '4:11:2026') => ({ From: from, To: to, 'Date From': from, 'Date To': to,
  Status: ['New ', 'Reopened ', 'In Progress ', 'Job Done ', 'On Hold '] })
const report = (df = filters(), stamp = time, link = 'https://synthetic.invalid/current.xlsx') => ({
  report_name: 'Helpdesk Report', status: 'Success', report_link: link, download_request_time: stamp, download_filters: df,
})
async function workbook(rows, reordered = false) {
  const wb = new ExcelJS.Workbook()
  const ws = wb.addWorksheet('Export')
  ws.addRow(['Society']); ws.addRow(['Helpdesk Report'])
  const headers = ['I.D', 'Created Date', 'Unused', 'Category', 'Sub Category', 'Unused 2', 'Flat', 'Unused 3', 'Subject', 'Status']
  if (reordered) [headers[0], headers[3]] = [headers[3], headers[0]]
  ws.addRow(headers)
  for (const row of rows) {
    const values = [row.id, row.date, '', row.category || 'Accounts Billing', 'Billing', '', 'Test flat', '', 'Test complaint', row.status || 'New']
    if (reordered) [values[0], values[3]] = [values[3], values[0]]
    ws.addRow(values)
  }
  return wb.xlsx.writeBuffer()
}
const goodBytes = await workbook([{ id: 1, date: '04-11-26, 03:35 PM' }])
async function run(options = {}) {
  const network = [], downloads = [], files = [], delays = [], store = {}
  let exportNumber = 0, currentFilters, pollNumber = 0
  const old = report(filters(), oldTime, 'https://synthetic.invalid/old.xlsx')
  const chrome = {
    runtime: { onMessage: { addListener() {} }, sendMessage: async () => {} },
    tabs: { query: async () => [{ id: 1 }] },
    scripting: { executeScript: async () => [{ result: 'synthetic-offline-token' }] },
    storage: { local: { set: async value => Object.assign(store, structuredClone(value)) } },
    downloads: { download: (f, cb) => { downloads.push(f); cb(1) } },
  }
  const response = data => ({ ok: true, json: async () => ({ data }) })
  const context = vm.createContext({ chrome, ExcelJS, Date, URL, Uint8Array, ArrayBuffer, btoa, atob,
    importScripts() {}, console: { error() {} }, setTimeout: (fn, ms) => { delays.push(ms); fn() },
    fetch: async (url, init) => {
      if (url.startsWith('https://synthetic.invalid/')) {
        files.push({ url, cache: init?.cache })
        return { ok: !options.fileError, arrayBuffer: async () => options.bytesByYear?.[currentFilters.From.split(':')[2]] || options.bytes || goodBytes }
      }
      const body = JSON.parse(init.body)
      network.push(body)
      const req = body.variables.requestData
      if (body.operationName === 'getDownloadReportList') {
        if (options.listError) return { ok: false, json: async () => ({}) }
        let records = [old]
        if (exportNumber) {
          pollNumber++
          const current = report(currentFilters, `2026-11-04 03:46:${44 + exportNumber} PM`, options.reuseLink ? old.report_link : `https://synthetic.invalid/current${exportNumber}.xlsx`)
          if (options.wrongFilters) current.download_filters.Status = ['In Progress ', 'Reopened ']
          if (options.conflictingDates) current.download_filters['Date From'] = '1:1:2025'
          records = options.timeout ? [old] : options.oldFirst && pollNumber === 1 ? [old] : [old, current]
          if (options.sharedLink) records.push(report(filters('1:1:2025', '31:12:2025'), time, current.report_link))
        }
        return response({ getDownloadReportList: { success: true, dataResponse: { data: records } } })
      }
      if (req.isDownload) {
        exportNumber++; currentFilters = structuredClone(req.downloadFilters); pollNumber = 0
        return response({ getAdminSrList: { success: !options.exportRejected } })
      }
      return response({ getAdminSrList: { success: true, dataResponse: { totalCount: options.count ?? 1 } } })
    },
  })
  await vm.runInContext(source + '\nrunMasterDumpExport(' + JSON.stringify({ fromDate: options.fromDate || '2026-01-01', toDate: '2026-11-04', requestDelayMs: 2000,
    ...(options.expectedOpenByCategory ? { expectedOpenByCategory: options.expectedOpenByCategory } : {}) }) + ')', context)
  return { state: store.dumpScanState, network, downloads, files, delays, exports: exportNumber }
}

// Reproduce the real failure: a same-date historical success is visible while
// the new cloud request is pending. It must never be downloaded or merged.
let result = await run({ oldFirst: true })
assert.equal(result.state.isDone, true)
assert.equal(result.files.some(f => f.url.endsWith('/old.xlsx')), false, 'accepted an earlier cloud export')
assert.equal(result.exports, 1)
assert.equal(result.state.totalRows, 1)
assert.ok(result.files.every(f => f.cache === 'no-store'))
console.log('PASS: waits for the new export, ignores historical success, downloads once without cache')

for (const [name, options, reason] of [
  ['wrong year', { bytes: await workbook([{ id: 1, date: '31-12-25, 03:35 PM' }]) }, /outside|date range/i],
  ['duplicate IDs', { count: 2, bytes: await workbook([{ id: 1, date: '01-01-2026' }, { id: 1, date: '02-01-2026' }]) }, /duplicate/i],
  ['missing rows', { count: 824 }, /count|expected/i],
  ['closed tickets', { bytes: await workbook([{ id: 1, date: '01-01-2026', status: 'Closed' }]) }, /status/i],
  ['unreadable date', { bytes: await workbook([{ id: 1, date: 'not a date' }]) }, /date/i],
  ['category mismatch', { expectedOpenByCategory: [{ category: 'Accounts Billing', open: 2 }] }, /count|expected/i],
  ['expired file', { fileError: true }, /download/i],
  ['export rejected', { exportRejected: true }, /rejected|request/i],
  ['shared cloud URL', { sharedLink: true }, /reused|shared/i],
  ['old cloud URL reused', { reuseLink: true }, /reused|shared/i],
]) {
  result = await run(options)
  assert.equal(result.state.isDone, false, name)
  assert.equal(result.state.isScanning, false, name)
  assert.equal(result.state.isFailed, true, name)
  assert.match(result.state.statusText, reason, name)
  assert.equal(result.downloads.length, 0, name)
  assert.equal(result.exports, 1, `${name}: an export was resubmitted`)
  console.log(`PASS: ${name} stops without a partial download or another export request`)
}
for (const options of [{ timeout: true }, { wrongFilters: true }, { conflictingDates: true }]) {
  result = await run({ fromDate: '2024-01-01', ...options })
  assert.equal(result.state.isFailed, true)
  assert.equal(result.exports, 1, 'continued to another year after a failed chunk')
  assert.equal(result.downloads.length, 0)
  assert.equal(result.files.length, 0)
}
console.log('PASS: timeout or wrong filters stop the entire dump; no later years or retries')
result = await run({ listError: true })
assert.equal(result.exports, 0)
assert.equal(result.state.isFailed, true)
console.log('PASS: failed export-history preflight never starts a cloud export')
result = await run({ count: 0, bytes: await workbook([]) })
assert.equal(result.state.isDone, true)
assert.equal(result.state.totalRows, 0)
console.log('PASS: a verified empty export is valid')
result = await run({ bytes: await workbook([{ id: 1, date: '04-11-2026' }], true) })
assert.equal(result.state.isDone, true)
console.log('PASS: column mapping follows each downloaded workbook header')
result = await run({ fromDate: '2024-01-01', bytesByYear: {
  2026: await workbook([{ id: 1, date: '04-11-26, 03:35 PM' }]),
  2025: await workbook([{ id: 2, date: '31-12-25, 03:35 PM' }], true),
  2024: await workbook([{ id: 3, date: '29-02-2024' }]),
}, expectedOpenByCategory: [{ category: 'Accounts  Billing', open: 3 }] })
assert.equal(result.state.isDone, true, result.state.statusText)
assert.equal(result.exports, 3)
assert.equal(result.files.length, 3)
assert.equal(result.state.totalRows, 3)
assert.equal(result.downloads.length, 1)
console.log('PASS: three yearly exports merge once with unique IDs, leap day and per-file header mapping')
result = await run({ count: 2, bytes: await workbook([{ id: 1, date: '01-01-2026' }, { id: 2, date: '02-01-2026', category: 'Design Related Issue' }]),
  expectedOpenByCategory: [{ category: 'Accounts Billing', open: 2 }] })
assert.equal(result.state.isFailed, true)
assert.equal(result.downloads.length, 0)
console.log('PASS: category mismatch is rejected even when grand totals agree')
const dateContext = vm.createContext({ chrome: { runtime: { onMessage: { addListener() {} } } }, importScripts() {}, Date })
vm.runInContext(source, dateContext)
for (const [value, expected] of [['31-12-25, 03:35 PM', '2025-12-31'], ['29/02/2024', '2024-02-29'], ['2026-11-04T10:00:00Z', '2026-11-04'], ['31-02-2026', null]]) {
  assert.equal(vm.runInContext('createdDateKey(' + JSON.stringify(value) + ')', dateContext), expected)
}
assert.equal(vm.runInContext('createdDateKey(new Date(Date.UTC(2026, 10, 4)))', dateContext), '2026-11-04')
console.log('PASS: date validation supports MyGate text dates, Excel dates and invalid-calendar rejection')
console.log('All dump regressions passed (offline; zero MyGate traffic).')

}
await runDumpRegressions()
