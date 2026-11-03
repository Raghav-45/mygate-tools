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
const params = { fromDate: '2024-01-01', toDate: '2024-01-02', requestDelayMs: 1000,
  slots: ['report', 'dump', 'analysis', 'summary'],
  to: ['test@example.invalid'], cc: [], subject: 'Test', includeTable: true }
const cloud = new ExcelJS.Workbook()
const sheet = cloud.addWorksheet('Export')
sheet.addRow(['DLF Independent Floors'])
sheet.addRow(['Helpdesk Report'])
sheet.addRow(['I.D', 'Created Date', 'Unused', 'Category', 'Sub Category', 'Unused 2', 'Flat', 'Unused 3', 'Subject', 'Status'])
sheet.addRow([123, '01-01-2024', '', 'Accounts Billing', 'Billing', '', 'Phase-1 A-101', '', 'Test complaint', 'New'])
const cloudBytes = await cloud.xlsx.writeBuffer()
const clone = value => structuredClone(value)
const delay = () => new Promise(resolve => setTimeout(resolve, 1))

function harness({ enableEmail = false } = {}) {
  const store = {}
  const downloads = []
  const network = []
  const outlook = []
  let router
  let startup
  let exportFilters
  const control = { noCloudReport: false, failDownload: false, badReport: false, failEmail: false, partialAttach: false }
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
    Buffer, atob, btoa, structuredClone, setTimeout: fn => setTimeout(fn, 0), clearTimeout, setInterval, clearInterval,
    fetch: async (url, options) => {
      if (url.startsWith('extension:///')) return { text: async () => fs.readFileSync(path.join(dir, url.slice(13)), 'utf8') }
      if (url.startsWith('https://synthetic.invalid/')) return { arrayBuffer: async () => cloudBytes }
      const body = JSON.parse(options.body)
      network.push(body)
      if (body.operationName === 'getDownloadReportList') return { ok: true, json: async () => ({ data: {
        getDownloadReportList: { dataResponse: { data: control.noCloudReport ? [] : [{ report_name: 'Helpdesk Report',
          status: 'Success', report_link: 'https://synthetic.invalid/dump.xlsx', download_filters: exportFilters }] } } } }) }
      const req = body.variables.requestData
      if (req.isDownload) {
        exportFilters = req.downloadFilters
        return { ok: true, json: async () => ({ data: { getAdminSrList: {} } }) }
      }
      if (control.badReport && req.conditions.some(c => c.name === 'category')) return { ok: false, status: 401 }
      const statuses = req.conditions.find(c => c.name === 'mygate_status').values
      const totalCount = statuses.length === 1 ? 3 : statuses.includes('closed') ? 10 : 7
      return { ok: true, json: async () => ({ data: { getAdminSrList: { dataResponse: { totalCount } } } }) }
    },
  })
  context.self = context
  context.globalThis = context
  context.importScripts = (...files) => files.forEach(file => vm.runInContext(fs.readFileSync(path.join(dir, file), 'utf8'), context, { filename: file }))
  context.importScripts('background.js')
  // Exercise the preserved email implementation separately from the shipped,
  // disabled release. Production behavior is tested below without this override.
  if (enableEmail) context.HandoffRelease = Object.freeze({ emailEnabled: true })
  vm.runInContext(`ensureOutlookTab = async () => 1; openCompose = async () => ({ok: true});`, context)
  context.ocCmd = async (_tab, op, payload) => {
    outlook.push({ op, payload })
    if (control.failEmail) throw new Error('Synthetic Outlook UI change')
    if (op === 'attach') return { ok: true, results: payload.files.map((f, i) => ({ name: f.name, ok: !control.partialAttach || i === 0 })) }
    return { ok: true }
  }
  return { store, downloads, network, outlook, control, context, startup: () => startup(),
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

const defaults = harness({ enableEmail: true })
const initial = await defaults.request({ type: 'HANDOFF_GET_STATE' })
assert.deepEqual(Array.from(initial.cfg.slots), ['report', 'dump', 'analysis'])
assert.equal(initial.cfg.requestDelayMs, 1000)
defaults.store.handoffCfg = { requestDelayMs: 2000 }
assert.equal((await defaults.request({ type: 'HANDOFF_GET_STATE' })).cfg.requestDelayMs, 1000)
let defaultState = await defaults.run(1, { ...params, slots: initial.cfg.slots, requestDelayMs: 1300 })
assert.equal(defaultState.phase, 'ready', defaultState.error)
assert.equal(defaults.downloads.length, 3)
assert.equal(defaults.store.handoffCfg.requestDelayMs, 1300)
assert.equal(defaultState.files.find(f => f.slot === 'summary').status, 'off')
assert.deepEqual(defaults.downloads.map(f => f.filename.split('_')[0]), ['Pending', 'MyGate', 'Mygate Complaints Dashboard as on 02.01.24.xlsx'])
console.log('PASS: default selection downloads Pending Tickets, Dump, Analysis in order; Summary is off and 1s default migrates')

const h = harness({ enableEmail: true })
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
assert.equal(h.outlook.find(c => c.op === 'attach').payload.files.length, 2)
assert.equal(h.outlook.some(c => c.op === 'send'), false)
console.log('PASS: Outlook failure preserves step 1; retry uses same batch without scans or downloads; draft is never sent')

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
  const one = harness({ enableEmail: true })
  const single = await one.run(1, { ...params, slots: selection })
  assert.equal(single.phase, 'ready', single.error)
  assert.deepEqual(Object.keys(one.store.handoffBatch.files), selection)
  assert.equal(one.downloads.length, 1)
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
console.log('PASS: individual tools can run alone; invalid dependencies/all-off are rejected; email uses only the completed batch')

const disabledEmail = harness()
await disabledEmail.run(1, { ...params, slots: ['report'] })
const readyBeforeDraft = clone(disabledEmail.store.handoffState)
const blocked = await disabledEmail.request({ type: 'HANDOFF_DRAFT', cfg: { ...params, emailEnabled: true } })
assert.equal(blocked.ok, false)
assert.match(blocked.error, /disabled/)
assert.equal(disabledEmail.outlook.length, 0)
assert.deepEqual(disabledEmail.store.handoffState, readyBeforeDraft)
console.log('PASS: shipped build blocks email requests even after generation and ignores saved email-enable flags')

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
    const chrome = { runtime: {
      sendMessage: async msg => {
        sent.push(msg);
        if (msg.type === 'HANDOFF_GET_STATE') await new Promise(r => setTimeout(r, 500));
        return {ok:true,cfg,state:{phase:sent.some(m=>m.type==='HANDOFF_GENERATE')?'ready':'idle',generationReady:sent.some(m=>m.type==='HANDOFF_GENERATE'),files:[{slot:'report',status:'downloaded'}],log:[]}};
      }, onMessage:{addListener(){}}
    }};
    let failure = null;
    function check(ok, text) { if (!ok) throw new Error(text); }
    function verify(fn) { try { fn(); } catch (e) { failure = e.message; } }
    window.addEventListener('load', () => {
      setTimeout(() => verify(() => {
        check(document.getElementById('generateBtn').disabled, 'Generate unlocked before settings loaded');
        for (const slot of ['report','dump','analysis','summary']) check(document.getElementById(slot+'Toggle').disabled, 'Selection editable before settings loaded');
        document.getElementById('generateBtn').click();
        check(!sent.some(m=>m.type==='HANDOFF_GENERATE'), 'Generation started while settings were pending');
      }), 20);
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
        check(document.getElementById('draftBtn').disabled, 'Email unlocked after generation');
        check(document.getElementById('emailFields').disabled, 'Email form is enabled');
        check(!document.getElementById('emailOverlay').classList.contains('hidden'), 'Email overlay hidden');
        document.getElementById('draftBtn').click();
        check(!sent.some(m=>m.type==='HANDOFF_DRAFT'), 'Disabled email sent a request');
      }), 1250);
      setTimeout(() => { document.body.dataset.testResult = failure || 'PASS'; }, 1500);
    });
  </script>`
  const html = fs.readFileSync(path.join(dir, 'popup.html'), 'utf8').replace('<script src="release-features.js"></script>', stub + '<script src="release-features.js"></script>')
  fs.writeFileSync(path.join(preview, 'popup.html'), html)
  const browser = spawnSync(chromePath, ['--headless', '--no-sandbox', '--disable-gpu', '--no-first-run',
    `--user-data-dir=${path.join(preview, 'profile')}`, '--virtual-time-budget=2500', '--dump-dom', `file://${path.join(preview, 'popup.html')}`], { encoding: 'utf8', timeout: 30000, maxBuffer: 1024 * 1024 })
  assert.equal(browser.status, 0, browser.stderr)
  assert.match(browser.stdout, /data-test-result="PASS"/, browser.stdout.match(/data-test-result="[^"]*"/)?.[0] || 'Popup test did not finish')
  console.log('PASS: real Chrome popup locks loading controls, immediately saves report-only selection, sends only report, and keeps email covered and disabled')
} else console.log('SKIP: popup browser regression requires Chrome or Chromium; worker checks passed')
console.log('All handoff integration checks passed.')
