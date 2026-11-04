// Outlook orchestration retained inside the handoff extension.
async function ensureOutlookTab() {
  const [tab] = await chrome.tabs.query({ url: OUTLOOK_TAB_URLS, active: true, currentWindow: true })
  if (!tab) throw new Error('Open your previous report email in Outlook Sent Items, then run step 2 from that tab.')
  return tab.id
}

async function openReply(tabId, payload) {
  let result
  try {
    result = await ocCmd(tabId, 'replyAll', payload, 45000)
  } catch (error) {
    result = error.aio || { ok: false, error: error.message }
  }
  if (!result?.ok) {
    for (const line of String(result?.dump || '').split('\n').filter(Boolean).slice(0, 24)) {
      await aioLog('OWA probe · ' + line, 'warn')
    }
    throw new Error(result?.error || 'Could not open Reply all on the report email.')
  }
  return result
}

const CAPS_PROBE_FN = () => ({
  subj: !!(
    document.querySelector('input[placeholder*="subject" i], input[aria-label*="subject" i], textarea[placeholder*="subject" i]')
  ),
  body: !!document.querySelector('[contenteditable="true"]'),
  to: !!document.querySelector('[aria-label="To"], input[placeholder*="email address" i]'),
  attach: !!(
    document.querySelector('[data-testid*="attach" i], [data-automation-id*="attach" i], input[type=file]') ||
    [...document.querySelectorAll('button, [role="button"]')].some((b) => /attach/i.test(b.getAttribute('aria-label') || ''))
  ),
})

// Which frame holds the elements each op needs? Returns a frameId or undefined.
const OPCAPS = {
  body: ['body'],
  subject: ['subj'],
  attach: ['attach', 'body'],
  recipients: ['to', 'subj'],
}

async function frameFor(tabId, wants) {
  try {
    const res = await chrome.scripting.executeScript({ target: { tabId, allFrames: true }, func: CAPS_PROBE_FN })
    let fallback = undefined
    for (const item of res || []) {
      const c = item.result || {}
      if (item.frameId === 0) fallback = 0
      for (const want of wants) {
        if (c[want]) return item.frameId
      }
    }
    return fallback
  } catch (e) {
    return undefined
  }
}

async function ocCmd(tabId, op, payload = {}, timeoutMs = 90000) {
  await chrome.scripting
    .executeScript({ target: { tabId, allFrames: true }, files: ['outlook.js'] })
    .catch(() => {})
  await sleep(100)

  // The report heading and Reply all controls belong to the main page. A
  // sidebar iframe must never win the response before the real reply opens.
  const frameId = op === 'replyAll' ? 0 : OPCAPS[op] ? await frameFor(tabId, OPCAPS[op]) : undefined
  const target = frameId != null ? { frameId } : {}

  let timer = null
  const timeout = new Promise((_, reject) => {
    timer = setTimeout(() => reject(new Error(`Outlook op "${op}" timed out`)), timeoutMs)
  })
  try {
    const res = await Promise.race([
      chrome.tabs.sendMessage(tabId, { aio: op, payload }, target).then(
        (r) => r,
        (err) => {
          throw new Error((err && err.message) || 'No Outlook script responded')
        },
      ),
      timeout,
    ])
    if (res && res.error) {
      // Attach the full payload (dump/deeplinks/results) to the Error so
      // the reply opener can surface them in the run log.
      const e = new Error(res.error)
      e.aio = res
      throw e
    }
    return res
  } finally {
    clearTimeout(timer)
  }
}

