// Outlook orchestration retained inside the handoff extension.
async function ensureOutlookTab() {
  let tabs = await chrome.tabs.query({ url: OUTLOOK_TAB_URLS })
  let tab = tabs.find((t) => t.active) || tabs[0]
  if (!tab) {
    tab = await chrome.tabs.create({ url: 'https://outlook.office.com/mail/', active: true })
    for (let i = 0; i < 60; i++) {
      const t = await chrome.tabs.get(tab.id).catch(() => null)
      if (t && t.status === 'complete') break
      await sleep(500)
    }
  } else {
    await chrome.tabs.update(tab.id, { active: true })
    await chrome.windows.update(tab.windowId, { focused: true }).catch(() => {})
  }
  await sleep(400)
  return tab.id
}

async function openCompose(tabId) {
  const attempt = async () => {
    try {
      return await ocCmd(tabId, 'newMail', {}, 45000)
    } catch (e) {
      // Preserve the content script's full payload (dump + deeplinks).
      if (e && e.aio) return e.aio
      return { ok: false, error: String((e && e.message) || e) }
    }
  }

  const logProbe = async (nm, max) => {
    if (nm && nm.dump) {
      for (const line of String(nm.dump).split('\n').slice(0, max)) {
        await aioLog('OWA probe · ' + line, 'warn')
      }
    }
  }

  let nm = await attempt()

  // Log the live probe of this Outlook build so selector mismatches are
  // diagnosable directly from the run log.
  await logProbe(nm, 40)

  if (!nm.ok && Array.isArray(nm.deeplinks) && nm.deeplinks.length) {
    await aioLog('New-mail click failed — opening the compose via the deep-link…', 'warn')
    for (const url of nm.deeplinks) {
      await chrome.tabs.update(tabId, { url, active: true }).catch(() => {})
      await sleep(1500)
      for (let i = 0; i < 40; i++) {
        const t = await chrome.tabs.get(tabId).catch(() => null)
        if (t && t.status === 'complete') break
        await sleep(500)
      }
      nm = await attempt()
      if (nm.ok) break
      await logProbe(nm, 24)
    }
  }

  if (!nm.ok) {
    // Surface the final probe so the run log always shows what the page had.
    await logProbe(nm, 30)
    throw new Error('Could not open the New mail compose: ' + (nm.error || 'unknown'))
  }
  return nm
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

  const frameId = OPCAPS[op] ? await frameFor(tabId, OPCAPS[op]) : undefined
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
      // callers like openCompose can still act on them.
      const e = new Error(res.error)
      e.aio = res
      throw e
    }
    return res
  } finally {
    clearTimeout(timer)
  }
}

