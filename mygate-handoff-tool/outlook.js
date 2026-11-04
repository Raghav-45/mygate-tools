/**
 * MyGate Handoff Tool — Outlook Compose Automation (content script)
 *
 * Injected into Outlook Web tabs (outlook.office.com / outlook.office365.com /
 * outlook.live.com) on demand by the background worker. It operates the
 * compose window just like a human would:
 *
 *   ping        → verifies the script is alive
 *   replyAll    → replies to the open report in Sent Items; retries reuse
 *                 only the reply created for the same report batch.
 *   body        → inserts the greeting/table/signature HTML (preserving the
 *                 Outlook signature that Outlook already inserted)
 *   attach      → attaches files via hidden file input, synthetic paste,
 *                 then drag&drop — verifying each attachment chip appears
 *   status      → reports current compose state
 */

(function () {
  'use strict'

  if (window.__AIO_OUTLOOK_INJECTED) return
  window.__AIO_OUTLOOK_INJECTED = true
  const signatures = new WeakMap()

  function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms))
  }

  async function until(getter, opts) {
    const o = Object.assign({ tries: 40, delay: 250 }, opts)
    for (let i = 0; i < o.tries; i++) {
      try {
        const v = getter()
        if (v) return v
      } catch (e) {}
      await sleep(o.delay)
    }
    return null
  }

  function isVisible(el) {
    if (!el) return false
    const r = el.getBoundingClientRect()
    return r.width > 0 && r.height > 0
  }

  function clickableElements() {
    return Array.from(
      document.querySelectorAll('button, [role="button"], a[href], [role="menuitem"], [role="option"], [data-automation-id]'),
    )
  }

  function labelOf(el) {
    let l = ((el.getAttribute('aria-label') || '') + ' ' + (el.getAttribute('title') || '')).trim()
    if (!l) {
      // aria-labelledBY wins in real accessibility trees — resolve it.
      const lb = el.getAttribute('aria-labelledby')
      if (lb) {
        const src = document.getElementById(lb)
        if (src) l = (src.textContent || '').trim()
      }
    }
    return l
  }

  function dataId(el) {
    return (
      el.getAttribute('data-automation-id') ||
      el.getAttribute('data-testid') ||
      el.getAttribute('data-test-id') ||
      ''
    )
  }

  // Integer-token label search on aria-label/title (no text noise).
  function findByLabel(patterns, { dataIdRx } = {}) {
    const rx = patterns.map((p) => new RegExp(p, 'i'))
    const rxData = dataIdRx ? new RegExp(dataIdRx, 'i') : null
    for (const el of clickableElements()) {
      if (!isVisible(el)) continue
      const l = labelOf(el)
      if (l && rx.some((r) => r.test(l))) return el
      if (rxData) {
        const d = dataId(el)
        if (d && rxData.test(d)) return el
      }
    }
    return null
  }

  // Text-content CSS fallback when no aria label is present.
  function findByVisibleText(words) {
    const els = [...document.querySelectorAll('span,div,button')]
    for (const want of words) {
      const rx = new RegExp(want, 'i')
      for (const el of els) {
        if (el.children.length > 0) continue
        const t = (el.textContent || '').trim()
        if (t && rx.test(t) && isVisible(el)) {
          const btn = el.closest('button, [role="button"]')
          if (btn) return btn
        }
      }
    }
    return null
  }

  // ---------------------------------------------------------------------------
  // Live DOM diagnostics — surfaced in the run log on failures
  // ---------------------------------------------------------------------------

  function truncate(s, n) {
    s = String(s == null ? '' : s)
    return s.length > n ? s.slice(0, n) + '…' : s
  }

  function fullLabel(el) {
    const parts = [labelOf(el), 'daid=' + (dataId(el) || '-'), 'text=' + truncate((el.textContent || '').trim(), 24)]
    return truncate(parts.join(' | ').replace(/\| - \| text= \|?/g, '').replace(/\| +\|/g, '|'), 90)
  }

  function describeEditable(el) {
    const tag = (el.tagName || '').toLowerCase()
    return (
      `${tag}[contenteditable] label=${truncate(labelOf(el), 40) || '-'} role=${el.getAttribute('role') || '-'} data-testid=${dataId(el) || '-'} pos=${Math.round(el.getBoundingClientRect().top)}x${Math.round(el.getBoundingClientRect().left)}`
    )
  }

  function diagnostics() {
    try {
      const lines = []
      const btns = [...document.querySelectorAll('button, [role="button"]')].filter(isVisible)
      for (const b of btns) {
        const l = fullLabel(b)
        if (/new|compose|mail.*message|message.*mail|send|attach/i.test(l)) lines.push('btn: ' + l)
      }
      const subs = [...document.querySelectorAll('input, textarea')].filter((x) => isVisible(x)).slice(0, 20)
      for (const s of subs) {
        lines.push(
          `input: tag=${(s.tagName || '').toLowerCase()} label=${truncate(labelOf(s), 40) || '-'} ph=${truncate(s.getAttribute('placeholder') || '', 30) || '-'} testid=${dataId(s) || '-'} id=${(s.id || '').slice(0, 24) || '-'}`,
        )
      }
      const ceds = [...document.querySelectorAll('[contenteditable="true"], [contenteditable=""]')].filter((x) => isVisible(x))
      for (const c of ceds.slice(0, 10)) lines.push('editable: ' + describeEditable(c))
      if (!ceds.length) lines.push('editable: NONE in this frame')

      // Frames — the compose body may live in an iframe.
      for (const fr of [...document.querySelectorAll('iframe')].filter((x) => x.getBoundingClientRect().width > 50 && x.getBoundingClientRect().height > 50)) {
        const r = fr.getBoundingClientRect()
        let inner = 'cross-origin/closed'
        try {
          const ed = [...fr.contentDocument.querySelectorAll('[contenteditable="true"]')].length
          const inputs = [...fr.contentDocument.querySelectorAll('input, textarea')].length
          const subj = !!([...fr.contentDocument.querySelectorAll('input[placeholder*="subject" i]')].length)
          inner = `same-origin: editables=${ed} inputs=${inputs} subjectLike=${subj}`
        } catch (e) {}
        lines.push(`iframe: ${fr.src.slice(0, 80)} ${Math.round(r.width)}x${Math.round(r.height)} ${inner}`)
      }

      const fs = [...document.querySelectorAll('input[type="file"]')]
      for (const fi of fs.slice(0, 6)) {
        lines.push(`fileinput: accept=${fi.getAttribute('accept') || '-'} multiple=${!!fi.multiple} inCompose=${!!(composeRootEl() && composeRootEl().contains(fi))}`)
      }
      if (!fs.length) lines.push('fileinput: none in this frame')
      return lines.slice(0, 60).join('\n') || '(page has no visible buttons yet — Outlook may still be loading)'
    } catch (e) {
      return '(probe error: ' + String(e && e.message) + ')'
    }
  }

  // ---------------------------------------------------------------------------
  // Compose form discovery
  // ---------------------------------------------------------------------------

  function subjectInput() {
    const cands = [
      ...document.querySelectorAll(
        'input[aria-label*="subject" i], input[placeholder*="subject" i], textarea[aria-label*="subject" i], textarea[placeholder*="subject" i], [data-testid*="subject"] input, [data-testid*="subject"] textarea, [data-automation-id*="subject"] input, [data-automation-id*="subject"] textarea',
      ),
    ]
    return cands.find(isVisible) || null
  }

  function bodyEditable() {
    const precise = [...document.querySelectorAll('div[contenteditable="true"][aria-label="Message body"], div[contenteditable="true"][aria-label*="message" i][aria-label*="body" i]')].find(isVisible)
    if (precise) return precise
    const cands = [
      // Precise, preferred: the composer labels its editor "Message body".
      ...document.querySelectorAll(
        'div[contenteditable="true"][aria-label="Message body"], div[contenteditable="true"][aria-label*="message" i][aria-label*="body" i]',
      ),
      ...document.querySelectorAll(
        'div[contenteditable="true"][role="textbox"], div[contenteditable="true"]',
      ),
    ]
    const vis = cands.filter(isVisible).filter(el => !/^(to|cc|bcc)$/i.test(el.getAttribute('aria-label') || ''))
    let best = null
    let bestArea = 0
    for (const el of vis) {
      const r = el.getBoundingClientRect()
      const area = r.width * r.height
      if (area > bestArea) {
        best = el
        bestArea = area
      }
    }
    return best
  }

  function composeRootEl() {
    const body = bodyEditable() || subjectInput()
    if (!body) return null
    const subject = subjectInput()
    if (subject && subject !== body) {
      // Outlook can nest the body more than six levels below its header and
      // attachment strip. The nearest shared ancestor is the reply boundary.
      let root = body.parentElement
      while (root && root !== document.body) {
        if (root.contains(subject)) return root
        root = root.parentElement
      }
      return null
    }
    let root = body
    for (let i = 0; i < 6; i++) {
      const p = root.parentElement
      if (!p) break
      if (p.getBoundingClientRect().height >= document.documentElement.scrollHeight * 0.85) break
      root = p
    }
    return root
  }

  function composeIsOpen() {
    // The subject field is the most stable anchor of an open compose. The
    // body may live in a separate frame/editor surface in some builds, so
    // subject + (body OR a recipient entry) both qualify.
    return !!(subjectInput() && (bodyEditable() || recipientInput('to')))
  }

  function recipientInput(kind) {
    const K = kind === 'cc' ? 'cc' : 'to'
    const exact = K === 'to' ? 'To' : 'Cc'
    const nameRx = new RegExp(`(?:^|\\W)${K}(?:\\W|$|\\W*recipients|addresses|emails?|groups?|names?|persons?)`, 'i')

    // a) attribute-labelled editable fields (aria-label exact first)
    const sel =
      kind === 'cc'
        ? 'input[aria-label="Cc"], textarea[aria-label="Cc"], input[aria-label="CC"], [role="combobox"][aria-label="Cc"], [role="textbox"][aria-label="Cc"], [contenteditable="true"][aria-label="Cc"], input[aria-label*="cc" i], textarea[aria-label*="cc" i], input[placeholder*="cc" i]'
        : 'input[aria-label="To"], textarea[aria-label="To"], [role="combobox"][aria-label="To"], [role="textbox"][aria-label="To"], [contenteditable="true"][aria-label="To"], input[aria-label*="To " i], textarea[aria-label^="To" i], input[aria-label^="To" i], input[placeholder*="name, group, or email" i], input[placeholder*="email address" i]'
    const cands = [...document.querySelectorAll(sel)].filter(isVisible)
      .filter(el => kind !== 'cc' || !/\bbcc\b/i.test(`${labelOf(el)} ${el.getAttribute('placeholder') || ''}`))
    const exactCands = cands.filter(el => (el.getAttribute('aria-label') || '').trim().toLowerCase() === exact.toLowerCase())
    if (exactCands.length) return exactCands[exactCands.length - 1]
    if (cands.length) return cands[cands.length - 1]

    const inputs = [...document.querySelectorAll('input:not([type="hidden"]):not([type="file"]), textarea')].filter((el) => isVisible(el))

    // b) honoured labels: aria-label, aria-labelledby, title, placeholder, test ids
    for (const el of inputs) {
      const bits = [
        el.getAttribute('aria-label'),
        el.getAttribute('aria-labelledby') ? (document.getElementById(el.getAttribute('aria-labelledby')) || {}).textContent || '' : '',
        el.getAttribute('title'),
        el.getAttribute('placeholder'),
        el.getAttribute('data-automation-id'),
        el.getAttribute('data-testid'),
      ]
        .filter(Boolean)
        .join(' ')
      if (bits && nameRx.test(bits) && !/^(search|filter)/i.test(bits) && !/subject/i.test(bits)) return el
    }

    // c) structural: a small "To" / "Cc" caption on a row that owns the input
    //    (works no matter what aria names the build uses)
    const root = composeRootEl()
    if (root) {
      for (const el of inputs) {
        if (!root.contains(el)) continue
        let row = el.parentElement
        for (let i = 0; i < 4 && row && row !== root; i++) {
          const t = (row.textContent || '').trim()
          if (new RegExp(`^${K}\\b`, 'i').test(t)) return el
          row = row.parentElement
        }
      }
    }
    return null
  }

  // ---------------------------------------------------------------------------
  // Attachment engine (one upload submission per file and reply)
  // ---------------------------------------------------------------------------

  function base64ToUint8(base64) {
    const bin = atob(base64)
    const bytes = new Uint8Array(bin.length)
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i)
    return bytes
  }

  function makeFile(f) {
    const bytes = base64ToUint8(f.base64)
    const mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    return new File([bytes], f.name, { type: mime })
  }

  function fileNameKey(name) {
    return String(name).toLowerCase().replace(/\s*\(\d+\)(?=\.(?:xlsx?|csv)\b)/g, '').replace(/[^a-z0-9]/g, '')
  }

  function attachmentNames(el) {
    const values = [el.getAttribute('data-filename'), el.getAttribute('title'), el.getAttribute('aria-label')]
    for (const child of el.querySelectorAll('[data-filename], [title], [aria-label]')) {
      values.push(child.getAttribute('data-filename'), child.getAttribute('title'), child.getAttribute('aria-label'))
    }
    return values.filter(Boolean).map(value => value.match(/[^\r\n<>]*?\.(?:xlsx?|csv)\b/i)?.[0].trim())
      .filter(value => value && !/[\u2026]|\.{3}/.test(value))
  }

  // Attachment cards belong to the reply, never to its editable body or the
  // original messages beneath it. Prefer Outlook's actual role=option cards.
  function chipList() {
    const root = composeRootEl()
    if (!root) return []
    const eligible = el => isVisible(el) && !el.closest('[contenteditable="true"], [role="document"]') && attachmentNames(el).length
    let elements = [...root.querySelectorAll('[role="option"]')].filter(eligible)
    if (!elements.length) {
      elements = [...root.querySelectorAll('[data-filename], .attachment, [title*=".xls" i], [aria-label*=".xls" i]')].filter(eligible)
      elements = elements.map(el => el.closest('[data-filename], .attachment') || el)
    }
    return [...new Set(elements)].filter(el => !elements.some(other => other !== el && other.contains(el)))
      .map(el => {
        const names = attachmentNames(el).map(fileNameKey)
        return { el, names, key: names[0] }
      })
  }

  function matchesFile(chip, file) {
    const key = fileNameKey(file.name)
    return !!key && chip.names.some(name => name.includes(key))
  }

  function fileChips(file) { return chipList().filter(chip => matchesFile(chip, file)) }
  function hasUnknownAttachmentCards() {
    const root = composeRootEl()
    return !!root && [...root.querySelectorAll('[role="listbox"][aria-label*="attachment" i] [role="option"]')]
      .some(el => isVisible(el) && !attachmentNames(el).length)
  }
  function chipVerified(file) {
    const chips = fileChips(file)
    return chips.length === 1 && !chips[0].el.querySelector('[role="progressbar"]') &&
      !/\b(uploading|failed|error|retry)\b/i.test(chips[0].el.textContent + ' ' + (chips[0].el.getAttribute('aria-label') || ''))
  }

  async function removeChip(chip) {
    const controls = [...chip.el.querySelectorAll('button, [role="button"]')].filter(isVisible)
    let remove = controls.find(el => /\b(remove|delete|close)\b/i.test(labelOf(el) + ' ' + el.textContent))
    if (!remove) {
      const more = controls.find(el => el.getAttribute('aria-haspopup') || /more actions|attachment options/i.test(labelOf(el)))
      if (!more) return false
      more.click()
      remove = await until(() => [...document.querySelectorAll('[role="menuitem"]')].find(el =>
        isVisible(el) && /^(remove|delete)( attachment)?$/i.test((labelOf(el) || el.textContent).trim())), { tries: 20, delay: 100 })
    }
    if (!remove) return false
    // One click only: syntheticClick dispatches click and then calls click(),
    // which can open and immediately close Outlook's attachment action menu.
    remove.click()
    return !!await until(() => !chip.el.isConnected || !isVisible(chip.el), { tries: 30, delay: 100 })
  }

  async function cleanAttachmentCards(files) {
    const keep = new Set()
    for (const chip of chipList()) {
      const file = files.find(file => matchesFile(chip, file))
      const key = file && fileNameKey(file.name)
      if (!file || keep.has(key)) {
        if (!await removeChip(chip)) return false
      } else keep.add(key)
    }
    return true
  }

  function uploadAnchors() { return [subjectInput(), bodyEditable()].filter(Boolean) }
  function attemptedFiles() {
    const anchors = uploadAnchors()
    const draftKey = anchors.map(el => el.getAttribute('data-mygate-handoff-draft')).find(Boolean) || ''
    const names = new Set()
    for (const anchor of anchors) {
      try {
        const saved = JSON.parse(anchor.getAttribute('data-mygate-handoff-upload-attempts') || '{}')
        if (saved.draftKey === draftKey) for (const name of saved.names || []) names.add(name)
      } catch (error) {}
    }
    return { anchors, draftKey, names }
  }

  function recordUploadAttempts(state, files) {
    for (const file of files) state.names.add(fileNameKey(file.name))
    const value = JSON.stringify({ draftKey: state.draftKey, names: [...state.names] })
    for (const anchor of state.anchors) anchor.setAttribute('data-mygate-handoff-upload-attempts', value)
  }

  function usableFileInputs(count) {
    const root = composeRootEl()
    return [...document.querySelectorAll('input[type="file"]')]
      .filter(el => !el.disabled && (count === 1 || el.multiple) &&
        (!el.accept || /spreadsheet|excel|octet-stream|\.(xlsx?|csv)/i.test(el.accept)))
      .sort((a, b) => Number(!!root?.contains(b)) - Number(!!root?.contains(a)))
  }

  async function submitFilesOnce(files, attempts) {
    let inputs = usableFileInputs(files.length)
    if (!inputs.length) {
      const opener = findByLabel(['attach'], { dataIdRx: 'attach' }) || findByVisibleText(['^Attach file$'])
      if (opener) {
        opener.click()
        await until(() => usableFileInputs(files.length).length, { tries: 20, delay: 100 })
        inputs = usableFileInputs(files.length)
      }
    }
    // Assignment alone does not upload. Try another input only if assignment
    // throws before dispatch; once change is dispatched, never submit again.
    for (const input of inputs) {
      try {
        const transfer = new DataTransfer()
        for (const file of files) transfer.items.add(makeFile(file))
        input.files = transfer.files
      } catch (error) { continue }
      recordUploadAttempts(attempts, files)
      input.dispatchEvent(new Event('change', { bubbles: true }))
      return true
    }
    return false
  }

  let attaching = false
  async function attachFiles(files) {
    if (attaching) return { ok: false, error: 'Attachments are already being prepared. Wait for the current step to finish.' }
    attaching = true
    try {
      if (!composeRootEl() || !uploadAnchors().length) return { ok: false, error: 'The reply editor is unavailable.' }
      if (!await cleanAttachmentCards(files)) {
        return { ok: false, error: 'Remove duplicate or outdated attachments from this reply in Outlook, then retry. No new upload was started.' }
      }
      const already = new Set(files.filter(chipVerified).map(file => file.name))
      const attempts = attemptedFiles()
      const missing = files.filter(file => !fileChips(file).length && !attempts.names.has(fileNameKey(file.name)))
      // Existing unverified cards may be uploads still appearing. Even if a
      // page reload removed our attempt markers, never add another batch on
      // top of an incomplete or unidentified attachment strip.
      if (missing.length && !chipList().length && !hasUnknownAttachmentCards()) await submitFilesOnce(missing, attempts)
      // Give a slow upload time to complete. If Outlook does not expose its
      // filename, stop and let the user inspect it instead of trying more inputs,
      // drag/drop or another upload on retry.
      await until(() => files.every(chipVerified) && !hasUnknownAttachmentCards(), { tries: 120, delay: 250 })
      const cleaned = await cleanAttachmentCards(files)
      const results = files.map(file => ({ name: file.name, slot: file.slot, ok: chipVerified(file), already: already.has(file.name) }))
      const ok = cleaned && results.every(result => result.ok) && chipList().length === files.length && !hasUnknownAttachmentCards()
      return { ok, results, count: chipList().length,
        error: ok ? undefined : 'Some attachments could not be verified. Check this reply and manually attach any missing files from Downloads. Existing upload attempts will not be repeated.' }
    } finally { attaching = false }
  }

  // ---------------------------------------------------------------------------
  // Operations
  // ---------------------------------------------------------------------------

  const ops = {
    ping() {
      return { ok: true, url: location.href }
    },

    async replyAll(p = {}) {
      if (!p.draftKey) return { ok: false, error: 'The generated report batch could not be identified. Please retry step 2.' }
      const editor = bodyEditable()
      if (editor || composeIsOpen()) {
        const ownReply = [editor, subjectInput()].some(el =>
          el?.getAttribute('data-mygate-handoff-draft') === p.draftKey &&
          el.getAttribute('data-mygate-handoff-reply') === 'true')
        if (ownReply) return { ok: true, reused: true,
          sourceUrl: (editor || subjectInput()).getAttribute('data-mygate-handoff-reply-source') }
        return { ok: false, error: 'Another draft is open. Finish or close it, then open the previous report in Sent Items and retry step 2.' }
      }

      const source = new URL(location.href)
      if (!/\/mail\/(?:\d+\/)?sentitems\/id\/[^/]+/i.test(source.pathname)) {
        return { ok: false, error: 'Open your previous report email in Outlook Sent Items, then retry step 2 from that tab.' }
      }
      const normalizeSubject = value => String(value || '').trim().replace(/^(?:(?:re|fw|fwd):\s*)+/i, '').trim().toLowerCase()
      const expectedSubject = normalizeSubject(p.subject)
      const pane = document.querySelector('[role="main"][aria-label="Reading Pane"]') || document
      const heading = [...pane.querySelectorAll('[role="heading"], h1, h2, h3')].find(el =>
        isVisible(el) && normalizeSubject(el.textContent) === expectedSubject)
      if (!expectedSubject || !heading) {
        return { ok: false, error: 'The open sent email does not match the report subject. Open the correct report and retry step 2.' }
      }
      const button = findByLabel(['^reply all$', '^reply to all$']) || findByVisibleText(['^Reply all$', '^Reply to all$'])
      if (!button) return { ok: false, error: 'Reply all is unavailable on the open report.', dump: diagnostics() }
      button.click()
      const ready = await until(() => bodyEditable() && recipientInput('to'), { tries: 40, delay: 250 })
      if (!ready) return { ok: false, error: 'Outlook did not open the reply-all editor.', dump: diagnostics() }
      for (const el of [bodyEditable(), subjectInput()].filter(Boolean)) {
        el.setAttribute('data-mygate-handoff-draft', p.draftKey)
        el.setAttribute('data-mygate-handoff-reply', 'true')
        el.setAttribute('data-mygate-handoff-reply-source', source.href)
      }
      return { ok: true, reused: false, sourceUrl: source.href }
    },

    async body(p) {
      const el = await until(bodyEditable, { tries: 20 })
      if (!el) return { ok: false, error: 'Message body not found', dump: diagnostics() }
      el.focus()
      await sleep(180)
      if (!signatures.has(el)) {
        // Retain the native reply's signature and any quoted history exactly
        // once; never capture the previously inserted report on a retry.
        const priorBody = el.querySelector('[data-mygate-handoff-body]')
        const priorSignature = el.querySelector('[data-mygate-handoff-signature]')
        signatures.set(el, priorBody ? priorSignature?.innerHTML || '' : el.innerHTML)
      }
      const existing = signatures.get(el)
      const signatureContent = document.createElement('div')
      signatureContent.innerHTML = existing || ''
      const hasContent = !!(signatureContent.textContent.replace(/\s|\u00a0/g, '') || signatureContent.querySelector('img, svg, hr'))
      const keepSig = p.keepSignature !== false && hasContent
      const range = document.createRange()
      range.selectNodeContents(el)
      const selection = window.getSelection()
      selection.removeAllRanges()
      selection.addRange(range)
      const signature = keepSig ? existing : String(p.fallbackSignatureHtml || '')
      const html = '<div data-mygate-handoff-body="true">' + String(p.html) + '</div>' +
        (signature ? '<br>'.repeat(10) + '<div data-mygate-handoff-signature="true">' + signature + '</div>' : '')
      const done = document.execCommand('insertHTML', false, html)
      if (!done) el.innerHTML = html
      el.blur()
      return { ok: true, keptSignature: keepSig }
    },

    async attach(p) {
      return attachFiles(p.files || [])
    },

    async status() {
      return { ok: true, url: location.href, hasCompose: composeIsOpen() }
    },
  }

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (!message || typeof message.aio !== 'string') return
    const handler = ops[message.aio]
    if (!handler) {
      sendResponse({ ok: false, error: 'Unknown op ' + message.aio })
    } else {
      Promise.resolve(handler(message.payload || {})).then(
        (res) => sendResponse(res),
        (err) => sendResponse({ ok: false, error: String((err && err.message) || err) }),
      )
    }
    return true // async
  })
})()
