/**
 * MyGate Handoff Tool — Outlook Compose Automation (content script)
 *
 * Injected into Outlook Web tabs (outlook.office.com / outlook.office365.com /
 * outlook.live.com) on demand by the background worker. It operates the
 * compose window just like a human would:
 *
 *   ping        → verifies the script is alive
 *   newMail     → opens a new compose: reuse → button → deep-link fallback.
 *                 On failure it returns a live DOM probe ("dump") of the
 *                 visible buttons so the run log shows exactly what the
 *                 current Outlook build calls its controls.
 *   recipients  → fills To/Cc with verification that each value committed
 *   subject     → fills the subject
 *   body        → inserts the greeting/table/signature HTML (preserving the
 *                 Outlook signature that Outlook already inserted)
 *   attach      → attaches files via hidden file input, synthetic paste,
 *                 then drag&drop — verifying each attachment chip appears
 *   send        → clicks Send (used only when auto-send is enabled)
 *   status      → reports current compose state
 */

(function () {
  'use strict'

  if (window.__AIO_OUTLOOK_INJECTED) return
  window.__AIO_OUTLOOK_INJECTED = true

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
        if (t && rx.test(t) && (isVisible(el) || true)) {
          const btn = el.closest('button, [role="button"]')
          if (btn) return btn
        }
      }
    }
    return null
  }

  function setNativeValue(input, value) {
    const proto =
      input instanceof HTMLTextAreaElement ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype
    const desc = Object.getOwnPropertyDescriptor(proto, 'value')
    const setter = desc && desc.set
    if (setter) setter.call(input, value)
    else input.value = value
    input.dispatchEvent(new Event('input', { bubbles: true }))
    input.dispatchEvent(new Event('change', { bubbles: true }))
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
    const cands = [
      // Precise, preferred: the composer labels its editor "Message body".
      ...document.querySelectorAll(
        'div[contenteditable="true"][aria-label="Message body"], div[contenteditable="true"][aria-label*="message" i][aria-label*="body" i]',
      ),
      ...document.querySelectorAll(
        'div[contenteditable="true"][role="textbox"], div[contenteditable="true"]',
      ),
    ]
    const vis = cands.filter(isVisible)
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

  function ensureCcRow() {
    if (recipientInput('cc')) return true
    // Cc expander: aria/title token match first.
    let btn = findByLabel(['\\bcc\\b', 'show cc', 'cc & bcc'])
    if (!btn) {
      btn = [...clickableElements()].find(
        (el) => isVisible(el) && /(^|\s)cc(\s|$|&)/i.test(labelOf(el)) && !/bcc/i.test(labelOf(el)),
      )
    }
    if (!btn) {
      // Some builds render a "Cc" text caption inside a button-like row.
      btn = findByVisibleText(['^Cc$', 'Cc'])
    }
    if (btn) {
      btn.click()
      return !!until(() => recipientInput('cc'), { tries: 12, delay: 200 })
    }
    return false
  }

  function fireKey(el, key, code) {
    const opts = { bubbles: true, cancelable: true, key, keyCode: code, which: code }
    el.dispatchEvent(new KeyboardEvent('keydown', opts))
    el.dispatchEvent(new KeyboardEvent('keyup', opts))
  }

  function recipientChipSeen(value) {
    const stem = escapeRegExp(value.split('@')[0].trim()).slice(0, 10)
    if (!stem) return false
    const root = composeRootEl() || document
    try {
      return new RegExp(stem, 'i').test(root.innerText || '')
    } catch (e) {
      return false
    }
  }

  function escapeRegExp(s) {
    return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  }

  function fieldHasText(el) {
    if (!el) return false
    const t = el.value != null ? el.value : el.textContent || ''
    return !!String(t).trim()
  }

  function clearField(el) {
    try {
      el.focus()
      if (el.setSelectionRange && el.value != null) el.setSelectionRange(0, (el.value || '').length)
      document.execCommand('selectAll', false, null)
      document.execCommand('delete')
    } catch (e) {}
  }

  function syntheticClick(el) {
    const base = { bubbles: true, cancelable: true, view: window, composed: true }
    try { el.dispatchEvent(new PointerEvent('pointerdown', { ...base, pointerType: 'mouse', isPrimary: true })) } catch (e) {}
    try { el.dispatchEvent(new MouseEvent('mousedown', base)) } catch (e) {}
    try { el.dispatchEvent(new PointerEvent('pointerup', { ...base, pointerType: 'mouse', isPrimary: true })) } catch (e) {}
    try { el.dispatchEvent(new MouseEvent('mouseup', base)) } catch (e) {}
    try { el.dispatchEvent(new MouseEvent('click', base)) } catch (e) {}
    try { el.click() } catch (e) {}
  }

  // Fluent pickers render the picker row for any typed-but-unknown address as
  // "Use this address: <value>". Clicking it is the only reliable commit on
  // this build (synthetic Enter/';' are ignored). Pick the smallest matching
  // row so huge ancestor containers never win.
  function suggestionItem(value) {
    const needle = String(value || '').trim().toLowerCase()
    const matchRx = (t) => /use this address/i.test(t) || (!!needle && t.includes(needle))
    const pick = (list) => {
      const rows = list
        .filter(isVisible)
        .filter((el) => matchRx(((el.textContent || '') + ' ' + (labelOf(el) || '')).toLowerCase()))
        .sort((a, b) => (a.textContent || '').length - (b.textContent || '').length)
      return rows[0] || null
    }
    const scope = composeRootEl() || document
    let hit = pick([...scope.querySelectorAll('[role="option"], [role="menuitem"], [data-automation-id*="suggestion" i], [class*="suggestion" i], li')])
    if (!hit) hit = pick([...document.querySelectorAll('[role="option"], [role="menuitem"], [class*="suggestion" i]')])
    return hit
  }

  async function typeRecipients(field, values) {
    const results = []
    for (const value of values) {
      const v = String(value || '').trim()
      if (!v) continue
      let el = recipientInput(field)
      if (!el && field === 'cc') {
        ensureCcRow()
        el = await until(() => recipientInput('cc'), { tries: 12, delay: 200 })
      }
      if (!el) {
        results.push({ value: v, ok: false })
        continue
      }

      // Fast path: this address is already present (committed chip or leftover
      // text OWA resolves on send) — nothing to do. Makes reuse runs instant.
      if (recipientChipSeen(v)) {
        results.push({ value: v, ok: true, already: true })
        continue
      }

      // The compose is often reused and still holds leftover text from an
      // earlier attempt — clear it completely before typing.
      el.focus()
      await sleep(120)
      clearField(el)
      await sleep(180)

      // Type the address (real typing first; native fallback for blocked editors).
      try {
        document.execCommand('insertText', false, v)
      } catch (e) {}
      await sleep(180)
      if (!fieldHasText(el)) {
        setNativeValue(el, v)
      }
      await sleep(450) // let the suggestion pane appear

      // Commit strategies until the picker empties the field (chip accepted).
      let committed = false
      for (const strategy of ['suggest', 'enter', 'semicolon', 'suggest2', 'blur', 'enter2']) {
        if (strategy === 'enter' || strategy === 'enter2') fireKey(el, 'Enter', 13)
        else if (strategy === 'semicolon') fireKey(el, ';', 186)
        else if (strategy === 'blur') {
          try {
            el.blur()
          } catch (e) {}
        } else {
          // "suggest" / "suggest2": click the "Use this address …" row.
          const item = suggestionItem(v)
          if (item) syntheticClick(item)
        }
        await sleep(450)
        // React swaps the input node when a chip lands — re-grab it so the
        // emptiness check reads the LIVE field.
        el = recipientInput(field) || el
        committed = !fieldHasText(el)
        if (committed) break
        try {
          el.focus()
        } catch (e) {}
        await sleep(100)
      }
      results.push({ value: v, ok: committed })
    }
    return results
  }

  // ---------------------------------------------------------------------------
  // Attachment engine (files via hidden input / DataTransfer paste / drop)
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

  function chipStem(name) {
    const stem = name.replace(/\.xlsx?$/i, '').toLowerCase()
    return stem.replace(/\s*\(\d+\)$/, '')
  }

  // Elements of the real attachment strip inside the compose. Never reads the
  // whole compose text — the subject/body can contain similar words and used
  // to fake attachments.
  function attachAreaEls() {
    const root = composeRootEl()
    if (!root) return []
    return [
      ...root.querySelectorAll(
        '[data-automation-id*="attachment" i], [data-testid*="attachment" i], [class*="attachment" i], [title*=".xls" i], [aria-label*=".xls" i], [title*=".csv" i], [aria-label*=".csv" i]',
      ),
    ].filter(isVisible)
  }

  // ONE entry per REAL chip on the attachment strip. Chip containers carry
  // the filename with its extension in title/aria-label; leaf text elements
  // qualify when chip-sized. Big containers (whole panes, class*=attachment
  // wrappers) are rejected — their text would fake-match every filename.
  function chipList() {
    const root = composeRootEl()
    if (!root) return []
    const cand = attachAreaEls()
    const accepted = []
    // Outermost first: accept the chip CARD before its inner text leaf, so
    // nested matches collapse into one entry.
    cand.sort((a, b) => b.getBoundingClientRect().width * b.getBoundingClientRect().height - a.getBoundingClientRect().width * a.getBoundingClientRect().height)
    for (const el of cand) {
      if (accepted.some((c) => c.el.contains(el))) continue
      const attrs = `${el.getAttribute('title') || ''} ${el.getAttribute('aria-label') || ''}`.toLowerCase()
      const isCard = /\.(xls|csv)/.test(attrs)
      const compact = `${attrs} ${el.textContent || ''}`
        .toLowerCase()
        .replace(/[^a-z0-9]/g, '')
      if (!compact) continue
      if (!isCard) {
        // Non-card matches must be leaf + chip-sized to count
        if (el.querySelector('[data-automation-id*="attachment" i], [data-testid*="attachment" i], [class*="attachment" i], [title*=".xls" i]')) continue
        if (compact.length > 160) continue
      }
      accepted.push({ el, compact })
    }
    return accepted
  }

  // Real chip check: the file's compact name must appear on an attachment
  // LIST element — not in a toast, a pane, the subject or the body.
  function chipVerified(f) {
    const nameCompact = chipStem(f.name).replace(/[^a-z0-9]/g, '')
    if (!nameCompact) return false
    const short = nameCompact.slice(0, 14)
    for (const c of chipList()) {
      if (c.compact.includes(short)) return true
    }
    return false
  }

  function attachmentCount() {
    return chipList().length
  }

  function dispatchPaste(file, target) {
    const dt = new DataTransfer()
    dt.items.add(file)
    target.dispatchEvent(new ClipboardEvent('paste', { bubbles: true, cancelable: true, clipboardData: dt }))
  }

  function dispatchDrop(file, target) {
    const dt = new DataTransfer()
    dt.items.add(file)
    dt.effectAllowed = 'copy'
    const rect = target.getBoundingClientRect()
    const cx = rect.left + rect.width / 2
    const cy = rect.top + rect.height / 2
    target.dispatchEvent(new DragEvent('dragenter', { bubbles: true, cancelable: true, view: window, dataTransfer: dt, clientX: cx, clientY: cy }))
    target.dispatchEvent(new DragEvent('dragover', { bubbles: true, cancelable: true, view: window, dataTransfer: dt, clientX: cx, clientY: cy }))
    setTimeout(() => {
      target.dispatchEvent(new DragEvent('drop', { bubbles: true, cancelable: true, view: window, dataTransfer: dt, clientX: cx, clientY: cy }))
    }, 80)
  }

  function pasteTargets() {
    const out = []
    const body = bodyEditable()
    const root = composeRootEl()
    if (body) out.push(body)
    if (root && out[0] !== root) out.push(root)
    const hook = document.querySelector('[data-test-id*="attach" i], [data-testid*="attach" i], [data-automation-id*="attach" i]')
    if (hook && isVisible(hook) && !out.includes(hook)) out.push(hook)
    return out
  }

  // Strategy 0 — drive the compose's hidden <input type=file> directly.
  // This bypasses click automation entirely and is the most reliable path
  // whenever the build exposes one.
  function composeFileInputs() {
    const root = composeRootEl()
    const all = [...document.querySelectorAll('input[type="file"]')]
    const inCompose = root ? all.filter((el) => root.contains(el)) : []
    const rest = all.filter((el) => !inCompose.includes(el))
    return [...inCompose, ...rest].filter((el) => !el.disabled)
  }

  async function attachViaFileInput(f, countBefore) {
    // countBefore = chip count when THIS file's turn started. If a late chip
    // from a previous file lands while we wait, the baseline is re-read so a
    // neighbour's chip never counts as OUR success.
    let base = countBefore
    const file = makeFile(f)
    // Never cycle a long input list: each dispatch is a real attach request.
    // One event per input, stop the moment anything lands.
    const inputs = composeFileInputs().slice(0, 3)
    for (const inp of inputs) {
      try {
        if (attachmentCount() > base) {
          if (chipVerified(f)) return true
          base = attachmentCount()
        }
        const dt = new DataTransfer()
        dt.items.add(file)
        inp.files = dt.files
        inp.dispatchEvent(new Event('change', { bubbles: true }))
        const grew = await until(() => chipVerified(f) || attachmentCount() > base, { tries: 5, delay: 320 })
        if (grew) {
          if (chipVerified(f)) return true
          await sleep(500) // settle — late chips from neighbours may land here
          if (chipVerified(f)) return true
          base = attachmentCount()
        }
      } catch (e) {}
    }
    return chipVerified(f)
  }

  // Opening the ribbon "Attach file" menu mounts the real <input type=file>
  // that "Browse this computer" drives. Assigning files to it directly skips
  // the native picker completely; Escape closes the menu afterwards.
  async function attachViaAttachMenu(f, countBefore) {
    const opener = findByLabel(['attach'], { dataIdRx: 'attach' }) || findByVisibleText(['^Attach file$', 'Attach'])
    if (!opener) return false
    let base = countBefore
    try {
      opener.click()
      await sleep(350)
      const file = makeFile(f)
      const inputs = [...document.querySelectorAll('input[type="file"]')].filter((el) => !el.disabled).slice(0, 2)
      for (const inp of inputs) {
        try {
          const dt = new DataTransfer()
          dt.items.add(file)
          inp.files = dt.files
          inp.dispatchEvent(new Event('change', { bubbles: true }))
          if (await until(() => chipVerified(f) || attachmentCount() > base, { tries: 5, delay: 320 })) {
            await sleep(300)
            fireKey(document.body || document, 'Escape', 27)
            await sleep(150)
            fireKey(document.body || document, 'Escape', 27)
            return chipVerified(f)
          }
          base = attachmentCount()
        } catch (e) {}
      }
      fireKey(document.body || document, 'Escape', 27)
      await sleep(150)
      fireKey(document.body || document, 'Escape', 27)
      return chipVerified(f)
    } catch (e) {
      fireKey(document.body || document, 'Escape', 27)
      return false
    }
  }

  // Primary strategy: hand ALL still-missing files to ONE hidden input in ONE
  // change event. One event → one batch upload → no duplicate chips from
  // repeated dispatches racing async handlers.
  async function attachBatch(files, countBefore) {
    const base = attachmentCount()
    const inputs = composeFileInputs().slice(0, 2)
    for (const inp of inputs) {
      try {
        const dt = new DataTransfer()
        for (const f of files) dt.items.add(makeFile(f))
        inp.files = dt.files
        inp.dispatchEvent(new Event('change', { bubbles: true }))
        if (await until(() => files.every((f) => chipVerified(f)) && attachmentCount() > countBefore, { tries: 10, delay: 300 })) {
          await sleep(450) // settle — let Outlook finish its upload handling
          return true
        }
        // Anything landed at all? Then another dispatch could duplicate it —
        // stop cycling inputs and let per-file work cover the leftovers.
        if (attachmentCount() > base) break
      } catch (e) {}
    }
    return files.every((f) => chipVerified(f))
  }

  // Remove accidental duplicate chips (e.g. two copies of the same workbook
  // after a retried dispatch). Keeps the FIRST chip of every filename group
  // and clicks the remove control on the rest — trusting a click only when
  // the chip count actually drops. Exposed so runs self-heal the pile.
  async function dedupeChips() {
    const removed = []
    for (let pass = 0; pass < 3; pass++) {
      const groups = new Map()
      for (const c of chipList()) {
        if (!groups.has(c.compact)) groups.set(c.compact, [])
        groups.get(c.compact).push(c)
      }
      const dupGroups = [...groups.values()].filter((g) => g.length > 1)
      if (!dupGroups.length) break
      let anyRemoved = false
      for (const g of dupGroups) {
        for (const extra of g.slice(1)) {
          const countBefore = chipList().length
          const rm = findRemoveControl(extra.el)
          if (!rm) continue
          syntheticClick(rm)
          await sleep(400)
          if (chipList().length < countBefore) {
            removed.push(g[0].compact.slice(0, 140))
            anyRemoved = true
          } else {
            // The click may have opened a menu instead of removing — abort
            // further clicks on this group and close whatever opened.
            fireKey(document.body || document, 'Escape', 27)
            await sleep(150)
          }
        }
      }
      if (!anyRemoved) break
      await sleep(400)
    }
    return removed
  }

  function findRemoveControl(chipEl) {
    const sels = [
      '[aria-label*="delete" i]',
      '[aria-label*="remove" i]',
      '[aria-label*="close" i]',
      '[title*="delete" i]',
      '[title*="remove" i]',
      '[title*="close" i]',
      '[data-automation-id*="delete" i]',
      '[data-automation-id*="remove" i]',
    ]
    for (const s of sels) {
      const b = chipEl.querySelector(s)
      if (b && isVisible(b)) return b
    }
    return null
  }

  // The email carries ONLY the configured files: any chip on the strip that
  // doesn't belong (older config, junk added manually) is clicked away. Trust
  // a removal only when the chip count actually drops afterwards.
  function chipMatchesAnyFile(c, files) {
    for (const f of files) {
      const compact = chipStem(f.name).replace(/[^a-z0-9]/g, '').slice(0, 14)
      if (compact && c.compact.includes(compact)) return true
    }
    return false
  }

  async function pruneForeignChips(keepFiles) {
    const removed = []
    if (!keepFiles.length) return removed
    for (let pass = 0; pass < 4; pass++) {
      const foreign = chipList().filter((c) => !chipMatchesAnyFile(c, keepFiles))
      if (!foreign.length) break
      let anyRemoved = false
      for (const c of foreign) {
        const countBefore = chipList().length
        const rm = findRemoveControl(c.el)
        if (!rm) continue
        syntheticClick(rm)
        await sleep(400)
        if (chipList().length < countBefore) {
          removed.push(c.compact.slice(0, 12))
          anyRemoved = true
        } else {
          fireKey(document.body || document, 'Escape', 27)
          await sleep(150)
        }
      }
      if (!anyRemoved) break
      await sleep(300)
    }
    return removed
  }

  async function syncAttachOne(f) {
    const file = makeFile(f)
    const isImage = /^image\//i.test(file.type || '')
    const already = chipVerified(f)
    if (already) return { name: f.name, ok: true, already: true }
    const before = attachmentCount()

    // Strategy 0 — hidden file input
    if (await attachViaFileInput(f, before)) return { name: f.name, ok: true }

    // Strategy 0.5 — open the Attach menu, drive its file input directly
    if (await attachViaAttachMenu(f, before)) return { name: f.name, ok: true }

    // Strategy 1 — drop on the compose surface (OWA accepts file drops as
    // attachments). Prefer the compose root / attach hook over the body text.
    const targets = pasteTargets()
    const dropList = [targets[1], targets[2], targets[0]].filter(Boolean)
    for (const t of dropList) {
      dispatchDrop(file, t)
      if (await until(() => chipVerified(f) || attachmentCount() > before, { tries: 4, delay: 300 })) {
        return { name: f.name, ok: true }
      }
    }

    // Strategy 2 — paste (only works for images on this build; OWA rejects
    // xlsx paste with the "unsupported image file type" toast, so skip it).
    if (isImage) {
      for (const t of [...targets, document.body]) {
        dispatchPaste(file, t)
        if (await until(() => chipVerified(f) || attachmentCount() > before, { tries: 4, delay: 300 })) {
          return { name: f.name, ok: true }
        }
      }
    }

    return { name: f.name, ok: false }
  }

  // ---------------------------------------------------------------------------
  // Operations
  // ---------------------------------------------------------------------------

  const ops = {
    ping() {
      return { ok: true, url: location.href }
    },

    async newMail() {
      // An open compose (even the empty "Draft saved" one sitting in the tab)
      // is reused as-is.
      if (composeIsOpen()) return { ok: true, reused: true }

      const attempts = [
        () => findByLabel(['^\\s*new\\s*mail', '^new mail', 'new mail'], { dataIdRx: 'newmail|new-mail|new_mail' }),
        () => findByVisibleText(['^New mail', '^New email']),
        // The current OWA build labels the split-button "$ New" — exactly "New"
        // inside the mail module. Exact-anchored so "New tab"/"New chat" never match.
        () => findByLabel(['^\\s*new\\s*$']),
        () => findByVisibleText(['^New$']),
      ]
      let btn = null
      // Outlook may still be rendering — give the buttons a few rounds to
      // appear before giving up (the deep-link fallback then takes over).
      for (let round = 0; round < 4 && !btn; round++) {
        if (round) {
          await sleep(700)
          if (composeIsOpen()) return { ok: true, reused: true }
        }
        for (const a of attempts) {
          btn = a()
          if (btn) break
        }
      }

      if (btn) {
        btn.click()
        const ready = await until(composeIsOpen, { tries: 40, delay: 250 })
        if (ready) return { ok: true }
      }

      // Everything else failed — hand the deep-link compose URL back to the
      // background so it can navigate this same tab (server-rendered opener
      // that works on every OWA surface) and try once more.
      return {
        ok: false,
        error: 'New mail button not found',
        deeplinks: [`${location.origin}/mail/deeplink/compose`, `${location.origin}/mail/0/deeplink/compose`],
        dump: diagnostics(),
      }
    },

    async recipients(p) {
      const resTo = await typeRecipients('to', p.to || [])
      const resCc = (p.cc || []).length ? await typeRecipients('cc', p.cc) : []
      const missing = [...resTo, ...resCc].filter((r) => !r.ok).map((r) => r.value)
      let dump = null
      if (missing.length) dump = diagnostics()
      return {
        ok: resTo.every((r) => r.ok) && resCc.every((r) => r.ok),
        resTo,
        resCc,
        missing,
        dump,
      }
    },

    async subject(p) {
      const el = await until(subjectInput, { tries: 20 })
      if (!el) return { ok: false, error: 'Subject field not found', dump: diagnostics() }
      el.focus()
      await sleep(100)
      if (el.tagName === 'DIV' || el.isContentEditable) {
        document.execCommand('selectAll', false)
        document.execCommand('insertText', false, String(p.subject || ''))
      } else {
        setNativeValue(el, String(p.subject || ''))
      }
      return { ok: true }
    },

    async body(p) {
      const el = await until(bodyEditable, { tries: 20 })
      if (!el) return { ok: false, error: 'Message body not found', dump: diagnostics() }
      el.focus()
      await sleep(180)
      const existing = el.innerHTML
      const hasContent = existing && existing.replace(/<br\s*\/?>/gi, '').replace(/\s|&nbsp;/g, '').length > 0
      const keepSig = p.keepSignature !== false && hasContent
      document.execCommand('selectAll')
      const html = keepSig ? String(p.html) + '<br>' + existing : String(p.html)
      const done = document.execCommand('insertHTML', false, html)
      if (!done) el.innerHTML = html
      el.blur()
      return { ok: true, keptSignature: keepSig }
    },

    async attach(p) {
      const files = p.files || []

      // 0) The strip may carry chips that this config no longer wants (older
      //    runs, manually added files) — remove them first so the final state
      //    is EXACTLY p.files.
      let removed = await pruneForeignChips(files)

      // 1) Self-heal: strip accidental duplicate chips from earlier attempts
      removed = removed.concat(await dedupeChips())

      // 2) Skip what's already there (a reused draft keeps old chips)
      const already = files.filter((f) => chipVerified(f))
      const missing = files.filter((f) => !chipVerified(f))

      // 3) Batch all missing files through one change event, then catch the
      //    leftovers one by one with strict single-dispatch strategies
      if (missing.length >= 2) await attachBatch(missing, attachmentCount())
      const results = already.map((f) => ({ name: f.name, slot: f.slot, ok: true, already: true }))
      for (const f of files.filter((x) => !chipVerified(x))) {
        results.push({ ...(await syncAttachOne(f)), slot: f.slot })
        await sleep(250)
      }

      // 4) Final sweep — belt and braces: no duplicate chips may survive
      const removedAfter = await dedupeChips()
      removed = removed.concat(removedAfter)
      const prunedAfter = await pruneForeignChips(files)
      removed = removed.concat(prunedAfter)

      const okAll = results.every((r) => r.ok)
      return { ok: okAll, results, deduped: removed.length, count: attachmentCount() }
    },

    async send() {
      let btn = findByLabel(['^send$', '^send ', 'send ^', '^发送$'])
      if (!btn) btn = findByVisibleText(['^Send$', 'Send'])
      if (!btn) return { ok: false, error: 'Send button not found', dump: diagnostics() }
      btn.click()
      const gone = await until(() => !composeIsOpen(), { tries: 16, delay: 500 })
      return { ok: true, closedCompose: !!gone }
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
