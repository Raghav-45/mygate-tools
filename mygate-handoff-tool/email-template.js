/**
 * MyGate Handoff Tool — Email & Table Builders
 *
 * Shared module imported by the background service worker (via importScripts).
 * Produces the exact Outlook body that replicates the manual email template:
 *
 *   Dear Ma'am/Sir,
 *   Please find the attached My Gate Complaint Summary for your reference
 *
 *   [ Pending Mygate Tickets - From 1-1-2024 To 2-11-2026  (blue banner row) ]
 *   [ Category | Total | Resolved | Open                   (blue header row) ]
 *   [ ...category rows...                                    (numbers right) ]
 *   [ Total    | ...                                       (final row)       ]
 *
 *   (signature — either preserved from Outlook or the SGV fallback below)
 */

(function (global) {
  'use strict'

  const TABLE_BLUE = '#2E75B6'
  const BORDER = '#000000'
  const FONT_STACK = "'Aptos', 'Calibri', 'Segoe UI', sans-serif"

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
  }

  // Email template uses plain un-formatted numbers (e.g. 28597), like the
  // original manual email — no thousand separators.
  function plainNumber(n) {
    const v = Number(n)
    return Number.isFinite(v) ? String(v) : '0'
  }

  function cell(value, opts) {
    const o = opts || {}
    const bg = o.bg ? `background:${o.bg};` : ''
    const color = o.color ? `color:${o.color};` : ''
    const weight = o.bold ? 'font-weight:bold;' : ''
    const align = o.align || 'left'
    return (
      `<td style="border:1px solid ${BORDER};padding:0 8pt;${bg}${color}${weight}` +
      `text-align:${align};vertical-align:middle;">${esc(value)}</td>`
    )
  }

  /**
   * Build the blue "Pending Mygate Tickets" table exactly like the manual email.
   * @param rows    [{name, total, resolved, open}]
   * @param summary {total, resolved, open}
   * @param fromFmt "1-1-2024"   (d-m-yyyy, no leading zeros)
   * @param toFmt   "2-11-2026"
   */
  function buildPendingTableHtml(rows, summary, fromFmt, toFmt) {
    const width = 'width:60%;width:10%;width:15%;width:15%'
    let html =
      `<table style="border-collapse:collapse;font-family:${FONT_STACK};` +
      `font-size:12pt;color:#000;margin:0;max-width:640px;"><tbody>`

    // Title banner row
    html +=
      `<tr>` +
      `<td colspan="4" style="border:1px solid ${BORDER};background:${TABLE_BLUE};` +
      `color:#FFFFFF;font-weight:bold;text-align:center;padding:2pt 8pt;">` +
      `Pending Mygate Tickets - From ${esc(fromFmt)} To ${esc(toFmt)}</td></tr>`

    // Header row
    html +=
      `<tr>` +
      `<td style="border:1px solid ${BORDER};background:${TABLE_BLUE};color:#FFFFFF;` +
      `font-weight:bold;text-align:left;padding:2pt 8pt;">Category</td>` +
      `<td style="border:1px solid ${BORDER};background:${TABLE_BLUE};color:#FFFFFF;` +
      `font-weight:bold;text-align:right;padding:2pt 8pt;">Total</td>` +
      `<td style="border:1px solid ${BORDER};background:${TABLE_BLUE};color:#FFFFFF;` +
      `font-weight:bold;text-align:right;padding:2pt 8pt;">Resolved</td>` +
      `<td style="border:1px solid ${BORDER};background:${TABLE_BLUE};color:#FFFFFF;` +
      `font-weight:bold;text-align:right;padding:2pt 8pt;">Open</td></tr>`

    // Category rows
    for (const r of rows) {
      html +=
        `<tr>` +
        cell(r.name, { align: 'left' }) +
        cell(plainNumber(r.total), { align: 'right' }) +
        cell(plainNumber(r.resolved), { align: 'right' }) +
        cell(plainNumber(r.open), { align: 'right' }) +
        `</tr>`
    }

    // Total row
    html +=
      `<tr>` +
      cell('Total', { align: 'left', bold: true }) +
      cell(plainNumber(summary.total), { align: 'right', bold: true }) +
      cell(plainNumber(summary.resolved), { align: 'right', bold: true }) +
      cell(plainNumber(summary.open), { align: 'right', bold: true }) +
      `</tr>`

    html += '</tbody></table>'
    void width
    return html
  }

  /**
   * Full body: greeting + optional table + static signature fallback.
   * (When Outlook already has a signature in the fresh compose, the content
   * script preserves it and this static part is replaced by it.)
   */
  function buildBodyHtml(opts) {
    const o = opts || {}
    const parts = []
    parts.push(
      `<div style="font-family:${FONT_STACK};font-size:11pt;color:#000;">` +
        `<p style="margin:0;">Dear Ma&rsquo;am/Sir,</p>` +
        `<p style="margin:0;">Please find the attached My Gate Complaint Summary for your reference</p>`,
    )
    if (o.tableHtml) parts.push(`<br>`, o.tableHtml)
    if (o.signatureHtml) {
      parts.push(`<br><br><br><br><br><br><br><br><br><br>`, o.signatureHtml)
    }
    parts.push('</div>')
    return parts.join('')
  }

  /** Static fallback signature (used only when the Outlook compose has none). */
  function fallbackSignatureHtml() {
    return (
      `<div style="font-family:${FONT_STACK};font-size:11pt;color:#000;">` +
      `<div style="border-top:1px solid #999999;margin:0 0 8pt 0;width:40%;">&nbsp;</div>` +
      `<p style="margin:0;">Siya</p>` +
      `<p style="margin:0;">Help desk</p>` +
      `<p style="margin:0;">+91 95992 94626</p>` +
      `<br>` +
      `<p style="margin:0;">DLF City Floors, Community Centre,</p>` +
      `<p style="margin:0;">Phase III, Gurugram - 122010</p>` +
      `</div>`
    )
  }

  const DEFAULTS = {
    subject: 'Mygate Complaint Summary.',
  }

  global.EmailTemplate = {
    TABLE_BLUE,
    buildPendingTableHtml,
    buildBodyHtml,
    fallbackSignatureHtml,
    DEFAULTS,
  }
})(self)
