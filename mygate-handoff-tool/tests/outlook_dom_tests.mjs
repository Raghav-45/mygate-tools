/** Exercise the real Outlook content script against browser DOM and file inputs. */
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'

const extension = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const browser = ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/usr/bin/google-chrome', '/usr/bin/chromium'].find(p => fs.existsSync(p))
if (!browser) {
  console.log('SKIP: Outlook DOM checks require Chrome or Chromium')
} else {
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'handoff-outlook-dom-'))
  try {
    const source = fs.readFileSync(path.join(extension, 'outlook.js'), 'utf8')
    const template = fs.readFileSync(path.join(extension, 'email-template.js'), 'utf8')
    const fixture = `<!doctype html><meta charset="utf-8">
      <style>body{min-height:1400px}main{width:750px;height:650px}input{display:block}#editor{width:700px;height:250px}.attachment{width:500px;height:45px}</style>
      <section role="main" aria-label="Reading Pane"><h3 id="report-subject">Mygate Complaint Summary.</h3>
        <button id="reply-all" aria-label="Reply all">Reply all</button><button id="new-mail">New mail</button>
        <div id="previous-message">Previous sent report and its attachments</div></section>
      <main id="compose"><div id="to" aria-label="To" contenteditable="true" style="width:700px;min-height:30px"></div><input aria-label="Subject">
        <button id="show-cc" aria-label="Show Cc">Cc</button>
        <div id="cc-row" hidden><div id="cc" aria-label="Cc" contenteditable="true" style="width:700px;min-height:30px"></div></div>
        <input id="bcc" aria-label="Bcc">
        <div id="suggestion" role="option" hidden></div>
        <div id="attachments" data-automation-id="attachments"></div>
        <input id="upload" type="file" multiple hidden>
        <div id="editor" contenteditable="true" role="textbox" aria-label="Message body"><p>Existing Outlook signature</p><img alt="Existing signature logo" src="data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7"><blockquote>Earlier report history</blockquote></div>
      </main><script>
      const normalTimeout = window.setTimeout.bind(window);
      window.setTimeout = (fn, ms, ...args) => normalTimeout(fn, Math.min(ms || 0, 5), ...args);
      let listener;
      const shim = {runtime:{onMessage:{addListener: fn => {listener = fn}}}};
      const pageLocation = {href:'https://outlook.cloud.microsoft/mail/sentitems/id/previous-report'};
      new Function('chrome', 'location', ${JSON.stringify(source)})(shim, pageLocation);
      new Function(${JSON.stringify(template)})();
      const run = (op, payload = {}) => new Promise(resolve => listener({aio:op,payload}, {}, resolve));
      const check = (ok, message) => {if (!ok) throw new Error(message)};
      const upload = document.getElementById('upload');
      const strip = document.getElementById('attachments');
      const to = document.getElementById('to');
      const cc = document.getElementById('cc');
      const suggestion = document.getElementById('suggestion');
      let replyClicks = 0, newMailClicks = 0;
      document.getElementById('reply-all').onclick = () => {
        replyClicks++;
        document.getElementById('compose').hidden = false;
        document.querySelector('[aria-label="Subject"]').value = 'Re: Mygate Complaint Summary.';
      };
      document.getElementById('new-mail').onclick = () => {newMailClicks++};
      let suggestionEditor = to;
      document.getElementById('show-cc').onclick = () => setTimeout(() => {document.getElementById('cc-row').hidden = false}, 30);
      function bindRecipientEditor(editor) { editor.addEventListener('input', () => {
        const copy = editor.cloneNode(true);
        copy.querySelectorAll('[contenteditable="false"]').forEach(el => el.remove());
        const value = copy.textContent.replace(/[\\u200b-\\u200d\\ufeff]/g, '').trim();
        suggestion.hidden = !value;
        suggestion.dataset.address = value;
        suggestion.textContent = 'Use this address: ' + value;
        suggestionEditor = editor;
      }) }
      bindRecipientEditor(to);
      bindRecipientEditor(cc);
      suggestion.onclick = () => {
        const value = suggestion.dataset.address;
        if (!value) return;
        const editor = suggestionEditor;
        [...editor.childNodes].filter(el => el.nodeType !== 1 || !el.matches('[contenteditable="false"]')).forEach(el => el.remove());
        const chip = document.createElement('span');
        chip.className = '_Entity';
        chip.setAttribute('contenteditable', 'false');
        chip.title = value;
        chip.textContent = value;
        editor.append('\\u200b', chip, '\\u200b');
        suggestion.hidden = true;
        suggestion.dataset.address = '';
      };
      let uploadEvents = 0;
      let blockedName = null;
      function addChip(file) {
        if (file.name === blockedName) return;
        const card = document.createElement('div');
        card.className = 'attachment';
        card.title = file.name;
        card.dataset.filename = file.name;
        const name = document.createElement('span');
        name.textContent = file.name;
        const remove = document.createElement('button');
        remove.setAttribute('aria-label', 'Remove attachment');
        remove.textContent = 'Remove';
        remove.onclick = () => card.remove();
        card.append(name, remove);
        strip.append(card);
      }
      upload.addEventListener('change', () => {uploadEvents++; [...upload.files].forEach(addChip)});
      document.getElementById('compose').addEventListener('drop', event => [...event.dataTransfer.files].forEach(addChip));
      const files = [
        {slot:'dump',name:'MyGate_Master_Helpdesk_Dump_2024-01-01_to_2024-01-02.xlsx',base64:btoa('synthetic dump')},
        {slot:'analysis',name:'Mygate Complaints Dashboard as on 02.01.24.xlsx',base64:btoa('synthetic analysis')}
      ];
      (async () => {
        const replyPayload = {draftKey:'batch-1',subject:'Mygate Complaint Summary.'};
        let result = await run('replyAll', replyPayload);
        check(!result.ok && replyClicks === 0, 'An unrelated draft was reused');
        document.getElementById('compose').hidden = true;
        pageLocation.href = 'https://outlook.cloud.microsoft/mail/inbox/id/other-message';
        result = await run('replyAll', replyPayload);
        check(!result.ok && replyClicks === 0, 'Reply opened outside Sent Items');
        pageLocation.href = 'https://outlook.cloud.microsoft/mail/sentitems/id/previous-report';
        document.getElementById('report-subject').textContent = 'Unrelated sent email';
        result = await run('replyAll', replyPayload);
        check(!result.ok && replyClicks === 0, 'Reply opened on the wrong sent message');
        document.getElementById('report-subject').textContent = 'Mygate Complaint Summary.';
        document.getElementById('reply-all').hidden = true;
        result = await run('replyAll', replyPayload);
        check(!result.ok && newMailClicks === 0, 'Missing Reply all fell back to New mail');
        document.getElementById('reply-all').hidden = false;
        result = await run('replyAll', replyPayload);
        check(result.ok && !result.reused && replyClicks === 1, 'The open sent report did not open a reply-all draft');
        check(document.querySelector('[aria-label="Subject"]').value === 'Re: Mygate Complaint Summary.', 'The native reply subject changed');
        result = await run('replyAll', replyPayload);
        check(result.ok && result.reused && replyClicks === 1, 'Retry opened a second reply instead of reusing its own draft');
        result = await run('replyAll', {...replyPayload,draftKey:'batch-2'});
        check(!result.ok && replyClicks === 1 && newMailClicks === 0, 'A new report batch reused the old reply draft');

        to.innerHTML = '<span contenteditable="false">Divya Gulati</span><span contenteditable="false">Pankaj Pawar</span><span contenteditable="false">SOORAJ</span>';
        document.getElementById('cc-row').hidden = false;
        cc.innerHTML = ['Ashok Kumar','Estate If3','Kirshan Chander','Anurag Yadav','Anurag Sharma'].map(name=>'<span contenteditable="false">'+name+'</span>').join('');
        const nativeTo = to.innerHTML, nativeCc = cc.innerHTML;
        result = await run('recipients', {to:['different@example.invalid'],cc:['tech.if3@sgvgroup.in']});
        check(!result.ok && to.innerHTML === nativeTo && cc.innerHTML === nativeCc, 'A recipient-entry command modified native display-name chips');

        result = await run('attach', {files});
        check(result.ok && result.results.length === 2 && result.results.every(r => r.ok), 'Batch upload did not report both attached files');
        check(uploadEvents === 1 && strip.children.length === 2, 'Batch upload dispatched duplicate files');
        result = await run('attach', {files});
        check(result.ok && result.results.every(r => r.already) && uploadEvents === 1, 'Retry uploaded existing attachments again');

        const nextDump = {...files[0],name:'MyGate_Master_Helpdesk_Dump_2024-01-01_to_2024-01-03.xlsx'};
        result = await run('attach', {files:[nextDump]});
        check(result.ok && result.results.length === 1 && result.results[0].ok, 'New dated dump failed to attach');
        check(strip.children.length === 1 && strip.firstElementChild.dataset.filename === nextDump.name, 'A similarly named old dump was mistaken for the new file');

        strip.replaceChildren();
        for (const el of [document.querySelector('[aria-label="Subject"]'),document.getElementById('editor')]) el.setAttribute('data-mygate-handoff-draft','partial-upload-batch');
        blockedName = files[1].name;
        result = await run('attach', {files});
        check(!result.ok && result.results.length === 2, 'Partial upload did not return one result per requested file');
        check(result.results.find(r => r.slot === 'dump').ok && !result.results.find(r => r.slot === 'analysis').ok, 'Partial upload marked a missing file successful');

        const html = EmailTemplate.buildBodyHtml({tableHtml:'<table><tr><td>Pending table</td></tr></table>'});
        const payload = {html,keepSignature:true,fallbackSignatureHtml:'<p>Fallback signature</p>'};
        await run('body', payload);
        await run('body', payload);
        let editor = document.getElementById('editor');
        check((editor.textContent.match(/Please find the attached/g) || []).length === 1, 'Retry duplicated the greeting or table');
        check((editor.textContent.match(/Existing Outlook signature/g) || []).length === 1 && !editor.textContent.includes('Fallback signature'), 'Existing signature was duplicated or combined with the fallback');
        check(editor.querySelectorAll('img[alt="Existing signature logo"]').length === 1, 'The native signature logo was removed or duplicated');
        check((editor.textContent.match(/Earlier report history/g) || []).length === 1, 'The quoted report history was removed or duplicated');
        check(document.getElementById('previous-message').textContent === 'Previous sent report and its attachments', 'The original sent message outside the editor changed');
        check(to.innerHTML === nativeTo && cc.innerHTML === nativeCc && document.getElementById('bcc').value === '', 'Draft preparation changed native To/Cc/Bcc recipients');
        const emptyEditor = editor.cloneNode(false);
        emptyEditor.innerHTML = '<div><br></div>';
        editor.replaceWith(emptyEditor);
        await run('body', payload);
        check(emptyEditor.textContent.includes('Fallback signature'), 'Fallback signature missing when Outlook has no signature');
        document.body.dataset.testResult = 'PASS';
      })().catch(error => {document.body.dataset.testResult = error.message});
      </script>`
    const file = path.join(temporary, 'outlook.html')
    fs.writeFileSync(file, fixture)
    const result = spawnSync(browser, ['--headless', '--no-sandbox', '--disable-gpu', '--no-first-run',
      `--user-data-dir=${path.join(temporary, 'profile')}`, '--virtual-time-budget=10000', '--dump-dom', `file://${file}`],
    { encoding: 'utf8', timeout: 30000, maxBuffer: 1024 * 1024 })
    assert.equal(result.status, 0, result.stderr)
    const outcome = result.stdout.match(/data-test-result="([^"]*)"/)?.[1]
    assert.equal(outcome, 'PASS', outcome || 'Outlook DOM check did not finish')
    console.log('PASS: real Outlook script preserves native To/Cc/Bcc, attachments, retries, signatures, logo, and quoted history')
  } finally { fs.rmSync(temporary, { recursive: true, force: true }) }
}
