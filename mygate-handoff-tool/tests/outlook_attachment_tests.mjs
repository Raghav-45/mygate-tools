/** Reproduce slow Outlook uploads, nested reply layouts, and menu-only removal. */
import assert from 'node:assert/strict'
import fs from 'node:fs'
import os from 'node:os'
import path from 'node:path'
import { fileURLToPath } from 'node:url'
import { spawnSync } from 'node:child_process'

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..')
const browser = ['/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/usr/bin/google-chrome', '/usr/bin/chromium'].find(p => fs.existsSync(p))
if (!browser) console.log('SKIP: slow Outlook attachment checks require Chrome or Chromium')
else {
  const temporary = fs.mkdtempSync(path.join(os.tmpdir(), 'handoff-slow-attachments-'))
  try {
    const source = fs.readFileSync(path.join(root, 'outlook.js'), 'utf8')
    const html = `<!doctype html><meta charset="utf-8"><style>
      body{min-height:1700px}main{width:800px;min-height:1500px}#editor{width:700px;height:260px}
      [role=option]{width:550px;height:50px}input{display:block}
      </style>
      <section id="original"><div role="option" aria-label="Old_report.xlsx Open 88 KB"><span title="Old_report.xlsx">Old report</span></div></section>
      <main id="compose"><input id="subject" aria-label="Subject" value="Re: Mygate Complaint Summary.">
      <div aria-label="To" contenteditable="true"><span contenteditable="false">Divya Gulati</span></div>
      <div aria-label="Cc" contenteditable="true"><span contenteditable="false">Ashok Kumar</span><span contenteditable="false">Estate If3</span></div>
      <div role="listbox" aria-label="file attachments" id="strip"></div>
      <input id="image-upload" type="file" accept="image/*" multiple hidden>
      <input id="upload" type="file" multiple hidden><input id="other-upload" type="file" multiple hidden>
      ${'<div>'.repeat(10)}<div id="editor" aria-label="Message body" role="textbox" contenteditable="true"><p>Signature</p></div>${'</div>'.repeat(10)}</main>
      <script>
      const normalTimeout = window.setTimeout.bind(window);
      window.setTimeout = (fn, ms, ...args) => normalTimeout(fn, Math.min(ms || 0, 2), ...args);
      let listener;
      new Function('chrome', ${JSON.stringify(source)})({runtime:{onMessage:{addListener:fn=>listener=fn}}});
      const run = files => new Promise(resolve=>listener({aio:'attach',payload:{files}}, {}, resolve));
      const check=(ok,message)=>{if(!ok)throw new Error(message)};
      const strip=document.getElementById('strip');
      const files=[{name:'MyGate_Master_Helpdesk_Dump_2024-01-01_to_2026-11-04.xlsx',base64:btoa('dump')},
        {name:'Mygate Complaints Dashboard as on 04.11.26.xlsx',base64:btoa('analysis')}];
      let uploadEvents=0,imageEvents=0,mode='normal',menuCard=null;
      const recipientsBefore=[...document.querySelectorAll('[aria-label=To],[aria-label=Cc]')].map(e=>e.innerHTML).join('|');
      function card(file, opaque=false) {
        const el=document.createElement('div');el.setAttribute('role','option');el.dataset.testFilename=file.name;
        const label=document.createElement('span');
        if(opaque){el.setAttribute('aria-label','Uploading file');label.textContent=file.name.slice(0,18)+'…'}
        else{el.setAttribute('aria-label',file.name+' Open 88 KB');label.title=file.name;label.textContent=file.name.slice(0,22)+'…'}
        const more=document.createElement('button');more.setAttribute('aria-label','More actions');more.setAttribute('aria-haspopup','menu');
        more.onclick=()=>{document.querySelector('[role=menu]')?.remove();menuCard=el;const menu=document.createElement('div');menu.setAttribute('role','menu');const remove=document.createElement('button');remove.setAttribute('role','menuitem');remove.textContent='Remove attachment';remove.onclick=()=>{menuCard.remove();menu.remove()};menu.append(remove);document.body.append(menu)};
        el.append(label,more);strip.append(el);return el;
      }
      document.getElementById('image-upload').onchange=()=>imageEvents++;
      for(const id of ['upload','other-upload'])document.getElementById(id).onchange=event=>{
        uploadEvents++;const incoming=[...event.target.files];normalTimeout(()=>incoming.forEach((f,i)=>{if(mode!=='partial'||i===0)card(f,mode==='opaque')}),80);
      };
      let generation=0;
      function reset(nextMode='normal') {mode=nextMode;strip.replaceChildren();uploadEvents=0;imageEvents=0;generation++;for(const e of [document.getElementById('subject'),document.getElementById('editor')])e.setAttribute('data-mygate-handoff-draft','batch-'+generation)}
      (async()=>{
        reset();let result=await run(files);
        check(result.ok,'Nested reply failed to recognize its delayed attachment cards');
        check(uploadEvents===1&&imageEvents===0&&strip.children.length===2,'A slow upload dispatched duplicate batches or used the image picker');
        result=await run(files);check(result.ok&&uploadEvents===1,'Retry submitted a completed upload again');
        check(document.getElementById('original').children.length===1,'Attachments on the original message were changed');
        check(recipientsBefore===[...document.querySelectorAll('[aria-label=To],[aria-label=Cc]')].map(e=>e.innerHTML).join('|'),'Native display-name recipients were changed');

        reset('opaque');result=await run(files);check(!result.ok,'Unverifiable filenames were marked complete');
        check(uploadEvents===1&&strip.children.length===2,'Uncertain recognition triggered another upload method');
        result=await run(files);check(!result.ok&&uploadEvents===1,'An uncertain upload was repeated on retry');
        for(const e of [document.getElementById('subject'),document.getElementById('editor')])e.removeAttribute('data-mygate-handoff-upload-attempts');
        result=await run(files);check(!result.ok&&uploadEvents===1,'Unidentified existing attachments caused another upload after attempt markers were lost');
        for(const el of strip.children){el.setAttribute('aria-label',el.dataset.testFilename+' Open 88 KB');el.firstElementChild.title=el.dataset.testFilename}
        result=await run(files);check(result.ok&&uploadEvents===1,'A verified late upload could not recover without uploading again');

        reset('partial');result=await run(files);check(!result.ok&&uploadEvents===1&&strip.children.length===1,'Partial batch triggered extra upload attempts');
        for(const e of [document.getElementById('subject'),document.getElementById('editor')])e.removeAttribute('data-mygate-handoff-upload-attempts');
        result=await run(files);check(!result.ok&&uploadEvents===1,'Missing file was automatically submitted again');
        card(files[1]);result=await run(files);check(result.ok&&uploadEvents===1,'Manually attached missing file was not accepted');

        reset();for(const file of files)for(let i=0;i<3;i++)card(file);
        result=await run(files);check(result.ok&&uploadEvents===0&&strip.children.length===2,'Menu-only removal did not reduce three copies of each file to one');

        reset();for(const file of files)card(file);const opaqueExtra=card(files[0],true);
        result=await run(files);check(!result.ok&&uploadEvents===0&&strip.children.length===3,'An unidentified extra attachment was ignored when marking the reply complete');
        opaqueExtra.remove();result=await run(files);check(result.ok&&uploadEvents===0&&strip.children.length===2,'Removing an unidentified extra attachment did not recover the reply');

        reset();card(files[0]);const unremovable=card(files[0]);unremovable.querySelector('button').remove();
        result=await run(files);check(!result.ok&&uploadEvents===0&&strip.children.length===2,'Unremovable duplicates triggered another upload');

        reset();const concurrent=await Promise.all([run(files),run(files)]);
        check(concurrent.some(r=>r.ok)&&concurrent.some(r=>!r.ok)&&uploadEvents===1&&strip.children.length===2,'Concurrent attachment commands duplicated uploads');
        document.body.dataset.testResult='PASS';
      })().catch(error=>document.body.dataset.testResult=error.message);
      </script>`
    const file = path.join(temporary, 'outlook.html')
    fs.writeFileSync(file, html)
    const result = spawnSync(browser, ['--headless', '--no-sandbox', '--disable-gpu', '--no-first-run',
      `--user-data-dir=${path.join(temporary, 'profile')}`, '--virtual-time-budget=20000', '--dump-dom', `file://${file}`],
    { encoding: 'utf8', timeout: 30000, maxBuffer: 1024 * 1024 })
    assert.equal(result.status, 0, result.stderr)
    const outcome = result.stdout.match(/data-test-result="([^"]*)"/)?.[1]
    assert.equal(outcome, 'PASS', outcome || 'Slow attachment checks did not finish')
    console.log('PASS: nested replies, delayed/opaque/partial uploads, retry/concurrency guards, native recipients, and menu-only duplicate removal')
  } finally { fs.rmSync(temporary, { recursive: true, force: true }) }
}
