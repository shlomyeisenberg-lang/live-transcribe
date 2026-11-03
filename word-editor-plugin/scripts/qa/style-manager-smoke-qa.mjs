/** Fresh, focused user-path proof for live style editing, creation and rename. */
import assert from 'node:assert/strict';
import {openApp,sleep} from './harness.mjs';
const app=await openApp({name:'style-manager-smoke',port:Number(process.env.QA_PORT??9876),settle:1000});
try{
  await app.cdp.send('Emulation.setDeviceMetricsOverride',{width:1440,height:413,deviceScaleFactor:1,mobile:false});
  await sleep(500);
  await app.tab('בית');
  const ask=async(code)=>JSON.parse(await app.js(`(async()=>JSON.stringify(await (${code})))()`));
  const text=()=>ask('window.__qa.doc().getText({})');
  const until=async(predicate)=>{for(let i=0;i<40;i++){if(await predicate())return;await sleep(250)}throw new Error('style dialog did not settle: '+await app.js(`document.querySelector('.sd-error')?.textContent||document.querySelector('.status-message')?.textContent||'no error'`))};
  const fill=async(selector,value)=>app.js(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});if(!e)throw Error('missing field');e.value=${JSON.stringify(value)};e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}))})()`);
  const rightClick=async(selector)=>{
    const r=await ask(`(()=>{const e=document.querySelector(${JSON.stringify(selector)});e.scrollIntoView({block:'nearest',inline:'nearest'});const r=e.getBoundingClientRect();return {x:r.x+r.width/2,y:r.y+r.height/2}})()`);
    for(const type of ['mouseMoved','mousePressed','mouseReleased']){await app.cdp.send('Input.dispatchMouseEvent',{type,x:Math.round(r.x),y:Math.round(r.y),button:type==='mouseMoved'?'none':'right',buttons:type==='mousePressed'?2:0,clickCount:1});await sleep(50)}
    await sleep(200);
  };
  const click=async(selector)=>{
    await app.js(`document.querySelector(${JSON.stringify(selector)})?.scrollIntoView({block:'nearest',inline:'nearest'})`);
    return app.clickSel(selector,0,{after:100});
  };
  const open=async(action)=>{assert.equal(await click(`[data-testid="style-menu"] [data-action="${action}"]`),true);await until(()=>app.exists('[data-testid="style-dialog"]'));};
  const submit=async()=>{assert.equal(await click('[data-testid="style-dialog"] .sd-btn-primary'),true);await until(async()=>!(await app.exists('[data-testid="style-dialog"]')));};
  await app.caretPara('',{verifyMs:10000});await app.type('בדיקת סגנונות בטוחה',30);
  await app.js('window.__styleSmokeHost=window.__qa.sd()');
  const before=await ask('window.__qa.doc().history.get().then(h=>h.undoDepth)');
  await rightClick('.style-card');await open('modify');
  await fill('[data-testid="style-dialog"] input[aria-label="גודל, בנקודות"]','17');await submit();
  assert.equal(await text(),'בדיקת סגנונות בטוחה');
  const after=await ask('window.__qa.doc().history.get().then(h=>h.undoDepth)');assert.equal(after,before+1);
  await sleep(300);
  const font=await app.js(`(()=>{const e=document.querySelector('.editor-stack .superdoc-text-run');return e?getComputedStyle(e).fontSize:null})()`);assert.equal(font,'22.6667px');
  assert.equal(await app.js('window.__qa.sd()===window.__styleSmokeHost'),true);
  await app.tab('בית');assert.equal(await click('[data-testid="style-gallery-more"]'),true);await open('create');
  await fill('#sd-name','סגנון בדיקה בטוח');await fill('[data-testid="style-dialog"] input[aria-label="גודל, בנקודות"]','18');await submit();
  assert.equal(await text(),'בדיקת סגנונות בטוחה');
  const created=await ask(`window.__qa.doc().styles.getCatalog({view:'currentDocument'}).then(c=>c.items.find(s=>s.name==='סגנון בדיקה בטוח'))`);assert.ok(created?.id);await until(()=>app.exists('.style-card[data-tip-title="סגנון בדיקה בטוח"]'));
  await rightClick('.style-card[data-tip-title="סגנון בדיקה בטוח"]');await open('rename');await fill('#sd-name','סגנון בדיקה חדש');await submit();
  const renamed=await ask(`window.__qa.doc().styles.getCatalog({view:'currentDocument'}).then(c=>c.items.find(s=>s.id===${JSON.stringify(created.id)}))`);assert.equal(renamed.name,'סגנון בדיקה חדש');await until(()=>app.exists('.style-card[data-tip-title="סגנון בדיקה חדש"]'));
  assert.equal(await app.js('window.__qa.sd()===window.__styleSmokeHost'),true);
  const files=await app.docx();assert.ok(files['word/document.xml'].includes(`w:val="${created.id}"`));assert.ok(files['word/styles.xml'].includes('סגנון בדיקה חדש'));
  console.log('PASS: physical pointer gallery/dialog modify, live rendering, create/apply/rename, content/history and instance preserved');
}finally{app.close()}
