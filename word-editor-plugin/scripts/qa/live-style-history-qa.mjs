/** Focused 2.20 browser proof: style transactions preserve body, unrelated styles/defaults and Undo/Redo. */
import assert from 'node:assert/strict';
import {buildSync} from 'esbuild';
import {openApp} from './harness.mjs';
const bundled = buildSync({entryPoints:['src/engine/live-style-edit.ts'],bundle:true,platform:'browser',format:'iife',globalName:'__styleLive',write:false}).outputFiles[0].text;
const app=await openApp({name:'live-style-history',port:Number(process.env.QA_PORT??9875),settle:1000,extra:`<script>${bundled.replaceAll('</script','<\\/script')}</script>`});
try{
  const ask=async(code)=>JSON.parse(await app.js(`(async()=>JSON.stringify(await (${code})))()`));
  const doc='window.__qa.doc()';
  const text=()=>ask(`${doc}.getText({})`);
  const history=()=>ask(`${doc}.history.get().then(h=>({undo:h.undoDepth,redo:h.redoDepth}))`);
  const run=async(xml,edit,revision)=>ask(`__styleLive.applyLiveStyleEdit(window.__qa.sd(),${JSON.stringify(xml)},${JSON.stringify(edit)},${JSON.stringify(revision)})`);
  const info=()=>ask(`${doc}.info({})`);
  const baseline=await app.docx();
  await app.js('window.__styleOriginalHost=window.__qa.sd()');
  const block=await ask(`${doc}.blocks.list().then(r=>r.blocks[0].nodeId)`);
  for(const [value,offset] of [['FIRST',0],['SECOND',5]]){
    const target={kind:'selection',start:{kind:'text',blockId:block,offset},end:{kind:'text',blockId:block,offset}};
    await ask(`${doc}.insert({value:${JSON.stringify(value)},type:'text',target:${JSON.stringify(target)}})`);
  }
  const prior=await history();assert.ok(prior.undo>=2);
  assert.equal(await text(),'FIRSTSECOND');
  let revision=(await info()).revision;
  let result=await run(baseline['word/styles.xml'],{kind:'create',name:'Unused Regression Style',type:'paragraph',basedOn:'Normal',next:null,format:{fontSizePt:17}},revision);
  assert.equal(result.outcome.ok,true);
  const styleId=result.id;
  const created=await app.docx();
  const node=(xml,id)=>xml.match(new RegExp(`<w:style\\b[^>]*w:styleId="${id}"[^>]*>[\\s\\S]*?</w:style>`))?.[0];
  const unusedCustom=node(created['word/styles.xml'],styleId);assert.ok(unusedCustom);
  revision=(await info()).revision;
  result=await run(created['word/styles.xml'],{kind:'modify',id:'Normal',format:{fontSizePt:19}},revision);assert.equal(result.outcome.ok,true);
  const changed=await app.docx();
  assert.equal(node(changed['word/styles.xml'],styleId),unusedCustom);
  assert.equal(node(changed['word/styles.xml'],'Heading9'),node(baseline['word/styles.xml'],'Heading9'));
  const singleton=(xml,name)=>xml.match(new RegExp(`<w:${name}\\b[^>]*>[\\s\\S]*?</w:${name}>|<w:${name}\\b[^>]*/>`))?.[0];
  for(const name of ['docDefaults','latentStyles'])assert.equal(singleton(changed['word/styles.xml'],name),singleton(baseline['word/styles.xml'],name));
  for(const part of ['word/settings.xml','word/theme/theme1.xml','word/numbering.xml','word/fontTable.xml','word/_rels/document.xml.rels'])assert.equal(changed[part],baseline[part],part);
  assert.equal(await text(),'FIRSTSECOND');assert.equal(await app.js('window.__qa.sd()===window.__styleOriginalHost'),true);
  assert.equal((await history()).undo,prior.undo+2);
  await ask(`${doc}.history.undo()`);assert.equal(await text(),'FIRSTSECOND');
  await ask(`${doc}.history.undo()`);assert.equal(await text(),'FIRSTSECOND');
  for(let i=0;i<prior.undo;i++)await ask(`${doc}.history.undo()`);
  assert.equal(await text(),'');
  for(let i=0;i<prior.undo+2;i++)await ask(`${doc}.history.redo()`);
  assert.equal(await text(),'FIRSTSECOND');
  const afterRedo=await app.docx();assert.ok(node(afterRedo['word/styles.xml'],'Normal').includes('w:val="38"'));
  assert.equal(node(afterRedo['word/styles.xml'],styleId),unusedCustom);
  const stale=await run(baseline['word/styles.xml'],{kind:'modify',id:'Normal',format:{fontSizePt:13}},'0');
  assert.equal(stale.outcome.ok,false);assert.match(stale.outcome.message,/REVISION_MISMATCH/);
  assert.equal(await text(),'FIRSTSECOND');
  console.log('PASS: existing text + style Undo/Redo, same instance, unused custom/builtin styles, defaults/sidecars, stale revision refusal');
}finally{app.close()}
