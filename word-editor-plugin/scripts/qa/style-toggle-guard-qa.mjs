/** Actual pinned renderer refusal: unsupported inherited bold off leaves the document intact. */
import assert from 'node:assert/strict';
import {buildSync} from 'esbuild';
import {openApp,sleep} from './harness.mjs';
const root=process.cwd();
const bundle=buildSync({entryPoints:[root+'/src/engine/live-style-edit.ts'],bundle:true,platform:'browser',format:'iife',globalName:'__styleLive',write:false}).outputFiles[0].text;
const app=await openApp({name:'pr97-toggle-proof',port:Number(process.env.QA_PORT??9877),settle:1000,extra:`<script>${bundle.replaceAll('</script','<\\/script')}</script>`});
try{
const ask=async(code)=>JSON.parse(await app.js(`(async()=>JSON.stringify(await(${code})))()`));
const weight=()=>app.js(`getComputedStyle(document.querySelector('.editor-stack .superdoc-text-run')).fontWeight`);
const revision=()=>ask('window.__qa.doc().info({}).then(x=>x.revision)');
const create=async(xml,edit)=>ask(`__styleLive.applyLiveStyleEdit(window.__qa.sd(),${JSON.stringify(xml)},${JSON.stringify(edit)},${JSON.stringify(await revision())})`);
await app.caretPara('',{verifyMs:10000});await app.type('בדיקת הדגשה',30);
let files=await app.docx();const base=await create(files['word/styles.xml'],{kind:'create',name:'BoldBase',type:'paragraph',basedOn:'Normal',next:null,format:{bold:true}});assert.equal(base.outcome.ok,true);
const appliedBase=await ask(`window.__qa.ui().commands.executeAsync('linked-style',{style:${JSON.stringify(base.id)}})`);assert.ok(appliedBase===true||appliedBase.success===true);await sleep(1000);
console.log('Base',base.id,'weight',await weight());assert.equal(await weight(),'700');
files=await app.docx();const beforeRevision=await revision();const beforeHistory=await ask('window.__qa.doc().history.get()');const beforeText=await ask('window.__qa.doc().getText({})');
const derived=await create(files['word/styles.xml'],{kind:'create',name:'BoldDerived',type:'paragraph',basedOn:base.id,next:null,format:{bold:false}});
assert.equal(derived.outcome.ok,false);assert.equal(derived.outcome.reason,'unsupported-style-inheritance');
assert.equal(await revision(),beforeRevision);assert.equal(await weight(),'700');assert.deepEqual(await ask('window.__qa.doc().history.get()'),beforeHistory);assert.equal(await ask('window.__qa.doc().getText({})'),beforeText);
const catalog=await ask("window.__qa.doc().styles.getCatalog({view:'currentDocument'})");assert.ok(!catalog.items.some(s=>s.name==='BoldDerived'));
console.log('PASS: native Base bold700; unsupported Derived-off refused before mutation, revision/history/catalog/text unchanged');
}finally{app.close()}
