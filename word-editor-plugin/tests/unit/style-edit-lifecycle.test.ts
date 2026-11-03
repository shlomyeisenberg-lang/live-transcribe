/** Exercise the real shell function across awaited operations, not a duplicate of its logic. */
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import ts from 'typescript';
import { describe, it, expect } from 'vitest';

const source = readFileSync('src/App.vue', 'utf8');
const start = source.indexOf('const styleEditsInFlight =');
const end = source.indexOf('provide(STYLE_DOCUMENT, {', start);
const javascript = ts.transpileModule(source.slice(start, end), {compilerOptions:{target:ts.ScriptTarget.ES2022}}).outputText;

function deferred<T>() {
  let resolve!: (value: T) => void;
  const promise = new Promise<T>(r => { resolve = r; });
  return {promise, resolve};
}

function setup() {
  const A = { id:'A', pendingRestore:false, swap:{current:null as unknown}, keeper:{state:{}}, text: 'A unsaved', dirty: false, save: {snapshot:{isSaving:false}, hold: () => () => {}, markDirty: () => { A.dirty = true; }} };
  const B = { id:'B', pendingRestore:false, swap:{current:null as unknown}, keeper:{state:{}}, text: 'B unsaved important', dirty: false, save: {snapshot:{isSaving:false}, hold: () => () => {}, markDirty: () => { B.dirty = true; }} };
  const activeSession = {value:A};
  let target={blockId:'first',offset:0};
  const current = {superdoc:A,ui:{selection:{getSnapshot:()=>({selectionTarget:target})}}};
  const swap = {current};
  A.swap.current=current; B.swap.current={superdoc:B,ui:current.ui};
  const exports = deferred<string>();
  const edits: unknown[] = [];
  const env = {activeSession, swap, console, Set, isOpenBusy:()=>false, HEAVY_ACTION_HOLD_MAX_MS:60000,
    styleRevision:async()=> '12', exportDocx:()=>exports.promise, readDocxStyles:async()=>'<styles/>',
    applyLiveStyleEdit:async(host:unknown,xml:string,edit:unknown)=>{edits.push({host,xml,edit});return {outcome:{ok:true},id:'Normal'};},
    sessions:new Map([['A',A],['B',B]]), activeEntry:()=>({document:{}}), commandAdapter:{value:null as {run:(id:string,payload:unknown)=>Promise<{ok:boolean}>}|null}, stylePayload:(id:string)=>id};
  const context = vm.createContext(env);
  vm.runInContext(javascript, context);
  for(const [first,last] of [['function sleepCandidate(', '/**\n * סינון מוקדם'], ['async function onDocumentTabClose(', '/** כפתור „+”']]) {
    const from=source.indexOf(first);const to=source.indexOf(last,from);
    vm.runInContext(ts.transpileModule(source.slice(from,to),{compilerOptions:{target:ts.ScriptTarget.ES2022}}).outputText,context);
  }
  const apply=(edit={kind:'modify',id:'Normal',format:{fontSizePt:18}},options={})=>vm.runInContext('applyStyleEdit',context)(edit,options);
  return {A,B,env,exports,edits,apply,moveSelection(){target={blockId:'second',offset:0};}, canSleep:()=>vm.runInContext('sleepCandidate(activeSession.value, false)',context), closeA:()=>vm.runInContext('onDocumentTabClose("A")',context),switchToB(){activeSession.value=B;swap.current={superdoc:B,ui:current.ui};}};
}

describe('style edit session lifecycle',()=>{
  it('does not mutate either document if a tab changes during export',async()=>{
    const s=setup();const pending=s.apply();await Promise.resolve();s.switchToB();s.exports.resolve('snapshotA');
    expect(await pending).toMatchObject({ok:false,reason:'superseded'});
    expect(s.edits).toEqual([]);expect(s.B.text).toBe('B unsaved important');expect(s.B.dirty).toBe(false);
  });

  it('marks only the originating session when a tab changes while its template transaction runs',async()=>{
    const s=setup();const reached=deferred<void>();const finished=deferred<{outcome:{ok:boolean},id:string}>();
    s.env.applyLiveStyleEdit=async(host,xml,edit)=>{s.edits.push({host,xml,edit});reached.resolve();return finished.promise;};
    const pending=s.apply();s.exports.resolve('snapshotA');await reached.promise;s.switchToB();
    finished.resolve({outcome:{ok:true},id:'Normal'});expect(await pending).toMatchObject({ok:true});
    expect(s.A.dirty).toBe(true);expect(s.B.dirty).toBe(false);expect(s.B.text).toBe('B unsaved important');
  });

  it('never replaces a body snapshot, preserving edits made after export starts',async()=>{
    const s=setup();const pending=s.apply();await Promise.resolve();s.A.text+=' NEW TEXT';s.exports.resolve('old snapshot');
    expect(await pending).toMatchObject({ok:true});expect(s.A.text).toBe('A unsaved NEW TEXT');
  });

  it('rejects concurrent operations on the same session without another export',async()=>{
    const s=setup();const first=s.apply();expect(await s.apply()).toMatchObject({ok:false,reason:'not-ready'});
    s.exports.resolve('snapshot');expect(await first).toMatchObject({ok:true});expect(s.edits).toHaveLength(1);
  });
  it('keeps an in-flight session alive when background trimming or closing is requested',async()=>{
    const s=setup();const pending=s.apply();
    expect(s.canSleep()).toMatchObject({isOpening:true});
    await s.closeA();
    expect(s.env.sessions.has('A')).toBe(true);
    s.exports.resolve('snapshot');expect(await pending).toMatchObject({ok:true});
    expect(s.canSleep()).toMatchObject({isOpening:false});
  });

  it('does not apply a new style to a selection moved while the source is exported',async()=>{
    const s=setup();const pending=s.apply(undefined,{applyToSelection:true});await Promise.resolve();s.moveSelection();s.exports.resolve('snapshot');
    expect(await pending).toMatchObject({ok:false,reason:'superseded'});expect(s.edits).toHaveLength(0);
  });

});
