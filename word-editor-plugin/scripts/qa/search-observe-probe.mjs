/** מה בדיוק `observe` פולט סביב חיפוש, ומתי — כדי לבנות המתנה על אות אמיתי. */
import { openApp } from './harness.mjs';
const app = await openApp({ name: 'observe', port: Number(process.env.QA_PORT ?? 9495) });
try {
  await app.caret(0);
  await app.type(('zzq ').repeat(8).trim());
  await app.sleep(1200);
  const out = await app.js(`(async function(){
    var E = window.__otzariaEditor, sd = E.superdoc;
    var s = (E.ui && E.ui.search) || (sd.ui && sd.ui.search);
    var sleep = function(ms){ return new Promise(function(r){ setTimeout(r, ms); }); };
    var t0 = performance.now(), ev = [];
    s.open && s.open();
    var off = s.observe(function(v){ ev.push({ t: Math.round(performance.now()-t0), total: v && v.total, q: v && v.query }); });
    var immediate = ev.length;
    var ret = (typeof s.find === 'function') ? s.find('zzq') : s.search('zzq');
    ev.push({ t: Math.round(performance.now()-t0), mark: 'search-returned', total: ret && ret.total });
    await sleep(2500);
    off && off();
    return JSON.stringify({ emitsBeforeSearch: immediate, events: ev });
  })()`);
  console.log(out);
} finally { app.close(); }
