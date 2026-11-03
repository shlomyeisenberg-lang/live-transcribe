/** האם אפשר לתת לרשימה מספור עברי, ומה מצויר על המסך. */
import { openApp, createReport } from './harness.mjs';

const report = createReport('מספור רשימה עברי', { strict: true });
const app = await openApp({ name: 'heblist', port: Number(process.env.QA_PORT ?? 9492) });

try {
  await app.caret(0);
  await app.type('aleph');
  await app.sleep(700);
  await app.tab('בית');
  await app.click('מספור');
  await app.sleep(1200);

  const listing = await app.js(`(async function(){
    var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    var l = await doc.lists.list();
    return JSON.stringify(l);
  })()`);
  console.log('רשימות אחרי לחיצה על „מספור":', String(listing).slice(0, 500));

  const setRes = await app.js(`(async function(){
    var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    var l = await doc.lists.list();
    var items = (l && l.items) || [];
    if (!items.length) return JSON.stringify({ err: 'no lists', listing: l });
    var it = items[0];
    var out = { item: it };
    for (var style of ['hebrew1','hebrew2']) {
      try {
        out[style] = await doc.lists.setLevelNumberStyle({ target: it.address, level: 0, numberStyle: style });
      } catch (e) { out[style + 'Err'] = String(e && e.message || e); }
    }
    return JSON.stringify(out);
  })()`);
  console.log('setLevelNumberStyle:', String(setRes).slice(0, 900));

  await app.sleep(1200);
  const files = await app.docx();
  const numbering = files['word/numbering.xml'] || '';
  const fmts = [...numbering.matchAll(/<w:numFmt w:val="([^"]+)"/g)].map((x) => x[1]);
  const heb = fmts.filter((f) => f.startsWith('hebrew'));
  console.log('numFmt ב-numbering.xml:', JSON.stringify(fmts));
  console.log('מסך:', JSON.stringify((await app.screenText() || '').slice(0, 200)));
  const marker = await app.js(`(function(){
    var els = document.querySelectorAll('.superdoc-list-marker, [class*="list-marker"], [class*="marker"]');
    return JSON.stringify(Array.prototype.slice.call(els, 0, 6).map(function(e){ return e.textContent; }));
  })()`);
  console.log('סמני רשימה על המסך:', marker);

  heb.length ? report.pass('מספור עברי לרשימה', heb.join(',')) : report.fail('מספור עברי לרשימה', fmts.join(',') || 'אין');
  console.log('לוג:', JSON.stringify(await app.log()));
} finally {
  app.close();
}
process.exit(report.print() > 0 ? 1 : 0);
