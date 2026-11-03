/**
 * בדיקה ממוקדת ל„החלף הכל”: כמה מופעים באמת מוחלפים, ומה מדווח.
 * נמדד מול ה-docx המיוצא, לא מול הדיווח של המנוע.
 */
import { openApp, createReport } from './harness.mjs';

const report = createReport('החלף הכל', { strict: true });
const app = await openApp({ name: 'replaceall', port: Number(process.env.QA_PORT ?? 9470) });

const count = (s, needle) => s.split(needle).length - 1;

try {
  await app.caret(0);
  const NEEDLE = 'zzq';
  const line = (NEEDLE + ' ').repeat(8).trim();
  await app.type(line);
  await app.sleep(1200);

  let files = await app.docx();
  let doc = files['word/document.xml'] || '';
  const before = count(doc, NEEDLE);
  console.log('מופעים לפני ההחלפה ב-docx:', before);
  console.log('טקסט על המסך:', JSON.stringify((await app.screenText() || '').slice(0, 200)));

  const api = await app.js(`(function(){
    var E = window.__otzariaEditor || {};
    var sd = E.superdoc;
    var cands = {
      'E.ui.search': E.ui && E.ui.search,
      'sd.activeEditor.ui.search': sd && sd.activeEditor && sd.activeEditor.ui && sd.activeEditor.ui.search,
      'sd.ui.search': sd && sd.ui && sd.ui.search
    };
    var out = {};
    for (var k in cands) {
      var s = cands[k];
      out[k] = s ? { keys: Object.keys(s), hasFind: typeof s.find === 'function', hasSearch: typeof s.search === 'function', hasReplaceAll: typeof s.replaceAll === 'function' } : null;
    }
    return JSON.stringify(out);
  })()`);
  console.log('משטח החיפוש:', api);

  const res = await app.js(`(async function(){
    var E = window.__otzariaEditor;
    var sd = E.superdoc;
    var s = (E.ui && E.ui.search) || (sd.ui && sd.ui.search);
    if (!s) return JSON.stringify({ error: 'no search surface' });
    var sleep = function(ms){ return new Promise(function(r){ setTimeout(r, ms); }); };
    s.open && s.open();
    var first = (typeof s.find === 'function') ? s.find('${NEEDLE}') : s.search('${NEEDLE}');
    var trail = [ first && first.total ];
    // ההמתנה להתייצבות: התוצאה מגיעה מ-worker ומתפרסמת מאוחר יותר
    for (var i = 0; i < 40; i++) {
      await sleep(100);
      var snap = s.getSnapshot();
      trail.push(snap.total);
      if (snap.total > 0 && trail[trail.length-1] === trail[trail.length-2]) break;
    }
    var settled = s.getSnapshot().total;
    var r = await s.replaceAll('YYY');
    await sleep(600);
    return JSON.stringify({ firstTotal: first && first.total, settledTotal: settled, trail: trail, replaceAll: r, after: s.getSnapshot() });
  })()`);
  console.log('תוצאת ההחלפה:', res);

  await app.sleep(1500);
  files = await app.docx();
  doc = files['word/document.xml'] || '';
  const left = count(doc, NEEDLE);
  const put = count(doc, 'YYY');
  console.log(`אחרי: נשארו ${left} מופעים של "${NEEDLE}", נכתבו ${put} מופעים של "YYY"`);
  console.log('טקסט על המסך אחרי:', JSON.stringify((await app.screenText() || '').slice(0, 200)));

  if (before !== 8) report.fail('הכנת הבדיקה', `נמצאו ${before} מופעים במקום 8`);
  else if (left === 0 && put === 8) report.pass('החלף הכל', 'כל 8 המופעים הוחלפו');
  else report.fail('החלף הכל', `נשארו ${left}, הוחלפו ${put} מתוך 8`);

  console.log('לוג הדף:', JSON.stringify(await app.log()));
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
