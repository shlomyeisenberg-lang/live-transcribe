/** המחוונים נשמרים ללא תפוגה, אך עריכה חדשה פוסלת את זהותם הישנה. */
import { openApp, createReport } from './harness.mjs';
import { loadHebrewPreset } from './autocorrect-helpers.mjs';
const report = createReport('סימון נשאר על כל מילה', { strict: true });
const app = await openApp({ name: 'acorr-persist', port: Number(process.env.QA_PORT ?? 9633) });

/** מרכז המופע ה-n של מחרוזת בדף, או null. */
const wordBox = (needle, nth) =>
  app
    .js(`(function(){
    var w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT), n, k = 0;
    while ((n = w.nextNode())) {
      if (!n.parentElement || !n.parentElement.closest('[data-source-node-id][data-pm-start]')) continue;
      var from = 0, i;
      while ((i = n.nodeValue.indexOf(${JSON.stringify(needle)}, from)) >= 0) {
        if (k++ === ${nth}) {
          var r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + ${JSON.stringify(needle)}.length);
          var b = r.getBoundingClientRect();
          var r1 = document.createRange(); r1.setStart(n, i); r1.setEnd(n, i + 1);
          var f = r1.getBoundingClientRect();
          return JSON.stringify({ x: (b.left + b.right) / 2, y: (b.top + b.bottom) / 2, fl: f.left, fr: f.right, bottom: b.bottom });
        }
        from = i + 1;
      }
    }
    return 'null';
  })()`)
    .then(JSON.parse);

/** מרחף מעל המילה ובודק שהפס מתחת לאות הראשונה שלה. */
async function marked(needle, nth) {
  const box = await wordBox(needle, nth);
  if (!box) return 'אין מילה בדף';
  await app.js(`(document.elementFromPoint(${box.x},${box.y})||document).dispatchEvent(new PointerEvent('pointermove',{bubbles:true,clientX:${box.x},clientY:${box.y}}))`);
  await app.sleep(120);
  const chip = await app.js(`(function(){var c=document.querySelector('.otzaria-acorr-chip'); if(!c||c.style.display!=='block') return 'null'; var r=c.getBoundingClientRect(); return JSON.stringify({l:r.left,r:r.right,t:r.top});})()`).then(JSON.parse);
  if (!chip) return 'אין סימון';
  const dx = Math.abs(chip.l - box.fl) + Math.abs(chip.r - box.fr);
  return dx < 3 && Math.abs(chip.t - (box.bottom - 2)) < 3 ? 'ok' : `סימון במקום שגוי dx=${dx.toFixed(1)}`;
}

try {
  await loadHebrewPreset(app);
  await app.caret(0);
  await app.type('(tm) שלום (c) ועוד אאכ ');
  await app.sleep(1200);

  for (const [w, n, label] of [['™', 0, 'המילה הראשונה'], ['©', 0, 'המילה האמצעית'], ['אא"כ', 0, 'המילה האחרונה']]) {
    const r = await marked(w, n);
    r === 'ok' ? report.pass(`${label} מסומנת אחרי התיקון`) : report.fail(`${label} מסומנת אחרי התיקון`, r);
  }

  await app.sleep(20_000);
  const late = await marked('™', 0);
  late === 'ok' ? report.pass('הסימון נשאר אחרי 20 שניות') : report.fail('הסימון נשאר אחרי 20 שניות', late);

  // עריכה חדשה פוסלת את המחוונים הישנים: אין ניחוש של זהות אחרי שינוי טקסט.
  await app.press('Home', 'Home', 36);
  await app.sleep(400);
  await app.type('הקדמה ארוכה ');
  await app.sleep(1200);
  const text = await app.js("document.querySelector('[data-source-node-id][data-pm-start]').textContent");
  text.startsWith('הקדמה') ? report.pass('ההקלדה נכנסה לתחילת הפסקה', text) : report.fail('ההקלדה נכנסה לתחילת הפסקה', text);
  for (const [w, n, label] of [['™', 0, 'הראשונה'], ['©', 0, 'האמצעית'], ['אא"כ', 0, 'האחרונה']]) {
    const r = await marked(w, n);
    r === 'אין סימון' ? report.pass(`עריכה חדשה פוסלת את המחוון של ${label}`) : report.fail(`עריכה חדשה פוסלת את המחוון של ${label}`, r);
  }
} finally {
  report.print();
  app.close();
}
