/**
 * מיקום ועקביות סימון התיקון האוטומטי מול המסך האמיתי: כמה מילים שהוחלפו
 * (כולל שתי זהות), והסימון עדיין זמין אחרי שממשיכים להקליד. הפס חייב לשבת מתחת
 * לאות הראשונה של **המילה הנכונה**.
 */
import { openApp, createReport } from './harness.mjs';
import { loadHebrewPreset } from './autocorrect-helpers.mjs';
const report = createReport('גיאומטריית סימון', { strict: true });
const app = await openApp({ name: 'acorr-geo', port: Number(process.env.QA_PORT ?? 9632) });

/** כל המופעים המצוירים של מחרוזת, בסדר ה-DOM. */
const probeAll = (needle) =>
  app
    .js(`(function(){
    var walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    var out = [], n;
    while ((n = walker.nextNode())) {
      var p = n.parentElement;
      if (!p || !p.closest('[data-source-node-id][data-pm-start]')) continue;
      var from = 0, i;
      while ((i = n.nodeValue.indexOf(${JSON.stringify(needle)}, from)) >= 0) {
        var r = document.createRange(); r.setStart(n, i); r.setEnd(n, i + ${JSON.stringify(needle)}.length);
        var b = r.getBoundingClientRect();
        var r1 = document.createRange(); r1.setStart(n, i); r1.setEnd(n, i + 1);
        var f = r1.getBoundingClientRect();
        if (b.width > 0) out.push({ left: b.left, right: b.right, top: b.top, bottom: b.bottom, fl: f.left, fr: f.right });
        from = i + 1;
      }
    }
    return JSON.stringify(out);
  })()`)
    .then(JSON.parse);

const chipBox = () =>
  app
    .js(`(function(){var c=document.querySelector('.otzaria-acorr-chip'); if(!c||c.style.display!=='block') return 'null'; var r=c.getBoundingClientRect(); return JSON.stringify({left:r.left,right:r.right,top:r.top,bottom:r.bottom});})()`)
    .then(JSON.parse);

/** מרחף מרכז המילה, ומחזיר את הפס שהופיע. */
async function hover(word) {
  const y = (word.top + word.bottom) / 2;
  const x = (word.left + word.right) / 2;
  // תנועה מחוץ למילה קודם, כדי שמצב קודם לא ישפיע
  await app.js(`(document.elementFromPoint(${word.left - 40},${y}) || document.body).dispatchEvent(new PointerEvent('pointermove',{bubbles:true,clientX:${word.left - 40},clientY:${y}}))`);
  await app.sleep(60);
  await app.js(`(document.elementFromPoint(${x},${y}) || document.body).dispatchEvent(new PointerEvent('pointermove',{bubbles:true,clientX:${x},clientY:${y}}))`);
  await app.sleep(120);
  return chipBox();
}

function check(name, word, chip) {
  if (!chip) return report.fail(name, 'הסימון לא הופיע בריחוף על המילה');
  // הפס מתחת לאות הראשונה: מרכז הפס קרוב למרכז האות הראשונה (בעברית — הימנית)
  const dx = Math.round((chip.left + chip.right) / 2 - (word.fl + word.fr) / 2);
  const dy = Math.round(chip.top - word.bottom);
  const ok = Math.abs(dx) <= 4 && Math.abs(dy) <= 4;
  (ok ? report.pass : report.fail)(name, `dx=${dx} dy=${dy}`);
}

try {
  await loadHebrewPreset(app);
  await app.js("window.otzariaAutocorrectDebug = true; window.__logs = []; var o = console.info; console.info = function(){ window.__logs.push(Array.prototype.slice.call(arguments).map(function(a){return typeof a==='string'?a:JSON.stringify(a)}).join(' ')); o.apply(console, arguments); }");
  await app.caret(0);
  // שתי החלפות זהות ועוד שתיים, ואז המשך הקלדה
  await app.type('אאכ אאכ (c) אאכ ');
  await app.sleep(1500);

  console.log('לוגים:', await app.js("JSON.stringify(window.__logs.filter(function(l){return /סימון|הוחלף/.test(l)}))"));
  console.log('DOM:', await app.js("(function(){var f=document.querySelector('[data-source-node-id][data-pm-start]');return f.getAttribute('data-source-node-id')+' '+f.getAttribute('data-pm-start')+'-'+f.getAttribute('data-pm-end')+' | '+Array.from(f.querySelectorAll('[data-pm-start]')).map(function(r){return r.getAttribute('data-pm-start')+':'+JSON.stringify(r.textContent)}).join(' ')})()"));
  const hebrew = await probeAll('אא"כ');
  const copyright = await probeAll('©');
  console.log('מופעים: אא"כ =', hebrew.length, ' © =', copyright.length);
  hebrew.length === 3 ? report.pass('שלוש החלפות של אא"כ בדף') : report.fail('שלוש החלפות של אא"כ בדף', String(hebrew.length));

  // כל ההחלפות באותה הערכה מסומנות לפני העריכה הבאה.
  for (const [i, word] of hebrew.entries()) check(`אא"כ מס' ${i + 1}`, word, await hover(word));
  for (const [i, word] of copyright.entries()) check(`© מס' ${i + 1}`, word, await hover(word));

  await app.type('ועוד טקסט שהוקלד אחרי ');
  await app.sleep(800);
  // טקסט רגיל שלא הוחלף: אין סימון
  const plain = await probeAll('ועוד');
  if (plain[0]) {
    const chip = await hover(plain[0]);
    chip ? report.fail('מילה שלא הוחלפה אינה מסומנת', JSON.stringify(chip)) : report.pass('מילה שלא הוחלפה אינה מסומנת');
  }
} finally {
  report.print();
  app.close();
}
