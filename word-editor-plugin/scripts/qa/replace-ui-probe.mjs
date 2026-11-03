/** „החלף הכל" דרך הדיאלוג האמיתי, כפי שהמשתמש עושה זאת. */
import { openApp, createReport } from './harness.mjs';

const report = createReport('החלף הכל — דרך הממשק', { strict: true });
const app = await openApp({ name: 'replui', port: Number(process.env.QA_PORT ?? 9497) });
const NEEDLE = 'zzq';
const texts = (x) => (x.match(/<w:t[^>]*>[^<]*<\/w:t>/g) ?? []).join('');
const count = (s, n) => (s.split(n).length - 1);

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1600, height: 1000, deviceScaleFactor: 1, mobile: false });
  await app.sleep(400);
  await app.caret(0);
  await app.type((NEEDLE + ' ').repeat(8).trim());
  await app.sleep(1500);

  await app.tab('בית');
  await app.click('החלפה');
  await app.sleep(900);

  const fill = async (sel, value) => app.js(`(function(){
    var el = document.querySelector('${sel}');
    if (!el) return 'no-el';
    var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    setter.call(el, ${JSON.stringify(value)});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return 'ok';
  })()`);

  console.log('שדה החלפה מראש:', await fill('#fr-replace-input', 'YYY'));
  console.log('שדה חיפוש:', await fill('#fr-search-input', NEEDLE));
  await app.sleep(Number(process.env.WAIT_MS ?? 1800));
  console.log('שדה החלפה:', await fill('#fr-replace-input', 'YYY'));
  await app.sleep(300);

  const before = texts((await app.docx())['word/document.xml'] || '');
  const b = count(before, NEEDLE);
  const counter = await app.js(`(function(){var e=document.querySelector('.fr-counter,.find-replace-dialog .fr-count');return e?e.textContent.trim():null;})()`);
  console.log(`מופעים לפני: ${b} | מונה בדיאלוג: ${counter}`);

  const btns = JSON.parse(await app.js(`JSON.stringify(Array.from(document.querySelectorAll('.find-replace-dialog .fr-btn')).map(x=>x.textContent.trim()))`));
  console.log('כפתורים:', JSON.stringify(btns));
  const idx = btns.indexOf('החלף הכל');
  const rect = JSON.parse(await app.js(`(function(){var el=document.querySelectorAll('.find-replace-dialog .fr-btn')[${idx}];if(!el)return 'null';var r=el.getBoundingClientRect();return JSON.stringify({x:r.x+r.width/2,y:r.y+r.height/2});})()`));
  await app.clickAt(rect.x, rect.y);
  await app.sleep(3500);

  const after = texts((await app.docx())['word/document.xml'] || '');
  const a = count(after, NEEDLE), q = count(after, 'YYY');
  console.log(`אחרי: נשארו ${a}, הוחלפו ${q} מתוך ${b}`);
  console.log('שורת מצב:', JSON.stringify(await app.status()));
  console.log('לוג:', JSON.stringify(await app.log()).slice(0, 400));

  (a === 0 && q === b && b === 8)
    ? report.pass('החלף הכל דרך הממשק', `כל ${b} המופעים הוחלפו`)
    : report.fail('החלף הכל דרך הממשק', `${b} → נשארו ${a}, הוחלפו ${q}`);
} finally {
  app.close();
}
process.exit(report.print() > 0 ? 1 : 0);
