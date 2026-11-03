/**
 * "החלף הכל" על מסמך רב-פסקאות — דרך הדיאלוג האמיתי, כפי שהמשתמש עושה זאת.
 *
 * זה בדיוק התרחיש שהיה שבור: `ui.search` (המנוע) נמדד מוצא רק חלק מהמופעים
 * במסמך עם כמה פסקאות (8 מופעים, 4 נמצאים — "1 מתוך 4" בדיאלוג), עם שגיאת
 * אבחון מהמנוע עצמו: `projection-incomplete: exact-complete projection did
 * not cover the full document`. ראו docs/superdoc-2.10-review.md ו-
 * docs/button-audit.md.
 *
 * `replace-all-probe.mjs`/`replace-ui-probe.mjs` הקיימים בודקים מסמך של
 * פסקה **אחת** עם שמונה מופעים לצידה (מופרדים ברווח) — התרחיש הזה נמדד
 * שעבד גם לפני התיקון. השער הזה בודק את התרחיש שבאמת נכשל: פסקה נפרדת
 * לכל מופע (Enter בין כל אחת), שבו `ui.search` איבד את הרוב.
 *
 * המימוש שנבדק כאן (`engine/search.ts`) אינו קורא ל-`ui.search` בכלל —
 * הוא קורא ישירות ל-`doc.blocks.list`/`doc.replace` הציבוריים. השער מאמת
 * גם מול ה-docx המיוצא וגם מול הטקסט על המסך.
 */
import { openApp, createReport } from './harness.mjs';

const report = createReport('החלף הכל — מסמך רב-פסקאות', { strict: true });
const app = await openApp({ name: 'replmultipara', port: Number(process.env.QA_PORT ?? 9499) });

const NEEDLE = 'zzq';
const COUNT = 8;
const count = (s, needle) => s.split(needle).length - 1;

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1600,
    height: 1000,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await app.sleep(400);

  await app.caret(0);
  for (let i = 0; i < COUNT; i++) {
    await app.type(`${NEEDLE} paragraph ${i}`);
    if (i < COUNT - 1) {
      await app.press('Enter', 'Enter', 13);
      await app.sleep(150);
    }
  }
  await app.sleep(1200);

  let files = await app.docx();
  let doc = files['word/document.xml'] || '';
  const before = count(doc, NEEDLE);
  console.log(`מופעים לפני ב-docx: ${before} (${COUNT} פסקאות נפרדות, אחד בכל אחת)`);
  if (before !== COUNT) {
    report.fail('הכנת הבדיקה', `נמצאו ${before} מופעים במקום ${COUNT}`);
  }

  // הדיאלוג האמיתי: הרצועה, לא שינוי ישיר של state.
  await app.tab('בית');
  const opened = await app.click('החלפה');
  if (!opened) throw new Error('כפתור "החלפה" לא נמצא ברצועה');
  await app.sleep(600);

  const fill = async (sel, value) => app.js(`(function(){
    var el = document.querySelector('${sel}');
    if (!el) return 'no-el';
    var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    setter.call(el, ${JSON.stringify(value)});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return 'ok';
  })()`);

  console.log('שדה חיפוש:', await fill('#fr-search-input', NEEDLE));
  console.log('שדה החלפה:', await fill('#fr-replace-input', 'YYY'));
  // השקטת החיפוש-בזמן-הקלדה (SEARCH_DEBOUNCE_MS) + קריאת blocks.list.
  await app.sleep(900);

  const counter = await app.js(`(function(){
    var e = document.querySelector('.fr-counter');
    return e ? e.textContent.trim() : null;
  })()`);
  console.log('מונה בדיאלוג:', counter);
  if (counter !== `${COUNT} תוצאות` && !new RegExp(`מתוך ${COUNT}$`).test(counter ?? '')) {
    report.fail('מונה הדיאלוג', `ציפינו למונה שמדבר על ${COUNT}, קיבלנו: ${counter}`);
  } else {
    report.pass('מונה הדיאלוג', `${counter}`);
  }

  const btns = JSON.parse(
    await app.js(
      `JSON.stringify(Array.from(document.querySelectorAll('.find-replace-dialog .fr-btn')).map(x=>x.textContent.trim()))`,
    ),
  );
  const idx = btns.indexOf('החלף הכל');
  if (idx < 0) throw new Error(`כפתור "החלף הכל" לא נמצא. כפתורים: ${JSON.stringify(btns)}`);
  const rect = JSON.parse(
    await app.js(
      `(function(){var el=document.querySelectorAll('.find-replace-dialog .fr-btn')[${idx}];if(!el)return 'null';var r=el.getBoundingClientRect();return JSON.stringify({x:r.x+r.width/2,y:r.y+r.height/2});})()`,
    ),
  );
  await app.clickAt(rect.x, rect.y);
  await app.sleep(1500);

  const status = await app.status();
  console.log('שורת מצב אחרי הלחיצה:', JSON.stringify(status));

  files = await app.docx();
  doc = files['word/document.xml'] || '';
  const left = count(doc, NEEDLE);
  const replaced = count(doc, 'YYY');
  const screen = (await app.screenText()) || '';
  console.log(`אחרי, ב-docx: נשארו ${left} מופעים של "${NEEDLE}", נכתבו ${replaced} מופעים של "YYY"`);
  console.log('על המסך מכיל YYY פעמים:', count(screen, 'YYY'), '| מכיל zzq פעמים:', count(screen, NEEDLE));

  if (left === 0 && replaced === COUNT) {
    report.pass('החלף הכל — מסמך רב-פסקאות', `כל ${COUNT} המופעים הוחלפו ב-docx המיוצא`);
  } else {
    report.fail('החלף הכל — מסמך רב-פסקאות', `${before} → נשארו ${left}, הוחלפו ${replaced} מתוך ${COUNT}`);
  }

  if (count(screen, 'YYY') === COUNT && count(screen, NEEDLE) === 0) {
    report.pass('החלף הכל — תצוגה על המסך', 'כל השמונה מוחלפים גם בתצוגה');
  } else {
    report.fail(
      'החלף הכל — תצוגה על המסך',
      `YYY מופיע ${count(screen, 'YYY')} פעמים, ${NEEDLE} מופיע ${count(screen, NEEDLE)} פעמים`,
    );
  }

  console.log('לוג הדף:', JSON.stringify(await app.log()).slice(0, 500));
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
