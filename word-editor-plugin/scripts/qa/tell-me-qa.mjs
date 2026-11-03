/**
 * שער ה-QA של חיפוש הפקודות („Tell Me”).
 *
 * הבאג שהוליד את השער: „השלמה מהספר” — כפתור גדול בלשונית „אוצריא” — החזירה
 * „לא נמצאו פקודות מתאימות”. הקטלוג (`ui/shell/tell-me-actions.ts`) מתוחזק
 * ביד, ולכן פקד שנוסף לרצועה אינו מגיע לחיפוש מעצמו: 55 מ-96 תוויות הרצועה
 * לא היו ניתנות למציאה.
 *
 * בדיקת היחידה סופרת את אותו כיסוי מול המקור. מה שהיא **אינה** יכולה לומר הוא
 * שהתיבה עצמה מציגה את מה שהקטלוג מחזיר, ושלחיצה על פריט אכן עושה משהו — וזה
 * מה שנמדד כאן, בדפדפן ובלחיצות אמיתיות:
 *   1. כל תווית של פקד ברצועה מופיעה ברשימה הנפתחת של התיבה.
 *   2. „השלמה מהספר” נלחצת ומדליקה את הכפתור שברצועה.
 *   3. פריט שכל תפקידו להוליך ללשונית — מחליף לשונית בפועל.
 *
 * הרצה:  node scripts/qa/tell-me-qa.mjs
 * היציאה 9633 שמורה לשער הזה בלבד.
 */
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { openApp, createReport, ROOT, sleep } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9633);
const report = createReport('חיפוש פקודות (Tell Me)', { strict: true });
const log = (...a) => console.log(...a);

/** `label="..."` בלבד — `:label="expr"` הוא ביטוי ולא תווית לחיפוש. */
const STATIC_LABEL = /(?<![:\w-])label="([^"]+)"/g;
const TABS_DIR = join(ROOT, 'src/ui/ribbon/tabs');

const RIBBON_LABELS = [
  ...new Set(
    readdirSync(TABS_DIR)
      .filter((f) => f.endsWith('.vue'))
      .flatMap((f) => [...readFileSync(join(TABS_DIR, f), 'utf8').matchAll(STATIC_LABEL)])
      .map((m) => m[1].replace(/[.…]+$/u, '').trim()),
  ),
];

/** מקליד שאילתה בתיבה ומחזיר את הפריטים שברשימה הנפתחת. */
async function ask(app, query) {
  await app.js(`(() => {
    const input = document.querySelector('.tell-me-input');
    if (!input) throw new Error('תיבת ה-Tell Me לא נמצאה');
    input.focus();
    input.value = ${JSON.stringify(query)};
    input.dispatchEvent(new Event('input', { bubbles: true }));
  })()`);
  await sleep(120);
  return app
    .js(`JSON.stringify([...document.querySelectorAll('.tell-me-dropdown .tell-me-item:not(.tell-me-item-find) .tell-me-item-title')].map((e) => e.textContent.trim()))`)
    .then(JSON.parse);
}

/**
 * סוגר את הרשימה הנפתחת.
 *
 * חובה לפני כל לחיצה ברצועה: התפריט הוא 380×420 שיושב מתחת לפס הכותרת, כלומר
 * **מעל** רצועת הלשוניות — ולחיצה על לשונית בזמן שהוא פתוח נוחתת עליו.
 */
async function closeBox(app) {
  await app.js(`(() => {
    const input = document.querySelector('.tell-me-input');
    if (!input) return;
    input.value = '';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    input.blur();
  })()`);
  await sleep(150);
}

/** לוחץ על פריט לפי הכותרת שלו ברשימה הנפתחת. */
async function clickResult(app, title) {
  const rect = await app
    .js(`(() => {
      const item = [...document.querySelectorAll('.tell-me-dropdown .tell-me-item')].find(
        (e) => (e.querySelector('.tell-me-item-title')?.textContent || '').trim() === ${JSON.stringify(title)},
      );
      if (!item) return 'null';
      const r = item.getBoundingClientRect();
      return JSON.stringify({ x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2) });
    })()`)
    .then(JSON.parse);
  if (!rect) return false;
  await app.clickAt(rect.x, rect.y);
  await sleep(400);
  return true;
}

async function step(name, fn) {
  log(`\n──────── ${name} ────────`);
  try {
    await fn();
  } catch (error) {
    log('!! זרק:', error?.message);
    report.fail(name, `הצעד זרק: ${error?.message}`);
  }
}

const app = await openApp({ name: 'tell-me', port: PORT });

try {
  await step('כל פקד ברצועה נמצא בחיפוש', async () => {
    const missing = [];
    for (const label of RIBBON_LABELS) {
      const titles = await ask(app, label);
      // מה שנמדד כאן הוא הסימפטום שדווח: „לא נמצאו פקודות מתאימות”. ההתאמה
      // המדויקת (כותרת או מילת מפתח זהה) נשמרת בבדיקת היחידה, שיש לה גישה
      // לקטלוג עצמו; לתיבה יש רק את מה שהיא מציגה.
      if (titles.length === 0) missing.push(label);
    }
    log(`נבדקו ${RIBBON_LABELS.length} תוויות; ללא תוצאה: ${missing.length}`);
    if (missing.length) {
      report.fail('כיסוי הרצועה', `אין תוצאה עבור: ${missing.join(', ')}`);
    } else {
      report.pass('כיסוי הרצועה', `כל ${RIBBON_LABELS.length} התוויות מחזירות תוצאה`);
    }
    await closeBox(app);
  });

  await step('„השלמה מהספר” — הבאג שדווח', async () => {
    const titles = await ask(app, 'השלמה מהספר');
    log('תוצאות:', titles.slice(0, 3).join(' | ') || '(אין)');
    if (titles[0] !== 'השלמה מהספר') {
      report.fail('השלמה מהספר', `הפריט אינו ראשון ברשימה (התקבל: ${titles[0] ?? 'אין תוצאות'})`);
      return;
    }
    report.pass('השלמה מהספר', 'הפריט מוחזר ראשון');

    await closeBox(app);
    await app.tab('✦ אוצריא');
    const before = await app.state('השלמה מהספר');
    await ask(app, 'השלמה מהספר');
    if (!(await clickResult(app, 'השלמה מהספר'))) {
      report.fail('לחיצה על „השלמה מהספר”', 'הפריט לא נמצא ללחיצה');
      return;
    }
    const after = await app.state('השלמה מהספר');
    log(`מצב הכפתור ברצועה: לפני active=${before.active}, אחרי active=${after.active}`);
    if (before.active === after.active) {
      report.fail('לחיצה על „השלמה מהספר”', 'מצב הכפתור ברצועה לא השתנה — הלחיצה לא עשתה דבר');
    } else {
      report.pass('לחיצה על „השלמה מהספר”', `הטוגל עבר ל-${after.active}`);
    }
  });

  await step('פריט שמוליך ללשונית מחליף לשונית', async () => {
    await closeBox(app);
    await app.tab('בית');
    await ask(app, 'גבולות עמוד');
    if (!(await clickResult(app, 'גבולות עמוד'))) {
      report.fail('ניווט ללשונית', '„גבולות עמוד” לא נמצא ברשימה');
      return;
    }
    const active = await app.js('window.__qa.activeTab()');
    log('לשונית פעילה אחרי הלחיצה:', active);
    if (active !== 'פריסה') {
      report.fail('ניווט ללשונית', `הלשונית הפעילה היא „${active}” ולא „פריסה”`);
    } else {
      report.pass('ניווט ללשונית', '„גבולות עמוד” פתח את „פריסה”');
    }
  });
} finally {
  app.close();
}

report.print();
