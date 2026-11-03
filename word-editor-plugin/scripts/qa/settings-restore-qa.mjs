/**
 * שער „העדפה שנשמרה חוזרת” — המסלול היחיד שאף שער אחר אינו מודד.
 *
 * ## מה נמדד כאן, ולמה לא נמדד עד כה
 *
 * לתוסף יש שורת העדפות שנקראות **בעלייה**: שמירה אוטומטית, הסרגל, בדיקת
 * האיות, „המסמכים האחרונים”, וזיכרון האפשרויות של הדיאלוגים. כולן עוברות
 * ב-`storage.get` של המאחז, וכולן נכתבות במסלול נפרד מזה שקורא אותן.
 *
 * שערים קיימים מודדים את **הכתיבה**: הם לוחצים על המתג, רואים שהמצב התחלף,
 * ובודקים ש-`storage.set` הגיע למאחז (`file-ops-qa.mjs`, שלב `autosave`).
 * אף אחד לא מדד את הכיוון ההפוך — „ההעדפה כבר שם, האם התוסף מכבד אותה” —
 * ולכן `host-stub.js` יכול היה להחזיר את הערך **עטוף** ב-`{ value }` לאורך
 * ~60 שערים בלי ששורה אחת תאדים: כל העדפה נדחתה בנרמול, נפלה לברירת המחדל,
 * ואיש לא ביקש ממנה דבר אחר.
 *
 * ## המבנה: שני חוזי חוט ואחריהם חמש התנהגויות
 *
 * שתי השורות הראשונות מודדות את **הגשר** ישירות — מה שנכתב חוזר כמו שהוא,
 * ומפתח שאינו קיים חוזר כ-`null`. הן הקנרית: כל צרכן, מרשני כמו
 * `ShulchanTab.vue` (`raw => raw`) או קפדני כמו `loadRulerVisible`
 * (`=== true`), נשען עליהן.
 *
 * חמש השורות שאחריהן מודדות את **המוצר**: הן זורעות העדפה לפני שהתוסף קם,
 * ובודקות מה המשתמש רואה. חוזה בלי התנהגות אינו אומר שההעדפה מגיעה למסך,
 * והתנהגות בלי חוזה אינה אומרת איפה זה נשבר.
 *
 * הזריעה נכנסת דרך `extra` של המסגרת — כלומר **אחרי** `host-stub.js` ולפני
 * הבאנדל, בדיוק החלון שבו אוצריא האמיתית כבר מחזיקה את ההעדפות של המשתמש
 * והתוסף עוד לא נשאל.
 *
 * כל שלב בדפדפן משלו: העדפה נקראת פעם אחת בעלייה, ולכן אין דרך לזרוע שנייה
 * באותה הרצה.
 *
 *   node scripts/qa/settings-restore-qa.mjs              # הכול
 *   node scripts/qa/settings-restore-qa.mjs wire ruler   # שלבים נבחרים
 */
import { openApp, createReport, sleep } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9645);
const report = createReport('שער שחזור העדפות מאחסון', { strict: true });

function withTimeout(promise, ms, label) {
  let timer;
  return Promise.race([
    promise.finally(() => clearTimeout(timer)),
    new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error(`פג הזמן: ${label}`)), ms);
    }),
  ]);
}

/**
 * זורעת העדפות ל-`__qaHost.storage` לפני שהבאנדל נטען.
 *
 * הזריעה היא **השמה ישירה**, ולא `storage.set` דרך הגשר, משתי סיבות: כך
 * נראית אוצריא שכבר מחזיקה את ההעדפה מהפעלה קודמת, וכך הזריעה אינה תלויה
 * בתקינות מסלול הכתיבה שהיא באה למדוד את הקריאה שלו.
 */
function seed(values) {
  return `
<script>
(function () {
  var H = window.__qaHost;
  if (!H) throw new Error('host-stub.js לא נטען לפני הזריעה');
  var v = ${JSON.stringify(values)};
  Object.keys(v).forEach(function (k) { H.storage[k] = v[k]; });
})();
</script>
`;
}

async function stage(name, values, body, { budgetMs = 180_000 } = {}) {
  const wanted = process.argv.slice(2);
  if (wanted.length && !wanted.includes(name)) return;

  console.log(`\n────── ${name} ──────`);
  console.log(`  נזרע: ${JSON.stringify(values)}`);
  let app = null;
  try {
    app = await openApp({ name: `settings-${name}`, port: PORT, extra: seed(values) });
    await app.cdp.send('Emulation.setDeviceMetricsOverride', {
      width: 1600, height: 1000, deviceScaleFactor: 1, mobile: false,
    });
    await sleep(400);
    await withTimeout(body(app), budgetMs, name);
  } catch (error) {
    console.log(`  ✗ ${error.message}`);
    report.fail(`${name} — השלב נכשל`, error.message);
  } finally {
    if (app) app.close();
  }
}

/** ממתינה למצב שמגיע באיחור (מילון, טעינת דיאלוג) בלי לנחש זמן קבוע. */
async function until(read, ok, { ms = 20_000, step = 500 } = {}) {
  let last = null;
  for (let waited = 0; waited < ms; waited += step) {
    last = await read();
    if (ok(last)) return { ok: true, value: last, waited };
    await sleep(step);
  }
  return { ok: false, value: last, waited: ms };
}

/* ------------------------------------------------------------------ */
/* 1. חוזה החוט                                                        */
/* ------------------------------------------------------------------ */

await stage('wire', { 'qa-seeded-key': { mm: 7 } }, async (app) => {
  // מה שנכתב דרך הגשר חוזר כמו שהוא.
  const roundTrip = await app.js(`(async function () {
    await window.Otzaria.call('storage.set', { key: 'qa-round-trip', value: { a: 1, b: [2, 3] } });
    var res = await window.Otzaria.call('storage.get', { key: 'qa-round-trip' });
    return JSON.stringify(res.data);
  })()`);
  // ומה שנזרע בהשמה ישירה — המסלול שכל שער משתמש בו — חוזר כמו שהוא.
  const seeded = await app.js(
    `(async function () { var r = await window.Otzaria.call('storage.get', { key: 'qa-seeded-key' }); return JSON.stringify(r.data); })()`,
  );
  const missing = await app.js(
    `(async function () { var r = await window.Otzaria.call('storage.get', { key: 'qa-no-such-key' }); return JSON.stringify(r.data); })()`,
  );
  console.log(`  כתיבה→קריאה: ${roundTrip} | זריעה→קריאה: ${seeded} | מפתח חסר: ${missing}`);

  const wantRound = JSON.stringify({ a: 1, b: [2, 3] });
  const wantSeed = JSON.stringify({ mm: 7 });
  roundTrip === wantRound && seeded === wantSeed
    ? report.pass('storage.get מחזיר את הערך עצמו', `${roundTrip} · ${seeded}`)
    : report.fail(
        'storage.get מחזיר את הערך עצמו',
        `התקבל ${roundTrip} ו-${seeded} במקום ${wantRound} ו-${wantSeed} — עטיפה כאן מפילה כל העדפה לברירת המחדל בשקט`,
      );

  missing === 'null'
    ? report.pass('מפתח שאינו קיים חוזר כ-null')
    : report.fail(
        'מפתח שאינו קיים חוזר כ-null',
        `התקבל ${missing} — אובייקט אמיתי, ולכן השומר של „אין ערך” ב-loadSetting אינו נורה`,
      );
});

/* ------------------------------------------------------------------ */
/* 2. שמירה אוטומטית — ברירת המחדל דלוקה, ולכן „כבויה” היא המדידה        */
/* ------------------------------------------------------------------ */

await stage('autosave', { 'autosave-enabled': false }, async (app) => {
  const found = await app.js("!!document.querySelector('.autosave-toggle')");
  if (!found) return report.fail('שמירה אוטומטית כבויה נשארת כבויה', 'המתג אינו ב-DOM');

  const checked = await app.js("document.querySelector('.autosave-toggle').getAttribute('aria-checked')");
  const active = await app.js("document.querySelector('.autosave-toggle').classList.contains('active')");
  console.log(`  aria-checked=${checked} | active=${active}`);

  checked === 'false' && active === false
    ? report.pass('שמירה אוטומטית כבויה נשארת כבויה', 'aria-checked=false')
    : report.fail(
        'שמירה אוטומטית כבויה נשארת כבויה',
        `aria-checked=${checked}, active=${active} — ההעדפה נזרעה כ-false והמתג קם דלוק`,
      );
});

/* ------------------------------------------------------------------ */
/* 3. הסרגל — ברירת המחדל כבויה, ולכן „דלוק” היא המדידה                 */
/* ------------------------------------------------------------------ */

await stage('ruler', { 'ruler-visible': true }, async (app) => {
  // מצב המנוע הוא האמת: `rulerPreference` מוחל דרך `adapter.run('ruler')`,
  // ולא כדגל של המעטפת (App.vue, „ההעדפה נקראת לפני הסנכרון”).
  const engine = await until(() => app.cmd('ruler'), (s) => s && s.active === true, { ms: 15_000 });
  await app.tab('תצוגה');
  const button = await app.state('סרגל');
  console.log(`  מנוע: ${JSON.stringify(engine.value)} (${engine.waited}ms) | כפתור: ${JSON.stringify(button)}`);

  engine.ok && button.active === true
    ? report.pass('סרגל דלוק חוזר דלוק', `המנוע והכפתור מסכימים אחרי ${engine.waited}ms`)
    : report.fail(
        'סרגל דלוק חוזר דלוק',
        `מנוע=${JSON.stringify(engine.value)}, כפתור active=${button.active} — ההעדפה נזרעה כ-true`,
      );
});

/* ------------------------------------------------------------------ */
/* 4. בדיקת איות — ההעדפה מפעילה משיכה של 1.3MB, ולכן היא מגיעה באיחור   */
/* ------------------------------------------------------------------ */

await stage('spellcheck', { 'spellcheck-enabled': true }, async (app) => {
  await app.tab('סקירה');
  const toggle = await until(() => app.state('בדיקת איות'), (s) => s && s.active === true, { ms: 40_000 });
  console.log(`  מתג: ${JSON.stringify(toggle.value)} (${toggle.waited}ms)`);

  toggle.ok
    ? report.pass('בדיקת איות דלוקה חוזרת דלוקה', `המילון הגיע והמתג נדלק אחרי ${toggle.waited}ms`)
    : report.fail(
        'בדיקת איות דלוקה חוזרת דלוקה',
        `המתג נשאר ${JSON.stringify(toggle.value)} אחרי ${toggle.waited}ms — ההעדפה נזרעה כ-true`,
      );
});

/* ------------------------------------------------------------------ */
/* 5. „המסמכים האחרונים” — רשימה שלמה, לא דגל                          */
/* ------------------------------------------------------------------ */

const RECENTS = [
  { token: 'qa-token-1', name: 'שולחן ערוך.docx', size: 20_480, openedAt: 1_700_000_000_000, writable: true, pinned: false },
  { token: 'qa-token-2', name: 'משנה ברורה.docx', size: 40_960, openedAt: 1_700_000_100_000, writable: false, pinned: true },
];

await stage('recents', { 'recent-documents': RECENTS }, async (app) => {
  await app.tab('קובץ');
  // „פתח קובץ” פותח את מסך „פתח מסמך”, שבתוכו יושבת הרשימה (ראו file-ops-qa).
  await app.click('פתח קובץ', { after: 1_500 });
  const screenOpen = await app.js("!!document.querySelector('.open-dialog')");
  if (!screenOpen) return report.fail('„אחרונים” שנשמרו נטענים', 'מסך „פתח מסמך” לא נפתח');

  const rows = await app.js("document.querySelectorAll('.rec-row').length");
  const empty = await app.js("!!document.querySelector('.rec-empty')");
  const names = await app.js(
    "JSON.stringify(Array.prototype.map.call(document.querySelectorAll('.rec-name'), function (e) { return e.textContent.trim(); }))",
  );
  console.log(`  שורות=${rows} | מצב „אין אחרונים”=${empty} | שמות=${names}`);

  const listed = JSON.parse(names);
  const bothThere = RECENTS.every((r) => listed.includes(r.name));
  rows === RECENTS.length && !empty && bothThere
    ? report.pass('„אחרונים” שנשמרו נטענים', `${rows} שורות: ${names}`)
    : report.fail(
        '„אחרונים” שנשמרו נטענים',
        `שורות=${rows}, „אין אחרונים”=${empty}, שמות=${names} — נזרעו ${RECENTS.length} רשומות תקינות`,
      );
});

/* ------------------------------------------------------------------ */
/* 6. זיכרון אפשרויות של דיאלוג — מסלול `loadSetting` + `mergeRemembered` */
/* ------------------------------------------------------------------ */

// ברירת המחדל היא 10 (`DEFAULT_CROP_MARKS_MM`), ולכן 7 הוא מה שרק זיכרון
// יכול להסביר.
await stage('dialog', { 'shulchan-dialog:crop-marks': { mm: 7 } }, async (app) => {
  await app.tab('שולחן העורך');
  const opened = await app.click('סימני חיתוך', { after: 1_500 });
  if (!opened) return report.fail('זיכרון אפשרויות של דיאלוג נטען', 'הפקד „סימני חיתוך” לא נמצא');

  const field = await until(
    () => app.js("(function () { var i = document.querySelector('.shcrop-dialog input[type=number]'); return i ? i.value : null; })()"),
    (v) => v === '7',
    { ms: 8_000 },
  );
  console.log(`  שדה המילימטרים: ${JSON.stringify(field.value)} (${field.waited}ms)`);

  field.ok
    ? report.pass('זיכרון אפשרויות של דיאלוג נטען', `הדיאלוג נפתח על 7 ולא על 10 (${field.waited}ms)`)
    : report.fail(
        'זיכרון אפשרויות של דיאלוג נטען',
        `השדה מציג ${JSON.stringify(field.value)} — נזרע {mm:7}, וברירת המחדל היא 10`,
      );
});

/* ------------------------------------------------------------------ */
/* 7. צבע הבד — העדפה שסופה פיקסל מצויר, ולא מצב של פקד                 */
/* ------------------------------------------------------------------ */

/**
 * ההעדפה הזאת שונה מכל קודמותיה בשרשרת שהיא עוברת: אחסון → נרמול → טוקן CSS
 * על שורש המסמך → כלל ב-`.editor-stack` → הרקע שהדפדפן מצייר. חמש חוליות,
 * ואף אחת מהן אינה נראית במצב של פקד — לכן הנמדד כאן הוא הצבע **המחושב**
 * של הבד עצמו, לא מה שהבורר בלשונית „תצוגה” מציג.
 *
 * המדידה השנייה היא הסיבה ש-`--word-canvas-bg` קיים בכלל: הסרגלים נצבעים
 * מטוקן ערכת הנושא, ודריסה שלו במקום טוקן נפרד הייתה צובעת גם אותם. שער
 * שמודד רק את הבד היה מאשר בירוק גם את המימוש הרחב מדי.
 *
 * `#3366cc` ולא צבע מהפלטה: הוא אינו ערך של שום טוקן בערכה, ולכן אי-אפשר
 * להגיע אליו במקרה.
 */
const CANVAS_SEED = '#3366cc';
const CANVAS_PAINTED = 'rgb(51, 102, 204)';

// הסרגל נזרע דלוק כדי שיהיה מה למדוד עליו: ברירת המחדל שלו כבויה, ובלי
// הזריעה מדידת „הצבע אינו נוזל לסרגלים” הייתה מדלגת על עצמה בשקט.
await stage('canvas', { 'canvas-color': CANVAS_SEED, 'ruler-visible': true }, async (app) => {
  const read = () =>
    app.js(`(function () {
      var stack = document.querySelector('.editor-stack');
      if (!stack) return null;
      var ruler = document.querySelector('.doc-vruler');
      var root = getComputedStyle(document.documentElement);
      return JSON.stringify({
        canvas: getComputedStyle(stack).backgroundColor,
        ruler: ruler ? getComputedStyle(ruler).backgroundColor : null,
        canvasToken: root.getPropertyValue('--word-canvas-bg').trim(),
        themeToken: root.getPropertyValue('--color-surface-container-highest').trim(),
      });
    })()`);

  const paint = await until(read, (raw) => raw !== null && JSON.parse(raw).canvas === CANVAS_PAINTED, {
    ms: 15_000,
  });
  const m = paint.value === null ? {} : JSON.parse(paint.value);
  console.log(`  בד=${m.canvas} | סרגל=${m.ruler} (${paint.waited}ms)`);
  console.log(`  טוקנים: --word-canvas-bg=${m.canvasToken} | ערכת נושא=${m.themeToken}`);

  paint.ok
    ? report.pass('צבע בד שנשמר חוזר ונצבע', `${m.canvas} אחרי ${paint.waited}ms`)
    : report.fail(
        'צבע בד שנשמר חוזר ונצבע',
        `הבד מצויר ${m.canvas} — נזרע ${CANVAS_SEED}, כלומר ${CANVAS_PAINTED}`,
      );

  // הדריסה חייבת להיות על הטוקן של הבד בלבד. אילו היא נעשתה על טוקן ערכת
  // הנושא, שני הערכים היו זהים — וכל מה שנצבע ממנו (הסרגלים, הפינה ביניהם,
  // כפתור היציאה ממצב מיקוד) היה נצבע יחד עם הבד.
  m.themeToken && m.themeToken.toLowerCase() !== CANVAS_SEED
    ? report.pass('הדריסה על טוקן הבד בלבד', `ערכת הנושא נשארה ${m.themeToken}`)
    : report.fail(
        'הדריסה על טוקן הבד בלבד',
        `--color-surface-container-highest הוא ${m.themeToken} — ההעדפה נכתבה על טוקן ערכת הנושא`,
      );

  m.ruler === null
    ? report.fail('צבע הבד אינו נוזל לסרגלים', 'הסרגל נזרע דלוק ואינו ב-DOM — אין מה למדוד')
    : m.ruler !== CANVAS_PAINTED
      ? report.pass('צבע הבד אינו נוזל לסרגלים', `הסרגל נשאר ${m.ruler}`)
      : report.fail(
          'צבע הבד אינו נוזל לסרגלים',
          `הסרגל נצבע ${m.ruler} — אותו צבע כמו הבד`,
        );
});

process.exit(report.print() > 0 ? 1 : 0);
