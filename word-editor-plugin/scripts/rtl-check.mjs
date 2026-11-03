/**
 * שער הכיווניות: האם **מסמך חדש** באמת נפתח מימין לשמאל, על ה-dist הארוז
 * מ-`file://`.
 *
 * למה שער ולא בדיקת יחידה: הניסיון הראשון להחיל RTL עשה זאת בפקודות ה-Ribbon
 * (`direction-rtl`, `text-align`), הן נכשלו ב-`selection-required` — לפסקה של
 * מסמך שנפתח כרגע אין עוד סמן — הכשל נבלע ב-`void`, ובדיקות היחידה שאישרו
 * ש„הפקודה נשלחה” המשיכו לעבור. מה שהיה חסר הוא מדידה על המנוע האמיתי.
 *
 * מה נמדד:
 *   1. `data-document-direction="rtl"` על שורש ה-HTML. זו אינה הצהרה קוסמטית:
 *      האפליקציה קובעת אותה **רק** כששלוש הקבלות של המנוע — docDefaults, מקטע
 *      ופסקה — חזרו בהצלחה, ומוחקת אותה בכל כשל. כלומר התכונה הזאת היא
 *      הקבלות של המנוע, במקום שאפשר לקרוא אותו מבחוץ.
 *   2. **הכיווניות שנראית בפועל**: `dir` ו-`text-align` המחושבים על הפסקה
 *      שהמנוע רינדר. זו התוספת החשובה, והיא נולדה מחור אמיתי — עד כאן השער
 *      מדד את ההצהרה של האפליקציה על עצמה בלבד, כלומר קבלות שחזרו „הצלחה”.
 *      קבלה מוצלחת אינה פיקסל: `styles.apply` מחזירה `{success: true}` גם
 *      כשלא נכתב דבר (ראו את בדיקת `after` ב-engine/document-defaults.ts), ולכן
 *      השער היה עובר על מסמך שמרונדר משמאל לימין. ההמתנה גדלה בגלל זה, והיא
 *      ממתינה לשני הדברים יחד — ראו את לולאת הדגימה: הסדר ביניהם אינו מובטח.
 *   3. אין שורת console של „כיווניות המסמך החדש לא הוחלה” — הנוסח שהאפליקציה
 *      כותבת ללוג של אוצריא כשההחלה נכשלה חלקית.
 *   4. **הצד שסרגל הגלילה האנכי צויר בו.** הצד נגזר מכיווניות מיכל הגלילה
 *      עצמו, ומיכל שיורש rtl מקבל סרגל בשמאל; `.editor-stack__host` מוצהר
 *      `direction: ltr` בשביל זה בדיוק (ראו styles/shell.css). המדידה כאן היא
 *      גם ההצהרה וגם הפיקסל — `clientLeft` על מיכל בלי גבול הוא רוחב הסרגל
 *      כשהוא בשמאל ואפס כשהוא בימין. השער הזה הוא הבית של המדידה מפני
 *      ש-`direction` על מיכל הגלילה הוא בדיוק מה שהוא מודד ממילא: אותה הצהרה
 *      אחת קובעת גם את צד הסרגל וגם מה שפסקה יורשת, ובדיקה שתמדוד אותה בשני
 *      מקומות תיפול פעמיים על אותו שינוי.
 *
 * מה **אינו** נמדד כאן, במפורש: התוסף אינו חושף את מופע SuperDoc ל-`window`,
 * ולכן השער אינו קורא את `sectionDirection` ואת `bidi` מהמסמך בעצמו, ואינו יוצר
 * פסקה שנייה כדי לאמת את שכבת `docDefaults` (זו שקובעת לפסקות הבאות). שתי
 * הקריאות האלה נעשו ידנית ב-CDP בזמן המימוש — המנוע החזיר `sectionDirection:
 * "rtl"`, פסקה שנוצרה אחרי הפתיחה רונדרה `dir="rtl"`, והייצוא נשא `w:bidi`
 * ב-`w:pPrDefault`, ב-`w:sectPr` ובפסקה — והחלפתן בשער הייתה מחייבת לפתוח את
 * המנוע ל-`window` בבנייה ארוזה: API ציבורי חדש לצורך בדיקה, שאינו שווה את
 * המחיר. שכבת `docDefaults` מכוסה במקום זה בבדיקת ה-`after` שלה.
 *
 * שמות המחלקות `.superdoc-page` ו-`.superdoc-fragment` הם של המנוע ולא שלנו.
 * שדרוג מנוע שישנה אותם יפיל את השער עם „לא נמצאה פסקה מרונדרת”, וזו התנהגות
 * מכוונת: עדיף שער שנשבר בקול על שינוי מבנה מאשר שער שממשיך לעבור בלי למדוד.
 *
 *   npm run build && npm run check:rtl
 */
import { existsSync, readFileSync, writeFileSync, rmSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage, requireChrome, sleep } from './cdp.mjs';

const DIST = join(dirname(fileURLToPath(import.meta.url)), '..', 'dist');
const INDEX = join(DIST, 'index.html');
/** נדיב: המנוע והמסמך הריק עולים ב-~500ms ארוז, וההחלה היא אחרי onReady. */
const SETTLE_MS = 9000;
/**
 * הרינדור עצמו איטי בהרבה מההחלה — נמדד שהפסקה מופיעה בין 10 ל-20 שניות
 * ב-headless. במקום להגדיל את ההמתנה הקבועה לגבול העליון, השער דוגם עד שיש מה
 * למדוד: מכונה מהירה מסיימת מוקדם, ומכונה עמוסה אינה נכשלת על תזמון.
 */
const RENDER_TIMEOUT_MS = 45000;
const POLL_MS = 500;

if (!existsSync(INDEX)) {
  console.error('dist/index.html אינו קיים — הריצו npm run build תחילה');
  process.exit(1);
}
requireChrome();

/** Host-דמה שעונה ל-RPC, כדי שהמעטפת תעלה בלי אוצריא. */
const STUB = `
<script>
  (function () {
    var BOOT = {
      plugin: { id: 'rtl-check', version: '0' },
      app: { version: '9.9.9', platform: 'rtl-check' },
      theme: { mode: 'light', colorScheme: {}, typography: {} },
      connectivity: { isOnline: false },
      permissions: []
    };
    window.Otzaria = {
      call: function (method) {
        if (method === 'app.getInfo') return Promise.resolve({ success: true, data: BOOT.app, error: null });
        if (method === 'app.getTheme') return Promise.resolve({ success: true, data: BOOT.theme, error: null });
        return Promise.resolve({ success: false, data: null, error: { message: 'לא נתמך: ' + method } });
      },
      on: function () {},
      off: function () {}
    };
    // console של תוסף מגיע ללוג של אוצריא; כאן הוא נאסף כדי שהשער יראה אזהרה.
    window.__rtlCheckLog = [];
    ['warn', 'error'].forEach(function (level) {
      var original = console[level];
      console[level] = function () {
        window.__rtlCheckLog.push([].map.call(arguments, String).join(' '));
        original.apply(console, arguments);
      };
    });
    window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));
  })();
</script>
`;

const html = readFileSync(INDEX, 'utf8');
const afterLatch = html.indexOf('</script>') + '</script>'.length;
const path = join(DIST, 'rtl-check-tmp.html');
writeFileSync(path, html.slice(0, afterLatch) + STUB + html.slice(afterLatch));

const PROBE = `(function () {
  var fragments = [].slice.call(document.querySelectorAll('.superdoc-fragment'));
  // מיכל הגלילה של המסמך. \`clientLeft\` הוא רוחב הגבול השמאלי ועוד רוחב הסרגל
  // **כשהוא בשמאל** — ואפס כשהוא בימין. לא נמצא כאן גבול, ולכן הערך הוא מדידה
  // ישירה של הצד שבו הדפדפן צייר את הסרגל, ולא הצהרה על עצמנו.
  var host = document.querySelector('.editor-stack__host');
  return {
    direction: document.documentElement.getAttribute('data-document-direction'),
    status: (document.getElementById('status') || {}).textContent,
    scroller: host === null ? null : {
      direction: getComputedStyle(host).direction,
      scrollable: host.scrollHeight > host.clientHeight,
      barWidth: host.offsetWidth - host.clientWidth,
      gutterStart: host.clientLeft
    },
    rendered: fragments.map(function (el) {
      var style = getComputedStyle(el);
      return { dir: el.getAttribute('dir'), direction: style.direction, textAlign: style.textAlign };
    }),
    log: window.__rtlCheckLog || []
  };
})()`;

const page = await openPage(`file://${path}`, { label: 'rtl' });
let report;
try {
  await sleep(SETTLE_MS);
  report = await page.cdp.evaluate(PROBE);
  /*
   * דגימה עד שיש מה למדוד: פסקה מרונדרת **וגם** הכרעה של האפליקציה — הצהרה,
   * או אזהרת הכיווניות שהיא כותבת כשההחלה נכשלה. הדגימה נעצרת ברגע ששניהם
   * שם, ולכן ה-timeout הוא גבול ולא זמן ריצה.
   *
   * שניהם ולא רק הפסקה: הפסקה מצוירת כשהמנוע מסיים לפתוח, וההצהרה נקבעת אחרי
   * שהקבלות של ההחלה חזרו — כלומר אחריה, בשלב „מסדר את התצוגה”. במנוע שבו
   * הפתיחה איטית מההחלה זה לא נראה; ב-superdoc 2.15 השער דגם בדיוק בין
   * השניים ונפל על `data-document-direction=חסר`, על מסמך שנפתח תקין. כשל
   * אמיתי עדיין נופל: הצהרה שלא מגיעה לעולם ממצה את הגבול ומדווחת.
   */
  const settled = (r) =>
    (r?.rendered?.length ?? 0) > 0 &&
    (r.direction === 'rtl' || (r.log ?? []).some((line) => line.includes('כיווניות')));
  const deadline = Date.now() + RENDER_TIMEOUT_MS;
  while (!settled(report) && Date.now() < deadline) {
    await sleep(POLL_MS);
    report = await page.cdp.evaluate(PROBE);
  }
} finally {
  page.close();
  rmSync(path, { force: true });
}

const errors = [];
if (report?.direction !== 'rtl') {
  errors.push(
    `data-document-direction=${report?.direction ?? 'חסר'} — האפליקציה לא הצהירה על RTL, ` +
      'כלומר לפחות אחת מקבלות המנוע נכשלה',
  );
}
const rendered = report?.rendered ?? [];
if (rendered.length === 0) {
  errors.push(
    'לא נמצאה פסקה מרונדרת (.superdoc-fragment) — או שהרינדור לא הסתיים בזמן, ' +
      'או שהמנוע שינה את מבנה ה-DOM שלו ויש לעדכן את השער',
  );
}
// כל פסקה ולא רק הראשונה: מסמך ריק נפתח עם אחת, ובדיקה של „לפחות אחת RTL”
// הייתה עוברת גם על מסמך שבו השאר משמאל לימין.
for (const [index, fragment] of rendered.entries()) {
  if (fragment.direction !== 'rtl') {
    errors.push(`פסקה ${index + 1} מרונדרת ב-direction=${fragment.direction}, ולא rtl`);
  }
  // `start` ב-RTL מתפרש כימין, ולכן שני הערכים תקינים; `left` אינו.
  if (fragment.textAlign !== 'right' && fragment.textAlign !== 'start') {
    errors.push(`פסקה ${index + 1} מיושרת ל-${fragment.textAlign}, ולא לימין`);
  }
}

// צד סרגל הגלילה. הוא נמדד כאן ולא בבדיקת יחידה מאותו טעם שכל השער הזה חי:
// `direction: ltr` בגיליון הוא הצהרה, והצד שהדפדפן צייר בו את הסרגל הוא
// הפיקסל. שתי המדידות נחוצות — ההצהרה מגנה מפני מחיקה בעריכה, והפיקסל מגן
// מפני כלל אחר שידרוס אותה.
const scroller = report?.scroller ?? null;
if (scroller === null) {
  errors.push('לא נמצא מיכל הגלילה (.editor-stack__host) — המעטפת לא עלתה, או שהמחלקה שונתה');
} else {
  if (scroller.direction !== 'ltr') {
    errors.push(
      `מיכל הגלילה ב-direction=${scroller.direction} — סרגל הגלילה האנכי יצויר בשמאל, ` +
        'ראו את ההערה על .editor-stack__host ב-styles/shell.css',
    );
  }
  if (scroller.gutterStart > 0) {
    errors.push(`סרגל הגלילה תופס ${scroller.gutterStart}px בקצה השמאלי של המיכל, ולא בימני`);
  }
}

for (const line of report?.log ?? []) {
  if (line.includes('כיווניות')) errors.push(`אזהרה מהתוסף: ${line}`);
}

// `barWidth=0` על מיכל שגולל אינו כשל — כך נראה סרגל overlay — אבל הוא כן
// אומר שהמדידה של `gutterStart` לא הוכיחה דבר. מודפס כדי שלא ייקרא כירוק.
const bar =
  scroller === null
    ? 'אין מיכל'
    : `${scroller.direction} גולל=${scroller.scrollable ? 'כן' : 'לא'} ` +
      `סרגל=${scroller.barWidth}px שמאל=${scroller.gutterStart}px`;

console.log(
  `direction=${report?.direction ?? '?'} פסקות=${rendered.length} ` +
    `מרונדר=${rendered.map((f) => `${f.direction}/${f.textAlign}`).join(', ') || 'אין'} ` +
    `מיכל=${bar} status="${(report?.status ?? '').slice(0, 40)}"`,
);

if (errors.length) {
  for (const error of errors) console.error(`שגיאה: ${error}`);
  process.exit(1);
}
console.log(
  'שער הכיווניות עבר: מסמך חדש נפתח RTL — שלוש הקבלות, והפסקה מרונדרת מימין לשמאל.',
);
