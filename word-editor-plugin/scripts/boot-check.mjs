/**
 * שער האתחול: מה קורה לתוסף כשאוצריא משגרת plugin.boot **לפני** שהבאנדל שלו
 * נטען.
 *
 * זה הכשל שנצפה בפועל — „אוצריא לא סיימה לאתחל את התוסף” בטעינה ראשונה, ותוסף
 * שעולה רק אחרי רענון. האירוע חד-פעמי ואינו משוחזר, ולכן הרשמה בזמן טעינת
 * המודול אינה מספיקה: 10MB של באנדל נטענים אחרי שהאירוע כבר נורה.
 *
 * שלושה מצבים, כולם על ה-dist האמיתי מ-file://, עם Host-דמה שמוזרק לפני app.js:
 *
 *   1. ה-latch + boot מוקדם            → האירוע נתפס  (`data-boot="event"`).
 *   2. בלי latch + boot מוקדם + RPC מת → האתחול נכשל  (`data-boot="failed"`).
 *                                        הבקרה: בלעדיה אין הוכחה שה-latch הוא
 *                                        מה שמציל.
 *   3. בלי latch + בלי boot + RPC חי    → שחזור ב-RPC  (`data-boot="recovered"`).
 *
 * מה מסמן את התוצאה, ולמה זה השתנה: הגרסה הראשונה של השער בדקה אם הכפתור
 * „פתיחת קובץ Word” נפתח, מפני שאז כשל אתחול הקפיא את הממשק. המעטפת הנוכחית
 * נכשלת פתוח — היא מרכיבה את Vue ופותחת מסמך ריק בלי להמתין ל-boot, וכשל
 * משאיר רק את ערכת הנושא של ברירת המחדל — ולכן הכפתור נפתח **בכל שלושת
 * המצבים**, כולל בבקרה. שער שהבקרה שלו עוברת אינו מודד כלום. מעכשיו נמדדת
 * התוצאה עצמה, `data-boot` על שורש ה-HTML (src/main.ts), שגם מפרידה בין תפיסה
 * ב-latch לשחזור ב-RPC — הבחנה שהסימן הקודם לא ידע לעשות.
 *
 * מונע דרך CDP ולא ב-`--dump-dom`: ברגע שהמנוע עולה, אירוע ה-load אינו מגיע
 * (נמדד — הדפדפן נתלה), ו-`--virtual-time-budget` נתקע מול ה-workers.
 *
 *   npm run build && npm run check:boot
 */
import { readFileSync, writeFileSync, existsSync, rmSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage, requireChrome, sleep } from './cdp.mjs';

const DIST = join(dirname(fileURLToPath(import.meta.url)), '..', 'dist');
/** מכסה את BOOT_TIMEOUT_MS (15 שניות) בתוספת עלייה של המנוע. */
const OBSERVE_MS = 28_000;

if (!existsSync(join(DIST, 'index.html'))) {
  console.error('dist/index.html אינו קיים — הריצו npm run build תחילה');
  process.exit(1);
}
requireChrome();

/**
 * Host-דמה. `answersRpc` קובע אם `app.getInfo`/`app.getTheme` עונים — כלומר אם
 * מסלול השחזור יכול להצליח; `firesBoot` אם האירוע נורה בכלל.
 */
function stub({ firesBoot, answersRpc }) {
  return `
<script>
  (function () {
    var BOOT = {
      plugin: { id: 'boot-check', version: '0' },
      app: { version: '9.9.9', platform: 'boot-check' },
      theme: { mode: 'light', colorScheme: {}, typography: {} },
      connectivity: { isOnline: false },
      permissions: []
    };
    var answers = ${answersRpc};
    window.Otzaria = {
      call: function (method) {
        if (answers && method === 'app.getInfo') {
          return Promise.resolve({ success: true, data: BOOT.app, error: null });
        }
        if (answers && method === 'app.getTheme') {
          return Promise.resolve({ success: true, data: BOOT.theme, error: null });
        }
        return Promise.resolve({
          success: false, data: null, error: { message: 'Otzaria SDK not ready yet' }
        });
      },
      on: function () {},
      off: function () {}
    };
    if (${firesBoot}) {
      // כאן העיקר: האירוע נורה עכשיו, לפני שאף שורה מ-app.js הורצה.
      window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));
    }
  })();
</script>
`;
}

const html = readFileSync(join(DIST, 'index.html'), 'utf8');
const latchEnd = html.indexOf('</script>');
if (latchEnd === -1 || !html.slice(0, latchEnd).includes('__otzariaBoot')) {
  console.error('ה-latch אינו הסקריפט הראשון ב-dist/index.html — check:dist היה אמור לתפוס זאת');
  process.exit(1);
}
const afterLatch = latchEnd + '</script>'.length;
const withLatch = (injected) => html.slice(0, afterLatch) + injected + html.slice(afterLatch);
const withoutLatch = (injected) => injected + html.slice(afterLatch);

const CASES = [
  {
    name: 'ה-latch + boot מוקדם',
    body: withLatch(stub({ firesBoot: true, answersRpc: false })),
    expect: 'event',
  },
  {
    name: 'בקרה: בלי latch, RPC מת',
    body: withoutLatch(stub({ firesBoot: true, answersRpc: false })),
    expect: 'failed',
  },
  {
    name: 'שחזור: בלי latch ובלי boot, RPC חי',
    body: withoutLatch(stub({ firesBoot: false, answersRpc: true })),
    expect: 'recovered',
  },
];

/**
 * מצב התוסף. `boot` הוא התוצאה שהמעטפת מצהירה עליה; שורת המצב נקראת רק
 * לתיעוד — היא ממשיכה להתחלף (מסמך שנטען, „טרם נשמר”) ולכן היא סימן גרוע.
 */
async function probe(cdp) {
  return (
    (await cdp.evaluate(`(function () {
      // הדף עשוי להיות בעיצומה של הניווט: אין עוד documentElement, ולא
      // כדאי להכשיל את השער על תזמון.
      var root = document.documentElement;
      var status = document.getElementById('status');
      return {
        boot: root ? root.getAttribute('data-boot') : null,
        status: status ? status.textContent : null
      };
    })()`)) ?? { boot: null, status: null }
  );
}

/** `pending` הוא מצב ביניים: ה-boot טרם הוכרע, וממתינים עד ה-deadline. */
function classify({ boot }) {
  return boot ?? 'pending';
}

let failures = 0;

for (const [index, testCase] of CASES.entries()) {
  const path = join(DIST, 'boot-check-tmp.html');
  writeFileSync(path, testCase.body);

  let outcome = 'pending';
  let status = null;
  let page;
  try {
    page = await openPage(`file://${path}`, { label: `boot-${index}` });
    const deadline = Date.now() + OBSERVE_MS;
    while (Date.now() < deadline) {
      const reading = await probe(page.cdp);
      status = reading.status;
      outcome = classify(reading);
      if (outcome !== 'pending') break;
      await sleep(200);
    }
  } catch (error) {
    console.error(`${testCase.name}: ${error.message}`);
    failures++;
    continue;
  } finally {
    page?.close();
    rmSync(path, { force: true });
  }

  const verdict = outcome === testCase.expect ? '✓' : '✗';
  console.log(`${verdict} ${testCase.name}: ${outcome} (צפוי ${testCase.expect}) — "${(status ?? '').slice(0, 80)}"`);
  if (outcome !== testCase.expect) failures++;
}

if (failures) {
  console.error(
    'שער האתחול נכשל. אם הבקרה הצליחה לאתחל בכל זאת — הבדיקה אינה מודדת כלום ' +
      'ויש לתקן אותה; אם מצב 1 או 3 לא אותחל — התוסף יאבד את ערכת הנושא של ' +
      'אוצריא, ובגרסה שתחזיר המתנה ל-boot הוא ייתקע כמו קודם.',
  );
  process.exit(1);
}
console.log('שער האתחול עבר: boot מוקדם נתפס ב-latch, ובהיעדרו השחזור ב-RPC מאתחל את התוסף.');
