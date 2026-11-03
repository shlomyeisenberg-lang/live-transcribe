/**
 * שער QA לזמינות הגופן: מה שהבורר **מכריז** מול מה שהדפדפן **פותר**.
 *
 * ## למה זה השער היחיד שיכול למדוד את התיקון הזה
 *
 * `isFamilyAvailable` מודד רוחב על canvas, וב-jsdom אין canvas — ולכן היא
 * מחזירה `true` תמיד („בלי canvas אין לנו מה לומר”). כלומר **אף בדיקת vitest
 * אינה מריצה מדידה אמיתית**: הן מזריקות כפיל, ובצדק, אבל בכך הן מודדות את
 * ההיגיון ולא את החיווט. השרשרת המלאה — מיזוג ← composable ← פקד — רצה עם
 * canvas שיכול לומר `false` רק כאן.
 *
 * ## ההשוואה, ומה שנלמד בדרך הקשה
 *
 * הגלגול הראשון קרא את `getComputedStyle` של השורה, וזה עבד **לפני** התיקון
 * בלבד: מרגע ששורה לא-זמינה מפסיקה להכריז `font-family` היא יורשת את גופן
 * הממשק, והמדידה מדווחת „נפתר”. הוא היה מדפיס אפס כשלים גם אם התיקון יוסר.
 *
 * הגלגול השני השווה `resolves(data-value)` מול `.unavailable` על **כל** שורה
 * — ונכשל מעצם בנייתו: חמשת הגופנים שאוצריא מזריקה מוכרזים זמינים בלי מדידה
 * (`available` ב-engine/font-options.ts), ובכרום נקי שאין בו מארח הם באמת
 * אינם נפתרים. השער מדד את הסביבה, לא את המוצר.
 *
 * לכן ההשוואה נעשית **רק על השורות שהמוצר עצמו מכריז שנמדדו**
 * (`data-availability`), והאוכלוסייה נגזרת מהמוצר ולא מטבלת שמות בשער —
 * טבלה כזאת הייתה נשברת בכל גופן חדש, והתשובה הייתה „להרחיב את הטבלה”.
 *
 * ומה שמונע „ירוק על כלום”: השער דורש מספר מזערי של שורות שנמדדו. אם `verify`
 * יוסר מהמקורות, כל השורות יהפכו ל-`declared`, האוכלוסייה תתרוקן — והשער
 * ייפול. זה מה שהופך אותו לשומר על התיקון ולא רק לתיאור שלו.
 *
 * יציאה 9634.
 */
import { openApp, createReport } from './harness.mjs';

const report = createReport('זמינות הגופן בבורר', { strict: true });
const app = await openApp({ name: 'font-availability', port: Number(process.env.QA_PORT ?? 9634) });

const note = (...p) =>
  console.log(p.map((x) => (typeof x === 'string' ? x : JSON.stringify(x))).join(' '));

const T = (p, label, ms = 60_000) =>
  Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error(`תקיעה ב-${label}`)), ms))]);

/** מתחת לזה אין רשימה למדוד בכלל. */
const MIN_ROWS = 20;
/**
 * לפחות שורה אחת נמדדה — ובכוונה אחת ולא „כמספר הלטיניות”.
 *
 * רף שסופר פריטים ברשימה הוא מספר מקובע בתחפושת: הוא נופל ביום שבו
 * מישהו משנה את הרשימה מסיבה לגיטימית, והתשובה היא „לעדכן את המספר”.
 * מה שבאמת שומר על המדידה הוא הטענה על `Aptos` בסוף השער, שנגזרת
 * מהמוצר; הרף כאן מונע רק „ירוק על כלום”.
 */
const MIN_MEASURED = 1;

const SCAN = `(function () {
  /* אותה מחרוזת בדיוק של PROBE_TEXT ב-src/engine/docx-fonts.ts. */
  var PROBE = '\\u05d0\\u05d1\\u05d2\\u05d3\\u05d4\\u05d5\\u05d6\\u05d7\\u05d8\\u05d9 ABCDEFGHIJ';
  var ctx = document.createElement('canvas').getContext('2d');
  if (!ctx) return JSON.stringify({ error: 'אין canvas בדף — השער אינו יכול למדוד' });

  function width(font) { ctx.font = font; return ctx.measureText(PROBE).width; }
  /* זהה ל-isFamilyAvailable: שני בסיסים, ומי שאינו נפתר מודד כמו שניהם. */
  function resolves(name) {
    var q = JSON.stringify(name);
    return width('72px ' + q + ', monospace') !== width('72px monospace')
        || width('72px ' + q + ', serif') !== width('72px serif');
  }

  var rows = Array.prototype.slice.call(
    document.querySelectorAll('.ribbon-combo-list .ribbon-combo-option'),
  );
  var counts = { declared: 0, measured: 0, missing: 0 };
  var mismatched = [];
  var badMarked = [];

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    /* השם עצמו, ולא הסגנון המחושב — ראו הערת הפתיחה. */
    var value = row.getAttribute('data-value') || '';
    if (!value) continue;
    var state = row.getAttribute('data-availability') || 'declared';
    var flagged = row.classList.contains('unavailable');
    counts[state] = (counts[state] || 0) + 1;

    /* שורה מסומנת חייבת להסביר, ולא לצייר — בכל מצב. */
    if (flagged) {
      var painted = (row.getAttribute('style') || '').indexOf('font-family') !== -1;
      if (!row.getAttribute('data-tip-title') || painted) badMarked.push(value);
    }

    /* ההשוואה עצמה — רק על מי שהמוצר מכריז שמדד. */
    if (state !== 'measured' && state !== 'missing') continue;
    var can = resolves(value);
    var expected = state === 'measured';
    if (can !== expected || flagged !== (state === 'missing')) {
      mismatched.push({ value: value, state: state, resolves: can, flagged: flagged });
    }
  }

  return JSON.stringify({ total: rows.length, counts: counts, mismatched: mismatched, badMarked: badMarked });
})()`;

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1600, height: 1000, deviceScaleFactor: 1, mobile: false,
  });
  await app.sleep(500);
  await app.tab('בית');
  await app.sleep(300);
  await app.click('גופן');
  await app.sleep(1_200);

  const scan = await T(app.js(SCAN), 'scan').then((s) => JSON.parse(s));
  if (scan.error) throw new Error(scan.error);

  const checked = scan.counts.measured + scan.counts.missing;
  note('נסרקו:', scan.total, '| לפי הכרזה:', scan.counts);
  if (scan.mismatched.length) note('אי-התאמות:', scan.mismatched);

  scan.total >= MIN_ROWS
    ? report.pass('יש רשימה למדוד', `${scan.total} שורות`)
    : report.fail('יש רשימה למדוד', `רק ${scan.total} שורות — מתחת ל-${MIN_ROWS}`);

  checked >= MIN_MEASURED
    ? report.pass('יש שורות שנמדדו — השער אינו ירוק על כלום', `${checked} שורות נמדדו`)
    : report.fail(
        'יש שורות שנמדדו — השער אינו ירוק על כלום',
        `רק ${checked} שורות מכריזות שנמדדו, ונדרשות ${MIN_MEASURED}. כך נראה verify שהוסר מהמקורות`,
      );

  scan.mismatched.length === 0
    ? report.pass(
        'ההכרזה מסכימה עם הדפדפן',
        `${checked} שורות נמדדו (${scan.counts.missing} חסרות), ו-${scan.counts.declared} מוכרזות ופטורות`,
      )
    : report.fail(
        'ההכרזה מסכימה עם הדפדפן',
        `${scan.mismatched.length} שורות חולקות, הראשונה ${JSON.stringify(scan.mismatched[0])}`,
      );

  scan.badMarked.length === 0
    ? report.pass('שורה לא-זמינה מסבירה ואינה מצוירת', `${scan.counts.missing} שורות עומדות בשניים`)
    : report.fail(
        'שורה לא-זמינה מסבירה ואינה מצוירת',
        `חסר הסבר או שיש ציור ב-${JSON.stringify(scan.badMarked)}`,
      );

  /*
   * `Aptos` היא ברירת המחדל של Word 365 והיא תמיד ברשימה
   * (`LATIN_FONT_FAMILIES`), ולכן היא המקרה שהתיקון נבנה בשבילו. הנכונות שלה
   * כבר נבדקה למעלה; כאן נמדד רק שהיא **נמדדה** — שורה שחזרה להיות „מוכרזת”
   * פירושה שהמדידה הפסיקה לחול עליה.
   */
  const aptos = await app.js(`(function () {
    var el = document.querySelector('.ribbon-combo-list .ribbon-combo-option[data-value="Aptos"]');
    return el ? el.getAttribute('data-availability') || 'declared' : 'null';
  })()`);
  aptos === 'measured' || aptos === 'missing'
    ? report.pass('Aptos נמדדת — המקרה שהתיקון נבנה בשבילו', aptos === 'missing' ? 'אינה מותקנת, ומסומנת' : 'מותקנת במכונה הזאת')
    : report.fail('Aptos נמדדת — המקרה שהתיקון נבנה בשבילו', `ההכרזה עליה היא ${aptos}`);
} catch (error) {
  console.error('השער נפל:', error);
  report.fail('השער השלים את הריצה', String((error && error.message) || error));
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
