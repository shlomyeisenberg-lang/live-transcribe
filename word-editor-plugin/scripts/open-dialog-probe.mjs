/**
 * שער „פתח מסמך” — בדפדפן אמיתי, על ה-dist הארוז.
 *
 * ## למה זה לא יכול להיות בדיקת רכיב
 *
 * `tests/component/open-document-dialog.test.ts` מודד את החיווט: אילו אירועים
 * יוצאים, מה מרונדר, ומה ספירת הצורות בכל תצוגה מקדימה. כל אלה נכונים
 * ב-jsdom, שאינו מחשב פריסה כלל — ולכן **שלוש הטענות המרכזיות של העיצוב אינן
 * נבדקות שם בכלל**:
 *
 * 1. **הגיליון מצויר בגודל אמיתי.** ה-SVG מקבל `height: 6.5em` ו-
 *    `aspect-ratio: 214/301`, ובלי מנוע פריסה שניהם מחרוזות. גיליון ברוחב אפס
 *    הוא בדיוק סוג הרגרסיה שעוברת ירוקה ב-jsdom.
 * 2. **הרשת גולשת לשורות ואינה גולשת אופקית.** `auto-fit` עם `minmax(148px)`
 *    הוא הבטחה על מספר עמודות בכל רוחב חלון, והיא נגזרת מחישוב.
 * 3. **הדיאלוג אינו יוצר גלילה אופקית בגוף הדף** — הכשל הקלאסי של רוחב קבוע
 *    בממשק RTL.
 *
 * ומעבר למדידה: הוא מצלם. `tmp/open-dialog-*.png` הם מה שמאפשר להסתכל על
 * המסך במקום להאמין לתיאור שלו.
 *
 *   npm run build && node scripts/open-dialog-probe.mjs
 */
import { readFileSync, writeFileSync, existsSync, mkdirSync, rmSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { pathToFileURL } from 'node:url';
import { openPage, requireChrome, sleep } from './cdp.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const DIST = join(ROOT, 'dist');
const TMP = join(ROOT, 'tmp');

const READY_MS = 40_000;

/**
 * שלושת הרוחבים שהמפרט נוקב בהם (§1.3): מעל 875 — חמש עמודות; 705–874 —
 * ארבע; מתחת ל-534 — שתיים. הם נבדקים ולא מוצהרים, כי הנוסחה שמאחוריהם
 * (`min(960, 0.94vw)` פחות ריפוד) נשברת מכל שינוי בריפוד או ב-gap.
 */
const VIEWPORTS = [
  { width: 1440, height: 900, label: 'wide', expectColumns: 5 },
  { width: 800, height: 900, label: 'medium', expectColumns: 4 },
  { width: 520, height: 900, label: 'narrow', expectColumns: 2 },
];

if (!existsSync(join(DIST, 'index.html'))) {
  console.error('dist/index.html אינו קיים — הריצו npm run build תחילה');
  process.exit(1);
}
requireChrome();
mkdirSync(TMP, { recursive: true });

/**
 * שלושה מסמכים אחרונים, ושניים מהם מכסים מצב קצה: אחד מוצמד (הקו מתחת
 * למוצמדת האחרונה), ואחד עם `openedAt: 0` ו-`size: 0` — „לא ידוע”, שאסור לו
 * להצטייר כ„0 בייט”.
 */
const RECENTS = [
  { token: 't1', name: 'שולחן ערוך — אורח חיים.docx', size: 1_482_000, openedAt: Date.now() - 3_600_000, pinned: true },
  { token: 't2', name: 'בראשית — הגהות.docx', size: 24_500, openedAt: Date.now() - 7_200_000, pinned: false },
  { token: 't3', name: 'קונטרס.docx', size: 0, openedAt: 0, pinned: false },
];

const stubFor = (recents) => `<script>
  window.__otzariaStorage = {
    'recent-documents': ${JSON.stringify(recents)}
  };
  window.Otzaria = {
    call: function (method, params) {
      if (method === 'app.getInfo') return Promise.resolve({ success: true, data: { version: '9', platform: 'p' }, error: null });
      if (method === 'app.getTheme') return Promise.resolve({ success: true, data: { mode: 'light', colorScheme: {}, typography: {} }, error: null });
      if (method === 'storage.get') {
        return Promise.resolve({ success: true, data: window.__otzariaStorage[params && params.key] || null, error: null });
      }
      if (method === 'storage.set') return Promise.resolve({ success: true, data: true, error: null });
      return Promise.resolve({ success: false, data: null, error: { message: 'no' } });
    },
    on: function () {},
    off: function () {}
  };
</script>`;

const html = readFileSync(join(DIST, 'index.html'), 'utf8');
const latchEnd = html.indexOf('</script>') + '</script>'.length;

/** דף עם רשימת אחרונים נתונה. שני התרחישים למטה נבדלים בזה בלבד. */
function writePage(name, recents) {
  const path = join(DIST, name);
  writeFileSync(path, html.slice(0, latchEnd) + stubFor(recents) + html.slice(latchEnd));
  return path;
}

const pagePath = writePage('open-dialog-tmp.html', RECENTS);

const failures = [];
function check(ok, message) {
  console.log(`${ok ? '  ✓' : '  ✗'} ${message}`);
  if (!ok) failures.push(message);
}

/**
 * שני תרחישים, ורשימת האחרונים היא כל ההבדל ביניהם.
 *
 * הוא אינו קוסמטי. „פתח קובץ” נוחת על `.open-pane`, ונחיתה דורשת מרחק
 * גלילה — ומסך בלי אחרונים הוא בדיוק המקום שבו מרחק כזה יכול להיעלם. זו
 * הטענה שהתרחיש השני נועד להפיל: `min-block-size: 100%` על הפן, והגובה
 * הקבוע של החלון שמאפשר לו להיפתר.
 *
 * הרשת, הגיליון והרשימה אינם תלויים ברשימת האחרונים ולכן נמדדים בשלושת
 * הרוחבים בתרחיש הראשון בלבד.
 */
const SCENARIOS = [
  { key: 'recents', label: 'עם אחרונים', recents: RECENTS, viewports: VIEWPORTS },
  { key: 'empty', label: 'בלי אחרונים', recents: [], viewports: [VIEWPORTS[0]] },
];

/**
 * הנחיתה כפי שהעין רואה אותה: איפה `.open-pane` יושב ביחס לחלון הגלילה, כמה
 * גלילה נעשתה כדי להביא אותו לשם, והאם רצועת התבניות אכן יצאה מהתמונה.
 */
const LANDING = `(() => {
  const body = document.querySelector('.open-body');
  const pane = document.querySelector('.open-pane');
  const tpl = document.querySelector('.tpl-section');
  const bodyRect = body.getBoundingClientRect();
  const paneRect = pane.getBoundingClientRect();
  return {
    scrollTop: Math.round(body.scrollTop),
    maxScroll: Math.round(body.scrollHeight - body.clientHeight),
    padding: Math.round(Number.parseFloat(getComputedStyle(body).paddingBlockStart)),
    paneFromTop: Math.round(paneRect.top - bodyRect.top),
    paneHeight: Math.round(paneRect.height),
    viewport: Math.round(body.clientHeight),
    templatesBottom: Math.round(tpl.getBoundingClientRect().bottom - bodyRect.top),
    focus: document.activeElement ? document.activeElement.className : '',
  };
})()`;

for (const scenario of SCENARIOS) {
  console.log(`\n════ ${scenario.label} ════`);
  const pagePath = writePage(`open-dialog-${scenario.key}-tmp.html`, scenario.recents);
  const page = await openPage(pathToFileURL(pagePath).href, { label: `open-dialog-${scenario.key}` });
  try {
    /*
     * ההמתנה היא ל**סוף הפתיחה הראשונה**, לא להופעת המעטפת.
     *
     * `openOpenDialog` יוצא מיד כש-`isOpenBusy()` — התנהגות מכוונת: אין לפתוח
     * „פתח מסמך” בזמן שמסמך נטען לתוך אותו טאב. המצב הזה נמשך שניות בעלייה,
     * ובדיקה שדוחפת Ctrl+O בתוכו מודדת את המשמר במקום את הדיאלוג (נמדד: אחרי
     * 600ms הדיאלוג אינו נפתח, ואחרי ~3 שניות כן).
     *
     * `isOpening` אינו חשוף ל-DOM (מחוון הטעינה בשורת המצב הוא תצוגה נפרדת
     * ואינו אותו דגל), ולכן ההמתנה היא **על התוצאה**: הקיצור נדחף שוב עד
     * שהדיאלוג נפתח. זה גם מה שמשתמש עושה, וזה מודד בדיוק את מה שהשער טוען.
     */
    async function openDialog(code) {
      const key = code === 'KeyN' ? 'n' : 'o';
      const deadline = Date.now() + READY_MS;
      for (;;) {
        await page.cdp.evaluate(
          `window.dispatchEvent(new KeyboardEvent('keydown', { key: '${key}', code: '${code}', ctrlKey: true, bubbles: true }))`,
        );
        await sleep(400);
        if (await page.cdp.evaluate("!!document.querySelector('.open-dialog')")) return true;
        if (Date.now() > deadline) return false;
      }
    }

    async function closeDialog() {
      await page.cdp.evaluate("document.querySelector('.open-close-btn').click()");
      await sleep(250);
    }

    async function shoot(name) {
      const shot = await page.cdp.send('Page.captureScreenshot', { format: 'png' });
      writeFileSync(join(TMP, `${name}.png`), Buffer.from(shot.result.data, 'base64'));
      console.log(`  📸 tmp/${name}.png`);
    }

    /*
     * הסימן הקנוני של המאגר לעלייה שהסתיימה — אותו ביטוי בדיוק כמו
     * context-menu-probe, readout-probe ו-session-probe. הוא מכסה גם את מסך
     * הטעינה: `otzaria-splash` יושב **מעל** הממשק, וצילום שנלקח לפני שהוא הוסר
     * מראה אותו ולא את הדיאלוג (נמדד — זה מה שקרה בריצה הראשונה כאן).
     */
    const shellDeadline = Date.now() + READY_MS;
    for (;;) {
      const ready = await page.cdp.evaluate(
        '!!window.__otzariaEditor && !document.getElementById("otzaria-splash")',
      );
      if (ready) break;
      if (Date.now() > shellDeadline) throw new Error('העלייה לא הסתיימה תוך 40 שניות');
      await sleep(250);
    }

    for (const viewport of scenario.viewports) {
      console.log(`\n[${viewport.label}] ${viewport.width}×${viewport.height}`);
      await page.cdp.send('Emulation.setDeviceMetricsOverride', {
        width: viewport.width,
        height: viewport.height,
        deviceScaleFactor: 1,
        mobile: false,
      });
      await sleep(200);

      /* ---------------------------------------------------------------- */
      /* Ctrl+O — הנחיתה על „עיון בקבצים…”                                 */
      /* ---------------------------------------------------------------- */

      // פתיחה דרך הקיצור — אותו מסלול שהמשתמש עובר בו.
      const opened = await openDialog('KeyO');
      check(opened, 'Ctrl+O פתח את הדיאלוג');
      if (!opened) break;

      const landing = await page.cdp.evaluate(LANDING);
      // הפן בגובה חלון גלילה שלם: זה מה שהופך את הנחיתה לזהה בין „יש עשרים
      // אחרונים” ל„אין אף אחד”, ובלעדיו אין בכלל מה לגלול בתרחיש הריק.
      check(
        landing.paneHeight >= landing.viewport - 2 * landing.padding - 1,
        `הפן בגובה חלון גלילה שלם: ${landing.paneHeight}px מתוך ${landing.viewport}px`,
      );
      check(landing.scrollTop > 0, `Ctrl+O גלל בפועל (${landing.scrollTop}px)`);
      check(
        Math.abs(landing.paneFromTop - landing.padding) <= 1,
        `„עיון בקבצים…” בראש חלון הגלילה: ${landing.paneFromTop}px מול ריפוד ${landing.padding}px`,
      );
      check(
        landing.templatesBottom <= landing.padding + 1,
        `רצועת התבניות יצאה מעל קו הגלילה (תחתיתה ב-${landing.templatesBottom}px)`,
      );
      check(
        landing.focus.includes('open-browse'),
        `המיקוד על „עיון בקבצים…” (נמדד „${landing.focus}”)`,
      );

      const measured = await page.cdp.evaluate(`(() => {
        const dialog = document.querySelector('.open-dialog');
        if (!dialog) return { open: false };
        const cards = [...document.querySelectorAll('.tpl-card')];
        const tops = new Set(cards.map((c) => Math.round(c.getBoundingClientRect().top)));
        const firstRow = cards.filter((c) => Math.round(c.getBoundingClientRect().top) === Math.min(...tops));
        const sheet = document.querySelector('.tpl-sheet').getBoundingClientRect();
        // גיליון A4 מול גיליון A5 — הדבר היחיד כאן שהעין כמעט אינה יכולה
        // לשפוט, ולכן הוא נמדד: הקנה הוא 148/210 = 0.704762 (ראו §5.4).
        const sheets = [...document.querySelectorAll('.pv-sheet')].map((el) => el.getBoundingClientRect().width);
        const a5Scale = sheets.length > 1 ? Math.min(...sheets) / Math.max(...sheets) : 0;
        const rows = [...document.querySelectorAll('.rec-row')];
        const meta = rows.map((r) => {
          const m = r.querySelector('.rec-meta');
          return m ? m.textContent.trim() : null;
        });
        const list = document.querySelector('.rec-list');
        return {
          open: true,
          columns: firstRow.length,
          a5Scale: Math.round(a5Scale * 10000) / 10000,
          sheetWidth: Math.round(sheet.width * 10) / 10,
          sheetHeight: Math.round(sheet.height * 10) / 10,
          dialogWidth: Math.round(dialog.getBoundingClientRect().width),
          bodyOverflowX: document.documentElement.scrollWidth - document.documentElement.clientWidth,
          rows: rows.length,
          lastPinned: document.querySelectorAll('.rec-row--last-pinned').length,
          meta,
          // פס הגלילה של הרשימה נגזר מכיווניות המיכל — ראו rec-list ב-
          // OpenDocumentDialog.vue. כאן נמדדת הכיווניות המחושבת עצמה.
          listDir: list ? getComputedStyle(list).direction : null,
          rowDir: rows.length > 0 ? getComputedStyle(rows[0]).direction : null,
        };
      })()`);

      check(
        measured.columns === viewport.expectColumns,
        `${viewport.expectColumns} כרטיסים בשורה הראשונה (נמדד ${measured.columns})`,
      );
      check(measured.sheetWidth > 30, `הגיליון מצויר ברוחב אמיתי: ${measured.sheetWidth}px`);
      check(
        Math.abs(measured.sheetHeight / measured.sheetWidth - 301 / 214) < 0.02,
        `יחס הגיליון הוא של דף: ${measured.sheetWidth}×${measured.sheetHeight}`,
      );
      check(
        Math.abs(measured.a5Scale - 148 / 210) < 0.005,
        `גיליון ה-A5 מוקטן ל-148/210 (נמדד ${measured.a5Scale})`,
      );
      check(measured.bodyOverflowX === 0, `אין גלילה אופקית בגוף הדף (${measured.bodyOverflowX}px)`);

      if (scenario.recents.length > 0) {
        check(
          measured.rows === scenario.recents.length,
          `שלוש שורות אחרונים (נמדד ${measured.rows})`,
        );
        check(measured.lastPinned === 1, 'קו אחד בדיוק מתחת למוצמדת האחרונה');
        check(
          measured.listDir === 'ltr' && measured.rowDir === 'rtl',
          'פס הגלילה בימין: רשימה ltr, שורה rtl',
        );
        check(
          measured.meta[2] === null,
          'שורה בלי גיל ובלי גודל אינה מציגה „0 בייט” (המטא אינו מרונדר)',
        );
      } else {
        check(measured.rows === 0, 'אין שורות אחרונים, ומצב הריק הוא שממלא את הפן');
      }

      await shoot(`open-dialog-${scenario.key}-${viewport.label}-open`);
      await closeDialog();

      /* ---------------------------------------------------------------- */
      /* Ctrl+N — אותו מסך בדיוק, נחיתה בראש                               */
      /* ---------------------------------------------------------------- */

      const openedNew = await openDialog('KeyN');
      check(openedNew, 'Ctrl+N פתח את אותו דיאלוג');
      if (openedNew) {
        const top = await page.cdp.evaluate(LANDING);
        check(top.scrollTop === 0, `Ctrl+N נוחת בראש המסך (${top.scrollTop}px)`);
        check(
          top.focus.includes('tpl-card'),
          `והמיקוד על כרטיס התבנית הראשון (נמדד „${top.focus}”)`,
        );
        await shoot(`open-dialog-${scenario.key}-${viewport.label}-new`);
        await closeDialog();
      }
    }
  } finally {
    page.close();
    rmSync(pagePath, { force: true });
  }
}

if (failures.length > 0) {
  console.error(`\nשער „פתח מסמך” נכשל ב-${failures.length} טענות.`);
  process.exit(1);
}
console.log(
  '\n✓ שער „פתח מסמך” עבר: הרשת, הגיליון, הרשימה, הכיווניות — ושתי הנחיתות, גם בלי היסטוריה.',
);
