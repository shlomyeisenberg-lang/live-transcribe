/**
 * שער על המרווח האנכי שסביב ערימת העמודים — „המשטח האפור” של Word.
 *
 * התלונה שהתיקון נולד ממנה היא ויזואלית („העמוד מגיע עד ראש המסך, נראה
 * חתוך”), ולכן גם הבדיקה מודדת פיקסלים ומצלמת מסך. מה שנמדד:
 *
 *   1. **שלושת המרווחים זהים** — מעל העמוד הראשון, בין עמוד לעמוד, ומתחת
 *      לאחרון. זו הטענה המרכזית, והיא זו שמחזיקה את המספר הקשיח בכנות:
 *      ה-24px ב-styles/shell.css אינו ערך שהמצאנו אלא בדיוק ה-`gap` שהמנוע
 *      כותב בין העמודים, והיום שבו המנוע ישנה אותו ייפול כאן — ולא אצל
 *      המשתמש כמרווח שאינו מסתדר.
 *
 *   2. **בכל אחוז זום.** המרווחים יושבים בתוך ה-`scale()` של המנוע, ולכן
 *      שלושתם אמורים לגדול ולהתכווץ יחד. מדידה באחוז אחד בלבד הייתה מפספסת
 *      בדיוק את סוג הרגרסיה שהגיע לכאן פעם — ריפוד שנקבע במיכל שאינו מוקטן,
 *      ונראה נכון ב-100% ושבור ב-300%.
 *
 *   3. **המרווח אינו אפס.** נשמע מיותר, אבל זה בדיוק המצב שלפני התיקון:
 *      `gap` בין עמודים היה קיים, ובשני הקצוות היה `0px`. שוויון בלבד היה
 *      עובר גם על „אפס בכל מקום”.
 *
 *   4. **ההדפסה מאפסת אותו.** בגיליון אין „משטח אפור”, ו-24px מעל העמוד
 *      הראשון מול `@page { margin: 0 }` גולשים לגיליון נוסף. נמדד במדיית
 *      print אמיתית (`Emulation.setEmulatedMedia`) ולא בקריאת CSS.
 *
 * המסמך שנמדד עליו נבנה כאן: מסמך חדש הוא עמוד אחד, ובעמוד אחד אין „בין
 * עמודים”. המילוי נעשה דרך ה-API הציבורי (`doc.insert`) ולא דרך ה-DOM.
 *
 * הרקע המלא — styles/shell.css (הכלל וההנמקה) ו-styles/print.css (האיפוס).
 *
 *   npm run build && node scripts/page-gutter-probe.mjs
 */
import { readFileSync, writeFileSync, existsSync, rmSync, mkdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage, requireChrome, sleep } from './cdp.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const DIST = join(ROOT, 'dist');
const TMP = join(ROOT, 'tmp');
const OBSERVE_MS = 40_000;

/** חלון קבוע: העימוד תלוי ברוחב, ובדיקה שאינה קובעת אותו אינה בדיקה. */
const VIEWPORT = { width: 1440, height: 900 };

/** המרווח שהכלל ב-styles/shell.css קובע, ב-100%. שאר האחוזים נגזרים ממנו. */
const GUTTER_PX = 24;

/** עיגול תת-פיקסלי בלבד. מעבר לכך זה כבר מרווח שאינו מה שנקבע. */
const TOLERANCE = 1.5;

if (!existsSync(join(DIST, 'index.html'))) {
  console.error('dist/index.html אינו קיים — הריצו npm run build תחילה');
  process.exit(1);
}
requireChrome();
mkdirSync(TMP, { recursive: true });

const STUB = `<script>
  window.Otzaria = {
    call: function (method) {
      if (method === 'app.getInfo') return Promise.resolve({ success: true, data: { version: '9', platform: 'p' }, error: null });
      if (method === 'app.getTheme') return Promise.resolve({ success: true, data: { mode: 'light', colorScheme: {}, typography: {} }, error: null });
      return Promise.resolve({ success: false, data: null, error: { message: 'no' } });
    },
    on: function () {},
    off: function () {}
  };
</script>`;

const html = readFileSync(join(DIST, 'index.html'), 'utf8');
const latchEnd = html.indexOf('</script>') + '</script>'.length;
const path = join(DIST, 'gutter-tmp.html');
writeFileSync(path, html.slice(0, latchEnd) + STUB + html.slice(latchEnd));

const READY = `(async () => {
  const settle = (ms) => new Promise((r) => setTimeout(r, ms));
  const deadline = Date.now() + ${OBSERVE_MS};
  while (Date.now() < deadline) {
    if (document.querySelector('.word-tab-strip') && window.__otzariaEditor) break;
    await settle(250);
  }
  const sd = window.__otzariaEditor && window.__otzariaEditor.superdoc;
  if (!sd) return false;
  while (Date.now() < deadline) {
    if (sd.ui && sd.ui.commands.get('zoom').getState().enabled) break;
    await settle(250);
  }
  await settle(1500);
  return true;
})()`;

/** שלושה עמודים לפחות: שניים בשביל „בין”, שלישי כדי שזה לא יהיה מקרה. */
const FILL = `(async () => {
  const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
  const paragraphs = [];
  for (let i = 0; i < 130; i++) {
    paragraphs.push('<p>שורה ' + (i + 1) + ' — טקסט מילוי לשער המרווחים.</p>');
  }
  await doc.insert({ value: paragraphs.join(''), type: 'html' });
  await new Promise((r) => setTimeout(r, 4000));
  return document.querySelectorAll('.superdoc-page').length;
})()`;

/**
 * שלושת המרווחים, כולם מול **מיכל העמודים עצמו** ולא מול מיכל הגלילה.
 *
 * זו ההבחנה שמאפשרת למדוד בכל אחוז זום באותה נוסחה: המיכל נושא את ה-`scale()`
 * ולכן גם הריפוד שלו וגם ה-`gap` שבתוכו מוקטנים באותו יחס, ושניהם נמדדים כאן
 * במרחב המוקטן. מדידה מול מיכל הגלילה הייתה מערבבת פנימה את גובה אזור הגלילה,
 * שאינו מתכווץ בהקטנה (נמדד: ב-60% הוא נשאר גובה הפריסה הלא-מוקטן) — כלומר
 * „המרווח מתחת” היה יוצא מאות פיקסלים ולא מה שנצבע.
 */
const MEASURE = `(() => {
  const wrapper = document.querySelector('.presentation-editor');
  if (!wrapper) return null;
  const pages = Array.from(document.querySelectorAll('.superdoc-page'))
    .map((page) => ({ index: Number(page.dataset.pageIndex), box: page.getBoundingClientRect() }))
    .filter((page) => Number.isInteger(page.index))
    .sort((a, b) => a.index - b.index);
  if (pages.length < 2) return { pageCount: pages.length };

  const box = wrapper.getBoundingClientRect();
  const gaps = [];
  for (let i = 1; i < pages.length; i += 1) gaps.push(pages[i].box.top - pages[i - 1].box.bottom);

  return {
    pageCount: pages.length,
    above: pages[0].box.top - box.top,
    below: box.bottom - pages[pages.length - 1].box.bottom,
    gaps,
    zoom: Number.parseFloat(getComputedStyle(wrapper).transform.replace('matrix(', '')) || 1,
  };
})()`;

/** הריפוד המחושב במדיית ההדפסה. `padding-block` נקרא בשני קצותיו. */
const PRINT_PADDING = `(() => {
  const wrapper = document.querySelector('.presentation-editor');
  if (!wrapper) return null;
  const style = getComputedStyle(wrapper);
  return { top: style.paddingTop, bottom: style.paddingBottom };
})()`;

async function screenshot(cdp, file) {
  const response = await cdp.send('Page.captureScreenshot', { format: 'png' });
  const data = response?.result?.data;
  if (!data) {
    console.error(`לא ניתן לצלם (${file})`);
    return;
  }
  writeFileSync(file, Buffer.from(data, 'base64'));
  console.log(`📷 ${file}`);
}

let failures = 0;

function check(ok, message) {
  console.log(`${ok ? '✓' : '✗'} ${message}`);
  if (!ok) failures += 1;
}

const round = (value) => Math.round(value * 10) / 10;

let page;
try {
  page = await openPage(`file://${path}`, { label: 'gutter' });
  await page.cdp.send('Emulation.setDeviceMetricsOverride', {
    ...VIEWPORT,
    deviceScaleFactor: 1,
    mobile: false,
  });
  if (!(await page.cdp.evaluate(READY))) throw new Error('האתחול לא הושלם');

  const pageCount = await page.cdp.evaluate(FILL);
  check(pageCount >= 3, `המסמך שנמדד נפרס ל-${pageCount} עמודים`);

  for (const percent of [100, 60, 150, 300]) {
    await page.cdp.evaluate(`window.__otzariaEditor.superdoc.setZoom(${percent})`);
    await sleep(1500);
    if (percent === 100) await screenshot(page.cdp, join(TMP, 'page-gutter-100.png'));

    const measured = await page.cdp.evaluate(MEASURE);
    if (!measured || measured.pageCount < 2) {
      console.error(`✗ ${percent}%: אין מספיק עמודים למדידה`);
      failures += 1;
      continue;
    }

    // הציפייה נגזרת מהזום שנמדד על המיכל, ולא מהאחוז שביקשנו: ה-`scale()` הוא
    // של המנוע, ושער שמניח אותו במקום לקרוא אותו מודד את ההנחה שלנו.
    const expected = GUTTER_PX * measured.zoom;
    const all = [measured.above, ...measured.gaps, measured.below];
    console.log(
      `\n${percent}% (scale ${round(measured.zoom)}): מעל ${round(measured.above)}px, ` +
        `בין ${measured.gaps.map(round).join('/')}px, מתחת ${round(measured.below)}px`,
    );

    check(
      measured.above > TOLERANCE,
      `${percent}%: יש מרווח מעל העמוד הראשון (${round(measured.above)}px)`,
    );
    check(
      all.every((value) => Math.abs(value - expected) <= TOLERANCE),
      `${percent}%: מעל, בין ומתחת זהים ל-${round(expected)}px`,
    );
  }

  // ההדפסה. מדיית print אמיתית ולא קריאת גיליון: מה שנמדד הוא מה שהמדפסת
  // מקבלת, אחרי כל שרשרת הקדימויות בין הכלל שלנו לכלל של המנוע.
  await page.cdp.send('Emulation.setEmulatedMedia', { media: 'print' });
  await sleep(600);
  const printPadding = await page.cdp.evaluate(PRINT_PADDING);
  check(
    printPadding?.top === '0px' && printPadding?.bottom === '0px',
    `בהדפסה הריפוד מאופס (${printPadding?.top} / ${printPadding?.bottom})`,
  );
  await page.cdp.send('Emulation.setEmulatedMedia', { media: '' });
} catch (error) {
  console.error(`page-gutter-probe נכשל: ${error.message}`);
  failures += 1;
} finally {
  page?.close();
  rmSync(path, { force: true });
}

process.exit(failures ? 1 : 0);
