/**
 * שער חי על מרכוז העמוד בזום — התלונה שהתיקון נולד ממנה היא ויזואלית, ולכן
 * גם הבדיקה מודדת פיקסלים ומצלמת מסך.
 *
 * שלושת המצבים שנמדדים, כולם בחלון קבוע 1440×900:
 *
 *   1. **הקטנה (60%)** — העמוד קטן מהמאגס וחייב להיות ממורכז. זה המצב
 *      שהתגלה ראשון: המנוע מציב `transform-origin: left top`, והתוכן ברח
 *      אל מחוץ למסך.
 *   2. **הגדלה שנכנסת (150%)** — העמוד גדל אך עדיין צר מהמאגס, וגם כאן הוא
 *      חייב להישאר ממורכז. בלי התיקון מרכזו סוטה ב-199px, כי `margin: 0 auto`
 *      של המנוע קורס לאפס ברגע שקופסת הפריסה צרה מהעמוד.
 *   3. **הגדלה שגולשת (300%)** — העמוד רחב מהמאגס. „ממורכז” כבר אינו הגדרה,
 *      והדרישה מתפצלת לשתיים: כל העמוד נגיש בגלילה (הפריסה מצמידה את העודף
 *      לצד ה-end של המיכל), והגלילה נחה על תחילת השורה מיד אחרי שינוי הזום
 *      (הצמדה בקוד, ראו engine/zoom-center.ts). מיכל הגלילה הוא ltr — הצהרה
 *      על צד פס הגלילה, ראו styles/shell.css — ולכן שני אלה אינם אותו דבר.
 *
 * הרקע המלא בהערות של engine/zoom-center.ts ו-styles/shell.css.
 *
 *   npm run build && node scripts/zoom-center-probe.mjs
 */
import { readFileSync, writeFileSync, existsSync, rmSync, mkdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openPage, requireChrome } from './cdp.mjs';

const DIST = join(dirname(fileURLToPath(import.meta.url)), '..', 'dist');
const TMP = join(dirname(fileURLToPath(import.meta.url)), '..', 'tmp');
const OBSERVE_MS = 40_000;

/** חלון קבוע: הענף „נכנס/גולש” תלוי ברוחב, ובדיקה שאינה קובעת אותו אינה בדיקה. */
const VIEWPORT = { width: 1440, height: 900 };

/** רוחב פס גלילה ושולי עימוד. מעבר לכך זו כבר „בריחה הצידה”. */
const TOLERANCE = 24;

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
const path = join(DIST, 'center-tmp.html');
writeFileSync(path, html.slice(0, latchEnd) + STUB + html.slice(latchEnd));

/**
 * מודדת את העמוד מול תיבת התוכן של מיכל הגלילה, ובודקת נגישות בקצוות: גלילה
 * עד הסוף לכל צד ומדידה מה נראה שם. גלישה מעבר לקצה ההתחלה של מיכל גלילה
 * אינה נכנסת לאזור הגלילה, ולכן „נגיש” אינו נובע מ„קיים”.
 *
 * המדידה נייטרלית לכיווניות בכוונה: `clientLeft` הוא רוחב הגבול השמאלי ועוד
 * רוחב פס הגלילה **כשהוא בשמאל**, ואפס כשהוא בימין — ולכן הוא, ולא הנחה על
 * הצד, מה שמגדיר איפה תיבת התוכן מתחילה. הנחה קשיחה כאן הייתה מדווחת סטייה
 * ברוחב פס הגלילה בכל פעם שהצד מתחלף, ומסתירה סטייה אמיתית באותו סדר גודל.
 */
const MEASURE = `(async () => {
  const settle = (ms) => new Promise((r) => setTimeout(r, ms));
  const q = (s) => document.querySelector(s);
  const host = q('.editor-stack__host');
  const page = q('.superdoc-page');
  const wrapper = q('.presentation-editor');
  if (!host || !page || !wrapper) return null;

  const edges = () => {
    const b = page.getBoundingClientRect();
    return { left: Math.round(b.left), right: Math.round(b.right), width: Math.round(b.width) };
  };

  // תיבת התוכן של המיכל, בלי להניח באיזה צד פס הגלילה יושב.
  const box = host.getBoundingClientRect();
  const left = Math.round(box.left + host.clientLeft);
  const view = { left, right: left + host.clientWidth };
  const direction = getComputedStyle(host).direction;
  // כיווניות **המסמך**, לא של המיכל: תחילת שורה עברית היא הקצה הימני של
  // העמוד בכל מקרה. התכונה נכתבת ב-engine/document-defaults.ts.
  const documentDirection = document.documentElement.getAttribute('data-document-direction');

  // rest נמדד לפני שנוגעים בגלילה: זה מה שהמשתמש רואה מיד אחרי שינוי הזום,
  // כלומר גם מה שההצמדה של zoom-center.ts קבעה.
  const rest = edges();
  const scroll = { left: Math.round(host.scrollLeft), max: host.scrollWidth - host.clientWidth };

  // שני הקצוות של אזור הגלילה, בלי להניח לאיזה סימן scrollLeft נע: ב-ltr
  // הוא 0..max, וב-rtl 0..-max.
  host.scrollLeft = 99999; await settle(200); const atMax = edges();
  host.scrollLeft = -99999; await settle(200); const atMin = edges();
  host.scrollLeft = 0; await settle(200);

  const cs = getComputedStyle(wrapper);
  return {
    view,
    direction,
    documentDirection,
    rest,
    scroll,
    extremes: [atMax, atMin],
    origin: cs.transformOrigin,
    wrapperLayoutWidth: wrapper.offsetWidth,
  };
})()`;

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

let page;
try {
  page = await openPage(`file://${path}`, { label: 'center' });
  await page.cdp.send('Emulation.setDeviceMetricsOverride', {
    ...VIEWPORT,
    deviceScaleFactor: 1,
    mobile: false,
  });
  if (!(await page.cdp.evaluate(READY))) throw new Error('האתחול לא הושלם');

  for (const percent of [100, 60, 150, 300]) {
    await page.cdp.evaluate(`window.__otzariaEditor.superdoc.setZoom(${percent})`);
    await new Promise((r) => setTimeout(r, 1500));
    await screenshot(page.cdp, join(TMP, `zoom-center-${percent}.png`));

    const m = await page.cdp.evaluate(MEASURE);
    if (!m) {
      console.error(`✗ ${percent}%: לא נמדד עמוד`);
      failures += 1;
      continue;
    }
    console.log(`\n${percent}%: ${JSON.stringify(m)}`);

    const viewWidth = m.view.right - m.view.left;
    const viewCenter = (m.view.right + m.view.left) / 2;
    const pageCenter = (m.rest.right + m.rest.left) / 2;

    if (m.rest.width <= viewWidth) {
      const drift = Math.abs(pageCenter - viewCenter);
      check(drift <= TOLERANCE, `${percent}%: העמוד נכנס למאגס וממורכז (סטייה ${Math.round(drift)}px)`);
    } else {
      // גולש. שתי דרישות נפרדות, וקל לבלבל ביניהן:
      //
      //   1. **נגישות** — כל העמוד חייב להיכנס לאזור הגלילה. גלישה מעבר לקצה
      //      ההתחלה של מיכל הגלילה אינה נגישה לצמיתות, ולכן ההצמדה שבנוסחה
      //      (shell.css) מאחדת את כל העודף בצד ה-end.
      //   2. **נקודת פתיחה** — מה שנראה מיד אחרי שינוי הזום צריך להיות תחילת
      //      השורה. במיכל ltr זהו הקצה הימני, ושם `scrollLeft` המקסימלי, ולכן
      //      זו הצמדה שנעשית בקוד (engine/zoom-center.ts) ולא בפריסה.
      // „נגיש” = קיימת עמדת גלילה שבה הקצה הזה מגיע לקצה המאגס. לכן מינימום
      // על שני קצות אזור הגלילה, ולא איחוד שלהם: איחוד היה עובר גם על עמוד
      // שחצי ממנו מחוץ להישג יד, שהרי הקצה הרחוק „קיים” באחת העמדות.
      const closest = (edge) => Math.min(...m.extremes.map((e) => Math.abs(e[edge] - m.view[edge])));
      check(
        closest('right') <= TOLERANCE,
        `${percent}%: הקצה הימני של העמוד נגיש בגלילה (סטייה ${Math.round(closest('right'))}px)`,
      );
      check(
        closest('left') <= TOLERANCE,
        `${percent}%: הקצה השמאלי של העמוד נגיש בגלילה (סטייה ${Math.round(closest('left'))}px)`,
      );
      // נקודת הפתיחה: מה שנראה מיד אחרי שינוי הזום. תחילת שורה עברית היא הקצה
      // הימני של העמוד — תכונה של המסמך ולא של המיכל — ומיכל ltr נח בקצה
      // השמאלי, ולכן זו הצמדה שנעשית בקוד (engine/zoom-center.ts).
      const lineStart = m.documentDirection === 'rtl' ? 'right' : 'left';
      const opening = Math.abs(m.rest[lineStart] - m.view[lineStart]);
      check(
        opening <= TOLERANCE,
        `${percent}%: אחרי שינוי הזום הגלילה נחה על תחילת השורה (סטייה ${Math.round(opening)}px)`,
      );
    }

    // ה-origin הנכון הוא פינת ההתחלה של קופסת הפריסה (ראו shell.css) — במיכל
    // ltr הקצה השמאלי, כלומר אפס. אין כאן כלל CSS: זה מה שהמנוע כותב בעצמו,
    // וכל חישוב המרכוז נשען עליו. סטייה כאן = הרגע להחזיר את ה-override.
    const expected = m.direction === 'rtl' ? m.wrapperLayoutWidth : 0;
    const actual = Number.parseFloat(m.origin);
    check(
      !Number.isNaN(actual) && Math.abs(actual - expected) <= 2,
      `${percent}%: transform-origin בפינת ההתחלה של קופסת הפריסה (${m.origin} מול ${expected}px)`,
    );
  }
} catch (error) {
  console.error(`center-probe נכשל: ${error.message}`);
  failures += 1;
} finally {
  page?.close();
  rmSync(path, { force: true });
}

process.exit(failures ? 1 : 0);
