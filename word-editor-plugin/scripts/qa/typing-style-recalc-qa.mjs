/**
 * שער QA: הקלדה אינה מחשבת מחדש את הסגנון של כל עץ המסמך.
 *
 * מה שנמדד, ולמה זה השער: ב-9.2026 דווח „לוקח כמה שניות עד שהתו מופיע”.
 * ה-trace של Chrome הראה שעל כל הקשה Blink חישב סגנון מחדש (`UpdateLayoutTree`)
 * ל-**כל** עץ המסמך — 2,748 אלמנטים במסמך של חמישה עמודים, 60–115ms לחישוב,
 * 42 חישובים ב-50 הקשות — בגלל כלל `:has()` יחיד שלנו שעוגנו ישב בתוך
 * `.superdoc` (engine-chrome.css, באנר `edit-rejected`). אחרי התיקון: 0.
 * ההסבר המלא ליד הכלל, וגדר סטטית ב-tests/unit/css-hygiene.test.ts.
 *
 * השער הזה הוא הגדר הדינמית: הוא מקליד במסמך של כמה עמודים תחת trace של
 * ה-renderer, וסופר חישובי סגנון שגודלם הוא סדר-הגודל של המסמך כולו. הוא
 * תופס גם מה שהסטטי אינו רואה — כלל ב-`.vue` שנוסח אחרת, או שדרוג מנוע שמשנה
 * את גיליון המנוע — כי הוא מודד את התוצאה ולא את המקור.
 *
 * למה ספירה ולא זמן: המכונה רועשת (ראו scripts/cdp.mjs) וזמן משתנה פי כמה בין
 * ריצות; מספר האלמנטים בחישוב הוא תכונה של ה-DOM וה-CSS, ואינו תלוי בעומס.
 * הסף (`WHOLE_TREE_ELEMENTS`) הוא מחצית ממה שנמדד למסמך של חמישה עמודים,
 * ורחוק בשני סדרי גודל מחישוב של הקשה תקינה (חציון 7 אלמנטים).
 *
 * הרצה: node scripts/qa/typing-style-recalc-qa.mjs   (יציאה 9391)
 */
import { openApp, createReport, sleep } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9391);
const report = createReport('הקלדה — אין חישוב סגנון של כל המסמך', { strict: true });

/** חישוב שנוגע ביותר מזה נחשב „כל המסמך”. */
const WHOLE_TREE_ELEMENTS = 1000;
/** כמה חישובים כאלה מותרים במהלך ההקלדה. אפס נמדד אחרי התיקון; אחד — רעש. */
const MAX_WHOLE_TREE_RECALCS = 1;
const KEYSTROKES = 40;
const MIN_PAGES = 3;

/**
 * חיבור CDP שני, לאותו דף, שיודע לקבל **אירועים**: `Tracing.dataCollected`
 * מגיע כאירוע, ו-`scripts/cdp.mjs` מטפל רק בתשובות לפי id. CDP מרשה כמה
 * לקוחות לאותו יעד, ולכן זה אינו משבש את ההרצה עצמה.
 */
async function attachTracer(port, urlPart) {
  const list = await (await fetch(`http://127.0.0.1:${port}/json/list`)).json();
  const page = list.find((t) => t.type === 'page' && t.url.includes(urlPart));
  if (!page) throw new Error('דף השער לא נמצא ברשימת היעדים');
  const socket = new WebSocket(page.webSocketDebuggerUrl);
  await new Promise((resolve, reject) => {
    socket.addEventListener('open', resolve, { once: true });
    socket.addEventListener('error', () => reject(new Error('חיבור ה-trace נכשל')), { once: true });
  });
  let nextId = 0;
  const pending = new Map();
  const collected = [];
  let complete = null;
  socket.addEventListener('message', (event) => {
    const message = JSON.parse(event.data);
    if (message.id === undefined) {
      if (message.method === 'Tracing.dataCollected') collected.push(...message.params.value);
      if (message.method === 'Tracing.tracingComplete') complete?.();
      return;
    }
    const settle = pending.get(message.id);
    if (settle) {
      pending.delete(message.id);
      settle(message);
    }
  });
  const send = (method, params) =>
    new Promise((resolve) => {
      const id = ++nextId;
      pending.set(id, resolve);
      socket.send(JSON.stringify({ id, method, params }));
    });
  return {
    start: () =>
      send('Tracing.start', {
        categories: 'devtools.timeline,disabled-by-default-devtools.timeline,blink.user_timing',
        transferMode: 'ReportEvents',
      }),
    async stop() {
      const done = new Promise((resolve) => {
        complete = resolve;
      });
      await send('Tracing.end');
      await Promise.race([done, sleep(60_000)]);
      socket.close();
      return collected;
    },
  };
}

const app = await openApp({ name: 'typing-style-recalc', port: PORT });

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1400, height: 900, deviceScaleFactor: 1, mobile: false });
  await sleep(800);
  await app.caret(0);

  /* ---- מסמך של כמה עמודים: הכשל תלוי בגודל העץ, ומסמך ריק אינו מודד אותו ----
   * דרך ה-Document API ולא בהקלדה: זה תנאי המדידה, לא מה שנמדד (כמו
   * column-selection-probe.mjs). 90 פסקאות של ~130 תווים הן כ-5 עמודים. */
  await app.js(`(async () => {
    const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    const filler = 'בראשית ברא אלהים את השמים ואת הארץ והארץ היתה תהו ובהו וחשך על פני תהום ורוח אלהים מרחפת על פני המים ויאמר אלהים יהי אור';
    const paras = [];
    for (let i = 1; i <= 90; i++) paras.push('פסקה ' + i + ' — ' + filler);
    await doc.insert({ value: paras.join('\\n\\n'), type: 'markdown' });
  })()`);
  await sleep(3000);
  await app.caret(2);
  const pages = Number(await app.js('document.querySelectorAll("[data-page-index]").length'));
  const elements = Number(await app.js('window.__otzariaEditor.container.querySelectorAll("*").length'));
  report[pages >= MIN_PAGES ? 'pass' : 'fail'](
    'מסמך הבדיקה גדול דיו',
    `${pages} עמודים, ${elements} אלמנטים בעץ המסמך`,
    pages >= MIN_PAGES ? '' : `נדרשים לפחות ${MIN_PAGES} עמודים כדי שהמדידה תהיה משמעותית`,
  );

  /* ---- הקלדה תחת trace ---- */
  const tracer = await attachTracer(PORT, '__qa-typing-style-recalc');
  await tracer.start();
  await sleep(300);
  await app.js('performance.mark("qa:typing-start")');
  await app.type('אבגדהוזחטיכלמנסעפצקרשת אבגדהוזחטיכלמנסעפצקרשת'.slice(0, KEYSTROKES), 80);
  await sleep(600);
  await app.js('performance.mark("qa:typing-end")');
  await sleep(300);
  const events = await tracer.stop();

  const mark = (name) => events.find((e) => e.name === name && e.cat?.includes('blink.user_timing'))?.ts;
  const from = mark('qa:typing-start');
  const to = mark('qa:typing-end');
  if (from === undefined || to === undefined) {
    report.fail('ה-trace נאסף', `אירועים: ${events.length}`, 'סימני ההתחלה/סיום של ההקלדה לא נמצאו ב-trace');
  } else {
    const recalcs = events.filter((e) => e.name === 'UpdateLayoutTree' && e.ph === 'X' && e.ts >= from && e.ts <= to);
    const counts = recalcs.map((e) => e.args?.elementCount ?? 0).sort((a, b) => a - b);
    const whole = recalcs.filter((e) => (e.args?.elementCount ?? 0) >= WHOLE_TREE_ELEMENTS);
    const wholeMs = Math.round(whole.reduce((sum, e) => sum + (e.dur ?? 0), 0) / 1000);
    const median = counts.length ? counts[Math.floor(counts.length / 2)] : 0;
    const max = counts.length ? counts[counts.length - 1] : 0;

    report[recalcs.length > 0 ? 'pass' : 'fail'](
      'ההקלדה אכן גרמה לחישובי סגנון (המדידה חיה)',
      `${recalcs.length} חישובים ב-${KEYSTROKES} הקשות; חציון ${median} אלמנטים, מקסימום ${max}`,
      recalcs.length > 0 ? '' : 'אפס חישובים — ההקלדה לא הגיעה למסמך, או שה-trace ריק',
    );
    report[whole.length <= MAX_WHOLE_TREE_RECALCS ? 'pass' : 'fail'](
      `אין חישוב סגנון של כל המסמך (≥ ${WHOLE_TREE_ELEMENTS} אלמנטים) בזמן הקלדה`,
      `${whole.length} חישובים כאלה, ${wholeMs}ms יחד`,
      whole.length <= MAX_WHOLE_TREE_RECALCS
        ? ''
        : 'כל הקשה מחשבת מחדש את כל עץ המסמך — ראו engine-chrome.css, הערת הכלל של באנר edit-rejected',
    );
  }
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
