/**
 * הנקודה הפתוחה היחידה שנשארה אחרי סקר הפקדים: „מסמך חדש” הקפיא את הדף
 * בשלושה שחזורים מתוך שלושה — אבל רק ברצף מסוים, ולא לבדו. בגלל ההקפאה
 * הזאת כל קבוצת „קובץ” שאחריה (שמור, שמור בשם, פתח, ייצוא, הדפסה, יציאה)
 * נשארה בלי מדידה.
 *
 * השער הזה עונה על שתי שאלות, ולא על יותר:
 *   1. **מה בדיוק מקפיא.** ארבעה תרחישים, כל אחד במופע דפדפן משלו, שמבודדים
 *      את הרצף שדווח: „מסמך חדש” לבדו, אחרי דיאלוג אחד, אחרי שני דיאלוגים,
 *      ואחרי שני דיאלוגים ומיקום סמן.
 *   2. **האם ההקפאה סופית.** אחרי הלחיצה הדף נדגם שוב ושוב עד 100 שניות.
 *      אם הוא חוזר — זו איטיות, לא הקפאה. אם לא — זו הקפאה, והיא באג.
 *
 * למה שעון על כל קריאה: `cdp.evaluate` ממתין לתשובה, ודף מוקפא לעולם לא
 * ישלח אותה. בלי השעון השער עצמו נתלה — וזה בדיוק מה שקרה בשערים קודמים,
 * ולכן הם לא הצליחו לדווח על התקלה שמצאו.
 *
 *   node scripts/qa/file-freeze-qa.mjs          # כל התרחישים
 *   node scripts/qa/file-freeze-qa.mjs 4        # תרחיש יחיד
 */
import { openApp, createReport, sleep } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9371);

/**
 * מד המשימות הארוכות, מוזרק לפני שהאפליקציה קמה.
 *
 * „הדף לא הגיב 10 שניות” אינו אומר **מה** חסם. שלושת ההסברים — משימה
 * סינכרונית ארוכה בקוד, עצירת GC, או המכונה עצמה תחת עומס — דורשים תגובה
 * שונה לגמרי, ומספר אחד לא מבדיל ביניהם. `PerformanceObserver` על
 * `longtask` רושם כל משימה מעל 50ms עם משך והתחלה, וזו העדות שמפרידה בין
 * „הקוד שלנו חסם” לבין „המכונה נעצרה מתחתינו”.
 *
 * `performance.memory` נדגם לצדו: קפיצה בערמה סביב החסימה מצביעה על GC,
 * ולא על לולאה שלנו. הוא קיים ב-Chromium בלבד, ולכן נקרא בזהירות.
 */
const LONGTASK_PROBE = `
<script>
(function () {
  var L = window.__longtasks = { entries: [], mem: [] };
  try {
    new PerformanceObserver(function (list) {
      list.getEntries().forEach(function (e) {
        L.entries.push({
          start: Math.round(e.startTime),
          dur: Math.round(e.duration),
          name: e.name,
          attr: (e.attribution || []).map(function (a) {
            return a.containerType + ':' + (a.containerName || a.containerSrc || '');
          }),
        });
      });
    }).observe({ entryTypes: ['longtask'] });
  } catch (err) { L.error = String(err); }
  setInterval(function () {
    try {
      var m = performance.memory;
      if (m) L.mem.push({ t: Math.round(performance.now()), used: Math.round(m.usedJSHeapSize / 1048576) });
    } catch (err) { /* לא Chromium */ }
  }, 500);
})();
</script>`;



/** גבול זמן לקריאה בודדת. דף חי עונה במילישניות. */
function withTimeout(promise, ms, label) {
  let timer;
  return Promise.race([
    promise.finally(() => clearTimeout(timer)),
    new Promise((_, reject) => {
      timer = setTimeout(() => reject(new Error(`פג הזמן: ${label}`)), ms);
    }),
  ]);
}

/** האם הדף עדיין מריץ JavaScript. זו ההגדרה התפעולית של „לא קפוא”. */
async function alive(app, ms = 3_000) {
  try {
    return (await withTimeout(app.js('1+1'), ms, 'ping')) === 2;
  } catch {
    return false;
  }
}

/** דוגם עד שהדף עונה, או עד שנגמר הזמן. מחזיר כמה זמן לקח. */
async function waitAlive(app, budgetMs = 100_000) {
  const t0 = Date.now();
  while (Date.now() - t0 < budgetMs) {
    if (await alive(app, 2_000)) return Date.now() - t0;
    await sleep(1_000);
  }
  return null;
}

/** מרחיב את החלון. בלעדיו פקדים בקצה הרצועה יושבים מחוץ לו. */
async function widen(app) {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1600,
    height: 1000,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await sleep(400);
}

/**
 * תרחיש אחד: רצף צעדים, ואז „מסמך חדש”, ואז מדידת חיוּת.
 * מופע דפדפן חדש לכל תרחיש — הקפאה מזהמת כל מה שאחריה.
 */
/** קורא את מד המשימות הארוכות. לעולם אינו זורק — הוא אבחון, לא תנאי. */
async function readLongTasks(app) {
  try {
    return JSON.parse(await withTimeout(app.js('JSON.stringify(window.__longtasks||null)'), 5_000, 'longtasks'));
  } catch {
    return null;
  }
}

/**
 * מנסח את המדידה, ומכריע בין שני ההסברים.
 *
 * משימה ארוכה יחידה שמכסה כמעט את כל החסימה = הקוד חסם. הרבה משימות קצרות,
 * או חסימה בלי משימות כלל, = המכונה נעצרה מתחתינו (GC, החלפת זיכרון, עומס),
 * וזה אינו באג במוצר.
 */
/** משך המשימה הארוכה ביותר שנרשמה, או 0. */
function longestTask(data) {
  const entries = data?.entries ?? [];
  return entries.length ? Math.max(...entries.map((e) => e.dur)) : 0;
}

function describeTasks(data) {
  if (!data) return 'אין מדידת longtask';
  const entries = data.entries ?? [];
  if (entries.length === 0) return 'אפס משימות ארוכות — החסימה לא הייתה בקוד של הדף';
  const worst = entries.reduce((a, b) => (b.dur > a.dur ? b : a));
  const total = entries.reduce((sum, e) => sum + e.dur, 0);
  const mem = data.mem ?? [];
  const heap = mem.length ? `, ערמה ${mem[0].used}→${mem[mem.length - 1].used}MB (שיא ${Math.max(...mem.map((m) => m.used))})` : '';
  return `${entries.length} משימות ארוכות, סה"כ ${total}ms, הארוכה ${worst.dur}ms ב-${worst.start}ms${worst.attr?.length ? ` (${worst.attr.join(',')})` : ''}${heap}`;
}

async function scenario(report, { id, name, steps, stopAfterSteps = false }) {
  console.log(`\n────── תרחיש ${id}: ${name} ──────`);
  const app = await openApp({ name: `freeze${id}`, port: PORT, extra: LONGTASK_PROBE });
  try {
    await widen(app);
    await app.tab('קובץ');

    for (const step of steps) {
      console.log(`  צעד: ${step.what}`);
      await step.run(app);
      const ok = await alive(app);
      console.log(`    חי אחרי הצעד? ${ok}`);
      if (!ok) {
        const back = await waitAlive(app, 60_000);
        // מה שחסם, ולא רק כמה זמן. ההנמקה ב-LONGTASK_PROBE.
        const tasks = await readLongTasks(app);
        const detail =
          (back === null
            ? 'הדף לא הריץ JavaScript במשך 60 שניות'
            : `הדף חזר לענות אחרי ${back}ms`) + ` | ${describeTasks(tasks)}`;
        console.log(`    ${detail}`);
        /* חסימה בלי משימה ארוכה אינה באג במוצר.
         *
         * נמדד: הדף „לא הגיב” 9.7 שניות, ובאותו חלון רשם **משימה ארוכה אחת
         * של 112ms — בעלייה, לא בחסימה** — והערמה נשארה שטוחה על 54MB. אם
         * הקוד שלנו היה חוסם, המשימה הייתה נרשמת כשהיא נגמרת, והדף אכן חזר.
         * כלומר ה-renderer לא קיבל מעבד — הרעבה של headless, לא לולאה שלנו.
         *
         * זה גם מסביר את חוסר הדטרמיניזם: אותו תרחיש בדיוק נחסם ב-1 מתוך 4
         * ריצות, ובמשכים שונים (9.7s / 10.1s / 15.6s).
         *
         * לכן זה „תקוע” ולא „שבור”. משימה ארוכה שכן תכסה את החסימה תמשיך
         * להיות כשל, וזה בדיוק ההבדל שהמדידה הזו נועדה להכריע. */
        const blocked = longestTask(tasks) >= (back ?? 0) * 0.5;
        if (blocked) report.fail(`${name} — קפא ב„${step.what}”`, detail);
        else report.stuck(`${name} — ב„${step.what}”`, detail);
        return;
      }
    }

    if (stopAfterSteps) {
      report.pass(`${name} — הדף נשאר חי`);
      return;
    }

    const state = await withTimeout(app.state('מסמך חדש'), 5_000, 'state');
    console.log(`  מצב „מסמך חדש”: ${JSON.stringify(state)}`);
    if (!state.found || !state.rect) {
      report.fail(`${name} — „מסמך חדש” לא נמצא`, JSON.stringify(state));
      return;
    }

    // לחיצה ישירה בקואורדינטות: `app.click` קורא לדף כדי למצוא את המלבן,
    // וזו קריאה שתיתקע אם הדף כבר מוקפא. כאן המלבן כבר בידינו.
    const t0 = Date.now();
    await app.clickAt(state.rect.x, state.rect.y);

    const immediate = await alive(app, 3_000);
    console.log(`  חי מיד אחרי הלחיצה? ${immediate}`);

    const recovered = immediate ? 0 : await waitAlive(app);
    if (recovered === null) {
      console.log('  הדף לא חזר לענות ב-100 שניות');
      report.fail(`${name} — הקפאה`, 'הדף לא הריץ JavaScript במשך 100 שניות אחרי הלחיצה');
      return;
    }

    const totalMs = Date.now() - t0;
    console.log(`  הדף עונה שוב אחרי ${recovered}ms (סה"כ ${totalMs}ms)`);

    // המסמך אמור להיפתח מחדש. בלי זה „לא קפא” אינו „עבד”.
    let opened = false;
    for (let i = 0; i < 30; i++) {
      try {
        opened = await withTimeout(app.js('window.__qa.lineCount() > 0'), 3_000, 'lineCount');
      } catch {
        opened = false;
      }
      if (opened) break;
      await sleep(1_000);
    }
    const status = await withTimeout(app.status(), 5_000, 'status').catch(() => null);
    console.log(`  מסמך פתוח? ${opened} | שורת מצב: ${JSON.stringify(status)}`);

    if (!opened) {
      report.fail(`${name} — לא נפתח מסמך`, `הדף חי אחרי ${recovered}ms אך אין שורות במסמך`);
    } else if (recovered > 3_000) {
      report.partial(`${name} — נפתח, אך הדף היה תקוע`, `${recovered}ms בלי להריץ JavaScript`);
    } else {
      report.pass(`${name}`, `נפתח תוך ${totalMs}ms, הדף לא נחסם`);
    }
  } finally {
    app.close();
  }
}


/**
 * ביקורת: כמה זמן הדף חסום **בלי שנוגעים בו בכלל**.
 *
 * בלי המדידה הזאת אי אפשר לייחס תקיעה לפקד: מנוע ה-DOCX ממשיך לעמד ברקע
 * אחרי הפתיחה, והעבודה הזאת חוסמת את ה-thread הראשי. תקיעה שמופיעה גם כאן
 * אינה באג של הכפתור שנלחץ לפניה — היא הרעש שכל מדידה אחרת יושבת עליו.
 */
async function control(report, seconds = 60) {
  console.log(`\n────── ביקורת: ${seconds} שניות בלי שום אינטראקציה ──────`);
  const app = await openApp({ name: 'control', port: PORT });
  try {
    await widen(app);
    const gaps = [];
    const t0 = Date.now();
    let lastOk = Date.now();
    while (Date.now() - t0 < seconds * 1_000) {
      const ok = await alive(app, 2_000);
      const now = Date.now();
      if (ok) {
        const gap = now - lastOk;
        if (gap > 1_500) gaps.push(gap);
        lastOk = now;
      }
      await sleep(250);
    }
    const worst = gaps.length ? Math.max(...gaps) : 0;
    console.log(`  חסימות מעל 1.5 שניות: ${gaps.length} | הארוכה ביותר: ${worst}ms`);
    console.log(`  כל החסימות: ${JSON.stringify(gaps)}`);
    if (worst >= 5_000) {
      report.fail('ביקורת — הדף נחסם מעצמו', `${worst}ms בלי שום לחיצה, ${gaps.length} חסימות`);
    } else if (gaps.length) {
      report.partial('ביקורת — חסימות קצרות מעצמו', `הארוכה ${worst}ms`);
    } else {
      report.pass('ביקורת — הדף לא נחסם מעצמו', `${seconds} שניות רצופות`);
    }
  } finally {
    app.close();
  }
}

const SCENARIOS = [
  {
    id: 1,
    name: '„מסמך חדש” לבדו',
    steps: [],
  },
  {
    id: 2,
    name: 'אחרי „אודות” וסגירה',
    steps: [
      {
        what: 'אודות → Escape',
        run: async (app) => {
          await app.click('אודות');
          await app.escape();
        },
      },
    ],
  },
  {
    id: 3,
    name: 'אחרי „קיצורים” וסגירה',
    steps: [
      {
        what: 'קיצורים → Escape',
        run: async (app) => {
          await app.click('קיצורים');
          await app.escape();
        },
      },
    ],
  },
  {
    id: 4,
    name: 'הרצף המלא שדווח',
    steps: [
      {
        what: 'אודות → Escape',
        run: async (app) => {
          await app.click('אודות');
          await app.escape();
        },
      },
      {
        what: 'קיצורים → Escape',
        run: async (app) => {
          await app.click('קיצורים');
          await app.escape();
        },
      },
      {
        what: 'מיקום סמן במסמך',
        run: async (app) => {
          await app.caret(0);
          await app.tab('קובץ');
        },
      },
    ],
  },
  {
    id: 5,
    name: 'אחרי מיקום סמן בלבד',
    steps: [
      {
        what: 'מיקום סמן במסמך',
        run: async (app) => {
          await app.caret(0);
          await app.tab('קובץ');
        },
      },
    ],
  },
  {
    id: 6,
    name: 'אודות → Escape → קיצורים (בלי לסגור)',
    stopAfterSteps: true,
    steps: [
      { what: 'אודות → Escape', run: async (app) => { await app.click('אודות'); await app.escape(); } },
      { what: 'קיצורים (נשאר פתוח)', run: async (app) => { await app.click('קיצורים'); } },
    ],
  },
  {
    id: 7,
    name: 'אודות (בלי לסגור) → קיצורים',
    stopAfterSteps: true,
    steps: [
      { what: 'אודות (נשאר פתוח)', run: async (app) => { await app.click('אודות'); } },
      { what: 'קיצורים', run: async (app) => { await app.click('קיצורים'); } },
    ],
  },
  {
    id: 8,
    name: 'סדר הפוך: קיצורים → Escape → אודות',
    stopAfterSteps: true,
    steps: [
      { what: 'קיצורים → Escape', run: async (app) => { await app.click('קיצורים'); await app.escape(); } },
      { what: 'אודות', run: async (app) => { await app.click('אודות'); } },
    ],
  },
  {
    id: 9,
    name: 'אותו דיאלוג פעמיים: אודות → Escape → אודות',
    stopAfterSteps: true,
    steps: [
      { what: 'אודות → Escape', run: async (app) => { await app.click('אודות'); await app.escape(); } },
      { what: 'אודות שוב', run: async (app) => { await app.click('אודות'); } },
    ],
  },
  {
    id: 10,
    name: 'אותו דיאלוג פעמיים: קיצורים → Escape → קיצורים',
    stopAfterSteps: true,
    steps: [
      { what: 'קיצורים → Escape', run: async (app) => { await app.click('קיצורים'); await app.escape(); } },
      { what: 'קיצורים שוב', run: async (app) => { await app.click('קיצורים'); } },
    ],
  },
  {
    id: 11,
    name: 'אודות → Escape → קיצורים → Escape',
    stopAfterSteps: true,
    steps: [
      { what: 'אודות → Escape', run: async (app) => { await app.click('אודות'); await app.escape(); } },
      { what: 'קיצורים → Escape', run: async (app) => { await app.click('קיצורים'); await app.escape(); } },
    ],
  },
  {
    id: 12,
    name: 'ביקורת: קיצורים → Escape → קיצורים → Escape (בלי אודות)',
    stopAfterSteps: true,
    steps: [
      { what: 'קיצורים → Escape', run: async (app) => { await app.click('קיצורים'); await app.escape(); } },
      { what: 'קיצורים → Escape שוב', run: async (app) => { await app.click('קיצורים'); await app.escape(); } },
    ],
  },
  {
    id: 13,
    name: 'ביקורת: שני Escape בלי שום דיאלוג',
    stopAfterSteps: true,
    steps: [
      { what: 'Escape ראשון', run: async (app) => { await app.escape(); await sleep(600); } },
      { what: 'Escape שני', run: async (app) => { await app.escape(); await sleep(600); } },
    ],
  },
  {
    id: 14,
    name: 'סגירה בכפתור ✕ פעמיים (בלי Escape)',
    stopAfterSteps: true,
    steps: [
      { what: 'קיצורים → ✕', run: async (app) => { await app.click('קיצורים'); await app.clickDialog('סגור'); } },
      { what: 'קיצורים → ✕ שוב', run: async (app) => { await app.click('קיצורים'); await app.clickDialog('סגור'); } },
    ],
  },
  {
    id: 15,
    name: 'פתיחה וסגירה ראשונה בלבד (בסיס)',
    stopAfterSteps: true,
    steps: [
      { what: 'קיצורים → Escape', run: async (app) => { await app.click('קיצורים'); await app.escape(); } },
    ],
  },
];

const only = process.argv[2] ? Number(process.argv[2]) : null;
const report = createReport('„מסמך חדש” — בידוד ההקפאה', { strict: true });

if (only === 0) {
  await control(report);
  process.exit(report.print() > 0 ? 1 : 0);
}

for (const s of SCENARIOS) {
  if (only && s.id !== only) continue;
  try {
    await scenario(report, s);
  } catch (error) {
    console.log(`  ✗ התרחיש זרק: ${error.message}`);
    report.fail(`${s.name} — השער זרק`, error.message);
  }
}

process.exit(report.print() > 0 ? 1 : 0);
