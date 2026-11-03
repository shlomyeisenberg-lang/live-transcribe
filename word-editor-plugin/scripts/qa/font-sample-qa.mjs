/**
 * שער QA לפס הדגימה של בורר הגופן — התצוגה החיה **שכן נשלחת**.
 *
 * הפס מציג את הטקסט שהמשתמש סימן, בגופן שהסימון עומד עליו ובגודל שבמסמך.
 * שני חצאים נמדדים כאן, ושניהם נדרשים:
 *
 * 1. **שהוא מראה את הדבר הנכון** — הטקסט המסומן, בגופן שמרחפים עליו, ובגודל
 *    שנגזר מגודל הבחירה ולא בגודל קבוע של הרשימה.
 * 2. **שהוא אינו נוגע במסמך** — וזה החצי החשוב. כל ההיסטוריה של הפיצ'ר הזה
 *    היא מסלולים שנשללו מפני שריחוף לבד כתב למסמך: דרגת undo, „לא נשמר”,
 *    autosave, ומחיקת ה-redo של המשתמש (`docs/engine-gaps.md`, „תצוגה חיה של
 *    גופן”). לכן השער מודד `undoDepth`, `redoDepth`, `isDirty` ו-
 *    `word/document.xml` לפני מעבר על הרשימה ואחריו, ודורש שאף אחד מהם לא זז.
 *
 * הבדיקות ב-`tests/` מוכיחות שהפקד **מצייר** את מה שהועבר לו; רק כאן נמדד
 * שהגודל האמיתי של המסמך אכן זורם אליו, ושהריחוף באמת אינו כותב.
 *
 * יציאה 9633.
 */
import { openApp, createReport } from './harness.mjs';

const report = createReport('פס הדגימה של בורר הגופן', { strict: true });
const app = await openApp({ name: 'font-sample', port: Number(process.env.QA_PORT ?? 9633) });

const note = (...p) =>
  console.log(p.map((x) => (typeof x === 'string' ? x : JSON.stringify(x))).join(' '));

const T = (p, label, ms = 40_000) =>
  Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error(`תקיעה ב-${label}`)), ms))]);

/** מריץ גוף async בדף. הגוף מחזיר JSON. */
const P = (body, label = 'eval') =>
  T(
    app.js(
      `(async () => { try { const doc = window.__qa.doc(); ${body} } catch (e) { return JSON.stringify({ error: String((e && e.message) || e) }); } })()`,
    ),
    label,
  ).then((raw) => (raw == null ? null : JSON.parse(raw)));

const docXml = async () => ((await T(app.docx(), 'docx')) || {})['word/document.xml'] || '';

/**
 * תמונת מצב של כל מה שריחוף אסור לו לשנות.
 *
 * `history.get()` הוא **Promise** — בלי `await` הוא נראה `{}` וכל שער שנשען
 * עליו „עובר” בלי למדוד דבר. זו מלכודת מתועדת, וזו הסיבה ל-`await` כאן.
 */
const snapshot = async () => {
  const engine = await P(`
    const h = await doc.history.get();
    const sessions = window.__otzariaEditors;
    const first = sessions ? Array.from(sessions.values())[0] : null;
    return JSON.stringify({
      undoDepth: h.undoDepth,
      redoDepth: h.redoDepth,
      isDirty: first ? first.save.snapshot.isDirty : null,
    });
  `, 'snapshot');
  return { ...engine, xml: await docXml() };
};

/** מה שפס הדגימה מצייר בפועל, מהסגנון המחושב — לא מה שהועבר כ-prop. */
const sampleBar = () =>
  app
    .js(`(function () {
      var el = document.querySelector('.ribbon-combo-sample');
      if (!el) return 'null';
      var cs = getComputedStyle(el);
      return JSON.stringify({
        text: (el.textContent || '').trim(),
        fontSize: cs.fontSize,
        fontFamily: cs.fontFamily.split(',')[0].replace(/"/g, ''),
        height: Math.round(el.getBoundingClientRect().height),
        clipped: el.scrollHeight > el.clientHeight + 1,
      });
    })()`)
    .then((raw) => (raw === 'null' ? null : JSON.parse(raw)));

/**
 * ממתינה לבחירה שהתיישבה, ולא קוראת פעם אחת.
 *
 * זו לא זהירות אלא רגרסיה שנמדדה: קריאה בודדת מיד אחרי `extendSelection`
 * החזירה `{status:'stale', empty:true}` בשלוש הרצות מתוך ארבע, והשער נפל
 * בהכנת הפיקסטורה במקום למדוד משהו. המנוע פותר את הבחירה מחדש אחרי הקלדה,
 * וזמן הפתרון תלוי בעומס המכונה — כלומר שער שקורא פעם אחת מודד את העומס,
 * לא את המוצר.
 */
async function waitForSelection(ms = 15_000) {
  const until = Date.now() + ms;
  let last = null;
  for (;;) {
    last = await app.selection();
    if (last && last.empty === false && last.status === 'ready') return last;
    if (Date.now() > until) return last;
    await app.sleep(400);
  }
}

/** האפשרות שהסימון עומד עליה — מה שהפס אמור להדגים. */
const highlighted = () =>
  app.js(`(function () {
    var el = document.querySelector('.ribbon-combo-list .ribbon-combo-option.active');
    return el ? el.getAttribute('data-value') : null;
  })()`);

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1600, height: 1000, deviceScaleFactor: 1, mobile: false,
  });
  await app.sleep(500);
  await app.tab('בית');

  /* ===== פיקסטורה: טקסט, בחירה, וגודל ידוע ===== */
  const TEXT = 'shalom olam';
  await app.caret(0);
  await app.type(TEXT, 25);
  await app.sleep(1_200);

  await app.press('Home', 'Home', 36);
  await app.sleep(150);
  await app.extendSelection(TEXT.length);
  await app.sleep(500);

  const sel = await waitForSelection();
  note('הבחירה:', sel);
  if (!sel || sel.empty !== false) {
    report.fail('הכנת הבדיקה — יש בחירה', `הבחירה דיווחה ${JSON.stringify(sel)}`);
    throw new Error('בלי בחירה אין פס דגימה למדוד');
  }

  /*
   * גודל ידוע, כדי שהמדידה תהיה מול מספר ולא מול „מה שיצא”. 20pt → 26.67px,
   * כלומר מעל התקרה (24px) — וזה בכוונה: התקרה היא חלק מהחוזה.
   *
   * הקלדה ו-Enter, ולא `app.selectValue`: העזר ההוא שולח `mousedown`, ו-
   * `RibbonCombo` מאזין ל-`pointerdown` בלבד (ראו את התבנית שלו, ואת הבדיקה
   * „`mousedown` לבדו אינו בוחר”). כלומר `selectValue` היה יוצא „ok” בלי
   * לבחור דבר, והשער היה מודד את גודל ברירת המחדל ומדווח הצלחה.
   */
  await app.click('גודל גופן');
  await app.sleep(400);
  await app.type('20', 40);
  await app.sleep(200);
  await app.press('Enter', 'Enter', 13);
  await app.sleep(1_500);
  const reported = (await app.cmd('font-size')).value;
  note('גודל שהמנוע מדווח:', reported);
  if (Number(reported) !== 20) {
    report.fail('הכנת הבדיקה — הגודל הוחל', `המנוע מדווח ${JSON.stringify(reported)} ולא 20`);
    throw new Error('בלי גודל ידוע אין מול מה למדוד את הפס');
  }

  const before = await snapshot();
  note('לפני:', { undoDepth: before.undoDepth, redoDepth: before.redoDepth, isDirty: before.isDirty });

  /* ===== א: הפס מציג את הטקסט המסומן, בגודל שנגזר מהמסמך ===== */
  await app.click('גופן');
  await app.sleep(700);

  const bar = await sampleBar();
  note('הפס:', bar, '| מסומן:', await highlighted());

  if (!bar) {
    report.fail('הפס מוצג כשיש בחירה ורשימה פתוחה', 'לא נמצא `.ribbon-combo-sample` ב-DOM');
  } else {
    bar.text === TEXT
      ? report.pass('הפס מציג את הטקסט המסומן', `„${bar.text}”`)
      : report.fail('הפס מציג את הטקסט המסומן', `הוצג „${bar.text}” במקום „${TEXT}”`);

    // 20pt × 4/3 = 26.67px, חסום ל-24px. הגודל הקבוע שהיה קודם הוא 14px,
    // ולכן 14px כאן פירושו שהגודל האמיתי אינו זורם לפס בכלל.
    const px = Number.parseFloat(bar.fontSize);
    if (px === 24) {
      report.pass('הפס מצייר בגודל שנגזר מהמסמך', `20pt → ${bar.fontSize} (חסום ל-24px)`);
    } else if (px === 14) {
      report.fail(
        'הפס מצייר בגודל שנגזר מהמסמך',
        'הפס ב-14px — הגודל הקבוע של הרשימה. הגודל שבמסמך אינו מגיע אליו',
      );
    } else {
      report.fail('הפס מצייר בגודל שנגזר מהמסמך', `נמדד ${bar.fontSize}, ולא 24px`);
    }

    bar.clipped
      ? report.fail('הפס אינו חותך את הדגימה', `scrollHeight > clientHeight בגובה ${bar.height}px`)
      : report.pass('הפס אינו חותך את הדגימה', `גובה ${bar.height}px, טקסט ב-${bar.fontSize}`);
  }

  /* ===== ב: הגופן עוקב אחרי הסימון ===== */
  const first = await highlighted();
  await app.press('ArrowDown', 'ArrowDown', 40);
  await app.sleep(400);
  const moved = await highlighted();
  const afterArrow = await sampleBar();
  note('אחרי חץ — מסומן:', moved, '| הפס:', afterArrow);

  if (moved && moved !== first && afterArrow) {
    afterArrow.fontFamily === moved
      ? report.pass('הפס מצייר בגופן שהסימון עומד עליו', `${first} → ${moved}`)
      : report.fail(
          'הפס מצייר בגופן שהסימון עומד עליו',
          `הסימון על „${moved}” והפס מצייר ב-„${afterArrow.fontFamily}”`,
        );
  } else {
    report.fail('החץ מזיז את הסימון ברשימה', `לפני „${first}”, אחרי „${moved}”`);
  }

  /* ===== ג: מעבר על הרשימה אינו נוגע במסמך — החצי החשוב ===== */
  for (let i = 0; i < 12; i += 1) {
    await app.press('ArrowDown', 'ArrowDown', 40);
    await app.sleep(120);
  }
  await app.sleep(1_000);
  const during = await snapshot();
  note('בזמן מעבר על הרשימה:', {
    undoDepth: during.undoDepth, redoDepth: during.redoDepth, isDirty: during.isDirty,
  });

  // Escape ולא בחירה: „יצאתי בלי לבחור” הוא המסלול שבו מסלול הצביעה היה
  // מחזיר את הגופן — כלומר מוטציה שנייה — ולכן זה המסלול שחייב להיות נקי.
  await app.escape();
  await app.sleep(1_500);
  const after = await snapshot();
  note('אחרי Escape:', { undoDepth: after.undoDepth, redoDepth: after.redoDepth, isDirty: after.isDirty });

  const still = (key) => before[key] === during[key] && before[key] === after[key];

  still('undoDepth')
    ? report.pass('מעבר על רשימת הגופנים אינו קונה דרגת undo', `undoDepth נשאר ${before.undoDepth}`)
    : report.fail(
        'מעבר על רשימת הגופנים אינו קונה דרגת undo',
        `${before.undoDepth} → ${during.undoDepth} → ${after.undoDepth}`,
      );

  still('redoDepth')
    ? report.pass('הריחוף אינו נוגע במחסנית ה-redo', `redoDepth נשאר ${before.redoDepth}`)
    : report.fail(
        'הריחוף אינו נוגע במחסנית ה-redo',
        `${before.redoDepth} → ${during.redoDepth} → ${after.redoDepth} — עבודה שהמתינה ל-redo נמחקה`,
      );

  still('isDirty')
    ? report.pass('הריחוף אינו מסמן „לא נשמר”', `isDirty נשאר ${before.isDirty}`)
    : report.fail(
        'הריחוף אינו מסמן „לא נשמר”',
        `${before.isDirty} → ${during.isDirty} → ${after.isDirty}`,
      );

  before.xml === after.xml
    ? report.pass('המסמך זהה אחרי מעבר על הרשימה', 'word/document.xml ללא שינוי')
    : report.fail(
        'המסמך זהה אחרי מעבר על הרשימה',
        `word/document.xml השתנה (${before.xml.length} → ${after.xml.length} תווים)`,
      );

  /* ===== ד: redo שממתין שורד את הריחוף — הבאג שסגר את מסלול הצביעה ===== */
  /*
   * זו השורה הספציפית לרגרסיה. מסלול הצביעה שנשלל (`format.apply` בריחוף,
   * `history.undo` ביציאה) נמדד כמוחק את מחסנית ה-redo של המשתמש ומחליף אותה
   * בגופן שאיש לא בחר — גם כשהמשתמש **ביטל** את התצוגה. כאן נבנה בדיוק אותו
   * מצב: עבודה שבוטלה וממתינה ל-redo, ואז מעבר על רשימת הגופנים.
   */
  await app.caret(0);
  await app.press('End', 'End', 35);
  await app.sleep(200);
  await app.type('Z', 40);
  await app.sleep(1_200);
  await app.press('z', 'KeyZ', 90, 2); // Ctrl+Z — עכשיו יש redo אמיתי שממתין
  await app.sleep(1_500);

  const pending = await snapshot();
  note('אחרי הקלדה+ביטול:', { undoDepth: pending.undoDepth, redoDepth: pending.redoDepth });

  if (pending.redoDepth < 1) {
    report.fail(
      'הכנת הבדיקה — יש redo שממתין',
      `redoDepth הוא ${pending.redoDepth}; בלי redo ממתין אין מה למדוד`,
    );
  } else {
    // בחירה מחדש, כדי שיהיה פס דגימה בכלל.
    await app.press('Home', 'Home', 36);
    await app.sleep(200);
    await app.extendSelection(TEXT.length);
    await waitForSelection();

    await app.click('גופן');
    await app.sleep(600);
    for (let i = 0; i < 8; i += 1) {
      await app.press('ArrowDown', 'ArrowDown', 40);
      await app.sleep(120);
    }
    await app.sleep(600);
    await app.escape();
    await app.sleep(1_500);

    const kept = await snapshot();
    note('אחרי ריחוף על הרשימה:', { undoDepth: kept.undoDepth, redoDepth: kept.redoDepth });

    kept.redoDepth === pending.redoDepth
      ? report.pass(
          'redo שממתין שורד מעבר על רשימת הגופנים',
          `redoDepth נשאר ${kept.redoDepth} — העבודה שבוטלה עדיין ניתנת לשחזור`,
        )
      : report.fail(
          'redo שממתין שורד מעבר על רשימת הגופנים',
          `redoDepth ${pending.redoDepth} → ${kept.redoDepth}: הריחוף מחק עבודה שהמתינה ל-redo`,
        );

    // וגם שהוא ה-redo **הנכון**: Ctrl+Y מחזיר את ה-Z שהוקלד, לא גופן.
    await app.press('y', 'KeyY', 89, 2);
    await app.sleep(1_500);
    const redone = (await docXml()).includes('Z');
    note('אחרי Ctrl+Y — ה-Z חזר למסמך:', redone);
    redone
      ? report.pass('ה-redo מחזיר את מה שהמשתמש הקליד', 'ה-Z חזר ל-word/document.xml')
      : report.fail(
          'ה-redo מחזיר את מה שהמשתמש הקליד',
          'ה-Z לא חזר — ה-redo מצביע על משהו אחר',
        );
  }

  const log = (await app.log()).filter((l) => !/addContextMenuItem|listInstalled/.test(l));
  note('לוג הדף:', log);
  log.length === 0
    ? report.pass('אין שגיאות בלוג', 'הריחוף לא הותיר אזהרה')
    : report.fail('אין שגיאות בלוג', JSON.stringify(log));
} catch (error) {
  console.error('השער נפל:', error);
  report.fail('השער השלים את הריצה', error.message);
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
