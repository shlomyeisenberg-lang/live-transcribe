/**
 * שני טורים במקטע `w:bidi` — איפה הם מצוירים, מה נבחר בגרירה, ומה עושה המקלדת.
 *
 * ## מה השתנה כאן, ולמה הגשש נשאר
 *
 * עד superdoc 2.11.0 המנוע מילא את הטורים שמאל→ימין גם תחת `w:bidi`, בניגוד
 * ל-ECMA-376 §17.6.1: שורה 01 נחתה בטור השמאלי, והמשתמש קיבל הודעה בפס המצב
 * (`rtlColumnNote`) שמסבירה את זה. הפער נסגר במנוע — `SD-4764`, נכלל
 * ב-2.12.0 — וההודעה ירדה. הגשש לא ירד איתה: הוא הפך מ**תיאור של פער**
 * ל**שער שמונע את חזרתו**, ובלי שער כזה רגרסיה בציור הטורים הייתה מגיעה
 * למשתמש בשקט מוחלט (ה-docx המיוצא תקין בשני המקרים — ראו שורה 2).
 *
 * ## מה נמדד
 *
 *   1. „עמודות ← שתיים” מהתפריט האמיתי מצליחה ואינה משאירה הודעה בפס המצב.
 *      הודעה כאן = ההערה שהוסרה חזרה, או שנוספה אחרת בלי שאיש שם לב.
 *   2. הייצוא: `w:bidi` ושני טורים ב-`sectPr`.
 *   3. **בצד ימין נוחת הטור הראשון** — הפער עצמו, בכיוונו החדש.
 *   4. גרירה מהטור הראשון אל השני: הטווח שנבחר הוא בדיוק המבוקש, והוא רצף.
 *      בעברית זו גם מחוות הקריאה (ימין ואז שמאל) — הן התלכדו עם התיקון.
 *   5. גרירה **בתוך הטור הראשון** אל מתחת לשורה האחרונה שבו — האם הראש מהודק
 *      לסוף הטור או קופץ לטור השכן.
 *   6. Shift+חץ מטה בגבול שבין הטורים — חוצה כראוי מ-superdoc 2.15.0-next.15.
 *   7. **בקרת LTR** לשתי השורות שמעליה. זו השורה שמכריעה מה שייך ל-`w:bidi`
 *      ומה שייך לטורים בכלל: אותה מחווה בדיוק על מקטע לועזי. נמדד שהמקלדת
 *      חוצה את הגבול **זהה** בשני הכיוונים — כלומר זו התנהגות טורים כללית,
 *      לא תלוית כיוון. ראו `docs/engine-gaps.md`.
 *
 * ## שלושה כללי מדידה שנלמדו כאן בדם, ואסור לוותר עליהם
 *
 * 1. **לחיצה שממקמת סמן חייבת השהיה בין `mousePressed` ל-`mouseReleased`.**
 *    בלעדיה המנוע אינו מאפס את העוגן, הבחירה הקודמת פשוט מורחבת, והשורה
 *    מודדת את הצעד הקודם ולא את עצמה. זה קרה כאן בפועל: שורת המקלדת „הוכיחה”
 *    פער שהיא כלל לא מדדה. לכן כל מיקום סמן עובר ב-`placeCaret`, שגם **מאמת**
 *    שהבחירה התאפסה לפני שממשיכים.
 * 2. **מרכז הפרגמנט הוא אמצע הטקסט, לא תחילתו.** לחיצה על `at(n)` מציבה את
 *    העוגן בתוך שורה n, ולכן המספר הראשון שנספר הוא n+1. כל קריטריון כאן
 *    לוקח את זה בחשבון במפורש.
 * 3. **ספירת מלבנים אינה קריטריון.** היא נמדדת ומדווחת כראיה, אבל מה שמכריע
 *    הוא הטווח שנבחר: „20 מלבנים ירדו ל-14” נראה כמו בחירה שנמחקה גם כשהטווח
 *    שנבחר מדויק לחלוטין. הקריטריון הישן כאן היה בדיוק זה, והוא סימן „חלקי”
 *    על גרירה תקינה אחרי שהמנוע כבר תוקן.
 *
 * הרצה:  CHROME=<נתיב> node scripts/qa/column-selection-probe.mjs
 */
import { openApp, createReport, sleep } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9385);
const report = createReport('שני טורים במקטע עברי: ציור, בחירה ומקלדת');

/** כמה פסקאות. די כדי לגלוש לטור שני ולהשאיר בכל טור שורות לזהות. */
const LINES = 40;

/**
 * שוליים עליונים ותחתונים באינץ'. גדולים בכוונה: שטח כתיבה נמוך ממלא שני
 * טורים ב-40 שורות, ובלעדיו היה צריך מאות — והמדידה הייתה לוקחת דקות.
 */
const TALL_MARGIN = 3.3;

/** שם הפקד ברצועה. `openMenu` מחפש לפי התווית, לא לפי ה-tooltip. */
const COLUMNS_BUTTON = 'עמודות';

const app = await openApp({ name: 'column-selection', port: PORT });
await app.cdp.send('Emulation.setDeviceMetricsOverride', {
  width: 1400,
  height: 1500,
  deviceScaleFactor: 1,
  mobile: false,
});
await sleep(1_500);

const mouse = (type, x, y, extra = {}) =>
  app.cdp.send('Input.dispatchMouseEvent', {
    type,
    x,
    y,
    button: 'left',
    buttons: type === 'mousePressed' ? 1 : 0,
    clickCount: 1,
    ...extra,
  });

/**
 * כמה מלבני בחירה מצוירים, ובאיזה צד של הדף.
 *
 * `ui.selection.getRects()` הוא ה-API הציבורי שמחזיר בדיוק את מה שמודגש על
 * המסך — ולכן זו המדידה שמייצגת „מה המשתמש רואה”, ולא ספירה של תווים. הוא
 * מחזיר קואורדינטות **חלון** כשלא מועבר `relativeTo`, ולכן ההשוואה ל-
 * `pageMiddle` שנמדד ב-`getBoundingClientRect` היא באותו מרחב.
 */
const rectsBySide = (middle) =>
  app.js(`(() => {
    const ui = window.__otzariaEditor.ui;
    let rects = [], error = null;
    try { rects = ui.selection.getRects() || []; } catch (e) { error = e.message; }
    let left = 0, right = 0;
    rects.forEach((r) => { if (r.left < ${middle}) left++; else right++; });
    // המספרים תמיד קיימים, גם בחריגה. אחרת Math.max על undefined יוצא NaN,
    // ההשוואה יוצאת false, וכשל של ה-API הסגור נרשם כאותה שורת דוח בדיוק כמו
    // הפער שהגשש כולו קיים בשבילו.
    return JSON.stringify({ n: rects.length, left, right, error });
  })()`).then(JSON.parse);

/**
 * מה נבחר, לפי מספרי השורות שבטקסט.
 *
 * `contiguous` נמדד ולא מונח: „25 שורות” אינו „25 שורות **רצופות**”, ובלי
 * הבדיקה הזאת הכותרת בדוח הייתה אומרת יותר ממה שהקוד יודע.
 */
const selectedLines = () =>
  app.js(`(async () => {
    const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    const current = await doc.selection.current({ includeText: true });
    const text = (current && current.text) || '';
    const nums = [...text.matchAll(/שורה (\\d\\d)/g)].map((m) => Number(m[1]));
    if (!nums.length) return JSON.stringify({ label: '(ריק)', count: 0, contiguous: true });
    const contiguous = nums.every((n, i) => i === 0 || n === nums[i - 1] + 1);
    return JSON.stringify({
      label: nums[0] + '..' + nums[nums.length - 1] + ' (' + nums.length + (contiguous ? '' : ', לא רצוף') + ')',
      first: nums[0], last: nums[nums.length - 1], count: nums.length, contiguous,
    });
  })()`).then(JSON.parse);

/**
 * מיקום סמן, **עם אימות**.
 *
 * `app.clickAt` של המסגרת ולא `mouse()` המקומי: הוא מכניס 40ms בין ה-press
 * ל-release, ובלעדיהם המנוע אינו מאפס את העוגן. מוחזר `true` רק כשהבחירה
 * באמת התאפסה — סמן אינו בוחר טקסט, ולכן `count === 0` הוא התנאי.
 *
 * מה שהוא **אינו** מאמת: **איפה** נחת הסמן. `count === 0` נכון לכל מיקום,
 * ולכן שורה שמסקנתה תלויה במיקום — כמו שורה 7, שכל מובנה הוא היחס לעוגן —
 * חייבת לאמת את המיקום בעצמה בהקשה בודדת לפני שהיא מודדת רצף הקשות.
 */
async function placeCaret(point) {
  await app.clickAt(point.x, point.y);
  await sleep(400);
  const after = await selectedLines();
  return after.count === 0;
}

/**
 * גרירה לאורך מסלול, עם דגימה אחרי כל צעד ודגימה נוספת **אחרי השחרור**.
 *
 * המסלול ולא קו ישר: „הכול של הטור הראשון התבטל” הוא אירוע **באמצע**
 * הגרירה, ומדידה של מצב הסיום בלבד מפספסת אותו. והדגימה שאחרי השחרור נפרדת,
 * מפני שהכרעה על סמך הדגימה שלפניו הייתה משווה מצב חי למצב שהתייצב.
 */
async function dragAlong(path, sample) {
  const [first] = path;
  await mouse('mouseMoved', first.x, first.y, { button: 'none', buttons: 0 });
  await mouse('mousePressed', first.x, first.y);
  await sleep(100);
  const trail = [];
  for (const point of path.slice(1)) {
    await mouse('mouseMoved', point.x, point.y, { buttons: 1 });
    await sleep(90);
    trail.push({ ...point, ...(await sample()) });
  }
  const last = path[path.length - 1];
  await mouse('mouseReleased', last.x, last.y);
  await sleep(600);
  return { trail, final: await sample(), threw: trail.find((p) => p.error)?.error ?? null };
}

try {
  /* -------------------- בניית המסמך -------------------- */

  // שוליים והטקסט דרך ה-API: זה תנאי המדידה, לא מה שנמדד.
  await app.js(`(async () => {
    const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    const first = (await doc.sections.list()).items[0];
    await doc.sections.setPageMargins({ target: first.address, top: ${TALL_MARGIN}, bottom: ${TALL_MARGIN}, left: 1, right: 1 });
    const paras = [];
    for (let i = 1; i <= ${LINES}; i++) paras.push('שורה ' + String(i).padStart(2, '0') + ' פסקה עברית');
    await doc.insert({ value: paras.join('\\n\\n'), type: 'markdown' });
  })()`);
  await sleep(2_000);

  /* -------------------- 1: ההודעה בשורת המצב -------------------- */

  // העמודות **מהתפריט האמיתי**, ולא מה-API: מה שנמדד כאן הוא המסלול שהמשתמש
  // עובר, כולל המדווח ושורת המצב. `reset` קודם, כדי ששורת מצב ישנה לא תיזקף
  // לזכות הפקודה הזאת.
  await app.reset();
  await app.tab('פריסה');
  const menu = await app.openMenu(COLUMNS_BUTTON);
  if (!menu) {
    report.stuck('„עמודות” לא נפתח', 'הפקד לא נמצא ברצועה');
  } else if (!(await app.clickMenu('שתיים', { after: 2_500 }))) {
    report.stuck('„שתיים” לא נלחץ', JSON.stringify(menu));
  } else {
    const status = await app.status();
    if (!status?.error && !status?.text) {
      report.pass('„עמודות ← שתיים” מצליחה בשקט', 'שורת המצב נשארה ריקה');
    } else {
      report.fail(
        'הפעולה הותירה טקסט בשורת המצב',
        'ההערה על סדר הטורים ירדה עם superdoc 2.12.0, ואין פקודה אחרת שאמורה ' +
          `לדבר כאן: ${JSON.stringify(status)}`,
      );
    }
  }
  await sleep(2_500);

  /* -------------------- 2: מה יצא ל-OOXML -------------------- */

  // ההוכחה שהמקטע באמת RTL אינה `sectionDirection` לבדו אלא ה-XML שיוצא.
  const files = await app.docx();
  const sectPr = (files['word/document.xml'] || '').match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)?.pop() ?? '';
  const hasBidi = /<w:bidi\b/.test(sectPr);
  const hasTwoCols = /<w:cols[^>]*w:num="2"/.test(sectPr);
  if (hasBidi && hasTwoCols) {
    report.pass('הייצוא נכון — `w:bidi` ושני טורים ב-sectPr', sectPr.match(/<w:cols[^>]*>/)?.[0] ?? '');
  } else {
    report.fail('sectPr אינו מה שהמדידה מניחה', `bidi=${hasBidi} cols2=${hasTwoCols} | ${sectPr}`);
  }

  /* -------------------- 3: איפה נחת כל טור -------------------- */

  /** נקרא פעמיים: פעם על המקטע העברי, ופעם על בקרת ה-LTR שבסוף. */
  const readGeometry = () =>
    app.js(`(() => {
    const page = document.querySelector('.superdoc-page').getBoundingClientRect();
    const rows = [];
    document.querySelectorAll('.superdoc-fragment').forEach((el) => {
      const match = /^שורה (\\d\\d)/.exec((el.textContent || '').trim());
      if (!match) return;
      const r = el.getBoundingClientRect();
      rows.push({ line: Number(match[1]), x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2) });
    });
    return JSON.stringify({ pageMiddle: Math.round(page.x + page.width / 2), rows });
  })()`).then(JSON.parse);

  const { pageMiddle, rows } = await readGeometry();
  const at = (line) => rows.find((r) => r.line === line);
  const side = (line) => (at(line).x < pageMiddle ? 'שמאלי' : 'ימני');
  const sample = () => rectsBySide(pageMiddle);

  /**
   * בלי **שני** טורים משני צדי אמצע הדף אין מה למדוד, וכל שורה שתירשם אחרי
   * זה תהיה שקר. הבדיקה היא על נוכחות בשני הצדדים ולא רק על „שני ערכי x
   * שונים”: פריסה שנשברה אחרת יכולה לתת שני ערכים באותו צד.
   */
  const leftRows = rows.filter((r) => r.x < pageMiddle);
  const rightRows = rows.filter((r) => r.x >= pageMiddle);
  if (rows.length < LINES || leftRows.length === 0 || rightRows.length === 0) {
    report.stuck(
      'המסמך לא הגיע לשני טורים — אין מה למדוד',
      `${rows.length}/${LINES} שורות, ${leftRows.length} משמאל ו-${rightRows.length} מימין`,
    );
    report.print();
    app.close();
    process.exit(0);
  }

  if (side(1) === 'ימני') {
    report.pass('הטור הראשון נוחת מימין, כמו בוורד', `שורה 01 ב-x=${at(1).x}, אמצע הדף ${pageMiddle}`);
  } else {
    report.fail(
      'סדר מילוי הטורים מתעלם מ-`w:bidi`',
      `שורה 01 נוחתת בטור ה${side(1)} (x=${at(1).x}, אמצע הדף ${pageMiddle}) — בוורד היא מתחילה מימין`,
    );
  }

  // הגבול בין הטורים בסדר המסמך: השורה האחרונה שעדיין בטור הראשון.
  // לפי הצד שנמדד, ולא לפי „ימין הוא הראשון”: אותו קוד משרת גם את בקרת ה-LTR
  // בסוף הקובץ, ושורה שמניחה כיוון אינה יכולה להכריע מה תלוי בכיוון.
  const firstColumnSide = side(1);
  const inFirst = rows.filter((r) => side(r.line) === firstColumnSide);
  const inSecond = rows.filter((r) => side(r.line) !== firstColumnSide);
  const lastOfFirst = inFirst[inFirst.length - 1].line;
  const firstOfSecond = lastOfFirst + 1;

  /* -------------------- 4: גרירה מהטור הראשון אל השני -------------------- */

  {
    // מחוות הקריאה: התחלה בטור שנקרא ראשון, ירידה בתוכו, ואז מעבר לשני. עד
    // 2.11.0 היא הייתה גם המחווה ההפוכה לסדר המסמך (הטור שנקרא ראשון היה
    // השני בסדר), ולכן היא הפכה את הבחירה סביב העוגן. עם התיקון השתיים
    // התלכדו, וזו בדיוק הסיבה שהשורה הזאת מודדת **טווח** ולא כיוון.
    const from = inFirst[Math.floor(inFirst.length * 0.3)];
    const down = inFirst[Math.floor(inFirst.length * 0.8)];
    const to = inSecond[Math.floor(inSecond.length * 0.75)];

    const path = [from];
    for (let y = from.y; y <= down.y; y += 30) path.push({ x: from.x, y });
    // הצעד האופקי חוצה את המרזב, ובכיוון שנמדד ולא בכיוון שמונח: הגבול הוא
    // המקום שבו הבחירה נשברה, ומסלול שמדלג עליו אינו מודד אותו.
    const stepX = to.x > down.x ? 60 : -60;
    for (let x = down.x; stepX > 0 ? x <= to.x : x >= to.x; x += stepX) path.push({ x, y: down.y });
    path.push(to);

    const { trail, final, threw } = await dragAlong(path, sample);
    const peak = Math.max(...trail.map((p) => p.right));
    const got = await selectedLines();
    // הראש בשורה `to.line` והזנב בשורה שאחרי העוגן (כלל המדידה 2).
    const wanted = `${from.line + 1}..${to.line}`;
    const rects = `מלבנים בסיום ${final.n} (${final.left} משמאל, ${final.right} מימין; שיא מימין ${peak})`;

    if (threw) {
      report.stuck('גרירה מהטור הראשון אל השני — `getRects` זרק', threw);
    } else if (got.contiguous && got.first === from.line + 1 && got.last === to.line) {
      report.pass('גרירה מהטור הראשון אל השני בוחרת בדיוק את הטווח', `${got.label} = ${wanted}; ${rects}`);
    } else if (!got.contiguous) {
      report.fail('גרירה חוצת-גבול מפרקת את הבחירה', `${got.label}, רצוי ${wanted}; ${rects}`);
    } else {
      report.fail(
        'גרירה חוצת-גבול בוחרת טווח אחר מהמבוקש',
        `${got.label} במקום ${wanted}; ${rects}`,
      );
    }
  }

  /* ------------- 5: גרירה בתוך הטור הראשון, אל מתחת לתחתיתו ------------- */

  {
    const bottomFirst = inFirst[inFirst.length - 1];
    const bottomSecond = inSecond[inSecond.length - 1];
    // נקודה ברוחב הטור הראשון, מתחת לשורה האחרונה שלו אך מעל תחתית השני.
    const belowY = Math.round((bottomFirst.y + bottomSecond.y) / 2);
    /** כל המלבנים בצד של הטור הראשון — הניסוח שאינו מניח איזה צד זה. */
    const onlyFirstSide = (final) =>
      firstColumnSide === 'ימני' ? final.left === 0 && final.right > 0 : final.right === 0 && final.left > 0;
    if (belowY <= bottomFirst.y + 10) {
      report.skip('גרירה אל מתחת לתחתית הטור הראשון', 'הטורים מסתיימים קרוב מדי זה לזה — אין רצועה למדוד');
    } else {
      // **העוגן בטור הראשון עצמו** — אחרת השורה מודדת משהו אחר לגמרי.
      const anchor = inFirst[Math.floor(inFirst.length * 0.3)];
      if (!(await placeCaret(anchor))) {
        report.stuck('גרירה אל מתחת לתחתית הטור הראשון', 'הלחיצה לא מיקמה סמן נקי');
      } else {
        const path = [anchor];
        for (let y = anchor.y + 40; y <= belowY; y += 40) path.push({ x: anchor.x, y });
        path.push({ x: anchor.x, y: belowY });

        const { trail, final, threw } = await dragAlong(path, sample);
        const before = trail[trail.length - 2] ?? trail[0];
        const got = await selectedLines();
        if (threw) {
          report.stuck('גרירה אל מתחת לתחתית הטור הראשון — `getRects` זרק', threw);
        } else if (onlyFirstSide(final)) {
          report.pass(
            'הראש מהודק לסוף הטור הראשון כשהסמן יורד מתחתיו',
            `${final.n} מלבנים, כולם בצד ה${firstColumnSide}; ${got.label}`,
          );
        } else {
          report.fail(
            'סמן ברוחב הטור הראשון מתחת לשורה האחרונה שבו — הראש אינו מהודק לסוף הטור',
            `בגובה ${belowY}: ${final.left} מלבנים משמאל ו-${final.right} מימין ` +
              `(בצעד שלפניו ${before.left}/${before.right}) — הנבחר: ${got.label}`,
          );
        }
      }
    }
  }

  /* -------------------- 6: המקלדת בגבול שבין הטורים -------------------- */

  {
    // Shift+חץ מטה חוצה את הגבול כראוי, ונותן למשתמש דרך לעקוף את
    // הגרירה. `placeCaret` ולא לחיצה גולמית: בלי אימות שהעוגן התאפס, השורה
    // הזאת מודדת את הבחירה של הצעד הקודם ומדווחת עליה כפער של המקלדת.
    const anchorLine = lastOfFirst - 2;
    if (!(await placeCaret(at(anchorLine)))) {
      report.stuck('Shift+חץ מטה בגבול הטורים', 'הלחיצה לא מיקמה סמן נקי — אין ממה למדוד');
    } else {
      // אימות **מיקום** ולא רק „אין בחירה”: `placeCaret` מאמת שהבחירה
      // התאפסה, ולא באיזו שורה נחת הסמן. הקשה בודדת אומרת את מספר השורה
      // בפועל, ובלעדיה כל מה שיימדד ב-6 ההקשות תלוי בהנחה על המיקום — כולל
      // המסקנה על הגלישה, שכל מובנה הוא היחס לעוגן.
      await app.press('ArrowDown', 'ArrowDown', 40, 8);
      await sleep(200);
      const step = await selectedLines();
      if (!(step.count === 1 && step.first === anchorLine + 1)) {
        report.stuck(
          'Shift+חץ מטה בגבול הטורים',
          `הקשה בודדת מסמן שהוצב בשורה ${anchorLine} בחרה ${step.label} ולא את שורה ` +
            `${anchorLine + 1} לבדה — מיקום הסמן אינו מה שהמדידה מניחה, ואין ממה להסיק`,
        );
      } else {
        for (let i = 0; i < 5; i++) {
          await app.press('ArrowDown', 'ArrowDown', 40, 8);
          await sleep(200);
        }
        const got = await selectedLines();
        // תקין = מתחיל מיד אחרי העוגן (ראו כלל המדידה 2) **וגם** חוצה את הגבול.
        const startsRight = got.first === anchorLine + 1;
        const crossed = got.last >= firstOfSecond;
        // הראש מוקדם מהעוגן ⟹ הבחירה נמשכה **אחורה** בסדר המסמך: מתחתית
        // הטור הראשון היא גלשה לתחילת המסמך ולא לראש הטור השני.
        const wrapped = got.first < anchorLine;
        if (startsRight && crossed) {
          report.pass('Shift+חץ מטה חוצה את הגבול בין הטורים', got.label);
        } else if (wrapped) {
          report.fail(
            'Shift+חץ מטה בתחתית הטור הראשון גולש לתחילת המסמך',
            `העוגן בשורה ${anchorLine}, ואומת בהקשה בודדת (${step.label}); 6 הקשות: ${got.label} — ` +
              `אחרי השורה האחרונה בטור הראשון (${lastOfFirst}) הראש נמצא בשורה ${got.first}, ` +
              `מוקדם מהעוגן, במקום בשורה ${firstOfSecond} שהיא ראש הטור השני`,
          );
        } else {
          report.fail(
            'Shift+חץ מטה אינו חוצה את הגבול בין הטורים',
            `העוגן בשורה ${anchorLine}, ואומת בהקשה בודדת (${step.label}); 6 הקשות: ${got.label} — ` +
              `${startsRight ? 'מתחיל נכון' : `מתחיל ב-${got.first} ולא ב-${anchorLine + 1}`}, ` +
              `${crossed ? 'חצה' : `לא הגיע לשורה ${firstOfSecond}`}`,
          );
        }
      }
    }
  }

  /* -------------------- 7: בקרת LTR -------------------- */

  /*
   * אותו מסמך, אותה מחווה, מקטע לועזי. בלי השורה הזאת אי אפשר לדעת אם מה
   * שנמדד למעלה שייך ל-`w:bidi` או לטורים בכלל — וההבדל הוא שההודעה שהוסרה
   * דיברה על מקטע עברי דווקא. נמדד: המקלדת חוצה את הגבול **זהה** בשני הכיוונים.
   */
  {
    await app.js(`(async () => {
      const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
      const first = (await doc.sections.list()).items[0];
      await doc.sections.setSectionDirection({ target: first.address, direction: 'ltr' });
    })()`);
    await sleep(2_500);

    const ltr = await readGeometry();
    const atL = (line) => ltr.rows.find((r) => r.line === line);
    const sideL = (line) => (atL(line).x < ltr.pageMiddle ? 'שמאלי' : 'ימני');
    if (ltr.rows.length < LINES) {
      report.stuck('בקרת LTR', `${ltr.rows.length}/${LINES} שורות אחרי החלפת הכיוון — אין מה להשוות`);
    } else {
      const firstSideL = sideL(1);
      const lastOfFirstL = ltr.rows.filter((r) => sideL(r.line) === firstSideL).pop().line;
      if (firstSideL === 'שמאלי') {
        report.pass('בקרת LTR: הטור הראשון עובר לשמאל עם הכיוון', `שורה 01 ב-x=${atL(1).x}, אמצע ${ltr.pageMiddle}`);
      } else {
        report.fail(
          'בקרת LTR: הטור הראשון נשאר מימין גם במקטע לועזי',
          `שורה 01 ב-x=${atL(1).x}, אמצע ${ltr.pageMiddle} — הכיוון אינו מה שמזיז את הטורים`,
        );
      }

      const anchorLine = lastOfFirstL - 2;
      if (!(await placeCaret(atL(anchorLine)))) {
        report.stuck('בקרת LTR: Shift+חץ מטה', 'הלחיצה לא מיקמה סמן נקי');
      } else {
        await app.press('ArrowDown', 'ArrowDown', 40, 8);
        await sleep(200);
        const step = await selectedLines();
        if (!(step.count === 1 && step.first === anchorLine + 1)) {
          report.stuck('בקרת LTR: Shift+חץ מטה', `הקשה בודדת בחרה ${step.label} — אין ממה להסיק`);
        } else {
          for (let i = 0; i < 5; i++) {
            await app.press('ArrowDown', 'ArrowDown', 40, 8);
            await sleep(200);
          }
          const got = await selectedLines();
          const crossed = got.first === anchorLine + 1 && got.last >= lastOfFirstL + 1;
          if (crossed) {
            report.pass('בקרת LTR: Shift+חץ מטה חוצה את הגבול — כלומר הפער שלמעלה הוא של RTL', got.label);
          } else {
            report.fail(
              'בקרת LTR: Shift+חץ מטה נשבר גם במקטע לועזי — הפער אינו של `w:bidi`',
              `עוגן ${anchorLine} (אומת: ${step.label}); 6 הקשות: ${got.label}; הגבול אחרי ${lastOfFirstL} — ` +
                'זהה למקטע העברי, ולכן זה פער של טורים ולא של כיוון',
            );
          }
        }
      }
    }
  }

  report.print();
} catch (error) {
  console.error('הגשש נפל:', error?.stack ?? error);
  process.exitCode = 1;
} finally {
  app.close();
}
