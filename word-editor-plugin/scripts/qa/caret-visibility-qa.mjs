/**
 * שער QA: הסמן נשאר בתצוגה כשמקלידים בחלון צר.
 *
 * ## הדיווח והמדידה
 *
 * „יש בעיה בהקלדה כשהחלון קטן — לא מזיזה את הדף מצד לצד.”
 *
 * עמוד A4 הוא 794px, ומיכל הגלילה צר ממנו ברגע שהחלון יורד מתחת ל-‎~800px.
 * המיכל הוא `direction: ltr` (הצהרה על צד פס הגלילה, styles/shell.css) ולכן
 * נח על `scrollLeft = 0` — הקצה השמאלי, שבמסמך עברי הוא *סוף* השורה. נמדד
 * בחלון 600px לפני התיקון: הבחירה בתחילת השורה, הסמן על `x = 698`, התצוגה
 * נגמרת ב-600, ו-`scrollLeft` נשאר 0 מתוך 209 זמינים — גם אחרי הקלדה שנכנסה
 * ל-OOXML בפועל. כלומר המשתמש מקליד, הטקסט נשמר, ודבר מזה אינו נראה.
 *
 * ## למה זה לא יכול להיות בדיקת רכיב
 *
 * „הסמן מחוץ לתצוגה” הוא יחס בין שני מלבנים אחרי פריסה. ב-jsdom
 * `getBoundingClientRect` מחזיר אפסים ו-`scrollWidth` הוא 0 תמיד — כלומר
 * הבאג עובר שם ירוק בכל צורה שלו, וכך גם כל תיקון שבור.
 * tests/unit/caret-visibility.test.ts מודד את הנוסחה; המפגש שלה עם פריסה
 * אמיתית נמדד רק כאן.
 *
 * ## מה נמדד
 *
 * הסמן מונח בחלון רחב ורק אז החלון מוקטן — זה גם התרחיש שדווח, וגם מה
 * שמונע לחיצה על נקודה שכבר אינה על המסך. בכל צעד אחרי ההקטנה נמדד מלבן
 * הסמן מול תיבת התוכן של מיכל הגלילה.
 *
 * ושלוש בקרות, בלעדיהן השער ירוק גם על ממשק שבור:
 *
 *   1. **המצב באמת מסוכן** — `scrollWidth > clientWidth` אחרי ההקטנה. בלי
 *      זה „הסמן בתצוגה” נכון מאליו, ואינו מעיד על דבר.
 *   2. **הגלילה באמת זזה** — `scrollLeft > 0` אחרי שהסמן הגיע לתחילת השורה.
 *      זה בדיוק המספר שהיה 0 לפני התיקון.
 *   3. **חלון רחב אינו נגרר** — בלי גלישה אסור שהגלילה תזוז בכלל.
 *
 * וארבע שורות שמחזיקות את ההנחות שהמודול בנוי עליהן, ולא את התנהגותו: הצורה
 * שהמנוע מדווח (יש בה `right`), שהגלילה שלנו אכן נשמעת כ-`scroll` על אותו
 * מיכל שמודולי ההשלמה מאזינים לו, שהדפדפן **מעגל** את מה שנכתב ל-`scrollLeft`,
 * ושהפער ב-`ui.viewport.scrollIntoView` עדיין קיים. אם אחת מהן תיפול, המודול
 * צריך להיכתב אחרת או לא להיכתב כלל.
 *
 * ## והצרכן, שקודם לא נמדד כלל
 *
 * שתי השורות האחרונות הן `isCaretFollowScroll` מקצה לקצה: רשימת המקורות של
 * „@” נפתחת בחלון צר, השאילתה מוקלדת בזמן שגלילת המעקב רצה, והרשימה חייבת
 * להישאר פתוחה — ובקרה הפוכה, שגלגל עכבר אמיתי על אותו מיכל **כן** סוגר
 * אותה. הגרסה הראשונה של ההבחנה נכשלה בשתיהן (היא נתפסה כמעט אף פעם, וכשכן
 * — לתמיד), וכל שאר השורות כאן נשארו ירוקות. זו הסיבה שהן קיימות.
 *
 * יציאה 9650 בלבד — שערים אחרים רצים במקביל על יציאות אחרות.
 *
 *   npm run build && node scripts/qa/caret-visibility-qa.mjs
 */
import { openApp, createReport, sleep } from './harness.mjs';

/**
 * הרוחב שבו העמוד גולש. נמדד: חלון 600 נותן מאגס `clientWidth` 585 (פס
 * הגלילה האנכי גורע את ההפרש) מול עמוד של 794 — כלומר 209px גלילה זמינה.
 */
const NARROW = 600;
/** רוחב הבקרה: העמוד נכנס בשלמותו, ואין גלישה כלל. */
const WIDE = 1100;

const report = createReport('הסמן בתצוגה בחלון צר', { strict: true });

const app = await openApp({ name: 'caret-visibility', port: 9650 });

/**
 * מיכל הגלילה של המסמך **הפעיל**, כביטוי לתוך הדף.
 *
 * ולא `document.querySelector('.editor-stack__host')`: זה מחזיר את הראשון
 * ב-DOM, והאפליקציה משאירה את הפאנל של המסמך הקודם במקומו עם `display: none`
 * (App.vue, `activateTab`) — מיכל שה-`clientWidth` שלו 0, וכל מדידה עליו
 * חסרת משמעות. `ui.viewport.getHost()` הוא בדיוק מה ש-`installCaretVisibility`
 * מקבל דרך `paintedHost`, ונמדד שהוא אותו אלמנט שמודולי ההשלמה מאזינים לו
 * (`editor.container === ui.viewport.getHost()`).
 */
const HOST = `(window.__otzariaEditor?.ui?.viewport?.getHost?.() ?? null)`;

/**
 * מלבן הסמן מול תיבת התוכן של מיכל הגלילה.
 *
 * הסמן נקרא מ-`ui.selection.getAnchorRect()` ולא מה-DOM: נמדד שאלמנט הסמן
 * נהרס ונוצר מחדש בכל הקשה, ושהוא חסר במשך ‎~130ms אחריה — כלומר מדידה
 * ממנו היא מירוץ. ה-API מדווח את אותו מקום בדיוק (נמדד: 777.47 מול 777).
 *
 * `clientWidth` ולא `getBoundingClientRect().width`: פס הגלילה האנכי גורע
 * 15px, וסמן שיושב מתחתיו אינו נראה.
 */
const measure = () =>
  app
    .js(`(() => {
      const host = ${HOST};
      if (!host) return JSON.stringify({ error: 'אין מיכל גלילה' });
      let rect = null;
      try { rect = window.__otzariaEditor?.ui?.selection?.getAnchorRect?.() ?? null; } catch (e) { rect = null; }
      const box = host.getBoundingClientRect();
      const viewLeft = box.left + host.clientLeft;
      return JSON.stringify({
        scrollLeft: Math.round(host.scrollLeft),
        overflow: Math.round(host.scrollWidth - host.clientWidth),
        viewLeft: Math.round(viewLeft),
        viewRight: Math.round(viewLeft + host.clientWidth),
        caretLeft: rect && typeof rect.left === 'number' ? Math.round(rect.left) : null,
        caretRight: rect && typeof rect.right === 'number' ? Math.round(rect.right) : null,
        caretTop: rect && typeof rect.top === 'number' ? Math.round(rect.top) : null,
      });
    })()`)
    .then(JSON.parse);

const resize = async (width) => {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width, height: 900, deviceScaleFactor: 1, mobile: false });
  await sleep(1_200);
};

/** תיאור קצר למדידה, לשורת הדיווח. */
const where = (m) =>
  `סמן ${m.caretLeft}..${m.caretRight}, תצוגה ${m.viewLeft}..${m.viewRight}, scrollLeft ${m.scrollLeft}/${m.overflow}`;

/** האם הסמן בתוך תיבת התוכן. סבילות פיקסל אחד — עיגול תת-פיקסלי. */
const inside = (m) => m.caretLeft !== null && m.caretLeft >= m.viewLeft - 1 && m.caretRight <= m.viewRight + 1;

try {
  /* -------------------- הסמן מונח בחלון רחב -------------------- */
  await resize(WIDE);
  await app.caretPara(0);
  await sleep(300);
  await app.type('שלום עולם זהו משפט בדיקה', 25);
  await sleep(700);

  const wide = await measure();
  if (wide.overflow <= 0) {
    report.pass('בקרה 3 — בחלון רחב אין גלישה', where(wide));
    if (wide.scrollLeft === 0) report.pass('בקרה 3 — והגלילה לא זזה', `scrollLeft ${wide.scrollLeft}`);
    else report.fail('בקרה 3 — הגלילה זזה בלי סיבה', where(wide));
  } else {
    report.fail('בקרה 3 — חלון רחב אמור להכיל את העמוד', where(wide));
  }

  /* -------------------- ההקטנה -------------------- */
  await resize(NARROW);
  const narrow = await measure();
  if (narrow.overflow > 0) report.pass('בקרה 1 — אחרי ההקטנה יש גלישה אופקית', `${narrow.overflow}px`);
  else report.fail('בקרה 1 — אין גלישה, והמדידה שאחריה חסרת משמעות', where(narrow));

  /* -------------------- הקלדה בחלון הצר -------------------- */
  /*
   * ההקלדה נמשכת עד ש**השורה גולשת**, ולא מספר צעדים שרירותי: בעברית
   * ההקלדה מזיזה את הסמן שמאלה, כלומר *פנימה*, ולכן שורה אחת אינה מבחינה
   * בין תקין לשבור. נמדד — סבב מוטציה על המודול השאיר את הצעד הזה ירוק.
   * מה שכן מבחין הוא הרגע שבו הטקסט עובר לשורה הבאה והסמן קופץ בחזרה לקצה
   * הימני. הגלישה מזוהה לפי `top` של הסמן, שיורד בדיוק אז.
   */
  const firstLineTop = (await measure()).caretTop;
  let broke = null;
  let wrapped = null;
  for (let step = 0; step < 12 && !broke && !wrapped; step += 1) {
    await app.type('ועוד מילים ', 25);
    await sleep(500);
    const m = await measure();
    if (!inside(m)) broke = `בצעד ${step + 1}: ${where(m)}`;
    else if (m.caretTop !== null && firstLineTop !== null && m.caretTop > firstLineTop) wrapped = where(m);
  }
  if (broke) report.fail('הקלדה בחלון צר — הסמן יצא מהתצוגה', broke);
  else if (wrapped) report.pass('גלישת שורה בחלון צר — הסמן נשאר בתצוגה', wrapped);
  else report.fail('הקלדה בחלון צר — השורה לא גלשה, והמדידה אינה מבחינה', where(await measure()));

  /* -------------------- שורה חדשה: הסמן קופץ לתחילת השורה -------------------- */
  /*
   * זה המקרה החמור, ולא הקצה: `Enter` בעברית מחזיר את הסמן לקצה **הימני**
   * של עמודת הטקסט — הנקודה שנמדדה על `x = 698` מול תצוגה שנגמרת ב-600.
   * אחרי `Enter` הפסקה ריקה, ולכן תו אחד מספיק כדי שיהיה מה למדוד.
   */
  await app.press('Enter', 'Enter', 13);
  await sleep(600);
  await app.type('א', 60);
  await sleep(700);

  const fresh = await measure();
  if (inside(fresh)) report.pass('שורה חדשה — הסמן בתחילת השורה נראה', where(fresh));
  else report.fail('שורה חדשה — הסמן בתחילת השורה מחוץ לתצוגה', where(fresh));

  if (fresh.scrollLeft > 0) report.pass('בקרה 2 — הגלילה האופקית אכן זזה', `scrollLeft ${fresh.scrollLeft}/${fresh.overflow}`);
  else report.fail('בקרה 2 — הגלילה נשארה באפס', where(fresh));

  /* -------------------- הטקסט באמת נכתב -------------------- */
  const files = await app.docx();
  const body = files['word/document.xml'] ?? '';
  if (body.includes('ועוד מילים')) report.pass('הטקסט שהוקלד נכתב למסמך', 'נמצא ב-word/document.xml');
  else report.fail('הטקסט שהוקלד לא הגיע למסמך', 'אין „ועוד מילים” ב-word/document.xml');

  /* -------------------- הצורה שהמנוע מדווח -------------------- */
  /*
   * המודול קורא `rect.left` ו-`rect.right` ישירות. שני מודולי ההשלמה
   * מכריזים על אותו מלבן כ-`{left, top, width, height}` — הכרזה על מה שהם
   * צורכים — ומי שיסיק מזה ש-`right` אינו קיים יגזור אותו מ-`left + width`
   * בלי סיבה. השורה הזאת היא מה שמכריע בין השניים, והיא גם מה שיאדים אם
   * המנוע יצמצם את המלבן בגרסה הבאה.
   */
  const shape = JSON.parse(
    await app.js(`JSON.stringify((function () {
      try {
        var r = window.__otzariaEditor.ui.selection.getAnchorRect({ placement: 'end' });
        return r ? { hasRect: true, keys: Object.keys(r), hasRight: typeof r.right === 'number' } : { hasRect: false };
      } catch (e) { return { error: String(e && e.message) }; }
    })())`),
  );
  if (shape.error) report.stuck('צורת המלבן שהמנוע מדווח', shape.error);
  // „אין מלבן” אינו „המלבן חדל לדווח `right`” — זה מצב רגיל של המנוע כשאין
  // בחירה פתורה, והמדידה פשוט לא התבצעה. דיווח כשל כאן היה שולח לחפש שינוי
  // בצורת המלבן במקום בחירה שנעלמה.
  else if (!shape.hasRect) report.stuck('צורת המלבן שהמנוע מדווח', 'אין בחירה פתורה — המנוע החזיר null');
  else if (shape.hasRight) report.pass('צורת המלבן שהמנוע מדווח', shape.keys.join(', '));
  else report.fail('המלבן חדל לדווח `right`', shape.keys.join(', '));

  /* -------------------- הגלילה נראית למי שמאזין -------------------- */
  /*
   * המיכל שהמודול גולל הוא בדיוק זה ששני מודולי ההשלמה מאזינים לו
   * (`engine/at-mention-overlay.ts`, `engine/book-completion-overlay.ts`),
   * ושניהם סוגרים את המושב שלהם על כל `scroll` — במשמעות „המשתמש גלל משם”.
   * השורה הזאת מודדת שהגלילה שלנו אכן מגיעה אליהם, וזה מה שמצדיק את
   * `isCaretFollowScroll`.
   */
  /*
   * ובאותה נשימה — **העיגול**. הערך נכתב כשבר, נקרא בחזרה, ושני המספרים
   * מדווחים כמו שהם. זו העובדה שעליה עומד כל התכנון של `isCaretFollowScroll`:
   * גרסה קודמת שלה השוותה את `scrollLeft` למספר ששלחנו, והדפדפן מעגל אותו —
   * כלומר ההשוואה כמעט אף פעם לא נתפסה. שורה שמעגלת את הקריאה לפני הדיווח
   * (וכך היה כאן) מסתירה בדיוק את זה.
   */
  const heard = JSON.parse(
    await app.js(`(async () => {
      var host = ${HOST};
      var seen = [];
      var onScroll = function (e) { seen.push({ at: host.scrollLeft, isHost: e.target === host }); };
      host.addEventListener('scroll', onScroll, true);
      var was = host.scrollLeft;
      var max = host.scrollWidth - host.clientWidth;
      var wrote = (was > 0 ? was - Math.min(40, was) : Math.min(40, max)) + 0.47;
      host.scrollLeft = wrote;
      // מיד, ובאותה משימה: אחרי המתנה כל דבר אחר בדף יכול להיות זה שהזיז את
      // המיכל, והשורה הייתה מדווחת „שונה” מסיבה שאינה עיגול.
      var read = host.scrollLeft;
      await new Promise(function (r) { setTimeout(r, 300); });
      host.removeEventListener('scroll', onScroll, true);
      host.scrollLeft = was;
      return JSON.stringify({ fired: seen.length, seen: seen, wrote: wrote, read: read });
    })()`),
  );
  if (heard.fired > 0) {
    const targets = heard.seen.every((s) => s.isHost) ? 'כולם על המיכל עצמו' : 'חלקם על אלמנט אחר';
    report.pass(
      'גלילה תכנותית נשמעת כ-scroll על אותו מיכל',
      `${heard.fired} אירועים, ${targets}, scrollLeft ${heard.seen.map((s) => s.at).join('/')}`,
    );
    if (heard.read !== heard.wrote) {
      report.pass('הדפדפן מעגל את מה שנכתב ל-scrollLeft', `נכתב ${heard.wrote}, נקרא ${heard.read}`);
    } else {
      report.partial(
        'הדפדפן מעגל את מה שנכתב ל-scrollLeft',
        `נכתב ${heard.wrote} ונקרא אותו דבר — ההנמקה ב-isCaretFollowScroll צריכה מדידה מחדש`,
      );
    }
  } else {
    report.stuck('גלילה תכנותית נשמעת כ-scroll על אותו מיכל', 'לא נורה אירוע — אין מה למדוד');
    report.stuck('הדפדפן מעגל את מה שנכתב ל-scrollLeft', 'אותה סיבה');
  }

  /* -------------------- הפער שבגללו המודול קיים -------------------- */
  /*
   * `ui.viewport.scrollIntoView` מדווח `{ success: true }` ואינו נוגע בציר
   * האופקי — נמדד לראשונה על 2.14.0-next.5, ונמדד כאן מחדש על מה שמותקן.
   * הבחירה מועברת לתחילת הפסקה (בעברית: הקצה הימני), הגלילה מוחזרת לאפס
   * **ביד** — אירוע בחירה אינו נורה, ולכן `follow` אינו מתקן אותה — ורק אז
   * נקרא ה-API. הבקרה היא השורה האחרונה: הגלילה זמינה, והסמן נכנס לתצוגה
   * ברגע שמבצעים אותה.
   *
   * זו שורת תיעוד, לא דרישה: אם המנוע יתחיל לגלול, היא תדווח „חלקי” וזה
   * הרמז לעדכן את docs/engine-gaps.md ולשקול להסיר את המודול.
   */
  const gap = JSON.parse(
    await app.js(`(async () => {
      var ui = window.__otzariaEditor.ui;
      var doc = window.__qa.doc();
      var host = ${HOST};
      var wait = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
      var cur = await doc.selection.current();
      var seg = ((cur && cur.target && cur.target.segments) || []).find(function (s) { return s.blockId; });
      if (!seg) return JSON.stringify({ error: 'אין כתובת בחירה' });
      var target = { kind: 'text', blockId: seg.blockId, range: { start: 0, end: 0 } };
      var at = { kind: 'text', blockId: seg.blockId, offset: 0 };
      ui.selection.apply({ kind: 'selection', start: at, end: at });
      await wait(800);
      var caretX = function () { try { var r = ui.selection.getAnchorRect(); return r ? Math.round(r.left) : null; } catch (e) { return null; } };
      var rows = [];
      var blocks = ['nearest', 'center', 'start'];
      for (var i = 0; i < blocks.length; i++) {
        host.scrollLeft = 0;
        await wait(150);
        var caretBefore = caretX();
        var answer = null;
        try { answer = await ui.viewport.scrollIntoView({ target: target, block: blocks[i], behavior: 'instant' }); }
        catch (e) { answer = { error: String(e && e.message) }; }
        await wait(300);
        rows.push({ block: blocks[i], answer: JSON.stringify(answer), after: Math.round(host.scrollLeft), caretBefore: caretBefore, caretAfter: caretX() });
      }
      host.scrollLeft = 0;
      await wait(150);
      var byHand = { before: caretX(), max: Math.round(host.scrollWidth - host.clientWidth) };
      host.scrollLeft = byHand.max;
      await wait(300);
      byHand.after = caretX();
      return JSON.stringify({ viewRight: Math.round(host.clientWidth), rows: rows, byHand: byHand });
    })()`),
  );
  const GAP_ROW = 'הפער: `scrollIntoView` מדווח הצלחה ואינו מזיז את הציר האופקי';
  if (gap.error) {
    report.stuck(GAP_ROW, gap.error);
  } else if (gap.rows.some((r) => r.caretBefore === null) || gap.byHand.max <= 0) {
    report.stuck(GAP_ROW, 'אין גלישה או אין מלבן סמן — אין מה למדוד');
  } else {
    const outside = gap.rows.filter((r) => r.caretBefore > gap.viewRight);
    const moved = gap.rows.filter((r) => r.after > 0);
    const detail = gap.rows
      .map((r) => `${r.block}: ${r.answer}, scrollLeft ${r.after}, סמן ${r.caretBefore}→${r.caretAfter}`)
      .join(' | ');
    const control = `בקרה ביד: scrollLeft ${gap.byHand.max}, סמן ${gap.byHand.before}→${gap.byHand.after}`;
    if (outside.length !== gap.rows.length) {
      report.stuck(GAP_ROW, `הסמן לא היה מחוץ לתצוגה (${gap.viewRight}) — ${detail}`);
    } else if (moved.length === 0) {
      report.pass(GAP_ROW, `${detail} — ${control}`);
    } else {
      report.partial(GAP_ROW, `המנוע כן גלל ב-${moved.length} מ-${gap.rows.length}: ${detail}`);
    }
  }

  /* -------------------- הצרכן: רשימת „@” תוך כדי הקלדה -------------------- */
  /*
   * זה התרחיש שבגללו `isCaretFollowScroll` נכתבה, וזו השורה היחידה שמודדת
   * אותו מקצה לקצה. בלעדיה אין בשער שום דבר שנוגע בהבחנה עצמה: שאר השורות
   * מודדות את צורת המלבן, שהגלילה נשמעת, ושהפער ב-`scrollIntoView` קיים —
   * וכולן נשארו ירוקות כשההבחנה לא עבדה בכלל.
   *
   * ההקלדה נמשכת עד שהסמן קרוב לקצה **השמאלי** של התצוגה, ורק אז מוקלד
   * האזכור: בעברית ההקלדה מזיזה את הסמן שמאלה, וגלילת המעקב מתרחשת בדיוק
   * כשהוא חוצה את המרווח שם. לכן השורה מוודאת קודם שנורתה גלילה תוך כדי
   * הקלדת השאילתה — בלי זה היא „ירוקה” בלי שמדדה דבר, ולכן היא מדווחת
   * `stuck` ולא `pass`.
   *
   * ובקרה: גלגל עכבר אמיתי על אותו מיכל **כן** סוגר. בלעדיה כל השורה עוברת
   * גם על מודול שהפסיק לסגור לגמרי — וזה בדיוק הכשל השני של הגרסה הקודמת,
   * שבה הסימן לא פג ובלע כל גלילה שהיא.
   */
  const MENTION_ROW = 'רשימת „@” נשארת פתוחה כשהסמן גורר גלילה';
  const WHEEL_ROW = 'בקרה: גלגל עכבר על אותו מיכל כן סוגר את הרשימה';

  await app.js(`window.__qaHost.replies['library.resolveRef'] = function () {
    return Promise.resolve({
      success: true,
      error: null,
      data: [{ id: 42, bookId: 'פסחים', bookUid: 'id:42', type: 'text', title: 'פסחים',
               reference: 'פסחים דף לד', index: 1234, isPdf: false, isSourceLine: true,
               isUserBook: false, bookPath: 'ש"ס, בבלי' }],
    });
  };`);

  /* שורה חדשה ונקייה: הסמן מתחיל בקצה הימני, ומכאן הוא הולך שמאלה. */
  await app.caretPara(0);
  await app.press('End', 'End', 35);
  await sleep(300);
  await app.press('Enter', 'Enter', 13);
  await sleep(600);

  let filled = await measure();
  for (let step = 0; step < 14 && filled.caretLeft !== null && filled.caretLeft > filled.viewLeft + 110; step += 1) {
    await app.type('מילה ', 25);
    await sleep(350);
    filled = await measure();
  }

  const arm = () =>
    app.js(`(function () {
      var host = ${HOST};
      window.__followProbe = { events: 0, from: host.scrollLeft };
      window.__followOn = function () { window.__followProbe.events += 1; };
      host.addEventListener('scroll', window.__followOn, true);
    })()`);
  const disarm = () =>
    app
      .js(`(function () {
        var host = ${HOST};
        host.removeEventListener('scroll', window.__followOn, true);
        return JSON.stringify({
          events: window.__followProbe.events,
          from: window.__followProbe.from,
          to: host.scrollLeft,
          open: !!document.querySelector('.otzaria-at-mention'),
        });
      })()`)
      .then(JSON.parse);

  /*
   * הרשימה נפתחת על התו הראשון, ומסומנת — ורק אז נמדדת ההקלדה שאחריו.
   *
   * ‏**הסימון הוא העיקר.** „פתוחה בסוף” אינו מודד דבר: כל הקשה מריצה
   * ‏`evaluate` מחדש, ולכן רשימה שנסגרה באמצע נפתחת שוב מיד אחריה, והמדידה
   * בסוף רואה רשימה פתוחה גם כשההבחנה שבורה לגמרי — נמדד: סבב מוטציה שהחזיר
   * את ההשוואה למיקום השאיר את השורה הזאת ירוקה כשהיא נשאלה רק בסוף.
   * ‏`ensurePopup` שומר את אותו אלמנט כל עוד המושב חי (`closeSession` מוחק
   * אותו), ולכן סימון על האלמנט הוא בדיוק „המושב לא נסגר מאז”.
   */
  // שני תווים, ולא אחד: `MIN_QUERY_LENGTH` הוא 2, ועל תו אחד הרשימה אינה
  // נפתחת כלל — נמדד (השורה דיווחה „לא נפתחה כלל”).
  await app.type('@פס', 45);
  await sleep(900);
  const opened = JSON.parse(
    await app.js(`JSON.stringify((function () {
      var el = document.querySelector('.otzaria-at-mention');
      if (!el) return { open: false };
      el.__qaKeep = true;
      return { open: true };
    })())`),
  );

  if (!opened.open) {
    report.stuck(MENTION_ROW, 'הרשימה לא נפתחה כלל — אין מה למדוד');
    report.skip(WHEEL_ROW, 'אותה סיבה');
  } else {
    /*
     * ההקלדה נמשכת **עד שנורתה גלילה**, ולא מספר תווים קבוע: כמה תווים נשארו
     * עד שהסמן חוצה את המרווח תלוי במקום שבו הפילר עצר ובשאלה אם השורה גלשה
     * בינתיים (נמדד: שלושה תווים קבועים נתנו 0 אירועי גלילה, כלומר שורה שאינה
     * מודדת דבר). השאילתה נשארת מילה אחת — רווח היה מסיים אותה.
     */
    await arm();
    const samples = [];
    for (const ch of 'חיםפסחיםפסחיםפסחיםפסחים') {
      await app.type(ch, 45);
      await sleep(250);
      samples.push(
        JSON.parse(
          await app.js(`JSON.stringify((function () {
            var el = document.querySelector('.otzaria-at-mention');
            return {
              ch: ${JSON.stringify(ch)},
              present: !!el,
              kept: !!(el && el.__qaKeep),
              events: window.__followProbe.events,
            };
          })())`),
        ),
      );
      if (samples[samples.length - 1].events > 0) break;
    }
    await sleep(600);
    const mention = await disarm();
    const lost = samples.filter((s) => !s.present || !s.kept);
    const trace = samples
      .map((s) => `${s.ch}:${s.present ? (s.kept ? 'אותה רשימה' : '**נפתחה מחדש**') : '**אין**'}`)
      .join(', ');
    const mentionWhere = `${mention.events} אירועי גלילה, scrollLeft ${mention.from}→${mention.to} — ${trace}`;

    if (mention.events === 0) {
      report.stuck(MENTION_ROW, `לא נורתה גלילה תוך כדי השאילתה — ${mentionWhere}`);
    } else if (lost.length === 0 && mention.open) {
      report.pass(MENTION_ROW, mentionWhere);
    } else {
      report.fail(MENTION_ROW, `המושב נסגר באמצע ההקלדה — ${mentionWhere}`);
    }

    /*
     * נקודת הגלגל נלקחת מקצהו העליון של המיכל ולא ממרכזו: הרשימה נפתחת מתחת
     * לסמן, ואירוע גלגל שנוחת עליה אינו גולל את המיכל שמאחוריה.
     *
     * הבקרה רצה גם כשהשורה שמעליה נכשלה: „הרשימה נסגרה כשצריך” הוא בדיוק מה
     * שמבדיל בין הבחנה שעובדת לבין אחת שבולעת הכול, וזה נכון במיוחד כשהראשונה
     * אדומה.
     */
    const at = JSON.parse(
      await app.js(`JSON.stringify((function () {
        var b = ${HOST}.getBoundingClientRect();
        return { x: Math.round(b.left + b.width / 2), y: Math.round(b.top + 20) };
      })())`),
    );
    await arm();
    await app.cdp.send('Input.dispatchMouseEvent', {
      type: 'mouseWheel',
      x: at.x,
      y: at.y,
      deltaX: -60,
      deltaY: 0,
      button: 'none',
      buttons: 0,
      clickCount: 0,
    });
    await sleep(700);
    const wheel = await disarm();
    const wheelWhere = `${wheel.events} אירועי גלילה, scrollLeft ${wheel.from}→${wheel.to}`;
    if (wheel.events === 0) report.stuck(WHEEL_ROW, `הגלגל לא גלל את המיכל — ${wheelWhere}`);
    else if (wheel.open) report.fail(WHEEL_ROW, `הרשימה נשארה פתוחה אחרי גלילת משתמש — ${wheelWhere}`);
    else report.pass(WHEEL_ROW, wheelWhere);
  }

  await app.escape();

  const log = await app.log();
  const noise = log.filter((line) => !line.includes('reader.addContextMenuItem'));
  if (noise.length === 0) report.pass('אין שגיאות בדף', `${log.length} שורות, כולן ידועות`);
  else report.fail('שגיאות בדף', noise.join(' | '));
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
