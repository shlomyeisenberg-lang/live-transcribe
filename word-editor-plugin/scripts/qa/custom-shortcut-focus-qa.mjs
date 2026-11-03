/**
 * שער QA לקיצור אישי לפי **מיקום המיקוד** — שתי הרגרסיות שדווחו על „החלת
 * העיצוב”, ושתיהן בלתי נראות ב-jsdom.
 *
 * ## למה זה לא יכול לחיות ב-tests/
 *
 *   1. **„הצירוף מקליד את התו”.** המנוע מחזיק מאזין `keydown` בשלב ה-capture
 *      על `.v2-super-editor__stage`, והוא מטפל ב-`Alt+<מקש>` כהקלדת תו: הוא
 *      מכניס את התו לבחירה וקורא `preventDefault`. המנתב שלנו יושב ב-bubble,
 *      ולכן נטש על `defaultPrevented` — הקיצור לא רץ, ובמקומו נכתבה „ס”
 *      במסמך (פריסה עברית). אין מנוע ב-jsdom, ולכן אין את המאזין ההוא, ולכן
 *      הבאג אינו קיים שם. הבעלות נמדדה ב-custom-shortcut-owner-probe.mjs.
 *   2. **„הוחל, ושום דבר לא השתנה”.** ההוכחה היא ה-OOXML שיוצא מול מה
 *      שנכתב בשורת המצב, ולא „הפקודה החזירה true” — המנוע **מדווח הצלחה**
 *      במצב הזה ומכניס את הערכה ל-stored marks.
 *
 * ## למה ארבע עליות ולא אחת
 *
 * שלושת המצבים אינם מפרידים ביניהם במסמך אחד: ההחלה הראשונה משנה את הבחירה,
 * ו„העיצוב הקודם” של המתג הופך את הלחיצה הבאה להחזרה. עלייה לכל מצב היא מה
 * שמפריד — ונמדד שבלעדיה שורה עברה על שארית של קודמתה.
 *
 *   npm run build && node scripts/qa/custom-shortcut-focus-qa.mjs
 */
import { openApp, createReport, sleep } from './harness.mjs';

const ALT = 1;
const SHIFT = 8;

/**
 * `Alt+X` בלבד, בלי Ctrl — זה הצירוף שהבאג חי בו, וזה מה שהמשתמש הצמיד.
 * `KeyQ` אינו משמש כאן: הוא „ספר לי” ברג'יסטרי, ורשומה אישית עליו נושרת.
 */
const SEEDED = [
  {
    id: 'alt-x',
    name: 'ערכת עבודה',
    kind: 'format-preset',
    combo: { code: 'KeyX', ctrl: false, shift: false, alt: true },
    preset: { fontSizePt: 22, bold: true },
  },
];

const EXTRA = `<script>
(function () {
  function install() {
    if (!window.__qaHost) return setTimeout(install, 10);
    var H = window.__qaHost;
    H.storage['custom-shortcuts'] = ${JSON.stringify(SEEDED)};
    H.replies['storage.get'] = function (payload) {
      var key = payload && payload.key;
      return Promise.resolve({ success: true, data: H.storage[key], error: null });
    };
  }
  install();

  /* יומן שורת המצב: כל טקסט שנכתב אליה, בסדר. */
  window.__statusLog = [];
  var last = null;
  setInterval(function () {
    var el = document.querySelector('.status-message');
    if (!el) return;
    var text = (el.textContent || '').trim();
    if (text === last) return;
    last = text;
    if (text !== '') window.__statusLog.push(text);
  }, 25);
})();
</script>`;

const report = createReport('קיצור אישי לפי מיקום המיקוד', { strict: true });

const boldTags = (xml) =>
  (xml.match(/<w:b(?:\s[^>]*)?\/>/g) ?? []).filter((tag) => !tag.startsWith('<w:bCs'));
const boldOn = (xml) => boldTags(xml).some((tag) => !/w:val="(0|false|off)"/.test(tag));
const hasSize44 = (xml) => /<w:sz\s+w:val="44"\s*\/>/.test(xml);

/** הטקסט שהמנוע צייר — כדי לראות שתו לא נכנס למסמך. */
const screen = (app) =>
  app.js("(document.querySelector('.editor-stack').innerText||'').replace(/\\s+/g,' ').trim()");

const statusLog = (app) => app.js('JSON.stringify(window.__statusLog)').then(JSON.parse);

/** בחירת „abcd” בפסקה הראשונה. זה מה שהערכה אמורה להיכתב עליו. */
async function selectWord(app) {
  await app.caret(0);
  await app.type('abcd', 60);
  await sleep(600);
  await app.press('Home', 'Home', 36);
  await sleep(200);
  for (let i = 0; i < 4; i += 1) {
    await app.press('ArrowRight', 'ArrowRight', 39, SHIFT);
    await sleep(60);
  }
  await sleep(800);
}

const pressAltX = (app) => app.press('x', 'KeyX', 88, ALT);

async function withApp(port, name, run) {
  const app = await openApp({ name, port, extra: EXTRA });
  try {
    await run(app);
  } finally {
    app.close();
  }
}

/* ================================================================== */
/* 1 — הסמן במסמך: הערכה מוחלת, והתו אינו נכתב                        */
/* ================================================================== */
await withApp(9666, 'csf1', async (app) => {
  await selectWord(app);
  await app.js('window.__statusLog.length = 0');
  const textBefore = await screen(app);

  await pressAltX(app);
  await sleep(1_600);

  const xml = (await app.docx())?.['word/document.xml'] ?? '';
  const textAfter = await screen(app);
  const log = await statusLog(app);

  if (hasSize44(xml) && boldOn(xml)) {
    report.pass('הסמן במסמך — הערכה ב-OOXML', `w:sz=44 ו-${JSON.stringify(boldTags(xml))}`);
  } else {
    report.fail(
      'הסמן במסמך — הערכה ב-OOXML',
      `w:sz=44 ${hasSize44(xml) ? 'נמצא' : 'חסר'}, תגי w:b: ${JSON.stringify(boldTags(xml))}`,
    );
  }

  /*
   * השורה שהבאג נופל עליה. לפני התיקון המנוע החליף את הבחירה בתו של המקש
   * („abcd” → „x”, ובפריסה עברית „ס”), והקיצור עצמו לא רץ כלל.
   */
  if (textAfter === textBefore) {
    report.pass('הסמן במסמך — התו אינו נכתב במסמך', `הטקסט נשאר „${textAfter}”`);
  } else {
    report.fail('הסמן במסמך — התו אינו נכתב במסמך', `„${textBefore}” הפך ל„${textAfter}”`);
  }

  const applied = log.find((text) => text.includes('הוחל'));
  if (applied) report.pass('הסמן במסמך — הודעת ההחלה', `„${applied}”`);
  else report.fail('הסמן במסמך — הודעת ההחלה', `יומן שורת המצב: ${JSON.stringify(log)}`);
});

/* ================================================================== */
/* 2 — המיקוד ברצועה ואין בחירה: סירוב, ולא „הוחל”                    */
/* ================================================================== */
await withApp(9666, 'csf2', async (app) => {
  /* לשונית ברצועה גוזלת את המיקוד, והמסמך ריק — אין לאן להחיל. */
  await app.tab('הוספה');
  await sleep(500);
  await app.js('window.__statusLog.length = 0');

  await pressAltX(app);
  await sleep(1_600);

  const xml = (await app.docx())?.['word/document.xml'] ?? '';
  const log = await statusLog(app);

  if (log.some((text) => text.includes('יש למקם את הסמן במסמך'))) {
    report.pass('אין לאן להחיל — נאמר מה חסר', `„${log[log.length - 1]}”`);
  } else {
    report.fail('אין לאן להחיל — נאמר מה חסר', `יומן שורת המצב: ${JSON.stringify(log)}`);
  }

  if (log.some((text) => text.includes('הוחל'))) {
    report.fail('אין לאן להחיל — לא מוכרז „הוחל”', `יומן שורת המצב: ${JSON.stringify(log)}`);
  } else {
    report.pass('אין לאן להחיל — לא מוכרז „הוחל”', 'שורת המצב אינה מכריזה על החלה');
  }

  if (!hasSize44(xml) && !boldOn(xml)) {
    report.pass('אין לאן להחיל — המסמך לא נגע', 'אין w:sz=44 ואין w:b');
  } else {
    report.fail(
      'אין לאן להחיל — המסמך לא נגע',
      `w:sz=44 ${hasSize44(xml) ? 'נמצא' : 'חסר'}, תגי w:b: ${JSON.stringify(boldTags(xml))}`,
    );
  }
});

/* ================================================================== */
/* 3 — המיקוד מחוץ למסמך אך הבחירה קיימת: מוחל כרגיל                  */
/* ================================================================== */
await withApp(9666, 'csf3', async (app) => {
  await selectWord(app);
  /*
   * לחיצה על פקד ברצועה גוזלת את המיקוד מהעורך והבחירה נשארת. זה המצב שבו
   * ההחלה **חייבת** לעבוד — סירוב שנשען על המיקוד לבדו היה שובר אותו.
   */
  await app.js('document.activeElement && document.activeElement.blur(); document.body.focus();');
  await sleep(500);
  await app.js('window.__statusLog.length = 0');

  await pressAltX(app);
  await sleep(1_600);

  const xml = (await app.docx())?.['word/document.xml'] ?? '';
  const log = await statusLog(app);

  if (hasSize44(xml) && boldOn(xml)) {
    report.pass('מיקוד מחוץ למסמך עם בחירה — מוחל', `w:sz=44 ו-${JSON.stringify(boldTags(xml))}`);
  } else {
    report.fail(
      'מיקוד מחוץ למסמך עם בחירה — מוחל',
      `w:sz=44 ${hasSize44(xml) ? 'נמצא' : 'חסר'}, תגי w:b: ${JSON.stringify(boldTags(xml))}, ` +
        `שורת המצב: ${JSON.stringify(log)}`,
    );
  }
});

/* ================================================================== */
/* 4 — הרשומות המובנות של Alt: הספרה אינה נכתבת במסמך                 */
/* ================================================================== */
await withApp(9666, 'csf4', async (app) => {
  /*
   * אותו מאזין של המנוע בלע גם 10 רשומות **מובנות** — `Alt+1`…`Alt+9`
   * (מעבר בין מסמכים פתוחים) ו-`Alt+Q`. הן לא רצו מעולם כשהסמן במסמך,
   * ובמקומן נכתבה הספרה. את המעבר עצמו אי אפשר למדוד כאן (מסמך אחד פתוח,
   * ו„המסמך האחרון” הוא זה), אבל את מה שנכתב במסמך — כן, וזו בדיוק הרגרסיה:
   * הקשה שאמורה לנווט ובמקום זה מקלידה.
   */
  await selectWord(app);
  const before = await screen(app);

  await app.press('1', 'Digit1', 49, ALT);
  await sleep(700);
  await app.press('9', 'Digit9', 57, ALT);
  await sleep(700);

  const after = await screen(app);
  if (after === before) {
    report.pass('Alt+ספרה מובנה — הספרה אינה נכתבת במסמך', `הטקסט נשאר „${after}”`);
  } else {
    report.fail('Alt+ספרה מובנה — הספרה אינה נכתבת במסמך', `„${before}” הפך ל„${after}”`);
  }

  /*
   * `Alt+F8` **לפני** `Alt+Q`, ובכוונה: „ספר לי” גוזל את המיקוד לשדה טקסט
   * של הממשק, ו-`macro-manage` אינה `inTextEntry` — כלומר אחרי `Alt+Q` היא
   * מסרבת בדין, והשורה הייתה מודדת את סדר הלחיצות ולא את הקיצור.
   *
   * והיא נמדדת בכלל מפני שהיא הרשומה היחידה בשלב הזה שהמנוע **לא** בלע
   * (מקש פונקציה אינו מפיק תו — ראו `runsBeforeEngine`), ולכן היא היחידה
   * שהתיקון הזיז בלי שהיה לה באג משלה. מעבר בין שלבים בלי עדות הוא בדיוק מה
   * שנשבר בשקט.
   */
  await app.press('F8', 'F8', 119, ALT);
  await sleep(900);
  const macros = await app.js(
    "(function(){return JSON.stringify(!!document.querySelector('.macros-dialog'));})()",
  ).then(JSON.parse);
  const textAfterF8 = await screen(app);

  if (macros && textAfterF8 === before) {
    report.pass('Alt+F8 מובנה — ניהול מאקרו נפתח', `הטקסט נשאר „${textAfterF8}”`);
  } else {
    report.fail(
      'Alt+F8 מובנה — ניהול מאקרו נפתח',
      `הדיאלוג ${macros ? 'נפתח' : 'לא נפתח'}, הטקסט „${textAfterF8}”`,
    );
  }

  await app.js("(function(){var b=document.querySelector('.macros-dialog .md-close-btn');return b?(b.click(),1):0;})()");
  await sleep(400);
  await app.caret(0);
  await sleep(300);

  /*
   * הצד החיובי של אותה רשומה: `Alt+Q` הוא „ספר לי”, ופתיחת התיבה היא משהו
   * שאפשר לראות. שער `tell-me-qa` מודד את התיבה עצמה, אבל הוא פותח אותה
   * בלחיצה — כלומר הקיצור שלה מעולם לא נמדד.
   */
  await app.press('q', 'KeyQ', 81, ALT);
  await sleep(700);
  const tellMe = await app.js(
    "(function(){var i=document.querySelector('.tell-me-input');" +
      'return JSON.stringify({ found: !!i, focused: i === document.activeElement });})()',
  ).then(JSON.parse);
  const textAfterQ = await screen(app);

  /*
   * `focused` ולא `found` לבדו: תיבת „ספר לי” היא חלק מהרצועה והיא **תמיד**
   * ב-DOM (נמדד — היא רושמת את מאזין ה-`keydown` הראשון בדף, לפני שהמנוע
   * עולה). שורה שנשענת על קיומה בלבד הייתה ירוקה גם על קיצור שאינו עושה
   * דבר; מה שהקיצור מבטיח הוא המיקוד.
   */
  if (tellMe.focused && textAfterQ === before) {
    report.pass('Alt+Q מובנה — „ספר לי” נפתח', `ממוקד=${tellMe.focused}, הטקסט נשאר „${textAfterQ}”`);
  } else {
    report.fail(
      'Alt+Q מובנה — „ספר לי” נפתח',
      `התיבה ${tellMe.found ? 'קיימת' : 'חסרה'}, ממוקדת=${tellMe.focused}, הטקסט „${textAfterQ}”`,
    );
  }
});

report.print();
