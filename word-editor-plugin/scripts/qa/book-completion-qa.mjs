/**
 * שער „השלמה מהספר” — המסלול הוותיק, שעד עכשיו לא היה מכוסה בכלל.
 *
 * הפער הזה אינו תיאורטי: הוא איפשר רגרסיה שנמדדה בפועל — אחרי ה-Tab הראשון
 * ה-ghost נעלם, אבל ההצעה נשארה **חיה**, ולכן ה-Tab הבא נתפס
 * (`preventDefault`) והכניס חמש מילים מהספר בלי ששום דבר הוצג. שני השערים
 * האחרים לא יכלו לתפוס זאת: המאחז אינו מממש את הקורא, ולכן אצלם `cache`
 * נשאר null וההשלמה מהספר כלל אינה נכנסת למסלול.
 *
 * לכן השער הזה מדמה את הקורא של אוצריא (`reader.getCurrentState`,
 * `reader.getSectionTextMap`), ומודד את מה שנשבר: **מה שמוצג שווה למה
 * שנכתב**, גם בסבב ההמשך.
 *
 * הרצה:  node scripts/qa/book-completion-qa.mjs
 * היציאה 9615 שמורה לשער הזה בלבד.
 */
import { openApp, createReport } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9615);

/** ה"ספר" שהמאחז מחזיר. שורה אחת, בלי ניקוד — הנרמול נבדק במקום אחר. */
const LINE = 'בראשית ברא אלהים את השמים ואת הארץ והארץ היתה תהו ובהו וחשך על פני תהום';
/** מה שמוקלד, ומה שאמור להיות מוצע: חמש המילים שאחרי „בראשית”. */
const TYPED = 'בראשית ב';
const FIRST_GHOST = 'ברא אלהים את השמים ואת';

const report = createReport('שער השלמה מהספר', { strict: true });
const app = await openApp({ name: 'book-completion', port: PORT });

const log = (...a) => console.log(...a);

const ghostText = () =>
  app.js("(document.querySelector('.otzaria-book-completion-ghost') || {}).textContent || ''");

async function docText() {
  const files = (await app.docx()) ?? {};
  return (files['word/document.xml'] ?? '').replace(/<[^>]+>/g, '').replace(/\r?\n/g, '');
}

async function step(name, fn) {
  log(`\n──────── ${name} ────────`);
  try {
    await fn();
  } catch (error) {
    log('!! זרק:', error?.message);
    report.fail(name, `הצעד זרק: ${error?.message}`);
  }
}

try {
  await step('הקמת הקורא המדומה והדלקת ההשלמה', async () => {
    await app.js(`
      window.__qaHost.replies['reader.getCurrentState'] = function () {
        return Promise.resolve({ success: true, data: { currentBook: 'בראשית', currentBookId: 'b1', currentRef: 'בראשית א', currentIndex: 0 }, error: null });
      };
      window.__qaHost.replies['reader.getSectionTextMap'] = function (payload) {
        var text = payload && payload.sectionIndex === 0 ? ${JSON.stringify(LINE)} : '';
        return Promise.resolve({ success: true, data: { renderedText: text, sourceText: text, hasMore: false, nextCursor: null }, error: null });
      };
    `);

    await app.tab('✦ אוצריא');
    const clicked = await app.click('השלמה מהספר', { after: 900 });
    if (!clicked) return report.fail('הדלקת ההשלמה', 'הכפתור „השלמה מהספר” לא נמצא');
    const state = await app.state('השלמה מהספר');
    if (state?.active !== true) return report.fail('הדלקת ההשלמה', `הכפתור לא נדלק: ${JSON.stringify(state)}`);
    report.pass('הדלקת ההשלמה');
  });

  /** מה שההצעה הראשונה הציגה — נשמר כדי להשוות אל מה שנכתב. */
  let firstGhost = '';

  await step('הקלדה מהספר מציגה את המשך הפסוק', async () => {
    await app.caretPara(0);
    await app.type(TYPED);
    await app.sleep(1_600);

    firstGhost = (await ghostText()).trim();
    log('ghost:', JSON.stringify(firstGhost));
    if (firstGhost === FIRST_GHOST) report.pass('ה-ghost מציג את המשך הספר');
    else report.fail('ה-ghost מציג את המשך הספר', `הוצג ${JSON.stringify(firstGhost)} במקום „${FIRST_GHOST}”`);
  });

  await step('Tab כותב בדיוק את מה שהוצג', async () => {
    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(1_000);

    const text = await docText();
    log('המסמך:', JSON.stringify(text));
    if (firstGhost !== '' && text.endsWith(firstGhost)) report.pass('מה שנכתב שווה למה שהוצג');
    else report.fail('מה שנכתב שווה למה שהוצג', `במסמך ${JSON.stringify(text)}, הוצג ${JSON.stringify(firstGhost)}`);
  });

  /*
   * הצעד שתופס את הרגרסיה. שתי טענות שחייבות להתקיים יחד, ודווקא הצירוף הוא
   * העיקר: הצעה חיה בלי ghost פירושה Tab שכותב למסמך בלי ששום דבר הוצג.
   */
  let continuation = '';

  await step('סבב ההמשך מוצג, ולא רק חי', async () => {
    continuation = (await ghostText()).trim();
    log('ghost ההמשך:', JSON.stringify(continuation));
    if (continuation !== '') report.pass('ה-ghost של ההמשך מוצג');
    else report.fail('ה-ghost של ההמשך מוצג', 'אין ghost אחרי ה-Tab הראשון');
  });

  await step('ה-Tab הבא כותב את מה שהוצג, ורק אם הוצג', async () => {
    const before = await docText();
    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(1_000);
    const after = await docText();
    log('לפני:', JSON.stringify(before), '| אחרי:', JSON.stringify(after));

    /*
     * הטענה הזאת היא שהייתה אדומה ברגרסיה: המסמך גדל אחרי Tab בעוד שלא הוצג
     * שום ghost — כלומר Tab נתפס וכתב חמש מילים בלי ששום דבר נראה.
     */
    if (after.length > before.length && continuation === '') {
      report.fail('Tab כותב רק כשהוצג', 'המסמך גדל אחרי Tab בלי ששום ghost הוצג');
    } else {
      report.pass('Tab כותב רק כשהוצג');
    }

    /*
     * ולא „מה שנוסף שווה ל-ghost”: ה-ghost מציג את הטקסט מהמילה שתאמה, כולל
     * מילת ההקשר שכבר הוקלדה, וההכנסה מחליפה אותה כדי שלא יישאר משפט
     * חצי-מנוקד (ר' book-completion.ts). לכן הטענה היא על **סוף המסמך**.
     */
    if (continuation !== '' && !after.endsWith(continuation)) {
      report.fail('סוף המסמך שווה למה שהוצג', `במסמך ${JSON.stringify(after)}, הוצג ${JSON.stringify(continuation)}`);
    } else {
      report.pass('סוף המסמך שווה למה שהוצג');
    }
  });

  /*
   * המילים שעלו זו על זו: ה-ghost נשאר בעמדת הסמן של ההערכה הקודמת בזמן
   * שההקלדה כבר מציירת אותיות חדשות בדיוק שם. לפני התיקון 11 מתוך 18 דגימות
   * באמצע הקלדה היו כאלה. הטענה: בכל רגע, ghost שמוצג יושב בסמן.
   */
  await step('ghost אינו נשאר מעל טקסט שמוקלד', async () => {
    const sample = () =>
      app.js(`(function(){
        var g = document.querySelector('.otzaria-book-completion-ghost');
        var sd = window.__otzariaEditor && window.__otzariaEditor.superdoc;
        var c = sd && sd.ui && sd.ui.selection && sd.ui.selection.getAnchorRect({placement:'end'});
        if (!g || !c) return null;
        var r = g.getBoundingClientRect();
        return JSON.stringify({ dx: Math.round(r.right - c.left), dy: Math.round(r.top - c.top) });
      })()`).then((raw) => (raw ? JSON.parse(raw) : null));

    const stale = [];
    for (const ch of 'הארץ') {
      await app.type(ch, 0);
      for (let i = 0; i < 3; i += 1) {
        await app.sleep(40);
        const s = await sample();
        if (s && (s.dx > 2 || Math.abs(s.dy) > 2)) stale.push(`${ch}:${s.dx},${s.dy}`);
      }
    }
    if (stale.length === 0) report.pass('ghost אינו נשאר מעל טקסט שמוקלד');
    else report.fail('ghost אינו נשאר מעל טקסט שמוקלד', `ghost מחוץ לסמן: ${stale.join(' ')}`);
  });

  await step('כיבוי מסיר את ה-ghost', async () => {
    await app.click('השלמה מהספר', { after: 700 });
    const ghost = await ghostText();
    if (ghost === '') report.pass('הכיבוי מנקה את ה-ghost');
    else report.fail('הכיבוי מנקה את ה-ghost', `נשאר ${JSON.stringify(ghost)}`);
  });

  await step('ללא רעש', async () => {
    const lines = await app.log();
    const noisy = (lines ?? []).filter(
      (l) =>
        /error|שגיאה|failed|נכשל/i.test(l) &&
        !/DevTools|Download the Vue/i.test(l) &&
        !/reader\.addContextMenuItem|fonts\.listInstalled/.test(l) &&
        !/projection-incomplete|typing-mutation: engine pass failed/.test(l),
    );
    if (noisy.length) report.fail('ללא רעש', noisy.join(' | '));
    else report.pass('ללא רעש');
  });
} finally {
  app.close();
}

report.print();
process.exit(process.exitCode ?? 0);
