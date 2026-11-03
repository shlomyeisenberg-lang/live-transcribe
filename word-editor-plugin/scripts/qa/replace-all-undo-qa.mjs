/**
 * "החלף הכל" — Ctrl+Z **יחיד** מבטל את כל ההחלפות, לא N ביטולים נפרדים.
 *
 * רגרסיה שנייה שנתפסה אחרי תיקון הכיסוי (ראו replace-all-multiparagraph-qa.mjs):
 * המעבר מ-`handle.replaceAll()` (batch יחיד במנוע, `undo`-step אחד) ל-
 * `doc.replace()` נקודתי (קריאה נפרדת לכל מופע) פתר את הכיסוי אבל פתח N
 * שלבי ביטול נפרדים ל-N מופעים — נמדד בפועל: "החלף הכל" על 3 מופעים דרש
 * 3 Ctrl+Z נפרדים.
 *
 * הפתרון: `doc.mutations.apply({atomic:true, steps:[...text.rewrite]})` —
 * ראו התיעוד המלא ב-`engine/search.ts` (`applyAtomicRewriteChunks`). השער
 * הזה בודק בדיוק את מה שנמדד שבור: מסמך עם 8 מופעים על פני 8 פסקאות,
 * "החלף הכל" דרך הדיאלוג האמיתי, Ctrl+Z **אחד**, וכל השמונה חוזרים.
 */
import { openApp, createReport } from './harness.mjs';

const report = createReport('החלף הכל — undo יחיד', { strict: true });
const app = await openApp({ name: 'replundo', port: Number(process.env.QA_PORT ?? 9504) });

const NEEDLE = 'zzq';
const COUNT = 8;
const count = (s, n) => s.split(n).length - 1;

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1600,
    height: 1000,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await app.sleep(400);

  await app.caret(0);
  for (let i = 0; i < COUNT; i++) {
    await app.type(`${NEEDLE} paragraph ${i}`);
    if (i < COUNT - 1) {
      await app.press('Enter', 'Enter', 13);
      await app.sleep(150);
    }
  }
  await app.sleep(1200);

  let files = await app.docx();
  let doc = files['word/document.xml'] || '';
  const before = count(doc, NEEDLE);
  console.log(`מופעים לפני ב-docx: ${before}`);
  if (before !== COUNT) report.fail('הכנת הבדיקה', `נמצאו ${before} מופעים במקום ${COUNT}`);

  await app.tab('בית');
  const opened = await app.click('החלפה');
  if (!opened) throw new Error('כפתור "החלפה" לא נמצא ברצועה');
  await app.sleep(600);

  const fill = async (sel, value) => app.js(`(function(){
    var el = document.querySelector('${sel}');
    if (!el) return 'no-el';
    var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    setter.call(el, ${JSON.stringify(value)});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return 'ok';
  })()`);

  console.log('שדה חיפוש:', await fill('#fr-search-input', NEEDLE));
  console.log('שדה החלפה:', await fill('#fr-replace-input', 'YYY'));
  await app.sleep(900);

  const btns = JSON.parse(
    await app.js(
      `JSON.stringify(Array.from(document.querySelectorAll('.find-replace-dialog .fr-btn')).map(x=>x.textContent.trim()))`,
    ),
  );
  const idx = btns.indexOf('החלף הכל');
  if (idx < 0) throw new Error(`כפתור "החלף הכל" לא נמצא. כפתורים: ${JSON.stringify(btns)}`);
  const rect = JSON.parse(
    await app.js(
      `(function(){var el=document.querySelectorAll('.find-replace-dialog .fr-btn')[${idx}];var r=el.getBoundingClientRect();return JSON.stringify({x:r.x+r.width/2,y:r.y+r.height/2});})()`,
    ),
  );
  await app.clickAt(rect.x, rect.y);
  await app.sleep(1200);

  files = await app.docx();
  doc = files['word/document.xml'] || '';
  const afterReplace = { left: count(doc, NEEDLE), put: count(doc, 'YYY') };
  console.log(`אחרי "החלף הכל": נשארו ${afterReplace.left} ${NEEDLE}, נכתבו ${afterReplace.put} YYY`);
  if (afterReplace.left === 0 && afterReplace.put === COUNT) {
    report.pass('החלף הכל בוצע', `${afterReplace.put}/${COUNT}`);
  } else {
    report.fail('החלף הכל לא בוצע במלואו', `נשארו ${afterReplace.left}, הוחלפו ${afterReplace.put}`);
  }

  // סגירת הדיאלוג לפני Ctrl+Z: אחרת המיקוד בשדה הטקסט של הדיאלוג, ו-Ctrl+Z
  // עלול לפעול על תיבת הטקסט של הדיאלוג ולא על מסמך העריכה.
  await app.escape();
  await app.sleep(300);

  // Ctrl+Z **יחיד**. modifiers=2 הוא הביט של Ctrl ב-CDP Input.dispatchKeyEvent.
  await app.press('z', 'KeyZ', 90, 2);
  await app.sleep(1000);

  files = await app.docx();
  doc = files['word/document.xml'] || '';
  const afterOneUndo = { left: count(doc, NEEDLE), put: count(doc, 'YYY') };
  console.log(`אחרי Ctrl+Z יחיד: ${afterOneUndo.left} ${NEEDLE}, ${afterOneUndo.put} YYY`);

  if (afterOneUndo.left === COUNT && afterOneUndo.put === 0) {
    report.pass('Ctrl+Z יחיד מבטל את כל השמונה', 'התאוששות מלאה ב-undo-step אחד');
  } else if (afterOneUndo.left === COUNT - 1 && afterOneUndo.put === 1) {
    report.fail(
      'Ctrl+Z יחיד ביטל רק מופע אחד',
      `${afterOneUndo.left} ${NEEDLE} נשארו, ${afterOneUndo.put} YYY נשארו — צריך N ביטולים, לא אחד`,
    );
  } else {
    report.fail('Ctrl+Z — תוצאה לא צפויה', `${afterOneUndo.left} ${NEEDLE}, ${afterOneUndo.put} YYY`);
  }

  const screenAfterUndo = (await app.screenText()) || '';
  console.log('על המסך אחרי הביטול:', JSON.stringify(screenAfterUndo.slice(0, 200)));
  if (count(screenAfterUndo, NEEDLE) === COUNT) {
    report.pass('התצוגה תואמת לביטול', `${NEEDLE} מופיע ${COUNT} פעמים על המסך`);
  } else {
    report.fail('התצוגה לא תואמת לביטול', `${NEEDLE} מופיע ${count(screenAfterUndo, NEEDLE)} פעמים על המסך`);
  }

  console.log('לוג הדף:', JSON.stringify(await app.log()).slice(0, 500));
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
