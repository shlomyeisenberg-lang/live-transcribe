/**
 * ניווט וההחלפה הבודדת בדיאלוג האמיתי, על מסמך רב-פסקאות: המונה מתעדכן
 * נכון בין "מצא הבא"/"מצא קודם", ו-"החלף" (לא "החלף הכל") מחליף רק את
 * המופע הפעיל ומתקדם לבא בתור — לא את כולם ולא כלום.
 *
 * זה המשך ישיר ל-replace-all-multiparagraph-qa.mjs: אותו מסמך שהיה שבור
 * ב-ui.search (`projection-incomplete`), הפעם עם ניווט בין מופעים ולא רק
 * "החלף הכל".
 */
import { openApp, createReport } from './harness.mjs';

const report = createReport('חיפוש-ניווט והחלפה בודדת — מסמך רב-פסקאות', { strict: true });
const app = await openApp({ name: 'findnav', port: Number(process.env.QA_PORT ?? 9502) });

const NEEDLE = 'zzq';
const COUNT = 8;
const count = (s, n) => s.split(n).length - 1;

const counterText = async () =>
  app.js(`(function(){var e=document.querySelector('.fr-counter');return e?e.textContent.trim():null;})()`);

const fill = async (sel, value) =>
  app.js(`(function(){
    var el = document.querySelector('${sel}');
    if (!el) return 'no-el';
    var setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
    setter.call(el, ${JSON.stringify(value)});
    el.dispatchEvent(new Event('input', { bubbles: true }));
    return 'ok';
  })()`);

const clickDialogButton = async (label) => {
  const btns = JSON.parse(
    await app.js(
      `JSON.stringify(Array.from(document.querySelectorAll('.find-replace-dialog .fr-btn')).map(x=>x.textContent.trim()))`,
    ),
  );
  const idx = btns.indexOf(label);
  if (idx < 0) throw new Error(`כפתור "${label}" לא נמצא. כפתורים: ${JSON.stringify(btns)}`);
  const rect = JSON.parse(
    await app.js(
      `(function(){var el=document.querySelectorAll('.find-replace-dialog .fr-btn')[${idx}];var r=el.getBoundingClientRect();return JSON.stringify({x:r.x+r.width/2,y:r.y+r.height/2});})()`,
    ),
  );
  await app.clickAt(rect.x, rect.y);
  await app.sleep(500);
};

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
  await app.sleep(1000);

  await app.tab('בית');
  const openedFind = await app.click('חפש');
  if (!openedFind) throw new Error('כפתור "חפש" לא נמצא ברצועה');
  await app.sleep(500);
  console.log('שדה חיפוש:', await fill('#fr-search-input', NEEDLE));
  await app.sleep(900);

  const first = await counterText();
  console.log('מונה אחרי חיפוש ראשוני:', first);
  first === `1 מתוך ${COUNT}`
    ? report.pass('מונה ראשוני', first)
    : report.fail('מונה ראשוני', `ציפינו ל-"1 מתוך ${COUNT}", קיבלנו "${first}"`);

  await clickDialogButton('מצא הבא');
  const second = await counterText();
  console.log('מונה אחרי "מצא הבא":', second);
  second === `2 מתוך ${COUNT}`
    ? report.pass('מצא הבא', second)
    : report.fail('מצא הבא', `ציפינו ל-"2 מתוך ${COUNT}", קיבלנו "${second}"`);

  await clickDialogButton('מצא קודם');
  const back = await counterText();
  console.log('מונה אחרי "מצא קודם":', back);
  back === `1 מתוך ${COUNT}`
    ? report.pass('מצא קודם', back)
    : report.fail('מצא קודם', `ציפינו ל-"1 מתוך ${COUNT}", קיבלנו "${back}"`);

  // עכשיו החלפה בודדת: לעבור למצב "החלף", ולהחליף רק את המופע הפעיל (הראשון).
  await app.click('החלפה');
  await app.sleep(500);
  console.log('שדה חיפוש (שוב):', await fill('#fr-search-input', NEEDLE));
  await app.sleep(900);
  console.log('שדה החלפה:', await fill('#fr-replace-input', 'ONE'));

  let files = await app.docx();
  const beforeDoc = files['word/document.xml'] || '';
  const beforeCount = count(beforeDoc, NEEDLE);

  await clickDialogButton('החלף');
  await app.sleep(1000);

  files = await app.docx();
  const afterDoc = files['word/document.xml'] || '';
  const zzqLeft = count(afterDoc, NEEDLE);
  const oneCount = count(afterDoc, 'ONE');
  console.log(`לפני ההחלפה הבודדת: ${beforeCount} מופעי zzq. אחריה: ${zzqLeft} נשארו, ${oneCount} הוחלפו`);

  if (beforeCount === COUNT && zzqLeft === COUNT - 1 && oneCount === 1) {
    report.pass('החלף (בודד)', `בדיוק מופע אחד הוחלף, ${zzqLeft} נשארו`);
  } else {
    report.fail('החלף (בודד)', `לפני ${beforeCount}, אחרי נשארו ${zzqLeft}, הוחלפו ${oneCount}`);
  }

  const counterAfterSingle = await counterText();
  console.log('מונה אחרי החלפה בודדת:', counterAfterSingle);
  counterAfterSingle === `1 מתוך ${COUNT - 1}`
    ? report.pass('מונה אחרי החלפה בודדת', counterAfterSingle)
    : report.fail('מונה אחרי החלפה בודדת', `ציפינו ל-"1 מתוך ${COUNT - 1}", קיבלנו "${counterAfterSingle}"`);

  console.log('לוג הדף:', JSON.stringify(await app.log()).slice(0, 500));
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
