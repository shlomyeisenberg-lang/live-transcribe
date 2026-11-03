/**
 * שער QA לפריסת רשימת הגופנים — שתי העמודות, ופס הדגימה.
 *
 * ## למה זה לא יכול להיות בדיקת רכיב
 *
 * `tests/component/font-combo.test.ts` מודד את מה שהתבנית מכריעה: אילו שורות
 * מקבלות דגימה, ואיזו מחלקה יושבת על הרשימה. שתי ההבטחות של העיצוב עצמו
 * **אינן נבדקות שם בכלל**, מפני ש-jsdom אינו מחשב פריסה:
 *
 * 1. **השמות נצמדים לקצה הסיום.** `justify-content: space-between` עם
 *    `::before` שקיים גם כשהוא ריק — הפרש אורך השמות נופל באמצע השורה, ולא
 *    מצטבר בקצה. שורה שאיבדה את זה נראית „תקינה לחלוטין” בכל סריקת מקור.
 * 2. **הפנגרם נכנס בפס.** הפס מחזיק שלוש שורות של 227 פיקסלים (הרוחב נמדד כאן,
 *    והוא 260 של הרשימה פחות מסגרת, ריפוד ופס גלילה), והפסוק שבברירת המחדל
 *    (`composables/font-sample.ts`) הוא 63 תווים. „נכנס” כאן פירושו
 *    `scrollHeight <= clientHeight` בגופן שהסימון עומד עליו — וזה בדיוק
 *    המספר שאין לו קיום ב-jsdom.
 *
 * ומעבר למדידה: הוא מצלם. `tmp/font-list-layout.png` הוא הרשימה כפי שהיא
 * נפתחת, ו-`tmp/font-list-walked.png` היא אחרי סיור בחצים — כלומר הפס בגופן
 * אחר. שניהם מה שמאפשר להסתכל על הרשימה במקום להאמין לתיאור שלה.
 *
 * יציאה 9393 בלבד — שערים אחרים רצים במקביל על יציאות אחרות.
 *
 *   npm run build && node scripts/qa/font-list-layout-qa.mjs
 */
import { writeFileSync, mkdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openApp, createReport } from './harness.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const TMP = join(ROOT, 'tmp');

/** כמה גופנים לעבור עליהם בחצים — כל אחד מצייר את הפס בגופן אחר. */
const FONTS_TO_WALK = 10;

/* `strict`: זה שער ולא סקר — שורה שבורה כאן היא רגרסיה בפריסה, ומפילה. */
const report = createReport('פריסת רשימת הגופנים', { strict: true });
const app = await openApp({ name: 'font-list', port: Number(process.env.QA_PORT ?? 9393) });

/**
 * פתיחת הרשימה ומדידתה, בבת אחת.
 *
 * `input` ואחריו אירוע, ולא `focus()` לבד: ב-headless החלון אינו ממוקד ואירוע
 * ה-focus אינו בהכרח נורה. זו אותה פתיחה שב-`installed-fonts-qa.mjs`, ומאותה
 * סיבה.
 *
 * המדידה עצמה נמנעת משורות שה-`content-visibility` דילג עליהן: שורה שלא
 * צוירה מדווחת על גובה תחליפי, וקצוות הטקסט שלה אינם אומרים דבר. לכן נמדדות
 * רק שורות שנחתכות עם תיבת הרשימה.
 */
const MEASURE = `(async function () {
  var input = window.__qa.el('גופן', { scope: '.word-ribbon-body', selector: 'input[role="combobox"]' });
  if (!input) return JSON.stringify({ found: false, why: 'אין פקד' });
  input.focus();
  var setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
  setter.call(input, '');
  input.dispatchEvent(new Event('input', { bubbles: true }));
  await null;
  await null;
  await null;
  var list = document.getElementById(input.getAttribute('aria-controls'));
  if (!list) return JSON.stringify({ found: false, why: 'הרשימה לא נפתחה' });
  return JSON.stringify(window.__fontListMeasure(list));
})()`;

/**
 * הפונקציה שמותקנת בדף פעם אחת ונקראת שוב אחרי כל חץ. היא בדף ולא כאן מפני
 * שכל קריאה חוזרת דרך CDP הייתה שולחת את כל הגוף שלה מחדש.
 */
const INSTALL = `window.__fontListMeasure = function (list) {
  var lr = list.getBoundingClientRect();
  var rtl = getComputedStyle(list).direction === 'rtl';
  var rows = [];
  var options = list.querySelectorAll('[role="option"]');
  for (var i = 0; i < options.length; i++) {
    var li = options[i];
    var r = li.getBoundingClientRect();
    /* שורה שדולגה ב-content-visibility אינה נמדדת — ראו הערת הראש. */
    if (r.bottom <= lr.top + 1 || r.top >= lr.bottom - 1 || r.height === 0) continue;
    var node = null;
    for (var k = 0; k < li.childNodes.length; k++) {
      if (li.childNodes[k].nodeType === 3 && li.childNodes[k].nodeValue.trim() !== '') node = li.childNodes[k];
    }
    if (!node) continue;
    var range = document.createRange();
    range.selectNodeContents(node);
    var t = range.getBoundingClientRect();
    rows.push({
      value: li.getAttribute('data-value'),
      hebrew: li.classList.contains('hebrew'),
      /* המרחק מקצה הסיום של השורה אל קצה הסיום של השם. */
      endGap: Math.round(rtl ? t.left - r.left : r.right - t.right),
      /* המרחק מקצה ההתחלה — זה מה שאמור להשתנות בין שם ארוך לקצר. */
      startGap: Math.round(rtl ? r.right - t.right : t.left - r.left),
      width: Math.round(t.width)
    });
  }

  var bar = list.querySelector('.ribbon-combo-sample');
  var sample = null;
  if (bar) {
    var span = bar.firstElementChild;
    var br = bar.getBoundingClientRect();
    var range2 = document.createRange();
    range2.selectNodeContents(span);
    var rects = range2.getClientRects();
    var tops = {};
    for (var j = 0; j < rects.length; j++) tops[Math.round(rects[j].top)] = 1;
    sample = {
      text: span.textContent,
      family: bar.style.fontFamily || '(ברירת מחדל)',
      /* מצב הדגימה: שלוש שורות בגודל הרשימה, ולא שורה אחת בגודל שבמסמך. */
      specimen: bar.classList.contains('specimen'),
      fontSize: getComputedStyle(bar).fontSize,
      height: Math.round(br.height),
      inner: Math.round(bar.clientWidth - 16),
      lines: Object.keys(tops).length,
      /* הטקסט נקטע: יש בו יותר משלוש שורות. זו רגרסיה — הפס חייב להיות פנגרם. */
      truncated: span.scrollHeight > span.clientHeight + 0.5,
      /* הפס עצמו גולש מגבולותיו — זו רגרסיה בפריסה, ולא קטיעה של טקסט. */
      overflow: bar.scrollHeight > bar.clientHeight + 0.5,
      scrollPad: getComputedStyle(list).scrollPaddingBottom
    };
  }
  return { found: true, rows: rows, sample: sample, listWidth: Math.round(lr.width) };
};`;

const ARROW = `(function () {
  var input = window.__qa.el('גופן', { scope: '.word-ribbon-body', selector: 'input[role="combobox"]' });
  input.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
  return Promise.resolve().then(function () {
    var list = document.getElementById(input.getAttribute('aria-controls'));
    return JSON.stringify(window.__fontListMeasure(list).sample);
  });
})()`;

/** תצלום ל-`tmp/`. שם לכל צעד — הפס מצייר בגופן אחר בכל אחד מהם. */
async function shoot(name) {
  mkdirSync(TMP, { recursive: true });
  const shot = await app.cdp.send('Page.captureScreenshot', { format: 'png' });
  const data = shot?.result?.data;
  if (!data) {
    console.error(`לא ניתן לצלם (${name})`);
    return;
  }
  const file = join(TMP, `${name}.png`);
  writeFileSync(file, Buffer.from(data, 'base64'));
  console.log(`📸 tmp/${name}.png`);
}

try {
  await app.tab('בית');
  // המנייה נורית ב-onMounted ומוותרת על החוט כל 40 שמות; היא נוחתת אחרי
  // שהמסמך הראשון כבר פתוח.
  await app.sleep(2500);
  await app.js(INSTALL);

  const measured = JSON.parse(await app.js(MEASURE));
  if (!measured.found) {
    report.fail('הרשימה נפתחה', measured.why);
    throw new Error('אין רשימה — אין מה למדוד');
  }
  console.log(`נמדדו ${measured.rows.length} שורות, רוחב הרשימה ${measured.listWidth}px`);
  await shoot('font-list-layout');

  /* -------------------------------------------------------------- */
  /* 1 — השמות נצמדים לקצה הסיום                                     */
  /* -------------------------------------------------------------- */
  const gaps = [...new Set(measured.rows.map((r) => r.endGap))];
  if (measured.rows.length < 5) {
    report.skip('השמות על קו אחד', `רק ${measured.rows.length} שורות נראות`);
  } else if (gaps.length === 1) {
    report.pass('השמות על קו אחד', `${measured.rows.length} שורות, מרחק ${gaps[0]}px מהקצה`);
  } else {
    report.fail('השמות על קו אחד', `מרחקים שונים מהקצה: ${gaps.join(', ')}`);
  }

  /* המדידה שלמעלה עוברת גם ברשימה שכל שמותיה באותו אורך. זו מפרידה. */
  const starts = [...new Set(measured.rows.map((r) => r.startGap))];
  const widths = [...new Set(measured.rows.map((r) => r.width))];
  if (widths.length < 2) {
    report.skip('הרווח באמצע', 'כל השמות באותו רוחב — אין מה להפריד');
  } else if (starts.length > 1) {
    report.pass('הרווח באמצע', `${starts.length} מרחקים שונים מקצה ההתחלה, ${widths.length} רוחבי שם`);
  } else {
    report.fail('הרווח באמצע', 'השמות נצמדים לשני הקצוות — הרווח לא נפל באמצע');
  }

  /* -------------------------------------------------------------- */
  /* 2 — פס הדגימה מחזיק את הפסוק                                    */
  /* -------------------------------------------------------------- */
  const bar = measured.sample;
  if (!bar) {
    report.fail('פס הדגימה קיים', 'אין `.ribbon-combo-sample` ברשימה פתוחה');
  } else {
    console.log(
      `הפס: ${bar.height}px גובה, ${bar.inner}px רוחב פנימי, ${bar.fontSize}, ` +
        `scroll-padding ${bar.scrollPad}`,
    );
    console.log(`הטקסט (${Array.from(bar.text).length} תווים): ${bar.text}`);

    /*
     * בלי בחירה הפס אמור להיות במצב הדגימה: הפסוק בגודל הקבוע של הרשימה.
     * אילו הגודל שבמסמך היה זורם לכאן (עד 24px) הוא לא היה נכנס בשום גופן.
     */
    if (bar.specimen && bar.fontSize === '14px') {
      report.pass('בלי בחירה הפס במצב דגימה', `${bar.fontSize}, ${bar.height}px`);
    } else {
      report.fail(
        'בלי בחירה הפס במצב דגימה',
        bar.specimen ? `מצב דגימה אבל ב-${bar.fontSize}` : `אין מחלקת specimen (${bar.fontSize})`,
      );
    }

    if (bar.scrollPad === `${bar.height}px`) {
      report.pass('הפס וה-scroll-padding שווים', `${bar.height}px`);
    } else {
      report.fail('הפס וה-scroll-padding שווים', `פס ${bar.height}px מול ${bar.scrollPad}`);
    }

    const walked = [bar];
    for (let i = 0; i < FONTS_TO_WALK; i++) {
      const next = JSON.parse(await app.js(ARROW));
      if (next) walked.push(next);
    }
    for (const w of walked) {
      console.log(`  ${w.family}: ${w.lines} שורות${w.truncated ? ' ✂' : ''}`);
    }

    /* גלישה של הפס עצמו היא רגרסיה בפריסה: גובה שהוקטן, או קטיעה שהוסרה. */
    const overflowing = walked.filter((w) => w.overflow);
    if (overflowing.length === 0) {
      report.pass('הפס אינו גולש מגבולותיו', `${walked.length} גופנים`);
    } else {
      report.fail('הפס אינו גולש מגבולותיו', overflowing.map((w) => w.family).join(', '));
    }

    /*
     * קטיעה היא שבר: הפסוק קיים כדי להראות כל אות, והסתרת זנבו מסתירה חלק
     * מהאלפבית. שלוש שורות נבחרו במיוחד כדי ש-Courier New, שנחתך בשתיים,
     * ייכנס במלואו.
     */
    const cut = walked.filter((w) => w.truncated);
    if (cut.length === 0) {
      report.pass('הפסוק נראה במלואו', `בכל ${walked.length} הגופנים שנמדדו`);
    } else {
      report.fail('הפסוק נראה במלואו', `נחתך ב: ${cut.map((w) => w.family).join(', ')}`);
    }
  }

  /* התצלום השני — אחרי הסיור, כלומר על הגופן האחרון שהסימון עבר עליו. */
  await shoot('font-list-walked');
} finally {
  report.print();
  app.close();
}
