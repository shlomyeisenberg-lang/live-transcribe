/**
 * שער QA: מה נצבע מחוץ למקומו כשהחלון צר.
 *
 * ## הדיווח והמדידה
 *
 * „באג — כשהחלון קטן, הטקסט בתחתית עולה”: הודעת שורת המצב נצבעה **על גבי**
 * סרגל הזום ועל „100%”. הסיבה אינה שאין מקום, אלא ש-`min-width: 0` על
 * `.statusbar-start` מכווץ את **התיבה** בלבד — טקסט ב-`white-space: nowrap`
 * בתוכה ממשיך להיצבע במלואו מחוצה לה, ומה שיושב שם הוא בקרת הזום. נמדד
 * בחלון של 560 פיקסלים: חפיפה של 28 פיקסלים, וב-420 של 168 — כלומר הסרגל
 * ו„100%” יחד.
 *
 * הסקר שנעשה בעקבות הדיווח מצא בפסי המעטפת עוד מקום אחד, בסף אחר ובכשל אחר:
 * מתחת ל-376 פיקסלים **הגריד של פס הכותרת עצמו** רחב מהפס (160+160 מינימום
 * הצדדים, 32 ריפוד, 24 gap), ועמודת הסיום יושבת מחוצה לו — גלולת „טרם נשמר”
 * נחתכה ב-20 פיקסלים ב-340 וב-40 ב-320. שאר הפסים נקיים: סרגל הלשוניות של
 * הרצועה וסרגל לשוניות המסמכים גוללים, וגוף הרצועה גולל אופקית בכוונה.
 *
 * ## למה זה לא יכול להיות בדיקת רכיב
 *
 * `tests/component/shell-bars.test.ts` בודק מה התבנית פולטת. „הטקסט נצבע על
 * הכפתור” הוא **חפיפת מלבנים אחרי פריסה**, ו-jsdom מחזיר אפסים לכל
 * `getBoundingClientRect`. כלומר הבאג המדובר עובר שם ירוק בכל צורה שלו.
 *
 * ## מה נמדד כאן
 *
 * שתי מדידות לכל רוחב, ושתיהן נדרשות מפני שכל אחת מהן תופסת כשל אחר:
 *
 *   1. **תוכן שנצבע על קבוצת הסיום.** כל תיבת עלה בקבוצת ההתחלה נחתכת קודם
 *      עם כל אב שאינו `overflow: visible` — תיבה שנגללה מתחת לפס גלילה
 *      מדווחת מיקום שאינו נצבע, וללא החיתוך הזה כל לשונית שנגללה הייתה
 *      נספרת כשבורה. מה ששרד נבדק מול תיבת קבוצת הסיום.
 *   2. **קבוצה שיושבת מחוץ לפס.** הילדים הישירים בלבד, מול תיבת התוכן של
 *      הפס. זה הכשל של פס הכותרת, ולא היה נתפס במדידה הראשונה: העמודה כולה
 *      מחוץ למקום, ולא תוכן שגלש מתוכה.
 *
 * ובקרה הפוכה, בלעדיה השער היה ירוק גם על הודעה שנעלמה: ההודעה חייבת להיות
 * **חתוכה** בחלון הצר (scrollWidth > clientWidth) ולשאת את הטקסט המלא
 * בטולטיפ — כלומר קוצרה, ולא הוסתרה.
 *
 * יציאה 9640 בלבד — שערים אחרים רצים במקביל על יציאות אחרות.
 *
 *   npm run build && node scripts/qa/narrow-window-qa.mjs
 */
import { writeFileSync, mkdirSync } from 'node:fs';
import { join } from 'node:path';
import { openApp, createReport, sleep, ROOT } from './harness.mjs';

const TMP = join(ROOT, 'tmp');

/**
 * הרוחבים שנמדדים.
 *
 * 1400 הוא הבקרה (חלון רגיל), 560 הוא הרוחב שבו הדיווח צולם, ו-320 הוא הקצה
 * שמתחתיו אין ממשק לדבר עליו. שלושת הביניים הם המקום שבו כל אחד משני הכשלים
 * התחיל: 480 ו-420 להודעה, 360 ו-340 לגריד של פס הכותרת.
 */
const WIDTHS = [1400, 900, 700, 560, 480, 420, 380, 360, 340, 320];

/** הרוחב שבו נבדקת הבקרה ההפוכה — ההודעה קוצרה, ולא נעלמה. */
const TRUNCATION_WIDTH = 560;

/** הפסים שנמדדים. כל אחד בנוי „קבוצת התחלה … קבוצת סיום”. */
const BARS = ['.word-statusbar', '.word-titlebar', '.word-tab-bar'];

const report = createReport('פסי המעטפת בחלון צר', { strict: true });

/**
 * המדידה בדף. מותקנת פעם אחת ונקראת לכל רוחב — כל קריאה חוזרת דרך CDP הייתה
 * שולחת את כל הגוף שלה מחדש.
 */
const INSTALL = `window.__narrowMeasure = (function () {
  /*
   * התיבה שבאמת נצבעת: הצטלבות עם כל אב שחותך, עד לשורש.
   *
   * ולא עד לפס: הקבוצה עצמה היא לרוב החותכת (word-tab-strip גולל,
   * statusbar-start חותך), ועצירה מתחתיה מדווחת כשבורה כל לשונית שפשוט
   * נגללה. זו לא הערה תיאורטית — הגרסה הראשונה של השער נכתבה כך ודיווחה
   * שבע שורות שבורות על ממשק שנמדד נקי.
   */
  function painted(el) {
    var r = el.getBoundingClientRect();
    var box = { left: r.left, right: r.right };
    var p = el.parentElement;
    while (p) {
      var cs = getComputedStyle(p);
      if (cs.overflowX !== 'visible' || cs.overflowY !== 'visible') {
        var pr = p.getBoundingClientRect();
        box.left = Math.max(box.left, pr.left);
        box.right = Math.min(box.right, pr.right);
      }
      p = p.parentElement;
    }
    return box;
  }
  function leaves(root) {
    var out = [];
    var all = root.querySelectorAll('*');
    for (var i = 0; i < all.length; i++) {
      var el = all[i];
      if (el.children.length) continue;
      var r = el.getBoundingClientRect();
      if (r.width < 1 || r.height < 1) continue;
      out.push(el);
    }
    return out;
  }
  function name(el) {
    var cls = String(el.className || '').trim();
    return (cls ? '.' + cls.split(/\\s+/)[0] : el.tagName.toLowerCase()) +
      ' „' + (el.textContent || '').trim().slice(0, 34) + '”';
  }
  return function (selectors) {
    var out = [];
    for (var s = 0; s < selectors.length; s++) {
      var bar = document.querySelector(selectors[s]);
      if (!bar) { out.push({ sel: selectors[s], missing: true }); continue; }
      var kids = bar.children;
      if (kids.length < 2) { out.push({ sel: selectors[s], missing: true }); continue; }
      var startGroup = kids[0];
      var endBox = kids[kids.length - 1].getBoundingClientRect();
      var covered = [];
      var startLeaves = leaves(startGroup);
      for (var i = 0; i < startLeaves.length; i++) {
        var b = painted(startLeaves[i]);
        if (b.right - b.left < 1) continue;
        var ov = Math.min(b.right, endBox.right) - Math.max(b.left, endBox.left);
        if (ov > 1) covered.push({ what: name(startLeaves[i]), px: Math.round(ov) });
      }
      /*
       * הכשל השני, ואינו אותו כשל: לא תוכן שגלש מקבוצה, אלא **הקבוצה עצמה**
       * שיושבת מחוץ לפס. זה מה שקורה כשמינימומי העמודות של גריד גדולים
       * מהחלון. נמדדות העמודות (הילדים הישירים) ולא העלים, כי עלה בתוך מיכל
       * גולל אמור לצאת מהמיכל — זה מה שגלילה היא.
       *
       * ואי אפשר למדוד את זה מול גבולות החלון: ל-body יש overflow: hidden,
       * ולכן הכול נחתך לתוך המסך ושום דבר לעולם אינו „מחוץ לחלון”. נמדד:
       * גלולת „טרם נשמר” שנחתכה 20 פיקסלים ב-340 דווחה כתקינה.
       */
      var barRect = bar.getBoundingClientRect();
      var barStyle = getComputedStyle(bar);
      var inner = {
        left: barRect.left + parseFloat(barStyle.paddingLeft || '0'),
        right: barRect.right - parseFloat(barStyle.paddingRight || '0'),
      };
      var outside = [];
      for (var k = 0; k < kids.length; k++) {
        var kr = kids[k].getBoundingClientRect();
        if (kr.width < 1) continue;
        var off = Math.max(0, inner.left - kr.left, kr.right - inner.right);
        if (off > 1.5) outside.push({ what: name(kids[k]), px: Math.round(off) });
      }
      out.push({ sel: selectors[s], covered: covered, outside: outside });
    }
    var msg = document.querySelector('.status-message');
    return {
      W: innerWidth,
      bars: out,
      message: msg
        ? {
            text: (msg.textContent || '').trim(),
            tip: msg.getAttribute('data-tip-title') || '',
            clipped: msg.scrollWidth > msg.clientWidth + 1,
          }
        : null,
    };
  };
})()`;

const app = await openApp({ name: 'narrow-window', port: Number(process.env.QA_PORT ?? 9640), extra: '' });

async function shoot(name) {
  mkdirSync(TMP, { recursive: true });
  const shot = await app.cdp.send('Page.captureScreenshot', { format: 'png' });
  const data = shot?.result?.data;
  if (!data) return;
  writeFileSync(join(TMP, `${name}.png`), Buffer.from(data, 'base64'));
}

try {
  await app.js(INSTALL);

  /*
   * הודעה ארוכה אמיתית, מהמסלול שמייצר אותה — ולא טקסט שהוזרק לאלמנט. „חיפוש
   * באוצריא” בלי טקסט מסומן עונה בהודעה בת 52 תווים, וזו אותה מחלקה של הודעות
   * שהדיווח צולם עליה („שוחזרו שינויים שלא נשמרו מההפעלה הקודמת…”).
   */
  await app.tab('✦ אוצריא');
  await sleep(400);
  const asked = await app.click('חיפוש באוצריא', { after: 700 });
  const status = await app.status();
  await app.tab('בית');
  await sleep(400);

  if (!asked || !status?.text) {
    report.stuck('הודעה ארוכה בשורת המצב', 'לא הצלחנו לייצר הודעה — אין מה למדוד');
  } else {
    report.pass('הודעה ארוכה בשורת המצב', `${status.text.length} תווים`);
  }
  const full = status?.text ?? '';

  let coveredRows = 0;
  let outsideRows = 0;
  for (const width of WIDTHS) {
    await app.cdp.send('Emulation.setDeviceMetricsOverride', {
      width,
      height: 620,
      deviceScaleFactor: 1,
      mobile: false,
    });
    await sleep(500);
    const data = JSON.parse(await app.js(`JSON.stringify(window.__narrowMeasure(${JSON.stringify(BARS)}))`));

    const covered = [];
    const outside = [];
    for (const bar of data.bars) {
      if (bar.missing) continue;
      for (const c of bar.covered) covered.push(`${bar.sel}: ${c.what} מכסה ${c.px}px`);
      for (const o of bar.outside) outside.push(`${bar.sel}: ${o.what} ‏${o.px}px מחוץ לפס`);
    }

    if (covered.length) {
      coveredRows++;
      report.fail(`${width}px — תוכן נצבע על קבוצת הסיום`, covered.join(' | '));
    } else {
      report.pass(`${width}px — קבוצת הסיום נקייה`);
    }
    if (outside.length) {
      outsideRows++;
      report.fail(`${width}px — קבוצה יושבת מחוץ לפס`, outside.join(' | '));
    } else {
      report.pass(`${width}px — כל הקבוצות בתוך הפס`);
    }

    if (width === TRUNCATION_WIDTH) {
      await shoot('narrow-window-560');
      const message = data.message;
      if (!message) {
        report.fail('ההודעה קוצרה ולא הוסתרה', 'אין אלמנט הודעה בשורת המצב');
      } else if (!message.clipped) {
        report.fail('ההודעה קוצרה ולא הוסתרה', 'ההודעה אינה חתוכה — ייתכן שהיא נעלמה או שהמדידה אינה על הרוחב שדווח');
      } else if (message.tip !== full || message.text !== full) {
        report.fail(
          'הטקסט המלא נשאר נגיש',
          `טולטיפ „${message.tip.slice(0, 30)}…” מול הודעה „${full.slice(0, 30)}…”`,
        );
      } else {
        report.pass('ההודעה קוצרה ולא הוסתרה', 'חתוכה, והטקסט המלא בטולטיפ');
      }
    }
  }

  if (!coveredRows && !outsideRows) {
    report.pass('סיכום', `${WIDTHS.length} רוחבים, ${BARS.length} פסים — אין ציור מחוץ למקום`);
  }
} catch (error) {
  report.stuck('הריצה', error instanceof Error ? error.message : String(error));
} finally {
  app.close();
}

report.print();
