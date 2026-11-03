/**
 * שער QA לדיאלוגים „גופן מתקדם” ו„פסקה”: הפוטר שנשאר למטה, והגרירה בכותרת.
 *
 * ## למה זה לא יכול להיות בדיקת רכיב
 *
 * שתי ההבטחות כאן הן **פריסה**, ו-jsdom מחזיר אפס מכל `getBoundingClientRect`:
 *
 *   1. **„אישור” ו„ביטול” נראים תמיד.** שני הדיאלוגים היו `overflow-block:
 *      auto` על השורש — כלומר כל הדיאלוג היה אזור הגלילה, והפוטר היה השורה
 *      האחרונה בו. בחלון שגובהו 600px התקרה `calc(100vh - 200px)` היא 400px,
 *      והתוכן גבוה ממנה: נמדד לפני התיקון שתחתית שורת הכפתורים ב-896, כמעט
 *      300px מתחת לקצה המסך. מה שנמדד עכשיו: השורש **אינו** אזור גלילה, הגוף
 *      כן, והפוטר בתוך החלון — גם אחרי גלילה עד סוף הגוף.
 *   2. **הדיאלוג נגרר.** לחיצה אמיתית בכותרת ותזוזת עכבר אמיתית דרך
 *      `Input.dispatchMouseEvent`, ואחריה מדידה של המלבן. גרירה מעבר לקצה
 *      התחתון נבלמת — אחרת הכפתורים היו יוצאים מהמסך בדרך השנייה.
 *   3. **Enter מאשר.** מקש אמיתי דרך `Input.dispatchKeyEvent`, כשהמיקוד על
 *      שורש הדיאלוג — המצב שבו הוא נפתח, ושבו Enter לא עשה כלום קודם. נמדד
 *      גם המשקל של הכפתור המסומן: „ברירת מחדל” שאינה נראית אינה ברירת מחדל.
 *
 * הוא גם מצלם: `tmp/dialog-drag-open.png` ו-`tmp/dialog-drag-moved.png`.
 *
 * יציאה 9395 בלבד — שערים אחרים רצים במקביל על יציאות אחרות.
 *
 *   npm run build && node scripts/qa/dialog-drag-qa.mjs
 */
import { writeFileSync, mkdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openApp, createReport } from './harness.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const TMP = join(ROOT, 'tmp');

/** חלון נמוך בכוונה: זה מה שמכריח את הדיאלוג להגיע לתקרת הגובה שלו. */
const WINDOW = { width: 1280, height: 600 };

/* `strict`: זה שער ולא סקר — פוטר שיצא מהמסך הוא רגרסיה, ומפיל. */
const report = createReport('גרירת דיאלוג ופוטר דביק', { strict: true });
const app = await openApp({ name: 'dialog-drag', port: Number(process.env.QA_PORT ?? 9395) });

/**
 * המדידה, בתוך הדף.
 *
 * `elementFromPoint` הוא מה שבוחר את נקודת האחיזה: מרכז הכותרת הוא לרוב מקום
 * ריק, אבל כפתור הסגירה יושב בה גם הוא — ולחיצה עליו אינה גרירה אלא סגירה.
 */
const INSTALL = `window.__dragProbe = {
  /* המחלקות נמסרות מבחוץ: אותה מדידה משרתת את „גופן מתקדם” ואת „פסקה”. */
  measure: function (sel) {
    var root = document.querySelector(sel.root);
    if (!root) return { found: false };
    var body = root.querySelector(sel.body);
    var footer = root.querySelector(sel.footer);
    var header = root.querySelector(sel.header);
    var box = function (el) {
      var r = el.getBoundingClientRect();
      return { left: Math.round(r.left), top: Math.round(r.top), right: Math.round(r.right), bottom: Math.round(r.bottom), width: Math.round(r.width), height: Math.round(r.height) };
    };
    return {
      found: true,
      root: box(root),
      body: box(body),
      footer: box(footer),
      header: box(header),
      /* השורש אינו אמור לגלול כלל: הגלילה עברה לגוף. */
      rootScrolls: root.scrollHeight > root.clientHeight + 1,
      bodyScrolls: body.scrollHeight > body.clientHeight + 1,
      bodyScrollTop: Math.round(body.scrollTop),
      viewport: { width: window.innerWidth, height: window.innerHeight },
      /* כפתור ברירת המחדל: מה ש-Enter מפעיל, ומה שההדגשה אמורה לסמן. */
      defaultLabel: (function () {
        var d = root.querySelector('[data-default-action]');
        return d ? (d.textContent || '').trim() : null;
      })(),
      defaultWeight: (function () {
        var d = root.querySelector('[data-default-action]');
        return d ? getComputedStyle(d).fontWeight : null;
      })(),
    };
  },
  scrollBodyToEnd: function (sel) {
    var body = document.querySelector(sel.root + ' ' + sel.body);
    body.scrollTop = body.scrollHeight;
    return true;
  },
  grabPoint: function (sel) {
    var header = document.querySelector(sel.root + ' ' + sel.header);
    if (!header) return null;
    var r = header.getBoundingClientRect();
    var y = Math.round(r.top + r.height / 2);
    var fractions = [0.5, 0.35, 0.65, 0.25, 0.75];
    for (var i = 0; i < fractions.length; i++) {
      var x = Math.round(r.left + r.width * fractions[i]);
      var hit = document.elementFromPoint(x, y);
      if (hit && !hit.closest('button')) return { x: x, y: y };
    }
    return null;
  }
};
true`;

/** תצלום ל-`tmp/`. */
async function shoot(name) {
  mkdirSync(TMP, { recursive: true });
  const shot = await app.cdp.send('Page.captureScreenshot', { format: 'png' });
  const data = shot?.result?.data;
  if (!data) {
    console.error(`לא ניתן לצלם (${name})`);
    return;
  }
  writeFileSync(join(TMP, `${name}.png`), Buffer.from(data, 'base64'));
  console.log(`📸 tmp/${name}.png`);
}

/** הדיאלוגים שהפוטר שלהם עבר לקיבוע, ומחלקות החלקים שלהם. */
const FONT_ADVANCED = {
  title: 'גופן מתקדם',
  open: 'מתקדם',
  root: '.fontadv-dialog',
  body: '.fa-body',
  footer: '.fa-footer',
  header: '.fa-header',
};
const PARAGRAPH = {
  title: 'פסקה',
  open: 'תפריט פסקה',
  root: '.para-dialog',
  body: '.pd-body',
  footer: '.pd-footer',
  header: '.pd-header',
};

const measure = (sel) =>
  app.js(`JSON.stringify(window.__dragProbe.measure(${JSON.stringify(sel)}))`).then(JSON.parse);

/**
 * גרירה אמיתית: לחיצה, שלוש תזוזות ושחרור.
 *
 * התזוזות ביניים אינן קישוט — גרירה בקפיצה אחת עוברת גם על מימוש שמאזין
 * ל-`mousemove` של הכותרת בלבד, ולכן אינה מבדילה בינו לבין מאזין על החלון.
 */
async function drag(from, dx, dy) {
  const send = (type, x, y, buttons) =>
    app.cdp.send('Input.dispatchMouseEvent', {
      type,
      x: Math.round(x),
      y: Math.round(y),
      button: buttons ? 'left' : 'none',
      buttons,
      clickCount: type === 'mousePressed' || type === 'mouseReleased' ? 1 : 0,
    });

  await send('mouseMoved', from.x, from.y, 0);
  await send('mousePressed', from.x, from.y, 1);
  for (const step of [0.34, 0.67, 1]) {
    await send('mouseMoved', from.x + dx * step, from.y + dy * step, 1);
    await app.sleep(30);
  }
  await send('mouseReleased', from.x + dx, from.y + dy, 1);
  await app.sleep(120);
}

/** „אישור” ו„ביטול” על המסך: זה כל מה שהתיקון הבטיח. */
const footerVisible = (m) =>
  m.footer.bottom <= m.viewport.height + 1 && m.footer.top >= 0 && m.footer.height > 0;

try {
  // חלון נמוך: התקרה `calc(100vh - 200px)` היא 400px, והדיאלוג גבוה ממנה.
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: WINDOW.width,
    height: WINDOW.height,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await app.sleep(400);

  await app.tab('בית');
  await app.js(INSTALL);
  // „פסקה” קורא את עיצוב הפסקה לפני שהוא נפתח, ובלי סמן הקריאה נכשלת.
  await app.caret(0);

  /*
   * וטקסט **מסומן**, לא רק סמן.
   *
   * „גופן מתקדם” נועל את „אישור” כשאין בחירה, ואומר למה — העיצוב חל על מה
   * שמסומן, ו-`format.apply` מקבל SelectionTarget בלבד (engine/font-advanced.ts).
   * בלי הבחירה כאן, צעד 5 („Enter מאשר”) היה נמדד על כפתור שנעול מסיבה שאין
   * לה שום קשר ל-Enter — כלומר שער אדום שאינו מצביע על מה שנשבר. נמדד: זה
   * בדיוק מה שקרה.
   *
   * המסמך שנפתח בשער ריק, ולכן צריך גם מה לסמן. ארבעה תווים אינם משנים אף
   * מדידת גאומטריה שלמטה — כולן על הדיאלוג, לא על העמוד.
   */
  await app.type('abcd', 30);
  await app.press('Home', 'Home', 36);
  await app.sleep(120);
  for (let i = 0; i < 4; i += 1) {
    await app.press('ArrowRight', 'ArrowRight', 39, 8);
    await app.sleep(30);
  }
  await app.sleep(500);

  /**
   * שלוש המדידות של הפוטר, לדיאלוג אחד.
   *
   * מוחזר המלבן של הפתיחה — הגרירה ממשיכה ממנו.
   */
  async function checkFooter(dialog) {
    if (!(await app.click(dialog.open))) {
      report.fail(`${dialog.title} — נפתח`, `הכפתור „${dialog.open}” לא נמצא ברצועה`);
      return null;
    }
    await app.sleep(300);

    const opened = await measure(dialog);
    if (!opened.found) {
      report.fail(`${dialog.title} — נפתח`, `הלחיצה לא פתחה את ${dialog.root}`);
      return null;
    }
    report.pass(`${dialog.title} — נפתח`, `${opened.root.width}×${opened.root.height} בחלון ${opened.viewport.height}px`);

    if (!opened.bodyScrolls) {
      // בלי גלישה אין מה למדוד: הדיאלוג נכנס במלואו, והפוטר נראה בכל מקרה.
      report.skip(`${dialog.title} — הגוף הוא אזור הגלילה`, 'התוכן נכנס בלי גלילה');
    } else if (opened.rootScrolls) {
      report.fail(`${dialog.title} — הגוף הוא אזור הגלילה`, 'השורש עצמו גולל — הפוטר בתוך אזור הגלילה');
    } else {
      report.pass(`${dialog.title} — הגוף הוא אזור הגלילה`, `גוף ${opened.body.height}px, שורש אינו גולל`);
    }

    if (!footerVisible(opened)) {
      report.fail(`${dialog.title} — הפוטר נראה`, `תחתית הפוטר ב-${opened.footer.bottom} מול חלון ${opened.viewport.height}`);
    } else {
      report.pass(`${dialog.title} — הפוטר נראה`, `תחתית ב-${opened.footer.bottom} מתוך ${opened.viewport.height}`);
    }

    await app.js(`window.__dragProbe.scrollBodyToEnd(${JSON.stringify(dialog)})`);
    await app.sleep(120);
    const scrolled = await measure(dialog);
    if (!scrolled.bodyScrolls) {
      report.skip(`${dialog.title} — הפוטר קבוע בגלילה`, 'אין גלילה בגוף');
    } else if (scrolled.footer.top !== opened.footer.top || !footerVisible(scrolled)) {
      report.fail(
        `${dialog.title} — הפוטר קבוע בגלילה`,
        `זז מ-${opened.footer.top} ל-${scrolled.footer.top} (גלילה ${scrolled.bodyScrollTop}px)`,
      );
    } else {
      report.pass(`${dialog.title} — הפוטר קבוע בגלילה`, `גלילה ${scrolled.bodyScrollTop}px, הפוטר ב-${scrolled.footer.top}`);
    }
    return scrolled;
  }

  /* -------------------------------------------------------------- */
  /* 1-2 — הפוטר בשני הדיאלוגים שהשורש שלהם היה אזור הגלילה           */
  /* -------------------------------------------------------------- */
  const advanced = await checkFooter(FONT_ADVANCED);
  if (!advanced) throw new Error('אין דיאלוג — אין מה למדוד');
  await shoot('dialog-drag-open');

  /* -------------------------------------------------------------- */
  /* 3 — גרירה בכותרת                                                */
  /* -------------------------------------------------------------- */
  const grab = JSON.parse(await app.js(`JSON.stringify(window.__dragProbe.grabPoint(${JSON.stringify(FONT_ADVANCED)}))`));
  if (!grab) {
    report.fail('גרירה בכותרת', 'לא נמצאה נקודת אחיזה פנויה בכותרת');
  } else {
    const before = await measure(FONT_ADVANCED);
    /*
     * התזוזה נגזרת מהמקום שיש בפועל, ולא ממספר קבוע: הדיאלוג נפתח בעברית
     * בקצה הימני של החלון ובגובה 400px מתוך 600, ולכן גרירה קבועה של „ימינה
     * ולמטה” נבלמת מיד — ואז השער מדד את הבלימה וקרא לה כשל. הבלימה עצמה
     * נמדדת בצעד 4, בכוונה ובנפרד.
     */
    const dx = -Math.min(200, before.root.left);
    const dy = Math.min(60, before.viewport.height - before.root.bottom);
    await drag(grab, dx, dy);
    const after = await measure(FONT_ADVANCED);
    const moved = { x: after.root.left - before.root.left, y: after.root.top - before.root.top };
    if (moved.x === dx && moved.y === dy) {
      report.pass('גרירה בכותרת', `זז (${moved.x}, ${moved.y}) בדיוק כמו העכבר`);
    } else if (moved.x === 0 && moved.y === 0) {
      report.fail('גרירה בכותרת', 'הדיאלוג לא זז כלל');
    } else {
      report.fail('גרירה בכותרת', `זז (${moved.x}, ${moved.y}) במקום (${dx}, ${dy})`);
    }
    await shoot('dialog-drag-moved');
  }

  /* -------------------------------------------------------------- */
  /* 4 — הקצה התחתון בולם, כדי שהכפתורים לא ייצאו מהמסך               */
  /* -------------------------------------------------------------- */
  const grabAgain = JSON.parse(await app.js(`JSON.stringify(window.__dragProbe.grabPoint(${JSON.stringify(FONT_ADVANCED)}))`));
  if (!grabAgain) {
    report.fail('הקצה התחתון בולם', 'לא נמצאה נקודת אחיזה');
  } else {
    await drag(grabAgain, 0, 900);
    const pushed = await measure(FONT_ADVANCED);
    if (pushed.root.bottom <= pushed.viewport.height + 1 && footerVisible(pushed)) {
      report.pass('הקצה התחתון בולם', `תחתית הדיאלוג ב-${pushed.root.bottom} מתוך ${pushed.viewport.height}`);
    } else {
      report.fail('הקצה התחתון בולם', `תחתית ב-${pushed.root.bottom} מול חלון ${pushed.viewport.height}`);
    }
  }

  /* -------------------------------------------------------------- */
  /* 5 — „אישור” הוא ברירת המחדל: מודגש, ו-Enter מפעיל אותו           */
  /* -------------------------------------------------------------- */
  const marked = await measure(FONT_ADVANCED);
  if (marked.defaultLabel !== 'אישור') {
    report.fail('„אישור” מסומן כברירת מחדל', `הכפתור המסומן הוא „${marked.defaultLabel}”`);
  } else if (Number(marked.defaultWeight) < 600) {
    report.fail('„אישור” מסומן כברירת מחדל', `משקל ${marked.defaultWeight} — אינו מודגש`);
  } else {
    report.pass('„אישור” מסומן כברירת מחדל', `משקל ${marked.defaultWeight}`);
  }

  /*
   * Enter מהמקלדת, ולא `click()` על הכפתור: המיקוד יושב על שורש הדיאלוג —
   * שם הוא נוחת בפתיחה — וזו בדיוק הנקודה שבה Enter לא עשה כלום קודם.
   * ההוכחה שהוא פעל: הדיאלוג נסגר, כפי ש-`onFontAdvancedSubmit` סוגר אותו.
   *
   * שדה אחד מלא לפני כן, ובכוונה: „אישור” נעול על דיאלוג שאין בו מה להחיל
   * (כפתור שנלחץ, סוגר, ולא עושה דבר הוא „ביטול” בתחפושת), ולכן Enter על
   * דיאלוג ריק **אמור** לא לעשות דבר. מה שנמדד כאן הוא שה-Enter מגיע
   * לכפתור — לא שהוא מגיע גם כשאין לו מה לעשות.
   */
  const filled = await app.dialogFill('fa-scale', '120');
  if (filled !== 'ok') report.fail('Enter מאשר', `מילוי „מתיחה אופקית” החזיר ${filled}`);
  await app.sleep(150);
  await app.press('Enter', 'Enter', 13);
  await app.sleep(400);
  const afterEnter = await measure(FONT_ADVANCED);
  if (!afterEnter.found) {
    report.pass('Enter מאשר', 'הדיאלוג נסגר, כמו לחיצה על „אישור”');
  } else {
    /*
     * הדיאלוג נסגר על **הצלחה בלבד** (ראו HomeTab.onFontAdvancedSubmit), ולכן
     * „נשאר פתוח” מערבב שתי טענות שונות: „ה-Enter לא הגיע לכפתור” ו„ההחלה
     * נכשלה”. שורת המצב היא מה שמפריד ביניהן, ובלעדיה השער היה מאשים את
     * ה-Enter בכשלון שאין לו שום קשר אליו — אותה טעות ייחוס שהבחירה שנוספה
     * בראש הקובץ נועדה למנוע.
     */
    const status = await app.status();
    if (status && status.error && status.text) {
      report.fail('Enter מאשר', `ההחלה נכשלה, ולא ה-Enter: „${status.text}”`);
    } else {
      report.fail('Enter מאשר', 'הדיאלוג נשאר פתוח, ושורת המצב שקטה — ה-Enter לא הגיע לכפתור');
    }
  }

  /* -------------------------------------------------------------- */
  /* 6 — „פסקה”: אותו תיקון, ולכן אותה מדידה                          */
  /* -------------------------------------------------------------- */
  await app.press('Escape', 'Escape', 27);
  await app.sleep(200);
  await checkFooter(PARAGRAPH);

  const noise = (await app.log()).filter((line) => !line.startsWith('warn:'));
  if (noise.length) console.log('לוג הדף:', noise.slice(0, 5).join(' | '));
} finally {
  report.print();
  app.close();
}
