/**
 * שער ה-QA של אזכור „@”.
 *
 * הכלל היחיד, כמו בשאר השערים: `success: true` אינו הוכחה — מה שנכתב ל-OOXML
 * הוא ההוכחה. קישור ב-DOCX הוא שני דברים שחייבים להסכים: `<w:hyperlink r:id>`
 * ב-document.xml, ו-Relationship עם ה-Target ב-document.xml.rels. אחד בלי
 * השני נראה כמו הצלחה ואינו קישור.
 *
 * השער מודד גם את ארבעת הדברים שהמנוע לא נותן בקבלה, וכל אחד מהם היה באג
 * מדווח (‎#66, ‎#69, ‎#73):
 *
 *   - **שני קישורים בפסקה אחת.** `hyperlinks.insert` דוחה את השני; המסלול
 *     שרץ הוא `doc.insert` ואז `hyperlinks.wrap` (הערת המודול ב-
 *     engine/at-mention-overlay.ts).
 *   - **הסמן אחרי הקישור**, אחרת ההקלדה הבאה נדחפת לפניו.
 *   - **תיוג מחדש אחרי מחיקה**, ושלא נשאר `<w:hyperlink>` בלי `<w:t>`.
 *   - **הלחיצה פותחת**: המנוע אינו מצייר `otzaria://` כקישור, ולכן נמדד
 *     שהגשר (engine/otzaria-link-click.ts) הוא שמתרגם אותה לקריאה למאחז.
 *
 * הרצה:  node scripts/qa/at-mention-qa.mjs
 * היציאה 9610 שמורה לשער הזה בלבד.
 */
import { openApp, createReport } from './harness.mjs';

const PORT = Number(process.env.QA_PORT ?? 9610);
const report = createReport('אזכור „@”', { strict: true });

const log = (...a) => console.log(...a);

/** ההתאמות שהמאחז יחזיר במקום אוצריא. */
const HITS = [
  {
    id: 42,
    bookId: 'פסחים',
    bookUid: 'id:42',
    type: 'text',
    title: 'פסחים',
    reference: 'פסחים דף לד',
    index: 1234,
    isPdf: false,
    isSourceLine: true,
    isUserBook: false,
    bookPath: 'ש"ס, בבלי',
  },
  {
    id: 43,
    bookId: 'פסחים',
    bookUid: 'id:43',
    type: 'text',
    title: 'פסחים',
    reference: 'פסחים דף לה',
    index: 1300,
    isPdf: false,
    isSourceLine: false,
    isUserBook: false,
    bookPath: 'ש"ס, בבלי',
  },
];

/** ספר אישי: אין לו id חד-משמעי, ולכן הקישור אליו הוא „איתור מקורות”. */
const USER_HIT = {
  id: null,
  bookId: 'בסוגיא דדיורים בריבית',
  type: 'text',
  title: 'בסוגיא דדיורים בריבית',
  reference: 'בסוגיא דדיורים בריבית',
  index: 0,
  isPdf: false,
  isSourceLine: false,
  isUserBook: true,
  bookPath: 'ספרים אישיים',
};

/** הטקסט הנראה של המסמך, בלי התגיות. */
function flat(doc) {
  return doc.replace(/<[^>]+>/g, '');
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

/** תמונת מצב של המסמך. */
async function snap(app) {
  const files = (await app.docx()) ?? {};
  return {
    files,
    doc: files['word/document.xml'] ?? '',
    rels: files['word/_rels/document.xml.rels'] ?? '',
  };
}

/** מצב הרשימה הצפה, נקרא מה-DOM האמיתי. */
async function popup(app) {
  return JSON.parse(
    await app.js(`(function () {
      var el = document.querySelector('.otzaria-at-mention');
      if (!el) return JSON.stringify({ open: false });
      var rows = [].slice.call(el.querySelectorAll('[role="option"]'));
      var box = el.getBoundingClientRect();
      return JSON.stringify({
        open: true,
        count: rows.length,
        active: rows.findIndex(function (r) { return r.getAttribute('aria-selected') === 'true'; }),
        texts: rows.map(function (r) { return r.textContent; }),
        rect: { left: box.left, top: box.top, right: box.right, bottom: box.bottom },
        role: el.getAttribute('role'),
        activedescendant: el.getAttribute('aria-activedescendant'),
        inEngineTree: !!el.closest('[class*="superdoc"]'),
      });
    })()`),
  );
}

/** כמה קישורים יש במסמך. */
function countLinks(doc) {
  return (doc.match(/<w:hyperlink\b/g) ?? []).length;
}

/**
 * מנקה את המסמך בין תרחישים.
 *
 * `Meta` ולא `Ctrl`: השער רץ ב-Chrome על macOS, ושם Ctrl+A אינו „בחר הכול”.
 * ניקוי שאינו מנקה הוא הדבר הגרוע ביותר בשער — התרחיש הבא מודד את השאריות
 * של קודמו ונראה שבור בלי שיהיה.
 */
const SELECT_ALL_MODIFIER = process.platform === 'darwin' ? 4 : 2;

async function clearDoc(app) {
  await app.press('a', 'KeyA', 65, SELECT_ALL_MODIFIER);
  await app.sleep(150);
  await app.press('Backspace', 'Backspace', 8, 0);
  await app.sleep(500);
  const { doc } = await snap(app);
  const left = doc.replace(/<[^>]+>/g, '').trim();
  if (left) throw new Error(`הניקוי לא רוקן את המסמך: ${left.slice(0, 80)}`);
  // מחיקת הטקסט משאירה `<w:hyperlink>` ריק, והסמן נשאר בתוכו — הקלדה שם
  // נכשלת ב-hyperlink-nested-unsupported. פסקה חדשה מוציאה אותו החוצה.
  if (countLinks(doc) > 0) {
    await app.press('Enter', 'Enter', 13, 0);
    await app.sleep(400);
  }
}

/** מקליד אזכור וממתין ל-debounce ולתשובה. */
async function typeMention(app, text) {
  await app.type(text);
  await app.sleep(700);
}

const app = await openApp({ name: 'at-mention', port: PORT });

try {
  // המאחז עונה בהצלחה על מתודות הניווט — כך „הלחיצה פתחה” נמדד ולא מונח.
  await app.js(`window.__qaHost.replies['reader.openBook'] = function () {
    return Promise.resolve({ success: true, data: true, error: null });
  };
  window.__qaHost.replies['reader.openSearchTab'] = function () {
    return Promise.resolve({ success: true, data: true, error: null });
  };`);

  await app.js(`window.__qaHost.replies['library.resolveRef'] = function (payload) {
    window.__qaResolveCalls = (window.__qaResolveCalls || []).concat([payload]);
    return Promise.resolve({ success: true, data: ${JSON.stringify(HITS)}, error: null });
  };`);

  await step('הרשימה נפתחת על אזכור, ונשלחת ההפניה בלבד', async () => {
    await app.caretPara(0);
    await typeMention(app, 'ראה @פסחים לד');

    const state = await popup(app);
    if (!state.open) return report.fail('פתיחת הרשימה', 'הרשימה לא נפתחה אחרי הקלדת אזכור');
    report.pass('פתיחת הרשימה');

    const calls = JSON.parse(await app.js('JSON.stringify(window.__qaResolveCalls || [])'));
    const last = calls[calls.length - 1];
    if (last?.ref !== 'פסחים לד') {
      report.fail('ההפניה שנשלחה', `נשלח ${JSON.stringify(last?.ref)} במקום „פסחים לד”`);
    } else {
      report.pass('ההפניה שנשלחה');
    }

    if (state.count !== HITS.length) {
      report.fail('מספר ההצעות', `${state.count} שורות במקום ${HITS.length}`);
    } else {
      report.pass('מספר ההצעות');
    }

    // הרשימה חייבת לשבת מחוץ לעץ ה-DOM של המנוע — ראו engine-boundaries.
    if (state.inEngineTree) {
      report.fail('הרשימה מחוץ למנוע', 'הרשימה נמצאת בתוך עץ ה-DOM של SuperDoc');
    } else {
      report.pass('הרשימה מחוץ למנוע');
    }

    if (state.role !== 'listbox' || !state.activedescendant) {
      report.fail('ARIA', `role=${state.role} activedescendant=${state.activedescendant}`);
    } else {
      report.pass('ARIA');
    }

    const { rect } = state;
    const onScreen =
      rect.left >= 0 && rect.top >= 0 && rect.right <= 4000 && rect.bottom <= 4000 && rect.right > rect.left;
    if (!onScreen) {
      report.fail('מיקום הרשימה', `הרשימה מחוץ למסך: ${JSON.stringify(rect)}`);
    } else {
      report.pass('מיקום הרשימה');
    }
  });

  await step('Enter כותב קישור אמיתי ל-OOXML', async () => {
    await app.press('Enter', 'Enter', 13, 0);
    await app.sleep(900);

    const after = await popup(app);
    if (after.open) report.fail('סגירה אחרי בחירה', 'הרשימה נשארה פתוחה');
    else report.pass('סגירה אחרי בחירה');

    const { doc, rels } = await snap(app);

    if (!/<w:hyperlink[^>]*r:id="([^"]+)"/.test(doc)) {
      return report.fail('הקישור ב-document.xml', 'אין <w:hyperlink r:id> במסמך');
    }
    report.pass('הקישור ב-document.xml');

    const relId = doc.match(/<w:hyperlink[^>]*r:id="([^"]+)"/)[1];
    const rel = new RegExp(`Id="${relId}"[^>]*Target="([^"]+)"`).exec(rels);
    if (!rel) {
      return report.fail('ה-Relationship', `אין Relationship עבור ${relId} ב-document.xml.rels`);
    }
    // ה-Target עובר escaping של XML; & הוא &amp; בקובץ.
    const target = rel[1].replace(/&amp;/g, '&');
    if (target !== 'otzaria://open/book/42?index=1234&uid=id%3A42') {
      report.fail('יעד הקישור', `Target=${target}`);
    } else {
      report.pass('יעד הקישור');
    }

    // הטקסט הנראה הוא ההפניה שנפתרה, וה-@ נעלם.
    const text = doc.replace(/<[^>]+>/g, '');
    if (!text.includes('פסחים דף לד')) {
      report.fail('טקסט הקישור', `לא נמצא „פסחים דף לד” ב: ${text.slice(0, 200)}`);
    } else {
      report.pass('טקסט הקישור');
    }
    if (text.includes('@')) {
      report.fail('ה-@ הוחלף', `נשאר „@” בטקסט: ${text.slice(0, 200)}`);
    } else {
      report.pass('ה-@ הוחלף');
    }
    // המילה שלפני האזכור נשארת במקומה.
    if (!text.includes('ראה')) {
      report.fail('הטקסט שלפני נשמר', `„ראה” נעלם: ${text.slice(0, 200)}`);
    } else {
      report.pass('הטקסט שלפני נשמר');
    }
  });

  await step('הסמן נשאר אחרי הקישור, וההמשך אינו נדחף לפניו', async () => {
    // הבאג: `hyperlinks.insert` השאירה את הסמן לפני הקישור, ולכן „ראה @פסחים
    // לד” ואז „ וכן” יצא „ראה  וכןפסחים דף לד”. הטקסט הנכתב כאן הוא המדידה.
    await app.type(' סוף');
    await app.sleep(700);

    const text = flat((await snap(app)).doc);
    if (!text.includes('פסחים דף לד סוף')) {
      report.fail('הסמן אחרי הקישור', `ההמשך לא נחת אחרי הקישור: ${text.slice(0, 160)}`);
    } else {
      report.pass('הסמן אחרי הקישור');
    }
  });

  await step('קישור שני באותה פסקה — בדיוק מה ש-insert לא ידעה', async () => {
    await typeMention(app, ' וכן @פסחים לד');
    if (!(await popup(app)).open) return report.fail('רשימה לקישור השני', 'הרשימה לא נפתחה');

    await app.press('ArrowDown', 'ArrowDown', 40, 0);
    await app.sleep(200);
    await app.press('Enter', 'Enter', 13, 0);
    await app.sleep(1100);

    const status = await app.status();
    if (status?.error) {
      return report.fail('קישור שני בפסקה', `שורת המצב: ${status.text}`);
    }

    const { doc, rels } = await snap(app);
    // שני הקישורים חייבים לשבת **באותה פסקה** — זה מה שנכשל.
    const paragraph = (doc.match(/<w:p\b[\s\S]*?<\/w:p>/g) ?? []).find(
      (p) => (p.match(/<w:hyperlink\b/g) ?? []).length >= 2,
    );
    if (!paragraph) {
      return report.fail('קישור שני בפסקה', 'אין פסקה עם שני קישורים');
    }
    report.pass('קישור שני בפסקה');

    const ids = [...paragraph.matchAll(/<w:hyperlink[^>]*r:id="([^"]+)"/g)].map((m) => m[1]);
    const targets = ids.map((id) => {
      const rel = new RegExp(`Id="${id}"[^>]*Target="([^"]+)"`).exec(rels);
      return (rel?.[1] ?? '').replace(/&amp;/g, '&');
    });
    if (!targets.some((t) => t.includes('/book/42?')) || !targets.some((t) => t.includes('/book/43?'))) {
      report.fail('שני יעדים נפרדים', `היעדים: ${targets.join(' | ')}`);
    } else {
      report.pass('שני יעדים נפרדים');
    }

    // המזהה היציב נכתב לקישור — בלעדיו המאחז אינו יודע לפתור id עמום.
    if (!targets.every((t) => t.includes('uid='))) {
      report.fail('המזהה היציב בקישור', `היעדים: ${targets.join(' | ')}`);
    } else {
      report.pass('המזהה היציב בקישור');
    }
  });

  await step('תיוג מחדש אחרי מחיקת קישור באותה שורה', async () => {
    await clearDoc(app);
    await typeMention(app, '@פסחים לד');
    await app.press('Enter', 'Enter', 13, 0);
    await app.sleep(1000);

    const created = countLinks((await snap(app)).doc);
    if (created === 0) return report.fail('הקישור הראשון נוצר', 'לא נכתב קישור לפני המחיקה');

    // מחיקת הטקסט בלבד — בדיוק מה שהמשתמש עושה. הצומת נשאר ריק במסמך.
    for (let i = 0; i < 'פסחים דף לד'.length; i += 1) {
      await app.press('Backspace', 'Backspace', 8, 0);
      await app.sleep(40);
    }
    await app.sleep(700);

    await typeMention(app, '@פסחים לד');
    if (!(await popup(app)).open) return report.fail('רשימה אחרי מחיקה', 'הרשימה לא נפתחה');
    await app.press('Enter', 'Enter', 13, 0);
    await app.sleep(1100);

    const status = await app.status();
    if (status?.error) {
      return report.fail('תיוג מחדש', `שורת המצב: ${status.text}`);
    }
    report.pass('תיוג מחדש');

    const { doc } = await snap(app);
    const text = flat(doc);
    if (!text.includes('פסחים דף לד')) {
      report.fail('הקישור החדש נכתב', `הטקסט: ${text.slice(0, 160)}`);
    } else {
      report.pass('הקישור החדש נכתב');
    }

    // ‏`<w:hyperlink>` שאין בתוכו `<w:t>` הוא קישור בלתי-נראה בקובץ שיוצא.
    const blank = (doc.match(/<w:hyperlink\b[^>]*>(?:(?!<w:t)[\s\S])*?<\/w:hyperlink>/g) ?? []).length;
    if (blank > 0) {
      report.fail('אין קישור ריק', `נשארו ${blank} צומתי קישור בלי טקסט`);
    } else {
      report.pass('אין קישור ריק');
    }
  });

  await step('ספר אישי: q= הוא ההפניה שנבחרה, לא מה שהוקלד', async () => {
    await clearDoc(app);
    await app.js(`window.__qaHost.replies['library.resolveRef'] = function () {
      return Promise.resolve({ success: true, data: ${JSON.stringify([USER_HIT])}, error: null });
    };`);

    // הקלדת קידומת בלבד — בדיוק התרחיש שדווח: הרשימה מציגה את השם המלא.
    await typeMention(app, '@בסוגי');
    if (!(await popup(app)).open) return report.fail('רשימה לספר אישי', 'הרשימה לא נפתחה');
    await app.press('Enter', 'Enter', 13, 0);
    await app.sleep(1100);

    const { doc, rels } = await snap(app);
    const ids = [...doc.matchAll(/<w:hyperlink[^>]*r:id="([^"]+)"/g)].map((m) => m[1]);
    const last = ids[ids.length - 1];
    const rel = last ? new RegExp(`Id="${last}"[^>]*Target="([^"]+)"`).exec(rels) : null;
    const target = (rel?.[1] ?? '').replace(/&amp;/g, '&');
    const q = target.includes('?q=') ? decodeURIComponent(target.split('?q=')[1] ?? '') : '';

    if (q !== USER_HIT.reference) {
      report.fail('q= מלא', `q=${JSON.stringify(q)} במקום ${JSON.stringify(USER_HIT.reference)}`);
    } else {
      report.pass('q= מלא');
    }

    // והטקסט הנראה חייב להסכים איתו — שניהם נבנים מאותו מקור.
    if (!flat(doc).includes(USER_HIT.reference)) {
      report.fail('הטקסט הנראה', `לא נמצא „${USER_HIT.reference}” במסמך`);
    } else {
      report.pass('הטקסט הנראה');
    }

    await app.js(`window.__qaHost.replies['library.resolveRef'] = function (payload) {
      window.__qaResolveCalls = (window.__qaResolveCalls || []).concat([payload]);
      return Promise.resolve({ success: true, data: ${JSON.stringify(HITS)}, error: null });
    };`);
  });

  await step('לחיצה על הקישור פותחת אותו באוצריא', async () => {
    await clearDoc(app);
    await typeMention(app, '@פסחים לד');
    await app.press('Enter', 'Enter', 13, 0);
    await app.sleep(1100);

    // המנוע מצייר קישור בסכימה חסומה כ-span ולא כ-<a>; זה מה שמאשר שהמדידה
    // נעשית על המקרה האמיתי ולא על אנקור שהמנוע כבר מפעיל בעצמו.
    const run = JSON.parse(
      await app.js(`(function () {
        var el = document.querySelector('[data-link-rid]');
        if (!el) return JSON.stringify({ found: false });
        var r = el.getBoundingClientRect();
        return JSON.stringify({
          found: true, tag: el.tagName, blocked: el.getAttribute('data-link-blocked'),
          x: r.left + r.width / 2, y: r.top + r.height / 2, w: r.width,
        });
      })()`),
    );
    if (!run.found || run.w === 0) return report.fail('הקישור מצויר', 'לא נמצאה ריצה עם data-link-rid');
    report.pass('הקישור מצויר');

    await app.js('window.__qaHost.calls.length = 0');
    await app.clickAt(run.x, run.y);
    await app.sleep(800);

    const calls = JSON.parse(await app.js('JSON.stringify(window.__qaHost.calls)'));
    const opened = calls.find((c) => c.method === 'reader.openBook');
    if (!opened) {
      return report.fail('הלחיצה פותחת', `לא נקראה reader.openBook. נקראו: ${calls.map((c) => c.method).join(', ') || 'כלום'}`);
    }
    report.pass('הלחיצה פותחת');

    if (opened.payload?.id !== 42 || opened.payload?.index !== 1234) {
      report.fail('היעד שנשלח', `payload=${JSON.stringify(opened.payload)}`);
    } else {
      report.pass('היעד שנשלח');
    }
    // המזהה היציב הוא מה שמונע פתיחה של ספר אחר כשה-id עמום.
    if (opened.payload?.bookUid !== 'id:42') {
      report.fail('המזהה היציב נשלח', `bookUid=${JSON.stringify(opened.payload?.bookUid)}`);
    } else {
      report.pass('המזהה היציב נשלח');
    }
  });

  await step('חצים בוחרים הצעה אחרת', async () => {
    await clearDoc(app);
    await typeMention(app, '@פסחים');

    const before = await popup(app);
    if (!before.open) return report.fail('פתיחה שנייה', 'הרשימה לא נפתחה');

    await app.press('ArrowDown', 'ArrowDown', 40, 0);
    await app.sleep(200);
    const moved = await popup(app);
    if (moved.active !== 1) {
      return report.fail('חץ למטה', `הפריט הפעיל הוא ${moved.active} במקום 1`);
    }
    report.pass('חץ למטה');

    await app.press('Tab', 'Tab', 9, 0);
    await app.sleep(900);

    const { doc, rels } = await snap(app);
    // האחרון ולא הראשון: קישורים ריקים משאריות התרחישים הקודמים נשארים במסמך.
    const ids = [...doc.matchAll(/<w:hyperlink[^>]*r:id="([^"]+)"/g)].map((m) => m[1]);
    if (!ids.length) return report.fail('כתיבה אחרי חץ', 'לא נכתב קישור');
    const rel = new RegExp(`Id="${ids[ids.length - 1]}"[^>]*Target="([^"]+)"`).exec(rels);
    const target = (rel?.[1] ?? '').replace(/&amp;/g, '&');
    if (target !== 'otzaria://open/book/43?index=1300&uid=id%3A43') {
      report.fail('ההצעה השנייה נבחרה', `Target=${target}`);
    } else {
      report.pass('ההצעה השנייה נבחרה');
    }
  });

  await step('Escape סוגר בלי לכתוב', async () => {
    await clearDoc(app);
    const linksBefore = countLinks((await snap(app)).doc);
    await typeMention(app, '@פסחים');
    if (!(await popup(app)).open) return report.fail('פתיחה שלישית', 'הרשימה לא נפתחה');

    await app.press('Escape', 'Escape', 27, 0);
    await app.sleep(300);
    if ((await popup(app)).open) {
      report.fail('Escape סוגר', 'הרשימה נשארה פתוחה');
    } else {
      report.pass('Escape סוגר');
    }

    // ספירה ולא נוכחות: מחיקת הטקסט משאירה `<w:hyperlink>` ריק במסמך, ולכן
    // „יש קישור” אינו אומר „נכתב קישור עכשיו”.
    const { doc } = await snap(app);
    if (countLinks(doc) > linksBefore) {
      report.fail('Escape אינו כותב', 'נוסף קישור למרות Escape');
    } else {
      report.pass('Escape אינו כותב');
    }
    // הטקסט שהוקלד נשאר כפי שהוא — Escape מוותר על ההצעה, לא על ההקלדה.
    const text = doc.replace(/<[^>]+>/g, '');
    if (!text.includes('@פסחים')) {
      report.fail('הטקסט נשאר אחרי Escape', `הטקסט: ${text.slice(0, 120)}`);
    } else {
      report.pass('הטקסט נשאר אחרי Escape');
    }
  });

  await step('כתובת דוא"ל אינה פותחת רשימה', async () => {
    await clearDoc(app);
    await typeMention(app, 'dev@example');
    if ((await popup(app)).open) {
      report.fail('דוא"ל אינו טריגר', 'הרשימה נפתחה על כתובת דוא"ל');
    } else {
      report.pass('דוא"ל אינו טריגר');
    }
  });

  await step('אות שימוש כן פותחת רשימה', async () => {
    await clearDoc(app);
    await typeMention(app, 'כמובא ב@פסחים');
    const state = await popup(app);
    if (!state.open) {
      report.fail('אות שימוש היא טריגר', '„ב@” לא פתח רשימה');
    } else {
      report.pass('אות שימוש היא טריגר');
    }
    await app.press('Escape', 'Escape', 27, 0);
    await app.sleep(200);
  });

  await step('מדידה: איזה מסלול כתיבה קיים במנוע', async () => {
    const shape = JSON.parse(
      await app.js(`(function () {
        var doc = window.__otzariaEditor && window.__otzariaEditor.superdoc
          && window.__otzariaEditor.superdoc.activeEditor
          && window.__otzariaEditor.superdoc.activeEditor.doc;
        if (!doc) return JSON.stringify({ found: false });
        var h = doc.hyperlinks || {};
        return JSON.stringify({
          found: true,
          insert: typeof h.insert,
          wrap: typeof h.wrap,
          patch: typeof h.patch,
          remove: typeof h.remove,
        });
      })()`),
    );
    if (!shape.found) {
      report.fail('צורת ה-API', 'לא נמצאה ידית עורך');
    } else {
      log(`   hyperlinks: insert=${shape.insert} wrap=${shape.wrap} patch=${shape.patch}`);
      // אחד משני המסלולים חייב להיות זמין, אחרת הפיצ'ר אינו יכול לכתוב.
      const canWrite = shape.insert === 'function' || shape.wrap === 'function';
      if (!canWrite) report.fail('צורת ה-API', 'אין insert ואין wrap');
      else report.pass('צורת ה-API');
    }
  });

  await step('מדידה: האם המנוע עדיין חוסם את הסכימה', async () => {
    // זו מדידה של המנוע, לא של הפיצ'ר: הלחיצה עצמה נבדקה למעלה ועובדת דרך
    // הגשר. מה שנמדד כאן הוא האם הגשר עדיין נחוץ — ביום שהסכימה תעבור
    // במעלה הזרם הריצה תצויר כ-`<a>`, הגשר ידלג עליה, ו-`onActivate` ייקח
    // את התפקיד בלי שינוי קוד.
    const pageLog = (await app.log()) ?? [];
    const blocked = pageLog.some((l) => /Blocked potentially unsafe URL/.test(l));
    log(`   הסכימה ${blocked ? 'עדיין נחסמת' : 'עוברת'} — הגשר ${blocked ? 'נחוץ' : 'מיותר'}`);
    report.pass('מצב חסימת הסכימה');
  });

  await step('לא הצטבר רעש', async () => {
    const [status, messages, pageLog] = await Promise.all([app.status(), app.messages(), app.log()]);
    const bad = [];
    if (status?.error) bad.push(`status=${status.text}`);
    const errs = (messages ?? []).filter((m) => m.method === 'ui.showError');
    if (errs.length) bad.push(`showError=${errs.map((m) => m.text).join(' | ')}`);
    const noisy = (pageLog ?? []).filter(
      (l) =>
        !/DevTools|Download the Vue/i.test(l) &&
        // אזהרות של המאחז עצמו: הדמה אינו מממש את שתי המתודות האלה.
        !/reader\.addContextMenuItem|fonts\.listInstalled/.test(l) &&
        // פער מדוד במנוע, לא כשל של הפיצ'ר — נבדק בצעד משלו למטה.
        !/Blocked potentially unsafe URL/.test(l) &&
        // פער ותיק ומתועד של ה-projection במסמך רב-פסקאות (ראו
        // engine/text-search.ts ו-docs/superdoc-2.10-review.md). הוא צץ כאן
        // רק בגלל הפסקאות שהניקוי מוסיף, ואינו נוגע לאזכור.
        !/projection-incomplete|typing-mutation: engine pass failed/.test(l),
    );
    if (noisy.length) bad.push(`log=${noisy.join(' | ')}`);
    if (bad.length) report.fail('ללא רעש', bad.join('; '));
    else report.pass('ללא רעש');
  });
  await step('ספר שאינו נמצא מדווח למשתמש, ולא נבלע', async () => {
    // `reader.openBook` מחזיר `false` כשהספר לא נמצא — זה לא זריקה, וזה מה
    // שנבלע קודם: „לחצתי ולא קרה כלום, וגם לא נאמר לי למה”.

    // הצעד עומד בפני עצמו: הוא רץ אחרי בדיקת הרעש, ולכן אינו יכול להישען
    // על מסמך שצעד קודם השאיר.
    await clearDoc(app);
    await typeMention(app, '@פסחים לד');
    await app.press('Enter', 'Enter', 13, 0);
    await app.sleep(1100);

    const run = JSON.parse(
      await app.js(`(function () {
        var els = [].slice.call(document.querySelectorAll('[data-link-rid]'));
        for (var i = 0; i < els.length; i++) {
          var r = els[i].getBoundingClientRect();
          if (r.width > 0) return JSON.stringify({ found: true, x: r.left + r.width / 2, y: r.top + r.height / 2 });
        }
        return JSON.stringify({ found: false });
      })()`),
    );
    if (!run.found) return report.fail('קישור ללחיצה', 'לא נמצא קישור מצויר');

    await app.js(`window.__qaHost.replies['reader.openBook'] = function () {
      return Promise.resolve({ success: true, data: false, error: null });
    };`);
    await app.clickAt(run.x, run.y);
    await app.sleep(900);

    const status = await app.status();
    if (!status?.error || !/אינו נמצא בספרייה/.test(status.text ?? '')) {
      report.fail('כשל פתיחה מדווח', `שורת המצב: ${JSON.stringify(status)}`);
    } else {
      report.pass('כשל פתיחה מדווח');
    }

    await app.js(`window.__qaHost.replies['reader.openBook'] = function () {
      return Promise.resolve({ success: true, data: true, error: null });
    };`);
  });
} finally {
  app.close();
}

report.print();
process.exit(process.exitCode ?? 0);
