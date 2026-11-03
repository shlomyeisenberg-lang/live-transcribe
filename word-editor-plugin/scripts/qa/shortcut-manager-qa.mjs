/**
 * שער QA לקיצורים שהמשתמש מגדיר בעצמו: „ניהול קיצורים”, הלוכד, וההחלה/החזרה
 * במסמך אמיתי.
 *
 * ## למה זה לא יכול להיות בדיקת רכיב
 *
 * `tests/unit/` מכסה את ההכרעות הטהורות (נורמליזציה, התנגשות, `decideToggle`),
 * ו-`tests/component/` מרכיב את הדיאלוג ב-jsdom. שלושה דברים אינם קיימים שם
 * בכלל, וכל אחד מהם הוא הבטחה שהפיצ'ר עומד או נופל עליה:
 *
 *   1. **הלוכד לוכד מקש פיזי.** `event.code` בא מהמקלדת, ובדיקה שיוצרת
 *      `KeyboardEvent({code:'KeyK'})` ביד מודדת את מה שהיא כתבה. כאן ההקשה
 *      נשלחת דרך `Input.dispatchKeyEvent`, כלומר מהדפדפן אל האלמנט הממוקד.
 *   2. **מלכודת המקלדת של הלוכד.** „Ctrl+O לא פותח קובץ” הוא טענה על
 *      `stopPropagation` מול מאזין שיושב על `window` **ועל ברירת המחדל של
 *      הדפדפן**. ב-jsdom אין ברירת מחדל, ולכן אין מה למדוד.
 *   3. **ההחלה עצמה.** ערכת עיצוב מגיעה למסמך דרך `CommandAdapter` של מנוע
 *      אמיתי, וההוכחה היא ה-OOXML שיוצא (`w:sz`, `w:b`). ב-jsdom אין מנוע.
 *
 * וגם הפריסה: קבוצת „מידע” בלשונית „קובץ” עברה למחסנית של שניים ליד כפתור
 * גדול, ו-`getBoundingClientRect` של jsdom הוא אפס.
 *
 * ## מה נמדד, ולמה כל שורה קיימת
 *
 * | שורה | מה תיפול עליה |
 * | --- | --- |
 * | קריאה מהאחסון | רשומה פגומה או תפוסה שנשארת ברשימה = שורה שאי אפשר לתקן |
 * | הודעת ההשמטה | קיצור שנעלם בשקט מתגלה בלחיצה שלא עשתה כלום |
 * | גאומטריית „מידע” | מחסנית שגולשת = חיתוך שקט, או קפיצת מסמך בהחלפת לשונית |
 * | הלוכד נחמש | בלי המצב הגלוי המשתמש אינו יודע שהוא אמור להקיש |
 * | Ctrl+O בלוכד | הצירוף נלכד ולא פותח קובץ — שתי טענות, שתיהן חייבות |
 * | סירוב על מובנה | הודעה עברית שנוקבת בפעולה, וכפתור נעול |
 * | Escape בלוכד | מבטל לכידה ו**אינו** סוגר את הדיאלוג |
 * | Tab בלוכד | עוד מזיז מיקוד — אחרת אין דרך לצאת במקלדת |
 * | הוספה | הרשומה בתצוגה **וגם** נכתבה לאחסון |
 * | החלה | `w:sz`/`w:b` ב-OOXML, לא „success: true” |
 * | החזרה | הלחיצה השנייה מחזירה את מה שנקרא, ולא ערך מנוחש |
 * | Ctrl+B | מובנה גובר גם כשיש רשימה אישית |
 * | עריכה במקום | `upsertShortcut` בלי `id` **מוסיפה** — כפילות בהפרש של שורה |
 * | התנגשות אישי↔אישי | ענף אחר של `conflictMessage` מזה שנמדד על מובנה |
 * | מחיקה | הרשומה יורדת גם מהתצוגה וגם מהאחסון |
 * | סמן מכווץ | הזרימה שהפיצ'ר נועד לה: הערכה נכנסת ל-stored marks |
 * | רשימת העזרה | קיצור שאינו מופיע בה נראה כמו קיצור שלא נשמר |
 *
 * ## שורה אחת שאינה פסק דין
 *
 * „התנגשות מוצגת גם לפני שיש שם” היא `partial` בכוונה: `draftProblem` מציגה
 * הודעה אחת בסדר מילוי הטופס, וזו החלטה מתועדת. השורה מודדת מה המשתמש רואה
 * בפועל כשהוא מקליט לפני שהוא ממלא שם, ומשאירה את ההכרעה לאדם.
 *
 * שני תצלומים: `tmp/shortcut-manager-ribbon.png`, `tmp/shortcut-manager-dialog.png`.
 *
 * יציאה 9640 בלבד — שערים אחרים רצים במקביל על יציאות אחרות.
 *
 *   npm run build && node scripts/qa/shortcut-manager-qa.mjs
 */
import { writeFileSync, mkdirSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { openApp, createReport, sleep } from './harness.mjs';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..');
const TMP = join(ROOT, 'tmp');

/** קודי המודיפיירים של `Input.dispatchKeyEvent`. */
const ALT = 1;
const CTRL = 2;
const SHIFT = 8;

/**
 * הרשימה שנזרעת לאחסון **לפני** העלייה, ומה שאמור לקרות לה.
 *
 * שלוש רשומות בכוונה, אחת מכל סוג שהנורמליזציה מכירה: תקינה, פגומה, וכזאת
 * שהצירוף שלה תפוס ברג'יסטרי (`Ctrl+S` = שמירה). זה המסלול היחיד שמוכיח את
 * הכלל „מה שמובנה זוכה גם בדיעבד” — קיצור שנשמר בגרסה שבה הצירוף היה פנוי
 * חייב לנשור, ולא להאפיל על השמירה.
 */
const SEEDED = [
  {
    id: 'seed-ok',
    name: 'ערכה מהאחסון',
    kind: 'format-preset',
    combo: { code: 'KeyJ', ctrl: true, shift: false, alt: true },
    preset: { fontSizePt: 18 },
  },
  // פגומה: אין צירוף בכלל.
  { id: 'seed-broken', name: 'שבורה', kind: 'format-preset', combo: null, preset: { bold: true } },
  // תפוסה: Ctrl+S הוא „שמירת המסמך” ברג'יסטרי.
  {
    id: 'seed-taken',
    name: 'תפוסה',
    kind: 'format-preset',
    combo: { code: 'KeyS', ctrl: true, shift: false, alt: false },
    preset: { bold: true },
  },
];

/**
 * מה שמוזרק **לפני** הבאנדל.
 *
 * שני דברים, ושניהם חייבים להיות מוקדמים מהעלייה:
 *
 *   1. **`storage.get` שמחזיר את הערך עצמו.** דמה המאחז שב-`host-stub.js`
 *      עוטף אותו ב-`{ value }`, ואילו `call()` מחזירה את `data` כמו שהיא —
 *      כלומר כל טוען העדפות בתוסף מקבל `{value: X}` במקום `X` ונופל לברירת
 *      המחדל. `session-probe.mjs` מחזיר את הערך ישר, וזאת הצורה שהמאחז
 *      האמיתי מספק (ראו `loadLastDocument`, שקורא `raw.token`). בלי התיקון
 *      הזה כאן, „הרשימה נטענה מהאחסון” היה נמדד על מסלול שתמיד ריק.
 *   2. **יומן שורת המצב.** ההודעה על רשומות שנשרו נכתבת בזמן העלייה, ופתיחת
 *      המסמך הראשון דורסת אותה תוך שברירי שנייה. מדידה אחרי העלייה הייתה
 *      מפספסת אותה תמיד — כלומר שורה שעוברת בלי למדוד דבר.
 */
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

/**
 * המדידות שבתוך הדף.
 *
 * הכול דרך סלקטורים של הקומפוננטות עצמן, ולא דרך „מה שנראה נכון”: שורה
 * שמודדת אלמנט שאינו קיים מחזירה `found: false` ונופלת, במקום להחזיר אפס
 * שנקרא כמו „בסדר”.
 */
const INSTALL = `window.__smProbe = {
  /* קבוצת „מידע” בלשונית „קובץ”: מה יש בה, איפה, ומה גולש. */
  infoGroup: function () {
    var groups = Array.prototype.slice.call(document.querySelectorAll('.word-ribbon-group'));
    var body = document.querySelector('.word-ribbon-body');
    var group = null;
    groups.forEach(function (g) {
      var t = g.querySelector('.word-group-title');
      if (!group && t && (t.textContent || '').trim() === 'מידע') group = g;
    });
    if (!group || !body) return { found: false, groups: groups.length };

    var content = group.querySelector('.word-group-content');
    var box = function (el) {
      var r = el.getBoundingClientRect();
      return {
        left: Math.round(r.left), right: Math.round(r.right),
        top: Math.round(r.top), bottom: Math.round(r.bottom),
        width: Math.round(r.width), height: Math.round(r.height),
      };
    };

    /* גלישה אנכית מתוך אזור התוכן — אותו כלל של scripts/ribbon-geometry-probe.mjs. */
    var overflow = [];
    var contentBox = content.getBoundingClientRect();
    Array.prototype.forEach.call(content.children, function (child) {
      var cb = child.getBoundingClientRect();
      if (cb.bottom - contentBox.bottom > 0.5 || contentBox.top - cb.top > 0.5) {
        overflow.push(String(child.className).slice(0, 40) + ' ' + Math.round(cb.height) + 'px בתוך ' + Math.round(contentBox.height) + 'px');
      }
    });

    var stacks = group.querySelectorAll('.ribbon-stack');
    var stack = stacks[0] || null;
    var btn = function (el) {
      return {
        name: el.getAttribute('data-tip-title') || (el.textContent || '').trim(),
        variant: el.classList.contains('btn-large') ? 'large' : el.classList.contains('btn-small') ? 'small' : 'icon',
        box: box(el),
      };
    };

    return {
      found: true,
      groupBox: box(group),
      contentBox: box(content),
      bodyBox: box(body),
      /* מיקום הקבוצה בין שכנותיה — „end” מבטיח את הקצה השמאלי. */
      groupIndex: groups.indexOf(group),
      groupCount: groups.length,
      leftmost: groups.every(function (g) {
        return g === group || g.getBoundingClientRect().left >= group.getBoundingClientRect().left - 0.5;
      }),
      stacks: stacks.length,
      stackButtons: stack ? Array.prototype.map.call(stack.querySelectorAll('.word-btn'), btn) : [],
      /* כפתורי הקבוצה שאינם במחסנית. */
      looseButtons: Array.prototype.filter.call(group.querySelectorAll('.word-btn'), function (b) {
        return !b.closest('.ribbon-stack');
      }).map(btn),
      overflow: overflow,
    };
  },

  ribbonHeight: function () {
    var body = document.querySelector('.word-ribbon-body');
    return body ? Math.round(body.getBoundingClientRect().height * 100) / 100 : null;
  },

  /* דיאלוג „ניהול קיצורים”: מצבו, הטופס, והרשימה. */
  manager: function () {
    var root = document.querySelector('.shortmgr-dialog');
    if (!root) return { open: false, dialogs: document.querySelectorAll('[role="dialog"]').length };
    var r = root.getBoundingClientRect();
    var footer = root.querySelector('.sm-footer');
    var fb = footer ? footer.getBoundingClientRect() : null;
    var recorder = root.querySelector('.sm-recorder');
    var primary = root.querySelector('[data-default-action]');
    var error = root.querySelector('.sm-error');
    return {
      open: true,
      role: root.getAttribute('role'),
      modal: root.getAttribute('aria-modal'),
      dialogs: document.querySelectorAll('[role="dialog"]').length,
      box: { left: Math.round(r.left), top: Math.round(r.top), right: Math.round(r.right), bottom: Math.round(r.bottom), width: Math.round(r.width), height: Math.round(r.height) },
      footerBottom: fb ? Math.round(fb.bottom) : null,
      viewport: { width: window.innerWidth, height: window.innerHeight },
      armed: recorder ? recorder.classList.contains('sm-recorder--armed') : null,
      recorderText: recorder ? (recorder.textContent || '').trim() : null,
      recorderFocused: recorder ? document.activeElement === recorder : null,
      activeTag: document.activeElement ? document.activeElement.tagName.toLowerCase() + '.' + String(document.activeElement.className).split(' ')[0] : null,
      primaryLabel: primary ? (primary.textContent || '').trim() : null,
      primaryDisabled: primary ? !!primary.disabled : null,
      error: error ? (error.textContent || '').trim() : null,
      note: (function () { var n = root.querySelector('.sm-note'); return n ? (n.textContent || '').trim() : null; })(),
      /* מה שהטופס מציג עכשיו — כדי ש„עריכה טוענת את הרשומה” תהיה מדידה. */
      form: {
        name: (root.querySelector('#sm-name') || {}).value,
        size: (root.querySelector('#sm-size') || {}).value,
        deleteDisabled: (function () {
          var buttons = Array.prototype.slice.call(root.querySelectorAll('.sm-footer .sm-btn'));
          var hit = buttons.filter(function (b) { return (b.textContent || '').trim() === 'מחיקה'; })[0];
          return hit ? !!hit.disabled : null;
        })(),
        selects: Array.prototype.map.call(root.querySelectorAll('select'), function (s) {
          return { label: s.getAttribute('aria-label'), value: s.value };
        }),
      },
      items: Array.prototype.map.call(root.querySelectorAll('.sm-item'), function (el) {
        return {
          name: (el.querySelector('.sm-item-name') || {}).textContent || '',
          combo: ((el.querySelector('.sm-item-combo') || {}).textContent || '').trim(),
          preset: ((el.querySelector('.sm-item-preset') || {}).textContent || '').trim(),
        };
      }),
      empty: !!root.querySelector('.sm-empty'),
    };
  },

  /* מלבן של שורה ברשימה לפי השם שמוצג בה. */
  itemRect: function (name) {
    var hit = null;
    Array.prototype.forEach.call(document.querySelectorAll('.shortmgr-dialog .sm-item'), function (el) {
      var n = el.querySelector('.sm-item-name');
      if (!hit && n && (n.textContent || '').trim() === name) hit = el;
    });
    if (!hit) return null;
    var r = hit.getBoundingClientRect();
    if (r.width === 0 || r.height === 0) return null;
    return { x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2) };
  },

  /* מלבן של כפתור בפוטר לפי הכתובת שעליו. */
  footerRect: function (label) {
    var hit = null;
    Array.prototype.forEach.call(document.querySelectorAll('.shortmgr-dialog .sm-footer .sm-btn'), function (b) {
      if (!hit && (b.textContent || '').trim() === label) hit = b;
    });
    if (!hit) return null;
    var r = hit.getBoundingClientRect();
    if (r.width === 0 || r.height === 0) return null;
    return { x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2) };
  },

  /* מלבן של אלמנט בדיאלוג לפי סלקטור, ללחיצה אמיתית. */
  rect: function (selector) {
    var el = document.querySelector(selector);
    if (!el) return null;
    var r = el.getBoundingClientRect();
    if (r.width === 0 || r.height === 0) return null;
    return { x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2) };
  },

  /* בורר לפי aria-label — התוויות בטופס אינן data-tip. */
  setSelect: function (label, value) {
    var root = document.querySelector('.shortmgr-dialog');
    if (!root) return 'no-dialog';
    var hit = null;
    Array.prototype.forEach.call(root.querySelectorAll('select'), function (s) {
      if (!hit && s.getAttribute('aria-label') === label) hit = s;
    });
    if (!hit) return 'not-found';
    var ok = Array.prototype.some.call(hit.options, function (o) { return o.value === value; });
    if (!ok) return 'no-option';
    hit.value = value;
    hit.dispatchEvent(new Event('change', { bubbles: true }));
    return 'ok';
  },

  /* דיאלוג „קיצורי מקלדת”: הקבוצות והשורות שבהן. */
  help: function () {
    var root = document.querySelector('.shortcuts-dialog, [aria-labelledby="shortcuts-dialog-title"]');
    if (!root) return { open: false };
    var text = (root.textContent || '').replace(/\\s+/g, ' ');
    return {
      open: true,
      text: text,
      rows: Array.prototype.map.call(root.querySelectorAll('kbd, .shortcut-combo, .sc-combo'), function (el) {
        return (el.textContent || '').trim();
      }),
    };
  },

  statusLog: function () { return window.__statusLog || []; }
};
true`;

const report = createReport('קיצורים אישיים — ניהול, לכידה, החלה והחזרה', { strict: true });
const app = await openApp({
  name: 'shortcut-manager',
  port: Number(process.env.QA_PORT ?? 9640),
  extra: EXTRA,
});

const probe = (expr) => app.js(`JSON.stringify(window.__smProbe.${expr})`).then(JSON.parse);

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

/** גודל גופן כפי שהמנוע מדווח אותו — מספר או `'16pt'`. */
function sizeOf(value) {
  if (value === undefined || value === null) return null;
  const n = Number(String(value).replace(/\s*pt$/i, '').trim());
  return Number.isFinite(n) ? n : null;
}

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1440,
    height: 900,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await sleep(600);
  await app.js(INSTALL);

  /* ================================================================ */
  /* 1 — הקריאה מהאחסון: מה שנשאר, ומה שנשר                            */
  /* ================================================================ */
  const statusLog = await probe('statusLog()');
  /*
   * שתי רשומות נזרעו: אחת פגומה (`combo: null`) ואחת שהצירוף שלה תפוס
   * ברג'יסטרי (Ctrl+S). ההודעה חייבת להבחין ביניהן, ולא לתלות את שתיהן
   * בהתנגשות: משתמש שנשלח לחפש התנגשות על רשומה פגומה מחפש משהו שאינו
   * קיים. זה היה הכשל שנמדד כאן, והנוסח נגזר עכשיו מהספירה לפי סיבה
   * (`dropMessage` ב-ui/shortcuts/custom-shortcuts.ts).
   */
  const dropNotice = statusLog.find((line) => /לא נטענו|לא נטען|לא נקרא/.test(line));
  if (!dropNotice) {
    report.fail(
      'הודעה על קיצורים שנשרו',
      `שתי רשומות נזרעו כפגומות/תפוסות ואיש לא נאמר עליהן דבר. שורת המצב הציגה: ${JSON.stringify(statusLog.slice(0, 6))}`,
    );
  } else if (!/פעולות מובנות|פעולה מובנית/.test(dropNotice)) {
    report.fail('הודעה על קיצורים שנשרו', `הצירוף התפוס אינו מדווח: „${dropNotice}”`);
  } else if (!/לא נקרא/.test(dropNotice)) {
    // הכשל המדויק שתוקן: הרשומה הפגומה נבלעה בתוך „הצירוף תפוס”.
    report.fail(
      'הודעה על קיצורים שנשרו',
      `הרשומה הפגומה מדווחת כהתנגשות, וזה שולח לחפש משהו שאינו קיים: „${dropNotice}”`,
    );
  } else {
    report.pass('הודעה על קיצורים שנשרו', `שתי הסיבות בנפרד — „${dropNotice}”`);
  }

  /* ================================================================ */
  /* 2 — קבוצת „מידע”: מחסנית של שניים, כפתור גדול, בלי גלישה           */
  /* ================================================================ */
  await app.tab('בית');
  await sleep(300);
  const homeHeight = await probe('ribbonHeight()');
  await app.tab('קובץ');
  await sleep(400);
  const fileHeight = await probe('ribbonHeight()');
  const info = await probe('infoGroup()');

  if (!info.found) {
    report.fail('קבוצת „מידע” קיימת', `לא נמצאה קבוצה בשם „מידע” מתוך ${info.groups} קבוצות`);
  } else {
    const stackNames = info.stackButtons.map((b) => b.name);
    const loose = info.looseButtons.map((b) => `${b.name} (${b.variant})`);

    if (info.stacks !== 1) {
      report.fail('מחסנית אחת בקבוצה', `נמצאו ${info.stacks} מחסניות`);
    } else if (stackNames.length !== 2) {
      report.fail('מחסנית של שני פקדים', `${stackNames.length}: ${JSON.stringify(stackNames)}`);
    } else if (stackNames[0] !== 'קיצורים' || stackNames[1] !== 'ניהול קיצורים') {
      report.fail('סדר המחסנית', `${JSON.stringify(stackNames)} — מצופה [קיצורים, ניהול קיצורים]`);
    } else if (!(info.stackButtons[0].box.top < info.stackButtons[1].box.top)) {
      // „מעל” בפועל ולא בסדר ה-DOM: `flex-direction: column` הוא מה שמבטיח זאת.
      report.fail(
        'קיצורים מעל ניהול קיצורים',
        `top ${info.stackButtons[0].box.top} מול ${info.stackButtons[1].box.top}`,
      );
    } else if (info.stackButtons.some((b) => b.variant !== 'small')) {
      /*
        הווריאנט של **שניהם**, ולא של הראשון: כפתור גדול במחסנית אינו מרחיב
        את הרצועה — הוא גולש מגובה קבוע ונחתך בשקט (ראו
        scripts/ribbon-geometry-probe.mjs). נמדד: מוטציה ל-`variant="large"`
        השאירה את הגובה 98px בשתי הלשוניות, ורק שורת הגלישה תפסה אותה.
      */
      report.fail(
        'מחסנית „קיצורים” / „ניהול קיצורים”',
        `וריאנטים ${JSON.stringify(info.stackButtons.map((b) => b.variant))} — במחסנית מותר small בלבד`,
      );
    } else {
      report.pass(
        'מחסנית „קיצורים” / „ניהול קיצורים”',
        `top ${info.stackButtons[0].box.top} ואז ${info.stackButtons[1].box.top}, שניהם small`,
      );
    }

    const about = info.looseButtons.find((b) => b.name === 'אודות');
    if (!about) {
      report.fail('„אודות” גדול בקבוצה', `הפקדים שאינם במחסנית: ${JSON.stringify(loose)}`);
    } else if (about.variant !== 'large') {
      report.fail('„אודות” גדול בקבוצה', `הווריאנט הוא ${about.variant}`);
    } else {
      report.pass('„אודות” גדול בקבוצה', `${about.box.width}×${about.box.height}`);
    }

    if (info.overflow.length) {
      report.fail('אין גלישה מאזור הקבוצה', info.overflow.join(' | '));
    } else {
      report.pass('אין גלישה מאזור הקבוצה', `תוכן ${info.contentBox.height}px`);
    }

    /*
      „נצמדת לקצה השמאלי” — הטענה שכתובה ב-FileTab.vue. בעברית זה הקצה
      הפיזי השמאלי של גוף הרצועה, ולכן היא נמדדת כמרחק ממנו ולא כאינדקס.
    */
    const gap = info.groupBox.left - info.bodyBox.left;
    if (!info.leftmost || gap > 24) {
      report.fail('„מידע” בקצה השמאלי', `מרחק ${gap}px מקצה גוף הרצועה, leftmost=${info.leftmost}`);
    } else {
      report.pass('„מידע” בקצה השמאלי', `${gap}px מהקצה, ${info.groupIndex + 1}/${info.groupCount} קבוצות`);
    }
  }

  if (homeHeight === null || fileHeight === null) {
    report.fail('גובה הרצועה זהה בין הלשוניות', 'גוף הרצועה לא נמצא');
  } else if (Math.abs(homeHeight - fileHeight) > 0.5) {
    report.fail('גובה הרצועה זהה בין הלשוניות', `„בית” ${homeHeight}px מול „קובץ” ${fileHeight}px`);
  } else {
    report.pass('גובה הרצועה זהה בין הלשוניות', `${fileHeight}px בשתיהן`);
  }
  await shoot('shortcut-manager-ribbon');

  /* ================================================================ */
  /* 3 — הדיאלוג נפתח, והרשימה מציגה את מה ששרד                        */
  /* ================================================================ */
  await app.reset();
  if (!(await app.click('ניהול קיצורים'))) {
    throw new Error('הפקד „ניהול קיצורים” לא נמצא בלשונית „קובץ”');
  }
  await sleep(400);
  let m = await probe('manager()');
  if (!m.open) {
    throw new Error(`הלחיצה לא פתחה את „ניהול קיצורים” (דיאלוגים פתוחים: ${m.dialogs})`);
  }
  report.pass('הדיאלוג נפתח', `${m.box.width}×${m.box.height}, role=${m.role} aria-modal=${m.modal}`);

  /*
   * **דיאלוג שנפתח אינו מציג שגיאה.** כלל מתועד של הבית, ב-LinkDialog.vue:
   * „השגיאה מוצגת רק אחרי שהמשתמש הקליד משהו. הצגה על שדה ריק פירושה
   * דיאלוג שנפתח עם הודעת שגיאה, בלי שאיש עשה כלום.” הדיאלוג הזה הפר
   * אותו — הוא נפתח עם „יש לתת שם לקיצור” באדום ובתוך `role="alert"`.
   *
   * השורה הזאת נוספה **אחרי** סבב מוטציה שגילה שהשער אינו מודד את זה:
   * החזרת התקלה השאירה את 43 השורות ירוקות. שער שאינו מודד עובר בירוק
   * על כל רגרסיה, וזה גרוע ממנו שלא להיות.
   */
  if (m.error === null && m.note !== null) {
    report.pass('דיאלוג שנפתח אינו מציג שגיאה', `מוצג ההסבר: „${m.note}”`);
  } else {
    report.fail(
      'דיאלוג שנפתח אינו מציג שגיאה',
      `שגיאה על טופס נקי: „${m.error}” (ההסבר: ${JSON.stringify(m.note)})`,
    );
  }

  /* והכפתור עדיין נעול: „אין שגיאה” אינו „אפשר לשמור”. */
  if (m.primaryDisabled === true) {
    report.pass('„הוספה” נעול על טופס נקי', `התווית „${m.primaryLabel}”`);
  } else {
    report.fail('„הוספה” נעול על טופס נקי', 'הכפתור פעיל על טופס ריק');
  }

  if (m.footerBottom !== null && m.footerBottom <= m.viewport.height + 1) {
    report.pass('הפוטר על המסך', `תחתית ב-${m.footerBottom} מתוך ${m.viewport.height}`);
  } else {
    report.fail('הפוטר על המסך', `תחתית הפוטר ב-${m.footerBottom} מול חלון ${m.viewport.height}`);
  }

  if (m.items.length === 1 && m.items[0].name === 'ערכה מהאחסון') {
    report.pass('הרשימה נטענה מהאחסון', `${JSON.stringify(m.items[0])}`);
  } else {
    report.fail(
      'הרשימה נטענה מהאחסון',
      `מצופה רשומה אחת („ערכה מהאחסון”); בפועל ${m.items.length}: ${JSON.stringify(m.items)}`,
    );
  }
  await shoot('shortcut-manager-dialog');

  /* ================================================================ */
  /* 4 — הלוכד: מלכודת המקלדת, והסירוב על צירוף מובנה                  */
  /* ================================================================ */
  const recorderRect = await probe(`rect(".sm-recorder")`);
  if (!recorderRect) throw new Error('הלוכד לא נמצא בדיאלוג');

  /*
    השם והעיצוב ממולאים **לפני** מדידת ההתנגשות, ולא אחריה.
    `draftProblem` מציגה הודעה אחת בכל רגע, בסדר מילוי הטופס: שם → צירוף →
    ערכה → התנגשות. כלומר על טופס ריק ההודעה היא „יש לתת שם לקיצור” גם כשהצירוף
    שנלכד תפוס — נמדד. שער שמדד את ההתנגשות על טופס ריק היה מאשים את בדיקת
    ההתנגשות בסדר של הטופס, כלומר אדום שאינו מצביע על מה שנשבר.
    (הצד ההפוך — שהמשתמש שמקליט לפני שהוא ממלא שם אינו רואה שהצירוף תפוס —
    נמדד בשורה נפרדת למטה, כי זו טענה על הטופס ולא על ההתנגשות.)
  */
  const NAME = 'ערכת בדיקה';
  const filledName = await app.dialogFill('sm-name', NAME);
  const filledSize = await app.dialogFill('sm-size', '22');
  const setBold = await app.js(
    `window.__smProbe.setSelect(${JSON.stringify('מודגש — מה שהקיצור יקבע')}, "on")`,
  );
  if (filledName !== 'ok' || filledSize !== 'ok' || setBold !== 'ok') {
    throw new Error(`מילוי הטופס נכשל: שם=${filledName}, גודל=${filledSize}, מודגש=${setBold}`);
  }
  await sleep(250);
  await app.clickAt(recorderRect.x, recorderRect.y);
  await sleep(250);
  m = await probe('manager()');
  if (m.armed && m.recorderFocused) {
    report.pass('הלוכד נחמש בלחיצה', `„${m.recorderText}”, המיקוד עליו`);
  } else {
    report.fail('הלוכד נחמש בלחיצה', `armed=${m.armed}, focused=${m.recorderFocused}, טקסט „${m.recorderText}”`);
  }

  /*
    Ctrl+O — שתי טענות בהקשה אחת: הצירוף **נלכד** (כלומר הגיע ללוכד), ו„פתח
    קובץ” **לא רץ**. השנייה נמדדת פעמיים: אין דיאלוג שני, ואין קריאה למאחז.
  */
  await app.press('o', 'KeyO', 79, CTRL);
  await sleep(400);
  m = await probe('manager()');
  const hostAfterCtrlO = (await app.hostCalls()).filter((c) => /^fs\./.test(c.method));
  if (m.recorderText !== 'Ctrl+O') {
    report.fail('Ctrl+O נלכד ולא פתח קובץ', `הלוכד מציג „${m.recorderText}”`);
  } else if (m.dialogs !== 1 || !m.open) {
    report.fail('Ctrl+O נלכד ולא פתח קובץ', `${m.dialogs} דיאלוגים פתוחים — משהו נוסף נפתח`);
  } else if (hostAfterCtrlO.length) {
    report.fail('Ctrl+O נלכד ולא פתח קובץ', `נקראו למאחז: ${hostAfterCtrlO.map((c) => c.method).join(', ')}`);
  } else {
    report.pass('Ctrl+O נלכד ולא פתח קובץ', 'הלוכד מציג Ctrl+O, אין דיאלוג נוסף, אין קריאת fs');
  }

  // ואותה הקשה היא גם הסירוב: Ctrl+O הוא „פתיחת קובץ” ברג'יסטרי.
  if (m.error && m.error.includes('פתיחת קובץ') && m.primaryDisabled === true) {
    report.pass('סירוב על צירוף מובנה (Ctrl+O)', `„${m.error}”, „${m.primaryLabel}” נעול`);
  } else {
    report.fail(
      'סירוב על צירוף מובנה (Ctrl+O)',
      `הודעה: ${JSON.stringify(m.error)}, primaryDisabled=${m.primaryDisabled}`,
    );
  }

  /* Ctrl+S — הצירוף שהמשימה נוקבת בו במפורש. */
  await app.clickAt(recorderRect.x, recorderRect.y);
  await sleep(200);
  await app.press('s', 'KeyS', 83, CTRL);
  await sleep(350);
  m = await probe('manager()');
  const savedCalls = (await app.hostCalls()).filter((c) => /^fs\.|^ui\.save/.test(c.method));
  if (m.recorderText === 'Ctrl+S' && m.error && m.error.includes('שמירת המסמך') && m.primaryDisabled === true && !savedCalls.length) {
    report.pass('סירוב על Ctrl+S', `„${m.error}”`);
  } else {
    report.fail(
      'סירוב על Ctrl+S',
      `לוכד „${m.recorderText}”, הודעה ${JSON.stringify(m.error)}, נעול=${m.primaryDisabled}, קריאות ${JSON.stringify(savedCalls.map((c) => c.method))}`,
    );
  }

  /* Ctrl+B — „מודגש”. אותו כלל, וזה הצירוף שהמשימה דורשת שיישאר מובנה. */
  await app.clickAt(recorderRect.x, recorderRect.y);
  await sleep(200);
  await app.press('b', 'KeyB', 66, CTRL);
  await sleep(350);
  m = await probe('manager()');
  if (m.recorderText === 'Ctrl+B' && m.error && m.error.includes('מודגש')) {
    report.pass('סירוב על Ctrl+B', `„${m.error}”`);
  } else {
    report.fail('סירוב על Ctrl+B', `לוכד „${m.recorderText}”, הודעה ${JSON.stringify(m.error)}`);
  }

  /*
    ומה רואה מי שמקליט **לפני** שהוא ממלא שם.
    שורת מדידה ולא שורת דין: סדר ההודעות הוא החלטה מתועדת ב-`draftProblem`,
    ולכן „חלקי” — היא מודדת מה מוצג בפועל, ואינה מכריעה שהסדר שגוי.
  */
  await app.dialogFill('sm-name', '');
  await sleep(200);
  await app.clickAt(recorderRect.x, recorderRect.y);
  await sleep(200);
  await app.press('s', 'KeyS', 83, CTRL);
  await sleep(350);
  const noName = await probe('manager()');
  if (noName.error && noName.error.includes('שמירת המסמך')) {
    report.pass('התנגשות מוצגת גם לפני שיש שם', `„${noName.error}”`);
  } else {
    report.partial(
      'התנגשות מוצגת גם לפני שיש שם',
      `הלוכד מציג „${noName.recorderText}” אך ההודעה היא „${noName.error}” — הצירוף התפוס אינו מסומן עד שהטופס יושלם`,
    );
  }
  await app.dialogFill('sm-name', NAME);
  await sleep(200);

  /* Escape בזמן לכידה: מבטל את הלכידה, ו**אינו** סוגר את הדיאלוג. */
  await app.clickAt(recorderRect.x, recorderRect.y);
  await sleep(200);
  await app.press('Escape', 'Escape', 27);
  await sleep(300);
  m = await probe('manager()');
  if (m.open && m.armed === false) {
    report.pass('Escape מבטל לכידה ולא סוגר', `הדיאלוג פתוח, armed=${m.armed}`);
  } else {
    report.fail('Escape מבטל לכידה ולא סוגר', m.open ? `הדיאלוג פתוח אך armed=${m.armed}` : 'הדיאלוג נסגר');
  }

  /* Tab בזמן לכידה: המיקוד זז, הלכידה מתבטלת, הדיאלוג נשאר. */
  await app.clickAt(recorderRect.x, recorderRect.y);
  await sleep(200);
  const beforeTab = await probe('manager()');
  await app.press('Tab', 'Tab', 9);
  await sleep(350);
  m = await probe('manager()');
  if (m.open && m.recorderFocused === false && beforeTab.recorderFocused === true) {
    report.pass('Tab עוד מזיז מיקוד', `מהלוכד אל ${m.activeTag}`);
  } else if (m.open && beforeTab.recorderFocused === true) {
    report.fail('Tab עוד מזיז מיקוד', `המיקוד נשאר על הלוכד (activeElement=${m.activeTag})`);
  } else {
    report.fail('Tab עוד מזיז מיקוד', `open=${m.open}, לפני=${beforeTab.recorderFocused}, אחרי=${m.recorderFocused}`);
  }

  /* ================================================================ */
  /* 5 — הוספת קיצור: לכידה, טופס, שמירה לאחסון                        */
  /* ================================================================ */
  await app.clickAt(recorderRect.x, recorderRect.y);
  await sleep(200);
  await app.press('k', 'KeyK', 75, CTRL | ALT);
  await sleep(350);
  m = await probe('manager()');
  if (m.recorderText === 'Ctrl+Alt+K') {
    report.pass('Ctrl+Alt+K נלכד', `הלוכד מציג „${m.recorderText}”`);
  } else {
    report.fail('Ctrl+Alt+K נלכד', `הלוכד מציג „${m.recorderText}”`);
  }

  m = await probe('manager()');
  if (m.error !== null) {
    report.fail('טופס שלם אינו מונע שמירה', `הטופס מלא ועדיין יש הודעת מניעה: „${m.error}”`);
  } else if (m.primaryDisabled !== false) {
    report.fail('טופס שלם אינו מונע שמירה', `„${m.primaryLabel}” נעול בלי הודעה שמסבירה למה`);
  } else {
    report.pass('טופס שלם אינו מונע שמירה', `„${m.primaryLabel}” פעיל, אין הודעת מניעה`);
  }

  await app.reset();
  const addRect = await probe(`rect(".sm-btn-primary")`);
  if (!addRect) throw new Error('כפתור „הוספה” לא נמצא');
  await app.clickAt(addRect.x, addRect.y);
  await sleep(600);
  m = await probe('manager()');
  const added = m.items.find((row) => row.name === NAME);
  if (added) {
    report.pass('„הוספה” הוסיפה לרשימה', `${JSON.stringify(added)}`);
  } else {
    report.fail('„הוספה” הוסיפה לרשימה', `הרשימה: ${JSON.stringify(m.items)}`);
  }

  /* וההוכחה שהיא נשמרת בין הפעלות: הכתיבה לאחסון של אוצריא. */
  const writes = (await app.hostCalls()).filter(
    (c) => c.method === 'storage.set' && c.payload && c.payload.key === 'custom-shortcuts',
  );
  const lastWrite = writes[writes.length - 1];
  const savedNames = Array.isArray(lastWrite?.payload?.value)
    ? lastWrite.payload.value.map((e) => e && e.name)
    : null;
  if (savedNames && savedNames.includes(NAME)) {
    report.pass('נכתב לאחסון', `storage.set(custom-shortcuts) עם ${JSON.stringify(savedNames)}`);
  } else {
    report.fail('נכתב לאחסון', `${writes.length} כתיבות; האחרונה: ${JSON.stringify(lastWrite?.payload ?? null)}`);
  }

  if (m.primaryLabel === 'הוספה' && m.recorderText && m.recorderText.includes('לחצו')) {
    report.pass('הטופס התאפס אחרי הוספה', `הכפתור חזר ל„${m.primaryLabel}”`);
  } else {
    report.fail('הטופס התאפס אחרי הוספה', `כפתור „${m.primaryLabel}”, לוכד „${m.recorderText}”`);
  }

  /* ================================================================ */
  /* 5ב — עריכה במקום ומחיקה: שאר ה-CRUD של הדיאלוג                    */
  /* ================================================================ */
  /*
   * למה זה סעיף ולא „ברור מאליו”: `upsertShortcut` מחליפה **לפי מזהה**, ומי
   * שמקבל `draft` בלי `id` מוסיף רשומה חדשה. כלומר „עריכה שיוצרת כפילות” היא
   * שורת קוד אחת של הפרש, והדרך היחידה לראות אותה היא לערוך בפועל ולמנות את
   * הרשימה. באותה נשימה נמדדת גם ההתנגשות מול קיצור אישי אחר — הענף
   * `kind: 'custom'` של `conflictMessage`, שהוא **לא** זה שנמדד למעלה.
   */
  const seedRect = await probe(`itemRect(${JSON.stringify('ערכה מהאחסון')})`);
  if (!seedRect) throw new Error('שורת „ערכה מהאחסון” לא נמצאה ברשימה');
  await app.clickAt(seedRect.x, seedRect.y);
  await sleep(400);
  m = await probe('manager()');
  if (
    m.primaryLabel === 'שמירה' &&
    m.form.deleteDisabled === false &&
    m.recorderText === 'Ctrl+Alt+J' &&
    m.form.name === 'ערכה מהאחסון' &&
    Number(m.form.size) === 18
  ) {
    report.pass('עריכה טוענת את הרשומה לטופס', `„${m.form.name}”, ${m.form.size} נק', ${m.recorderText}`);
  } else {
    report.fail(
      'עריכה טוענת את הרשומה לטופס',
      `כפתור „${m.primaryLabel}”, מחיקה נעולה=${m.form.deleteDisabled}, לוכד „${m.recorderText}”, שם „${m.form.name}”, גודל „${m.form.size}”`,
    );
  }

  /* התנגשות מול קיצור אישי אחר — ענף אחר של `conflictMessage`. */
  await app.clickAt(recorderRect.x, recorderRect.y);
  await sleep(200);
  await app.press('k', 'KeyK', 75, CTRL | ALT);
  await sleep(350);
  m = await probe('manager()');
  if (m.error && m.error.includes(NAME) && m.primaryDisabled === true) {
    report.pass('התנגשות מול קיצור אישי אחר', `„${m.error}”`);
  } else {
    report.fail('התנגשות מול קיצור אישי אחר', `הודעה ${JSON.stringify(m.error)}, נעול=${m.primaryDisabled}`);
  }

  /* אותו צירוף על הרשומה שנערכת עצמה **אינו** התנגשות (`exceptId`). */
  await app.clickAt(recorderRect.x, recorderRect.y);
  await sleep(200);
  await app.press('j', 'KeyJ', 74, CTRL | ALT);
  await sleep(350);
  m = await probe('manager()');
  if (m.error === null && m.primaryDisabled === false) {
    report.pass('הרשומה שנערכת אינה מתנגשת בעצמה', 'הצירוף המקורי חזר, אין הודעת מניעה');
  } else {
    report.fail('הרשומה שנערכת אינה מתנגשת בעצמה', `הודעה ${JSON.stringify(m.error)}, נעול=${m.primaryDisabled}`);
  }

  const EDITED = 'ערכה שנערכה';
  await app.dialogFill('sm-name', EDITED);
  await sleep(250);
  await app.reset();
  const saveRect = await probe(`footerRect(${JSON.stringify('שמירה')})`);
  if (!saveRect) throw new Error('כפתור „שמירה” לא נמצא בפוטר');
  await app.clickAt(saveRect.x, saveRect.y);
  await sleep(600);
  m = await probe('manager()');
  const names = m.items.map((row) => row.name);
  if (names.length === 2 && names.includes(EDITED) && names.includes(NAME)) {
    report.pass('שמירה מחליפה במקום ואינה מכפילה', JSON.stringify(names));
  } else {
    report.fail('שמירה מחליפה במקום ואינה מכפילה', `${names.length} שורות: ${JSON.stringify(names)}`);
  }

  /* מחיקה. */
  const editedRect = await probe(`itemRect(${JSON.stringify(EDITED)})`);
  if (!editedRect) {
    report.fail('מחיקה מסירה מהרשימה ומהאחסון', `„${EDITED}” לא נמצא ברשימה`);
  } else {
    await app.clickAt(editedRect.x, editedRect.y);
    await sleep(300);
    await app.reset();
    const delRect = await probe(`footerRect(${JSON.stringify('מחיקה')})`);
    if (!delRect) throw new Error('כפתור „מחיקה” לא נמצא בפוטר');
    await app.clickAt(delRect.x, delRect.y);
    await sleep(600);
    m = await probe('manager()');
    const afterDelete = m.items.map((row) => row.name);
    const delWrites = (await app.hostCalls()).filter(
      (c) => c.method === 'storage.set' && c.payload && c.payload.key === 'custom-shortcuts',
    );
    const storedAfter = Array.isArray(delWrites[delWrites.length - 1]?.payload?.value)
      ? delWrites[delWrites.length - 1].payload.value.map((e) => e && e.name)
      : null;
    if (afterDelete.length === 1 && afterDelete[0] === NAME && JSON.stringify(storedAfter) === JSON.stringify([NAME])) {
      report.pass('מחיקה מסירה מהרשימה ומהאחסון', `נשאר ${JSON.stringify(afterDelete)}, באחסון ${JSON.stringify(storedAfter)}`);
    } else {
      report.fail(
        'מחיקה מסירה מהרשימה ומהאחסון',
        `בתצוגה ${JSON.stringify(afterDelete)}, באחסון ${JSON.stringify(storedAfter)}`,
      );
    }
  }

  /* סגירה: „סגירה” בפוטר, לא Escape — זה המסלול של המשתמש. */
  await app.clickDialog('סגירה');
  await sleep(400);
  m = await probe('manager()');
  if (m.open) report.fail('הדיאלוג נסגר', 'הלחיצה על „סגירה” לא סגרה');
  else report.pass('הדיאלוג נסגר', 'אין `.shortmgr-dialog`');

  /* ================================================================ */
  /* 6 — ההחלה במסמך, וההחזרה בלחיצה נוספת                             */
  /* ================================================================ */
  await app.tab('בית');
  await sleep(300);
  await app.caret(0);
  await app.type('abcd', 40);
  await app.press('Home', 'Home', 36);
  await sleep(150);
  for (let i = 0; i < 4; i += 1) {
    await app.press('ArrowRight', 'ArrowRight', 39, SHIFT);
    await sleep(40);
  }
  await sleep(600);

  const before = { bold: await app.cmd('bold'), size: await app.cmd('font-size') };
  const baseSize = sizeOf(before.size.value);
  if (baseSize === null) {
    report.fail('בקרה: גודל הגופן נקרא לפני ההחלה', `font-size מדווח ${JSON.stringify(before.size)}`);
  } else if (baseSize === 22) {
    report.fail('בקרה: גודל הגופן נקרא לפני ההחלה', 'הגודל ההתחלתי הוא 22 — ההחלה לא תהיה מבחינה');
  } else {
    report.pass('בקרה: גודל הגופן נקרא לפני ההחלה', `${baseSize} נק', מודגש=${before.bold.active}`);
  }

  await app.reset();
  await app.press('k', 'KeyK', 75, CTRL | ALT);
  await sleep(1_200);
  const applyStatus = await app.status();
  const afterApply = { bold: await app.cmd('bold'), size: await app.cmd('font-size') };
  const appliedSize = sizeOf(afterApply.size.value);

  if (applyStatus.text && applyStatus.text.includes(NAME) && applyStatus.text.includes('הוחל')) {
    report.pass('הודעת ההחלה', `„${applyStatus.text}”`);
  } else {
    report.fail('הודעת ההחלה', `שורת המצב: ${JSON.stringify(applyStatus)}`);
  }

  if (appliedSize === 22 && afterApply.bold.active === true) {
    report.pass('המנוע מדווח את הערכה', `גודל ${appliedSize}, מודגש=${afterApply.bold.active}`);
  } else {
    report.fail(
      'המנוע מדווח את הערכה',
      `גודל ${appliedSize} (מצופה 22), מודגש=${afterApply.bold.active} (מצופה true)`,
    );
  }

  /*
   * ההוכחה: ה-OOXML. `w:sz` בחצאי נקודות, כלומר 44.
   *
   * `w:b` נמדד **לפי הערך שלו** ולא לפי נוכחות התג: כיבוי מודגש נכתב
   * `<w:b w:val="0"/>`, כלומר התג נשאר ורק המשמעות מתהפכת. נמדד — שער
   * שחיפש „האם התג קיים” קרא „ההחזרה לא עבדה” על מסמך שהוחזר כהלכה.
   */
  const boldTags = (xml) => (xml.match(/<w:b(?:\s[^>]*)?\/>/g) ?? []).filter((t) => !t.startsWith('<w:bCs'));
  const boldOn = (xml) => boldTags(xml).some((tag) => !/w:val="(0|false|off)"/.test(tag));

  const filesApplied = await app.docx();
  const xmlApplied = filesApplied?.['word/document.xml'] ?? '';
  const hasSize44 = /<w:sz\s+w:val="44"\s*\/>/.test(xmlApplied);
  const hasBold = boldOn(xmlApplied);
  if (hasSize44 && hasBold) {
    report.pass('ההחלה ב-OOXML', `w:sz=44 ו-${JSON.stringify(boldTags(xmlApplied))} ב-word/document.xml`);
  } else {
    report.fail(
      'ההחלה ב-OOXML',
      `w:sz=44 ${hasSize44 ? 'נמצא' : 'חסר'}, תגי w:b: ${JSON.stringify(boldTags(xmlApplied))}`,
    );
  }

  /* ולחיצה נוספת — החזרה. */
  await app.press('k', 'KeyK', 75, CTRL | ALT);
  await sleep(1_200);
  const restoreStatus = await app.status();
  const afterRestore = { bold: await app.cmd('bold'), size: await app.cmd('font-size') };
  const restoredSize = sizeOf(afterRestore.size.value);

  if (restoreStatus.text && restoreStatus.text.includes('העיצוב הקודם הוחזר')) {
    report.pass('הודעת ההחזרה', `„${restoreStatus.text}”`);
  } else {
    report.fail('הודעת ההחזרה', `שורת המצב: ${JSON.stringify(restoreStatus)}`);
  }

  if (restoredSize === baseSize && afterRestore.bold.active === before.bold.active) {
    report.pass('ההחזרה החזירה את מה שנקרא', `גודל ${restoredSize}, מודגש=${afterRestore.bold.active}`);
  } else {
    report.fail(
      'ההחזרה החזירה את מה שנקרא',
      `גודל ${restoredSize} (היה ${baseSize}), מודגש=${afterRestore.bold.active} (היה ${before.bold.active})`,
    );
  }

  const filesRestored = await app.docx();
  const xmlRestored = filesRestored?.['word/document.xml'] ?? '';
  const stillBold = boldOn(xmlRestored);
  const stillSize44 = /<w:sz\s+w:val="44"\s*\/>/.test(xmlRestored);
  const backTo12 = new RegExp(`<w:sz\\s+w:val="${baseSize * 2}"\\s*\\/>`).test(xmlRestored);
  if (!stillBold && !stillSize44) {
    report.pass(
      'ההחזרה ב-OOXML',
      `אין מודגש (תגים: ${JSON.stringify(boldTags(xmlRestored))}), w:sz=44 הוסר, ` +
        `w:sz=${baseSize * 2} ${backTo12 ? 'נכתב' : 'לא נכתב (ברירת המחדל של הסגנון)'}`,
    );
  } else {
    report.fail(
      'ההחזרה ב-OOXML',
      `מודגש ${stillBold ? 'נשאר' : 'הוסר'} (${JSON.stringify(boldTags(xmlRestored))}), w:sz=44 ${stillSize44 ? 'נשאר' : 'הוסר'}`,
    );
  }

  /* ================================================================ */
  /* 7 — מובנה גובר: Ctrl+B עוד מדגיש כשיש רשימה אישית                 */
  /* ================================================================ */
  await app.reset();
  const boldBefore = (await app.cmd('bold')).active;
  /*
    יומן שורת המצב ולא הטקסט שעליה: `setStatus` אינו מתנקה מעצמו, ולכן אחרי
    ההחזרה שלמעלה **נשארת** על הפס ההודעה „העיצוב הקודם הוחזר”. שער שקרא את
    הפס אחרי Ctrl+B מצא בה את שם הקיצור והסיק שהקיצור האישי נורה — הודעה
    בת שנייה קודמת שנקראה כעדות. מה שנמדד כאן הוא **הודעה חדשה**.
  */
  const logBefore = (await probe('statusLog()')).length;
  await app.press('b', 'KeyB', 66, CTRL);
  await sleep(700);
  const boldAfter = (await app.cmd('bold')).active;
  if (boldAfter === !boldBefore) {
    report.pass('Ctrl+B עוד מריץ „מודגש”', `${boldBefore} → ${boldAfter}`);
  } else {
    report.fail('Ctrl+B עוד מריץ „מודגש”', `נשאר ${boldAfter} (היה ${boldBefore})`);
  }
  // ובקרה הפוכה: לא נורה עליו קיצור אישי (הוא סורב, ולכן אין לו רשומה).
  const newStatus = (await probe('statusLog()')).slice(logBefore);
  const firedCustom = newStatus.filter((line) => line.includes(NAME));
  if (firedCustom.length) {
    report.fail('Ctrl+B לא הפעיל קיצור אישי', `הודעות חדשות: ${JSON.stringify(firedCustom)}`);
  } else {
    report.pass('Ctrl+B לא הפעיל קיצור אישי', `אין הודעה חדשה על „${NAME}” (${newStatus.length} הודעות חדשות)`);
  }
  // מחזירים את המצב, כדי שהבדיקה הבאה לא תרוץ על טקסט שהשתנה.
  await app.press('b', 'KeyB', 66, CTRL);
  await sleep(500);

  /* ================================================================ */
  /* 8 — „קיצורים”: הרשימה מציגה גם את מה שהמשתמש הגדיר                */
  /* ================================================================ */
  await app.tab('קובץ');
  await sleep(300);
  if (!(await app.click('קיצורים', { exact: true }))) {
    report.fail('„קיצורים” נפתח', 'הפקד לא נמצא');
  } else {
    await sleep(500);
    const help = await probe('help()');
    if (!help.open) {
      report.fail('„קיצורים” נפתח', 'הדיאלוג לא נמצא ב-DOM');
    } else {
      report.pass('„קיצורים” נפתח', `${help.text.length} תווים`);
      const hasGroup = help.text.includes('קיצורים שהגדרתם');
      const hasCombo = help.text.includes('Ctrl+Alt+K');
      const hasName = help.text.includes(NAME);
      if (hasGroup && hasCombo && hasName) {
        report.pass('הקיצור האישי ברשימת העזרה', 'הקבוצה, הצירוף והשם — שלושתם');
      } else {
        report.fail(
          'הקיצור האישי ברשימת העזרה',
          `קבוצה=${hasGroup}, Ctrl+Alt+K=${hasCombo}, שם=${hasName}`,
        );
      }
    }
    await app.escape();
  }

  /* ================================================================ */
  /* 9 — הדיאלוג בחלון נמוך: „הוספה” עוד בהישג יד                      */
  /* ================================================================ */
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1280,
    height: 600,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await sleep(500);
  await app.click('ניהול קיצורים');
  await sleep(500);
  const low = await probe('manager()');
  if (!low.open) {
    report.fail('הפוטר על מסך נמוך', 'הדיאלוג לא נפתח בחלון 600px');
  } else if (low.footerBottom <= low.viewport.height + 1) {
    report.pass('הפוטר על מסך נמוך', `תחתית ב-${low.footerBottom} מתוך ${low.viewport.height}`);
  } else {
    report.fail(
      'הפוטר על מסך נמוך',
      `תחתית הפוטר ב-${low.footerBottom} מול חלון ${low.viewport.height} — „הוספה” מחוץ למסך, ולדיאלוג אין max-height כמו ל„גופן מתקדם”`,
    );
  }

  /* ================================================================ */
  /* 10 — סמן מכווץ: הזרימה שהפיצ'ר נועד לה                            */
  /* ================================================================ */
  /*
   * למה זה סעיף בפני עצמו, ולא וריאנט של סעיף 6.
   *
   * „לעבור לעיצוב עבודה ולחזור ממנו” (ראש custom-shortcuts.ts) הוא בעיקרו
   * **הקלדה**, לא עיצוב של טקסט מסומן: הסמן מכווץ, הקיצור נלחץ, מה שמוקלד
   * מכאן ואילך נושא את הערכה. זה מסלול אחר לגמרי במנוע — ההחלה נכנסת ל-
   * stored marks ולא לטווח — וגם מסלול אחר ב-`readFormat`: על סמן מיושב
   * שתיקת המנוע על צבע נקראת כ„אין צבע” ולא כ„לא נודע” (ראו `readColor`).
   *
   * ההוכחה היא התו שמוקלד **אחרי** הלחיצה: אם הערכה לא נכנסה ל-stored marks,
   * הוא יֵצא בגודל ברירת המחדל ובלי הדגשה.
   */
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1440,
    height: 900,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await sleep(400);
  await app.clickDialog('סגירה');
  await sleep(400);

  await app.caret(0);
  await app.press('End', 'End', 35);
  await sleep(400);
  const collapsed = await app.selection();
  await app.reset();
  const caretLogBefore = (await probe('statusLog()')).length;
  await app.press('k', 'KeyK', 75, CTRL | ALT);
  await sleep(1_200);
  const caretApplyStatus = await app.status();
  await app.type('Z', 60);
  await sleep(900);

  const filesTyped = await app.docx();
  const xmlTyped = filesTyped?.['word/document.xml'] ?? '';
  /* הריצה שמחזיקה את ה-Z, ורק היא: העיצוב חל על מה שהוקלד אחרי הלחיצה. */
  const zRun = (xmlTyped.match(/<w:r>(?:(?!<\/w:r>)[\s\S])*?<w:t[^>]*>[^<]*Z[^<]*<\/w:t>[\s\S]*?<\/w:r>/) ?? [])[0] ?? '';
  const zBold = boldOn(zRun);
  const zSize = /<w:sz\s+w:val="44"\s*\/>/.test(zRun);

  if (caretApplyStatus.error) {
    report.fail(
      'סמן מכווץ — הערכה נכנסת להקלדה הבאה',
      `הלחיצה החזירה שגיאה: „${caretApplyStatus.text}” (בחירה: ${JSON.stringify(collapsed)})`,
    );
  } else if (zBold && zSize) {
    report.pass('סמן מכווץ — הערכה נכנסת להקלדה הבאה', `הריצה של ה-Z: w:sz=44 ו-${JSON.stringify(boldTags(zRun))}`);
  } else {
    report.fail(
      'סמן מכווץ — הערכה נכנסת להקלדה הבאה',
      `w:sz=44 ${zSize ? 'נמצא' : 'חסר'}, מודגש ${zBold ? 'נמצא' : 'חסר'} — שורת המצב אמרה „${caretApplyStatus.text}”. הריצה: ${zRun.slice(0, 220) || '(לא נמצאה ריצה עם Z)'}`,
    );
  }

  /* ומה שהמנוע דיווח על סמן מכווץ — כדי שהודעה מטעה תיראה. */
  const caretMessages = (await probe('statusLog()')).slice(caretLogBefore);
  console.log('הודעות בזרימת הסמן המכווץ:', JSON.stringify(caretMessages));

  /* ================================================================ */
  /* 11 — הלוג של הדף                                                  */
  /* ================================================================ */
  const log = await app.log();
  const hard = log.filter((line) => /^uncaught:|^rejected:|^error:/.test(line));
  const vue = log.filter((line) => /Vue warn|\[Vue/.test(line));
  if (hard.length) {
    report.fail('אין שגיאות בדף', hard.slice(0, 4).join(' | '));
  } else {
    report.pass('אין שגיאות בדף', `${log.length} שורות לוג, אף אחת אינה שגיאה`);
  }
  if (vue.length) {
    report.fail('אין אזהרות Vue', vue.slice(0, 4).join(' | '));
  } else {
    report.pass('אין אזהרות Vue', 'הרכבת הדיאלוג נקייה');
  }
} finally {
  report.print();
  app.close();
}
process.exit(report.rows.some((r) => r.verdict === 'שבור') ? 1 : 0);
