/**
 * „לחצתי במסמך והוא לא מקליד” — לחיצה בשטח האפור שסביב העמוד, במנוע אמיתי.
 *
 * הדיווח: „לפעמים במעבר מהמסמך לכפתורים, כשחוזרים למסמך צריך ללחוץ כמה
 * פעמים בשביל שיקליד”. מה שנמדד: כפתורי הרצועה אינם מעורבים (הם מבטלים
 * `pointerdown` והמיקוד אינו יוצא מהמסמך), אבל **לחיצה בשטח האפור שסביב
 * העמוד** העיפה את המיקוד ל-`<body>` — בעוד הבחירה של המנוע שורדת והסמן
 * ממשיך להיות מצויר. כלומר סמן מהבהב, הקלדה שנבלעת, ולחיצה נוספת שמצילה.
 * התיקון והמדידה: engine/click-focus.ts.
 *
 * ארבע שורות, וכל אחת מהן נכשלה לפני התיקון או מגנה עליו:
 *
 *   1. בקרה — לחיצה על טקסט מקלידה (שהמדידה עצמה עובדת).
 *   2. לחיצה באפור כשהמיקוד במסמך — ההקלדה שאחריה נכנסת.
 *   3. לחיצה באפור כשהמיקוד בתיבת הרצועה — ההקלדה שאחריה נכנסת.
 *   4. פס הגלילה עוד גולל — ה-`preventDefault` נוגע גם בו (המיכל הגולל הוא
 *      ה-container), ואסור שהוא ישבור גלילה.
 *   5. רקע של פס מעטפת (שורת המצב, פס הלשוניות, הכותרת, שורת הלשוניות של
 *      הרצועה, וגוף הרצועה שבין הכפתורים) — כל אחד מהם העיף את המיקוד
 *      ל-`<body>` (ui/shell/caret-keeper.ts).
 *   6. ובקרה הפוכה: פקד בתוך פס עוד מקבל מיקוד, אחרת השמירה הזאת שברה את
 *      תיבות הרצועה.
 *
 * מסמך הבדיקה נבנה כאן (60 פסקאות) כדי שיהיה עמוד ברוחב מוגדר עם אפור סביבו.
 *
 *   CHROME=<נתיב> node scripts/qa/click-focus-qa.mjs
 */
import { openApp, createReport, sleep } from './harness.mjs';
import { buildDocx } from './docx-fixtures.mjs';
import { readFileSync, writeFileSync, rmSync, mkdtempSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

const report = createReport('מיקוד אחרי לחיצה מחוץ לעמוד', { strict: true });

const ROW = {
  control: 'בקרה: לחיצה על טקסט מקלידה',
  gray: 'לחיצה בשטח האפור — ההקלדה שאחריה נכנסת',
  ribbon: 'לחיצה באפור כשהמיקוד בתיבת הרצועה — ההקלדה שאחריה נכנסת',
  scroll: 'פס הגלילה של אזור המסמך עוד גולל',
  chrome: 'רקע של פס מעטפת אינו לוקח את הסמן',
  barControl: 'פקד בפס עוד מקבל מיקוד',
};

/** מסמך עם פסקאות ארוכות: עמוד מלא, ואפור משני צדיו. */
const WORDS = 'בראשית ברא אלהים את השמים ואת הארץ והארץ היתה תהו ובהו וחשך על פני תהום'.split(' ');
const paragraph = (i) =>
  `<w:p><w:pPr><w:bidi/></w:pPr>` +
  Array.from({ length: 12 }, (_, k) => {
    const word = WORDS[(i * 7 + k) % WORDS.length];
    return `<w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">${word} </w:t></w:r>`;
  }).join('') +
  `</w:p>`;
const body = Array.from({ length: 60 }, (_, i) => paragraph(i)).join('');

const dir = mkdtempSync(join(tmpdir(), 'owe-click-focus-'));
const docxPath = join(dir, 'click-focus.docx');
writeFileSync(docxPath, buildDocx({ body }));

/** דמה של „פתח קובץ”, כמו ב-typing-latency-probe.mjs. */
function docxStub(path) {
  const b64 = readFileSync(path).toString('base64');
  return `
<script>
(function () {
  function install() {
    if (!window.__qaHost) return setTimeout(install, 20);
    var bytes = Uint8Array.from(atob(${JSON.stringify(b64)}), function (c) { return c.charCodeAt(0); });
    var url = URL.createObjectURL(new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }));
    window.__qaHost.replies['fs.pickUserFile'] = function () {
      return Promise.resolve({ success: true, data: { cancelled: false, token: 'click-focus', url: url, name: 'click-focus.docx', size: bytes.length, access: 'read' }, error: null });
    };
  }
  install();
})();
</script>`;
}
const app = await openApp({
  name: 'click-focus',
  port: Number(process.env.QA_PORT ?? 9621),
  extra: docxStub(docxPath),
});

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1400,
    height: 1200,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await sleep(1_200);

  /* פתיחת המסמך דרך הרצועה, כמו המשתמש */
  await app.tab('קובץ');
  await sleep(400);
  if (!(await app.click('פתח קובץ'))) throw new Error('הפקד „פתח קובץ” לא נמצא');
  await sleep(800);
  const browse = JSON.parse(
    await app.js(
      `JSON.stringify((function(){var d=document.querySelector('[role="dialog"]');if(!d)return null;var b=Array.prototype.find.call(d.querySelectorAll('button'),function(x){return /עיון בקבצים/.test(x.textContent||'')});return b?window.__qa.rectOf(b):null;})())`,
    ),
  );
  if (!browse) throw new Error('„עיון בקבצים” לא נמצא');
  await app.clickAt(browse.x, browse.y);
  let len = 0;
  for (let waited = 0; waited < 90_000; waited += 500) {
    await sleep(500);
    len = await app.js('(window.__otzariaEditor && window.__otzariaEditor.container.textContent.length) || 0');
    if (len > 2_000) break;
  }
  if (len <= 2_000) throw new Error('המסמך לא נפתח בזמן');
  // הפריסה „מסתדרת” עוד כמה שניות אחרי שהטקסט בעץ; מדידה בזמן הזה מודדת פתיחה.
  await sleep(5_000);
  await app.tab('בית');
  await sleep(600);

  const geometry = () =>
    app
      .js(`JSON.stringify((function () {
        var W = innerWidth, H = innerHeight;
        var root = window.__otzariaEditor.container;
        var page = [].slice.call(root.querySelectorAll('[data-page-index]')).filter(function (p) {
          var b = p.getBoundingClientRect();
          return b.height > 100 && b.bottom > 200 && b.top < H - 100;
        })[0];
        if (!page) return null;
        var pb = page.getBoundingClientRect();
        var line = null;
        var frags = [].slice.call(root.querySelectorAll('.superdoc-fragment[data-source-node-id]'))
          .filter(function (f) { return !f.closest('.superdoc-page-header, .superdoc-page-footer'); });
        for (var i = 0; i < frags.length; i++) {
          var box = frags[i].querySelector('.superdoc-line') || frags[i];
          var b = box.getBoundingClientRect();
          if (b.height < 6 || b.top < 160 || b.bottom > H - 45) continue;
          var right = Math.min(b.right, W - 30), left = Math.max(b.left, 20);
          if (right - left < 60) continue;
          line = { x: Math.round(right - Math.min(60, (right - left) / 2)), y: Math.round(b.top + b.height / 2) };
          break;
        }
        return {
          line: line,
          gray: pb.left > 60 ? { x: Math.round(pb.left / 2), y: line ? line.y : Math.round(pb.top + 60) } : null,
          scroller: { pad: root.offsetWidth - root.clientWidth, right: Math.round(pb.right), hostRight: Math.round(root.getBoundingClientRect().right), midY: Math.round(root.getBoundingClientRect().top + root.getBoundingClientRect().height / 2), scrollTop: Math.round(root.scrollTop) },
        };
      })())`)
      .then((s) => JSON.parse(s));

  const focusInfo = () =>
    app
      .js(`JSON.stringify((function () {
        var a = document.activeElement;
        var c = window.__otzariaEditor.container;
        return { tag: a ? a.tagName : null, inEditor: !!(a && c.contains(a)) };
      })())`)
      .then((s) => JSON.parse(s));

  const bodyLen = () => app.js('window.__otzariaEditor.container.textContent.length');

  /** הקשה אחת, בלי לחיצה נוספת. `true` = נכנסה למסמך. */
  async function typed() {
    const before = await bodyLen();
    await app.press('x', 'KeyX', 88, 0, 'x');
    for (let waited = 0; waited < 3_000; waited += 100) {
      await sleep(100);
      if ((await bodyLen()) > before) return true;
    }
    return false;
  }

  const geo = await geometry();
  if (!geo?.line) throw new Error('אין שורת טקסט גלויה למדידה');
  const clickLine = async () => {
    const g = await geometry();
    const line = g?.line ?? geo.line;
    await app.clickAt(line.x, line.y);
    await sleep(450);
  };

  /* 1. בקרה */
  await clickLine();
  const controlFocus = await focusInfo();
  if (await typed()) report.pass(ROW.control, `מיקוד ${controlFocus.tag}`);
  else report.fail(ROW.control, `ההקשה לא נכנסה; מיקוד ${controlFocus.tag}`);

  /* 2. לחיצה באפור כשהמיקוד במסמך */
  if (!geo.gray) {
    report.skip(ROW.gray, 'העמוד תופס את כל הרוחב — אין אפור למדוד');
  } else {
    await clickLine();
    await app.clickAt(geo.gray.x, geo.gray.y);
    await sleep(500);
    const after = await focusInfo();
    if (await typed()) report.pass(ROW.gray, `מיקוד אחרי הלחיצה: ${after.tag}`);
    else report.fail(ROW.gray, `המיקוד עבר ל-${after.tag} וההקלדה נבלעה`);
  }

  /* 3. לחיצה באפור כשהמיקוד בתיבת הרצועה */
  if (!geo.gray) {
    report.skip(ROW.ribbon, 'אין אפור למדוד');
  } else {
    const combo = JSON.parse(await app.js(`JSON.stringify(window.__qa.rect('גודל גופן', {}))`));
    if (!combo) {
      report.skip(ROW.ribbon, 'תיבת „גודל גופן” לא נמצאה');
    } else {
      await clickLine();
      await app.clickAt(combo.x, combo.y);
      await sleep(500);
      const inCombo = await focusInfo();
      await app.clickAt(geo.gray.x, geo.gray.y);
      await sleep(600);
      const after = await focusInfo();
      if (await typed()) report.pass(ROW.ribbon, `${inCombo.tag} → ${after.tag}`);
      else report.fail(ROW.ribbon, `המיקוד נשאר ב-${after.tag} וההקלדה נבלעה`);
    }
  }

  /* 4. פס הגלילה */
  {
    const g = await geometry();
    const pad = g?.scroller?.pad ?? 0;
    if (pad < 4) {
      report.skip(ROW.scroll, 'אין פס גלילה מקומי במיכל');
    } else {
      await clickLine();
      const start = (await geometry()).scroller.scrollTop;
      const x = Math.round(g.scroller.hostRight - Math.max(4, pad / 2));
      await app.clickAt(x, Math.round(g.scroller.midY + 150));
      await sleep(900);
      const end = (await geometry()).scroller.scrollTop;
      if (end !== start) report.pass(ROW.scroll, `scrollTop ${start} → ${end}`);
      else report.fail(ROW.scroll, `scrollTop נשאר ${start} — הביטול שבר את הגלילה`);
    }
  }
  /* 5. רקע של פסי המעטפת */
  {
    const CONTROL =
      'button, input, select, textarea, a, [role="button"], [role="tab"], [role="combobox"], [role="menuitem"], [role="option"], [tabindex]';
    const BARS = [
      ['שורת המצב', '.word-statusbar'],
      ['פס לשוניות המסמכים', '.word-doctabs-bar'],
      ['סרגל הכותרת', '.word-titlebar'],
      ['שורת לשוניות הרצועה', '.word-tab-strip'],
      ['גוף הרצועה', '.word-ribbon-body'],
    ];
    /** נקודה בפס שאינה על פקד. */
    const emptyIn = (selector) =>
      app
        .js(`JSON.stringify((function () {
          var bar = document.querySelector(${JSON.stringify(selector)});
          if (!bar) return null;
          var b = bar.getBoundingClientRect();
          if (!b.width || !b.height) return null;
          var y = Math.round(b.y + b.height / 2);
          for (var x = Math.round(b.x + 4); x < b.x + b.width - 4; x += 6) {
            var el = document.elementFromPoint(x, y);
            if (!el || !bar.contains(el)) continue;
            if (el.closest(${JSON.stringify(CONTROL)})) continue;
            return { x: x, y: y };
          }
          return null;
        })())`)
        .then((s) => JSON.parse(s));

    const dead = [];
    const measured = [];
    for (const [label, selector] of BARS) {
      const point = await emptyIn(selector);
      if (!point) continue;
      await clickLine();
      await app.clickAt(point.x, point.y);
      await sleep(450);
      const after = await focusInfo();
      const ok = await typed();
      measured.push(label);
      if (!ok) dead.push(`${label} (${after.tag})`);
    }
    if (measured.length === 0) report.skip(ROW.chrome, 'לא נמצאה נקודה ריקה באף פס');
    else if (dead.length === 0) report.pass(ROW.chrome, `${measured.length} פסים: ${measured.join(', ')}`);
    else report.fail(ROW.chrome, `ההקלדה מתה אחרי: ${dead.join(', ')}`);
  }

  /* 6. בקרה הפוכה: פקד בפס עוד מקבל מיקוד */
  {
    const combo = JSON.parse(await app.js(`JSON.stringify(window.__qa.rect('גודל גופן', {}))`));
    if (!combo) {
      report.skip(ROW.barControl, 'תיבת „גודל גופן” לא נמצאה');
    } else {
      await clickLine();
      await app.clickAt(combo.x, combo.y);
      await sleep(450);
      const after = await focusInfo();
      if (after.tag === 'INPUT') report.pass(ROW.barControl, 'תיבת הגודל קיבלה מיקוד');
      else report.fail(ROW.barControl, `המיקוד נשאר ב-${after.tag} — השמירה חסמה את הפקד`);
    }
  }
} finally {
  app.close();
  rmSync(dir, { recursive: true, force: true });
}

process.exit(report.print() > 0 ? 1 : 0);
