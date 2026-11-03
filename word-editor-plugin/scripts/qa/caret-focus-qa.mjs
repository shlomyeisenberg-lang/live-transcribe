/**
 * „אני על חלון אחר, והסמן ממשיך להבהב” — במנוע אמיתי, בדפדפן אמיתי.
 *
 * ## למה שער ולא בדיקת יחידה
 *
 * שתי הבדיקות שמכסות את התיקון בודקות חצי כל אחת: `tests/unit/caret-focus`
 * מודדת את ההכרעה מול `document.hasFocus` מוחלף, ו-`tests/contract/caret-blink`
 * מודדת שהכלל ב-CSS קיים ושהמעטפת קושרת אליו תכונה. אף אחת מהן אינה יכולה
 * לאבד מיקוד באמת, ואף אחת מהן אינה מריצה את מנוע העימוד.
 *
 * מה שנמדד כאן הוא מה שהמשתמש רואה: `getComputedStyle` על הסמן עצמו, אחרי
 * שחלון אחר לקח את המיקוד. וזה מה שנמדד **לפני** התיקון, ובגללו הוא נכתב:
 *
 *   | מצב              | `document.hasFocus()` | `activeElement`   | האנימציה |
 *   |------------------|-----------------------|-------------------|----------|
 *   | בסיס             | `true`                | `TEXTAREA` במסמך  | רצה      |
 *   | **לשונית אחרת**  | **`false`**           | `TEXTAREA` — לא זז | **רצה**  |
 *
 * ## איך מאבדים מיקוד בשער
 *
 * דרך ה-HTTP של ה-CDP: `PUT /json/new` פותח לשונית חדשה והיא נעשית הפעילה,
 * ו-`/json/activate/<id>` מחזיר את שלנו. נמדד שזה אכן מהפך את
 * `document.hasFocus()` ומייצר `blur` על ה-`window` — בדיוק כמו מעבר לחלון
 * אחר. אין צורך בחלון שני אמיתי.
 *
 * ## ושתי שורות שמודדות את הדפדפן, לא אותנו
 *
 * השתיים האחרונות מודדות מה קורה כשהאלמנט הממוקד **נמחק** — כלומר מה שסגירת
 * טאב עושה למשטח ההקלדה של המנוע. הן מה שמתיר ל-`ui/shell/caret-focus.ts`
 * להסתפק בארבעה מאזיני מיקוד, ו-jsdom אינו יכול להחליף אותן: שם האירועים
 * האלה אינם נורים כלל.
 *
 *   CHROME=<נתיב> node scripts/qa/caret-focus-qa.mjs
 */
import { openApp, createReport, sleep } from './harness.mjs';

const report = createReport('סמן הכתיבה כשאין מיקוד', { strict: true });

const ROW = {
  control: 'בקרה: מיקוד במסמך — הסמן מהבהב',
  away: 'חלון אחר לקח את המיקוד — הסמן נעלם',
  stays: 'והוא נשאר נעלם, לא רק בפריים אחד',
  back: 'חזרה לחלון — הסמן מהבהב שוב',
  control_focus: 'פקד של הממשק לקח את המיקוד — הסמן נעלם',
  restored: 'בקרה הפוכה: חזרה מהפקד למסמך — מהבהב שוב',
  removed_self: 'האלמנט הממוקד הוסר — הדפדפן מדווח, והסמן נעלם',
  removed_ancestor: 'אב של האלמנט הממוקד הוסר — אותו דבר',
  hidden_display: 'אב הוסתר ב-`display: none` — המסלול שהחלפת טאב באמת עוברת בו',
  hidden_visibility: 'אב הוסתר ב-`visibility: hidden` — אותו דבר',
  attribute: 'הסמן החי נושא את `data-v2-local-selection-caret`',
};

/** שם השורה לכל מבנה שנמדד בבלוק „המיקוד שנעלם מתחת לרגליים”. */
const REMOVAL_ROW = {
  self: 'removed_self',
  ancestor: 'removed_ancestor',
  display: 'hidden_display',
  visibility: 'hidden_visibility',
};

const BLINK = 'sd-v2-local-caret-blink';
const PORT = Number(process.env.QA_PORT ?? 9629);

const app = await openApp({ name: 'caret-focus', port: PORT });

/**
 * מצב הסמן כפי שהדפדפן מצייר אותו.
 *
 * ‏`animationName` ולא `opacity` לבדו: כשההבהוב רץ, ה-`opacity` המחושב הוא
 * מה שבמקרה נדגם באותו רגע במחזור — 0 בחצי מהדגימות — ובדיקה שנשענת עליו
 * הייתה מדווחת „נעלם” על סמן תקין. שם האנימציה הוא מה שמבדיל.
 */
const caret = () =>
  app
    .js(`JSON.stringify((function () {
      var shell = document.querySelector('.word-app-shell');
      var el = document.querySelector('.sd-v2-local-selection-caret');
      var byAttr = document.querySelector('[data-v2-local-selection-caret]');
      var cs = el ? getComputedStyle(el) : null;
      return {
        hasFocus: document.hasFocus(),
        idleAttr: shell ? shell.getAttribute('data-caret-idle') : null,
        present: !!el,
        /* shell.css מכבה את ההבהוב דרך שני בוררים, ואת השני המנוע מציב דרך
           dataset — ההסבר המלא ליד שורת הדיווח. */
        attr: !!byAttr,
        attrIsSame: !!byAttr && byAttr === el,
        anim: cs ? cs.animationName : null,
        opacity: cs ? cs.opacity : null,
      };
    })())`)
    .then((s) => JSON.parse(s));

const blinking = (c) => c.present && c.anim === BLINK;
/** „נעלם” = או שאין אלמנט כלל, או שהוא שקוף והאנימציה כבויה. */
const gone = (c) => !c.present || (c.anim === 'none' && c.opacity === '0');
const describe = (c) =>
  `hasFocus=${c.hasFocus}, data-caret-idle=${c.idleAttr}, ` +
  (c.present ? `animation=${c.anim}, opacity=${c.opacity}` : 'אין אלמנט סמן');

/** לשונית אחרת שלוקחת את המיקוד, ופונקציה שמחזירה אותו. */
async function leaveWindow() {
  const created = await (await fetch(`http://127.0.0.1:${PORT}/json/new?about:blank`, { method: 'PUT' })).json();
  await sleep(1_000);
  return async () => {
    const list = await (await fetch(`http://127.0.0.1:${PORT}/json/list`)).json();
    const mine = list.find((t) => t.url.includes('__qa-caret-focus'));
    if (!mine) throw new Error('דף הבדיקה נעלם מרשימת הלשוניות');
    await fetch(`http://127.0.0.1:${PORT}/json/activate/${mine.id}`);
    await fetch(`http://127.0.0.1:${PORT}/json/close/${created.id}`);
    await sleep(1_000);
  };
}

try {
  /* לחיצה בטקסט: סמן אמיתי במקום אמיתי, ולא זה שנפתח עם המסמך הריק. */
  const line = JSON.parse(
    await app.js(`JSON.stringify((function () {
      var root = window.__otzariaEditor && window.__otzariaEditor.container;
      var f = root && root.querySelector('.superdoc-fragment');
      if (!f) return null;
      var b = f.getBoundingClientRect();
      if (b.height < 4) return null;
      return { x: Math.round(b.left + Math.min(40, b.width / 2)), y: Math.round(b.top + b.height / 2) };
    })())`),
  );
  if (line) {
    await app.clickAt(line.x, line.y);
    await sleep(600);
  }

  /* 1. בקרה */
  const base = await caret();
  if (blinking(base)) report.pass(ROW.control, describe(base));
  else report.fail(ROW.control, describe(base));

  /*
   * ‏1ב. הבורר השני של shell.css, על הסמן החי.
   *
   * ‏`tests/contract/caret-blink.test.ts` מציב את התכונה בעצמו, ולכן הוא מודד
   * את המפל ולא את המנוע; והשורות כאן שואלות על המחלקה. כלומר עד עכשיו שום
   * דבר בעץ לא היה מאדים אילו המנוע היה מפסיק להציב את התכונה — והכלל
   * ב-shell.css היה הופך לבורר מת בלי שאיש יבחין.
   */
  if (!base.present) {
    report.stuck(ROW.attribute, 'אין אלמנט סמן — אין על מה למדוד');
  } else if (base.attr && base.attrIsSame) {
    report.pass(ROW.attribute, 'אותו אלמנט נענה גם למחלקה וגם לתכונה');
  } else if (base.attr) {
    report.fail(ROW.attribute, 'התכונה נמצאה על אלמנט אחר מזה שנושא את המחלקה');
  } else {
    report.fail(ROW.attribute, 'המנוע הפסיק להציב את התכונה — הבורר השני ב-shell.css מת');
  }

  /* 2. + 3. חלון אחר */
  const comeBack = await leaveWindow();
  const away = await caret();
  if (away.hasFocus) {
    report.stuck(ROW.away, 'הלשונית החדשה לא לקחה את המיקוד — אין מה למדוד');
    report.stuck(ROW.stays, 'אותה סיבה');
  } else if (gone(away)) {
    report.pass(ROW.away, describe(away));

    const samples = [];
    for (let i = 0; i < 4; i++) {
      await sleep(400);
      samples.push(await caret());
    }
    const visible = samples.filter((s) => !gone(s));
    if (visible.length === 0) report.pass(ROW.stays, `4 דגימות על פני 1.6 שניות, כולן נעלמות`);
    else report.fail(ROW.stays, `הסמן חזר ב-${visible.length} מ-4 דגימות: ${describe(visible[0])}`);
  } else {
    report.fail(ROW.away, describe(away));
    report.skip(ROW.stays, 'השורה שמעליה כבר שבורה');
  }

  /* 4. חזרה */
  await comeBack();
  const back = await caret();
  if (blinking(back)) report.pass(ROW.back, describe(back));
  else report.fail(ROW.back, describe(back));

  /* 5. + 6. פקד של הממשק */
  const combo = JSON.parse(await app.js(`JSON.stringify(window.__qa.rect('גודל גופן', {}))`));
  if (!combo) {
    report.skip(ROW.control_focus, 'תיבת „גודל גופן” לא נמצאה');
    report.skip(ROW.restored, 'אותה סיבה');
  } else {
    await app.clickAt(combo.x, combo.y);
    await sleep(600);
    const inControl = await caret();
    const focusedControl = await app.js(
      `!!(document.activeElement && !document.querySelector('.editor-stack').contains(document.activeElement))`,
    );
    if (!focusedControl) {
      report.stuck(ROW.control_focus, 'התיבה לא לקחה מיקוד — אין מה למדוד');
    } else if (gone(inControl)) {
      report.pass(ROW.control_focus, describe(inControl));
    } else {
      report.fail(ROW.control_focus, describe(inControl));
    }

    if (line) {
      await app.clickAt(line.x, line.y);
      await sleep(700);
      const restored = await caret();
      if (blinking(restored)) report.pass(ROW.restored, describe(restored));
      else report.fail(ROW.restored, describe(restored));
    } else {
      report.skip(ROW.restored, 'אין שורת טקסט לחזור אליה');
    }
  }

  /* 7. המיקוד שנעלם מתחת לרגליים */
  /*
   * סגירת טאב או החלפת מסמך מסירות מ-`.editor-stack` את הפאנל שבתוכו יושב
   * משטח ההקלדה של המנוע. אם הדפדפן אינו מדווח על כך, המודול נשאר על „מהבהב”
   * בזמן שהמיקוד כבר על `<body>` — אותו שקר בדיוק, והפעם בלי שום אירוע
   * שיסגיר אותו.
   *
   * זו הסיבה שאין ב-`ui/shell/caret-focus.ts` שום משגיח DOM: נמדד כאן שהדפדפן
   * **כן** מדווח `blur` ו-`focusout`. ‏jsdom אינו יורה את האירועים האלה, ולכן
   * זו המדידה היחידה שאפשר לסמוך עליה, וזו הסיבה שהיא כאן ולא בבדיקת יחידה.
   *
   * ארבעה מבנים ולא שניים: **הסרה** היא מה שסגירת טאב עושה, אבל החלפת טאב
   * (App.vue, `activateTab`) אינה מסירה דבר — היא מציבה `display: none` על
   * הפאנל של הקודם. זה המסלול השכיח, והוא נמדד כאן במפורש, לצד
   * `visibility: hidden` שגם הוא בשימוש (`.editor-stack__host--pending`).
   *
   * הכפיל **מוסף** לאזור המסמך ולא נלקח ממנו: הסרת משטח ההקלדה של המנוע
   * באמצע השער הייתה משאירה את כל מה שאחריה בלי מסמך.
   */
  const removals = JSON.parse(
    await app.js(`(async () => {
      var stack = document.querySelector('.editor-stack');
      var shell = document.querySelector('.word-app-shell');
      if (!stack || !shell) return JSON.stringify({ error: 'אין אזור מסמך' });
      var wait = function (ms) { return new Promise(function (r) { setTimeout(r, ms); }); };
      var run = async function (mode) {
        var pane = document.createElement('div');
        var probe = document.createElement('textarea');
        probe.style.cssText = 'position:absolute;left:-9999px;width:10px;height:10px';
        pane.appendChild(probe);
        stack.appendChild(pane);
        var events = [];
        var note = function (e) { events.push(e.type); };
        document.addEventListener('focusin', note, true);
        document.addEventListener('focusout', note, true);
        probe.addEventListener('blur', note);
        probe.focus();
        await wait(200);
        var idleWhileFocused = shell.getAttribute('data-caret-idle');
        events.length = 0;
        if (mode === 'self') probe.remove();
        else if (mode === 'ancestor') pane.remove();
        else if (mode === 'display') pane.style.display = 'none';
        else pane.style.visibility = 'hidden';
        await wait(300);
        document.removeEventListener('focusin', note, true);
        document.removeEventListener('focusout', note, true);
        pane.remove();
        return {
          mode: mode,
          events: events.slice(),
          active: document.activeElement ? document.activeElement.tagName : null,
          idleWhileFocused: idleWhileFocused,
          idleAfter: shell.getAttribute('data-caret-idle'),
        };
      };
      return JSON.stringify([await run('self'), await run('ancestor'), await run('display'), await run('visibility')]);
    })()`),
  );

  if (removals.error) {
    for (const key of Object.values(REMOVAL_ROW)) report.stuck(ROW[key], removals.error);
  } else {
    for (const row of removals) {
      const name = ROW[REMOVAL_ROW[row.mode]];
      const detail =
        `אירועים: ${row.events.join(', ') || 'אין'}, activeElement=${row.active}, ` +
        `data-caret-idle ${row.idleWhileFocused} → ${row.idleAfter}`;
      // שתי השורות יחד: הדפדפן דיווח, **וגם** המודול הסיק מזה שאין סמן.
      if (row.events.includes('focusout') && row.active === 'BODY' && row.idleAfter === 'true') {
        report.pass(name, detail);
      } else {
        report.fail(name, detail);
      }
    }
  }

} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
