/**
 * שער „מותקן שאינו מצוייר”: גופן שהמארח מדווח עליו ושהדפדפן אינו פותר.
 *
 * ## למה זה השער היחיד שיכול למדוד את התיקון
 *
 * הצירוף הזה הסתיר 43 מ-287 המשפחות המותקנות במכונה שנמדדה, חמש מהן עבריות
 * (`Guttman Kav-Light`, `Segoe UI Semilight` ועוד) — כולן מופיעות ברשימת
 * הגופנים של אוצריא עצמה, ואף אחת מהן לא הייתה נגישה בבורר. `keepAvailable`
 * מחקה אותן, מפני שהיא סוננה את רשימת המארח באותה מסננת שמסננת את רשימת
 * הניחושים (src/engine/system-fonts.ts).
 *
 * בדיקות היחידה מזריקות את `available`, ולכן הן מודדות את ההכרעות ולא את
 * המדידה — ומדידה של שם גופן דורשת canvas אמיתי ופותר גופנים אמיתי. וגם השערים
 * הקיימים אינם מגיעים לכאן: המאחז אינו מממש `fonts.listInstalled`, ולכן בהם
 * התוסף נופל תמיד לשכבת הניחושים ומסלול המארח אינו נבדק כלל.
 *
 * ## למה השמות מומצאים ולא אמיתיים
 *
 * `Segoe UI Semilight` הוא דוגמה אמיתית — נמדד כלא-נפתר בכרום ב-Windows —
 * אבל שער שנשען עליו מודד את מערכת הגופנים של המכונה ולא את הקוד: במכונה
 * אחרת הוא עשוי להיפתר, ואז השער עובר בלי לבדוק דבר. שם מומצא אינו נפתר
 * **בשום** מכונה, ולכן הוא הקלט היחיד שהופך את השער לדטרמיניסטי.
 *
 * ומי שמכריז שהוא מותקן הוא המאחז — כלומר בדיוק התפקיד של המארח במסלול הזה.
 * מה שנמדד הוא הקוד: „המארח אמר מותקן, הדפדפן אינו פותר — מה הבורר עושה”.
 *
 * ## חמש השאלות
 *
 * 1. **השורה בכלל שם?** לפני התיקון היא נמחקה, וזו התלונה עצמה.
 * 2. **היא מסומנת?** שורה שמכריזה שם ומציירת גופן אחר היא השקר שהדגל מונע.
 * 3. **ההסבר נכון?** „אינו מותקן במכונה” הוא שקר על גופן שהמשתמש התקין.
 * 4. **הבחירה שולפת בייטים?** ועם השם עצמו כתחליף, בלי הגשה של גופן אחר.
 * 5. **השורה מתעוררת?** אחרי ההזרקה — דגימה, בלי סימון, בלי רענון ידני.
 *
 *   npm run build && node scripts/qa/undrawable-fonts-qa.mjs
 *
 * יציאה 9369.
 */
import { openApp, createReport } from './harness.mjs';

const report = createReport('מותקן שאינו מצוייר', { strict: true });

/** שם שאף מכונה אינה פותרת. ראו „למה השמות מומצאים” למעלה. */
const UNDRAWABLE = 'QA Undrawable Family';
/** שם שכל Windows פותר, ושאינו באף אחת מהרשימות הקבועות של התוסף. */
const DRAWABLE = 'Georgia';

const NOT_INSTALLED_TIP = 'הגופן אינו מותקן במכונה — אין דגימה להציג';
const UNDRAWABLE_TIP = 'הגופן מותקן אך הדפדפן אינו מצייר אותו — אין דגימה להציג';

/**
 * המאחז, מוגדר **לפני** שהאפליקציה עולה.
 *
 * `loadInstalledFonts` נורית פעם אחת ב-onMounted, ולכן תשובה שנרשמת אחרי
 * העלייה אינה משנה דבר. `extra` מוזרק אחרי המאחז ולפני החבילה, וזה החלון.
 *
 * ו-`resolveFamilies` מגיש `local("Arial")` תחת השם המבוקש: אלה בייטים
 * שהדפדפן באמת יודע לצייר, כלומר ההזרקה הופכת את השם לנפתר — בדיוק מה
 * שהמארח האמיתי עושה עם הבייטים של הגופן. גופן ארוז בשער היה מוסיף מאות
 * קילובייטים ל-dist בלי להוסיף דבר למדידה.
 */
const EXTRA = `<script>
(function () {
  var H = window.__qaHost;
  H.replies['fonts.listInstalled'] = {
    platform: 'windows',
    families: [
      { name: ${JSON.stringify(DRAWABLE)}, scripts: ['hebrew', 'latin'], monospace: false },
      { name: ${JSON.stringify(UNDRAWABLE)}, scripts: ['hebrew'], monospace: false }
    ]
  };
  H.resolveCalls = [];
  H.replies['fonts.resolveFamilies'] = function (payload) {
    var families = (payload && payload.families) || [];
    H.resolveCalls.push(JSON.parse(JSON.stringify(families)));
    var css = families
      .map(function (f) {
        return '@font-face{font-family:' + JSON.stringify(f.name) + ';src:local("Arial");}';
      })
      .join('\\n');
    return Promise.resolve({
      success: true,
      data: { css: css, resolved: families.map(function (f) { return f.name; }) },
      error: null
    });
  };
})();
</script>`;

const app = await openApp({
  name: 'undrawable-fonts',
  port: Number(process.env.QA_PORT ?? 9369),
  extra: EXTRA,
});

/** שורה אחת בבורר, עם מה ש-`Q.options` אינו מחזיר: הסימון וההסבר. */
const readRow = (value) =>
  app.js(`(() => {
    const box = [...document.querySelectorAll('.ribbon-combo')]
      .find((el) => (el.getAttribute('title') || el.textContent || '').includes('גופן'));
    const input = document.querySelector('.ribbon-combo input');
    if (!input) return JSON.stringify({ error: 'אין בורר' });
    input.focus();
    input.dispatchEvent(new FocusEvent('focus'));
    return new Promise((resolve) => setTimeout(() => {
      const row = [...document.querySelectorAll('[role="option"]')]
        .find((el) => el.getAttribute('data-value') === ${JSON.stringify(value)});
      resolve(JSON.stringify(row === undefined ? { present: false } : {
        present: true,
        availability: row.getAttribute('data-availability'),
        tip: row.getAttribute('data-tip-title'),
        aria: row.getAttribute('aria-label'),
        group: row.getAttribute('data-group'),
        unavailable: row.classList.contains('unavailable'),
        fontFamily: getComputedStyle(row).fontFamily,
      }));
    }, 250));
  })()`);

const parse = (raw) => {
  try {
    return JSON.parse(raw);
  } catch {
    return null;
  }
};

try {
  await app.tab('בית');
  // המנייה נורית ב-onMounted ומוותרת על החוט; היא נוחתת אחרי המסמך הראשון.
  await app.sleep(2500);

  /* -------------------------------------------------------------- */
  /* 1 + 2 + 3 — השורה שם, מסומנת, וההסבר נכון                       */
  /* -------------------------------------------------------------- */
  const before = parse(await readRow(UNDRAWABLE));
  console.log('השורה לפני בחירה:', JSON.stringify(before));

  if (!before) {
    report.stuck('קריאת השורה', 'הדף לא החזיר JSON');
    throw new Error('אין מה למדוד');
  }

  before.present
    ? report.pass('גופן מותקן שאינו נפתר מופיע בבורר', `בקבוצה „${before.group}”`)
    : report.fail(
        'גופן מותקן שאינו נפתר מופיע בבורר',
        'השורה נעדרת — זו בדיוק המחיקה ש-43 גופנים מותקנים נפלו בה',
      );

  if (before.present) {
    before.unavailable && before.availability === 'missing'
      ? report.pass('השורה מסומנת ואינה מכריזה דגימה', `data-availability=${before.availability}`)
      : report.fail(
          'השורה מסומנת ואינה מכריזה דגימה',
          `unavailable=${before.unavailable} availability=${before.availability}`,
        );

    if (before.tip === UNDRAWABLE_TIP) {
      report.pass('ההסבר אומר „מותקן” ולא „אינו מותקן”', before.tip);
    } else if (before.tip === NOT_INSTALLED_TIP) {
      report.fail(
        'ההסבר אומר „מותקן” ולא „אינו מותקן”',
        'השורה אומרת „אינו מותקן במכונה” על גופן שהמארח דיווח כמותקן',
      );
    } else {
      report.fail('ההסבר אומר „מותקן” ולא „אינו מותקן”', `הסבר לא מזוהה: ${before.tip}`);
    }
  } else {
    report.skip('השורה מסומנת ואינה מכריזה דגימה', 'אין שורה');
    report.skip('ההסבר אומר „מותקן” ולא „אינו מותקן”', 'אין שורה');
  }

  /* -------------------------------------------------------------- */
  /* השורה שכן נפתרת — כדי שהסימון יאמר משהו                          */
  /* -------------------------------------------------------------- */
  const drawable = parse(await readRow(DRAWABLE));
  console.log('שורה שנפתרת:', JSON.stringify(drawable));
  drawable?.present && drawable.unavailable === false
    ? report.pass('גופן מהמנייה שכן נפתר אינו מסומן', `availability=${drawable.availability}`)
    : report.fail(
        'גופן מהמנייה שכן נפתר אינו מסומן',
        `present=${drawable?.present} unavailable=${drawable?.unavailable}`,
      );

  /* -------------------------------------------------------------- */
  /* 4 — הבחירה שולפת בייטים, עם השם כתחליף של עצמו                  */
  /* -------------------------------------------------------------- */
  const chose = await app.selectValue('גופן', UNDRAWABLE);
  console.log('בחירה:', chose);
  await app.sleep(1200);

  const asked = parse(await app.js('JSON.stringify(window.__qaHost.resolveCalls || [])')) ?? [];
  console.log('בקשות resolveFamilies:', JSON.stringify(asked));

  const mine = asked.find((batch) => batch.some((f) => f.name === UNDRAWABLE));
  if (!mine) {
    report.fail('הבחירה שולפת את הבייטים מאוצריא', 'לא נשלחה שום בקשה לשם שנבחר');
  } else {
    const entry = mine.find((f) => f.name === UNDRAWABLE);
    const substitutes = entry.substitutes ?? [];
    substitutes.length === 1 && substitutes[0] === UNDRAWABLE
      ? report.pass('הבחירה שולפת את הבייטים מאוצריא', 'השם כתחליף של עצמו, בלי תחליפים')
      : report.fail(
          'הבחירה שולפת את הבייטים מאוצריא',
          `נשלחו תחליפים — גופן אחר יוגש תחת השם שנבחר: ${JSON.stringify(substitutes)}`,
        );
  }

  /* -------------------------------------------------------------- */
  /* 5 — השורה מתעוררת: דגימה, בלי סימון, בלי רענון ידני              */
  /* -------------------------------------------------------------- */
  const after = parse(await readRow(UNDRAWABLE));
  console.log('השורה אחרי הזרקה:', JSON.stringify(after));

  if (!after?.present) {
    report.fail('השורה מתעוררת אחרי ההזרקה', 'השורה נעלמה');
  } else if (after.unavailable === false && after.availability === 'measured') {
    report.pass('השורה מתעוררת אחרי ההזרקה', `נמדדה זמינה, ${after.fontFamily}`);
  } else {
    report.fail(
      'השורה מתעוררת אחרי ההזרקה',
      `הבייטים הוזרקו והשורה עדיין מסומנת — unavailable=${after.unavailable} ` +
        `availability=${after.availability}. המיזוג לא הורכב מחדש, או שהמדידה לא נשכחה.`,
    );
  }
} catch (error) {
  report.stuck('השער כולו', String(error && error.message ? error.message : error));
} finally {
  report.print();
  app.close();
}
