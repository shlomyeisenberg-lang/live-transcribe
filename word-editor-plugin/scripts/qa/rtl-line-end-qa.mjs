/**
 * שער רגרסיה ל„הנקודה עוברת לתחילת המסמך”.
 *
 * מה שדווח: משתמש הקליד משפט בעברית, לחץ `End` כדי לחזור לסוף השורה, הקליד
 * נקודה — ובקובץ הוורד השמור הנקודה יושבת בתחילת המסמך. השורש שנמדד
 * (superdoc 2.12.0): בפסקה עם `<w:bidi/>` המנוע מפרש `End` כ„תחילת השורה”,
 * כלומר בדיוק כמו `Home`, ולכן כל מה שמוקלד אחריו נכנס בהיסט 0.
 *
 * מה שהשער מודד, על ה-dist הארוז ובלחיצות והקשות אמיתיות:
 *
 *   1. `End` בשורה עברית גולשת מגיע לסוף השורה — היעד נגזר מהמסמך המצויר
 *      (`data-pm-end` של השורה פחות `data-pm-start` של הפסקה) ולא ממספר קשיח.
 *   2. הנקודה שמוקלדת אחרי `End` יושבת בסוף הטקסט ב-`word/document.xml`,
 *      ולא בראשו. זו הצורה שהמשתמש דיווח עליה.
 *   3. `End` פעמיים אינו מטייל לשורה הבאה.
 *   3א. המסלול שדווח מקצה לקצה על פסקה קצרה: `Home`, `End`, נקודה — והפסקה
 *      נגמרת בנקודה במקום להיפתח בה.
 *   3ב. `Home` ואז `End` בתחילת שורה גולשת שנייה מגיעים לסופה.
 *   4. `Shift+End` בוחר קדימה, מהסמן ועד סוף השורה.
 *   5. באנגלית `End` ממשיך לעשות מה שהוא עשה — היירוט אינו נוגע בשורה שאינה
 *      עברית.
 *   6. `Ctrl+End` עדיין מגיע לסוף המסמך.
 *
 * שימוש:  CHROME=<path> node scripts/qa/rtl-line-end-qa.mjs   (QA_PORT דורס 9371)
 */
import { openApp, createReport } from './harness.mjs';
import { buildDocx, p } from './docx-fixtures.mjs';

const RTL = '<w:bidi/>';
/** פסקה עברית שגולשת לשתי שורות מצוירות. */
const HEB =
  'אמר רבי חנינא לא מצאתי לגוף טוב אלא שתיקה וכל המרבה דברים מביא חטא והנה זה כלל גדול בתורה ואין לך דבר שעומד בפני התשובה ולכן צריך אדם לפשפש במעשיו';
/** פסקה קצרה בת שורה אחת — המסמך שהמשתמש תיאר. */
const SHORT = 'משפט קצר בעברית';
const ENG = 'the quick brown fox jumps over the lazy dog and keeps running across the wide field';

const report = createReport('„הנקודה עוברת לתחילת המסמך” — End בשורה עברית', { strict: true });
const app = await openApp({ name: 'rtl-line-end', port: Number(process.env.QA_PORT ?? 9371) });

async function openDocx(buffer, name) {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(buffer).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-${name}',url:${JSON.stringify(dataUrl)},name:${JSON.stringify(name + '.docx')},size:${buffer.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  await app.click('פתח קובץ', { after: 2500 });
  // הכפתור ברצועה פותח את מסך „פתח מסמך”; הבורר יושב מאחורי „עיון בקבצים…”.
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.clickSel('.open-browse', 0, { after: 12000 });
  return app.js("document.querySelector('.doc-title-input')?.value");
}

/** השורות המצוירות, עם טווח ה-pm וההיסט שלהן בתוך הפסקה. */
const paintedLines = () =>
  app
    .js(
      `(function(){
        var out = [];
        Array.prototype.forEach.call(document.querySelectorAll('[data-source-node-id][data-pm-start]'), function(frag){
          var base = Number(frag.getAttribute('data-pm-start'));
          Array.prototype.forEach.call(frag.children, function(line){
            var s = Number(line.getAttribute('data-pm-start')), e = Number(line.getAttribute('data-pm-end'));
            if (!isFinite(s) || !isFinite(e)) return;
            var r = line.getBoundingClientRect();
            out.push({ blockId: frag.getAttribute('data-source-node-id'), base: base,
              startOffset: s - base, endOffset: e - base, rtl: line.getAttribute('dir') === 'rtl',
              text: (line.textContent || '').slice(0, 18),
              x: Math.round((r.left + r.right) / 2), y: Math.round((r.top + r.bottom) / 2) });
          });
        });
        return JSON.stringify(out);
      })()`,
    )
    .then(JSON.parse);

const selection = () =>
  app
    .js(
      `(async function(){
        var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
        var info = await doc.selection.current();
        var t = info && info.selectionTarget;
        return JSON.stringify({ empty: !!(info && info.empty), block: t && t.start && t.start.blockId,
          start: t && t.start && t.start.offset, end: t && t.end && t.end.offset });
      })()`,
    )
    .then(JSON.parse);

/** טקסט הפסקה כפי שהוא ב-docx המיוצא — ההוכחה שהמשתמש רואה. */
async function paragraphText(index) {
  const xml = (await app.docx())['word/document.xml'] || '';
  const paragraphs = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g) || [];
  return (paragraphs[index] || '').replace(/<[^>]+>/g, '');
}

/**
 * הגאומטריה המצוירת, מחדש.
 *
 * **למה זה קיים.** השער לכד את השורות פעם אחת והשתמש בקואורדינטות שלהן גם
 * אחרי ש-`app.type('.')` שינה את המסמך. שלוש הטענות שנשענו על לכידה ישנה הן
 * בדיוק שלוש הטענות שהמתחזק מדד אדומות — `Home→End` בשורה השנייה, `Shift+End`
 * ו-`End` בשורה אנגלית — בעוד שש הטענות שלכדו מחדש עברו אצלו. ההתאמה היא
 * שלוש מתוך שלוש, ולכן זו אינה השערה על מהירות המכונה אלא על מה שהשער עושה.
 *
 * תו שנוסף יכול להזיז פריסה, ואז `clickAt(line.x, line.y)` נוחת על שורה אחרת
 * ו-`line.endOffset` שייך לשורה שכבר אינה שם. אצלנו הפסקה הקצרה אינה גולשת
 * מהוספת נקודה ולכן הפריסה אינה זזה, וזו הסיבה שכל התשעה עוברים כאן על אותו
 * קומיט שאצלו אדום: ההבדל הוא במטריקות הגופן, לא בקוד.
 *
 * **נמדד שזה ולא תזמון:** הרצה עם המתנה אפס — קריאת הסמן מיד אחרי ההקשה, בלי
 * שום `sleep` — נתנה כאן 9 מתוך 9. לו התזמון היה הגורם, היא הייתה מאדימה.
 * לכן לא נוספה כאן המתנה-על-תנאי: אין לה טריגר נמדד.
 */
async function rtlLinesNow() {
  return (await paintedLines()).filter((l) => l.rtl);
}

async function caretAt(line) {
  await app.clickAt(line.x, line.y);
  await app.sleep(500);
  return selection();
}

try {
  const opened = await openDocx(
    buildDocx({ body: p(HEB, RTL) + p(SHORT, RTL) + p(ENG) }),
    'rtl-line-end',
  );
  if (opened !== 'rtl-line-end') {
    report.fail('פתיחת המסמך', `שם המסמך אחרי הפתיחה: ${opened}`);
  } else {
    const lines = await paintedLines();
    const hebrew = lines.filter((l) => l.rtl);
    const latin = lines.filter((l) => !l.rtl);
    console.log('שורות מצוירות:', JSON.stringify(lines.map(({ text, rtl, startOffset, endOffset }) => ({ text, rtl, startOffset, endOffset }))));

    if (hebrew.length < 2) {
      report.fail('הפסקה העברית גולשת לשתי שורות', `נמצאו ${hebrew.length}`);
    } else {
      /* 1. End מגיע לסוף השורה */
      const first = hebrew[0];
      const before = await caretAt(first);

      /* 1א. ה-`story` בתצלום הסינכרוני — זה מה שנכתב בחזרה למנוע.
         המודול אינו ממציא ברירת מחדל, ולכן נמדד כאן שהשדה אכן מגיע. */
      const snap = JSON.parse(
        await app.js(
          `(function(){
            var h = window.__otzariaEditor.superdoc.activeEditor.host;
            var s = h.readLiveSelectionSyncSnapshot();
            var t = s && s.selectionTarget;
            return JSON.stringify({
              onTarget: !!(t && t.story), onEnd: !!(t && t.end && t.end.story),
              story: t && (t.story || (t.end && t.end.story)) });
          })()`,
        ),
      );
      if (snap.onTarget || snap.onEnd)
        report.pass('ה-story בתצלום', `target=${snap.onTarget} end=${snap.onEnd} — ${JSON.stringify(snap.story)}`);
      else report.fail('ה-story בתצלום', 'התצלום לא נשא story באף אחד מהשניים — הכתיבה תצא בלעדיו');

      await app.press('End', 'End', 35);
      await app.sleep(400);
      const after = await selection();
      if (after.start === first.endOffset)
        report.pass('End בשורה עברית', `${before.start} → ${after.start} (סוף השורה)`);
      else
        report.fail(
          'End בשורה עברית',
          `${before.start} → ${after.start}, וסוף השורה הוא ${first.endOffset}` +
            (after.start === first.startOffset ? ' — זו בדיוק התקלה שדווחה' : ''),
        );

      /* 2. End פעמיים אינו מטייל לשורה הבאה */
      await app.press('End', 'End', 35);
      await app.sleep(400);
      const twice = await selection();
      if (twice.start === first.endOffset) report.pass('End פעמיים', `נשאר על ${twice.start}`);
      else report.fail('End פעמיים', `זז ל-${twice.start} במקום להישאר על ${first.endOffset}`);

      /* 3. הנקודה שמוקלדת אחריו יושבת בסוף, לא בראש */
      await app.type('.');
      await app.sleep(800);
      const text = await paragraphText(0);
      if (text.startsWith('.'))
        report.fail('נקודה אחרי End', `הפסקה נפתחת בנקודה: ${JSON.stringify(text.slice(0, 24))}`);
      else if (text.slice(0, first.endOffset + 1).endsWith('.'))
        report.pass('נקודה אחרי End', `נכנסה בהיסט ${first.endOffset}, בסוף השורה`);
      else
        report.fail('נקודה אחרי End', `לא נמצאה בסוף השורה: ${JSON.stringify(text.slice(0, 40))}`);

      /* 4. המסלול שדווח, מקצה לקצה, על פסקה קצרה בת שורה אחת */
      const fresh = (await paintedLines()).filter((l) => l.rtl);
      const short = fresh.find((l) => l.blockId !== first.blockId);
      // התרחיש מניח פסקה בת שורה אחת: `End` בפסקה גולשת מגיע לסוף השורה ולא
      // לסוף הפסקה, ואז הנקודה נוחתת באמצע בצדק. ההנחה נאמרת כאן במפורש, כדי
      // שמכונה שמטריקות הגופן שלה מגלישות את הפסקה תקבל אבחנה ולא „שבור” סתום.
      const shortLines = fresh.filter((l) => short && l.blockId === short.blockId).length;
      if (!short) {
        report.fail('המסלול שדווח', 'לא נמצאה פסקה עברית קצרה');
      } else if (shortLines !== 1) {
        report.fail(
          'המסלול שדווח',
          `הפסקה הקצרה גולשת ל-${shortLines} שורות על המכונה הזאת, והתרחיש מניח אחת — ` +
            'הפיקסטורה צריכה להתקצר, לא הקוד להשתנות',
        );
      } else {
        await caretAt(short);
        await app.press('Home', 'Home', 36);
        await app.sleep(300);
        await app.press('End', 'End', 35);
        await app.sleep(300);
        await app.type('.');
        await app.sleep(800);
        const text = await paragraphText(1);
        if (text.endsWith('.') && !text.startsWith('.'))
          report.pass('המסלול שדווח', `„${text}”`);
        else report.fail('המסלול שדווח', `הפסקה יצאה „${text}”`);
      }

      /* 4א. Home בתחילת שורה גולשת שנייה, ואז End, מגיע לסופה. */
      // לכידה מחדש: הנקודה שהוקלדה למעלה יכולה להזיז את הפריסה. ראו rtlLinesNow.
      const afterDot = await rtlLinesNow();
      const wrapped = afterDot.filter((l) => l.blockId === first.blockId);
      const second = wrapped[1];
      if (!second) {
        report.fail('Home ואז End בשורה שנייה', 'לא נמצאה שורה גולשת שנייה');
      } else {
        await caretAt(second);
        await app.press('Home', 'Home', 36);
        await app.sleep(300);
        await app.press('End', 'End', 35);
        await app.sleep(400);
        const after = await selection();
        if (after.start === second.endOffset)
          report.pass('Home ואז End בשורה שנייה', `${after.start} (סוף השורה)`);
        else report.fail('Home ואז End בשורה שנייה', `${after.start} במקום ${second.endOffset}`);
      }

      /* 5. Shift+End בוחר קדימה */
      // לכידה מחדש מאותו טעם — הטענה משווה ל-endOffset של השורה שנלחצה.
      const forShift = (await rtlLinesNow())[0];
      const anchor = await caretAt(forShift);
      await app.press('End', 'End', 35, 8);
      await app.sleep(400);
      const range = await selection();
      if (!range.empty && range.start === anchor.start && range.end === forShift.endOffset)
        report.pass('Shift+End', `${range.start}..${range.end}`);
      else
        report.fail(
          'Shift+End',
          `${range.start}..${range.end}, והעוגן היה ${anchor.start} וסוף השורה ${forShift.endOffset}`,
        );
    }

    /* 6. אנגלית — היירוט אינו נוגע */
    if (!latin.length) {
      report.fail('End בשורה אנגלית', 'לא נמצאה שורה שאינה עברית');
    } else {
      // הלכידה הראשונית קדמה לשתי ההקלדות, ולכן נלכדת מחדש.
      const line = (await paintedLines()).filter((l) => !l.rtl)[0];
      if (!line) throw new Error('שורה לטינית נעלמה אחרי ההקלדות');
      await caretAt(line);
      await app.press('End', 'End', 35);
      await app.sleep(400);
      const after = await selection();
      if (after.start === line.endOffset) report.pass('End בשורה אנגלית', `הגיע ל-${after.start}`);
      else report.fail('End בשורה אנגלית', `${after.start} במקום ${line.endOffset}`);
    }

    /* 7. Ctrl+End עדיין סוף המסמך */
    if (hebrew.length) {
      await caretAt(hebrew[0]);
      await app.press('End', 'End', 35, 2);
      await app.sleep(600);
      const last = (await paintedLines()).at(-1);
      const after = await selection();
      if (after.block === last.blockId && after.start === last.endOffset)
        report.pass('Ctrl+End', `סוף המסמך: ${after.block}:${after.start}`);
      else
        report.fail(
          'Ctrl+End',
          `${after.block}:${after.start}, וסוף המסמך הוא ${last.blockId}:${last.endOffset}`,
        );
    }
  }
} finally {
  app.close();
}

process.exit(report.print() > 0 ? 1 : 0);
