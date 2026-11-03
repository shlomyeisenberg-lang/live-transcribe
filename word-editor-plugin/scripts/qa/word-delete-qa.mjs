/**
 * שער: ‏`Ctrl+Backspace`/`Ctrl+Delete` מוחקים מילה גם אחרי פקודת רצועה.
 *
 * הדיווח: „קיצור מקשים קונטרול ומחיקה לא תמיד מוחק מילה”. אחרי „מספור” או
 * „תבליטים” ברצועה — או `lists.create` דרך ה-API — מחיקת המילה אינה עושה
 * דבר, ו-Backspace רגיל עובד (superdoc/docx-editor#4084). ההסבר והעקיפה
 * ב-`src/engine/word-delete-refresh.ts`.
 *
 * כל שורה בודקת את הטקסט של הפסקה **במנוע** (`ranges.resolve`) לפני ואחרי,
 * ושורת מעקב השינויים בודקת את ה-OOXML המיוצא (`w:delText`).
 *
 * שורות שנוספו אחרי סקירה: העיצוב הממתין (Ctrl+B בלי בחירה) שורד את
 * העקיפה; ביטול/הקלדה חזרה לאותו היסט אינם נחשבים הקשה מפגרת; ומקש מוחזק
 * אחרי „מספור” אינו כותב היסט ישן ואינו מוחק מילה אחרת.
 *
 *   npm run build && node scripts/qa/word-delete-qa.mjs   (QA_PORT דורס 9741)
 */
import { openApp, createReport } from './harness.mjs';

const CTRL = 2;
const report = createReport('Ctrl+Backspace / Ctrl+Delete — מחיקת מילה', { strict: true });
const app = await openApp({ name: 'word-delete', port: Number(process.env.QA_PORT ?? 9741) });

const readPara = () =>
  app
    .js(
      `(async function(){
    try {
      var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
      var info = await doc.selection.current();
      var seg = info && info.target && info.target.segments && info.target.segments[0];
      var st = info.selectionTarget || {};
      if (!seg) return JSON.stringify({ err: 'no block' });
      var r = await doc.ranges.resolve({
        start: { kind: 'point', point: { kind: 'text', blockId: seg.blockId, offset: 0 } },
        end: { kind: 'point', point: { kind: 'text', blockId: seg.blockId, offset: 400 } },
      });
      return JSON.stringify({ text: (r && r.preview && r.preview.text) || '', caret: st.start ? st.start.offset : null });
    } catch (e) { return JSON.stringify({ err: String((e && e.message) || e) }); }
  })()`,
    )
    .then(JSON.parse);

async function setText(value) {
  await app.js(`(async function(){
    var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    await doc.clearContent({});
    await new Promise(function(r){ setTimeout(r, 300); });
    await doc.insert({ type: 'text', value: ${JSON.stringify(value)} });
  })()`);
  await app.sleep(700);
  await app.caret(0);
  await app.sleep(300);
}

const MAKE_LIST = `(async function(){
  var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
  var info = await doc.selection.current();
  var seg = info.target.segments[0];
  await doc.lists.create({
    mode: 'fromParagraphs',
    target: { kind: 'block', nodeType: 'paragraph', nodeId: seg.blockId },
    kind: 'ordered',
    style: { version: 1, levels: [{ level: 0, numFmt: 'hebrew1', lvlText: '%1.', markerFont: '' }] },
    sequence: { mode: 'new', startAt: 1 },
  });
})()`;

const KEYS = {
  Home: ['Home', 'Home', 36, 0],
  End: ['End', 'End', 35, 0],
  '^BS': ['Backspace', 'Backspace', 8, CTRL],
  '^DEL': ['Delete', 'Delete', 46, CTRL],
};
const press = (name) => app.press(...KEYS[name]);

const focusInEditor = () => app.js('document.activeElement && document.activeElement.tagName');

/** תרחיש: טקסט, הכנה, ההקשה; מחזיר את הפסקה אחרי. */
async function scenario(name, { text = 'אחד שתיים שלוש', caretKey = 'End', setup, key = '^BS', expect: want }) {
  try {
    await setText(text);
    await press(caretKey);
    await app.sleep(300);
    if (setup) await setup();
    await app.sleep(900);
    const focus = await focusInEditor();
    const before = await readPara();
    await press(key);
    await app.sleep(1000);
    const after = await readPara();
    const detail = `${JSON.stringify(before.text)} → ${JSON.stringify(after.text)} (פוקוס: ${focus})`;
    if (after.text === want) report.pass(name, detail);
    else report.fail(name, detail, `צפוי ${JSON.stringify(want)}`);
  } catch (error) {
    report.fail(name, String(error?.message ?? error));
  }
}

const ribbon = (label) => async () => {
  const ok = await app.click(label, { after: 700 });
  if (!ok) throw new Error(`הפקד „${label}” לא נמצא ברצועה`);
};

try {
  // בקרה: אותו מקש בלי שום פקודה לפניו — מה שעבד גם לפני התיקון.
  await scenario('בקרה: לחיצה בפסקה, Ctrl+Backspace', { expect: 'אחד שתיים' });

  // הדיווח: פקודת רצועה, ומיד מחיקת מילה.
  await scenario('„מספור” ברצועה, Ctrl+Backspace', { setup: ribbon('מספור'), expect: 'אחד שתיים' });
  await scenario('„תבליטים” ברצועה, Ctrl+Backspace', { setup: ribbon('תבליטים'), expect: 'אחד שתיים' });
  await scenario('„מספור” ברצועה, Ctrl+Delete בתחילת הפסקה', {
    caretKey: 'Home',
    setup: ribbon('מספור'),
    key: '^DEL',
    expect: 'שתיים שלוש',
  });
  await scenario('lists.create דרך ה-API, Ctrl+Backspace', { setup: () => app.js(MAKE_LIST), expect: 'אחד שתיים' });

  // Ctrl+Delete משאיר את הסמן באותו היסט אחרי מחיקה תקינה. תצלום כזה
  // אינו מפגר. לחיצות נפרדות מוחקות כל מילה. במקש מוחזק שומרים על אותה
  // מדיניות בטיחות כמו במחיקה אחורה: ייתכן דילוג עד commit, אך לעולם לא
  // מחיקת מילה אחרת; ההקשה הנפרדת הבאה חייבת להמשיך ולא להיתקע לנצח.
  for (const [label, offset, gap] of [
    ['תחילת הפסקה, לחיצות נפרדות', 0, 1000],
    ['אמצע הפסקה, לחיצות נפרדות', 4, 1000],
    ['תחילת הפסקה, מקש מוחזק', 0, 33],
    ['אמצע הפסקה, מקש מוחזק', 4, 33],
  ]) {
    const name = `„מספור”, Ctrl+Delete חוזר: ${label}`;
    try {
      const text = 'אחד שתיים שלוש ארבע חמש שש';
      await setText(text);
      await app.js(`(async function(){
        var e = window.__otzariaEditor.superdoc.activeEditor;
        var s = e.host.readLiveSelectionSyncSnapshot().selectionTarget;
        var p = { ...s.start, offset: ${offset} };
        await e.authoring.setSelectionTarget({ target: { ...s, start: p, end: p }, focus: true });
      })()`);
      await app.sleep(400);
      await ribbon('מספור')();
      await app.sleep(900);
      const before = await readPara();
      if (before.caret !== offset) throw new Error(`הסמן לפני: ${before.caret}`);
      for (let i = 0; i < 3; i++) {
        await app.cdp.send('Input.dispatchKeyEvent', {
          type: 'rawKeyDown', key: 'Delete', code: 'Delete',
          windowsVirtualKeyCode: 46, nativeVirtualKeyCode: 46, modifiers: CTRL, autoRepeat: gap < 100 && i > 0,
        });
        await app.sleep(gap);
        if (gap >= 100) await app.cdp.send('Input.dispatchKeyEvent', {
          type: 'keyUp', key: 'Delete', code: 'Delete', windowsVirtualKeyCode: 46, nativeVirtualKeyCode: 46, modifiers: CTRL,
        });
      }
      await app.cdp.send('Input.dispatchKeyEvent', {
        type: 'keyUp', key: 'Delete', code: 'Delete', windowsVirtualKeyCode: 46, nativeVirtualKeyCode: 46, modifiers: CTRL,
      });
      await app.sleep(1000);
      const after = await readPara();
      const prefix = text.slice(0, offset);
      const words = text.slice(offset).split(' ');
      const remaining = (deleted) => prefix + words.slice(deleted).join(' ');
      let deleted = -1;
      for (let n = 1; n <= 3; n++) if (after.text === remaining(n)) deleted = n;
      if (after.caret !== offset || deleted < 1 || (gap >= 100 && deleted !== 3)) {
        report.fail(name, JSON.stringify(after), gap >= 100 ? `צפוי ${JSON.stringify(remaining(3))}@${offset}` : 'נמחק רק רצף מילים קדימה והסמן נשאר במקומו');
      } else if (gap < 100) {
        await press('^DEL');
        await app.sleep(1000);
        const next = await readPara();
        const detail = `${JSON.stringify(after)} נמחקו ${deleted}/3; הקשה נפרדת נוספת: ${JSON.stringify(next)}`;
        if (next.text === remaining(deleted + 1) && next.caret === offset) report.pass(name, detail);
        else report.fail(name, detail, `צפוי ${JSON.stringify(remaining(deleted + 1))}@${offset}`);
      } else report.pass(name, JSON.stringify(after));
    } catch (error) {
      report.fail(name, String(error?.message ?? error));
    }
  }

  // הקלדה אינה מרפאת את המצב במנוע, ולכן גם המשתמש שהמשיך לכתוב נפגע.
  await scenario('„מספור”, הקלדת מילה, Ctrl+Backspace', {
    setup: async () => {
      await ribbon('מספור')();
      await app.type(' ארבע', 120);
    },
    expect: 'אחד שתיים שלוש',
  });

  // העקיפה כותבת את הסמן מחדש — אסור שתחזיר אותו להיסט ישן בהקלדה מהירה.
  for (const gap of [40, 120]) {
    const name = `הקלדה בפער ${gap}ms ומיד Ctrl+Backspace — הסמן אינו נגרר אחורה`;
    try {
      await setText('אבג');
      await press('End');
      await app.sleep(300);
      await app.type(' שלום עולם', gap);
      await press('^BS');
      await app.sleep(1500);
      const after = await readPara();
      if (after.text === 'אבג שלום') report.pass(name, JSON.stringify(after));
      else report.fail(name, JSON.stringify(after), 'צפוי "אבג שלום"');
    } catch (error) {
      report.fail(name, String(error?.message ?? error));
    }
  }

  // עיצוב ממתין: Ctrl+B בלי בחירה, Ctrl+Backspace, הקלדה — הטקסט מודגש, כמו
  // בלי העקיפה. ‏setSelectionTarget מאפס אותו, והעקיפה משחזרת.
  for (const [label, setup] of [
    ['בלי פקודה', null],
    ['אחרי „מספור”', ribbon('מספור')],
  ]) {
    const name = `${label}: Ctrl+B, Ctrl+Backspace, הקלדה — הטקסט מודגש`;
    try {
      await setText('אחד שתיים שלוש');
      await press('End');
      await app.sleep(300);
      if (setup) await setup();
      await app.sleep(600);
      await app.press('b', 'KeyB', 66, CTRL);
      await app.sleep(500);
      await press('^BS');
      await app.sleep(900);
      await app.type(' זזז', 120);
      await app.sleep(1000);
      const xml = (await app.docx())['word/document.xml'];
      const run = (xml.match(/<w:r>(?:(?!<\/w:r>)[\s\S])*זזז[\s\S]*?<\/w:r>/) || [''])[0];
      const after = await readPara();
      const bold = /<w:b\/>|<w:bCs\/>/.test(run);
      const detail = `${JSON.stringify(after.text)} bold=${bold}`;
      if (after.text === 'אחד שתיים זזז' && bold) report.pass(name, detail);
      else report.fail(name, detail, 'צפוי "אחד שתיים זזז" מודגש');
    } catch (error) {
      report.fail(name, String(error?.message ?? error));
    }
  }

  // פסקה ממוספרת מתיישנת מחדש אחרי כל מחיקה; אחרי ביטול או הקלדה חזרה הסמן
  // שב לאותו היסט — וזו אינה הקשה מפגרת, ולכן חייבת להימחק.
  for (const [label, between, want] of [
    ['ביטול', () => app.press('z', 'KeyZ', 90, CTRL), 'אחד שתיים'],
    ['הקלדה חזרה', () => app.type(' שלוש', 120), 'אחד שתיים'],
  ]) {
    const name = `„מספור”, Ctrl+Backspace, ${label}, Ctrl+Backspace`;
    try {
      await setText('אחד שתיים שלוש');
      await press('End');
      await app.sleep(300);
      await ribbon('מספור')();
      await app.sleep(900);
      await press('^BS');
      await app.sleep(1000);
      const mid = await readPara();
      await between();
      await app.sleep(1000);
      const back = await readPara();
      await press('^BS');
      await app.sleep(1000);
      const after = await readPara();
      const detail = `${JSON.stringify(mid.text)} → ${JSON.stringify(back.text)}@${back.caret} → ${JSON.stringify(after.text)}`;
      if (mid.text === 'אחד שתיים' && back.text === 'אחד שתיים שלוש' && after.text === want) report.pass(name, detail);
      else report.fail(name, detail, `צפוי "אחד שתיים" → "אחד שתיים שלוש" → ${JSON.stringify(want)}`);
    } catch (error) {
      report.fail(name, String(error?.message ?? error));
    }
  }

  // מקש מוחזק (הקשות ב-33ms) מיד אחרי „מספור”, הסמן באמצע הפסקה. התצלום
  // הסינכרוני מפגר אחרי ההקשה הקודמת; כתיבה של ההיסט הישן מזיזה את הסמן
  // קדימה, אל מעבר למה שכבר נמחק. שני תנאים: נמחקו רק מילים צמודות משמאל
  // לסמן (אף מילה אחרת), וכל כתיבה של העקיפה נמצאת לפני הקודמת לה — היסט
  // שאינו יורד הוא היסט ישן. הקשות שאבדו מדווחות, ואינן כשל: בפסקה ממוספרת
  // המנוע עצמו מחמיץ אותן (ב-main רק הראשונה מוחקת).
  {
    const TEXT = 'א ב ג ד ה ו ז ח ט י כ ל מ נ ס ע';
    const MID = 19; // אחרי „י”
    const N = 6;
    // קצבים מהירים מעוררים את הפיגור; 33ms הוא קצב החזרה של Windows.
    for (const gap of [10, 15, 33]) {
      const name = `מקש מוחזק (${gap}ms) אחרי „מספור”, באמצע הפסקה — אין מחיקה של מילה אחרת ואין היסט ישן`;
      try {
        await setText(TEXT);
        await press('End');
        await app.sleep(300);
        for (let i = 0; i < TEXT.length - MID; i++) {
          // ‏ArrowRight בפסקה עברית = אחורה; לפני הפקודה, כי ניווט מרפא את המנוע.
          await app.press('ArrowRight', 'ArrowRight', 39, 0);
          await app.sleep(60);
        }
        await app.sleep(400);
        await ribbon('מספור')();
        await app.sleep(1200);
        const before = await readPara();
        if (before.caret !== MID) throw new Error(`הסמן לפני: ${before.caret}`);
        await app.js(`(function(){ window.__wdWrites = []; var a = window.__otzariaEditor.superdoc.activeEditor.authoring;
          if (!a.__wdOrig) { a.__wdOrig = a.setSelectionTarget; a.setSelectionTarget = function(input){ window.__wdWrites.push(input.target.start.offset); return a.__wdOrig.call(a, input); }; }
          return 1; })()`);
        for (let i = 0; i < N; i++) {
          await app.cdp.send('Input.dispatchKeyEvent', {
            type: 'rawKeyDown', key: 'Backspace', code: 'Backspace',
            windowsVirtualKeyCode: 8, nativeVirtualKeyCode: 8, modifiers: CTRL, autoRepeat: i > 0,
          });
          await app.sleep(gap);
        }
        await app.cdp.send('Input.dispatchKeyEvent', {
          type: 'keyUp', key: 'Backspace', code: 'Backspace', windowsVirtualKeyCode: 8, nativeVirtualKeyCode: 8, modifiers: CTRL,
        });
        await app.sleep(1500);
        const writes = JSON.parse(await app.js('JSON.stringify(window.__wdWrites)'));
        await app.js(`(function(){ var a = window.__otzariaEditor.superdoc.activeEditor.authoring; if (a.__wdOrig) { a.setSelectionTarget = a.__wdOrig; delete a.__wdOrig; } return 1; })()`);
        const after = await readPara();
        let deleted = -1;
        for (let k = 0; k <= N; k++) if (after.text === TEXT.slice(0, Math.max(0, MID - 2 * k)) + TEXT.slice(MID)) deleted = k;
        const descending = writes.every((offset, i) => i === 0 || offset < writes[i - 1]);
        const detail = `${JSON.stringify(after.text)} נמחקו ${deleted}/${N}; כתיבות ${JSON.stringify(writes)}`;
        if (deleted >= 1 && descending) report.pass(name, detail);
        else report.fail(name, detail, deleted < 0 ? 'נמחק טקסט שאינו המילים שלפני הסמן' : deleted < 1 ? 'לא נמחקה אף מילה' : 'נכתב היסט ישן');
      } catch (error) {
        report.fail(name, String(error?.message ?? error));
      }
    }
  }

  // בחירה פתוחה: המנוע מוחק את הבחירה, והעקיפה אינה נוגעת.
  {
    const name = 'בחירה פתוחה + Ctrl+Backspace מוחק את הבחירה בלבד';
    try {
      await setText('אחד שתיים שלוש');
      await press('End');
      await app.sleep(300);
      // ‏Shift+חץ שייך למנוע (ראו rtl-caret.ts); בפסקה עברית הוא מרחיב אחורה.
      await app.press('ArrowRight', 'ArrowRight', 39, 8);
      await app.sleep(200);
      await app.press('ArrowRight', 'ArrowRight', 39, 8);
      await app.sleep(600);
      const selected = await app.js(
        `(async function(){ var i = await window.__otzariaEditor.superdoc.activeEditor.doc.selection.current({ includeText: true }); return i.text; })()`,
      );
      if (selected !== 'וש') throw new Error(`הבחירה לפני ההקשה: ${JSON.stringify(selected)}`);
      await press('^BS');
      await app.sleep(1000);
      const after = await readPara();
      if (after.text === 'אחד שתיים של') report.pass(name, JSON.stringify(after));
      else report.fail(name, JSON.stringify(after), 'צפוי "אחד שתיים של"');
    } catch (error) {
      report.fail(name, String(error?.message ?? error));
    }
  }

  // מעקב שינויים: אותו מסלול, והמחיקה נרשמת כ-w:del.
  {
    const name = 'מעקב שינויים: lists.create, Ctrl+Backspace → w:delText';
    try {
      await setText('אחד שתיים שלוש');
      await app.tab('סקירה');
      if (!(await app.click('עקוב אחר שינויים', { after: 800 }))) throw new Error('„עקוב אחר שינויים” לא נמצא');
      const mode = await app.cmd('document-mode');
      if (mode.value !== 'suggesting') throw new Error(`document-mode=${mode.value}`);
      await app.tab('בית');
      await app.caret(0);
      await app.sleep(300);
      await press('End');
      await app.sleep(300);
      await app.js(MAKE_LIST);
      await app.sleep(900);
      await press('^BS');
      await app.sleep(1200);
      const xml = (await app.docx())['word/document.xml'];
      const deleted = [...xml.matchAll(/<w:delText(?:\s[^>]*)?>([^<]*)<\/w:delText>/g)].map((m) => m[1]).join('');
      if (deleted.includes('שלוש')) report.pass(name, `w:delText=${JSON.stringify(deleted)}`);
      else report.fail(name, `w:delText=${JSON.stringify(deleted)}`, 'המילה לא נרשמה כמחיקה');
    } catch (error) {
      report.fail(name, String(error?.message ?? error));
    }
  }

  const log = (await app.log()).filter((line) => /^(error|uncaught|rejected)/.test(line));
  if (log.length === 0) report.pass('אין שגיאות בדף');
  else report.fail('אין שגיאות בדף', log.join(' | ').slice(0, 600));
} finally {
  app.close();
}
process.exit(report.print() > 0 ? 1 : 0);
