/**
 * טבלת „החלף בעת ההקלדה” עם רצפי סימנים שהמשתמש מגדיר — `/` ← `"`, `--` ← `:`,
 * `==` ← `?` — מול העורך והדיאלוג האמיתיים.
 *
 * מה נמדד: ההחלפה מיידית גם בתוך מילה; קצב אנושי וקצב מהיר נותנים אותה תוצאה;
 * דפוס קצר שהוא תחילתו של דפוס ארוך (`--` מול `-->`) ממתין לתו הבא ואינו תלוי
 * בקצב. אחרי עצירה קצרה Enter מחליף אותו; Enter מהיר שומר את המקור ללא ניחוש,
 * ו-Enter שמבוטל אינו גורר החלפה מאוחרת. Ctrl+Z מחזיר את המקור בצעד אחד; כתובת
 * אינה נפגעת; ההחלפה יורשת את עיצוב הריצה.
 *
 * `QA_SHOTS=<תיקייה>` שומר צילום של הדיאלוג.
 */
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { openApp, createReport } from './harness.mjs';

const report = createReport('טבלת החלפה של סימנים', { strict: true });
const app = await openApp({ name: 'acorr-replace-table', port: Number(process.env.QA_PORT ?? 9641) });
const read = () => app.js(`(async () => {
  const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
  const target = (await doc.selection.current()).selectionTarget;
  const range = await doc.ranges.resolve({
    start: {kind:'point', point:{kind:'text', blockId:target.start.blockId, offset:0}},
    end: {kind:'point', point:{kind:'text', blockId:target.start.blockId, offset:target.start.offset}}
  });
  return range.preview.text;
})()`);
const check = (label, actual, expected) =>
  actual === expected ? report.pass(label) : report.fail(label, JSON.stringify({ actual, expected }));
const undo = () => app.press('z', 'KeyZ', 90, 2);
const newPara = async () => {
  await app.press('Enter', 'Enter', 13);
  await app.sleep(300);
};

async function shot(name) {
  const dir = process.env.QA_SHOTS;
  if (!dir) return;
  const rect = await app.js(`(() => {
    const r = document.querySelector('.acorr-dialog').getBoundingClientRect();
    return { x: r.x, y: r.y, width: r.width, height: r.height };
  })()`);
  const response = await app.cdp.send('Page.captureScreenshot', { format: 'png', clip: { ...rect, scale: 1 } });
  const data = response?.result?.data;
  if (!data) throw new Error('הצילום נכשל');
  mkdirSync(dir, { recursive: true });
  writeFileSync(join(dir, `${name}.png`), Buffer.from(data, 'base64'));
}

async function addEntry(from, to) {
  await app.js(`(() => {
    const fields = document.querySelectorAll('.acorr-input');
    [${JSON.stringify(from)}, ${JSON.stringify(to)}].forEach((value, i) => {
      fields[i].value = value;
      fields[i].dispatchEvent(new Event('input', { bubbles:true }));
    });
  })()`);
  await app.sleep(100);
  await app.js("document.querySelector('.acorr-table-actions .acorr-btn').scrollIntoView({ block:'nearest' })");
  await app.clickSel('.acorr-table-actions .acorr-btn');
}

try {
  await app.caret(0);
  await app.tab('סקירה');
  if (!(await app.click('תיקון אוטומטי'))) throw new Error('AutoCorrect dialog did not open');
  await app.sleep(300);
  await addEntry('/', '"');
  await addEntry('--', ':');
  await addEntry('==', '?');
  await shot('dialog');
  const modes = await app.js(`Array.from(document.querySelectorAll('.acorr-row'))
    .filter(row => ['/', '--', '1/2'].includes(row.querySelector('.acorr-cell')?.textContent))
    .map(row => row.querySelector('.acorr-mode')?.textContent ?? '')`);
  check('הטבלה מסמנת את הדפוסים שמוחלפים מיד', JSON.stringify(modes), JSON.stringify(['', 'מיד', 'מיד']));
  await app.clickSel('.acorr-btn-primary');
  await app.caretPara(0);

  // 1. בתוך מילה, בקצב אנושי ובקצב מהיר.
  await app.type('רש/י ');
  await app.sleep(800);
  check('/ ← " בתוך מילה, בקצב אנושי', await read(), 'רש"י ');
  await newPara();
  await app.type('רמב/ם ', 5);
  await app.sleep(800);
  check('/ ← " בתוך מילה, בקצב מהיר', await read(), 'רמב"ם ');

  // 2. Ctrl+Z מחזיר את מה שהוקלד, ואינו מוחק אותו.
  await newPara();
  await app.type('א/');
  await app.sleep(800);
  check('ההחלפה המיידית מוחלת לפני כל תו נוסף', await read(), 'א"');
  await undo();
  await app.sleep(600);
  check('Ctrl+Z אחד מחזיר את / (צעד אחד)', await read(), 'א/');
  await app.type('ב');
  await app.sleep(800);
  check('ההקלדה שאחרי הביטול אינה מחליפה שוב', await read(), 'א/ב');

  // 3. דפוס קצר שהוא תחילתו של דפוס ארוך אינו תלוי בקצב.
  await newPara();
  await app.type('א--');
  await app.sleep(800);
  check('-- ממתין כשהתו הבא יכול להשלים -->', await read(), 'א--');
  await app.type('ב');
  await app.sleep(800);
  check('-- ← : אחרי שהתו הבא הוקלד', await read(), 'א:ב');
  await newPara();
  await app.type('א--');
  await app.sleep(800);
  await app.type('>');
  await app.sleep(800);
  check('--> בהקלדה איטית עדיין חץ', await read(), 'א→');
  await newPara();
  await app.type('מה==', 5);
  await app.sleep(800);
  await app.type(' ');
  await app.sleep(800);
  check('== ← ? בקצב מהיר', await read(), 'מה? ');

  // 4. כתובת אינה נפגעת.
  await newPara();
  await app.type('http://a.b/c ');
  await app.sleep(800);
  check('/ בכתובת נשאר', await read(), 'http://a.b/c ');

  // 5. דפוס שממתין בסוף שורה: Enter מכריע שלא יגדל, ומחליף אותו בשורה שהסתיימה.
  //    הסמן נשאר בשורה החדשה, וההקלדה בה ממשיכה.
  await newPara();
  await app.type('ואלו הן--');
  await app.sleep(800);
  await newPara();
  await app.sleep(1200);
  await app.type('א');
  await app.sleep(800);
  const texts = await app.js(
    "JSON.stringify(Array.from(document.querySelectorAll('.superdoc-fragment[data-source-node-id]')).map(f => f.textContent))",
  ).then(JSON.parse);
  check('-- ← : כש-Enter סיים את השורה', texts.includes('ואלו הן:') ? 'ואלו הן:' : JSON.stringify(texts.slice(-3)), 'ואלו הן:');
  check('אחרי ההחלפה הסמן נשאר בשורה החדשה', await read(), 'א');

  // Cancellation is required even when the real engine's asynchronous range API is used.
  await app.type('undo--', 5);
  await app.sleep(700);
  await app.press('Enter', 'Enter', 13);
  await undo();
  await app.sleep(900);
  check('Enter ואז Ctrl+Z אינם משאירים החלפה מאוחרת', await read(), 'אundo--');

  // End-of-line settlement deliberately requires a previously verified suffix.
  // Fast Enter has no such anchor; preserving the typed source is the safe fallback.
  await newPara();
  await app.type('fast--', 5);
  await app.press('Enter', 'Enter', 13);
  await app.sleep(700);
  const fastTexts = await app.js(
    "Array.from(document.querySelectorAll('.superdoc-fragment[data-source-node-id]')).map(f => f.textContent)",
  );
  check('Enter מהיר שומר טקסט ללא עוגן מאומת', fastTexts.includes('fast--') ? 'fast--' : JSON.stringify(fastTexts), 'fast--');

  // 6. עיצוב הריצה עובר להחלפה.
  await newPara();
  await app.press('b', 'KeyB', 66, 2);
  await app.type('ש/ב');
  await app.sleep(800);
  check('ההחלפה בתוך ריצה מודגשת', await read(), 'ש"ב');
  const files = await app.docx();
  const xml = files['word/document.xml'] ?? '';
  // הפסקה שהטקסט שלה ש"ב; כל ריצה בה שנושאת טקסט — מודגשת.
  const textOf = (body) => [...body.matchAll(/<w:t\b[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1].replaceAll('&quot;', '"')).join('');
  const para = [...xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)].map((m) => m[0]).find((body) => textOf(body) === 'ש"ב');
  const runs = para ? [...para.matchAll(/<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g)].map((m) => m[1]).filter((body) => textOf(body) !== '') : [];
  const unbold = runs.filter((body) => !/<w:b\/>|<w:b w:val="(1|true|on)"\/>/.test(body));
  runs.length > 0 && unbold.length === 0
    ? report.pass('הגרשיים יורשים את ההדגשה של הריצה', `${runs.length} ריצות`)
    : report.fail('הגרשיים יורשים את ההדגשה של הריצה', para ? unbold.join(' | ').slice(0, 600) : 'לא נמצאה פסקה ש"ב');
} catch (error) {
  report.fail('הרצת השער', String(error?.stack ?? error));
} finally {
  app.close();
  process.exit(report.print() > 0 ? 1 : 0);
}
