/** AutoCorrect text-safety regressions against the real editor and dialog. */
import { openApp, createReport } from './harness.mjs';
import { loadHebrewPreset } from './autocorrect-helpers.mjs';

const report = createReport('בטיחות תיקון אוטומטי', { strict: true });
const app = await openApp({ name: 'acorr-safety', port: Number(process.env.QA_PORT ?? 9635) });
const read = () => app.js(`(async () => {
  const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
  const target = (await doc.selection.current()).selectionTarget;
  const range = await doc.ranges.resolve({
    start: {kind:'point', point:{kind:'text', blockId:target.start.blockId, offset:0}},
    end: {kind:'point', point:{kind:'text', blockId:target.start.blockId, offset:target.start.offset}}
  });
  return range.preview.text;
})()`);
const check = (label, actual, expected) => actual === expected ? report.pass(label) : report.fail(label, JSON.stringify({ actual, expected }));

try {
  await app.caret(0);
  const ordinary = 'הילד למד תורה והמטוס טס ';
  await app.type(ordinary);
  await app.sleep(800);
  check('ברירת המחדל משמרת מילים תקינות', await read(), ordinary);

  await app.press('Enter', 'Enter', 13);
  await loadHebrewPreset(app);
  await app.caretPara(1);
  // Insert pre-existing text through the API, rather than counting it as typing.
  await app.js(`(async () => {
    const doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    await doc.insert({ type:'text', value:'למד ', target:(await doc.selection.current()).selectionTarget });
  })()`);
  await app.sleep(300);
  await app.press('End', 'End', 35);
  await app.type('(c)');
  await app.sleep(800);
  check('תיקון סמל אינו משנה את המילה הקיימת', await read(), 'למד ©');
  await app.type(' ');
  await app.sleep(800);
  check('המשך הקלדה אחרי קיצור אינו מתקן טקסט ישן', await read(), 'למד © ');

  await app.press('Enter', 'Enter', 13);
  await app.type('מלמד ' + 'x'.repeat(196), 5);
  await app.sleep(1000);
  const long = await read();
  long.startsWith('מלמד ') ? report.pass('חלון 200 תווים אינו חותך ומתקן מילה') : report.fail('חלון 200 תווים אינו חותך ומתקן מילה', long.slice(0, 12));

  await app.press('Enter', 'Enter', 13);
  if (!(await app.click('תיקון אוטומטי'))) throw new Error('AutoCorrect dialog did not open');
  await app.js(`(() => {
    const fields = document.querySelectorAll('.acorr-input');
    ['foo bar', 'baz'].forEach((value, i) => {
      fields[i].value = value;
      fields[i].dispatchEvent(new Event('input', { bubbles:true }));
    });
  })()`);
  await app.sleep(100);
  await app.js("document.querySelector('.acorr-table-actions .acorr-btn').scrollIntoView({ block:'nearest' })");
  await app.clickSel('.acorr-table-actions .acorr-btn');
  await app.clickSel('.acorr-btn-primary');
  await app.caretPara(3);
  await app.type('foo bar ');
  await app.sleep(800);
  check('ביטוי מהדיאלוג מוחל בהקלדה', await read(), 'baz ');
  await app.press('z', 'KeyZ', 90, 2);
  await app.sleep(800);
  check('Ctrl+Z מחזיר את הביטוי המקורי', await read(), 'foo bar ');
} finally {
  app.close();
  report.print();
}
