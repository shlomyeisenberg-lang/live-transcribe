/**
 * Ctrl+Shift+E / Ctrl+Shift+8 במסמך: מה קורה לפסקה, ומה קורה לכיווניות.
 *
 * המנוע קושר את שני הצירופים בעצמו (מירכוז, רשימת תבליטים), ולכן הרשומות
 * שלנו — מעקב שינויים וסימני עיצוב — חייבות לרוץ לפניו. ובשלב ה-capture
 * הבליעה מסתירה את האות ממאזין כיווניות ב-bubble: מי שמשחרר את ה-Shift
 * לפני האות היה מקבל גם היפוך כיוון. נבדקים שני סדרי השחרור.
 *
 * QA_PORT=9377 node scripts/qa/ctrl-shift-chord-probe.mjs [shotDir prefix]
 */
import { writeFileSync } from 'node:fs';
import { openApp, sleep } from './harness.mjs';

const [shotDir, prefix = 'x'] = process.argv.slice(2);
const app = await openApp({ name: 'ctrlshift', port: Number(process.env.QA_PORT ?? 9377) });
const { cdp } = app;
const CTRL = 2;
const SHIFT = 8;

async function key(type, k, code, vk, modifiers, location = 0) {
  await cdp.send('Input.dispatchKeyEvent', { type, key: k, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers, location });
  await sleep(60);
}
async function shot(name) {
  if (!shotDir) return;
  const s = await cdp.send('Page.captureScreenshot', {});
  writeFileSync(`${shotDir}/${prefix}-${name}.png`, Buffer.from((s.result ?? s).data, 'base64'));
}
/** `rolled`: ה-Shift משוחרר לפני האות. */
async function chord(letter, { rolled = false } = {}) {
  await key('rawKeyDown', 'Control', 'ControlLeft', 17, CTRL, 1);
  await key('rawKeyDown', 'Shift', 'ShiftLeft', 16, CTRL | SHIFT, 1);
  await key('rawKeyDown', letter.key, letter.code, letter.vk, CTRL | SHIFT);
  if (rolled) {
    await key('keyUp', 'Shift', 'ShiftLeft', 16, CTRL, 1);
    await key('keyUp', letter.key, letter.code, letter.vk, CTRL);
  } else {
    await key('keyUp', letter.key, letter.code, letter.vk, CTRL | SHIFT);
    await key('keyUp', 'Shift', 'ShiftLeft', 16, CTRL, 1);
  }
  await key('keyUp', 'Control', 'ControlLeft', 17, 0, 1);
  await sleep(600);
}
async function plainShift(side) {
  const loc = side === 'ShiftRight' ? 2 : 1;
  await key('rawKeyDown', 'Control', 'ControlLeft', 17, CTRL, 1);
  await key('rawKeyDown', 'Shift', side, 16, CTRL | SHIFT, loc);
  await key('keyUp', 'Shift', side, 16, CTRL, loc);
  await key('keyUp', 'Control', 'ControlLeft', 17, 0, 1);
  await sleep(600);
}
async function ppr() {
  const xml = (await app.docx())['word/document.xml'];
  const p = (xml.match(/<w:p[ >][\s\S]*?<\/w:p>/) || [''])[0];
  return (p.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0];
}
const mode = () => app.cmd('document-mode').then((s) => s?.value ?? null);
const marks = () => app.cmd('formatting-marks').then((s) => s?.active ?? s?.value ?? null);
const E = { key: 'E', code: 'KeyE', vk: 69 };
const EIGHT = { key: '*', code: 'Digit8', vk: 56 };

try {
  await app.caret(0);
  await cdp.send('Input.insertText', { text: 'שלום עולם' });
  await sleep(800);
  await shot('0-before');
  const p0 = await ppr();
  console.log('start', p0, 'mode', JSON.stringify(await mode()));

  for (const rolled of [false, true]) {
    await chord(E, { rolled });
    const p1 = await ppr();
    console.log(`Ctrl+Shift+E rolled=${rolled}: pPr ${p1 === p0 ? 'unchanged' : 'CHANGED ' + p1}, mode=${JSON.stringify(await mode())}`);
    if (!rolled) await shot('1-after-ctrl-shift-e');
    // חוזרים למצב עריכה דרך אותו צירוף, כדי שהסיבוב הבא יימדד מאותה נקודה
    await chord(E);
    console.log(`  back: mode=${JSON.stringify(await mode())}, pPr ${(await ppr()) === p0 ? 'unchanged' : 'CHANGED'}`);
  }

  const m0 = await marks();
  await chord(EIGHT, { rolled: true });
  const p8 = await ppr();
  console.log(`Ctrl+Shift+8 rolled: pPr ${p8 === p0 ? 'unchanged' : 'CHANGED ' + p8}, marks ${JSON.stringify(m0)} -> ${JSON.stringify(await marks())}`);
  await chord(EIGHT);

  await plainShift('ShiftLeft');
  const pl = await ppr();
  await plainShift('ShiftRight');
  const pr = await ppr();
  console.log(`plain Ctrl+ShiftLeft: ${/<w:bidi\/>/.test(pl) ? 'still rtl (BAD)' : 'ltr (ok)'}; Ctrl+ShiftRight: ${/<w:bidi\/>/.test(pr) ? 'rtl (ok)' : 'ltr (BAD)'}`);
} finally {
  app.close();
}
