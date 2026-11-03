/** סקר: אילו Ctrl+Shift+<מקש> המנוע בולע, ומה הם עושים למסמך. */
import { openApp, sleep } from './harness.mjs';

const app = await openApp({ name: 'csurvey', port: Number(process.env.QA_PORT ?? 9385) });
const { cdp } = app;
async function key(type, key, code, vk, modifiers, location = 0) {
  await cdp.send('Input.dispatchKeyEvent', { type, key, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers, location });
  await sleep(50);
}
const para = async () => {
  const xml = (await app.docx())['word/document.xml'];
  const p = (xml.match(/<w:body>([\s\S]*?)<w:sectPr/) || ['', xml])[1];
  return p.replace(/ w:rsid\w*="[^"]*"/g, '');
};

await app.js(`(() => {
  window.__s = {};
  document.addEventListener('keydown', (e) => { if (e.ctrlKey && e.shiftKey && !/Shift|Control/.test(e.key)) window.__s.doc = { prevented: e.defaultPrevented }; });
  window.addEventListener('keydown', (e) => { if (e.ctrlKey && e.shiftKey && !/Shift|Control/.test(e.key)) window.__s.win = true; }, true);
  return 1;
})()`);

const skip = new Set(['KeyQ', 'KeyW', 'KeyN', 'KeyT', 'KeyI', 'KeyJ', 'KeyO', 'KeyM', 'KeyB', 'KeyS']);
const keys = [];
for (let c = 65; c <= 90; c++) keys.push({ key: String.fromCharCode(c), code: `Key${String.fromCharCode(c)}`, vk: c });
for (let d = 0; d <= 9; d++) keys.push({ key: String(d), code: `Digit${d}`, vk: 48 + d });
keys.push({ key: '>', code: 'Period', vk: 190 }, { key: '<', code: 'Comma', vk: 188 }, { key: '+', code: 'Equal', vk: 187 }, { key: '_', code: 'Minus', vk: 189 });

try {
  await app.caret(0);
  await cdp.send('Input.insertText', { text: 'אבג דהו' });
  await sleep(800);
  const base = await para();
  for (const k of keys) {
    if (skip.has(k.code)) continue;
    await app.caret(0);
    await app.js('window.__s = {}');
    await key('rawKeyDown', 'Control', 'ControlLeft', 17, 2, 1);
    await key('rawKeyDown', 'Shift', 'ShiftLeft', 16, 10, 1);
    await key('rawKeyDown', k.key, k.code, k.vk, 10);
    await key('keyUp', k.key, k.code, k.vk, 10);
    await key('keyUp', 'Shift', 'ShiftLeft', 16, 2, 1);
    await key('keyUp', 'Control', 'ControlLeft', 17, 0, 1);
    await sleep(500);
    const s = JSON.parse(await app.js('JSON.stringify(window.__s)'));
    const after = await para();
    const changed = after !== base;
    const dialog = await app.js(`!!document.querySelector('[role=dialog], dialog[open]')`);
    if (changed || s.doc?.prevented || dialog || !s.doc) {
      console.log(`Ctrl+Shift+${k.key}: reachedDoc=${!!s.doc} enginePrevented=${!!s.doc?.prevented} docChanged=${changed} dialog=${dialog}${changed ? ' pPr=' + (after.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0] : ''}`);
    }
    if (dialog) await app.escape();
    if (changed) {
      // undo עד שחוזרים
      for (let i = 0; i < 4 && (await para()) !== base; i++) {
        await key('rawKeyDown', 'Control', 'ControlLeft', 17, 2, 1);
        await key('rawKeyDown', 'z', 'KeyZ', 90, 2);
        await key('keyUp', 'z', 'KeyZ', 90, 2);
        await key('keyUp', 'Control', 'ControlLeft', 17, 0, 1);
        await sleep(400);
      }
    }
  }
  console.log('done');
} finally {
  app.close();
}
