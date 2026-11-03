/**
 * שער: צירופי `Ctrl+Shift` שהמנוע קושר — הטבלה ב-dispatch.ts עדיין נכונה,
 * והקיצורים שלנו עליהם רצים לפני המנוע.
 *
 * `ENGINE_CTRL_SHIFT` ב-src/ui/shortcuts/dispatch.ts היא טבלה קשיחה שנגזרה
 * מסקר (ctrl-shift-survey-probe.mjs). שדרוג מנוע שמוסיף צירוף — או מסיר —
 * היה עובר בשקט: רשומה שלנו על צירוף חדש שהמנוע בולע הייתה מתה ב-bubble בלי
 * שגיאה, בדיוק כמו Ctrl+Shift+E לפני התיקון. השער חוזר על הסקר ומשווה:
 *
 *   - כל צירוף שהמנוע קורא לו `preventDefault` (ומגיע ל-document), וכל צירוף
 *     שהמנתב שלנו בלע לפני המנוע, חייב להופיע בטבלה — ולהפך.
 *   - ‏Ctrl+Shift+E מפעיל מעקב שינויים ואינו נוגע ביישור ובכיוון, בשני סדרי
 *     השחרור; ‏Ctrl+Shift+8 מדליק סימני עיצוב ואינו פותח רשימה.
 *   - ‏Ctrl+Shift לבד: הימני מימין לשמאל, השמאלי משמאל לימין.
 *
 * מגבלה: האותיות שפותחות דיאלוג או טאב, וההדבקה (`SKIP`), אינן נסקרות, ולכן צירוף
 * מנוע חדש עליהן לא ייתפס כאן.
 * The table describes Windows Ctrl bindings. Set Windows navigator identity
 * before app/engine initialization so this survey is deterministic on macOS too.
 * This exercises the real engine in Chrome, not a native Windows WebView.
 *
 *   npm run build && node scripts/qa/ctrl-shift-engine-qa.mjs   (QA_PORT דורס 9395)
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { openApp, createReport, sleep, ROOT } from './harness.mjs';

const CTRL = 2;
const SHIFT = 8;
/** צירופים שהסקר עצמו הפיל בהם שגיאה — מדווחים בנפרד מהשורה „אין שגיאות”. */
const surveyErrors = [];
const report = createReport('Ctrl+Shift — הצירופים שהמנוע קושר', { strict: true });

/** הטבלה כפי שהיא כתובה במקור — השער משווה אליה, ולא לעותק משלו. */
function tableFromSource() {
  const src = readFileSync(join(ROOT, 'src', 'ui', 'shortcuts', 'dispatch.ts'), 'utf8');
  const block = src.match(/const ENGINE_CTRL_SHIFT[^=]*=\s*\{([\s\S]*?)\};/);
  if (!block) return null;
  return new Set([...block[1].matchAll(/(\w+)\s*:\s*'/g)].map((m) => m[1]));
}

// ‏`KeyV`: הדבקה כטקסט פשוט של המנוע. בלי לוח (headless) היא נכשלת בשגיאת מנוע
// („This edit couldn’t be completed”), והשורה „אין שגיאות” הייתה נצבעת בגללה.
const SKIP = new Set(['KeyQ', 'KeyW', 'KeyN', 'KeyT', 'KeyI', 'KeyJ', 'KeyO', 'KeyM', 'KeyB', 'KeyS', 'KeyV']);
const KEYS = [];
for (let c = 65; c <= 90; c++) KEYS.push({ key: String.fromCharCode(c), code: `Key${String.fromCharCode(c)}`, vk: c });
for (let d = 0; d <= 9; d++) KEYS.push({ key: String(d), code: `Digit${d}`, vk: 48 + d });
KEYS.push({ key: '>', code: 'Period', vk: 190 }, { key: '<', code: 'Comma', vk: 188 }, { key: '+', code: 'Equal', vk: 187 }, { key: '_', code: 'Minus', vk: 189 });

const app = await openApp({
  name: 'ctrl-shift-engine',
  port: Number(process.env.QA_PORT ?? 9395),
  extra: `<script>
    Object.defineProperty(navigator, 'platform', { value: 'Win32' });
    Object.defineProperty(navigator, 'userAgent', {
      value: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36',
    });
  </script>`,
});
const { cdp } = app;

async function key(type, k, code, vk, modifiers, location = 0, autoRepeat = false) {
  await cdp.send('Input.dispatchKeyEvent', { type, key: k, code, windowsVirtualKeyCode: vk, nativeVirtualKeyCode: vk, modifiers, location, autoRepeat });
  await sleep(60);
}
/** `rolled`: ה-Shift משוחרר לפני האות — הסדר שבו היפוך הכיוון הסתתר. */
async function chord(letter, { rolled = false, repeat = false } = {}) {
  await key('rawKeyDown', 'Control', 'ControlLeft', 17, CTRL, 1);
  await key('rawKeyDown', 'Shift', 'ShiftLeft', 16, CTRL | SHIFT, 1);
  await key('rawKeyDown', letter.key, letter.code, letter.vk, CTRL | SHIFT);
  if (repeat) {
    for (let i = 0; i < 3; i++) await key('rawKeyDown', letter.key, letter.code, letter.vk, CTRL | SHIFT, 0, true);
  }
  if (rolled) {
    await key('keyUp', 'Shift', 'ShiftLeft', 16, CTRL, 1);
    await key('keyUp', letter.key.toLowerCase(), letter.code, letter.vk, CTRL);
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
async function firstPara() {
  const xml = (await app.docx())['word/document.xml'];
  const body = (xml.match(/<w:body>([\s\S]*?)<w:sectPr/) || ['', ''])[1];
  const p = (body.match(/<w:p[ >][\s\S]*?<\/w:p>/) || [''])[0];
  const ppr = (p.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0];
  return {
    jc: (ppr.match(/<w:jc w:val="(\w+)"/) || [, null])[1],
    bidi: /<w:bidi\/>|<w:bidi w:val="(1|true|on)"/.test(ppr),
    list: /<w:numPr>/.test(ppr),
  };
}
const mode = () => app.cmd('document-mode').then((s) => s?.value ?? null);
const marks = () => app.cmd('formatting-marks').then((s) => s?.active ?? s?.value ?? null);
const E = { key: 'E', code: 'KeyE', vk: 69 };
const EIGHT = { key: '*', code: 'Digit8', vk: 56 };

try {
  await app.caret(0);
  await cdp.send('Input.insertText', { text: 'שלום עולם' });
  await sleep(800);

  // ── הסקר ──
  {
    const name = 'הטבלה ENGINE_CTRL_SHIFT תואמת את מה שהמנוע קושר';
    const table = tableFromSource();
    if (!table) {
      report.fail(name, 'ENGINE_CTRL_SHIFT לא נמצאה ב-dispatch.ts');
    } else {
      await app.js(`(() => {
        window.__cs = {};
        document.addEventListener('keydown', (e) => { if (e.ctrlKey && e.shiftKey && !/Shift|Control/.test(e.key)) window.__cs.doc = { prevented: e.defaultPrevented }; });
        return 1;
      })()`);
      const base = JSON.stringify(await firstPara()) + (await app.docx())['word/document.xml'].length;
      const owned = new Set();
      const errorsLen = async () => (await app.log()).filter((line) => /^(error|uncaught|rejected)/.test(line)).length;
      for (const k of KEYS) {
        if (SKIP.has(k.code)) continue;
        const errorsBefore = await errorsLen();
        await app.caret(0);
        await app.js('window.__cs = {}');
        await key('rawKeyDown', 'Control', 'ControlLeft', 17, CTRL, 1);
        await key('rawKeyDown', 'Shift', 'ShiftLeft', 16, CTRL | SHIFT, 1);
        await key('rawKeyDown', k.key, k.code, k.vk, CTRL | SHIFT);
        await key('keyUp', k.key, k.code, k.vk, CTRL | SHIFT);
        await key('keyUp', 'Shift', 'ShiftLeft', 16, CTRL, 1);
        await key('keyUp', 'Control', 'ControlLeft', 17, 0, 1);
        await sleep(350);
        const seen = JSON.parse(await app.js('JSON.stringify(window.__cs)'));
        // המנוע בלע (preventDefault), או שאנחנו בלענו לפניו (לא הגיע ל-document).
        if (!seen.doc || seen.doc.prevented) owned.add(k.code);
        if ((await errorsLen()) > errorsBefore) surveyErrors.push(k.code);
        if (await app.js(`!!document.querySelector('[role=dialog], dialog[open]')`)) await app.escape();
        // מחזירים את המצב: מעקב שינויים/סימני עיצוב שהודלקו, ושינוי במסמך
        if ((await mode()) === 'suggesting') await chord(E);
        if ((await marks()) === true) await chord(EIGHT);
        const snapshot = async () => JSON.stringify(await firstPara()) + (await app.docx())['word/document.xml'].length;
        for (let i = 0; i < 4 && (await snapshot()) !== base; i++) {
          await app.press('z', 'KeyZ', 90, CTRL);
          await sleep(400);
        }
      }
      const extra = [...owned].filter((c) => !table.has(c));
      const missing = [...table].filter((c) => !owned.has(c));
      if (extra.length === 0 && missing.length === 0) report.pass(name, [...owned].join(', '));
      else
        report.fail(
          name,
          `נבלעים ואינם בטבלה: ${extra.join(', ') || '—'}; בטבלה ואינם נבלעים: ${missing.join(', ') || '—'} — לעדכן את ENGINE_CTRL_SHIFT`,
        );
    }
  }

  // ── Ctrl+Shift+E / 8 ──
  await app.caret(0);
  await sleep(300);
  const p0 = await firstPara();
  for (const options of [{ rolled: false }, { rolled: true }, { rolled: true, repeat: true }]) {
    const { rolled, repeat } = options;
    const name = `Ctrl+Shift+E (${repeat ? 'מקש מוחזק' : rolled ? 'Shift משוחרר ראשון' : 'האות משוחררת ראשונה'}) — מעקב שינויים, בלי יישור ובלי כיוון`;
    await chord(E, options);
    const m = await mode();
    const p1 = await firstPara();
    const detail = `mode=${m} jc=${p1.jc} bidi=${p1.bidi}`;
    if (m === 'suggesting' && p1.jc === p0.jc && p1.bidi === p0.bidi) report.pass(name, detail);
    else report.fail(name, detail, `לפני: jc=${p0.jc} bidi=${p0.bidi}`);
    if (m === 'suggesting') await chord(E);
  }
  for (const repeat of [false, true]) {
    const name = `Ctrl+Shift+8${repeat ? ' (מקש מוחזק)' : ''} — סימני עיצוב, לא רשימת תבליטים`;
    const before = await marks();
    await chord(EIGHT, { rolled: true, repeat });
    const after = await marks();
    const p8 = await firstPara();
    const detail = `marks ${before}→${after} list=${p8.list} bidi=${p8.bidi}`;
    if (after === true && before !== true && !p8.list && p8.bidi === p0.bidi) report.pass(name, detail);
    else report.fail(name, detail);
    if (after === true) await chord(EIGHT);
  }
  {
    const name = 'Ctrl+Shift לבד: השמאלי משמאל לימין, הימני מימין לשמאל';
    await plainShift('ShiftLeft');
    const l = await firstPara();
    await plainShift('ShiftRight');
    const r = await firstPara();
    if (!l.bidi && r.bidi) report.pass(name);
    else report.fail(name, `אחרי השמאלי bidi=${l.bidi}, אחרי הימני bidi=${r.bidi}`);
  }

  const log = (await app.log()).filter((line) => /^(error|uncaught|rejected)/.test(line));
  console.log(`(שגיאות בזמן הסקר, אחרי: ${surveyErrors.join(', ') || 'אין'})`);
  if (log.length === 0) report.pass('אין שגיאות בדף');
  else report.fail('אין שגיאות בדף', log.join(' | ').slice(0, 600));
} finally {
  app.close();
}
process.exit(report.print() > 0 ? 1 : 0);
