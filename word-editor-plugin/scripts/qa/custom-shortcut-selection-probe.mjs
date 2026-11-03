/**
 * איזה קריאה מפרידה בין „יש בחירה במסמך” לבין „אין לאן להחיל”.
 *
 * `ui.selection.get()` נמדד כלא-מפריד: הוא מחזיר `{ready, empty:true}` גם על
 * סמן מכווץ וגם כשאין סמן כלל, והוא גם **מפגר** — מיד אחרי בחירה במקלדת הוא
 * מדווח `stale`. כאן נמדדת מולו הקריאה של ה-Document API,
 * `doc.selection.current()`, שהיא מה ש-engine/doc-selection.ts קיים בשבילו.
 *
 *   node scripts/qa/custom-shortcut-selection-probe.mjs
 */
import { openApp, sleep } from './harness.mjs';

const SHIFT = 8;

const CURRENT = `(async () => {
  var sd = window.__qa.sd();
  var doc = sd && sd.activeEditor && sd.activeEditor.doc;
  if (!doc || !doc.selection || !doc.selection.current) return JSON.stringify({ error: 'no-doc-api' });
  try {
    var info = await doc.selection.current({ includeText: true });
    return JSON.stringify({
      empty: info && info.empty,
      text: info && typeof info.text === 'string' ? info.text.slice(0, 20) : null,
      hasTarget: !!(info && info.target),
      kind: info && info.target && info.target.kind,
      segments: info && info.target && Array.isArray(info.target.segments) ? info.target.segments.length : null,
    });
  } catch (e) {
    return JSON.stringify({ error: String(e && e.message) });
  }
})()`;

const log = async (label, app) =>
  console.log(
    `${label}\n   ui.selection.get  : ${JSON.stringify(await app.selection())}` +
      `\n   doc.selection.cur : ${await app.js(CURRENT)}`,
  );

const app = await openApp({ name: 'cssel', port: 9665 });
try {
  await log('א. אחרי עלייה, בלי נגיעה', app);

  await app.tab('הוספה');
  await sleep(500);
  await log('ב. מסמך ריק, המיקוד על לשונית ברצועה', app);

  await app.tab('בית');
  await app.caret(0);
  await sleep(600);
  await log('ג. סמן מכווץ במסמך', app);

  await app.type('abcd', 60);
  await sleep(600);
  await app.press('Home', 'Home', 36);
  await sleep(200);
  for (let i = 0; i < 4; i += 1) {
    await app.press('ArrowRight', 'ArrowRight', 39, SHIFT);
    await sleep(60);
  }
  await sleep(300);
  await log('ד. מיד אחרי בחירת טווח במקלדת', app);

  await app.js('document.activeElement && document.activeElement.blur(); document.body.focus();');
  await sleep(500);
  await log('ה. אותה בחירה, המיקוד ב-body', app);

  /* המצב של השער: סמן מכווץ בסוף הטקסט, המיקוד מחוץ למסמך. */
  await app.caret(0);
  await sleep(400);
  await app.press('End', 'End', 35);
  await sleep(300);
  await app.js('document.activeElement && document.activeElement.blur(); document.body.focus();');
  await sleep(500);
  await log('ו. סמן מכווץ בטקסט, המיקוד ב-body', app);
} finally {
  app.close();
}
