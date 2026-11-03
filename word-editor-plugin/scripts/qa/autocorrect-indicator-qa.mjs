/**
 * סימון התיקון האוטומטי — מול Chrome אמיתי. מקליד `teh ` ובודק איפה הסמן אחרי
 * ההחלפה, אם המנוע מחזיר מלבן סמן, ואם הסימון מצויר.
 *   node scripts/qa/autocorrect-indicator-qa.mjs
 */
import { openApp, createReport } from './harness.mjs';
import { loadHebrewPreset } from './autocorrect-helpers.mjs';

const report = createReport('סימון תיקון אוטומטי', { strict: true });
const app = await openApp({ name: 'acorr-ind', port: Number(process.env.QA_PORT ?? 9631) });

const state = () =>
  app
    .js(`(async function(){
    var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    var ui = window.__otzariaEditor.superdoc.ui;
    var info = await doc.selection.current();
    var t = info && info.selectionTarget;
    var seg = info && info.target && info.target.segments && info.target.segments[0];
    var r = seg ? await doc.ranges.resolve({
      start:{kind:'point',point:{kind:'text',blockId:seg.blockId,offset:0}},
      end:{kind:'point',point:{kind:'text',blockId:seg.blockId,offset:400}}}) : null;
    var rect = ui && ui.selection && ui.selection.getAnchorRect ? ui.selection.getAnchorRect({placement:'end'}) : 'noapi';
    var chip = document.querySelector('.otzaria-acorr-chip');
    return JSON.stringify({
      text: r && r.preview && r.preview.text,
      caret: t && t.start && t.start.offset,
      rect: rect,
      chipDisplay: chip && chip.style.display,
    });
  })()`)
    .then(JSON.parse);

try {
  await loadHebrewPreset(app);
  await app.js("window.otzariaAutocorrectDebug = true; window.__logs = []; var o = console.info; console.info = function(){ window.__logs.push(Array.prototype.slice.call(arguments).map(function(a){return typeof a==='string'?a:JSON.stringify(a)}).join(' ')); o.apply(console, arguments); }");
  await app.caret(0);
  await app.type('(c) ');
  await app.sleep(1500);
  console.log('לוגים:', await app.js('JSON.stringify(window.__logs)'));
  const s = await state();
  console.log('אחרי ההקלדה:', JSON.stringify(s));
  s.text && s.text.startsWith('©') ? report.pass('המילה הוחלפה', s.text) : report.fail('המילה הוחלפה', JSON.stringify(s));
  console.log('caret expected (עם הרווח) =', 4, ' בפועל =', s.caret);
  s.rect && s.rect !== 'noapi' ? report.pass('getAnchorRect מחזיר מלבן') : report.fail('getAnchorRect מחזיר מלבן', JSON.stringify(s.rect));

  // ריחוף: סורקים את השורה משמאל לסמן עד שהסימון מופיע
  if (s.rect && s.rect !== 'noapi') {
    const y = s.rect.top + s.rect.height / 2;
    let shown = false;
    for (let dx = 4; dx > -80 && !shown; dx -= 4) {
      await app.js(`(document.elementFromPoint(${s.rect.left + dx},${y}) || document).dispatchEvent(new PointerEvent('pointermove',{bubbles:true,clientX:${s.rect.left + dx},clientY:${y}}))`);
      await app.sleep(60);
      shown = (await state()).chipDisplay === 'block';
    }
    shown ? report.pass('הסימון מוצג בריחוף') : report.fail('הסימון מוצג בריחוף');
  }

  // מילה עברית מהטבלה
  await app.caret(0);
  await app.js("window.__logs.length = 0");
  await app.type(' אאכ ');
  await app.sleep(1500);
  console.log('לוגים עברית:', await app.js('JSON.stringify(window.__logs.filter(function(l){return /סימון|הוחלף/.test(l)}))'));
  const h = await state();
  console.log('עברית:', JSON.stringify(h));
  const yh = h.rect.top + h.rect.height / 2;
  let shownHe = false;
  for (let dx = 4; dx > -120 && !shownHe; dx -= 4) {
    await app.js(`(document.elementFromPoint(${h.rect.left + dx},${yh}) || document).dispatchEvent(new PointerEvent('pointermove',{bubbles:true,clientX:${h.rect.left + dx},clientY:${yh}}))`);
    await app.sleep(60);
    shownHe = (await state()).chipDisplay === 'block';
  }
  console.log('chip:', await app.js("JSON.stringify(document.querySelector('.otzaria-acorr-chip').getBoundingClientRect())"), 'caret.left=', h.rect.left);
  if (shownHe) {
    await app.js("document.querySelector('.otzaria-acorr-chip').click()");
    await app.sleep(300);
    const labels = await app.js("JSON.stringify(Array.from(document.querySelectorAll('.otzaria-acorr-menu button')).map(function(b){return b.textContent}))");
    console.log('תפריט:', labels);
    const look = await app.js(`JSON.stringify((function(){
      var root = getComputedStyle(document.documentElement);
      var m = getComputedStyle(document.querySelector('.otzaria-acorr-menu'));
      var b = getComputedStyle(document.querySelector('.otzaria-acorr-menu button'));
      var c = getComputedStyle(document.querySelector('.otzaria-acorr-chip'));
      return { surface: root.getPropertyValue('--color-surface').trim(), menuBg: m.backgroundColor, menuBorder: m.borderTopColor,
        outline: root.getPropertyValue('--color-outline').trim(), font: m.fontFamily.slice(0, 30), radius: m.borderRadius,
        btnH: b.height, chipBorder: c.borderTopColor, primary: root.getPropertyValue('--color-primary').trim(),
        icons: document.querySelectorAll('.otzaria-acorr-menu .svg-icon svg path').length };
    })())`);
    console.log('מראה:', look);
    await app.js("document.querySelector('.otzaria-acorr-menu button').click()");
    await app.sleep(1200);
    const back = await state();
    console.log('אחרי חזרה:', back.text);
    back.text.includes('אאכ') && !back.text.includes('אא"כ') ? report.pass('חזור למקור עובד') : report.fail('חזור למקור עובד', back.text);
  }
  shownHe ? report.pass('סימון בעברית מוצג בריחוף') : report.fail('סימון בעברית מוצג בריחוף');
} finally {
  report.print();
  app.close();
}
