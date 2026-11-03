/**
 * גשש: מי מטפל ב-Backspace על פריט רשימה ריק — המנוע, או אנחנו?
 *
 * הרשימה כאן נוצרת דרך `lists.create` ולא בהקלדה, ולכן חלון הביטול של
 * `list-autoformat-install.ts` **אינו** חמוש והמדידה היא של המנוע לבדו.
 * אם המנוע מסיר את הרשימה בעצמו, הביטול שלנו מתחרה בו ולא משלים אותו.
 *
 *   node scripts/qa/list-backspace-owner-probe.mjs
 */
import { openApp } from './harness.mjs';

const app = await openApp({ name: 'bs-owner', port: Number(process.env.QA_PORT ?? 9627) });

const state = () =>
  app
    .js(
      `(async function(){
    try {
      var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
      var info = await doc.selection.current();
      var seg = info && info.target && info.target.segments && info.target.segments[0];
      if (!seg) return JSON.stringify({ err: 'no block' });
      var st = await doc.lists.getState({ target: { kind: 'block', nodeType: 'paragraph', nodeId: seg.blockId } });
      var r = await doc.ranges.resolve({
        start: { kind: 'point', point: { kind: 'text', blockId: seg.blockId, offset: 0 } },
        end: { kind: 'point', point: { kind: 'text', blockId: seg.blockId, offset: 60 } },
      });
      var markers = Array.from(document.querySelectorAll('[class*="list-marker"]'))
        .filter(function(n){ return n.getBoundingClientRect().width > 0; })
        .map(function(n){ return n.textContent.replace(/\\u200f/g,'').trim(); });
      return JSON.stringify({
        blockId: seg.blockId,
        isListItem: st && st.isListItem,
        text: (r && r.preview && r.preview.text) || '',
        caret: info.selectionTarget && info.selectionTarget.start ? info.selectionTarget.start.offset : null,
        markers: markers,
      });
    } catch (e) { return JSON.stringify({ err: String((e && e.message) || e) }); }
  })()`,
    )
    .then(JSON.parse);

try {
  await app.caret(0);
  await app.sleep(400);

  // רשימה ריקה, שנוצרה דרך ה-API — חלון הביטול שלנו אינו חמוש.
  const created = await app
    .js(
      `(async function(){
    var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    await doc.clearContent({});
    await new Promise(function(r){ setTimeout(r, 300); });
    var info = await doc.selection.current();
    var seg = info.target.segments[0];
    var out = await doc.lists.create({
      mode: 'fromParagraphs',
      target: { kind: 'block', nodeType: 'paragraph', nodeId: seg.blockId },
      kind: 'ordered',
      style: { version: 1, levels: [{ level: 0, numFmt: 'hebrew1', lvlText: '%1)', markerFont: '' }] },
      sequence: { mode: 'new', startAt: 1 },
    });
    return JSON.stringify(out && out.success);
  })()`,
    )
    .then(JSON.parse);
  await app.sleep(700);
  await app.caret(0);
  await app.sleep(400);

  console.log('נוצרה רשימה:', created);
  console.log('לפני Backspace:', JSON.stringify(await state()));
  console.log('ידית זיהוי הרשימות:', await app.js('JSON.stringify(window.__otzariaListAutoformat || null)'));

  await app.press('Backspace', 'Backspace', 8);
  await app.sleep(900);
  console.log('אחרי Backspace:', JSON.stringify(await state()));

  // ושוב, על פריט שיש בו טקסט והסמן בתחילתו.
  await app.js(`(async function(){
    var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    await doc.clearContent({});
  })()`);
  await app.sleep(600);
  await app.caret(0);
  await app.type('שלום', 50);
  await app.sleep(600);
  await app.js(`(async function(){
    var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
    var info = await doc.selection.current();
    var seg = info.target.segments[0];
    await doc.lists.create({
      mode: 'fromParagraphs',
      target: { kind: 'block', nodeType: 'paragraph', nodeId: seg.blockId },
      kind: 'ordered',
      style: { version: 1, levels: [{ level: 0, numFmt: 'hebrew1', lvlText: '%1)', markerFont: '' }] },
      sequence: { mode: 'new', startAt: 1 },
    });
  })()`);
  await app.sleep(800);
  // הסמן לתחילת הפריט
  await app.press('Home', 'Home', 36);
  await app.sleep(400);
  console.log('פריט עם טקסט, לפני Backspace:', JSON.stringify(await state()));
  await app.press('Backspace', 'Backspace', 8);
  await app.sleep(900);
  console.log('פריט עם טקסט, אחרי Backspace:', JSON.stringify(await state()));

  console.log('לוג:', JSON.stringify(await app.log()));
} finally {
  app.close();
}
