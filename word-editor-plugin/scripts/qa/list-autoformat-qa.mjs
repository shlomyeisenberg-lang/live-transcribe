/**
 * זיהוי רשימות בהקלדה — השער, מול Chrome אמיתי ועל ה-dist הבנוי.
 *
 * מה שנמדד כאן ולא ניתן למדוד ב-jsdom:
 *
 *   1. **`lists.create` בקריאה אחת.** האם המנוע מקבל `mode:'fromParagraphs'`
 *      עם `style` ו-`sequence` יחד — כלומר האם אפשר ליצור רשימה **ישר**
 *      בסגנון המבוקש, בלי לעבור דרך רשימה עשרונית ולתקן אותה אחר כך. זה
 *      ההבדל בין „‏`א)` מופיע” לבין „‏`1.` מהבהב ואז מתחלף”.
 *   2. **הסמן שמצויר על המסך** — `א)` ולא `1.` ולא ריק.
 *   3. **הסמן של המשתמש** נשאר במסמך אחרי ההמרה.
 *   4. **מה שנכתב ל-docx**: `numFmt` ו-`lvlText` ב-numbering.xml.
 *   5. Backspace מיד אחרי ההמרה מחזיר את הטקסט.
 *   6. „‏`ב.`” אינו מומר — ההגנה מפני „ב׳ בניסן”.
 *   7. Enter ואז סמן ממיר רק את הפסקה החדשה. סמן באמצע פסקה אינו מומר.
 *   8. Ctrl+Z מיד אחרי ההמרה מחזיר את „א) ” בהקשה אחת, בלי שסמן עשרוני
 *      יצויר בדרך; אחרי הקלדה — הקשה אחת להקלדה ואחת להמרה. Ctrl+Y מחזיר.
 *   9. „- ” מצייר מקף — ההמרה שלנו, ולא התבליט של המנוע.
 *  10. המתג כבוי: גם הצורות שהמנוע ממיר בעצמו („1. ”, „- ”) נשארות טקסט,
 *      והרווח נכנס במקומו גם בהקלדה.
 *
 *   node scripts/qa/list-autoformat-qa.mjs
 */
import { openApp, createReport } from './harness.mjs';

const report = createReport('זיהוי רשימות בהקלדה', { strict: true });
const app = await openApp({ name: 'list-autoformat', port: Number(process.env.QA_PORT ?? 9623) });

/** הסמנים שמצוירים בפועל — מה שהמשתמש רואה. */
const markers = () =>
  app
    .js(
      `JSON.stringify(Array.from(document.querySelectorAll('[class*="list-marker"]'))` +
        `.filter(function(n){ return n.getBoundingClientRect().width > 0; })` +
        `.map(function(n){ return n.textContent.replace(/\\u200f/g,'').trim(); }))`,
    )
    .then(JSON.parse);

/** הרמה הראשונה של הרשימה שבסמן, כפי שהמנוע מדווח אותה. */
const levelZero = () =>
  app
    .js(
      `(async function(){
    try {
      var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
      var info = await doc.selection.current();
      var seg = info && info.target && info.target.segments && info.target.segments[0];
      if (!seg) return JSON.stringify({ err: 'no block' });
      var address = { kind: 'block', nodeType: 'listItem', nodeId: seg.blockId };
      var state = await doc.lists.getState({ target: { kind: 'block', nodeType: 'paragraph', nodeId: seg.blockId } });
      if (!state || state.isListItem !== true) return JSON.stringify({ isListItem: false, state: state });
      var style = await doc.lists.getStyle({ target: address });
      var levels = (style && style.style && style.style.levels) || [];
      var zero = levels.filter(function(l){ return l && l.level === 0; })[0] || null;
      return JSON.stringify({ isListItem: true, level0: zero });
    } catch (e) { return JSON.stringify({ err: String((e && e.message) || e) }); }
  })()`,
    )
    .then(JSON.parse);

/** הטקסט של הפסקה שבסמן, וההיסט של הסמן בתוכה. */
const paragraph = () =>
  app
    .js(
      `(async function(){
    try {
      var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
      var info = await doc.selection.current();
      var t = info && info.selectionTarget;
      var seg = info && info.target && info.target.segments && info.target.segments[0];
      if (!seg) return JSON.stringify({ err: 'no block' });
      var r = await doc.ranges.resolve({
        start: { kind: 'point', point: { kind: 'text', blockId: seg.blockId, offset: 0 } },
        end: { kind: 'point', point: { kind: 'text', blockId: seg.blockId, offset: 400 } },
      });
      return JSON.stringify({
        text: (r && r.preview && r.preview.text) || '',
        caret: t && t.start ? t.start.offset : null,
        blockId: seg.blockId,
      });
    } catch (e) { return JSON.stringify({ err: String((e && e.message) || e) }); }
  })()`,
    )
    .then(JSON.parse);

/**
 * מרוקן את המסמך ומחזיר את הסמן לתחילתו.
 *
 * `clearContent` לבדו אינו מספיק: הוא מנקה את הטקסט ומשאיר את הפסקה **פריט
 * רשימה**, והמקרה הבא היה נמדד על בלוק שכבר ברשימה — כלומר על כלום. נמדד
 * בריצה הראשונה של השער, כשכל המקרים אחרי הראשון החזירו תשובות של המקרה
 * שלפניהם.
 */
async function clearDoc() {
  for (let attempt = 0; attempt < 8; attempt += 1) {
    await app.js(`(async function(){
      var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
      try { await doc.clearContent({}); } catch (e) {}
      try {
        var info = await doc.selection.current();
        var seg = info && info.target && info.target.segments && info.target.segments[0];
        if (seg) {
          var state = await doc.lists.getState({ target: { kind: 'block', nodeType: 'paragraph', nodeId: seg.blockId } });
          if (state && state.isListItem === true) {
            await doc.lists.remove({ target: { kind: 'block', nodeType: 'listItem', nodeId: seg.blockId } });
          }
        }
      } catch (e) {}
    })()`);
    await app.sleep(500);
    await app.caret(0);
    await app.sleep(300);
    const state = await levelZero();
    const text = await paragraph();
    if (state.isListItem !== true && text.text === '') {
      // clearContent הוא API תוכנתי, ולכן אינו שולח למאזיני הממשק את אירוע
      // הפקודה שהמשתמש היה מפעיל בסרגל. מדמים את אותה פעולת סרגל מחוץ לעורך,
      // כדי שהשער ימדוד את הנתיב הציבורי ולא מצב זיכרון בלתי־אפשרי בממשק.
      await app.js("document.body.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }))");
      await app.sleep(800);
      await app.caret(0);
      await app.sleep(700);
      return;
    }
  }
  throw new Error('הניקוי בין המקרים לא הצליח להחזיר פסקה ריקה שאינה רשימה');
}

/** ידית האבחון של המודול — installs/evaluates/applies. */
const debugHandle = () =>
  app.js('JSON.stringify(window.__otzariaListAutoformat || null)').then(JSON.parse);

try {
  // ── 0. האם המודול בכלל הותקן ─────────────────────────────────────────────
  const installed = await debugHandle();
  console.log('ידית האבחון בעלייה:', JSON.stringify(installed));
  installed && installed.installs > 0
    ? report.pass('המודול הותקן על המסמך', `installs=${installed.installs}`)
    : report.fail('המודול הותקן על המסמך', JSON.stringify(installed));
  // ── 1. `lists.create` בקריאה אחת ─────────────────────────────────────────
  await clearDoc();
  await app.caret(0);
  await app.type('x', 60);
  await app.sleep(700);

  const createReceipt = await app
    .js(
      `(async function(){
    try {
      var doc = window.__otzariaEditor.superdoc.activeEditor.doc;
      var info = await doc.selection.current();
      var seg = info.target.segments[0];
      var out = await doc.lists.create({
        mode: 'fromParagraphs',
        target: { kind: 'block', nodeType: 'paragraph', nodeId: seg.blockId },
        kind: 'ordered',
        style: { version: 1, levels: [{ level: 0, numFmt: 'hebrew1', lvlText: '%1)', markerFont: '' }] },
        sequence: { mode: 'new', startAt: 1 },
      });
      return JSON.stringify(out);
    } catch (e) { return JSON.stringify({ threw: String((e && e.message) || e) }); }
  })()`,
    )
    .then(JSON.parse);
  await app.sleep(900);

  const createdOk = createReceipt && createReceipt.success === true;
  const drawn = await markers();
  console.log('lists.create receipt:', JSON.stringify(createReceipt).slice(0, 400));
  console.log('סמנים אחרי create:', JSON.stringify(drawn));

  createdOk
    ? report.pass('lists.create מקבל style ו-sequence בקריאה אחת', JSON.stringify(drawn))
    : report.fail('lists.create מקבל style ו-sequence בקריאה אחת', JSON.stringify(createReceipt).slice(0, 200));

  drawn.some((m) => m.includes('א'))
    ? report.pass('הסמן שצויר הוא עברי', drawn.join(' '))
    : report.fail('הסמן שצויר הוא עברי', drawn.join(' ') || 'אין סמן');

  // ── 2. הקלדה אמיתית: „א) ” ───────────────────────────────────────────────
  await clearDoc();
  await app.type('א', 60);
  await app.type(')', 60);
  await app.type(' ', 60);
  await app.sleep(1200);

  console.log('עקבה אחרי „א) ”:', JSON.stringify(await debugHandle()));
  const afterHebrew = await levelZero();
  const paraHebrew = await paragraph();
  const markersHebrew = await markers();
  console.log('אחרי „א) ”:', JSON.stringify({ afterHebrew, paraHebrew, markersHebrew }));

  afterHebrew.isListItem === true
    ? report.pass('„א) ” הפך את הפסקה לרשימה', JSON.stringify(afterHebrew.level0))
    : report.fail('„א) ” הפך את הפסקה לרשימה', JSON.stringify(afterHebrew).slice(0, 200));

  afterHebrew.level0 && afterHebrew.level0.numFmt === 'hebrew1'
    ? report.pass('numFmt הוא hebrew1', String(afterHebrew.level0.numFmt))
    : report.fail('numFmt הוא hebrew1', JSON.stringify(afterHebrew.level0));

  afterHebrew.level0 && afterHebrew.level0.lvlText === '%1)'
    ? report.pass('המפריד שהוקלד נשמר בתבנית', String(afterHebrew.level0.lvlText))
    : report.fail('המפריד שהוקלד נשמר בתבנית', JSON.stringify(afterHebrew.level0));

  paraHebrew.text === ''
    ? report.pass('הסמן שהוקלד נבלע', JSON.stringify(paraHebrew.text))
    : report.fail('הסמן שהוקלד נבלע', JSON.stringify(paraHebrew.text));

  // התקלה שדווחה: „הסמן נעלם מהמקום שלו”.
  paraHebrew.caret === 0
    ? report.pass('הסמן נשאר בתחילת הפריט', String(paraHebrew.caret))
    : report.fail('הסמן נשאר בתחילת הפריט', String(paraHebrew.caret));

  markersHebrew.some((m) => m.includes('א'))
    ? report.pass('המסך מצייר סמן עברי ולא עשרוני', markersHebrew.join(' '))
    : report.fail('המסך מצייר סמן עברי ולא עשרוני', markersHebrew.join(' ') || 'אין');

  // ── 3. ההקלדה ממשיכה לתוך הפריט ──────────────────────────────────────────
  await app.type('טקסט', 50);
  await app.sleep(700);
  const typedInto = await paragraph();
  typedInto.text === 'טקסט'
    ? report.pass('ההקלדה נכנסת לפריט', JSON.stringify(typedInto.text))
    : report.fail('ההקלדה נכנסת לפריט', JSON.stringify(typedInto.text));

  // ── 3ב. Enter מפסקה רגילה, ומיד סמן ────────────────────────────────────
  await clearDoc();
  await app.type('פתיחה', 90);
  await app.sleep(900);
  await app.press('Enter', 'Enter', 13, 0, '\r');
  await app.sleep(120);
  await app.type('א) טקסט', 120);
  await app.sleep(1500);
  const afterEnter = JSON.parse(
    await app.js(`(async function(){
      var l = await window.__otzariaEditor.superdoc.activeEditor.doc.blocks.list({ includeText: true });
      return JSON.stringify((l.blocks || []).map(function(b){ return [b.nodeType, b.text]; }));
    })()`),
  );
  console.log('Enter ואז סמן:', JSON.stringify(afterEnter));
  JSON.stringify(afterEnter) === JSON.stringify([['paragraph', 'פתיחה'], ['listItem', 'טקסט']])
    ? report.pass('Enter ואז סמן: רק הפסקה החדשה הומרה', JSON.stringify(afterEnter))
    : report.fail('Enter ואז סמן: רק הפסקה החדשה הומרה', JSON.stringify(afterEnter));

  // ── 3ג. סמן שהוקלד באמצע פסקה אינו מומר ──────────────────────────────
  await clearDoc();
  await app.type('abc', 90);
  await app.sleep(600);
  // End מאפס את הרצף: „א)” שאחריו הוא רצף חדש, שמתחיל בהיסט 3.
  await app.press('End', 'End', 35);
  await app.sleep(150);
  await app.type('א) ', 120);
  await app.sleep(2000);
  const midState = await levelZero();
  const midPara = await paragraph();
  midState.isListItem !== true && midPara.text === 'abcא) '
    ? report.pass('סמן באמצע פסקה נשאר טקסט', JSON.stringify(midPara.text))
    : report.fail('סמן באמצע פסקה נשאר טקסט', JSON.stringify({ midState, midPara }));

  // ── 4. „ב. ” אינו מומר ───────────────────────────────────────────────────
  await clearDoc();
  await app.type('ב', 60);
  await app.type('.', 60);
  await app.type(' ', 60);
  await app.sleep(1000);
  const afterBet = await levelZero();
  const paraBet = await paragraph();
  console.log('אחרי „ב. ”:', JSON.stringify({ afterBet, paraBet }));

  afterBet.isListItem !== true && paraBet.text === 'ב. '
    ? report.pass('„ב. ” נשאר טקסט', JSON.stringify(paraBet.text))
    : report.fail('„ב. ” נשאר טקסט', JSON.stringify({ afterBet, text: paraBet.text }).slice(0, 200));

  // ── 5. Backspace אחרי ההמרה — הטיפול הוא של המנוע ─────────────────────
  /*
   * לא היה כאן מסלול משלנו, והוא ירד אחרי מדידה: המנוע מסיר את הרשימה
   * בתחילת פריט ומשאיר את הטקסט, טוב יותר מ-Word (שמשאיר את ההזחה). השער
   * נועל את ההתנהגות **שלו**, כי עליה אנחנו נשענים — ר' list-backspace-owner-probe.
   */
  await clearDoc();
  await app.type('1', 60);
  await app.type('.', 60);
  await app.type(' ', 60);
  await app.sleep(1200);
  const beforeBack = await levelZero();
  await app.type('חידוש', 50);
  await app.sleep(700);
  await app.press('Home', 'Home', 36);
  await app.sleep(400);
  await app.press('Backspace', 'Backspace', 8);
  await app.sleep(900);
  const afterBack = await levelZero();
  const paraBack = await paragraph();
  console.log('Backspace:', JSON.stringify({ beforeBack, afterBack, paraBack }));

  beforeBack.isListItem === true && afterBack.isListItem !== true && paraBack.text === 'חידוש'
    ? report.pass('Backspace בתחילת הפריט מסיר את הרשימה ומשאיר את הטקסט', JSON.stringify(paraBack.text))
    : report.fail(
        'Backspace בתחילת הפריט מסיר את הרשימה ומשאיר את הטקסט',
        JSON.stringify({ before: beforeBack.isListItem, after: afterBack.isListItem, text: paraBack.text }),
      );

  // ── 5א. Ctrl+Z מיד אחרי ההמרה — הקשה אחת ─────────────────────────────
  const ctrlKey = (key, code, vk) => app.press(key, code, vk, 2);
  const seenMarkers = () => app.js('JSON.stringify(window.__laSeen || [])').then(JSON.parse);
  await app.js(`(function(){
    window.__laSeen = [];
    if (!window.__laSeenObs) {
      window.__laSeenObs = new MutationObserver(function(){
        Array.prototype.forEach.call(document.querySelectorAll('[class*="list-marker"]'), function(n){
          if (!n.getBoundingClientRect().width) return;
          var t = n.textContent.replace(/\u200f/g, '').trim();
          if (window.__laSeen.indexOf(t) < 0) window.__laSeen.push(t);
        });
      });
      window.__laSeenObs.observe(document.body, { subtree: true, childList: true, characterData: true });
    }
    return 1;
  })()`);

  await clearDoc();
  await app.type('א) ', 90);
  await app.sleep(1200);
  await app.js('window.__laSeen = []');
  await ctrlKey('z', 'KeyZ', 90);
  await app.sleep(1200);
  const undone = await paragraph();
  const undoneState = await levelZero();
  const flashed = (await seenMarkers()).filter((m) => /^\d/.test(m));
  console.log('Ctrl+Z אחרי ההמרה:', JSON.stringify({ undone, list: undoneState.isListItem, flashed }));
  undone.text === 'א) ' && undoneState.isListItem !== true
    ? report.pass('Ctrl+Z אחד מחזיר את „א) ” כטקסט', JSON.stringify(undone.text))
    : report.fail('Ctrl+Z אחד מחזיר את „א) ” כטקסט', JSON.stringify({ undone, undoneState }));
  flashed.length === 0
    ? report.pass('בדרך לא צויר סמן עשרוני')
    : report.fail('בדרך לא צויר סמן עשרוני', flashed.join(' '));

  await ctrlKey('y', 'KeyY', 89);
  await app.sleep(1200);
  const redone = await levelZero();
  redone.isListItem === true && redone.level0 && redone.level0.numFmt === 'hebrew1'
    ? report.pass('Ctrl+Y מחזיר את הרשימה העברית')
    : report.fail('Ctrl+Y מחזיר את הרשימה העברית', JSON.stringify(redone));

  // ── 5ב. אחרי הקלדה: הקשה להקלדה, והקשה להמרה ─────────────────────────
  await clearDoc();
  await app.type('א) ', 90);
  await app.sleep(1000);
  await app.type('טקסט', 90);
  await app.sleep(1000);
  await ctrlKey('z', 'KeyZ', 90);
  await app.sleep(1000);
  const firstUndo = await paragraph();
  await ctrlKey('z', 'KeyZ', 90);
  await app.sleep(1200);
  const secondUndo = await paragraph();
  const secondState = await levelZero();
  console.log('הקלדה ואז Ctrl+Z פעמיים:', JSON.stringify({ firstUndo, secondUndo, list: secondState.isListItem }));
  secondUndo.text === 'א) ' && secondState.isListItem !== true && firstUndo.text !== 'א) טקסט'
    ? report.pass('אחרי הקלדה: שני Ctrl+Z מחזירים את „א) ”', `${JSON.stringify(firstUndo.text)} → ${JSON.stringify(secondUndo.text)}`)
    : report.fail('אחרי הקלדה: שני Ctrl+Z מחזירים את „א) ”', JSON.stringify({ firstUndo, secondUndo, secondState }));

  // ── 5ג. „- ” — מקף, ולא התבליט של המנוע ──────────────────────────────
  await clearDoc();
  await app.type('- ', 90);
  await app.sleep(1500);
  const dash = await levelZero();
  const dashMarkers = await markers();
  dash.isListItem === true && dashMarkers.includes('-')
    ? report.pass('„- ” מצייר מקף', dashMarkers.join(' '))
    : report.fail('„- ” מצייר מקף', JSON.stringify({ dash, dashMarkers }));

  // ── 6. מה שנכתב ל-docx ───────────────────────────────────────────────────
  await clearDoc();
  await app.type('א', 60);
  await app.type(')', 60);
  await app.type(' ', 60);
  await app.sleep(1000);
  await app.type('חידוש', 50);
  await app.sleep(900);

  const files = await app.docx();
  const numbering = files['word/numbering.xml'] || '';
  const fmts = [...numbering.matchAll(/<w:numFmt w:val="([^"]+)"/g)].map((x) => x[1]);
  const texts = [...numbering.matchAll(/<w:lvlText w:val="([^"]*)"/g)].map((x) => x[1]);
  console.log('numbering.xml numFmt:', JSON.stringify(fmts.slice(0, 10)));
  console.log('numbering.xml lvlText:', JSON.stringify(texts.slice(0, 10)));

  fmts.includes('hebrew1')
    ? report.pass('hebrew1 נכתב ל-numbering.xml', fmts.filter((f) => f.startsWith('hebrew')).join(','))
    : report.fail('hebrew1 נכתב ל-numbering.xml', fmts.slice(0, 8).join(',') || 'אין');

  texts.includes('%1)')
    ? report.pass('התבנית %1) נכתבה ל-numbering.xml', '%1)')
    : report.fail('התבנית %1) נכתבה ל-numbering.xml', texts.slice(0, 8).join(',') || 'אין');

  console.log('ידית האבחון בסוף:', JSON.stringify(await debugHandle()));
  console.log('לוג:', JSON.stringify(await app.log()));
} finally {
  app.close();
}

// ── 7. המתג כבוי — מהפעלה קודמת ───────────────────────────────────────────
/*
 * יציאה אחרת בכוונה: פתיחה שנייה על אותה יציאה באותו תהליך נמדדה מפילה את
 * node במכונה הזאת (replace-race-probe, גם ב-main).
 */
const off = await openApp({
  name: 'list-autoformat-off',
  port: Number(process.env.QA_PORT ?? 9623) + 1,
  extra: `<script>window.__qaHost.storage['list-autoformat-enabled'] = false;</script>`,
});
try {
  const offParagraph = () =>
    off.js(`(async function(){
      var l = await window.__otzariaEditor.superdoc.activeEditor.doc.blocks.list({ includeText: true });
      return JSON.stringify((l.blocks || []).map(function(b){ return [b.nodeType, b.text]; }));
    })()`).then(JSON.parse);

  await off.caret(0);
  await off.sleep(400);
  await off.type('1. אבג', 90);
  await off.press('Enter', 'Enter', 13, 0, '\r');
  await off.sleep(300);
  await off.type('- ד', 40);
  await off.sleep(1800);
  const offBlocks = await offParagraph();
  console.log('כבוי:', JSON.stringify(offBlocks));
  JSON.stringify(offBlocks) === JSON.stringify([['paragraph', '1. אבג'], ['paragraph', '- ד']])
    ? report.pass('כבוי: „1. ” ו-„- ” נשארים טקסט, והרווח במקומו', JSON.stringify(offBlocks))
    : report.fail('כבוי: „1. ” ו-„- ” נשארים טקסט, והרווח במקומו', JSON.stringify(offBlocks));

  /*
   * רצף ההקלדה נקטע לפני הרווח: תיקון טעות באמצע הסימן, או הזזת סמן
   * וחזרה. המנוע קורא את הפסקה ולא את ההקשות, ולכן המיר בכל זאת (נמדד).
   */
  for (const [label, prefix, interrupt] of [
    ['Backspace באמצע הסימן', '1x', [['Backspace', 'Backspace', 8]]],
    ['חץ וחזרה באמצע הסימן', '1', [['ArrowLeft', 'ArrowLeft', 37], ['ArrowRight', 'ArrowRight', 39]]],
  ]) {
    await off.press('Enter', 'Enter', 13, 0, '\r');
    await off.sleep(300);
    await off.type(prefix, 90);
    await off.sleep(400);
    for (const [key, code, which] of interrupt) {
      await off.press(key, code, which);
      await off.sleep(250);
    }
    await off.type('. ט', 90);
    await off.sleep(1800);
    const blocks = await offParagraph();
    const last = blocks[blocks.length - 1];
    console.log(`כבוי — ${label}: ${JSON.stringify(last)}`);
    JSON.stringify(last) === JSON.stringify(['paragraph', '1. ט'])
      ? report.pass(`כבוי: ${label} — „1. ” נשאר טקסט`, JSON.stringify(last))
      : report.fail(`כבוי: ${label} — „1. ” נשאר טקסט`, JSON.stringify(last));
  }
} finally {
  off.close();
}

process.exit(report.print() > 0 ? 1 : 0);
