/**
 * גשש היתכנות: האם אפשר להזיז את הסמן **חזותית**, כמו Word, בכל סוגי
 * הפסקאות — כולל שורה שמערבבת עברית עם לטינית או ספרות — ובכמה זה עולה.
 *
 * ## למה הגשש הזה קיים
 *
 * `docs/engine-gaps.md` („חצים אופקיים בעברית”) מדד שהמנוע בוחר מודל כיוון
 * **לכל פסקה בנפרד**: חזותי בפסקה עברית נקייה, לוגי (=הפוך על המסך) ברשימה,
 * בטאב, ובשורה עם ספרות או לטינית. מסקנת אותו סעיף הייתה שאין לנו מה לתקן,
 * מפני ש„היעד הנכון תלוי במודל הדו-כיווני שהמנוע מחזיק ואיננו רואים”.
 *
 * הגשש הזה בודק את ההנחה הזאת, ולא מקבל אותה: הטקסט **מצויר** ב-DOM אמיתי,
 * ולכן המיקום החזותי של כל תו ניתן למדידה מכאן. אם המיקומים האלה מספיקים
 * כדי לחשב „מי התו שמימין לסמן על המסך”, אין צורך במודל של המנוע.
 *
 * ## שלוש השאלות, ושום דבר מעבר להן
 *
 * 1. **מיפוי** — אפשר לגזור מה-DOM מיקום מצויר לכל היסט בשורה?
 * 2. **נכונות** — האם המיפוי הזול תואם את **האמת**, כלומר את המקום שבו
 *    המנוע עצמו מצייר את הסמן לכל היסט? האמת נמדדת בכוח הזרוע:
 *    `authoring.setSelectionTarget` לכל היסט בתורו, וקריאת ה-x המצויר.
 * 3. **מחיר** — כמה מילישניות עולה המיפוי הזול, על השורה הארוכה ביותר.
 *
 * הגשש **אינו מתקן דבר ואינו נוגע בקוד המוצר.** הוא עונה „אפשר / אי אפשר”,
 * ובכמה. כל מה שהוא כותב הוא לבחירה שהוא בעצמו מחזיר לאחור.
 *
 *   node scripts/qa/rtl-caret-visual-feasibility.mjs      (QA_PORT עוקף 9713)
 */
import { openApp, createReport } from './harness.mjs';
import { zipStored } from './docx-fixtures.mjs';

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const RTL = '<w:bidi/>';
const numPr = (numId, ilvl = 0) =>
  `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`;

const pRuns = (runs, pPr = '') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}${runs
    .map(
      (r) =>
        `<w:r>${r.rPr ? `<w:rPr>${r.rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${r.t}</w:t></w:r>`,
    )
    .join('')}</w:p>`;

const p1 = (text, pPr = '', rPr = '<w:rtl/>') => pRuns([{ t: text, rPr }], pPr);

const lvl = () =>
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
  `<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>` +
  `<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>`;

const numberingXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W}>
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvl()}</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

const stylesXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/><w:qFormat/><w:pPr><w:ind w:left="720"/><w:contextualSpacing/></w:pPr></w:style>
</w:styles>`;

function buildDocx(bodyXml) {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>${bodyXml}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:bidi/></w:sectPr></w:body></w:document>`;
  return zipStored({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    'word/document.xml': documentXml,
    'word/numbering.xml': numberingXml(),
    'word/styles.xml': stylesXml(),
  });
}

/* שורה עברית ארוכה למדידת המחיר: מספיק ארוכה כדי לגלוש, כלומר השורה שתימדד
   היא שורה מלאה ברוחב העמוד — המקרה היקר ביותר שקיים בפועל. */
const LONG = Array.from({ length: 45 }, (_, i) => `מילה${i}`).join(' ');

const BODY = [
  p1('שלום עולם פשוט מאוד', RTL),
  p1('פריט ראשון ברשימה', RTL + numPr(1)),
  p1('פרק 12 בספר', RTL),
  pRuns(
    [
      { t: 'מילה ', rPr: '<w:rtl/>' },
      { t: 'ABC', rPr: '' },
      { t: ' מילה', rPr: '<w:rtl/>' },
    ],
    RTL,
  ),
  p1(LONG, RTL),
  p1('plain latin line', '', ''),
  /* שני קצות השורה: אי לועזי שפותח שורה, ואי לועזי שסוגר אותה. בשניהם
     ה„תפר” הוא מול **הפסקה** ולא מול ריצה שנייה, וזה המקרה שלא נמדד בסבב
     הראשון — ובלעדיו אין לדעת אם החריץ הראשון והאחרון חוקיים. */
  pRuns([{ t: 'ABC', rPr: '' }, { t: ' מילה כאן', rPr: '<w:rtl/>' }], RTL),
  p1('לפני הטבלה 4', RTL),
].join('');

const CASES = [
  { key: 'clean', label: 'עברית נקייה', find: 'שלום עולם פשוט מאוד', truth: true },
  { key: 'list', label: 'פריט ברשימה ממוספרת', find: 'פריט ראשון ברשימה', truth: true },
  { key: 'digits', label: 'עברית עם ספרות', find: 'פרק 12 בספר', truth: true, seam: true },
  { key: 'latin', label: 'עברית עם אי לטיני', find: 'מילה ABC מילה', truth: true, seam: true },
  { key: 'startLatin', label: 'שורה שמתחילה באי לטיני', find: 'ABC מילה כאן', truth: true },
  { key: 'endDigit', label: 'שורה שנגמרת בספרה', find: 'לפני הטבלה 4', truth: true },
  { key: 'long', label: 'שורה עברית ארוכה (מחיר)', find: 'מילה0 מילה1', truth: false },
  { key: 'control', label: 'בקרה: שורה לטינית', find: 'plain latin line', truth: true },
];

const docx = buildDocx(BODY);
const report = createReport('היתכנות: תנועת סמן חזותית בשורה דו-כיוונית');
const app = await openApp({ name: 'rtl-vf', port: Number(process.env.QA_PORT ?? 9713) });

/** „פתח קובץ” פותח דיאלוג; הבורר נקרא רק מ„עיון בקבצים…” שבתוכו. */
async function openFixture() {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(docx).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-vf',url:${JSON.stringify(dataUrl)},name:'rtl-vf.docx',size:${docx.length},access:'readwrite'}})}`,
  );
  await app.tab('קובץ');
  if (!(await app.click('פתח קובץ', { after: 500 }))) throw new Error('„פתח קובץ” לא נמצא');
  await app.js("document.querySelector('.open-browse')?.scrollIntoView({ block: 'center' })");
  await app.click('עיון בקבצים…', { after: 8000 });
  for (let waited = 0; waited < 40_000; waited += 250) {
    await app.sleep(250);
    if (!(await app.exists('.status-load'))) break;
  }
  await app.sleep(2500);
  for (let attempt = 0; attempt < 8; attempt += 1) {
    try {
      await app.tab('בית');
      break;
    } catch {
      await app.sleep(1500);
    }
  }
  const paras = await app.paraCount();
  console.log(`פתיחה: paraCount=${paras}`);
  return paras >= 6;
}

/* ------------------------------------------------------------------ */
/* העוזר שרץ בתוך הדף                                                   */
/* ------------------------------------------------------------------ */

/**
 * שלוש פעולות בתוך הדף:
 *
 *   `structure` — מה מצויר: ה-fragment, השורות שבו, והתווים שבשורת הסמן.
 *   `raw`       — התיבות המצוירות של כל תו בשורה, ועלות החילוץ במילישניות.
 *   `truth`     — האמת: לכל היסט, לאן המנוע עצמו מצייר את הסמן.
 *
 * `raw` הוא מה שיירוץ בזמן אמת אם התיקון ייבנה, ולכן הוא זה שנמדד בזמן.
 * `truth` הוא כלי מדידה בלבד — הוא כותב בחירה לכל היסט, וזה יקר מדי מכדי
 * לרוץ אי פעם במוצר.
 */
const HELPER = `window.__vf = (function () {
  function editor() {
    var oe = window.__otzariaEditor;
    return (oe && oe.superdoc && oe.superdoc.activeEditor) || null;
  }

  function fragmentOf(blockId, caretOffset) {
    var frags = document.querySelectorAll('[data-source-node-id][data-pm-start]');
    for (var i = 0; i < frags.length; i++) {
      var f = frags[i];
      if (f.getAttribute('data-source-node-id') !== blockId) continue;
      var base = Number(f.getAttribute('data-pm-start'));
      var end = Number(f.getAttribute('data-pm-end'));
      if (!isFinite(base)) continue;
      var pm = base + caretOffset;
      if (isFinite(end) && (pm < base || pm > end)) continue;
      return { el: f, base: base, end: end };
    }
    return null;
  }

  function lineOf(frag, pm) {
    var kids = frag.el.children;
    for (var i = 0; i < kids.length; i++) {
      var s = Number(kids[i].getAttribute('data-pm-start'));
      var e = Number(kids[i].getAttribute('data-pm-end'));
      if (isFinite(s) && isFinite(e) && pm >= s && pm <= e) {
        return { el: kids[i], pmStart: s, pmEnd: e, index: i, dir: kids[i].getAttribute('dir') };
      }
    }
    return null;
  }

  /**
   * כל תו בשורה, בסדר ה-DOM — כלומר בסדר הלוגי — עם התיבה שבה הוא מצויר.
   *
   * השדה seg הוא מזהה הריצה שהתו שייך לה. הוא נחוץ מפני שכיוון אינו תכונה
   * של תו בודד: השוואה לשכן **מעבר לגבול הריצה** מסווגת את התו האחרון באי
   * הלטיני כעברי, וזה בדיוק הפער היחיד שנמדד בסבב הראשון.
   */
  function charBoxes(lineEl) {
    var out = [];
    var walker = document.createTreeWalker(lineEl, NodeFilter.SHOW_TEXT, null);
    var node;
    var range = document.createRange();
    var segs = [];
    while ((node = walker.nextNode())) {
      var host = node.parentElement;
      var declared = host && host.closest('[dir]');
      var dir = declared ? declared.getAttribute('dir') : null;
      var cls = host ? String(host.className) : '';
      var segIndex = segs.indexOf(host);
      if (segIndex < 0) { segs.push(host); segIndex = segs.length - 1; }
      for (var i = 0; i < node.data.length; i++) {
        range.setStart(node, i);
        range.setEnd(node, i + 1);
        var r = range.getBoundingClientRect();
        out.push({
          ch: node.data[i],
          left: Math.round(r.left * 10) / 10,
          right: Math.round(r.right * 10) / 10,
          width: Math.round(r.width * 10) / 10,
          dir: dir,
          seg: segIndex,
          run: cls.indexOf('superdoc-text-run') >= 0,
          cls: cls.slice(0, 40)
        });
      }
    }
    return out;
  }

  return {
    structure: function (blockId, caretOffset) {
      var frag = fragmentOf(blockId, caretOffset);
      if (!frag) return JSON.stringify({ error: 'no fragment' });
      var pm = frag.base + caretOffset;
      var line = lineOf(frag, pm);
      var lines = [];
      for (var i = 0; i < frag.el.children.length; i++) {
        var k = frag.el.children[i];
        lines.push({
          pmStart: Number(k.getAttribute('data-pm-start')),
          pmEnd: Number(k.getAttribute('data-pm-end')),
          dir: k.getAttribute('dir'),
          tag: k.tagName,
          cls: String(k.className).slice(0, 40),
          text: (k.textContent || '').slice(0, 40)
        });
      }
      if (!line) return JSON.stringify({ base: frag.base, end: frag.end, lines: lines, error: 'no line' });
      var chars = charBoxes(line.el);
      return JSON.stringify({
        base: frag.base,
        end: frag.end,
        caretPm: pm,
        lines: lines,
        line: { pmStart: line.pmStart, pmEnd: line.pmEnd, dir: line.dir, index: line.index },
        charCount: chars.length,
        span: line.pmEnd - line.pmStart,
        text: (line.el.textContent || '').slice(0, 80)
      });
    },

    /** התיבות עצמן + כמה זמן לקח לחלץ אותן. \`reps\` מדידות חוזרות לחציון. */
    raw: function (blockId, caretOffset, reps) {
      var frag = fragmentOf(blockId, caretOffset);
      if (!frag) return JSON.stringify({ error: 'no fragment' });
      var line = lineOf(frag, frag.base + caretOffset);
      if (!line) return JSON.stringify({ error: 'no line' });
      var times = [];
      var chars = null;
      var n = reps || 1;
      for (var i = 0; i < n; i++) {
        var t0 = performance.now();
        chars = charBoxes(line.el);
        times.push(performance.now() - t0);
      }
      times.sort(function (a, b) { return a - b; });
      return JSON.stringify({
        base: frag.base,
        lineStart: line.pmStart,
        lineEnd: line.pmEnd,
        lineDir: line.dir,
        chars: chars,
        ms: {
          median: Math.round(times[Math.floor(times.length / 2)] * 1000) / 1000,
          min: Math.round(times[0] * 1000) / 1000,
          max: Math.round(times[times.length - 1] * 1000) / 1000,
          reps: times.length
        }
      });
    },

    /*
     * המחיר **המלא** של הקשה, ולא של חילוץ התיבות בלבד.
     *
     * (בלי גרשיים אחוריים: כל הפונקציה הזאת יושבת בתוך תבנית מחרוזת.)
     *
     * raw מודד רכיב אחד מתוך המסלול; כאן נשלח keydown אמיתי אל ה-DOM, ולכן
     * נמדד כל מה שהיירוט עושה בפועל: הסמן המצויר ומלבנו, שלוש הסריקות על
     * המארח (blockStart פעמיים ו-findFragment), lineDistance לכל מועמדת,
     * readLineChars, ובקצה גם קריאה שנייה שלה — וכתיבת הבחירה. ההקשה נבלעת
     * בפסקה עברית, ולכן המנוע אינו מוסיף לזמן הזה.
     *
     * הסמן זז עם כל הקשה, בדיוק כמו אצל המשתמש; מי שקורא כאן צריך להחזיר
     * אותו אחר כך.
     */
    keypress: function (blockId, caretOffset, key, reps) {
      var frag = fragmentOf(blockId, caretOffset);
      if (!frag) return JSON.stringify({ error: 'no fragment' });
      var times = [];
      var prevented = 0;
      var n = reps || 1;
      for (var i = 0; i < n; i++) {
        var ev = new KeyboardEvent('keydown', { key: key, bubbles: true, cancelable: true });
        var t0 = performance.now();
        frag.el.dispatchEvent(ev);
        times.push(performance.now() - t0);
        if (ev.defaultPrevented) prevented++;
      }
      times.sort(function (a, b) { return a - b; });
      return JSON.stringify({
        median: Math.round(times[Math.floor(times.length / 2)] * 1000) / 1000,
        min: Math.round(times[0] * 1000) / 1000,
        max: Math.round(times[times.length - 1] * 1000) / 1000,
        reps: times.length,
        prevented: prevented
      });
    },

    /** מלבן שורת הסמן — כדי לסרוק אותה בלחיצות אמיתיות. */
    lineRect: function (blockId, caretOffset) {
      var frag = fragmentOf(blockId, caretOffset);
      if (!frag) return JSON.stringify({ error: 'no fragment' });
      var line = lineOf(frag, frag.base + caretOffset);
      if (!line) return JSON.stringify({ error: 'no line' });
      var r = line.el.getBoundingClientRect();
      return JSON.stringify({
        left: Math.round(r.left * 10) / 10,
        right: Math.round(r.right * 10) / 10,
        y: Math.round((r.top + r.bottom) / 2),
        pmStart: line.pmStart,
        pmEnd: line.pmEnd
      });
    },

    /** אחרי לחיצה: איזה היסט המנוע בחר, ואיפה הוא צייר את הסמן. */
    after: function () {
      var ed = editor();
      var snap = ed && ed.host && ed.host.readLiveSelectionSyncSnapshot && ed.host.readLiveSelectionSyncSnapshot();
      var end = snap && snap.selectionTarget && snap.selectionTarget.end;
      var el = document.querySelector('.sd-v2-local-selection-caret');
      var r = el ? el.getBoundingClientRect() : null;
      return JSON.stringify({
        off: end ? end.offset : null,
        blockId: end ? end.blockId : null,
        x: r ? Math.round(r.left * 10) / 10 : null
      });
    },

    /**
     * האם ל-API יש ידית ל„צד” של התפר.
     *
     * הלחיצה מגיעה לשני הבתים של אותו היסט; כתיבת בחירה לפי היסט מגיעה
     * לאחד. השאלה כאן היא אם אחת מצורות הכתיבה האחרות — עוגן מהצד השני,
     * collapse, או כיוון הפוך — מוסרת למנוע את אותו מידע.
     */
    affinity: async function (blockId, a, b, settleMs) {
      var ed = editor();
      var snap = ed && ed.host && ed.host.readLiveSelectionSyncSnapshot && ed.host.readLiveSelectionSyncSnapshot();
      var sel = snap && snap.selectionTarget;
      if (!sel || !sel.end) return JSON.stringify({ error: 'no snapshot' });
      var story = sel.story || sel.end.story || { kind: 'story', storyType: 'body' };
      var pt = function (o) { return { kind: 'text', blockId: blockId, offset: o, story: story }; };
      var variants = [
        { name: 'start=end=' + a, target: { start: pt(a), end: pt(a) } },
        { name: 'start=' + b + ' end=' + a + ' collapse=end', target: { start: pt(b), end: pt(a) }, collapse: 'end' },
        { name: 'start=' + a + ' end=' + b + ' collapse=start', target: { start: pt(a), end: pt(b) }, collapse: 'start' },
        { name: 'start=end=' + b, target: { start: pt(b), end: pt(b) } },
        { name: 'start=' + a + ' end=' + b + ' collapse=end', target: { start: pt(a), end: pt(b) }, collapse: 'end' }
      ];
      var out = [];
      for (var i = 0; i < variants.length; i++) {
        var v = variants[i];
        var input = { target: { kind: 'selection', start: v.target.start, end: v.target.end, story: story }, focus: true };
        if (v.collapse) input.collapse = v.collapse;
        try {
          await ed.authoring.setSelectionTarget(input);
        } catch (e) {
          out.push({ name: v.name, x: null, err: String(e).slice(0, 50) });
          continue;
        }
        await new Promise(function (r) { setTimeout(r, settleMs || 80); });
        var el = document.querySelector('.sd-v2-local-selection-caret');
        var r2 = el ? el.getBoundingClientRect() : null;
        out.push({ name: v.name, x: r2 ? Math.round(r2.left * 10) / 10 : null });
      }
      return JSON.stringify({ variants: out });
    },

    /**
     * האמת: לכל היסט בטווח, לאן המנוע מצייר את הסמן.
     *
     * כלי מדידה בלבד — כתיבת בחירה לכל היסט. הסמן מוחזר בסוף להיסט שממנו
     * התחלנו, כדי שהמדידה לא תשאיר את המסמך במצב אחר.
     */
    truth: async function (blockId, from, to, settleMs) {
      var ed = editor();
      if (!ed || !ed.authoring || !ed.authoring.setSelectionTarget) return JSON.stringify({ error: 'no authoring' });
      var snap = ed.host && ed.host.readLiveSelectionSyncSnapshot && ed.host.readLiveSelectionSyncSnapshot();
      var sel = snap && snap.selectionTarget;
      if (!sel || !sel.end) return JSON.stringify({ error: 'no snapshot' });
      var story = sel.story || sel.end.story || { kind: 'story', storyType: 'body' };
      var back = sel.end.offset;
      var out = [];
      for (var o = from; o <= to; o++) {
        var point = { kind: 'text', blockId: blockId, offset: o, story: story };
        try {
          await ed.authoring.setSelectionTarget({
            target: { kind: 'selection', start: point, end: point, story: story },
            focus: true
          });
        } catch (e) {
          out.push({ off: o, x: null, err: String(e).slice(0, 60) });
          continue;
        }
        await new Promise(function (r) { setTimeout(r, settleMs || 60); });
        var el = document.querySelector('.sd-v2-local-selection-caret');
        var r2 = el ? el.getBoundingClientRect() : null;
        out.push({ off: o, x: r2 ? Math.round(r2.left * 10) / 10 : null, y: r2 ? Math.round(r2.top * 10) / 10 : null });
      }
      var home = { kind: 'text', blockId: blockId, offset: back, story: story };
      try {
        await ed.authoring.setSelectionTarget({
          target: { kind: 'selection', start: home, end: home, story: story },
          focus: true
        });
      } catch (e) { /* החזרה בלבד */ }
      return JSON.stringify({ slots: out, restored: back });
    }
  };
})(); 'ok'`;

/** ההיסט של הסמן, מהמנוע. */
const engineCaret = () =>
  app
    .js(`(async function(){
      try {
        var ed = window.__otzariaEditor && window.__otzariaEditor.superdoc && window.__otzariaEditor.superdoc.activeEditor;
        if (!ed || !ed.doc) return JSON.stringify({error:'no editor'});
        var info = await ed.doc.selection.current();
        var seg = info && info.target && info.target.segments && info.target.segments[0];
        return JSON.stringify({ blockId: seg ? seg.blockId : null, start: seg && seg.range ? seg.range.start : null });
      } catch (e) { return JSON.stringify({error:String(e)}); }
    })()`)
    .then(JSON.parse);

/* ------------------------------------------------------------------ */
/* התאמת הכלל: איזה חוק ממפה היסט → x                                   */
/* ------------------------------------------------------------------ */

/**
 * הכלל המועמד: הסמן בהיסט o יושב על ה**קצה המוביל** של התו שבאינדקס o —
 * שמאל לתו לטיני, ימין לתו עברי — ובסוף השורה על הקצה הנגרר של האחרון.
 *
 * הכיוון של כל תו **נמדד מהגאומטריה**, לא מההצהרה: תו ש„הבא אחריו” מצויר
 * שמאלה ממנו הוא תו בקטע RTL. זו בדיוק המסקנה של
 * `bidi-declaration-is-not-direction` — כיוון הוא תכונה של מה שמצויר, לא של
 * מה שהוצהר — וכאן היא נבדקת מול האמת ולא מונחת.
 */
function slotsFromChars(all, lineStart, lineRtl = true) {
  /* סמן המספור של פריט רשימה מצויר בתוך השורה אך אינו טקסט של המסמך, ולכן
     אינו נושא היסט. בסבב הראשון הוא הזיז את כל המפה ב-3 — ההתאמה ל-x הייתה
     מושלמת, רק הכתובת הייתה שגויה. `superdoc-text-run` הוא ההבחנה. */
  const chars = all.filter((c) => c.run);
  if (!chars.length) return [];

  /* כיוון נקבע מ**צמידות**, לא מהשוואה לשכן ולא מההצהרה: שני תווים עוקבים
     שהתיבות שלהם נוגעות זו בזו שייכים לאותו קטע דו-כיווני, והצד שבו הן
     נוגעות הוא הכיוון. חלוקה לפי אלמנט ה-DOM אינה מספיקה — „פרק 12 בספר”
     הוא ריצה **אחת** שבתוכה אי לטיני, וזה מה שהחטיא את התו האחרון של האי. */
  const TOUCH = 0.6;
  const rtlAt = new Array(chars.length).fill(null);
  for (let i = 0; i + 1 < chars.length; i += 1) {
    const a = chars[i];
    const b = chars[i + 1];
    let dir = null;
    if (Math.abs(b.left - a.right) < TOUCH) dir = false;
    else if (Math.abs(a.left - b.right) < TOUCH) dir = true;
    if (dir === null) continue;
    if (rtlAt[i] === null) rtlAt[i] = dir;
    rtlAt[i + 1] = dir;
  }
  /* אי של תו **אחד** נוגע בשכניו בדיוק כמו תו עברי, ולכן הגאומטריה לבדה
     מסווגת אותו כעברי. זה נמדד כשגוי: ב„לפני הטבלה 4” הספרה קיבלה כך את שני
     הקצוות הפוכים מהמנוע (411 ו-412 התחלפו, ‏17.9.2026). המבחן הוא תכונת
     התו, בדיוק כמו ב-`isLtrIntrinsic` שבמוצר — ספרה או אות שאינה ממערכת כתב
     ימנית; סימן ניטרלי נשאר בכיוון השורה. */
  const RTL_SCRIPT =
    /[\p{Script=Hebrew}\p{Script=Arabic}\p{Script=Syriac}\p{Script=Thaana}\p{Script=Nko}\p{Script=Samaritan}\p{Script=Mandaic}\p{Script=Adlam}]/u;
  const ltrIntrinsic = (ch) =>
    !!ch && (/\p{Nd}/u.test(ch) || (/\p{Alphabetic}/u.test(ch) && !RTL_SCRIPT.test(ch)));
  for (let i = 0; i < chars.length; i += 1) {
    if (rtlAt[i] !== false && ltrIntrinsic(chars[i].ch)) rtlAt[i] = false;
    if (rtlAt[i] === null) rtlAt[i] = chars[i].dir === 'rtl' || lineRtl;
  }

  const slots = chars.map((c, i) => ({ off: lineStart + i, x: rtlAt[i] ? c.right : c.left }));
  const last = chars[chars.length - 1];
  slots.push({ off: lineStart + chars.length, x: rtlAt[chars.length - 1] ? last.left : last.right });
  return slots;
}

/** „מי השכן החזותי מימין” לפי מפת החריצים. זה מה שהחץ צריך לעשות. */
function visualRight(slots, off) {
  const here = slots.find((s) => s.off === off);
  if (!here) return null;
  let best = null;
  for (const s of slots) {
    if (s.x <= here.x + 0.5) continue;
    if (!best || s.x < best.x) best = s;
  }
  return best;
}

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1500, height: 1000, deviceScaleFactor: 1, mobile: false });
  await app.sleep(400);
  if (!(await openFixture())) {
    report.fail('פתיחת הקובץ', 'המסמך לא נפתח');
    throw new Error('no doc');
  }
  await app.js(HELPER);

  const costs = [];
  /** אותה שורה, אבל המסלול המלא: keydown אמיתי מקצה לקצה. */
  const keypresses = [];

  for (const c of CASES) {
    console.log(`\n================ ${c.label}`);
    try {
      await app.caretPara(c.find);
    } catch (e) {
      report.stuck(c.label, `מיקום סמן נכשל: ${String(e.message || e).slice(0, 80)}`);
      continue;
    }
    const caret = await engineCaret();
    if (!caret.blockId || typeof caret.start !== 'number') {
      report.stuck(c.label, 'המנוע לא מסר סמן');
      continue;
    }

    /* -------- 1. מבנה -------- */
    const structure = JSON.parse(
      await app.js(`window.__vf.structure(${JSON.stringify(caret.blockId)}, ${caret.start})`),
    );
    if (structure.error) {
      report.fail(`${c.label} — מיפוי`, `אין מבנה לקרוא: ${structure.error}`);
      continue;
    }
    console.log(`   שורות בפסקה: ${structure.lines.length}`);
    for (const l of structure.lines) {
      console.log(`     [${l.pmStart}..${l.pmEnd}] dir=${l.dir ?? '—'} <${l.tag}.${l.cls}> „${l.text}”`);
    }
    console.log(
      `   שורת הסמן: [${structure.line.pmStart}..${structure.line.pmEnd}] dir=${structure.line.dir ?? '—'}` +
        ` — ${structure.charCount} תווים מצוירים מול טווח ${structure.span}`,
    );


    /* -------- 2. התיבות והמחיר -------- */
    const raw = JSON.parse(
      await app.js(`window.__vf.raw(${JSON.stringify(caret.blockId)}, ${caret.start}, 40)`),
    );
    if (raw.error) {
      report.fail(`${c.label} — מיפוי`, raw.error);
      continue;
    }
    console.log(`   מחיר החילוץ: חציון ${raw.ms.median}ms, מקס ${raw.ms.max}ms (${raw.chars.length} תווים, ${raw.ms.reps} מדידות)`);
    costs.push({ label: c.label, chars: raw.chars.length, ...raw.ms });

    const runChars = raw.chars.filter((ch) => ch.run);
    const extra = raw.chars.length - runChars.length;
    const aligned = runChars.length === raw.lineEnd - raw.lineStart;
    if (extra) console.log(`   ${extra} תווים מצוירים שאינם טקסט המסמך (סמן רשימה): ${raw.chars.filter((ch) => !ch.run).map((ch) => `„${ch.ch}”(${ch.cls})`).join(' ')}`);
    if (!aligned) console.log(`   ⚠ גם אחרי הסינון: ${runChars.length} תווים מול טווח ${raw.lineEnd - raw.lineStart}`);

    const slots = slotsFromChars(raw.chars, raw.lineStart, raw.lineDir === 'rtl');
    console.log(`   חריצים: ${slots.map((s) => `${s.off}@${s.x}`).join(' ')}`);
    /* התיבות עצמן — מהן נבנים הקיבועים ב-`tests/unit/rtl-caret.test.ts`,
       ולכן הן נדפסות ולא רק נגזרות. */
    if (runChars.length <= 24) {
      console.log(`   תיבות: ${runChars.map((ch) => `„${ch.ch}”[${ch.left},${ch.right}]`).join(' ')}`);
    }

    /* -------- 2ב. המחיר המלא של הקשה, ולא של החילוץ בלבד -------- */
    const e2e = JSON.parse(
      await app.js(
        `window.__vf.keypress(${JSON.stringify(caret.blockId)}, ${caret.start}, 'ArrowRight', 40)`,
      ),
    );
    if (e2e.error) {
      console.log(`   מחיר ההקשה: ${e2e.error}`);
    } else {
      console.log(
        `   מחיר ההקשה המלא: חציון ${e2e.median}ms, מקס ${e2e.max}ms` +
          ` (${e2e.reps} הקשות, ${e2e.prevented} מהן נבלעו ביירוט)`,
      );
      keypresses.push({ label: c.label, chars: raw.chars.length, ...e2e });
    }
    await app.caretPara(c.find);

    if (!c.truth) {
      report.pass(`${c.label} — מחיר`, `חציון ${raw.ms.median}ms על ${raw.chars.length} תווים`);
      continue;
    }

    /* -------- 3. האמת מהמנוע -------- */
    const from = raw.lineStart - raw.base;
    const to = raw.lineEnd - raw.base;
    if (to - from > 40) {
      report.skip(`${c.label} — נכונות`, 'השורה ארוכה מכדי למדוד אמת היסט-היסט');
      continue;
    }
    const truth = JSON.parse(
      await app.js(`window.__vf.truth(${JSON.stringify(caret.blockId)}, ${from}, ${to}, 70)`),
    );
    if (truth.error) {
      report.stuck(`${c.label} — נכונות`, truth.error);
      continue;
    }
    console.log(`   אמת:     ${truth.slots.map((s) => `${raw.base + s.off}@${s.x ?? '—'}`).join(' ')}`);

    /* -------- 4. השוואה -------- */
    let matched = 0;
    let missing = 0;
    const diffs = [];
    for (const t of truth.slots) {
      const pm = raw.base + t.off;
      const mine = slots.find((s) => s.off === pm);
      if (t.x === null || !mine) {
        missing += 1;
        continue;
      }
      const d = Math.abs(mine.x - t.x);
      if (d <= 2) matched += 1;
      else diffs.push(`${pm}: שלי ${mine.x} מול ${t.x}`);
    }
    const total = truth.slots.length;
    console.log(`   התאמה: ${matched}/${total}` + (diffs.length ? ` — פערים: ${diffs.slice(0, 6).join('; ')}` : ''));

    /* -------- 5. האם הצעד החזותי אמיתי -------- */
    const ordered = [...truth.slots].filter((s) => s.x !== null).sort((a, b) => b.x - a.x);
    const uniqueX = new Set(ordered.map((s) => s.x));
    const seam = ordered.length - uniqueX.size;
    const startOff = raw.base + ordered[ordered.length - 1].off;
    const step = visualRight(slots, startOff);
    console.log(
      `   שכן חזותי מהקצה השמאלי (היסט ${startOff}): ${step ? `${step.off}@${step.x}` : '—'}` +
        `; חריצים שחולקים x: ${seam}`,
    );

    /* -------- 6. התפר: האם החריץ החסר בר-הגעה בכלל -------- */
    if (c.seam) {
      const lr = JSON.parse(await app.js(`window.__vf.lineRect(${JSON.stringify(caret.blockId)}, ${caret.start})`));
      /* רק על הגליפים עצמם: מלבן השורה הוא כל רוחב העמודה, ובפסקה עברית רובו
         שטח ריק משמאל לטקסט — סריקה שלו נחתה 200 פעם על סוף השורה. */
      const from = Math.round(Math.min(...runChars.map((ch) => ch.left))) - 4;
      const to = Math.round(Math.max(...runChars.map((ch) => ch.right))) + 4;
      const reachable = new Map();
      const hits = [];
      for (let x = from; x <= to; x += 3) {
        /* הרחק ואיטי: לחיצות צפופות באותו אזור נקראות כלחיצה כפולה, והמנוע
           מחליף אז בחירה ומאבד את הסמן — נמדד, כל הסריקה חזרה ריקה. */
        await app.clickAt(x, lr.y);
        await app.sleep(650);
        const a = JSON.parse(await app.js('window.__vf.after()'));
        hits.push(`${x}→${a.off === null ? '?' : raw.base + a.off}@${a.x ?? '—'}`);
        if (a.x !== null) reachable.set(a.x, a.off === null ? '?' : raw.base + a.off);
      }
      console.log(`   סריקת לחיצות ב-[${from}..${to}] y=${lr.y}:`);
      console.log(`     ${hits.join(' ')}`);
      const drawn = [...reachable.keys()].sort((a, b) => a - b);
      console.log(`   לחיצות, ייחודי: ${drawn.map((x) => `${reachable.get(x)}@${x}`).join(' ')}`);

      /* הגבולות המצוירים של כל תו — כל מקום שבו סמן *צריך* להיות. */
      const edges = new Set();
      for (const ch of runChars) {
        edges.add(Math.round(ch.left * 10) / 10);
        edges.add(Math.round(ch.right * 10) / 10);
      }
      const truthX = new Set(truth.slots.filter((s) => s.x !== null).map((s) => s.x));
      const clickX = new Set(drawn);
      const unreachable = [...edges]
        .filter((x) => ![...truthX].some((t) => Math.abs(t - x) <= 2))
        .filter((x) => ![...clickX].some((t) => Math.abs(t - x) <= 2))
        .sort((a, b) => a - b);
      console.log(
        `   גבולות גליפים: ${edges.size}; מהם בני-הגעה בכתיבת בחירה: ${truthX.size}, בלחיצה: ${clickX.size}` +
          `; בלתי-נגישים בשתי הדרכים: ${unreachable.length ? unreachable.join(', ') : 'אין'}`,
      );
      /* שני ההיסטים שחולקים מיקום — הבית הכפול של התפר. */
      const byX = new Map();
      for (const t of truth.slots) {
        if (t.x === null) continue;
        byX.set(t.x, [...(byX.get(t.x) ?? []), raw.base + t.off]);
      }
      const shared = [...byX.values()].find((list) => list.length > 1);
      if (shared) {
        const aff = JSON.parse(
          await app.js(`window.__vf.affinity(${JSON.stringify(caret.blockId)}, ${shared[0] - raw.base}, ${shared[1] - raw.base}, 90)`),
        );
        console.log(`   ידית ל„צד” של התפר (היסטים ${shared.join(' ו-')}):`);
        for (const v of aff.variants ?? []) console.log(`     ${v.name} → ${v.x ?? v.err ?? '—'}`);
        const xs = new Set((aff.variants ?? []).map((v) => v.x).filter((x) => x !== null));
        report[xs.size > 1 ? 'pass' : 'partial'](
          `${c.label} — ידית לתפר ב-API`,
          xs.size > 1 ? `יש: ${[...xs].join(' / ')}` : `אין — כל הצורות נותנות ${[...xs][0] ?? '—'}`,
        );
        await app.caretPara(c.find);
      }

      if (unreachable.length) {
        report.partial(`${c.label} — תפר`, `${unreachable.length} מיקומי סמן אינם קיימים במנוע כלל (x: ${unreachable.join(', ')})`);
      } else {
        report.pass(`${c.label} — תפר`, 'כל גבול גליף בר-הגעה');
      }
      await app.caretPara(c.find);
    }

    const detail =
      `${matched}/${total} חריצים תואמים את המנוע` +
      (seam ? `, ${seam} חריצים חולקים מיקום (תפר)` : '') +
      (aligned ? '' : ', מיפוי אינדקס→היסט אינו ישיר');

    if (matched === total && aligned) report.pass(`${c.label} — נכונות`, detail);
    else if (matched >= total - 1) report.partial(`${c.label} — נכונות`, detail);
    else report.fail(`${c.label} — נכונות`, detail);
  }

  console.log('\n---- מחיר, מרוכז ----');
  for (const c of costs) {
    console.log(`   ${c.label}: ${c.chars} תווים → חציון ${c.median}ms, מקס ${c.max}ms`);
  }
  console.log('\n---- מחיר ההקשה המלא, מרוכז ----');
  for (const c of keypresses) {
    console.log(
      `   ${c.label}: ${c.chars} תווים → חציון ${c.median}ms, מקס ${c.max}ms` +
        ` (${c.prevented}/${c.reps} נבלעו)`,
    );
  }
  const heaviest = keypresses.reduce((a, b) => (b.median > (a?.median ?? -1) ? b : a), null);
  if (heaviest) {
    const worstMax = keypresses.reduce((a, b) => (b.max > a ? b.max : a), 0);
    report[worstMax < 16 ? 'pass' : 'partial'](
      'מחיר: הקשה מקצה לקצה',
      `חציון ${heaviest.median}ms בכבדה ביותר (${heaviest.label}), והגרוע בכל הריצה ${worstMax}ms`,
    );
  }

  const worst = costs.reduce((a, b) => (b.max > (a?.max ?? -1) ? b : a), null);
  if (worst) {
    if (worst.max < 4) report.pass('מחיר: השורה היקרה ביותר', `${worst.max}ms במקרה הגרוע (${worst.label})`);
    else if (worst.max < 16) report.partial('מחיר: השורה היקרה ביותר', `${worst.max}ms — מתחת לפריים, אך לא זניח`);
    else report.fail('מחיר: השורה היקרה ביותר', `${worst.max}ms — מעל פריים`);
  }
} catch (error) {
  console.error(error);
  report.fail('הריצה', String(error.message || error).slice(0, 120));
} finally {
  report.print();
  app.close();
}
