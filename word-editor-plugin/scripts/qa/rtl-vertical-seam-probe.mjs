/**
 * גשש: שתי שאלות שמכריעות את המדיניות של תיקון החצים האנכיים.
 *
 * העקיפה של החצים האנכיים חייבת לכתוב בחירה דרך `authoring.setSelectionTarget`,
 * ושתי עובדות על הכתיבה הזאת אינן ידועות מהקוד — רק ממדידה:
 *
 * **1. תפר גלישה.** ההיסט שבין שתי שורות גולשות שייך לשתיהן: הוא `data-pm-end`
 * של הראשונה וגם `data-pm-start` של השנייה. ‏`rtl-line-end.ts` מדד שמקש `End`
 * על ההיסט הזה נשאר בשורה ה**קודמת** — כלומר „התפר שייך לקודמת”. אם אותו כלל
 * חל גם על כתיבה דרך ה-API, אז „רד לתחילת השורה הבאה” אינו ניתן לביטוי כלל,
 * והעקיפה חייבת לזהות את המקרה ולמסור אותו למנוע. אם לעומת זאת הכתיבה נוחתת
 * בתחילת השורה השנייה — אין מלכודת, והמדיניות פשוטה יותר.
 *
 * **2. עמודת המטרה אחרי כתיבה שלנו.** נמדד כבר שהמנוע מחזיק עמודת מטרה ומחזיר
 * אליה אחרי שורה קצרה. השאלה היא מה קורה לה כשאנחנו כותבים בחירה באמצע: אם
 * היא נדרסת, העקיפה חייבת להחזיק עמודה משלה ולהמשיך לטפל בהקשות הבאות; אם היא
 * שורדת, אפשר למסור למנוע כל הקשה שהוא צודק בה.
 *
 *   node scripts/qa/rtl-vertical-seam-probe.mjs      (QA_PORT עוקף 9718)
 */
import { openApp, createReport } from './harness.mjs';
import { zipStored } from './docx-fixtures.mjs';

const W =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const RTL = '<w:bidi/>';

const p1 = (text, pPr = '', rPr = '<w:rtl/>') =>
  `<w:p>${pPr ? `<w:pPr>${pPr}</w:pPr>` : ''}<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}` +
  `<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

const stylesXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${W}>
<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
</w:styles>`;

function buildDocx(bodyXml) {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${W}><w:body>${bodyXml}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:bidi/></w:sectPr></w:body></w:document>`;
  return zipStored({
    '[Content_Types].xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
    '_rels/.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
    'word/document.xml': documentXml,
    'word/styles.xml': stylesXml(),
  });
}

const HE_WRAP =
  'זהו משפט ארוך בעברית שנועד למלא את רוחב השורה כולו ולגלוש אל השורה הבאה כדי ' +
  'שאפשר יהיה למדוד לאן בדיוק החץ מוריד את הסמן כאשר השורה שמתחת קצרה יותר. סוף';
const HE_LONG = 'שורה עברית ארוכה שממלאת כמעט את כל רוחב העמוד מצד אחד אל הצד השני ואינה גולשת';
const HE_SHORT = 'שורה קצרה';
const HE_LONG2 = 'שורה עברית נוספת שגם היא ממלאת כמעט את כל רוחב העמוד ואינה גולשת אל הבאה';

const docx = buildDocx([p1(HE_WRAP, RTL), p1(HE_LONG, RTL), p1(HE_SHORT, RTL), p1(HE_LONG2, RTL)].join(''));
const report = createReport('תפר גלישה ועמודת מטרה אחרי כתיבה');
const app = await openApp({ name: 'rtl-seam', port: Number(process.env.QA_PORT ?? 9718) });

async function openFixture() {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(docx).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-seam',url:${JSON.stringify(dataUrl)},name:'rtl-seam.docx',size:${docx.length},access:'readwrite'}})}`,
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
  return (await app.paraCount()) >= 4;
}

const caretRect = () =>
  app
    .js(`(function(){
      var el = document.querySelector('.sd-v2-local-selection-caret');
      if (!el) return JSON.stringify({ x: null, y: null });
      var r = el.getBoundingClientRect();
      return JSON.stringify({ x: Math.round(r.left*10)/10, y: Math.round((r.top+r.bottom)/2*10)/10 });
    })()`)
    .then(JSON.parse);

const engineCaret = () =>
  app
    .js(`(async function(){
      try {
        var ed = window.__otzariaEditor.superdoc.activeEditor;
        var info = await ed.doc.selection.current();
        var seg = info && info.target && info.target.segments && info.target.segments[0];
        return JSON.stringify({ blockId: seg ? seg.blockId : null, start: seg && seg.range ? seg.range.start : null });
      } catch (e) { return JSON.stringify({ error: String(e) }); }
    })()`)
    .then(JSON.parse);

async function caretState() {
  const [engine, rect] = await Promise.all([engineCaret(), caretRect()]);
  return { block: engine.blockId, off: engine.start, x: rect.x, y: rect.y, err: engine.error };
}

/** השורות של ה-fragment שמכיל את הטקסט, עם טווחי ה-pm ומלבן הטקסט שלהן. */
const fragmentLines = (find) =>
  app
    .js(`(function(){
      var frags = document.querySelectorAll('[data-source-node-id][data-pm-start]');
      for (var i = 0; i < frags.length; i++) {
        if ((frags[i].textContent || '').indexOf(${JSON.stringify(find)}) < 0) continue;
        var out = [];
        var kids = frags[i].children;
        for (var j = 0; j < kids.length; j++) {
          var ps = Number(kids[j].getAttribute('data-pm-start'));
          var pe = Number(kids[j].getAttribute('data-pm-end'));
          if (!isFinite(ps) || !isFinite(pe)) continue;
          var runs = kids[j].querySelectorAll('.superdoc-text-run, .superdoc-tab');
          var left = Infinity, right = -Infinity, top = 0, bottom = 0;
          for (var k = 0; k < runs.length; k++) {
            var r = runs[k].getBoundingClientRect();
            if (!r.width) continue;
            left = Math.min(left, r.left); right = Math.max(right, r.right);
            top = r.top; bottom = r.bottom;
          }
          out.push({ pmStart: ps, pmEnd: pe, rtl: kids[j].getAttribute('dir') === 'rtl',
                     left: Math.round(left*10)/10, right: Math.round(right*10)/10,
                     y: Math.round((top+bottom)/2*10)/10 });
        }
        return JSON.stringify({ base: Number(frags[i].getAttribute('data-pm-start')),
                                nodeId: frags[i].getAttribute('data-source-node-id'), lines: out });
      }
      return JSON.stringify(null);
    })()`)
    .then(JSON.parse);

/** כותבת בחירה מכווצת דרך אותו מסלול שהעקיפה תשתמש בו. */
const writeCaret = (blockId, offset) =>
  app.js(`(function(){
    try {
      var ed = window.__otzariaEditor.superdoc.activeEditor;
      var snap = ed.host.readLiveSelectionSyncSnapshot();
      var sel = snap && snap.selectionTarget;
      var head = sel && sel.end;
      var story = (sel && sel.story) || (head && head.story);
      var point = { kind: 'text', blockId: ${JSON.stringify(blockId)}, offset: ${offset}, story: story };
      var input = { target: { kind: 'selection', start: point, end: point, story: story }, focus: true };
      if (sel && typeof sel.coordinateSpace === 'string') input.target.coordinateSpace = sel.coordinateSpace;
      ed.authoring.setSelectionTarget(input);
      return 'ok';
    } catch (e) { return 'ERR ' + String(e); }
  })()`);

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1500, height: 1000, deviceScaleFactor: 1, mobile: false });
  await app.sleep(400);
  if (!(await openFixture())) {
    report.fail('פתיחת הקובץ', 'המסמך לא נפתח');
    throw new Error('no doc');
  }

  /* -------- 1. תפר גלישה: כתיבה להיסט שהוא סוף שורה וגם תחילת הבאה -------- */
  console.log('\n---- 1. תפר גלישה ----');
  const wrap = await fragmentLines('זהו משפט ארוך בעברית');
  if (!wrap || wrap.lines.length < 2) {
    report.stuck('תפר גלישה', 'הפסקה לא גלשה לשתי שורות');
  } else {
    const [first, second] = wrap.lines;
    console.log(`   שורה 1: pm ${first.pmStart}..${first.pmEnd} y=${first.y} [${first.left}..${first.right}]`);
    console.log(`   שורה 2: pm ${second.pmStart}..${second.pmEnd} y=${second.y} [${second.left}..${second.right}]`);
    if (first.pmEnd !== second.pmStart) {
      report.skip('תפר גלישה', `אין תפר משותף (${first.pmEnd} מול ${second.pmStart}) — השאלה אינה קיימת`);
    } else {
      // מתחילים רחוק משם, כדי שהנחיתה לא תהיה במקרה במקום שכבר היינו בו.
      await app.clickAt(Math.round((second.left + second.right) / 2), Math.round(second.y));
      await app.sleep(500);
      const before = await caretState();
      const seam = first.pmEnd - wrap.base;
      console.log(`   כותבים היסט ${seam} (התפר). לפני: ${before.off}@${before.x},${before.y}`);
      console.log(`   ${await writeCaret(wrap.nodeId, seam)}`);
      await app.sleep(700);
      const after = await caretState();
      const onFirst = Math.abs(after.y - first.y) < Math.abs(after.y - second.y);
      console.log(`   אחרי: היסט ${after.off} ב-${after.x},${after.y} → ${onFirst ? 'השורה הראשונה (סופה)' : 'השורה השנייה (תחילתה)'}`);
      const detail = `היסט ${seam} צויר ב-y=${after.y}; שורה 1 ב-${first.y}, שורה 2 ב-${second.y}`;
      if (onFirst) report.fail('התפר שייך לשורה הקודמת — „תחילת שורת המשך” אינה ניתנת לכתיבה', detail);
      else report.pass('התפר נכתב כתחילת השורה השנייה — אין מלכודת', detail);
    }
  }

  /* -------- 2. עמודת המטרה אחרי כתיבה שלנו -------- */
  console.log('\n---- 2. עמודת המטרה אחרי כתיבה ----');
  const longLines = await fragmentLines('שורה עברית ארוכה שממלאת');
  const shortLines = await fragmentLines('שורה קצרה');
  if (!longLines || !shortLines) {
    report.stuck('עמודת המטרה אחרי כתיבה', 'לא נמצאו הפסקאות');
  } else {
    const src = longLines.lines[0];
    const shortLine = shortLines.lines[0];
    // סמן עמוק בתוך השורה הארוכה — עמודה שהשורה הקצרה אינה מגיעה אליה.
    const x0 = Math.round(src.right - (src.right - src.left) * 0.75);
    await app.clickAt(x0, Math.round(src.y));
    await app.sleep(500);
    const start = await caretState();
    // מדמים בדיוק את מה שהעקיפה תעשה: כותבים את **סוף** השורה הקצרה.
    const endOffset = shortLine.pmEnd - shortLines.base;
    console.log(`   סמן ב-${start.x}; כותבים לסוף השורה הקצרה (היסט ${endOffset})`);
    console.log(`   ${await writeCaret(shortLines.nodeId, endOffset)}`);
    await app.sleep(700);
    const written = await caretState();
    await app.press('ArrowDown', 'ArrowDown', 40);
    await app.sleep(400);
    const next = await caretState();
    console.log(`   אחרי הכתיבה: ${written.off}@${written.x}; אחרי ArrowDown: ${next.off}@${next.x}`);
    const keptOriginal = typeof next.x === 'number' && Math.abs(next.x - start.x) <= 8;
    const followedWrite = typeof next.x === 'number' && typeof written.x === 'number' && Math.abs(next.x - written.x) <= 8;
    const detail = `${start.x} → (כתיבה) ${written.x} → ${next.x}`;
    if (keptOriginal) report.pass('עמודת המטרה שורדת כתיבה — אפשר למסור למנוע', detail);
    else if (followedWrite) report.fail('הכתיבה דורסת את עמודת המטרה — העקיפה חייבת להחזיק עמודה משלה', detail);
    else report.partial('עמודת המטרה אחרי כתיבה — לא זו ולא זו', detail);
  }
} finally {
  app.close();
}

report.print();
