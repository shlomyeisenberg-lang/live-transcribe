/**
 * גשש: חצים אנכיים בשורה עברית — „יורד לשורה קצרה והסמן קופץ לתחילתה”.
 *
 * ## הדיווח
 *
 * שתי שורות זו מתחת לזו. מהשורה הארוכה יורדים בחץ למטה: כשהשורה שמתחת ארוכה
 * לפחות כמו זו שמעליה הסמן נוחת במקום הנכון, אבל כשהיא **קצרה יותר** הוא
 * קופץ ל**תחילת** השורה — בעברית זה הקצה הימני — במקום להישאר בקצה השמאלי,
 * שהוא סוף השורה ומה ש-Word עושה.
 *
 * ## למה גשש
 *
 * „קופץ לתחילת השורה” יכול להיות שלושה דברים שונים לגמרי, ורק מספרים מפרידים
 * ביניהם: (א) היעד נכון והציור שגוי, (ב) עמודת המטרה נשמרת אבל התרגום שלה
 * להיסט מקצץ לפי סדר לוגי, (ג) החץ בכלל אינו אנכי אלא נופל למסלול אחר. לכן
 * נמדדים כאן שלושה מספרים אחרי **כל** הקשה: ההיסט שהמנוע מדווח, ה-x שבו הוא
 * צייר את הסמן, ומלבן שורת היעד — ומהם נגזר אם הסמן על הקצה הלוגי הראשון של
 * השורה או האחרון.
 *
 * ## מה מפריד את החשודים
 *
 * לכל מקרה נמדדות **חמש נקודות התחלה** לאורך השורה שמעליה, מקצה לקצה. אם
 * הקפיצה תלויה בשאלה אם עמודת המטרה נופלת מעבר לסוף שורת היעד, היא תופיע רק
 * בנקודות שמעבר לו — וזה מפריד „קיצוץ לפי גאומטריה” מ„קיצוץ לפי היסט”.
 *
 * ובקרה **LTR**: אותו מבנה בדיוק באנגלית. אם גם שם הסמן קופץ לתחילת השורה,
 * זו אינה תקלת RTL אלא תקלת ניווט אנכי כללית — והניסוח של הדיווח למעלה תלוי
 * בזה.
 *
 * שני כיוונים: חץ למטה אל שורה קצרה, וחץ למעלה אל אותה שורה קצרה.
 *
 *   node scripts/qa/rtl-vertical-caret-probe.mjs      (QA_PORT עוקף 9717)
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

/* פסקה עברית שגולשת: השורה הראשונה מלאה, האחרונה קצרה. */
const HE_WRAP =
  'זהו משפט ארוך בעברית שנועד למלא את רוחב השורה כולו ולגלוש אל השורה הבאה כדי ' +
  'שאפשר יהיה למדוד לאן בדיוק החץ מוריד את הסמן כאשר השורה שמתחת קצרה יותר. סוף';

const HE_LONG = 'שורה עברית ארוכה שממלאת כמעט את כל רוחב העמוד מצד אחד אל הצד השני ואינה גולשת';
const HE_SHORT = 'שורה קצרה';
const HE_LONG2 = 'שורה עברית נוספת שגם היא ממלאת כמעט את כל רוחב העמוד ואינה גולשת אל הבאה';

const EN_WRAP =
  'This is a long English sentence meant to fill the whole width of the line and ' +
  'then wrap onto the next one so the landing spot can be measured. End';

const EN_LONG = 'A long English line that fills nearly the entire width of the page without wrapping';
const EN_SHORT = 'Short line';
const EN_LONG2 = 'Another long English line that also fills nearly the entire page width here';
const HE_LONG3 = 'עוד שורה עברית ארוכה שממלאת כמעט את כל רוחב העמוד ואינה גולשת אל הבאה';
const HE_INDENT = 'קצרה ומוזחת';
const EN_LONG3 = 'Yet another long English line filling nearly the whole width of the page';
const EN_INDENT = 'Short indented';
const HE_LONG4 = 'ועוד שורה עברית ארוכה שממלאת כמעט את כל רוחב העמוד ואינה גולשת';
const HE_CENTER = 'קצרה וממורכזת';
const EN_LONG4 = 'One more long English line that fills nearly the whole page width again';
const EN_CENTER = 'Short centered';

const BODY = [
  p1(HE_WRAP, RTL),
  p1(HE_LONG, RTL),
  p1(HE_SHORT, RTL),
  p1(HE_LONG2, RTL),
  p1(EN_WRAP, '', ''),
  p1(EN_LONG, '', ''),
  p1(EN_SHORT, '', ''),
  p1(EN_LONG2, '', ''),
  p1(HE_LONG3, RTL),
  p1(HE_INDENT, RTL + '<w:ind w:right="3600"/>'),
  p1(EN_LONG3, '', ''),
  p1(EN_INDENT, '<w:ind w:left="3600"/>', ''),
  p1(HE_LONG4, RTL),
  p1(HE_CENTER, RTL + '<w:jc w:val="center"/>'),
  p1(EN_LONG4, '', ''),
  p1(EN_CENTER, '<w:jc w:val="center"/>', ''),
].join('');

const docx = buildDocx(BODY);
const report = createReport('חצים אנכיים: נחיתה בשורה קצרה יותר');
const app = await openApp({ name: 'rtl-vert', port: Number(process.env.QA_PORT ?? 9717) });

/** „פתח קובץ” פותח דיאלוג; הבורר נקרא רק מ„עיון בקבצים…”. */
async function openFixture() {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(docx).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-vert',url:${JSON.stringify(dataUrl)},name:'rtl-vertical.docx',size:${docx.length},access:'readwrite'}})}`,
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
  return paras >= 16;
}

/**
 * כל השורות המצוירות במסמך, מקובצות לפי ה-y של הריצות.
 *
 * שורה חזותית אינה אלמנט — היא קבוצת `.superdoc-text-run` באותו גובה, ולכן
 * הקיבוץ הוא לפי `top` עם סבילות של 4px (עיגול תת-פיקסל בלבד).
 */
const allLines = () =>
  app
    .js(`(function(){
      var out = [];
      var frags = document.querySelectorAll('.superdoc-fragment');
      for (var i = 0; i < frags.length; i++) {
        var nodeId = frags[i].getAttribute('data-source-node-id');
        var runs = frags[i].querySelectorAll('.superdoc-text-run');
        var lines = [];
        for (var j = 0; j < runs.length; j++) {
          var r = runs[j].getBoundingClientRect();
          if (!r.width || !r.height) continue;
          var hit = null;
          for (var k = 0; k < lines.length; k++) {
            if (Math.abs(lines[k].top - r.top) < 4) { hit = lines[k]; break; }
          }
          if (!hit) { hit = { top: r.top, bottom: r.bottom, left: r.left, right: r.right, text: '' }; lines.push(hit); }
          hit.left = Math.min(hit.left, r.left);
          hit.right = Math.max(hit.right, r.right);
          hit.bottom = Math.max(hit.bottom, r.bottom);
          hit.text += runs[j].textContent || '';
        }
        for (var m = 0; m < lines.length; m++) {
          lines[m].nodeId = nodeId;
          lines[m].left = Math.round(lines[m].left * 10) / 10;
          lines[m].right = Math.round(lines[m].right * 10) / 10;
          lines[m].top = Math.round(lines[m].top * 10) / 10;
          lines[m].bottom = Math.round(lines[m].bottom * 10) / 10;
          out.push(lines[m]);
        }
      }
      out.sort(function(a, b) { return a.top - b.top; });
      return JSON.stringify(out);
    })()`)
    .then(JSON.parse);

/** הסמן המצויר של המנוע. */
const caretRect = () =>
  app
    .js(`(function(){
      var el = document.querySelector('.sd-v2-local-selection-caret');
      if (!el) return JSON.stringify({ x: null, y: null });
      var r = el.getBoundingClientRect();
      return JSON.stringify({ x: Math.round(r.left*10)/10, y: Math.round((r.top + r.bottom)/2*10)/10 });
    })()`)
    .then(JSON.parse);

/** ההיסט הלוגי מהמנוע: בלוק + מספר תווים מתחילתו. */
const engineCaret = () =>
  app
    .js(`(async function(){
      try {
        var ed = window.__otzariaEditor && window.__otzariaEditor.superdoc && window.__otzariaEditor.superdoc.activeEditor;
        if (!ed || !ed.doc) return JSON.stringify({ error: 'no editor' });
        var info = await ed.doc.selection.current();
        var seg = info && info.target && info.target.segments && info.target.segments[0];
        return JSON.stringify({
          blockId: seg ? seg.blockId : null,
          start: seg && seg.range ? seg.range.start : null,
          end: seg && seg.range ? seg.range.end : null
        });
      } catch (e) { return JSON.stringify({ error: String(e) }); }
    })()`)
    .then(JSON.parse);

async function caretState() {
  const [engine, rect] = await Promise.all([engineCaret(), caretRect()]);
  return { block: engine.blockId, off: engine.start, end: engine.end, x: rect.x, y: rect.y, err: engine.error };
}

/** השורה שהסמן נמצא בתוכה, לפי ה-y שלו. */
function lineAt(lines, y) {
  if (typeof y !== 'number') return null;
  for (const line of lines) {
    if (y >= line.top - 6 && y <= line.bottom + 6) return line;
  }
  return null;
}

/**
 * איפה הסמן יושב ביחס לשורה: בקצה הלוגי הראשון, האחרון, או באמצע.
 *
 * ב-RTL הקצה הלוגי הראשון הוא ה-`right` של השורה, והאחרון הוא ה-`left`.
 */
/**
 * עמודת המטרה ביחס לשורת היעד: לפני תחילתה, בתוכה, או מעבר לסופה.
 *
 * שלושה מצבים ולא שניים, כי הנכון שונה בכל אחד מהם: לפני התחילה —
 * תחילת השורה, מעבר לסוף — סוף השורה, ובתוך — המקום שמתחת לעמודה.
 * שורה מוזחת היא המקרה היחיד שבו המצב הראשון בכלל אפשרי.
 */
function sideOf(line, x, rtl) {
  if (!line || typeof x !== 'number') return '?';
  const startX = rtl ? line.right : line.left;
  const endX = rtl ? line.left : line.right;
  if (rtl ? x > startX + 1 : x < startX - 1) return 'לפני תחילת השורה';
  if (rtl ? x < endX - 1 : x > endX + 1) return 'מעבר לסוף השורה';
  return 'בתוך השורה';
}

function edgeOf(line, x, rtl) {
  if (!line || typeof x !== 'number') return '?';
  const startX = rtl ? line.right : line.left;
  const endX = rtl ? line.left : line.right;
  if (Math.abs(x - startX) <= 3) return 'תחילת השורה';
  if (Math.abs(x - endX) <= 3) return 'סוף השורה';
  return 'באמצע';
}

/**
 * מקרה אחד: מהשורה שמעליה (או שמתחתיה) אל השורה הקצרה, מחמש נקודות התחלה.
 */
async function measure({ label, sourceText, sourceLineIndex, key, vk, targetText, rtl }) {
  const lines = await allLines();
  const source = lines.filter((l) => l.nodeId && l.text.indexOf(sourceText) >= 0);
  const sourceLine =
    sourceLineIndex === 'last'
      ? source[source.length - 1]
      : sourceLineIndex === 'first'
        ? source[0]
        : source[sourceLineIndex];
  if (!sourceLine) throw new Error(`לא נמצאה שורת מקור ל„${sourceText.slice(0, 20)}”`);

  const targetLine = lines.find((l) => l.text.indexOf(targetText) >= 0);
  if (!targetLine) throw new Error(`לא נמצאה שורת יעד ל„${targetText.slice(0, 20)}”`);

  const width = sourceLine.right - sourceLine.left;
  const targetEndX = rtl ? targetLine.left : targetLine.right;
  const rows = [];

  for (const frac of [0.05, 0.25, 0.5, 0.75, 0.95]) {
    const x = Math.round(sourceLine.left + width * frac);
    const y = Math.round((sourceLine.top + sourceLine.bottom) / 2);
    await app.clickAt(x, y);
    await app.sleep(500);
    const before = await caretState();
    if (before.err) throw new Error(`קריאת הסמן נכשלה: ${before.err}`);

    await app.press(key, key, vk);
    await app.sleep(350);
    const after = await caretState();

    // הגאומטריה נלכדת מחדש: הקשה עלולה לגלול, והמלבנים הם מלבני מסך.
    const nowLines = await allLines();
    const landedOn = lineAt(nowLines, after.y);
    const nowTarget = nowLines.find((l) => l.text.indexOf(targetText) >= 0) ?? targetLine;

    const side = sideOf(nowTarget, before.x, rtl);
    const inTarget = landedOn && landedOn.text.indexOf(targetText) >= 0;
    const edge = inTarget ? edgeOf(nowTarget, after.x, rtl) : '(לא בשורת היעד)';

    rows.push({ frac, fromX: before.x, toX: after.x, off: after.off, side, inTarget, edge });
    console.log(
      `   ${String(Math.round(frac * 100)).padStart(3)}% : x ${before.x} → ${after.x}` +
        ` | היסט ${after.off} | עמודת המטרה ${side} | נחת ${edge}`,
    );
  }

  console.log(
    `   שורת המקור [${sourceLine.left}..${sourceLine.right}] ; שורת היעד [${targetLine.left}..${targetLine.right}], סוף לוגי ב-${targetEndX}`,
  );
  return rows;
}

const CASES = [
  {
    label: 'עברית: חץ למטה בתוך פסקה גולשת (שורה ראשונה → אחרונה, קצרה)',
    sourceText: 'זהו משפט ארוך בעברית',
    sourceLineIndex: 'first',
    key: 'ArrowDown',
    vk: 40,
    targetText: 'סוף',
    rtl: true,
  },
  {
    label: 'עברית: חץ למטה מפסקה ארוכה לפסקה קצרה',
    sourceText: 'שורה עברית ארוכה שממלאת',
    sourceLineIndex: 0,
    key: 'ArrowDown',
    vk: 40,
    targetText: 'שורה קצרה',
    rtl: true,
  },
  {
    label: 'עברית: חץ למעלה מפסקה ארוכה לפסקה קצרה',
    sourceText: 'שורה עברית נוספת שגם היא',
    sourceLineIndex: 0,
    key: 'ArrowUp',
    vk: 38,
    targetText: 'שורה קצרה',
    rtl: true,
  },
  {
    label: 'בקרה LTR: חץ למטה בתוך פסקה גולשת',
    sourceText: 'This is a long English',
    sourceLineIndex: 'first',
    key: 'ArrowDown',
    vk: 40,
    targetText: 'End',
    rtl: false,
  },
  {
    label: 'בקרה LTR: חץ למטה מפסקה ארוכה לפסקה קצרה',
    sourceText: 'A long English line that fills',
    sourceLineIndex: 0,
    key: 'ArrowDown',
    vk: 40,
    targetText: 'Short line',
    rtl: false,
  },
  {
    label: 'בקרה LTR: חץ למעלה מפסקה ארוכה לפסקה קצרה',
    sourceText: 'Another long English line',
    sourceLineIndex: 0,
    key: 'ArrowUp',
    vk: 38,
    targetText: 'Short line',
    rtl: false,
  },
  {
    label: 'עברית: חץ למטה לשורה מוזחת (עמודת המטרה לפני תחילתה)',
    sourceText: 'עוד שורה עברית ארוכה',
    sourceLineIndex: 0,
    key: 'ArrowDown',
    vk: 40,
    targetText: 'קצרה ומוזחת',
    rtl: true,
  },
  {
    label: 'בקרה LTR: חץ למטה לשורה מוזחת',
    sourceText: 'Yet another long English line',
    sourceLineIndex: 0,
    key: 'ArrowDown',
    vk: 40,
    targetText: 'Short indented',
    rtl: false,
  },
  {
    label: 'עברית: חץ למטה לשורה ממורכזת (שני הצדדים)',
    sourceText: 'ועוד שורה עברית ארוכה',
    sourceLineIndex: 0,
    key: 'ArrowDown',
    vk: 40,
    targetText: 'קצרה וממורכזת',
    rtl: true,
  },
  {
    label: 'בקרה LTR: חץ למטה לשורה ממורכזת',
    sourceText: 'One more long English line',
    sourceLineIndex: 0,
    key: 'ArrowDown',
    vk: 40,
    targetText: 'Short centered',
    rtl: false,
  },
];

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', {
    width: 1500,
    height: 1100,
    deviceScaleFactor: 1,
    mobile: false,
  });
  await app.sleep(400);
  if (!(await openFixture())) {
    report.fail('פתיחת הקובץ', 'המסמך לא נפתח');
    throw new Error('no doc');
  }

  const geom = await allLines();
  console.log('\n---- השורות שצוירו ----');
  for (const line of geom) {
    console.log(`   [${line.left}..${line.right}] y=${line.top} „${line.text.slice(0, 40)}”`);
  }

  for (const c of CASES) {
    console.log(`\n== ${c.label}`);
    let rows;
    try {
      rows = await measure(c);
    } catch (e) {
      report.stuck(c.label, String(e.message || e).slice(0, 100));
      continue;
    }

    const landed = rows.filter((r) => r.inTarget);
    if (!landed.length) {
      report.stuck(c.label, 'אף הקשה לא נחתה בשורת היעד');
      continue;
    }
    /*
     * מה נכון לכל צד: עמודת מטרה שנופלת מעבר לסוף השורה צריכה לנחות
     * בסוףה, ועמודה שלפני תחילתה — בתחילתה. זה מה ש-Word עושה, וזה
     * מה שהבקרה הלועזית מראה שהמנוע עושה בעצמו בפסקה LTR.
     */
    const want = {
      'מעבר לסוף השורה': 'סוף השורה',
      'לפני תחילת השורה': 'תחילת השורה',
    };
    const outside = landed.filter((r) => r.side !== 'בתוך השורה');
    const wrong = outside.filter((r) => r.edge !== want[r.side]);
    const inside = landed.filter((r) => r.side === 'בתוך השורה');
    const insideWrong = inside.filter((r) => Math.abs((r.toX ?? 0) - (r.fromX ?? 0)) > 8);

    const detail =
      `מחוץ לשורה: ${outside.length} נקודות, ${wrong.length} נחתו בקצה ההפוך` +
      (wrong.length ? ` (${wrong.map((r) => `${Math.round(r.frac * 100)}%→${r.edge}`).join(', ')})` : '') +
      `; בתוך השורה: ${inside.length} נקודות, ${insideWrong.length} שגויות`;

    if (wrong.length > 0) report.fail(c.label, detail);
    else if (!outside.length) report.skip(c.label, 'אף נקודת התחלה לא נפלה מחוץ לשורת היעד');
    else if (insideWrong.length) report.partial(c.label, detail);
    else report.pass(c.label, detail);
  }

  /* -------- עמודת המטרה: האם היא שורדת שורה קצרה -------- */
  /*
   * זה אינו קישוט. תיקון שכותב בחירה דרך ה-API מחליף את מה שהמנוע מחזיק,
   * ולכן חייבים לדעת מה הוא מחזיק: אם הוא זוכר את העמודה המקורית ומחזיר
   * אליה אחרי השורה הקצרה — כמו Word — אז עקיפה ש„רק מתקנת את הנחיתה”
   * תשבור את ההקשה הבאה; ואם הוא מאפס אותה בכל הקשה, אין מה לשמר.
   *
   * הבקרה הלועזית היא התשובה האמיתית: שם הנחיתה תקינה, ולכן מה שנמדד שם
   * הוא **הכוונה** של המנוע ולא תוצר של הבאג.
   */
  console.log('\n---- עמודת המטרה אחרי שורה קצרה ----');
  const GOAL_CASES = [
    { label: 'עברית: ארוכה → קצרה → ארוכה', source: 'שורה עברית ארוכה שממלאת', third: 'שורה עברית נוספת שגם היא', rtl: true },
    { label: 'בקרה LTR: ארוכה → קצרה → ארוכה', source: 'A long English line that fills', third: 'Another long English line', rtl: false },
  ];
  for (const c of GOAL_CASES) {
    try {
      const lines = await allLines();
      const sourceLine = lines.find((l) => l.text.indexOf(c.source) >= 0);
      if (!sourceLine) throw new Error('אין שורת מקור');
      // רבע מקצה ההתחלה: עמוק בתוך השורה הארוכה, והרחק מקצה השורה הקצרה.
      const width = sourceLine.right - sourceLine.left;
      const x = Math.round(c.rtl ? sourceLine.right - width * 0.75 : sourceLine.left + width * 0.75);
      const y = Math.round((sourceLine.top + sourceLine.bottom) / 2);
      await app.clickAt(x, y);
      await app.sleep(500);
      const start = await caretState();
      await app.press('ArrowDown', 'ArrowDown', 40);
      await app.sleep(350);
      const mid = await caretState();
      await app.press('ArrowDown', 'ArrowDown', 40);
      await app.sleep(350);
      const end = await caretState();

      const nowLines = await allLines();
      const thirdLine = nowLines.find((l) => l.text.indexOf(c.third) >= 0);
      const back = thirdLine && lineAt(nowLines, end.y) && lineAt(nowLines, end.y).text.indexOf(c.third) >= 0;
      const kept = typeof start.x === 'number' && typeof end.x === 'number' && Math.abs(end.x - start.x) <= 8;

      console.log(`\n== ${c.label}`);
      console.log(`   x: התחלה ${start.x} → אחרי הקשה ראשונה ${mid.x} → אחרי השנייה ${end.x}`);
      console.log(`   ${back ? 'חזר לשורה השלישית' : 'לא הגיע לשורה השלישית'}; ${kept ? 'העמודה נשמרה' : 'העמודה לא נשמרה'}`);

      const detail = `${start.x} → ${mid.x} → ${end.x}`;
      if (!back) report.stuck(`${c.label} — עמודת המטרה`, `לא הגיע לשורה השלישית (${detail})`);
      else if (kept) report.pass(`${c.label} — עמודת המטרה נשמרת`, detail);
      else report.fail(`${c.label} — עמודת המטרה אינה נשמרת`, detail);
    } catch (e) {
      report.stuck(`${c.label} — עמודת המטרה`, String(e.message || e).slice(0, 100));
    }
  }
} finally {
  app.close();
}

report.print();
