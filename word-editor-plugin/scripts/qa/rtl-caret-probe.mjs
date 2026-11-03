/**
 * גשש: תזוזת הסמן בחצים ימין/שמאל בטקסט עברי — הדיווח „לוחץ ימין והסמן הולך
 * שמאלה, בטקסט רגיל תקין וברשימה הפוך, ולפעמים נתקע".
 *
 * ## למה גשש, ולמה בכלל למדוד
 *
 * אין בקוד — לא שלנו ולא ב-SuperDoc הפתוח — שום מאזין לחצים בתוך העורך:
 * `ArrowRight` מופיע בבאנדל שלהם רק ברצועת הכלים. גם `contenteditable` אינו
 * קיים בשרשרת של הטקסט המצויר, והבחירה הנייטיבית יושבת על אלמנט מסך נסתר
 * (`.v2-super-editor__stage`, שם תלויים `keydown`/`beforeinput` בשלב capture).
 * כלומר הסמן הוא של **המנוע**: אין כאן CSS לתקן, והשאלה היחידה שאפשר לענות
 * עליה מכאן היא מה בדיוק מכתיב את הכיוון — ועל זה הגשש הזה עונה במספרים.
 *
 * ## מה יצא, בשורה אחת
 *
 * המנוע בוחר **מודל כיוון לכל פסקה בנפרד**, ולכן אותו מקש עושה שני דברים
 * הפוכים באותו מסמך:
 *
 *   | הפסקה (כולן עברית עם `w:bidi`)      | `ArrowRight`      | המודל  |
 *   |-------------------------------------|-------------------|--------|
 *   | פסקה רגילה                          | 9 → 8 → 7 → 6     | חזותי  |
 *   | פריט רשימה (מספור/תבליט/`hebrew1`)  | 6 → 7 → 8 → 9     | לוגי   |
 *   | רשימה עם `w:suff` רווח / כלום       | 6 → 7 → 8 → 9     | לוגי   |
 *   | פסקה עם `w:tab`                     | 2 → 3 → 4 → 5     | לוגי   |
 *   | פסקה עם ספרות או לטינית             | 2 → 3 → 4 → 5     | לוגי   |
 *   | עברית בלי `w:bidi`                  | קדימה             | לוגי   |
 *   | בקרה: הזחת רשימה **בלי** מספור      | 7 → 6 → 5 → 4     | חזותי  |
 *   | בקרה: שלושה מקטעי עיצוב עבריים      | 2 → 1 → 0         | חזותי  |
 *
 * שתי הבקרות האחרונות מפרפרות את שני החשודים המתבקשים: **ההזחה אינה הגורם**
 * ו**מספר מקטעי העיצוב אינו הגורם**. מה שמהפך הוא היות הפסקה פריט רשימה —
 * בכל סמן ובכל מפריד — או תוכן מעורב-כיוון בשורה.
 *
 * ו„נתקע" אינו באג נפרד אלא התוצאה: בגבול בין פסקה „חזותית" לפסקה „לוגית"
 * כל צד שולח את הסמן בחזרה לשני, והוא מקפץ בין השתיים לנצח —
 * `list:0 → paragraph:16 → list:0 → …`. אין מוצא בחצים, רק בעכבר.
 *
 * דווח למעלה: superdoc/docx-editor#3996 (ותיקון לניסוח ב-#3986, שם נכתב
 * בטעות שכל החצים תקינים — זה נכון לפסקאות רגילות בלבד).
 *
 * ## אם יוחלט לעקוף אצלנו
 *
 * יש לזה תבנית מוכחת: `src/engine/rtl-line-end.ts` (העקיפה של `End`) מיירט
 * `keydown` על ה-div שלנו לפני שהאירוע מגיע למנוע, קורא את הסמן
 * **סינכרונית** ב-`readLiveSelectionSyncSnapshot()` ומזיז אותו ב-
 * `authoring.setSelectionTarget`. הקושי כאן אינו המנגנון אלא ההחלטה: היפוך
 * הדלתא נכון לשורה עברית **אחידה**, ואינו נכון לשורה עם ספרות או לטינית —
 * שם תנועה חזותית אמיתית דורשת את סדר ה-bidi של השורה, שהמנוע אינו חושף.
 *
 * ## מה נמדד
 *
 * לכל שילוב, אחרי **כל** הקשה: מזהה הבלוק, ההיסט הלוגי בתוך הבלוק
 * (`selection.current()`), ו-x של הסמן המצויר (`.sd-v2-local-selection-caret`).
 * שלושתם יחד מבדילים בין שלוש התנהגויות שנראות דומות למשתמש:
 *   • **חזותי** — ימין מזיז ימינה על המסך, וההיסט יורד (כי הטקסט עברי). Word.
 *   • **לוגי** — ההיסט עולה והסמן זז שמאלה. זה הדיווח.
 *   • **מלכודת** — הסמן מקפץ בין שני בלוקים ואינו מתקדם.
 *
 * ## מה השילובים מפרידים
 *
 * החשודים נבדקים אחד מול השני, ולא „ברשימה זה נשבר": מספור מול תבליט,
 * מספור עם `w:suff` של טאב (ברירת המחדל של Word) מול רווח מול כלום, הזחה
 * בלי מספור בכלל, טאב בפסקה רגילה, ספרות ולטינית בתוך שורה עברית, שלושה
 * מקטעי עיצוב עבריים (בקרה: מספר המקטעים אינו הגורם), ורשימה **לטינית**
 * (בקרה: האם התופעה תלוית-כיוון).
 *
 *   node scripts/qa/rtl-caret-probe.mjs          (QA_PORT עוקף 9711)
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
        `<w:r>${r.rPr ? `<w:rPr>${r.rPr}</w:rPr>` : ''}${
          r.tab ? '<w:tab/>' : `<w:t xml:space="preserve">${r.t}</w:t>`
        }</w:r>`,
    )
    .join('')}</w:p>`;

const p1 = (text, pPr = '', rPr = '<w:rtl/>') => pRuns([{ t: text, rPr }], pPr);

/**
 * רמת מספור אחת. `suff` הוא מה שמפריד את הסמן מהטקסט: `tab` היא ברירת המחדל
 * של Word (וזו הצורה שכל מסמך אמיתי נושא), ולכן שלוש הווריאציות נמדדות זו
 * מול זו.
 */
const lvl = ({ fmt = 'decimal', text = '%1.', suff = null, jc = 'left' }) =>
  `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>` +
  (suff ? `<w:suff w:val="${suff}"/>` : '') +
  `<w:lvlText w:val="${text}"/><w:lvlJc w:val="${jc}"/>` +
  `<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>`;

const numberingXml = () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering ${W}>
<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>${lvl({})}</w:abstractNum>
<w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="hybridMultilevel"/>${lvl({ fmt: 'bullet', text: '•' })}</w:abstractNum>
<w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="hybridMultilevel"/>${lvl({ suff: 'space' })}</w:abstractNum>
<w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="hybridMultilevel"/>${lvl({ suff: 'nothing' })}</w:abstractNum>
<w:abstractNum w:abstractNumId="4"><w:multiLevelType w:val="hybridMultilevel"/>${lvl({})}</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
<w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>
<w:num w:numId="4"><w:abstractNumId w:val="3"/></w:num>
<w:num w:numId="5"><w:abstractNumId w:val="4"/></w:num>
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

const BODY = [
  p1('אבגדהוז ראשון גדול', RTL),
  p1('פסקה באמצע המסמך', RTL),
  p1('עברית בלי הכרזה', '', ''),
  p1('מספור עם טאב', RTL + numPr(1)),
  p1('מספור עם רווח', RTL + numPr(3)),
  p1('מספור בלי מפריד', RTL + numPr(4)),
  p1('תבליט עגול פריט', RTL + numPr(2)),
  p1('הזחה בלי מספור', RTL + '<w:pStyle w:val="ListParagraph"/>'),
  pRuns([{ t: 'קודם', rPr: '<w:rtl/>' }, { tab: true }, { t: 'אחרון', rPr: '<w:rtl/>' }], RTL),
  p1('עברית 123 עברית', RTL),
  p1('עברית ABC עברית', RTL),
  pRuns(
    [
      { t: 'לפני ', rPr: '<w:rtl/>' },
      { t: 'מודגש', rPr: '<w:b/><w:bCs/><w:rtl/>' },
      { t: ' אחרי', rPr: '<w:rtl/>' },
    ],
    RTL,
  ),
  p1('latin list item here', numPr(5), ''),
  p1('alpha beta gamma', '', ''),
].join('');

const CASES = [
  { label: 'בסיס: פסקה עברית ראשונה במסמך', find: 'אבגדהוז ראשון גדול' },
  { label: 'בסיס: פסקה עברית באמצע המסמך', find: 'פסקה באמצע המסמך' },
  { label: 'עברית בלי w:bidi (השורה מוכרזת LTR)', find: 'עברית בלי הכרזה' },
  { label: 'רשימה ממוספרת, מפריד טאב (ברירת המחדל של Word)', find: 'מספור עם טאב' },
  { label: 'רשימה ממוספרת, מפריד רווח (w:suff=space)', find: 'מספור עם רווח' },
  { label: 'רשימה ממוספרת, בלי מפריד (w:suff=nothing)', find: 'מספור בלי מפריד' },
  { label: 'רשימת תבליטים', find: 'תבליט עגול פריט' },
  { label: 'בקרה: הזחת רשימה בלי מספור', find: 'הזחה בלי מספור' },
  { label: 'פסקה עברית עם טאב באמצע', find: 'קודם' },
  { label: 'פסקה עברית עם ספרות באמצע', find: 'עברית 123 עברית' },
  { label: 'פסקה עברית עם לטינית באמצע', find: 'עברית ABC עברית' },
  { label: 'בקרה: שלושה מקטעי עיצוב עבריים', find: 'לפני מודגש אחרי' },
  { label: 'בקרה: רשימה ממוספרת לטינית', find: 'latin list item here' },
  { label: 'בקרה: פסקה לטינית רגילה', find: 'alpha beta gamma' },
];

const docx = buildDocx(BODY);
const report = createReport('חצים אופקיים בעברית — מה מכתיב את הכיוון');
const app = await openApp({ name: 'rtl-caret', port: Number(process.env.QA_PORT ?? 9711) });

/** „פתח קובץ” פותח דיאלוג; הבורר נקרא רק מ„עיון בקבצים…” שבתוכו. */
async function openFixture() {
  const dataUrl =
    'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' +
    Buffer.from(docx).toString('base64');
  await app.js(
    `window.__qaHost.replies['fs.pickUserFile']=function(){return Promise.resolve({success:true,error:null,` +
      `data:{token:'tok-rtl',url:${JSON.stringify(dataUrl)},name:'rtl-caret.docx',size:${docx.length},access:'readwrite'}})}`,
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
  return paras > 10;
}

/** הסמן המצויר של המנוע. */
const caretRect = () =>
  app
    .js(`(function(){
      var el = document.querySelector('.sd-v2-local-selection-caret');
      if (!el) return JSON.stringify({x:null});
      var r = el.getBoundingClientRect();
      return JSON.stringify({ x: Math.round(r.left*10)/10, y: Math.round(r.top*10)/10 });
    })()`)
    .then(JSON.parse);

/** ההיסט הלוגי מהמנוע: בלוק + מספר תווים מתחילתו. */
const engineCaret = () =>
  app
    .js(`(async function(){
      try {
        var ed = window.__otzariaEditor && window.__otzariaEditor.superdoc && window.__otzariaEditor.superdoc.activeEditor;
        if (!ed || !ed.doc) return JSON.stringify({error:'no editor'});
        var info = await ed.doc.selection.current();
        var seg = info && info.target && info.target.segments && info.target.segments[0];
        return JSON.stringify({
          blockId: seg ? seg.blockId : null,
          start: seg && seg.range ? seg.range.start : null,
          end: seg && seg.range ? seg.range.end : null,
          empty: info ? info.empty : null
        });
      } catch (e) { return JSON.stringify({error:String(e)}); }
    })()`)
    .then(JSON.parse);

async function caretState() {
  const [engine, rect] = await Promise.all([engineCaret(), caretRect()]);
  return {
    block: engine.blockId,
    off: engine.start,
    end: engine.end,
    x: rect.x,
    err: engine.error,
  };
}

/**
 * ממקמת את הסמן ב**אמצע הטקסט** של הפסקה, ולא במלבן השורה.
 *
 * `app.caretPara` לוחצת 14px מקצה השורה, וקצה השורה של פריט רשימה הוא אזור
 * ההזחה — הלחיצה נחתה שם על היסט 0 בכל פריטי הרשימה. בלי המדידה הזאת אפשר
 * היה לטעות ולתלות את ההיפוך בנקודת ההתחלה במקום בפסקה עצמה.
 */
async function caretMid(find) {
  const rect = JSON.parse(
    await app.js(`(function(){
      var frags = document.querySelectorAll('.superdoc-fragment');
      for (var i = 0; i < frags.length; i++) {
        if ((frags[i].textContent || '').indexOf(${JSON.stringify(find)}) < 0) continue;
        var runs = frags[i].querySelectorAll('.superdoc-text-run');
        var left = Infinity, right = -Infinity, top = 0, bottom = 0;
        for (var j = 0; j < runs.length; j++) {
          var r = runs[j].getBoundingClientRect();
          if (!r.width) continue;
          left = Math.min(left, r.left); right = Math.max(right, r.right);
          top = r.top; bottom = r.bottom;
        }
        if (right < 0) return JSON.stringify(null);
        return JSON.stringify({
          x: Math.round((left + right) / 2), y: Math.round((top + bottom) / 2),
          nodeId: frags[i].getAttribute('data-source-node-id')
        });
      }
      return JSON.stringify(null);
    })()`),
  );
  if (!rect) throw new Error(`אין טקסט מצויר לפסקה „${find}”`);
  await app.clickAt(rect.x, rect.y);
  await app.sleep(600);
  const state = await caretState();
  if (state.block !== rect.nodeId) {
    throw new Error(`הלחיצה נחתה על בלוק ${state.block} ולא על ${rect.nodeId}`);
  }
  return state;
}

async function walk(key, times) {
  const vk = key === 'ArrowRight' ? 39 : 37;
  const steps = [];
  for (let i = 0; i < times; i++) {
    await app.press(key, key, vk);
    await app.sleep(200);
    steps.push(await caretState());
  }
  return steps;
}

const trace = (steps) =>
  steps.map((s) => `${String(s.block ?? '?').slice(-2)}:${s.off ?? '?'}@${s.x ?? '?'}`).join(' → ');

/**
 * הכיוון שהמנוע נתן להקשות בתוך אותו בלוק, והאם יש מלכודת בגבול.
 *
 * צעדים שחוצים בלוק אינם נספרים ככיוון — הם גבול, וכיוון נמדד רק בתוך פסקה.
 * מלכודת מזוהה בחזרה לאותו (בלוק, היסט) שכבר היינו בו: זה בדיוק מה שהמשתמש
 * מתאר כ„נתקע", ובלי זיהוי מפורש הוא נראה כמו „שתי תזוזות רגילות".
 */
function verdictOf(steps, key) {
  const inside = [];
  for (let i = 1; i < steps.length; i++) {
    const a = steps[i - 1];
    const b = steps[i];
    if (a.block !== b.block) continue;
    if (typeof a.x !== 'number' || typeof b.x !== 'number') continue;
    inside.push(b.x - a.x);
  }
  const moved = inside.filter((dx) => Math.abs(dx) > 0.5);
  const rightward = moved.filter((dx) => dx > 0).length;
  const leftward = moved.filter((dx) => dx < 0).length;

  /*
   * עצירה בקצה המסמך היא אותו (בלוק, היסט) שוב ושוב, אבל אינה מלכודת:
   * אין פסקה שנייה שמחזירה את הסמן. מלכודת הגבול שנמדדה היא מחזור ממשי
   * בין שני בלוקים, למשל list:0 → paragraph:16 → list:0. לכן סופרים רק
   * חזרה של מצב אחרי מעבר לשני בלוק אחר, ולא כל מצב שכבר נראה.
   */
  let loop = 0;
  for (let i = 2; i < steps.length; i++) {
    const a = steps[i - 2];
    const b = steps[i - 1];
    const c = steps[i];
    if (a.block !== b.block && b.block !== c.block && a.block === c.block && a.off === c.off) {
      loop += 1;
    }
  }

  const expectRight = key === 'ArrowRight';
  return {
    rightward,
    leftward,
    frozen: inside.length - moved.length,
    loop,
    visual: expectRight ? rightward > 0 && leftward === 0 : leftward > 0 && rightward === 0,
    logical: expectRight ? leftward > 0 && rightward === 0 : rightward > 0 && leftward === 0,
  };
}

try {
  await app.cdp.send('Emulation.setDeviceMetricsOverride', { width: 1500, height: 1000, deviceScaleFactor: 1, mobile: false });
  await app.sleep(400);
  if (!(await openFixture())) {
    report.fail('פתיחת הקובץ', 'המסמך לא נפתח');
    throw new Error('no doc');
  }

  for (const c of CASES) {
    try {
      await app.caretPara(c.find);
    } catch (e) {
      report.stuck(c.label, `מיקום סמן נכשל: ${String(e.message || e).slice(0, 80)}`);
      continue;
    }
    const rightSteps = [await caretState(), ...(await walk('ArrowRight', 7))];
    await app.caretPara(c.find);
    const leftSteps = [await caretState(), ...(await walk('ArrowLeft', 7))];

    const right = verdictOf(rightSteps, 'ArrowRight');
    const left = verdictOf(leftSteps, 'ArrowLeft');

    console.log(`\n== ${c.label}`);
    console.log(`   ArrowRight: ${trace(rightSteps)}`);
    console.log(`   ArrowLeft : ${trace(leftSteps)}`);

    const shape =
      right.visual && left.visual
        ? 'חזותי (כמו Word)'
        : right.logical && left.logical
          ? 'לוגי — הפוך ממה שנראה על המסך'
          : 'מעורב';
    const trap = right.loop > 1 || left.loop > 1 ? `; מלכודת גבול (חזרות: ימין ${right.loop}, שמאל ${left.loop})` : '';
    const detail = `${shape}${trap}`;

    if (right.logical && left.logical) report.fail(c.label, detail);
    else if (trap) report.partial(c.label, detail);
    else if (right.visual && left.visual) report.pass(c.label, detail);
    else report.partial(c.label, detail);
  }
  /* -------- אותה שאלה, כשהסמן מתחיל באמצע הטקסט -------- */
  console.log('\n---- התחלה באמצע הטקסט (ולא בקצה השורה) ----');
  for (const c of CASES.filter((x) => /רשימה|בסיס|תבליט/.test(x.label))) {
    let start;
    try {
      start = await caretMid(c.find);
    } catch (e) {
      report.stuck(`${c.label} — מאמצע`, String(e.message || e).slice(0, 80));
      continue;
    }
    const steps = [start, ...(await walk('ArrowRight', 4))];
    const v = verdictOf(steps, 'ArrowRight');
    console.log(`\n== ${c.label} (מאמצע)`);
    console.log(`   ArrowRight: ${trace(steps)}`);
    const shape = v.visual ? 'חזותי' : v.logical ? 'לוגי' : 'מעורב';
    if (v.logical) report.fail(`${c.label} — מאמצע הטקסט`, `${shape}: ימין הזיז שמאלה`);
    else if (v.visual) report.pass(`${c.label} — מאמצע הטקסט`, `${shape}: ימין הזיז ימינה`);
    else report.partial(`${c.label} — מאמצע הטקסט`, shape);
  }

  /* -------- בחירה ב-Shift+חץ: לאן היא נמשכת -------- */
  console.log('\n---- Shift+ArrowRight ----');
  for (const c of CASES.filter((x) => /בסיס: פסקה עברית באמצע|מפריד טאב/.test(x.label))) {
    let start;
    try {
      start = await caretMid(c.find);
    } catch (e) {
      report.stuck(`${c.label} — Shift`, String(e.message || e).slice(0, 80));
      continue;
    }
    for (let i = 0; i < 3; i++) {
      await app.press('ArrowRight', 'ArrowRight', 39, 8);
      await app.sleep(200);
    }
    const after = await caretState();
    const forward = (after.end ?? 0) > (start.end ?? 0);
    const backward = (after.off ?? 0) < (start.off ?? 0);
    console.log(
      `\n== ${c.label}: לפני [${start.off},${start.end}] אחרי [${after.off},${after.end}] → ${
        forward ? 'נמשכה קדימה (שמאלה על המסך)' : backward ? 'נמשכה אחורה (ימינה על המסך)' : 'לא נמשכה'
      }`,
    );
    const detail = `[${start.off},${start.end}] → [${after.off},${after.end}]`;
    if (backward && !forward) report.pass(`${c.label} — Shift+ימין בוחר ימינה`, detail);
    else if (forward && !backward) report.fail(`${c.label} — Shift+ימין בוחר שמאלה`, detail);
    else report.partial(`${c.label} — Shift+ימין`, detail);
  }
} finally {
  app.close();
}

report.print();
