/**
 * מיקרו-בנצ'מרק: כמה עולה `postflightDocx` על מסמך גדול, ואיפה הזמן נשרף.
 *
 * ## למה הוא קיים
 *
 * הביקורת על PR #75: „ה-postflight פותח וסורק את כל חלקי ה-DOCX בכל שמירה
 * רגילה; במסמכים גדולים זה עיבוד כבד במסלול שמירה, ואין אופטימיזציה שמדלגת
 * מראש על מסמכים שאינם זקוקים לתיקון”. המאגר קובע שטענת ביצועים בלי מדידה
 * נפסלת — לשני הכיוונים. לכן קודם מודדים, ורק אחר כך מכריעים אם יש מה לדלג
 * עליו.
 *
 * ## למה ב-Node ובלי דפדפן
 *
 * ‏`docx-postflight.ts` ושכבת ה-zip שמתחתיו (`docx-parts.ts`) אינם נוגעים
 * ב-DOM. מה שהם כן צריכים — `CompressionStream`/`DecompressionStream`,
 * `TextEncoder`/`TextDecoder` — קיים ב-Node 20 ומעלה בדיוק באותו API של
 * הדפדפן. מדידה בדפדפן הייתה מוסיפה רעש של הרתמה ושל CDP על מדידה של מאות
 * מילישניות, ולא הייתה מוסיפה דיוק.
 *
 * ה-TypeScript נארז ב-esbuild לזיכרון ונטען כמודול אחד, ולכן אין כאן העתק
 * שני של הקוד שיכול להתיישן מול המקור.
 *
 * ## מה הוא מודד, ולמה שלושה מסמכים ולא אחד
 *
 * שלושת המצבים הם שלושה מסלולי זמן שונים, וממוצע שלהם אינו אומר כלום:
 *
 *   - `first`  — מסמך עברי שעוד לא תוקן. השמירה הראשונה: פריסה, סריקה,
 *                הכנסות, דחיסה מחדש וכתיבת ארכיון.
 *   - `steady` — **אותו** מסמך אחרי שכבר תוקן. זה המצב של כל שמירה שנייה
 *                ואילך, וזה המצב שהביקורת מדברת עליו: פריסה וסריקה מלאות,
 *                ואפס כתיבה.
 *   - `latin`  — מסמך בלי אף אות ימנית ובלי אף `<w:rtl/>`. שתי היציאות
 *                המוקדמות יחד: `markNeutralParagraphEnds` על היעדר אות ימנית,
 *                ו-`mirrorComplexScript` על היעדר הצהרה.
 *
 * הפירוק לשלבים אינו נמדד דרך `postflightDocx` אלא בקריאה ישירה לאותן
 * פונקציות מיוצאות, על אותם בייטים. זה מה שמאפשר לענות על השאלה האמיתית —
 * „האם היקר הוא פרישת ה-zip או הסריקה” — שהיא זו שמכריעה אם יציאה מוקדמת
 * בכלל אפשרית.
 *
 * הרצה:
 *   node scripts/qa/postflight-bench.mjs              # 40,000 פסקאות, 5 חזרות
 *   node scripts/qa/postflight-bench.mjs 5000 3       # פסקאות, חזרות
 *
 * זהו **בנצ'מרק ולא שער**: הוא אינו נכנס ל-`verify`, ואינו מחזיר קוד יציאה
 * שנגזר ממספרים. מספר שנמדד על מכונה רוויה אינו סף שאפשר להפיל עליו ריצה.
 */
import { build } from 'esbuild';
import { deflateRawSync } from 'node:zlib';
import { pathToFileURL } from 'node:url';
import { performance } from 'node:perf_hooks';

const ROOT = new URL('../../', import.meta.url).pathname.replace(/^\/([A-Za-z]:)/, '$1');

/** אורזת את מודולי המנוע לזיכרון וטוענת אותם כמודול אחד. */
async function loadEngine() {
  const result = await build({
    stdin: {
      contents: [
        "export { postflightDocx } from './src/engine/docx-postflight';",
        "export { readZip, readEntryText, rewriteEntry, writeZip, readZipComment, CONTENT_PARTS } from './src/engine/docx-parts';",
        "export { markNeutralParagraphEnds } from './src/engine/docx-neutral-mark';",
        "export { mirrorComplexScript } from './src/engine/docx-cs-mirror';",
        "export { uniqueNumberingIds } from './src/engine/docx-numbering-ids';",
      ].join('\n'),
      resolveDir: ROOT,
      loader: 'ts',
    },
    bundle: true,
    format: 'esm',
    platform: 'neutral',
    write: false,
  });
  const code = result.outputFiles[0].text;
  return import(`data:text/javascript;base64,${Buffer.from(code, 'utf8').toString('base64')}`);
}

// ── בניית הפיקסטורה ────────────────────────────────────────────────────────

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(buf) {
  let c = 0xffffffff;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/**
 * ארכיון **דחוס**, ולא `zipStored` של `docx-fixtures.mjs`.
 *
 * זה אינו פרט: ארכיון גלוי אינו עובר `inflateRaw` בכלל
 * (`readEntryText` מחזיר `decodeXml` מיד), כלומר מדידה עליו הייתה מוחקת את
 * בדיוק אותו שלב שהביקורת שואלת עליו. כל DOCX שיוצא מ-Word או מהמנוע דחוס.
 */
function zipDeflated(entries) {
  const locals = [];
  const centrals = [];
  let offset = 0;
  let count = 0;
  for (const [name, content] of Object.entries(entries)) {
    const nameBuf = Buffer.from(name, 'utf8');
    const raw = Buffer.from(content, 'utf8');
    const data = deflateRawSync(raw, { level: 6 });
    const crc = crc32(raw);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(8, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0x21, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(data.length, 18);
    local.writeUInt32LE(raw.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    local.writeUInt16LE(0, 28);
    const central = Buffer.alloc(46);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0x0800, 8);
    central.writeUInt16LE(8, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0x21, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(data.length, 20);
    central.writeUInt32LE(raw.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    locals.push(local, nameBuf, data);
    centrals.push(central, nameBuf);
    offset += local.length + nameBuf.length + data.length;
    count += 1;
  }
  const cd = Buffer.concat(centrals);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(0, 4);
  eocd.writeUInt16LE(0, 6);
  eocd.writeUInt16LE(count, 8);
  eocd.writeUInt16LE(count, 10);
  eocd.writeUInt32LE(cd.length, 12);
  eocd.writeUInt32LE(offset, 16);
  eocd.writeUInt16LE(0, 20);
  return new Uint8Array(Buffer.concat([...locals, cd, eocd]));
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const HEB = [
  'ואמר רבי יוחנן משום רבי שמעון בן יוחאי כל המקיים את התורה מעוני',
  'תנו רבנן שלושה דברים צריך אדם לומר בתוך ביתו ערב שבת עם חשכה',
  'מאי טעמא דהא קיימא לן כדברי המיקל בערוב וכל ספק ערוב להקל',
  'אמר ליה רב אשי לרב כהנא והא תנן אין בין יום טוב לשבת אלא אוכל נפש',
];
const LAT = [
  'The quick brown fox jumps over the lazy dog near the riverbank',
  'A paragraph of ordinary Latin prose with a closing period here',
  'Nothing in this document carries a right to left letter at all',
  'Measurement without a control is an anecdote and not a finding',
];

/**
 * פסקה אחת. `declared` מייצרת ריצה שמצהירה `w:rtl` **בלי** מחסנית כתב מורכב —
 * כלומר בדיוק מה ש-`mirrorComplexScript` מתקן, וזה מה שמכריח אותו לעבוד.
 */
function para(i, words, declared) {
  const text = `${words[i % words.length]} ${i}.`;
  const rPr = declared
    ? '<w:rPr><w:rFonts w:ascii="David" w:hAnsi="David"/><w:b/><w:sz w:val="28"/><w:rtl/></w:rPr>'
    : '<w:rPr><w:rFonts w:ascii="David" w:hAnsi="David"/><w:sz w:val="24"/></w:rPr>';
  return (
    '<w:p><w:pPr><w:bidi/><w:spacing w:after="120"/><w:jc w:val="both"/></w:pPr>' +
    `<w:r>${rPr}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  );
}

function documentXml(count, words, rtlRuns) {
  const body = [];
  for (let i = 0; i < count; i += 1) body.push(para(i, words, rtlRuns && i % 10 === 0));
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w:document ${W}><w:body>${body.join('')}` +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:bidi/></w:sectPr></w:body></w:document>'
  );
}

function stylesXml() {
  const styles = [];
  for (let i = 0; i < 120; i += 1) {
    styles.push(
      `<w:style w:type="paragraph" w:styleId="S${i}"><w:name w:val="Style ${i}"/>` +
        '<w:pPr><w:bidi/></w:pPr><w:rPr><w:rFonts w:ascii="David"/><w:sz w:val="24"/></w:rPr></w:style>',
    );
  }
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<w:styles ${W}><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault>` +
    `<w:pPrDefault><w:pPr><w:bidi/></w:pPr></w:pPrDefault></w:docDefaults>${styles.join('')}</w:styles>`
  );
}

/**
 * ‏`nsid` **ייחודי** לכל הגדרה, וזה תיקון של הבנצ'מרק ולא קישוט.
 *
 * הגרסה הראשונה נתנה לכל 24 ההגדרות את אותו `587013BA`, ולכן
 * `uniqueNumberingIds` שכתב 23 מהן בכל ריצה — כולל במקרה שהוצג כ„יציאה
 * מוקדמת”. כלומר השורה הלטינית מדדה תרחיש **כתיבה** והוצגה כ-no-op, והמספר
 * שדווח בעקבותיה היה גדול מהאמת. נתפס בסקירה של המתחזק.
 *
 * מי שרוצה למדוד דווקא את מסלול הכתיבה של המספור — להחזיר כאן ערך קבוע,
 * ולקרוא לשורה בשמה.
 */
function numberingXml() {
  const defs = [];
  for (let i = 0; i < 24; i += 1) {
    const nsid = (0x587013ba + i).toString(16).toUpperCase().padStart(8, '0');
    defs.push(
      `<w:abstractNum w:abstractNumId="${i}"><w:nsid w:val="${nsid}"/>` +
        '<w:multiLevelType w:val="hybridMultilevel"/>' +
        '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum>',
    );
    defs.push(`<w:num w:numId="${i + 1}"><w:abstractNumId w:val="${i}"/></w:num>`);
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering ${W}>${defs.join('')}</w:numbering>`;
}

function smallPart(root, count, words) {
  const body = [];
  for (let i = 0; i < count; i += 1) body.push(para(i, words, false));
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:${root} ${W}>${body.join('')}</w:${root}>`;
}

/**
 * `rtlRuns` מופרד מהשפה בכוונה: מסמך לטיני שיש בו `<w:rtl/>` אינו מסמך לטיני,
 * והוא היה מסתיר את היציאה המוקדמת של `mirrorComplexScript` שהוא בא למדוד.
 */
function buildFixture(paragraphs, words, rtlRuns = true) {
  return zipDeflated({
    '[Content_Types].xml':
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>',
    '_rels/.rels':
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/document.xml': documentXml(paragraphs, words, rtlRuns),
    'word/styles.xml': stylesXml(),
    'word/numbering.xml': numberingXml(),
    'word/footnotes.xml': smallPart('footnotes', 200, words),
    'word/endnotes.xml': smallPart('endnotes', 40, words),
    'word/header1.xml': smallPart('hdr', 3, words),
    'word/footer1.xml': smallPart('ftr', 3, words),
    'word/settings.xml': `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:settings ${W}><w:defaultTabStop w:val="720"/></w:settings>`,
    'docProps/app.xml':
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>otzaria</Application></Properties>',
  });
}

// ── מדידה ──────────────────────────────────────────────────────────────────

/** חציון — לא ממוצע. ריצה אחת שנתקעה על עומק אחר אינה אמורה להזיז את המספר. */
const median = (xs) => [...xs].sort((a, b) => a - b)[Math.floor(xs.length / 2)];
const ms = (x) => `${x.toFixed(1)}ms`;

async function repeat(times, fn) {
  const out = [];
  for (let i = 0; i < times; i += 1) {
    const start = performance.now();
    await fn();
    out.push(performance.now() - start);
  }
  return out;
}

/** פירוק לשלבים: אותה עבודה בדיוק, בקריאות ישירות, כדי לדעת מי אוכל מה. */
async function phases(engine, bytes) {
  const { readZip, readEntryText, rewriteEntry, writeZip, readZipComment, CONTENT_PARTS } = engine;
  const { markNeutralParagraphEnds, mirrorComplexScript, uniqueNumberingIds } = engine;
  const out = {};

  let start = performance.now();
  const entries = readZip(bytes);
  out.readZip = performance.now() - start;

  const content = entries.filter((entry) => CONTENT_PARTS.test(entry.name));
  const texts = new Map();
  start = performance.now();
  for (const entry of content) texts.set(entry, await readEntryText(entry));
  out.inflate = performance.now() - start;

  out.neutral = 0;
  out.mirror = 0;
  out.numbering = 0;
  const changed = new Map();
  for (const entry of content) {
    const xml = texts.get(entry);
    if (xml === null) continue;
    start = performance.now();
    let next = markNeutralParagraphEnds(xml) ?? xml;
    out.neutral += performance.now() - start;
    start = performance.now();
    next = mirrorComplexScript(next) ?? next;
    out.mirror += performance.now() - start;
    if (/^word\/numbering\d*\.xml$/i.test(entry.name)) {
      start = performance.now();
      next = uniqueNumberingIds(next) ?? next;
      out.numbering += performance.now() - start;
    }
    if (next !== xml) changed.set(entry, next);
  }

  start = performance.now();
  const patched = new Map();
  for (const [entry, next] of changed) {
    patched.set(entry, await rewriteEntry(entry, new TextEncoder().encode(next)));
  }
  out.deflate = performance.now() - start;

  start = performance.now();
  if (patched.size > 0) writeZip(entries.map((entry) => patched.get(entry) ?? entry), readZipComment(bytes));
  out.writeZip = performance.now() - start;

  out.changedParts = changed.size;
  return out;
}

async function measure(engine, label, bytes, times) {
  const total = await repeat(times, () => engine.postflightDocx(bytes));
  const phase = await phases(engine, bytes);
  const kb = (bytes.byteLength / 1024).toFixed(0);
  console.log(`\n── ${label}  (${kb}KB דחוס)`);
  console.log(`   postflightDocx  חציון ${ms(median(total))}   טווח ${ms(Math.min(...total))}–${ms(Math.max(...total))}`);
  console.log(
    `   פירוק:  readZip ${ms(phase.readZip)} | פריסה+פענוח ${ms(phase.inflate)} | ` +
      `neutral ${ms(phase.neutral)} | mirror ${ms(phase.mirror)} | numbering ${ms(phase.numbering)} | ` +
      `דחיסה ${ms(phase.deflate)} | writeZip ${ms(phase.writeZip)}   (${phase.changedParts} חלקים תוקנו)`,
  );
  return { total: median(total), phase };
}

async function main() {
  const paragraphs = Number(process.argv[2] ?? 40000);
  const times = Number(process.argv[3] ?? 5);
  if (!Number.isInteger(paragraphs) || paragraphs <= 0 || !Number.isInteger(times) || times <= 0) {
    throw new Error('paragraphs and repetitions must be positive integers');
  }
  const engine = await loadEngine();

  console.log(`postflight-bench — ${paragraphs.toLocaleString('he-IL')} פסקאות, ${times} חזרות, Node ${process.version}`);

  const hebrew = buildFixture(paragraphs, HEB);
  const first = await measure(engine, 'first — מסמך עברי שעוד לא תוקן', hebrew, times);

  const fixed = await engine.postflightDocx(hebrew);
  if (!fixed) throw new Error('הפיקסטורה העברית לא תוקנה — אין מצב steady למדוד');
  const steady = await measure(engine, 'steady — אותו מסמך אחרי שכבר תוקן', fixed, times);
  const again = await engine.postflightDocx(fixed);
  if (again !== null) throw new Error('steady fixture is not idempotent');
  console.log(`   אידמפוטנטי: מעבר שני מחזיר ${again === null ? 'null — אין מה לתקן' : '**בייטים חדשים**'}`);

  // `false` — בלי `w:rtl`. זה מה שמבודד את היציאה המוקדמת של המראה; עם
  // ברירת המחדל הפיקסטורה נושאת הצהרות, והמראה רצה במלואה.
  const latin = buildFixture(paragraphs, LAT, false);
  const plain = await measure(engine, 'latin — בלי אף אות ימנית', latin, times);
  if (plain.phase.changedParts !== 0 || (await engine.postflightDocx(latin)) !== null) {
    throw new Error('Latin control must not rewrite any parts');
  }

  console.log('\n── ההכרעה במספרים');
  const share = (part, whole) => `${((part / whole) * 100).toFixed(0)}%`;
  const unavoidable = steady.phase.readZip + steady.phase.inflate;
  console.log(`   שמירה חוזרת (steady): ${ms(steady.total)}, מזה ${ms(unavoidable)} (${share(unavoidable, steady.total)}) פרישת zip ופענוח`);
  console.log(`   הסריקות עצמן: ${ms(steady.phase.neutral + steady.phase.mirror + steady.phase.numbering)} (${share(steady.phase.neutral + steady.phase.mirror + steady.phase.numbering, steady.total)})`);
  console.log(`   שמירה ראשונה: ${ms(first.total)}, מזה דחיסה+כתיבה ${ms(first.phase.deflate + first.phase.writeZip)}`);
  console.log(`   מסמך לטיני: ${ms(plain.total)} — היציאה המוקדמת חוסכת ${ms(steady.total - plain.total)}`);
}

if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  main().catch((error) => {
    console.error(error);
    process.exitCode = 1;
  });
}
