/**
 * `w:nsid` ייחודי לכל הגדרת מספור — והתיקונים של הדרך החוצה כמקשה אחת.
 *
 * המקרה שממנו זה בא נמדד ב-Word: רשימה עברית שהעורך יצר חלקה nsid עם
 * ההגדרה העשרונית שממנה שוכפלה, ו-Word הציג אותה כ-„1. 2.”. ראו
 * engine/docx-numbering-ids.ts.
 */
import { describe, it, expect, vi } from 'vitest';
import { deflateRawSync, gzipSync, inflateRawSync } from 'node:zlib';
import { uniqueNumberingIds } from '../../src/engine/docx-numbering-ids';
import { postflightDocx } from '../../src/engine/docx-postflight';
import {
  CONTENT_PARTS,
  crc32,
  readEntryText,
  readZip,
  rewriteDocxXmlParts,
  writeZip,
  type Bytes,
  type ZipEntry,
} from '../../src/engine/docx-parts';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const abstractNum = (id: number, nsid: string, fmt: string, text: string): string =>
  `<w:abstractNum w:abstractNumId="${id}"><w:nsid w:val="${nsid}"/><w:multiLevelType w:val="hybridMultilevel"/>` +
  `<w:tmpl w:val="EE6417C4"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/>` +
  `<w:lvlText w:val="${text}"/></w:lvl></w:abstractNum>`;

const numbering = (body: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering ${W}>${body}` +
  `<w:num w:numId="3"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="4"><w:abstractNumId w:val="2"/></w:num>` +
  `</w:numbering>`;

/** צורת הקובץ שנמדדה: העשרוני והעברי חולקים nsid. */
const MEASURED = numbering(
  abstractNum(0, '16126B07', 'bullet', '') +
    abstractNum(1, '587013BA', 'decimal', '%1.') +
    abstractNum(2, '587013BA', 'hebrew1', '%1)'),
);

const nsids = (xml: string): string[] => [...xml.matchAll(/<w:nsid w:val="([^"]+)"\/>/g)].map((m) => m[1]!);

describe('uniqueNumberingIds', () => {
  it('ההופעה השנייה של nsid מקבלת ערך חדש, והראשונה נשארת', () => {
    const out = uniqueNumberingIds(MEASURED) ?? '';
    const [bullet, decimal, hebrew] = nsids(out);
    expect(bullet).toBe('16126B07');
    expect(decimal).toBe('587013BA');
    expect(hebrew).not.toBe('587013BA');
    expect(hebrew).toMatch(/^[0-9A-F]{8}$/);
  });

  it('רק התג עצמו משתנה — שאר הקובץ זהה', () => {
    const out = uniqueNumberingIds(MEASURED) ?? '';
    const fresh = nsids(out)[2]!;
    expect(out.replace(fresh, '587013BA')).toBe(MEASURED);
  });

  it('הערך החדש דטרמיניסטי — שתי שמירות, אותם בייטים', () => {
    expect(uniqueNumberingIds(MEASURED)).toBe(uniqueNumberingIds(MEASURED));
  });

  it('הערך החדש אינו מתנגש באף nsid קיים', () => {
    const many = numbering(
      [0, 1, 2, 3, 4, 5].map((id) => abstractNum(id, id < 3 ? 'AAAAAAAA' : `0000000${id}`, 'decimal', '%1.')).join(''),
    );
    const values = nsids(uniqueNumberingIds(many) ?? '');
    expect(new Set(values).size).toBe(values.length);
  });

  it('התנגשות של ההאש — הערך החדש מדלג עליה', () => {
    // הלולאה שמוסיפה מלח נכנסת רק כשהפלט של ההאש כבר קיים בקובץ, ולכן כאן
    // הוא נשתל שם: מריצים פעם אחת כדי לקבל אותו, ואז בונים קובץ שיש בו הגדרה
    // רביעית שה-nsid שלה הוא בדיוק הוא. בלי הלולאה, ההגדרה השלישית מקבלת את
    // הערך של הרביעית — כלומר nsid כפול, הבאג שהמודול בא לתקן.
    const collision = nsids(uniqueNumberingIds(MEASURED) ?? '')[2]!;
    const xml = numbering(
      abstractNum(0, '16126B07', 'bullet', '') +
        abstractNum(1, '587013BA', 'decimal', '%1.') +
        abstractNum(2, '587013BA', 'hebrew1', '%1)') +
        abstractNum(3, collision, 'decimal', '%1.'),
    );
    const values = nsids(uniqueNumberingIds(xml) ?? '');
    expect(values).toHaveLength(4);
    expect(values[2]).not.toBe(collision);
    expect(new Set(values).size).toBe(4);
  });

  // שלוש הבדיקות שמצפות ל-`null` (כאן, „nsid בתוך הערה”, ו„מסמך שאין בו מה
  // לתקן” שלמטה) עוברות גם על מימוש שמחזיר `null` תמיד. מה שהורג אותו הוא
  // הבדיקות החיוביות שלצדן באותו `describe`, ולכן הן נשארות זוגות. „nsid
  // בתוך הערה” נושאת משקל משלה: מחיקת `SKIPPED_SPANS` מהסורק צובעת אותה אדום.
  it('בלי כפילות — `null`', () => {
    expect(uniqueNumberingIds(numbering(abstractNum(0, '11111111', 'decimal', '%1.')))).toBeNull();
  });

  it('השוואה בלי תלות ברישיות', () => {
    const xml = numbering(abstractNum(0, 'abcdef12', 'decimal', '%1.') + abstractNum(1, 'ABCDEF12', 'hebrew1', '%1.'));
    const values = nsids(uniqueNumberingIds(xml) ?? '');
    expect(values[0]).toBe('abcdef12');
    expect(values[1]!.toUpperCase()).not.toBe('ABCDEF12');
  });

  it('קידומת אחרת נתמכת', () => {
    const xml = MEASURED.replace(/<(\/?)w:/g, '<$1ns0:').replace(/ w:/g, ' ns0:').replace('xmlns:ns0', 'xmlns:ns0');
    const out = uniqueNumberingIds(xml) ?? '';
    const values = [...out.matchAll(/<ns0:nsid ns0:val="([^"]+)"\/>/g)].map((m) => m[1]);
    expect(new Set(values).size).toBe(3);
  });

  it('nsid בתוך הערה אינו נספר', () => {
    const xml = numbering(
      `<!-- ${abstractNum(9, '587013BA', 'decimal', '%1.')} -->` + abstractNum(1, '587013BA', 'decimal', '%1.'),
    );
    expect(uniqueNumberingIds(xml)).toBeNull();
  });

  /**
   * המקרה שבאמת מגיע ללולאת המלח, ולא התנגשות ההאש שמעליו: שלוש הגדרות
   * שחולקות גם את ה-nsid וגם את ה-`abstractNumId` מקבלות את אותו קלט להאש
   * בדיוק, ולכן השנייה והשלישית נגזרות לאותו ערך. בלי הלולאה שתיהן יוצאות
   * `2DA62BE6` — nsid כפול שנכתב כאן בידיים, כלומר הבאג שהמודול בא לתקן.
   */
  it('שלוש הגדרות עם אותו nsid ואותו abstractNumId — שלושה ערכים שונים', () => {
    const xml = numbering(
      abstractNum(7, 'AAAAAAAA', 'decimal', '%1.') +
        abstractNum(7, 'AAAAAAAA', 'hebrew1', '%1)') +
        abstractNum(7, 'AAAAAAAA', 'bullet', ''),
    );
    const values = nsids(uniqueNumberingIds(xml) ?? '');
    expect(values).toHaveLength(3);
    expect(values[0]).toBe('AAAAAAAA');
    expect(new Set(values).size).toBe(3);
  });
});

const enc = new TextEncoder();

function entry(name: string, text: string): ZipEntry {
  const data = enc.encode(text) as Bytes;
  return {
    name,
    nameBytes: enc.encode(name) as Bytes,
    versionMadeBy: 20,
    versionNeeded: 10,
    flags: 0,
    method: 0,
    modTime: 0,
    modDate: 0x21,
    crc: crc32(data),
    internalAttrs: 0,
    externalAttrs: 0,
    data,
    uncompressedSize: data.byteLength,
  };
}

async function partText(bytes: Bytes, name: string): Promise<string> {
  const found = readZip(bytes)?.find((e) => e.name === name);
  return (found && (await readEntryText(found))) ?? '';
}

/**
 * CRC32 שאינו שלנו: ה-trailer של gzip הוא ה-CRC32 של הנתונים הלא-דחוסים
 * (RFC 1952 §2.3.1), ולכן `gzipSync` הוא מימוש ייחוס שקיים בכל גרסת Node.
 * (`zlib.crc32` היה ישיר יותר, אבל הוא מ-Node 20.15 והמאגר מצהיר `node >= 20`.)
 */
function referenceCrc32(bytes: Uint8Array): number {
  const gzipped = gzipSync(bytes);
  return new DataView(gzipped.buffer, gzipped.byteOffset, gzipped.byteLength).getUint32(
    gzipped.byteLength - 8,
    true,
  );
}

/**
 * קורא zip **עצמאי**: הולך על הכותרות המקומיות מתחילת הארכיון, פורס
 * ב-`inflateRawSync` של Node ומאמת את ה-CRC מול המימוש שלמעלה. אין בו שורה
 * אחת מ-docx-parts.ts, וזה כל העניין — כל שאר הבדיקות כאן סוגרות מעגל דרך
 * `readZip` שלנו, וקורא וכותב שטועים **באותה** טעות עוברים אותו בשלום.
 */
function readArchive(bytes: Bytes): { name: string; text: string; crcOk: boolean }[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const parts: { name: string; text: string; crcOk: boolean }[] = [];
  let at = 0;
  while (at + 30 <= bytes.byteLength && view.getUint32(at, true) === 0x04034b50) {
    const method = view.getUint16(at + 8, true);
    const crc = view.getUint32(at + 14, true);
    const compressed = view.getUint32(at + 18, true);
    const uncompressed = view.getUint32(at + 22, true);
    const nameLength = view.getUint16(at + 26, true);
    const extraLength = view.getUint16(at + 28, true);
    const from = at + 30 + nameLength + extraLength;
    const stored = bytes.subarray(from, from + compressed);
    const raw = method === 8 ? new Uint8Array(inflateRawSync(stored)) : stored;
    parts.push({
      name: new TextDecoder().decode(bytes.subarray(at + 30, at + 30 + nameLength)),
      text: new TextDecoder().decode(raw),
      crcOk: raw.byteLength === uncompressed && referenceCrc32(raw) === crc,
    });
    at = from + compressed;
  }
  return parts;
}

/** 2024-01-15, ‏10:30:40 — חותמת DOS אמיתית, ולא אפס. */
const MOD_TIME = (10 << 11) | (30 << 5) | 20;
const MOD_DATE = ((2024 - 1980) << 9) | (1 << 5) | 15;
/** „נכתב ב-Unix, גרסה 2.0” — מה ש-Info-ZIP כותב, ושונה מ-`versionNeeded`. */
const VERSION_MADE_BY = 0x0314;
/** „קובץ טקסט”, וביט הארכיון של DOS. */
const INTERNAL_ATTRS = 1;
const EXTERNAL_ATTRS = 0x20;

/**
 * כותב zip **עצמאי**, תאומו של הקורא שמעליו: אין בו שורה אחת מ-docx-parts.ts,
 * וה-CRC שלו הוא מימוש הייחוס.
 *
 * בלעדיו כל הבדיקות כאן בונות את הקלט שלהן ב-`writeZip` — הכותב הנבדק —
 * וטעות שסימטרית בין הכתיבה לקריאה מתבטלת: נמדד שמוטציה שמסדרת את הרשומות
 * הפוך ב-`writeZip` משאירה את הקובץ כולו ירוק, ובכלל זה את הבדיקה ששמה
 * „`[Content_Types].xml` נשאר הרשומה הראשונה”, מפני שהקלט שלה נכתב באותו
 * כותב הפוך ושתי ההיפוכים מבטלים זה את זה.
 *
 * תאום כזה קיים גם ב-`tests/unit/docx-preflight.test.ts` (`buildZip`). הוא
 * מועתק לכאן, ולא משותף, מפני שקובץ בדיקה שמייבא מקובץ בדיקה אחר גורר איתו
 * את ה-`describe` שלו.
 *
 * חמשת הערכים שלמעלה אינם אפס **בכוונה**: אלה השדות שאף בדיקה אינה מסתכלת
 * עליהם, ולכן ארכיון שכולו אפסים היה נותן לכתיבה שמשמיטה אותם לעבור. נמדד —
 * כשהחותמת הייתה אפס, מוטציה שמאפסת את `modDate` ב-`writeZip` השאירה את כל
 * 112 הבדיקות בשני הקבצים ירוקות.
 */
function buildArchive(parts: { name: string; content: string; deflate?: boolean }[]): Bytes {
  const records = parts.map((part) => {
    const raw = enc.encode(part.content);
    const stored = part.deflate ? new Uint8Array(deflateRawSync(raw)) : raw;
    return {
      nameBytes: enc.encode(part.name),
      raw,
      stored,
      method: part.deflate ? 8 : 0,
      crc: referenceCrc32(raw),
    };
  });

  const size =
    records.reduce((total, r) => total + 76 + 2 * r.nameBytes.byteLength + r.stored.byteLength, 0) + 22;
  const out = new Uint8Array(size) as Bytes;
  const view = new DataView(out.buffer);
  const offsets: number[] = [];
  let at = 0;

  for (const record of records) {
    offsets.push(at);
    view.setUint32(at, 0x04034b50, true);
    view.setUint16(at + 4, 20, true);
    view.setUint16(at + 8, record.method, true);
    view.setUint16(at + 10, MOD_TIME, true);
    view.setUint16(at + 12, MOD_DATE, true);
    view.setUint32(at + 14, record.crc, true);
    view.setUint32(at + 18, record.stored.byteLength, true);
    view.setUint32(at + 22, record.raw.byteLength, true);
    view.setUint16(at + 26, record.nameBytes.byteLength, true);
    at += 30;
    out.set(record.nameBytes, at);
    at += record.nameBytes.byteLength;
    out.set(record.stored, at);
    at += record.stored.byteLength;
  }

  const centralOffset = at;
  records.forEach((record, index) => {
    view.setUint32(at, 0x02014b50, true);
    view.setUint16(at + 4, VERSION_MADE_BY, true);
    view.setUint16(at + 6, 20, true);
    view.setUint16(at + 10, record.method, true);
    view.setUint16(at + 12, MOD_TIME, true);
    view.setUint16(at + 14, MOD_DATE, true);
    view.setUint32(at + 16, record.crc, true);
    view.setUint32(at + 20, record.stored.byteLength, true);
    view.setUint32(at + 24, record.raw.byteLength, true);
    view.setUint16(at + 28, record.nameBytes.byteLength, true);
    view.setUint16(at + 36, INTERNAL_ATTRS, true);
    view.setUint32(at + 38, EXTERNAL_ATTRS, true);
    view.setUint32(at + 42, offsets[index]!, true);
    at += 46;
    out.set(record.nameBytes, at);
    at += record.nameBytes.byteLength;
  });

  view.setUint32(at, 0x06054b50, true);
  view.setUint16(at + 8, records.length, true);
  view.setUint16(at + 10, records.length, true);
  view.setUint32(at + 12, at - centralOffset, true);
  view.setUint32(at + 16, centralOffset, true);
  return out;
}

/** מיקומי ה-EOCD והרשומות של הספרייה המרכזית, לארבע ההתעללויות שלמטה. */
function directory(bytes: Bytes): { eocd: number; records: { at: number; size: number }[] } {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.byteLength - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;

  const records: { at: number; size: number }[] = [];
  let at = view.getUint32(eocd + 16, true);
  for (let i = view.getUint16(eocd + 10, true); i > 0; i -= 1) {
    const size =
      46 + view.getUint16(at + 28, true) + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
    records.push({ at, size });
    at += size;
  }
  return { eocd, records };
}

/**
 * אותם בייטים בדיוק, עם ספרייה מרכזית שסדרה הפוך מהסדר הפיזי. זה ארכיון
 * חוקי — APPNOTE אינו מחייב שהספרייה תמנה את הרשומות בסדר שבו הן כתובות —
 * וזה המקרה שבו קורא שהולך אחרי הספרייה בלבד מזיז את `[Content_Types].xml`
 * ממקומו הראשון, שאותו OPC כן דורש.
 */
function reverseCentralDirectory(bytes: Bytes): Bytes {
  const { records } = directory(bytes);
  const out = bytes.slice() as Bytes;
  let write = records[0]!.at;
  for (const record of [...records].reverse()) {
    out.set(bytes.subarray(record.at, record.at + record.size), write);
    write += record.size;
  }
  return out;
}

/** הסימן של ZIP64 בגודל הלא-דחוס של הרשומה הראשונה. */
function poisonUncompressedSize(bytes: Bytes): Bytes {
  const out = bytes.slice() as Bytes;
  new DataView(out.buffer).setUint32(directory(bytes).records[0]!.at + 24, 0xffffffff, true);
  return out;
}

/**
 * הרשומה האחרונה בספרייה נחתכת, וה-EOCD נדחף אחורה — כלומר ארכיון שנקטע
 * אבל עדיין נראה שלם מסופו. `subarray` מעבר לקצה אינו זורק אלא מקצר בשקט,
 * ולכן בלי בדיקת גבולות השם של הרשומה הזאת בולע את ה-EOCD.
 */
/**
 * ה-EOCD מצביע על ספרייה מרכזית שמתחילה בבייט האחרון של הקובץ.
 *
 * זה המקרה היחיד שבו הגבול שבראש לולאת הקריאה עושה משהו, ומה שהוא שומר עליו
 * הוא ההבטחה „`null` ולא זריקה”: בלעדיו `getUint32` על היסט שחורג **זורק**
 * `RangeError` (נמדד), והוא עולה מ-`readZip` דרך `postflightDocx` ומפיל את
 * השמירה כולה במקום להחזיר את המסמך כמות שהוא.
 */
function centralOffsetPastEnd(bytes: Bytes): Bytes {
  const out = bytes.slice() as Bytes;
  new DataView(out.buffer).setUint32(directory(bytes).eocd + 16, out.byteLength - 1, true);
  return out;
}

function truncateLastRecord(bytes: Bytes, cut: number): Bytes {
  const { eocd, records } = directory(bytes);
  const last = records[records.length - 1]!;
  const out = new Uint8Array(bytes.byteLength - cut) as Bytes;
  out.set(bytes.subarray(0, last.at + last.size - cut), 0);
  out.set(bytes.subarray(eocd), last.at + last.size - cut);
  return out;
}

describe('postflightDocx', () => {
  const paragraph =
    `<w:p><w:pPr><w:bidi/></w:pPr><w:r><w:t xml:space="preserve">שלום עולם.</w:t></w:r></w:p>`;
  const documentXml = `<w:document ${W}><w:body>${paragraph}</w:body></w:document>`;
  /** גדול מספיק כדי שהדחיסה תהיה קטנה מהמקור, כלומר שהרשומה תיכתב deflate. */
  const bigDocumentXml = `<w:document ${W}><w:body>${paragraph.repeat(200)}</w:body></w:document>`;
  const pack = (body: string): Bytes =>
    writeZip([
      entry('[Content_Types].xml', '<Types/>'),
      entry('word/document.xml', body),
      entry('word/numbering.xml', MEASURED),
    ]);

  it('שני התיקונים, כל אחד בחלק שלו', async () => {
    const input = writeZip([entry('word/document.xml', documentXml), entry('word/numbering.xml', MEASURED)]);
    const out = await postflightDocx(input);
    expect(out).not.toBeNull();
    // ‏U+200F — ה-RLM שהתיקון השני מוסיף אחרי הנקודה. ראו docx-neutral-mark.ts.
    expect(await partText(out!, 'word/document.xml')).toContain('שלום עולם.‏');
    expect(new Set(nsids(await partText(out!, 'word/numbering.xml'))).size).toBe(3);
  });

  // בזוג עם הבדיקה שמעליה: מימוש שמחזיר `null` תמיד עובר כאן ונופל שם.
  it('מסמך שאין בו מה לתקן — `null`', async () => {
    const plain = `<w:document ${W}><w:body><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:body></w:document>`;
    const input = writeZip([entry('word/document.xml', plain)]);
    expect(await postflightDocx(input)).toBeNull();
  });

  it('הארכיון שנכתב נקרא בקורא עצמאי — שמות, סדר, גדלים ו-CRC', async () => {
    const out = await postflightDocx(pack(bigDocumentXml));
    expect(out).not.toBeNull();

    const parts = readArchive(out!);
    expect(parts.map((part) => part.name)).toEqual([
      '[Content_Types].xml',
      'word/document.xml',
      'word/numbering.xml',
    ]);
    expect(parts.map((part) => part.crcOk)).toEqual([true, true, true]);
    // על התוכן נבדק רק מה ששכבת החבילה אחראית לו: החלק שנכתב מחדש נפרס
    // בחזרה שלם. איזה תיקון הוא עבר זו שאלה של השכבה שמעליה.
    expect(new Set(nsids(parts[2]!.text)).size).toBe(3);
    expect(parts[1]!.text.startsWith('<w:document ')).toBe(true);
  });

  it('`[Content_Types].xml` נשאר הרשומה הראשונה, גם כשהספרייה המרכזית מונה אחרת', async () => {
    const out = await postflightDocx(reverseCentralDirectory(pack(documentXml)));
    expect(out).not.toBeNull();
    expect(readArchive(out!)[0]!.name).toBe('[Content_Types].xml');
  });

  /**
   * שלושה ארכיונים שאיננו מבינים, ולכן המסמך יוצא כפי שהמנוע כתב אותו.
   * שלושתם מקרים שהתקבלו עד שנמדדו: הסימן של ZIP64 ב-`uncompressedSize` —
   * אחד משלושת שדות הרשומה שהכתיבה מחשבת מחדש — רשומה מרכזית אחרונה שנחתכה,
   * שאת השם שלה `subarray` קיצר בשקט במקום לזרוק, והיסט ספרייה שחורג מהקובץ,
   * שעליו `getUint32` זורק. השלישי הוא **גם** מה שמוכיח את הגבול שבראש
   * הלולאה: בלעדיו הבדיקה הזאת אינה מקבלת `null` אלא `RangeError`.
   */
  it.each([
    ['ZIP64 ב-uncompressedSize', poisonUncompressedSize],
    ['רשומה מרכזית אחרונה קטועה', (bytes: Bytes) => truncateLastRecord(bytes, 10)],
    ['היסט ספרייה מרכזית שחורג מהקובץ', centralOffsetPastEnd],
  ])('ארכיון שאינו נקרא — `null` (%s)', async (_label, damage) => {
    expect(await postflightDocx(damage(pack(documentXml)))).toBeNull();
  });

  /**
   * הקלט הזה נבנה בכותב **עצמאי**, ולא ב-`writeZip`, והפלט נקרא בקורא
   * העצמאי — כלומר אין כאן אף שורה משכבת החבילה משני הצדדים. זה מה שכל שאר
   * הבדיקות כאן אינן יכולות לעשות: נמדד שמוטציה שמסדרת את הרשומות הפוך
   * ב-`writeZip` משאירה את הקובץ כולו ירוק, מפני שהקלט והפלט עוברים דרך
   * אותו כותב הפוך פעמיים.
   */
  it('קלט מכותב עצמאי, פלט לקורא עצמאי — שמות, סדר, CRC ותוכן', async () => {
    const input = buildArchive([
      { name: '[Content_Types].xml', content: '<Types/>' },
      { name: 'word/document.xml', content: bigDocumentXml, deflate: true },
      { name: 'word/numbering.xml', content: MEASURED },
    ]);

    const out = await postflightDocx(input);
    expect(out).not.toBeNull();

    const parts = readArchive(out!);
    expect(parts.map((part) => part.name)).toEqual([
      '[Content_Types].xml',
      'word/document.xml',
      'word/numbering.xml',
    ]);
    expect(parts.map((part) => part.crcOk)).toEqual([true, true, true]);
    expect(parts[1]!.text).toContain('שלום עולם.‏');
    expect(new Set(nsids(parts[2]!.text)).size).toBe(3);
  });

  /**
   * ארכיון שעבר קריאה וכתיבה ולא נגעו בו — בייט-בבייט אותו קובץ.
   *
   * זו ההבטחה של `writeZip` („כל מה שאינו התוכן שתוקן נכתב בחזרה כמות שהוא”)
   * בצורתה החזקה ביותר, ועד עכשיו היא לא נבדקה בשום מקום: הבדיקות האחרות
   * משוות חלקים, לא בייטים. היא נמדדה על שלושה-עשר קובצי ה-docx שבעץ
   * ועוברת בכולם, ואלה אינם בעץ הגיט — ולכן הקובץ כאן נבנה בכותב העצמאי,
   * שהוא ממילא העד הנכון: הוא אינו יודע דבר על סדר השדות של `writeZip`.
   */
  it('קריאה וכתיבה בלי תיקון — אותם בייטים בדיוק', () => {
    const input = buildArchive([
      { name: '[Content_Types].xml', content: '<Types/>' },
      { name: 'word/document.xml', content: bigDocumentXml, deflate: true },
      { name: 'word/numbering.xml', content: MEASURED },
      { name: 'word/settings.xml', content: '<w:settings/>' },
    ]);

    const entries = readZip(input);
    expect(entries).not.toBeNull();
    expect(writeZip(entries!)).toEqual(input);
  });

  /**
   * ההבטחה של השכבה היא `null` ולא זריקה, ולכן תיקון שזורק על חלק אחד אינו
   * רשאי להפיל את השמירה כולה: החלק הזה נשאר כמות שהוא, והשאר מתוקן.
   */
  it('תיקון שזורק מפיל חלק אחד בלבד, ומיומן', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const out = await rewriteDocxXmlParts(
        pack(documentXml),
        (name) => CONTENT_PARTS.test(name),
        (xml, name) => {
          if (name === 'word/document.xml') throw new Error('נפילה בחלק אחד');
          return uniqueNumberingIds(xml);
        },
      );

      expect(out).not.toBeNull();
      expect(await partText(out!, 'word/document.xml')).toBe(documentXml);
      expect(new Set(nsids(await partText(out!, 'word/numbering.xml'))).size).toBe(3);
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });

  /**
   * וזה הכשל שקורה בפועל, ולא זריקה: `TextEncoder.encode` אינו זורק על מה
   * שאינו מחרוזת אלא מקודד את הצורה שלו. נמדד — בלי השומר, `async` שנשתל
   * בטעות בחתימה כותב לחלק את המחרוזת `"[object Promise]"`, ו-`return`
   * שנשמט כותב חלק **ריק**; בשני המקרים הארכיון שיוצא תקין לגמרי, על כל
   * החתימות וה-CRC שבו, והמסמך של המשתמש הרוס בשקט.
   */
  it.each([
    ['callback שהפך ל-async', async (xml: string) => `${xml} `],
    ['callback בלי return', () => undefined],
  ])('תיקון שאינו מחזיר טקסט משאיר את החלק, ומיומן (%s)', async (_label, broken) => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const out = await rewriteDocxXmlParts(
        pack(documentXml),
        (name) => CONTENT_PARTS.test(name),
        (xml, name) =>
          name === 'word/document.xml'
            ? (broken(xml) as unknown as string | null)
            : uniqueNumberingIds(xml),
      );

      expect(out).not.toBeNull();
      expect(await partText(out!, 'word/document.xml')).toBe(documentXml);
      expect(new Set(nsids(await partText(out!, 'word/numbering.xml'))).size).toBe(3);
      expect(warn).toHaveBeenCalledTimes(1);
    } finally {
      warn.mockRestore();
    }
  });
});
