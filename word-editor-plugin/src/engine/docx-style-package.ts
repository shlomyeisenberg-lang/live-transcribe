/**
 * עריכת סגנון על חבילת DOCX שלמה — הצד של ה-zip של `docx-style-edit.ts`.
 *
 * שני חלקים נוגעים: `word/styles.xml` תמיד, ובמחיקה גם כל חלק תוכן שיש בו
 * פסקאות (גוף, כותרות עליונות ותחתונות, הערות שוליים וסיום, הערות) — פסקה
 * שהצביעה על סגנון שנמחק חוזרת ל„רגיל”, כמו ב-Word, ואינה נשארת עם הפניה
 * יתומה.
 *
 * `stylesWithEffects.xml` אינו נכתב: זה עותק של Word 2010 בשביל Word 2007,
 * ו-Word עצמו קורא את `styles.xml` כשהוא קיים. המנוע אינו קורא אותו כלל.
 */
import { readEntryText, readZip, rewriteDocxXmlParts, type Bytes } from './docx-parts';
import { editStyles, stripStyleReferences, type StyleEdit } from './docx-style-edit';

const STYLES_PART = 'word/styles.xml';

/** החלקים שיש בהם פסקאות, וממילא הפניות `pStyle`/`rStyle`. */
const PARAGRAPH_PARTS = /^word\/(?:document|footnotes|endnotes|comments|header|footer)\d*\.xml$/i;

const DOCX_MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

async function bytesOf(blob: Blob): Promise<Bytes> {
  return new Uint8Array(await blob.arrayBuffer());
}

/** ‏`styles.xml` של המסמך, או `null` כשאין או שאינו נקרא. */
export async function readDocxStyles(docx: Blob): Promise<string | null> {
  const entries = readZip(await bytesOf(docx));
  const entry = entries?.find((candidate) => candidate.name === STYLES_PART);
  return entry ? readEntryText(entry) : null;
}

export type DocxStyleEditResult =
  | { ok: true; docx: Blob; id: string }
  | { ok: false; message: string };

/**
 * מחילה עריכה על החבילה. אינה זורקת. כשל בקריאת ה-zip — הודעה ולא מסמך
 * פגום: המסמך הפתוח נשאר כפי שהוא.
 */
export async function editDocxStyles(docx: Blob, edit: StyleEdit): Promise<DocxStyleEditResult> {
  let outcome: ReturnType<typeof editStyles> | null = null;

  const bytes = await rewriteDocxXmlParts<null>(
    await bytesOf(docx),
    (name) => name === STYLES_PART || (edit.kind === 'delete' && PARAGRAPH_PARTS.test(name)),
    (xml, name) => {
      if (!outcome?.ok) return null;
      if (name === STYLES_PART) return outcome.xml === xml ? null : outcome.xml;
      return stripStyleReferences(xml, outcome.removed);
    },
    async (read) => {
      const styles = await read(STYLES_PART);
      outcome = styles === null ? { ok: false, message: 'במסמך אין גיליון סגנונות' } : editStyles(styles, edit);
      return null;
    },
  );

  const result = outcome as ReturnType<typeof editStyles> | null;
  if (!result) return { ok: false, message: 'המסמך אינו קריא' };
  if (!result.ok) return result;
  // ‏`null` מ-`rewriteDocxXmlParts` = שום חלק לא השתנה. עריכה שאינה משנה דבר
  // (דיאלוג שנסגר ב„אישור” בלי שינוי) אינה כשל.
  return { ok: true, docx: bytes ? new Blob([bytes], { type: DOCX_MIME }) : docx, id: result.id };
}
