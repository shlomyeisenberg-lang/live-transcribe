/**
 * „ציטוטים וביבליוגרפיה” — הקבוצה החמישית בלשונית „הפניות” של Word העברי,
 * דרך `doc.citations`.
 *
 * ## הכול כאן נמדד בדפדפן. אלה התוצאות
 *
 * Chrome headless, `file://`, ה-dist האמיתי, מסמך שנוקה ב-`clearContent`,
 * ושבע הרצאות שכולן פורקו גם ברמת ה-zip של `export.toDocx`. הסיבה לפירוק
 * היא הלקח של גל 6: `success: true` אינו עדות, וקוד שדה של Word נבדק מול
 * הקובץ שנכתב ולא מול הקבלה.
 *
 * 1. **המקורות נכתבים לחלק ביבליוגרפיה תקני של OOXML.** לא JSON פרטי ולא
 *    מצאה: `customXml/item1.xml` הוא `<b:Sources>` בסכימת
 *    `…/2006/bibliography`, עם `itemProps1.xml` שמצהיר על ה-`schemaRef`,
 *    עם רלציה מ-`document.xml.rels` ועם `Override` ב-`[Content_Types].xml`.
 *    זה בדיוק המקום שבו Word עצמו מחזיק את „נהל מקורות”.
 * 2. **העברית עוברת שלמה.** נמדד אחד לאחד בקובץ שנכתב: `שו״ת הרמב״ם`
 *    (גרשיים), `בן מימון` / `משה` בשמות, `תתקצ״ה` בשנה, `מוסד הרב קוק`
 *    במוציא לאור, `ירושלים` בעיר. אין בריחה, אין היפוך ואין איבוד תווים.
 * 3. **`citations.insert` כותב שדה `CITATION` תקני**, וזו הנקודה שהפילה את
 *    גל 5 ואת גל 6. נמדד ב-docx:
 *    `<w:fldSimple w:instr="CITATION src-mt82e4rn-golgbf">`, ובאותו קובץ
 *    `<b:Tag>src-mt82e4rn-golgbf</b:Tag>`. כלומר הארגומנט של השדה **הוא**
 *    ה-`Tag` שקיים בחלק המקורות — שלא כמו `REF SDXREF` של גל 5, שהצביע
 *    לשום מקום, ושלא כמו `TA` של גל 6, שאיבד את הציטוט הארוך.
 * 4. **הביבליוגרפיה באמת נבנית ובאמת מתעדכנת.** זו הבדיקה שהכריעה בשלושת
 *    הגלים האחרונים, וכאן היא עוברת: אחרי `bibliography.insert` הבלוק כבר
 *    הכיל את שני המקורות, מקור שנוסף **אחריו** נכנס ב-`rebuild`
 *    (`sourceCount` 2 → 3, והטקסט על המסך גדל בשורה), ועריכת כותרת של מקור
 *    קיים השתקפה גם היא ב-`rebuild`.
 * 5. **ההסרה נקייה.** `bibliography.remove` מפיל את הבלוק כולו: לפני
 *    ההסרה שני בלוקים, אחריה אחד, ו-`getText` חזר למה שהיה. אין שיירים
 *    ואין תלות ב-`blocks.deleteRange` — ההפך מתוכן העניינים של גל 4.
 * 6. **מקור שאינו קיים נדחה.** `citations.insert` עם `sourceIds:['src-nope']`
 *    החזיר `TARGET_NOT_FOUND` ולא כתב כלום. שוב, ההפך מ-`crossRefs.insert`
 *    שהחזיר `success: true` על סימנייה מדומה.
 *
 * ## שני הפערים שבגללם הקבוצה נשלחת חלקית
 *
 * ### א. ציטוט של יותר ממקור אחד — לא נשלח
 *
 * `citations.insert` מקבל `sourceIds: string[]`, ובשניים הוא כותב
 * `CITATION src-a;src-b` (נמדד, גם ב-docx). ב-Word התחביר לריבוי מקורות הוא
 * המתג `\m`: `{ CITATION Tag1 \m Tag2 }`. אסימון אחד שמחבר שני תגים בנקודה
 * ופסיק אינו tag קיים, ולכן Word לא יפתור אותו. המודול שולח לכן **תמיד
 * מקור אחד**, וזה גם מה שהממשק מאפשר.
 *
 * ### ב. „סגנון” — לא נשלח
 *
 * `bibliography.configure` עובד בצד אחד ושבור בצד שני, ושני הצדדים נמדדו
 * באותו קובץ:
 *
 * - **הצד שעובד:** הסגנון מגיע למקום הנכון. `configure({style:'Chicago'})`
 *   כתב `<b:Sources SelectedStyle="/CHICAGO.XSL" StyleName="Chicago"
 *   Version="16">`, וזה בדיוק המקום שבו Word מחזיק את בחירת הסגנון. אחד
 *   עשר השמות הקנוניים ממופים נכון (APA, MLA, Chicago, Harvard - Anglia,
 *   IEEE, שני ה-ISO 690, שני ה-GOST, SIST02, Turabian).
 * - **הצד השבור:** אותה קריאה כותבת גם `BIBLIOGRAPHY \sdStyle "Chicago"`
 *   לתוך קוד השדה. `\sdStyle` אינו מתג של Word — המתגים המתועדים לשדה
 *   `BIBLIOGRAPHY` הם `\l` ו-`\f` — ואין דרך לבקש את הראשון בלי השני.
 *   אותו `\sdStyle` נכתב גם במסלול השני, `bibliography.insert({style})`,
 *   ולכן אין כאן קריאה „נקייה” לעקוף אליה.
 * - ובנוסף, כרגיל: `configure({style:'zigzag'})` חזר `success: true` וכתב
 *   `SelectedStyle="/zigzag.XSL"`, גיליון סגנון שאינו קיים.
 *
 * ההכרעה: לא לחשוף פקד סגנון. בלעדיו כל קוד שדה שהמודול כותב למסמך הוא
 * קנוני — `CITATION <tag>` ו-`BIBLIOGRAPHY` חשוף — והסגנון נשאר ברירת
 * המחדל שגם Word מתחיל בה (APA). המחיר ידוע ומוצהר; החלופה היא לכתוב
 * לכל מסמך מתג שאין לו תיעוד, וזה בדיוק מה ששלושת הגלים הקודמים לימדו לא
 * לעשות.
 *
 * ## איך מוצאים ביבליוגרפיה קיימת, ולמה לא דרך `blocks`
 *
 * ל-`citations.bibliography` אין `list`, ו-`blocks.list` **אינו** מסמן את
 * הבלוק: הביבליוגרפיה מופיעה שם כ-`nodeType: 'paragraph'` רגילה (נמדד).
 * כלומר שתי הדרכים המתבקשות אינן עובדות, והדרך שכן היא `fields.list` —
 * הוא מחזיר `fieldType: 'BIBLIOGRAPHY'` עם `address.blockId`, וכתובת
 * שנבנתה ממנו מניעה את `get`, את `rebuild` ואת `remove` (נמדד).
 *
 * זה גם מה שהופך את „עדכן” ואת „הסר” לשימושיים על מסמך שהגיע **מ-Word**
 * עם ביבליוגרפיה שנוצרה שם, ולא רק על אחת שנוצרה בפעלה הזאת.
 *
 * הכתובת מאומתת במנוע ואינה נבלעת: `rebuild` על מזהה של פסקה רגילה החזיר
 * `TARGET_NOT_FOUND` ולא נגע במסמך.
 *
 * ## מה המנוע בולע בשקט, ולכן נבדק כאן
 *
 * כל אלה חזרו `success: true` וכתבו למסמך מקור פגום:
 *
 * - `fields: {}` — מקור בלי שום שדה. `tag` חוזר `"Source"`.
 * - `fields: { title: '' }` ו-`{ title: '   ' }` — מקור בלי כותרת שאי אפשר
 *   לזהות ברשימה ואי אפשר למחוק בלי לדעת שהוא שם. אותה מלכודת בדיוק כמו
 *   `XE "   "` בגל 5.
 * - `type: 'zigzag'` — סוג שאינו בחוזה, נשמר כמות שהוא.
 * - שדה שאינו בחוזה (`zigzag: '1'`) — נשמר כמות שהוא.
 *
 * ומצד שני, מחבר בלי `last` **זורק** `TypeError` גולמי
 * („Cannot read properties of undefined (reading 'trim')”) ולא קבלה. שתי
 * ההתנהגויות מובילות לאותה מסקנה: הוולידציה יושבת כאן, לפני הקריאה.
 *
 * ## מחיקת מקור שיש לו ציטוט במסמך
 *
 * נמדד: `sources.remove` על מקור שיש עליו שדה `CITATION` מחזיר
 * `success: true`, מוחק את המקור, ו**משאיר** את השדה במסמך מצביע לתג שכבר
 * אינו קיים — כלומר בדיוק המסמך השבור שגל 5 נדחה בגללו, רק שכאן אנחנו אלה
 * שיוצרים אותו. לכן `removeCitationSource` מסרב, ומספר כמה ציטוטים מחזיקים
 * במקור. „הסר את הציטוטים תחילה” היא הוראה שאפשר לבצע; מסמך עם תג יתום
 * אינו מצב שאפשר לתקן בלי לדעת שהוא קיים.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';

/* ------------------------------------------------------------------ */
/* צורת ה-API. מוגדרת כאן ואינה מיובאת — ראו document-api.ts            */
/* ------------------------------------------------------------------ */

/** `CitationSourceAddress` — מה שכל פעולת מקור מקבלת כ-`target`. */
interface SourceAddress {
  kind: 'entity';
  entityType: 'citationSource';
  sourceId: string;
}

/** `BibliographyAddress`. `nodeId` הוא ה-`blockId` שמגיע מ-`fields.list`. */
interface BibliographyAddress {
  kind: 'block';
  nodeType: 'bibliography';
  nodeId: string;
}

/** `CitationPerson` של החוזה. `last` הוא היחיד שאינו אופציונלי שם. */
interface Person {
  first?: string;
  last: string;
}

/** `CitationSourceFields` בחלק שנחשף כאן. ראו `SOURCE_FORM` לנימוק הבחירה. */
interface SourceFields {
  title?: string;
  authors?: Person[];
  year?: string;
  publisher?: string;
  city?: string;
  journalName?: string;
  volume?: string;
  pages?: string;
}

/** `CitationSourceDomain` עטוף ב-`DiscoveryItem`. */
interface RawSource {
  id?: string;
  sourceId?: string;
  tag?: string;
  type?: string;
  fields?: SourceFields;
}

/** `CitationDomain` בחלק שנצרך כאן: רק „מי מצוטט”. */
interface RawCitation {
  sourceIds?: readonly string[];
}

/** `FieldDomain` בחלק שנצרך כאן. ראו „איך מוצאים ביבליוגרפיה קיימת”. */
interface RawField {
  address?: { blockId?: string };
  fieldType?: string;
}

interface DiscoveryPage<T> {
  items?: readonly T[];
  total?: number;
}

export interface CitationsDocumentApi {
  fields?: {
    list?: (query?: { limit?: number; offset?: number }) => MaybePromise<
      DiscoveryPage<RawField> | undefined
    >;
  };
  selection?: {
    current?: (input?: { includeText?: boolean }) => MaybePromise<
      | {
          target?: {
            segments?: readonly { blockId?: string; range?: { start?: number; end?: number } }[];
          } | null;
        }
      | undefined
    >;
  };
  citations?: {
    list?: (query?: { limit?: number; offset?: number }) => MaybePromise<
      DiscoveryPage<RawCitation> | undefined
    >;
    insert?: (input: { at: unknown; sourceIds: string[] }) => MaybePromise<DocReceipt>;
    sources?: {
      list?: (query?: { limit?: number; offset?: number }) => MaybePromise<
        DiscoveryPage<RawSource> | undefined
      >;
      insert?: (input: { type: string; fields: SourceFields }) => MaybePromise<DocReceipt>;
      update?: (input: { target: SourceAddress; patch: SourceFields }) => MaybePromise<DocReceipt>;
      remove?: (input: { target: SourceAddress }) => MaybePromise<DocReceipt>;
    };
    bibliography?: {
      insert?: (input: { at: { kind: 'documentEnd' } }) => MaybePromise<DocReceipt>;
      rebuild?: (input: { target: BibliographyAddress }) => MaybePromise<DocReceipt>;
      remove?: (input: { target: BibliographyAddress }) => MaybePromise<DocReceipt>;
    };
  };
}

/** מה שנדרש מ-SuperDoc: רק הפאסדה של המסמך. ראו page-setup.ts. */
export interface CitationsHost {
  activeEditor?: { doc?: CitationsDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type CitationsTarget = SuperDoc | CitationsHost | null | undefined;

/* ------------------------------------------------------------------ */
/* נוסחי הכשל                                                          */
/* ------------------------------------------------------------------ */

/**
 * ביטוי שלם עם הטיית הכשל בעברית תקנית, ולא שם עצם: „המקור” זכר,
 * „הביבליוגרפיה” נקבה. ראו document-api.ts.
 */
const ADD_SOURCE_FAILED = 'הוספת המקור נכשלה';
const EDIT_SOURCE_FAILED = 'עריכת המקור נכשלה';
const DROP_SOURCE_FAILED = 'מחיקת המקור נכשלה';
const CITE_FAILED = 'הוספת הציטוט נכשלה';
const BIB_INSERT_FAILED = 'הוספת הביבליוגרפיה נכשלה';
const BIB_REBUILD_FAILED = 'עדכון הביבליוגרפיה נכשל';
const BIB_REMOVE_FAILED = 'הסרת הביבליוגרפיה נכשלה';
const READ_FAILED = 'קריאת המקורות נכשלה';

const LOADING_DETAIL = 'המסמך עדיין נטען';
const NO_BIB_DETAIL = 'אין במסמך ביבליוגרפיה';
const AMBIGUOUS_BIB_DETAIL =
  'יש במסמך יותר מביבליוגרפיה אחת, ואין דרך לדעת על איזו מהן הפעולה חלה';
/**
 * `citations.insert` דורש יעד **מכווץ** ולא טווח: נמדד ש-`{start:0,end:5}`
 * חוזר `INVALID_TARGET` עם „requires a collapsed text target”. לכן המודול
 * מכווץ בעצמו לסוף הבחירה — וזו גם ההתנהגות של Word, שמכניס את הציטוט אחרי
 * הטקסט המסומן ולא במקומו.
 */
const NO_CARET_DETAIL = 'יש למקם את הסמן במסמך במקום שבו יופיע הציטוט';

function unavailable(failedAction: string, detail: string, reason: string): CommandOutcome {
  return { ok: false, message: `${failedAction}: ${detail}`, reason };
}

/** הנוסח שהתכנית קובעת ב-§12, וזהה לזה שהיכולת מחזירה. ראו footnotes.ts. */
function unsupported(failedAction: string): CommandOutcome {
  return {
    ok: false,
    message: `${failedAction}: אינו זמין בגרסה זו`,
    reason: 'command-unsupported',
  };
}

/** קריאה למנוע שאינה זורקת החוצה. ראו הערת ה„לעולם לא זורקת” ב-footnotes.ts. */
async function attempt<T>(
  failedAction: string,
  call: () => MaybePromise<T>,
): Promise<{ ok: true; value: T } | { ok: false; outcome: CommandOutcome }> {
  try {
    return { ok: true, value: await call() };
  } catch (error) {
    return {
      ok: false,
      outcome: { ok: false, message: thrownText(failedAction, error), reason: 'threw' },
    };
  }
}

/** כשל הקבלה, או `null` כשהיא הצליחה. `NO_OP` נחשב הצלחה — ראו header-footer.ts. */
function failureOf(failedAction: string, receipt: DocReceipt | undefined): CommandOutcome | null {
  const code = receipt?.failure?.code;
  if (receipt?.success !== false || code === 'NO_OP') return null;
  return { ok: false, message: receiptFailureText(failedAction, receipt), reason: code };
}

function docOf(host: CitationsTarget): CitationsDocumentApi | null {
  return (host as CitationsHost | null | undefined)?.activeEditor?.doc ?? null;
}

/** גודל העמוד בכל שאיבה, ובלם מפני מנוע שיחזיר `total` שאינו יורד. ראו fields.ts. */
const PAGE_SIZE = 200;
const PAGE_GUARD = 1000;

/**
 * כל הפריטים של פעולת discovery, בשאיבת עמודים עד `total`.
 *
 * `total` ולא `items.length`: שלוש הרשימות כאן הן `DiscoveryOutput`, כלומר
 * `items` הוא עמוד תחת `limit`/`offset`. ספר תורני עם רשימת מקורות ארוכה הוא
 * בדיוק המסמך שבו העמוד הראשון אינו הכול — וזה התרחיש של אוצריא, לא מקרה קצה.
 */
async function collectAll<T>(
  failedAction: string,
  list: (query: { limit: number; offset: number }) => MaybePromise<DiscoveryPage<T> | undefined>,
): Promise<{ ok: true; items: T[] } | { ok: false; items: T[]; outcome: CommandOutcome }> {
  const items: T[] = [];
  let offset = 0;
  let guard = 0;

  for (;;) {
    const listed = await attempt(failedAction, () => list({ limit: PAGE_SIZE, offset }));
    if (!listed.ok) return { ok: false, items, outcome: listed.outcome };

    const page = listed.value?.items ?? [];
    items.push(...page);
    if (page.length === 0) return { ok: true, items };

    offset += page.length;

    const total = listed.value?.total;
    if (!Number.isFinite(total) || offset >= (total as number)) return { ok: true, items };
    if (++guard > PAGE_GUARD) return { ok: true, items };
  }
}

/* ------------------------------------------------------------------ */
/* ולידציה — כאן ולא במנוע                                             */
/* ------------------------------------------------------------------ */

/**
 * סוגי המקור שנחשפים, עם התווית שהמשתמש רואה.
 *
 * שמונה מתוך שנים-עשר. מה שנשאר בחוץ — `patent`, `case`, `statute`, `film` —
 * הוא ארבעה סוגים שכולם משפטיים או קולנועיים, כלומר קטגוריות של כתיבה
 * אקדמית מערבית שאין להן מקבילה בספר תורני. הרשימה גם עצמה: `misc` („אחר”)
 * קולט כל דבר שאינו מכוסה, ולכן שום מקור אינו נשאר בלי סוג.
 *
 * הרשימה נבדקת כאן ולא במנוע מפני שנמדד שהוא בולע `type: 'zigzag'` בשקט
 * וכותב אותו לקובץ.
 */
export const CITATION_SOURCE_TYPES = [
  { value: 'book', label: 'ספר' },
  { value: 'journalArticle', label: 'מאמר בכתב עת' },
  { value: 'conferenceProceedings', label: 'מאמר בקובץ' },
  { value: 'thesis', label: 'עבודת מחקר' },
  { value: 'report', label: 'דוח' },
  { value: 'website', label: 'אתר אינטרנט' },
  { value: 'interview', label: 'ראיון' },
  { value: 'misc', label: 'אחר' },
] as const;

export type CitationSourceType = (typeof CITATION_SOURCE_TYPES)[number]['value'];

export const DEFAULT_CITATION_SOURCE_TYPE: CitationSourceType = 'book';

/**
 * הסוגים שבהם שדות כתב העת רלוונטיים. מודל ולא תצוגה: הדיאלוג שואל את
 * הפונקציה הזאת אילו שדות להראות, והמודול משתמש בה כדי לא לשלוח `volume`
 * ו-`journalName` על ספר — שדות שהמנוע יכתוב לקובץ גם כשאין להם מובן.
 */
export function usesJournalFields(type: CitationSourceType): boolean {
  return type === 'journalArticle' || type === 'conferenceProceedings';
}

function isKnownSourceType(value: string): value is CitationSourceType {
  return CITATION_SOURCE_TYPES.some((entry) => entry.value === value);
}

/** מה שמוצג כשהכותרת נדחית. */
export const CITATION_TITLE_HINT = 'יש להקליד את כותרת המקור';

/** מה שמוצג כשסוג המקור נדחה. אינו אמור להגיע למשתמש — הבחירה היא רשימה סגורה. */
export const CITATION_TYPE_HINT = 'סוג המקור אינו מוכר';

/**
 * הכותרת כפי שתישלח, או `null` כשהיא ריקה.
 *
 * ה-`trim` אינו קוסמטיקה, וזו אותה מלכודת בדיוק שנמדדה ב-`XE` בגל 5: המנוע
 * מקבל בהצלחה `title: '   '` וכותב `<b:Title>   </b:Title>` — מקור שאי אפשר
 * לזהות ברשימה ואי אפשר למחוק בלי לדעת שהוא שם.
 */
export function normalizeCitationTitle(raw: string): string | null {
  const title = raw.trim();
  return title === '' ? null : title;
}

/**
 * שורות המחברים כפי שהמשתמש מקליד אותן, למבנה של החוזה.
 *
 * הצורה החופשית היא הכרעה ולא עצלות. `CitationPerson` דורש `last` ומציע
 * `first`, כלומר המודל של Word הוא „שם משפחה, שם פרטי” — ולמחצית מהמקורות
 * בספר תורני אין פיצול כזה בכלל: „רמב״ם”, „חזון איש”, „שולחן ערוך” הם שם
 * אחד ולא שניים. לכן שורה בלי פסיק נכנסת כולה ל-`last`, ושורה עם פסיק
 * מתפצלת — מי שרוצה „כהן, יוסף” מקבל אותו, ומי שכותב „רמב״ם” אינו נאלץ
 * להמציא שם פרטי.
 *
 * שורות ריקות מדולגות, ו-`last` ריק אינו נשלח לעולם: נמדד שמחבר בלי `last`
 * מפיל את המנוע ב-`TypeError` גולמי ולא בקבלה.
 */
export function parseCitationAuthors(raw: string): Person[] {
  const people: Person[] = [];
  for (const line of raw.split('\n')) {
    const trimmed = line.trim();
    if (trimmed === '') continue;

    const comma = trimmed.indexOf(',');
    if (comma === -1) {
      people.push({ last: trimmed });
      continue;
    }

    const last = trimmed.slice(0, comma).trim();
    const first = trimmed.slice(comma + 1).trim();
    if (last === '') continue;
    people.push(first === '' ? { last } : { last, first });
  }
  return people;
}

/** ההפך מ-`parseCitationAuthors`, למילוי הטופס בעריכת מקור קיים. */
export function formatCitationAuthors(people: readonly Person[] | undefined): string {
  if (!Array.isArray(people)) return '';
  return people
    .map((person) => {
      const last = typeof person?.last === 'string' ? person.last : '';
      const first = typeof person?.first === 'string' ? person.first : '';
      return first === '' ? last : `${last}, ${first}`;
    })
    .filter((line) => line !== '')
    .join('\n');
}

/* ------------------------------------------------------------------ */
/* המודל שהממשק עובד מולו                                              */
/* ------------------------------------------------------------------ */

/** מה שהטופס מחזיק. מחרוזות בלבד — זו הצורה שקלט טקסט נותן. */
export interface CitationSourceDraft {
  type: CitationSourceType;
  title: string;
  /** שורה למחבר. ראו `parseCitationAuthors`. */
  authors: string;
  year: string;
  publisher: string;
  city: string;
  journalName: string;
  volume: string;
  pages: string;
}

export function emptyCitationSourceDraft(): CitationSourceDraft {
  return {
    type: DEFAULT_CITATION_SOURCE_TYPE,
    title: '',
    authors: '',
    year: '',
    publisher: '',
    city: '',
    journalName: '',
    volume: '',
    pages: '',
  };
}

/** מקור אחד כפי שהדיאלוגים מציגים אותו. */
export interface CitationSourceSummary {
  /** `sourceId` של המנוע. גם מפתח `v-for` וגם היעד של כל פעולה. */
  id: string;
  /** מה שמוצג ברשימה: כותרת, ואחריה המחבר והשנה כשהם קיימים. */
  label: string;
  /** הטופס מלא מראש, לעריכה. */
  draft: CitationSourceDraft;
  /** כמה שדות `CITATION` במסמך מצביעים אל המקור. 0 = אפשר למחוק בבטחה. */
  citedCount: number;
}

/** מה שהממשק צריך לדעת. תצלום ולא מנוי, כמו header-footer.ts. */
export interface CitationsState {
  sources: readonly CitationSourceSummary[];
  /** מספר הציטוטים במסמך. */
  citationCount: number;
  /** מספר הביבליוגרפיות במסמך. 0 = אין מה לעדכן, >1 = אי-אפשר להסיר. */
  bibliographyCount: number;
}

export function emptyCitationsState(): CitationsState {
  return { sources: [], citationCount: 0, bibliographyCount: 0 };
}

/** התווית שברשימה. הכותרת היא העוגן; המחבר והשנה מבדילים בין מהדורות. */
function labelOf(draft: CitationSourceDraft): string {
  const detail = [draft.authors.split('\n')[0]?.trim(), draft.year.trim()]
    .filter((part) => part !== undefined && part !== '')
    .join(', ');
  const title = draft.title.trim() === '' ? 'מקור בלי כותרת' : draft.title.trim();
  return detail === '' ? title : `${title} (${detail})`;
}

function text(value: string | undefined): string {
  return typeof value === 'string' ? value : '';
}

/** `CitationSourceDomain` → הטופס. סוג שאינו מוכר יורד ל„אחר”, ולא נעלם. */
function draftOf(raw: RawSource): CitationSourceDraft {
  const fields = raw.fields ?? {};
  const type = typeof raw.type === 'string' && isKnownSourceType(raw.type) ? raw.type : 'misc';
  return {
    type,
    title: text(fields.title),
    authors: formatCitationAuthors(fields.authors),
    year: text(fields.year),
    publisher: text(fields.publisher),
    city: text(fields.city),
    journalName: text(fields.journalName),
    volume: text(fields.volume),
    pages: text(fields.pages),
  };
}

/**
 * הטופס → `CitationSourceFields`.
 *
 * **בהוספה** שדה ריק אינו נשלח כלל ולא נשלח כמחרוזת ריקה: נמדד שהמנוע כותב
 * לקובץ כל מה שהוא מקבל, ו-`<b:Year></b:Year>` ריק בחלק הביבליוגרפיה הוא
 * רעש ש-Word יצטרך להתמודד איתו.
 *
 * **בעריכה זה הפוך**, ומטעם שאין דרך לעקוף אותו: `patch` הוא
 * `Partial<CitationSourceFields>` אמיתי — שדה שאינו בו נשאר במסמך כמו
 * שהיה, ורק מחרוזת ריקה מוחקת אותו (נמדד: patch בלי `year` השאיר
 * `תש״ף` במסמך). משתמש שמחק את השנה מהטופס ולחץ „שמור שינויים” היה מקבל
 * `{ok:true}` ורואה את השנה חוזרת ברענון — הצלחה מדומה. לכן שדה שרוקן
 * נשלח בעריכה כמחרוזת ריקה.
 *
 * מה שנשאר מושמט בשני המצבים הוא שדה שאינו רלוונטי לסוג: שדות כתב העת
 * נשלחים רק לסוגים שיש להם בהם מובן (ראו `usesJournalFields`), ובסוג אחר
 * אין להם ערך למחוק.
 */
function fieldsOf(
  draft: CitationSourceDraft,
  title: string,
  mode: 'insert' | 'update',
): SourceFields {
  const fields: SourceFields = { title };
  const erasing = mode === 'update';

  const authors = parseCitationAuthors(draft.authors);
  if (authors.length > 0 || erasing) fields.authors = authors;

  const year = draft.year.trim();
  if (year !== '' || erasing) fields.year = year;
  const publisher = draft.publisher.trim();
  if (publisher !== '' || erasing) fields.publisher = publisher;
  const city = draft.city.trim();
  if (city !== '' || erasing) fields.city = city;

  if (usesJournalFields(draft.type)) {
    const journalName = draft.journalName.trim();
    if (journalName !== '' || erasing) fields.journalName = journalName;
    const volume = draft.volume.trim();
    if (volume !== '' || erasing) fields.volume = volume;
    const pages = draft.pages.trim();
    if (pages !== '' || erasing) fields.pages = pages;
  }

  return fields;
}

/* ------------------------------------------------------------------ */
/* קריאה                                                               */
/* ------------------------------------------------------------------ */

/** כמה ציטוטים מצביעים לכל `sourceId`. ריק כשאין `citations.list`. */
async function citedCounts(
  host: CitationsTarget,
  failedAction: string,
): Promise<{ ok: boolean; counts: Map<string, number>; total: number }> {
  const list = docOf(host)?.citations?.list;
  if (typeof list !== 'function') return { ok: false, counts: new Map(), total: 0 };

  const listed = await collectAll<RawCitation>(failedAction, (query) => list(query));
  const counts = new Map<string, number>();
  for (const citation of listed.items) {
    for (const sourceId of citation.sourceIds ?? []) {
      if (typeof sourceId !== 'string') continue;
      counts.set(sourceId, (counts.get(sourceId) ?? 0) + 1);
    }
  }
  return { ok: listed.ok, counts, total: listed.items.length };
}

/**
 * המקורות שבמסמך. לעולם אינה זורקת: כשל של קריאה מחזיר רשימה ריקה, כלומר
 * הדיאלוג יאמר „אין מקורות” — ולא ימציא רשומות.
 *
 * מקור בלי `sourceId` מדולג ואינו מוחזר: הצגתו הייתה מייצרת שורה שלחיצה
 * עליה שולחת `undefined` ל-`sources.remove`.
 */
export async function listCitationSources(
  host: CitationsTarget,
): Promise<CitationSourceSummary[]> {
  return collectSources(host, (await citedCounts(host, READ_FAILED)).counts);
}

/** הגוף של `listCitationSources`, עם ספירת הציטוטים שכבר נקראה. */
async function collectSources(
  host: CitationsTarget,
  counts: Map<string, number>,
): Promise<CitationSourceSummary[]> {
  const list = docOf(host)?.citations?.sources?.list;
  if (typeof list !== 'function') return [];

  const listed = await collectAll<RawSource>(READ_FAILED, (query) => list(query));

  const sources: CitationSourceSummary[] = [];
  for (const raw of listed.items) {
    const id = typeof raw.sourceId === 'string' ? raw.sourceId : raw.id;
    if (typeof id !== 'string' || id === '') continue;
    const draft = draftOf(raw);
    sources.push({ id, label: labelOf(draft), draft, citedCount: counts.get(id) ?? 0 });
  }
  return sources;
}

/**
 * הכתובות של הביבליוגרפיות שבמסמך.
 *
 * דרך `fields.list` ולא `blocks.list`: הביבליוגרפיה מופיעה ב-`blocks.list`
 * כפסקה רגילה, ואין ל-`citations.bibliography` פעולת `list` משלה. ההנמקה
 * המלאה בהערת הפתיחה.
 */
async function listBibliographies(
  host: CitationsTarget,
  failedAction: string,
): Promise<{ ok: true; items: BibliographyAddress[] } | { ok: false; outcome: CommandOutcome }> {
  const list = docOf(host)?.fields?.list;
  if (typeof list !== 'function') return { ok: false, outcome: unsupported(failedAction) };

  const listed = await collectAll<RawField>(failedAction, (query) => list(query));
  if (!listed.ok) return { ok: false, outcome: listed.outcome };

  const items: BibliographyAddress[] = [];
  for (const field of listed.items) {
    if (field.fieldType !== 'BIBLIOGRAPHY') continue;
    const nodeId = field.address?.blockId;
    if (typeof nodeId !== 'string' || nodeId === '') continue;
    items.push({ kind: 'block', nodeType: 'bibliography', nodeId });
  }
  return { ok: true, items };
}

/**
 * קוראת את מצב הציטוטים במסמך. לעולם אינה זורקת: כשל של קריאה מחזיר „אין”,
 * כלומר ה-tooltip יאמר שאין מה לעדכן — ולא ימציא מספר.
 */
export async function readCitationsState(host: CitationsTarget): Promise<CitationsState> {
  const doc = docOf(host);
  if (!doc) return emptyCitationsState();

  // שלוש הקריאות במקביל, וספירת הציטוטים נקראת **פעם אחת** ומועברת פנימה:
  // היא נדרשת גם למונה וגם ל-`citedCount` של כל מקור, וקריאה כפולה הייתה
  // מכפילה את זמן הפתיחה של הדיאלוג על מסמך עם מאות ציטוטים.
  const [bibliographies, cited] = await Promise.all([
    listBibliographies(host, READ_FAILED),
    citedCounts(host, READ_FAILED),
  ]);

  return {
    sources: await collectSources(host, cited.counts),
    citationCount: cited.ok ? cited.total : 0,
    bibliographyCount: bibliographies.ok ? bibliographies.items.length : 0,
  };
}

/**
 * הביבליוגרפיה היחידה שבמסמך, או כשל מנומק.
 *
 * שלושת המצבים נפרדים בכוונה: „אין”, „יש כמה” ו„הקריאה נכשלה” הן שלוש
 * הודעות שונות למשתמש, ואיחוד שלהן לאחת היה שולח אותו לחפש את הבעיה הלא
 * נכונה. אותה החלטה בדיוק כמו ב-index-field.ts.
 */
async function soleBibliography(
  host: CitationsTarget,
  failedAction: string,
): Promise<{ ok: true; address: BibliographyAddress } | { ok: false; outcome: CommandOutcome }> {
  const listed = await listBibliographies(host, failedAction);
  if (!listed.ok) return listed;

  if (listed.items.length === 0) {
    return { ok: false, outcome: unavailable(failedAction, NO_BIB_DETAIL, 'no-bibliography') };
  }
  if (listed.items.length > 1) {
    return {
      ok: false,
      outcome: unavailable(failedAction, AMBIGUOUS_BIB_DETAIL, 'ambiguous-bibliography'),
    };
  }
  return { ok: true, address: listed.items[0] };
}

/* ------------------------------------------------------------------ */
/* ניהול מקורות                                                        */
/* ------------------------------------------------------------------ */

/** הבדיקות שקודמות לכל שליחה של מקור. ראו „מה המנוע בולע בשקט”. */
function validateDraft(
  failedAction: string,
  draft: CitationSourceDraft,
): { ok: true; title: string } | { ok: false; outcome: CommandOutcome } {
  if (!isKnownSourceType(draft.type)) {
    return {
      ok: false,
      outcome: { ok: false, message: `${failedAction}: ${CITATION_TYPE_HINT}`, reason: 'invalid-type' },
    };
  }

  const title = normalizeCitationTitle(draft.title);
  if (title === null) {
    return {
      ok: false,
      outcome: { ok: false, message: `${failedAction}: ${CITATION_TITLE_HINT}`, reason: 'invalid-title' },
    };
  }
  return { ok: true, title };
}

/** „מקור חדש” — מוסיפה מקור לחלק הביבליוגרפיה של המסמך. */
export async function addCitationSource(
  host: CitationsTarget,
  draft: CitationSourceDraft,
): Promise<CommandOutcome> {
  const valid = validateDraft(ADD_SOURCE_FAILED, draft);
  if (!valid.ok) return valid.outcome;

  const doc = docOf(host);
  if (!doc) return unavailable(ADD_SOURCE_FAILED, LOADING_DETAIL, 'document-api-unavailable');

  const insert = doc.citations?.sources?.insert;
  if (typeof insert !== 'function') return unsupported(ADD_SOURCE_FAILED);

  const added = await attempt(ADD_SOURCE_FAILED, () =>
    insert({ type: draft.type, fields: fieldsOf(draft, valid.title, 'insert') }),
  );
  if (!added.ok) return added.outcome;

  return failureOf(ADD_SOURCE_FAILED, added.value) ?? { ok: true };
}

/**
 * „ערוך מקור” — מחליפה את שדות המקור.
 *
 * `patch` הוא `Partial` אמיתי: השמטה משמרת את הערך שבמסמך, ורק מחרוזת
 * ריקה מוחקת אותו (נמדד). זה טופס עריכה, ומחיקת שנה בשדה חייבת למחוק
 * אותה גם במסמך — ולכן כל שדה שרלוונטי לסוג נשלח כאן גם כשהוא ריק, ולא
 * מושמט. ראו `fieldsOf`. הסוג עצמו אינו ניתן לשינוי —
 * `CitationSourceUpdateInput.patch` הוא `Partial<CitationSourceFields>`
 * בלבד, ואין בחוזה מסלול שמשנה `type`.
 */
export async function updateCitationSource(
  host: CitationsTarget,
  sourceId: string,
  draft: CitationSourceDraft,
): Promise<CommandOutcome> {
  const valid = validateDraft(EDIT_SOURCE_FAILED, draft);
  if (!valid.ok) return valid.outcome;

  const doc = docOf(host);
  if (!doc) return unavailable(EDIT_SOURCE_FAILED, LOADING_DETAIL, 'document-api-unavailable');

  const update = doc.citations?.sources?.update;
  if (typeof update !== 'function') return unsupported(EDIT_SOURCE_FAILED);

  const updated = await attempt(EDIT_SOURCE_FAILED, () =>
    update({
      target: { kind: 'entity', entityType: 'citationSource', sourceId },
      patch: fieldsOf(draft, valid.title, 'update'),
    }),
  );
  if (!updated.ok) return updated.outcome;

  return failureOf(EDIT_SOURCE_FAILED, updated.value) ?? { ok: true };
}

/**
 * „מחק מקור” — מסירה מקור מחלק הביבליוגרפיה.
 *
 * מסרבת כשיש למקור ציטוט במסמך, וזה **לב הפקד** ולא זהירות יתר: נמדד
 * שהמנוע מוחק בהצלחה ומשאיר את שדה ה-`CITATION` מצביע לתג שאינו קיים —
 * כלומר מייצר בדיוק את המסמך השבור שגל 5 נדחה בגללו. ההנמקה המלאה בהערת
 * הפתיחה.
 */
export async function removeCitationSource(
  host: CitationsTarget,
  sourceId: string,
): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(DROP_SOURCE_FAILED, LOADING_DETAIL, 'document-api-unavailable');

  const remove = doc.citations?.sources?.remove;
  if (typeof remove !== 'function') return unsupported(DROP_SOURCE_FAILED);

  const cited = await citedCounts(host, DROP_SOURCE_FAILED);
  // כשל של הספירה מונע את המחיקה: „לא הצלחנו לבדוק” אינו „אין ציטוטים”,
  // והמחיר של הטעות הוא שדה יתום שאי אפשר לזהות בלי לחפש אותו.
  if (!cited.ok) {
    return unavailable(
      DROP_SOURCE_FAILED,
      'לא ניתן היה לבדוק אילו ציטוטים מפנים אל המקור',
      'citations-unreadable',
    );
  }

  const uses = cited.counts.get(sourceId) ?? 0;
  if (uses > 0) {
    const detail =
      uses === 1
        ? 'יש במסמך ציטוט אחד שמפנה אל המקור. יש להסיר אותו תחילה'
        : `יש במסמך ${uses} ציטוטים שמפנים אל המקור. יש להסיר אותם תחילה`;
    return unavailable(DROP_SOURCE_FAILED, detail, 'source-in-use');
  }

  const removed = await attempt(DROP_SOURCE_FAILED, () =>
    remove({ target: { kind: 'entity', entityType: 'citationSource', sourceId } }),
  );
  if (!removed.ok) return removed.outcome;

  return failureOf(DROP_SOURCE_FAILED, removed.value) ?? { ok: true };
}

/* ------------------------------------------------------------------ */
/* הוספת ציטוט                                                         */
/* ------------------------------------------------------------------ */

/**
 * ה-`TextTarget` המכווץ שאליו הציטוט נכנס, או כשל מנומק.
 *
 * הכיווץ הוא **לסוף** הבחירה ולא לתחילתה: `citations.insert` דורש יעד
 * מכווץ (נמדד — טווח חוזר `INVALID_TARGET`), ו-Word מכניס את הציטוט אחרי
 * הטקסט המסומן. משתמש שסימן „שולחן ערוך” וביקש ציטוט מצפה לראות אותו
 * אחרי המילים, לא לפניהן.
 *
 * הבחירה נקראת כאן ולא דרך `readDocSelection`: אותו מודול מחזיר את היעד
 * **כמו שהוא**, וזה בדיוק מה שאסור כאן — יעד עם טווח ייפסל במנוע. מה
 * שנדרש הוא ההיסט עצמו, והוא אינו חלק מהתצלום שהמודול ההוא מחזיר.
 */
async function caretTarget(
  host: CitationsTarget,
  failedAction: string,
): Promise<{ ok: true; at: unknown } | { ok: false; outcome: CommandOutcome }> {
  const current = docOf(host)?.selection?.current;
  if (typeof current !== 'function') return { ok: false, outcome: unsupported(failedAction) };

  const read = await attempt(failedAction, () => current());
  if (!read.ok) return { ok: false, outcome: read.outcome };

  // המקטע ה**אחרון** התקין ולא הראשון: בחירה שחוצה פסקאות מדווחת מקטע לכל
  // פסקה, וסופה של הבחירה הוא סוף האחרון. עצירה על הראשון הייתה שותלת את
  // הציטוט באמצע הטקסט המסומן. מקטע פסול מדולג כדי שמקטע ריק לא יבטל בחירה
  // שיש לה סוף תקין.
  let at: unknown;
  for (const segment of read.value?.target?.segments ?? []) {
    const blockId = segment?.blockId;
    const end = segment?.range?.end ?? segment?.range?.start;
    if (typeof blockId !== 'string' || blockId === '' || typeof end !== 'number') continue;
    at = { kind: 'text', segments: [{ blockId, range: { start: end, end } }] };
  }
  if (at !== undefined) return { ok: true, at };

  return { ok: false, outcome: unavailable(failedAction, NO_CARET_DETAIL, 'no-caret') };
}

/**
 * „הוסף ציטוט” — מכניסה שדה `CITATION` במקום הסמן.
 *
 * מקור **אחד** ולא רשימה, וזו הכרעה שנמדדה: שניים מייצרים
 * `CITATION src-a;src-b`, ותחביר ריבוי המקורות של Word הוא המתג `\m`.
 * ההנמקה המלאה בהערת הפתיחה.
 */
export async function insertCitation(
  host: CitationsTarget,
  sourceId: string,
): Promise<CommandOutcome> {
  if (typeof sourceId !== 'string' || sourceId.trim() === '') {
    return unavailable(CITE_FAILED, 'יש לבחור מקור', 'no-source');
  }

  const doc = docOf(host);
  if (!doc) return unavailable(CITE_FAILED, LOADING_DETAIL, 'document-api-unavailable');

  const insert = doc.citations?.insert;
  if (typeof insert !== 'function') return unsupported(CITE_FAILED);

  const caret = await caretTarget(host, CITE_FAILED);
  if (!caret.ok) return caret.outcome;

  const inserted = await attempt(CITE_FAILED, () =>
    insert({ at: caret.at, sourceIds: [sourceId] }),
  );
  if (!inserted.ok) return inserted.outcome;

  return failureOf(CITE_FAILED, inserted.value) ?? { ok: true };
}

/* ------------------------------------------------------------------ */
/* ביבליוגרפיה                                                         */
/* ------------------------------------------------------------------ */

/**
 * „ביבליוגרפיה” — מכניסה בלוק `BIBLIOGRAPHY` בסוף המסמך.
 *
 * `documentEnd` ולא מיקום הסמן, מאותו טעם בדיוק כמו במפתח: רשימת מקורות
 * היא נספח, והמקום שלה בספר הוא הסוף. ביבליוגרפיה שנשתלת באמצע פרק היא
 * נזק שקשה לבטל.
 *
 * בלי `style`: `bibliography.insert` מקבל אותו, וכל ערך שהוא מקבל נכתב גם
 * כמתג `\sdStyle` שאינו של Word. ההנמקה המלאה בהערת הפתיחה.
 */
export async function insertBibliography(host: CitationsTarget): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(BIB_INSERT_FAILED, LOADING_DETAIL, 'document-api-unavailable');

  const insert = doc.citations?.bibliography?.insert;
  if (typeof insert !== 'function') return unsupported(BIB_INSERT_FAILED);

  const inserted = await attempt(BIB_INSERT_FAILED, () => insert({ at: { kind: 'documentEnd' } }));
  if (!inserted.ok) return inserted.outcome;

  return failureOf(BIB_INSERT_FAILED, inserted.value) ?? { ok: true };
}

/**
 * מה שמוחזר כשהעדכון רץ אך לא על כל הביבליוגרפיות. „לא הושלם” ולא „נכשל”:
 * משהו כן עודכן, והודעה שאומרת „נכשל” הייתה שולחת את המשתמש לבדוק דווקא
 * את זו שנבנתה מחדש. אותה הבחנה כמו ב-toc.ts וב-index-field.ts.
 */
function partialRebuild(unreachable: number): CommandOutcome {
  const left = unreachable === 1 ? 'אחת מהן לא עודכנה' : `${unreachable} מהן לא עודכנו`;
  return {
    ok: false,
    message: `עדכון הביבליוגרפיה לא הושלם: יש במסמך כמה ביבליוגרפיות שאינן ניתנות להבחנה זו מזו, ולכן ${left}`,
    reason: 'ambiguous-bibliography',
  };
}

/**
 * „עדכן ביבליוגרפיה” — בונה מחדש כל ביבליוגרפיה שאפשר לפנות אליה.
 *
 * רצה על כולן ולא על היחידה: שלא כמו בהסרה, בנייה מחדש אינה הרסנית ואינה
 * דו-משמעית — כל ביבליוגרפיה נבנית מאותם מקורות. הבלם על כתובת חוזרת נשאר
 * בכל זאת, מאותו טעם כמו במפתח: גל 4 מדד מנוע שנותן לשני עצמים את אותו
 * `nodeId`, ולולאה תמימה במצב כזה הייתה בונה את הראשונה פעמיים ומדווחת
 * „בוצע” על שנייה שנשארה מיושנת.
 *
 * עצירה בכשל הראשון: מסמך שחציו עודכן וחציו לא הוא מצב שאי אפשר לתאר
 * למשתמש.
 */
export async function rebuildBibliography(host: CitationsTarget): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(BIB_REBUILD_FAILED, LOADING_DETAIL, 'document-api-unavailable');

  const rebuild = doc.citations?.bibliography?.rebuild;
  if (typeof rebuild !== 'function') return unsupported(BIB_REBUILD_FAILED);

  const listed = await listBibliographies(host, BIB_REBUILD_FAILED);
  if (!listed.ok) return listed.outcome;
  if (listed.items.length === 0) {
    return unavailable(BIB_REBUILD_FAILED, NO_BIB_DETAIL, 'no-bibliography');
  }

  const sent = new Set<string>();
  let unreachable = 0;

  for (const address of listed.items) {
    if (sent.has(address.nodeId)) {
      unreachable++;
      continue;
    }
    sent.add(address.nodeId);

    const rebuilt = await attempt(BIB_REBUILD_FAILED, () => rebuild({ target: address }));
    if (!rebuilt.ok) return rebuilt.outcome;
    const failure = failureOf(BIB_REBUILD_FAILED, rebuilt.value);
    if (failure) return failure;
  }

  if (unreachable > 0) return partialRebuild(unreachable);

  return { ok: true };
}

/**
 * „הסר ביבליוגרפיה” — מוחקת את בלוק ה-`BIBLIOGRAPHY`.
 *
 * צעד אחד, בלי ניקוי שיירים: נמדד שאחרי `remove` לא נשארת ולו פסקה אחת —
 * ההפך מתוכן העניינים של גל 4. המקורות עצמם נשארים במסמך, וזו הכוונה:
 * הם עבודת המשתמש, והביבליוגרפיה היא רק התצוגה שלהם.
 *
 * על היחידה ולא על כולן: מחיקה היא הרסנית, ואין דרך ציבורית לדעת על איזו
 * מהן המשתמש התכוון — `doc.selection` מחזיר `blockId` של פסקה בלבד. אותה
 * החלטה כמו ב„הסר מפתח”.
 */
export async function removeBibliography(host: CitationsTarget): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(BIB_REMOVE_FAILED, LOADING_DETAIL, 'document-api-unavailable');

  const remove = doc.citations?.bibliography?.remove;
  if (typeof remove !== 'function') return unsupported(BIB_REMOVE_FAILED);

  const sole = await soleBibliography(host, BIB_REMOVE_FAILED);
  if (!sole.ok) return sole.outcome;

  const removed = await attempt(BIB_REMOVE_FAILED, () => remove({ target: sole.address }));
  if (!removed.ok) return removed.outcome;

  return failureOf(BIB_REMOVE_FAILED, removed.value) ?? { ok: true };
}
