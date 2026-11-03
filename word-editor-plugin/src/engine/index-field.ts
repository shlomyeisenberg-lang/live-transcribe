/**
 * „מפתח” — הקבוצה השלישית בלשונית „הפניות” של Word העברי (סמן ערך, הוסף
 * מפתח, עדכן מפתח), דרך `doc.index`.
 *
 * הקובץ אינו `index.ts` בכוונה: קובץ בשם הזה בתיקיית `engine/` היה נחטף
 * כברירת המחדל של `import … from './engine'`, וכל ייבוא של מודול אחר בתיקייה
 * היה עלול להיפתר אליו.
 *
 * ## הכול כאן נמדד בדפדפן. אלה התוצאות
 *
 * Chrome headless, `file://`, ה-dist האמיתי, מסמך שנוקה ב-`clearContent`
 * וזרוע בשמונה פסקאות HTML, ביום 25.8.2026:
 *
 * 1. **`index.entries.insert` כותב שדה `XE` תקני של Word**, והעברית עוברת
 *    אותו שלמה. נמדד אחד לאחד: `XE "רש״י"` (גרשיים), `XE "ר' עקיבא"` (גרש),
 *    `XE "בְּרֵאשִׁית"` (מנוקד, כולל דגש ושווא), `XE "Rashi"`. אין בריחה, אין
 *    היפוך ואין איבוד ניקוד — לא כמו שם הסימנייה בגל 3.
 * 2. **`index.insert` כותב `INDEX \h "A"`** — קוד Word תקני — והמפתח נבנה
 *    **מלא כבר ביצירה**: `getText` קיבל מיד את כל שמונת הערכים.
 * 3. **`index.rebuild` מחזיר `success: true` ואינו משנה דבר במסמך ריק־עדכון.**
 *    נמדד: `getText` לפני ואחרי זהים תו בתו, ו-`entryCount` נשאר 8. זו אינה
 *    סיבה לוותר עליו — הוא „עדכן מפתח” של Word, והוא זה שיאסוף ערכים שסומנו
 *    אחרי היצירה — אבל הוא **אינו** מה שממלא את המפתח בפעם הראשונה.
 *
 * ## הממצא הגדול: המפתח שעל המסך אינו מפתח
 *
 * זו התגלית שקובעת מה מותר להבטיח למשתמש. המנוע מרנדר את בלוק ה-`INDEX`
 * כרשימת הערכים **בסדר הופעתם במסמך**, בלי מיון כלשהו, בלי מספרי עמודים
 * ובלי כותרות אותיות.
 *
 * נמדד במפורש: נשלחו שמונה ערכים בסדר `תשובה, אברהם, בית, גמרא, Zayin,
 * Alef, משנה, דעת`, ואחרי `insert` ואחרי `rebuild` הבלוק הכיל
 * `"תשובה\nאברהם\nבית\nגמרא\nZayin\nAlef\nמשנה\nדעת"` — כלומר **בדיוק סדר
 * השליחה**. זה לא מיון עברי ולא מיון ASCII: זה היעדר מיון. מספרי עמודים לא
 * הופיעו כלל, וגם לא אות מפרידה, למרות ש-`\h "A"` בקוד השדה.
 *
 * המסקנה המעשית, וההשלכה על נוסח ההודעות: השדות שנכתבים למסמך תקינים
 * לחלוטין, ו-Word עצמו — שהוא זה שמרנדר `INDEX` בפועל — ימיין וימספר אותם
 * כשהמסמך ייפתח בו. מה שאסור הוא להבטיח למשתמש מפתח ממוין **כאן**. לכן
 * הדיאלוג והפקדים אומרים „המפתח נבנה במלואו ב-Word” ולא „הערכים ימוינו”.
 *
 * ## תת-ערך: הקידוד שנבחר, ולמה לא זה שבחוזה
 *
 * `IndexEntryData.subEntry` קיים בחוזה, מתקבל, ומחזיר `success: true` — וכותב
 * למסמך `XE "אבות" \s "יצחק"`. **`\s` אינו מתג של שדה `XE` ב-Word.** מתגי
 * ה-`XE` הם `\b \f \i \r \t \y`; `\s` הוא מתג של שדה ה-`INDEX` (מזהה רצף).
 * כלומר Word יתעלם ממנו, ותת-הערך פשוט ייעלם — בדיוק התבנית של
 * `crossRefs.insert` בגל 3: פעולה שמדווחת הצלחה וכותבת שדה שבור.
 *
 * הקידוד הקנוני של Word לתת-ערך הוא נקודתיים בתוך הטקסט עצמו:
 * `XE "אבות:אברהם"`. נמדד שהמנוע **מקבל אותו ומפרק אותו בחזרה**: נשלח
 * `text: 'אבות:אברהם'`, ה-`instruction` שנכתב הוא `XE "אבות:אברהם"`,
 * ו-`entries.list` החזיר `text: 'אבות'` עם `subEntry: 'אברהם'`. כלומר
 * הצורה הזאת נכונה בשני הצדדים — גם ב-Word וגם במנוע — ולכן היא זו שנשלחת,
 * ו-`subEntry` אינו נשלח לעולם.
 *
 * המחיר: נקודתיים שהמשתמש הקליד בתוך טקסט הערך ייקראו כתת-ערך. זו
 * ההתנהגות של Word עצמו (שם מבריחים נקודתיים ספרותית ב-`\:`), ולכן היא
 * נשמרת ואינה „מתוקנת”.
 *
 * ## מה שנמדד ולכן **אינו** כאן
 *
 * `IndexConfig` מציע עשרה מתגים. אף אחד מהם אינו משנה את מה שעל המסך (נמדד:
 * שלוש-עשרה קריאות `configure` רצופות, וטקסט הבלוק זהה בכולן) — כולם קוד
 * שדה בלבד, שמשמעותו תתגלה ב-Word. מתוכם נחשפים כאן שניים בלבד, אלה שנמדדו
 * גם משפיעים על ה-`instruction` וגם **הפיכים לחלוטין**:
 *
 * - `columns` → `\c N`. נמדד: `2` → `\c 2`, `3` → `\c 3`, `1` → `\c 1`.
 * - `runIn` → `\r`. נמדד: `true` → `\r`, ו-`false` **מסיר אותו** מה-instruction.
 *
 * ואלה שנשארו בחוץ, כל אחד והסיבה:
 *
 * - `columns` בולע בשקט כל ערך שאינו מספר שלם חיובי: `0`, `-5` ו-`2.5` חזרו
 *   `success: true` ופשוט לא הופיעו ב-`instruction`. לכן הבדיקה כאן ולא במנוע.
 * - `letterRange` (`\p`) מקבל **כל** מחרוזת: נשלח `{from:'zigzag', to:'9'}`
 *   ונכתב `\p "zigzag-9"` עם `success: true`. פקד שכותב טווח אותיות שאין לו
 *   משמעות הוא פקד שמייצר מסמך שבור בשקט.
 * - `accentedSorting` (`\a`) הוא מיון אותיות מוטעמות — סימני ניקוד לטיניים.
 *   אין לו משמעות בעברית, כמו `format.smallCaps`, ולכן הוא מדולג (ראו הבריף).
 * - `headingSeparator` (`\h`), `entryPageSeparator` (`\e`),
 *   `pageRangeSeparator` (`\g`), `sequenceId` (`\s`), `entryTypeFilter` (`\f`)
 *   ו-`pageRangeBookmark` — כולם נכנסים ל-`instruction` כמצופה, אבל אין להם
 *   פקד בדיאלוג „מפתח” של Word העברי, ופקד שאין לו מקבילה שם הוא פקד
 *   שהמשתמש לא יידע לְמה הוא מכוון.
 * - שדה שאינו בחוזה כלל (`flavour: 'zigzag'`) נבלע בשקט עם `success: true`.
 *   כלומר גם כאן, כמו ב-`tabLeader` בגל 4, ההצלחה אינה עדות.
 *
 * ## ההסרה, ולמה **אין** כאן ניקוי שיירים
 *
 * זו ההפתעה הנעימה, והיא ההפך מתוכן העניינים. במודל של המנוע מפתח הוא
 * **בלוק אחד**: `blocks.list` מציג אותו כפסקה יחידה שה-`textPreview` שלה הוא
 * כל הערכים מופרדים ב-`\n`. נמדד על מפתח בן שישה ערכים: לפני ההסרה שמונה
 * בלוקים, אחרי `index.remove` שבעה — הבלוק היחיד נעלם, `index.list` מדווח 0,
 * ו-`getText` נקי לגמרי. אין פסקאות יתומות, ולכן אין כאן `blocks.deleteRange`
 * ואין תלות ב-`blocks.*` ביכולת. ניקוי שאין לו מה לנקות היה רק דרך למחוק
 * פסקה של המשתמש.
 *
 * שדות ה-`XE` **נשארים** במסמך אחרי ההסרה (נמדד: `entries.list` עדיין 6).
 * זו ההתנהגות הנכונה וגם זו של Word: מוחקים את הטבלה, לא את הסימונים.
 *
 * ## למה „מפתח אחד” בהסרה ובהתאמה אישית
 *
 * שלא כמו בתוכן העניינים, שני מפתחות שנוצרו באותו מסמך מקבלים **כתובות
 * שונות** (נמדד: `07043394` ו-`30CE0D77` עם אותו `instruction` בדיוק). כלומר
 * המלכודת של גל 4 אינה חוזרת, ו„עדכן מפתח” יכול לרוץ על כולם בבטחה.
 *
 * אבל להסרה ולהתאמה אישית זה לא מספיק: `doc.selection` מחזיר `blockId` של
 * פסקה, ואין דרך ציבורית לדעת על איזה מהמפתחות המשתמש התכוון. לכן שתי
 * הפעולות האלה פועלות על המפתח היחיד שבמסמך, ובמסמך שיש בו יותר מאחד הן
 * מסרבות ואומרות למה — אותה החלטה בדיוק כמו ב-toc.ts, מטעם אחר.
 *
 * ההגנה מפני כתובת כפולה נשארת בכל זאת ב„עדכן מפתח”: היא עולה כשורה אחת,
 * והמדידה של גל 4 מוכיחה שהמנוע יודע לחלוק מזהים בין שני עצמים.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';
import { readDocSelection, type SelectionDocumentApi, type SelectionTarget } from './doc-selection';

/** `IndexAddress` — מה ש-`get`/`configure`/`rebuild`/`remove` מקבלים כ-`target`. */
interface IndexAddress {
  kind: 'block';
  nodeType: 'index';
  nodeId: string;
}

/** `Position` של החוזה. עוגן של שדה `XE` נבנה משניים כאלה. */
interface IndexPosition {
  blockId?: string;
  offset?: number;
}

/**
 * `IndexEntryAddress` — הכתובת של שדה `XE` בודד.
 *
 * `anchor` ולא `nodeId`: שלא כמו ערך `TC`, שדה `XE` ממוען במיקום בתוך הפסקה
 * (`{start:{blockId,offset}, end:{…}}`). כלומר הכתובת תקפה לתצלום שממנו
 * נקראה בלבד — ראו `removeIndexEntry`.
 */
interface IndexEntryAddress {
  kind: 'inline';
  nodeType: 'indexEntry';
  anchor: { start: IndexPosition; end: IndexPosition };
}

/** החלק מ-`IndexConfig` שנמדד כמשפיע והפיך. ראו הערת הפתיחה. */
interface IndexConfigPatch {
  columns?: number;
  runIn?: boolean;
}

/** `IndexDomain` בחלק שנצרך כאן. */
interface IndexItem {
  address?: IndexAddress;
  config?: { columns?: number; runIn?: boolean };
  entryCount?: number;
}

/**
 * `IndexEntryDomain` עטוף ב-`DiscoveryItem` — ולכן `id` הוא חלק מהחוזה
 * הציבורי (`DiscoveryItem<TDomain>` ב-types/discovery.d.ts) ולא שדה runtime.
 */
interface MarkedIndexEntry {
  id?: string;
  address?: IndexEntryAddress;
  text?: string;
  subEntry?: string;
}

export interface IndexDocumentApi extends SelectionDocumentApi {
  index?: {
    list?: (query?: {
      limit?: number;
      offset?: number;
    }) => MaybePromise<{ items?: readonly IndexItem[]; total?: number } | undefined>;
    insert?: (input: {
      at: { kind: 'documentEnd' };
      config?: IndexConfigPatch;
    }) => MaybePromise<DocReceipt>;
    configure?: (input: {
      target: IndexAddress;
      patch: IndexConfigPatch;
    }) => MaybePromise<DocReceipt>;
    rebuild?: (input: { target: IndexAddress }) => MaybePromise<DocReceipt>;
    remove?: (input: { target: IndexAddress }) => MaybePromise<DocReceipt>;
    entries?: {
      list?: (query?: {
        limit?: number;
        offset?: number;
      }) => MaybePromise<{ items?: readonly MarkedIndexEntry[]; total?: number } | undefined>;
      insert?: (input: {
        at: unknown;
        entry: { text: string };
      }) => MaybePromise<DocReceipt>;
      remove?: (input: { target: IndexEntryAddress }) => MaybePromise<DocReceipt>;
    };
  };
}

/** מה שנדרש מ-SuperDoc: רק הפאסדה של המסמך. ראו page-setup.ts. */
export interface IndexHost {
  activeEditor?: { doc?: IndexDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type IndexTarget = SuperDoc | IndexHost | null | undefined;

/**
 * הטיית הכשל בעברית תקנית. „המפתח” זכר, „הערך” זכר — הביטוי נשמר שלם ואינו
 * נגזר ממזהה. ראו document-api.ts.
 */
const INSERT_FAILED = 'הוספת המפתח נכשלה';
const REBUILD_FAILED = 'עדכון המפתח נכשל';
const REMOVE_FAILED = 'הסרת המפתח נכשלה';
const CONFIGURE_FAILED = 'שינוי הגדרות המפתח נכשל';
const MARK_FAILED = 'סימון הערך נכשל';
const UNMARK_FAILED = 'ביטול סימון הערך נכשל';
const READ_FAILED = 'קריאת המפתח נכשלה';

/**
 * ההודעות שחוזרות על עצמן. „אין במסמך מפתח” מוחזר ככשל ולא כהצלחה שקטה,
 * מאותו טעם בדיוק כמו בתוכן העניינים: המפתח הוא **עצם אחד גלוי**, ולחיצה על
 * „עדכן מפתח” במסמך שאין בו מפתח היא טעות שהשתיקה עליה משאירה את המשתמש
 * בטוח שמשהו קרה.
 */
const NO_INDEX_DETAIL = 'אין במסמך מפתח';
const AMBIGUOUS_DETAIL = 'יש במסמך יותר ממפתח אחד, ואין דרך לדעת על איזה מהם הפעולה חלה';
/**
 * `index.entries.insert` מקבל `at: TextTarget`, כלומר טווח טקסט — ולא עוגן
 * של פסקה כמו `toc.markEntry`. זו גם ההתנהגות של Word: „סמן ערך” מסמן את
 * הטקסט שנבחר. לכן בחירה ריקה היא כשל מנומק ולא ניחוש על תחילת הפסקה.
 */
const NO_SELECTION_DETAIL = 'יש לסמן במסמך את הטקסט שהערך יצביע אליו';

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

function docOf(host: IndexTarget): IndexDocumentApi | null {
  return (host as IndexHost | null | undefined)?.activeEditor?.doc ?? null;
}

/** גודל העמוד בכל שאיבה, ובלם מפני מנוע שיחזיר `total` שאינו יורד. ראו fields.ts. */
const PAGE_SIZE = 200;
const PAGE_GUARD = 1000;

/**
 * כל הפריטים של פעולת discovery, בשאיבת עמודים עד `total`.
 *
 * `total` ולא `items.length`: שתי הרשימות כאן הן `DiscoveryOutput`, כלומר
 * `items` הוא עמוד תחת `limit`/`offset` (נמדד: `limit:2` על חמישה ערכים החזיר
 * שניים עם `total: 5`). ספר תורני עם מפתח ערכים הוא בדיוק המסמך שבו העמוד
 * הראשון אינו הכול — וזה התרחיש של אוצריא, לא מקרה קצה.
 */
async function collectAll<T>(
  failedAction: string,
  list: (query: { limit: number; offset: number }) => MaybePromise<
    { items?: readonly T[]; total?: number } | undefined
  >,
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

/** מה שמוצג כשטקסט הערך נדחה. */
export const INDEX_ENTRY_HINT = 'יש להקליד את טקסט הערך';

/**
 * טקסט הערך כפי שיישלח, או `null` כשהוא ריק.
 *
 * ה-`trim` אינו קוסמטיקה: נמדד שהמנוע **זורק** על מחרוזת ריקה
 * („requires a non-empty entry.text string”) אבל **מקבל בהצלחה** את המחרוזת
 * `'   '` וכותב למסמך `XE "   "` — ערך בלתי נראה שאי אפשר למצוא ואי אפשר
 * למחוק בלי לדעת שהוא שם.
 */
export function normalizeIndexEntryText(raw: string): string | null {
  const text = raw.trim();
  return text === '' ? null : text;
}

/**
 * טקסט הערך כפי ש-Word מקודד תת-ערך: `ראשי:משני`.
 *
 * ההנמקה המלאה, כולל המדידה שמראה שהמנוע מפרק את זה בחזרה ל-`text`+`subEntry`
 * בעוד ש-`subEntry` הישיר כותב מתג שאינו קיים ב-Word — בהערת הפתיחה.
 */
export function buildIndexEntryText(text: string, subEntry: string): string {
  const sub = subEntry.trim();
  return sub === '' ? text : `${text}:${sub}`;
}

/** התקרה של הדיאלוג „מפתח” ב-Word: 1 עד 4 טורים. */
export const INDEX_COLUMNS_MIN = 1;
export const INDEX_COLUMNS_MAX = 4;

/**
 * מה שמוצג כשמספר הטורים נדחה. הבדיקה כאן ולא במנוע מפני שנמדד שהמנוע
 * **בולע בשקט** כל ערך שאינו מספר שלם חיובי — `0`, `-5` ו-`2.5` חזרו כולם
 * `success: true` ופשוט לא הופיעו ב-`instruction`, כלומר המשתמש היה לוחץ
 * „אישור” ולא היה קורה כלום.
 */
export const INDEX_COLUMNS_HINT = `מספר הטורים הוא מספר שלם בין ${INDEX_COLUMNS_MIN} ל-${INDEX_COLUMNS_MAX}`;

/**
 * ברירת המחדל של מספר הטורים: שניים, בדיוק כמו המפתח שנוצר ב-Word
 * (`INDEX \c "2"`). המנוע כותב `INDEX \h "A"` בלי `\c` בכלל (נמדד), ולכן
 * „המסמך לא הצהיר על טורים” הוא המצב הרגיל ולא מקרה קצה.
 */
export const DEFAULT_INDEX_COLUMNS = 2;

export function isValidIndexColumns(columns: number): boolean {
  return (
    Number.isInteger(columns) && columns >= INDEX_COLUMNS_MIN && columns <= INDEX_COLUMNS_MAX
  );
}

/* ------------------------------------------------------------------ */
/* קריאה                                                               */
/* ------------------------------------------------------------------ */

/** ערך `XE` כפי שהדיאלוג מציג אותו. */
export interface IndexEntrySummary {
  /** `DiscoveryItem.id` — מזהה יציב לתצלום, ומפתח `v-for`. */
  id: string;
  text: string;
  subEntry: string;
  /** הכתובת שנשלחת ל-`entries.remove`. אטומה לממשק בכוונה. */
  address: unknown;
}

/** מה שהממשק צריך לדעת. תצלום ולא מנוי, כמו header-footer.ts. */
export interface IndexState {
  /** מספר המפתחות במסמך. 0 = אין מה לעדכן, >1 = אי-אפשר להסיר. */
  count: number;
  /** מספר הטורים של המפתח הראשון, או `null` כשהמנוע לא הצהיר עליו. */
  columns: number | null;
  /** האם הערכים רצופים (`\r`) ולא כל אחד בשורה. */
  runIn: boolean;
  /** הערכים שסומנו במסמך. */
  entries: readonly IndexEntrySummary[];
}

export function emptyIndexState(): IndexState {
  return { count: 0, columns: null, runIn: false, entries: [] };
}

/**
 * הערכים שסומנו במסמך.
 *
 * מיוצאת בנפרד מ-`readIndexState` מפני שהיא השאלה שנשאלת גם בלי המפתח עצמו:
 * המשתמש מסמן מאות ערכים לפני שהוא מוסיף את המפתח בכלל, והדיאלוג צריך
 * להראות לו אותם גם אז.
 *
 * ערך בלי כתובת מדולג ואינו מוחזר: הצגתו הייתה מייצרת שורה שלחיצה עליה
 * שולחת `undefined` ל-`entries.remove`.
 */
export async function listIndexEntries(host: IndexTarget): Promise<IndexEntrySummary[]> {
  const list = docOf(host)?.index?.entries?.list;
  if (typeof list !== 'function') return [];

  const listed = await collectAll<MarkedIndexEntry>(READ_FAILED, (query) => list(query));

  const entries: IndexEntrySummary[] = [];
  for (const entry of listed.items) {
    const address = entry.address;
    if (!address?.anchor) continue;
    entries.push({
      id: typeof entry.id === 'string' && entry.id !== '' ? entry.id : JSON.stringify(address),
      text: typeof entry.text === 'string' ? entry.text : '',
      subEntry: typeof entry.subEntry === 'string' ? entry.subEntry : '',
      address,
    });
  }
  return entries;
}

/**
 * קוראת את מצב המפתח במסמך. לעולם אינה זורקת: כשל של קריאה מחזיר „אין
 * מפתח”, כלומר ה-tooltip יאמר שאין מה לעדכן — ולא ימציא מספר.
 *
 * ההגדרות נקראות מהמפתח **הראשון** בלבד, מפני שהדיאלוג ממילא פועל רק על
 * מסמך שיש בו מפתח אחד; במסמך עם שניים הוא מסרב לפני שהוא מציג משהו.
 */
export async function readIndexState(host: IndexTarget): Promise<IndexState> {
  const list = docOf(host)?.index?.list;
  if (typeof list !== 'function') return emptyIndexState();

  // כשל של הקריאה מחזיר „אין מפתח” ולא ספירה חלקית: `collectAll` מחזיר את
  // העמודים שהספיק לקרוא גם כשהוא נכשל, וספירה חלקית הייתה מגיעה ל-tooltip
  // כאילו היא המספר האמיתי. „אין” הוא תשובה כנה; „שניים מתוך שמונה” אינה.
  const listed = await collectAll<IndexItem>(READ_FAILED, (query) => list(query));
  if (!listed.ok) return emptyIndexState();

  const first = listed.items[0];
  const columns = first?.config?.columns;

  return {
    count: listed.items.length,
    columns: typeof columns === 'number' ? columns : null,
    runIn: first?.config?.runIn === true,
    entries: await listIndexEntries(host),
  };
}

/**
 * המפתח היחיד שבמסמך, או כשל מנומק.
 *
 * שלושת המצבים נפרדים בכוונה: „אין”, „יש כמה” ו„הקריאה נכשלה” הם שלוש
 * הודעות שונות למשתמש, ואיחוד שלהם לאחת היה שולח אותו לחפש את הבעיה
 * הלא נכונה.
 */
async function soleIndex(
  host: IndexTarget,
  failedAction: string,
): Promise<{ ok: true; address: IndexAddress } | { ok: false; outcome: CommandOutcome }> {
  const list = docOf(host)?.index?.list;
  if (typeof list !== 'function') return { ok: false, outcome: unsupported(failedAction) };

  const listed = await collectAll<IndexItem>(failedAction, (query) => list(query));
  if (!listed.ok) return { ok: false, outcome: listed.outcome };

  if (listed.items.length === 0) {
    return { ok: false, outcome: unavailable(failedAction, NO_INDEX_DETAIL, 'no-index') };
  }
  if (listed.items.length > 1) {
    return { ok: false, outcome: unavailable(failedAction, AMBIGUOUS_DETAIL, 'ambiguous-index') };
  }

  const address = listed.items[0].address;
  if (!address?.nodeId) {
    // מפתח בלי כתובת אינו יעד חוקי, ושליחתו הייתה חריגת `TARGET_NOT_FOUND`.
    return { ok: false, outcome: unavailable(failedAction, NO_INDEX_DETAIL, 'no-index') };
  }
  return { ok: true, address };
}

/* ------------------------------------------------------------------ */
/* סימון ערך (שדה XE)                                                  */
/* ------------------------------------------------------------------ */

/** מה שהדיאלוג שולח. `subEntry` ריק = ערך ראשי בלבד. */
export interface IndexEntryDraft {
  text: string;
  subEntry: string;
}

/**
 * „סמן ערך” — מכניסה שדה `XE` על הטקסט שנבחר במסמך.
 *
 * זו הפעולה שהמשתמש יבצע מאות פעמים בספר, ולכן שני הדברים שנבדקים כאן לפני
 * שנוגעים במנוע הם אלה שנמדדו כמסוכנים: טקסט של רווחים בלבד (מתקבל בהצלחה
 * וכותב ערך בלתי נראה) ותת-ערך (החוזה כותב מתג שאינו של Word). ההנמקה
 * המלאה לשניהם בהערת הפתיחה.
 *
 * ה-`at` הוא ה-`TextTarget` שהמנוע עצמו החזיר, ומועבר **כמו שהוא** ולא נבנה
 * מחדש — אותה החלטה כמו ב-doc-selection.ts, ומאותו טעם.
 */
export async function markIndexEntry(
  host: IndexTarget,
  draft: IndexEntryDraft,
): Promise<CommandOutcome> {
  const text = normalizeIndexEntryText(draft.text);
  if (text === null) {
    return { ok: false, message: `${MARK_FAILED}: ${INDEX_ENTRY_HINT}`, reason: 'invalid-text' };
  }

  const doc = docOf(host);
  if (!doc) return unavailable(MARK_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const insert = doc.index?.entries?.insert;
  if (typeof insert !== 'function') return unsupported(MARK_FAILED);

  const selection = await readDocSelection(host as SelectionTarget);
  if (!selection.target) return unavailable(MARK_FAILED, NO_SELECTION_DETAIL, 'no-selection');

  const marked = await attempt(MARK_FAILED, () =>
    insert({
      at: selection.target,
      entry: { text: buildIndexEntryText(text, draft.subEntry) },
    }),
  );
  if (!marked.ok) return marked.outcome;

  return failureOf(MARK_FAILED, marked.value) ?? { ok: true };
}

/**
 * „בטל סימון ערך” — מוחקת שדה `XE` שסומן קודם.
 *
 * הכתובת מגיעה מ-`listIndexEntries` ואינה נבנית כאן, והיא **מיקומית**:
 * `{start:{blockId,offset}, end:{…}}`. כלומר היא תקפה לתצלום שממנו נקראה
 * בלבד — מחיקת ערך מזיזה את ההיסטים של הערכים שאחריו באותה פסקה, ורשימה
 * ישנה הייתה מצביעה על הערך הלא נכון. לכן ReferencesTab קורא מחדש אחרי כל
 * פעולה, והדיאלוג מנקה את הבחירה ומאמת אותה מול הרשימה שהתקבלה.
 */
export async function removeIndexEntry(host: IndexTarget, address: unknown): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(UNMARK_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const remove = doc.index?.entries?.remove;
  if (typeof remove !== 'function') return unsupported(UNMARK_FAILED);

  const removed = await attempt(UNMARK_FAILED, () =>
    remove({ target: address as IndexEntryAddress }),
  );
  if (!removed.ok) return removed.outcome;

  return failureOf(UNMARK_FAILED, removed.value) ?? { ok: true };
}

/* ------------------------------------------------------------------ */
/* הוספה, עדכון, הסרה                                                  */
/* ------------------------------------------------------------------ */

/** מה שהדיאלוג שולח. שני השדות נמדדו כמשפיעים והפיכים — ראו הערת הפתיחה. */
export interface IndexSettings {
  columns: number;
  runIn: boolean;
}

/**
 * „הוסף מפתח” — מכניסה בלוק `INDEX` בסוף המסמך.
 *
 * `documentEnd` ולא מיקום הסמן: מפתח ערכים הוא נספח, והמקום שלו בספר תורני
 * הוא הסוף. `TocCreateLocation` מציע גם `before`/`after` בלוק, אבל אין דרך
 * ציבורית לדעת שהבלוק שהסמן בו הוא זה שהמשתמש התכוון לו — ומפתח שנשתל
 * באמצע פרק הוא נזק שקשה לבטל.
 *
 * ההגדרות נשלחות **ביצירה** ולא ב-`configure` שאחריה: נמדד ש-`insert`
 * מקבל `config` וכותב אותו מיד ל-`instruction` (`INDEX \h "A" \c 2 \r`),
 * וקריאה אחת עדיפה על שתיים שהשנייה בהן עלולה להיכשל ולהשאיר מפתח שאינו
 * במצב שהמשתמש ביקש.
 */
export async function insertIndex(
  host: IndexTarget,
  settings: IndexSettings,
): Promise<CommandOutcome> {
  if (!isValidIndexColumns(settings.columns)) {
    return {
      ok: false,
      message: `${INSERT_FAILED}: ${INDEX_COLUMNS_HINT}`,
      reason: 'invalid-columns',
    };
  }

  const doc = docOf(host);
  if (!doc) return unavailable(INSERT_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const insert = doc.index?.insert;
  if (typeof insert !== 'function') return unsupported(INSERT_FAILED);

  const inserted = await attempt(INSERT_FAILED, () =>
    insert({
      at: { kind: 'documentEnd' },
      config: { columns: settings.columns, runIn: settings.runIn },
    }),
  );
  if (!inserted.ok) return inserted.outcome;

  return failureOf(INSERT_FAILED, inserted.value) ?? { ok: true };
}

/**
 * מה שמוחזר כשהעדכון רץ אך לא על כל המפתחות. „לא הושלם” ולא „נכשל”: משהו
 * כן עודכן, והודעה שאומרת „נכשל” הייתה שולחת את המשתמש לבדוק מפתח שדווקא
 * נבנה מחדש. אותה הבחנה בדיוק כמו ב-toc.ts.
 */
function partialRebuild(unreachable: number): CommandOutcome {
  const left = unreachable === 1 ? 'אחד מהם לא עודכן' : `${unreachable} מהם לא עודכנו`;
  return {
    ok: false,
    message: `עדכון המפתח לא הושלם: יש במסמך כמה מפתחות שאינם ניתנים להבחנה זה מזה, ולכן ${left}`,
    reason: 'ambiguous-index',
  };
}

/**
 * „עדכן מפתח” — בונה מחדש כל מפתח שאפשר לפנות אליו במסמך.
 *
 * רץ על כולם ולא על היחיד: שלא כמו בהסרה, בנייה מחדש אינה הרסנית ואינה
 * דו-משמעית — כל מפתח נבנה מאותם שדות `XE` — ונמדד ששני מפתחות מקבלים
 * כתובות **שונות**, כלומר אפשר באמת לפנות לשניהם.
 *
 * הבלם על כתובת חוזרת נשאר בכל זאת: הוא עולה שורה, וגל 4 מדד מנוע שנותן
 * לשני עצמים את אותו `nodeId`. לולאה תמימה במצב כזה הייתה בונה את הראשון
 * פעמיים ומדווחת „בוצע” על שני שנשאר מיושן.
 *
 * עצירה בכשל הראשון: מסמך שחציו עודכן וחציו לא הוא מצב שאי אפשר לתאר
 * למשתמש. אותה החלטה כמו `rebuildAllFields`.
 */
export async function rebuildIndex(host: IndexTarget): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(REBUILD_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const list = doc.index?.list;
  const rebuild = doc.index?.rebuild;
  if (typeof list !== 'function' || typeof rebuild !== 'function') {
    return unsupported(REBUILD_FAILED);
  }

  const listed = await collectAll<IndexItem>(REBUILD_FAILED, (query) => list(query));
  if (!listed.ok) return listed.outcome;
  if (listed.items.length === 0) {
    return unavailable(REBUILD_FAILED, NO_INDEX_DETAIL, 'no-index');
  }

  const sent = new Set<string>();
  let unreachable = 0;

  for (const item of listed.items) {
    const address = item.address;
    if (!address?.nodeId) continue;
    if (sent.has(address.nodeId)) {
      unreachable++;
      continue;
    }
    sent.add(address.nodeId);

    const rebuilt = await attempt(REBUILD_FAILED, () => rebuild({ target: address }));
    if (!rebuilt.ok) return rebuilt.outcome;
    const failure = failureOf(REBUILD_FAILED, rebuilt.value);
    if (failure) return failure;
  }

  if (unreachable > 0) return partialRebuild(unreachable);

  return { ok: true };
}

/**
 * „הסר מפתח” — מוחקת את בלוק ה-`INDEX`.
 *
 * צעד אחד, בלי ניקוי שיירים: המפתח הוא בלוק **יחיד** במודל של המנוע, ונמדד
 * שאחרי `remove` לא נשארת ממנו ולו פסקה אחת. ההנמקה המלאה, כולל ההשוואה
 * לתוכן העניינים שכן משאיר שיירים, בהערת הפתיחה.
 *
 * שדות ה-`XE` נשארים במסמך, וזו הכוונה: הסימונים הם עבודת המשתמש, והמפתח
 * הוא רק התצוגה שלהם.
 */
export async function removeIndex(host: IndexTarget): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(REMOVE_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const remove = doc.index?.remove;
  if (typeof remove !== 'function') return unsupported(REMOVE_FAILED);

  const sole = await soleIndex(host, REMOVE_FAILED);
  if (!sole.ok) return sole.outcome;

  const removed = await attempt(REMOVE_FAILED, () => remove({ target: sole.address }));
  if (!removed.ok) return removed.outcome;

  return failureOf(REMOVE_FAILED, removed.value) ?? { ok: true };
}

/**
 * „מפתח מותאם אישית” — מספר הטורים והאם הערכים רצופים.
 *
 * שני השדות היחידים שנמדדו גם משפיעים על ה-`instruction` וגם הפיכים
 * לחלוטין, והם גם השניים שיש להם פקד בדיאלוג „מפתח” של Word העברי. ראו
 * הערת הפתיחה לרשימת מה שנשאר בחוץ ולמה.
 *
 * שינוי ההגדרות אינו בונה את המפתח מחדש מהשדות — הוא כותב את המתגים בלבד
 * (נמדד: טקסט הבלוק זהה תו בתו אחרי כל אחת מ-13 קריאות `configure`). לכן
 * הפקד ברצועה מריץ אחריו „עדכן מפתח”, וזו החלטה של הממשק ולא של המודול.
 */
export async function configureIndex(
  host: IndexTarget,
  settings: IndexSettings,
): Promise<CommandOutcome> {
  // נדחה לפני שנוגעים במנוע: נמדד שהוא בולע בשקט כל ערך שאינו מספר שלם חיובי.
  if (!isValidIndexColumns(settings.columns)) {
    return {
      ok: false,
      message: `${CONFIGURE_FAILED}: ${INDEX_COLUMNS_HINT}`,
      reason: 'invalid-columns',
    };
  }

  const doc = docOf(host);
  if (!doc) return unavailable(CONFIGURE_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const configure = doc.index?.configure;
  if (typeof configure !== 'function') return unsupported(CONFIGURE_FAILED);

  const sole = await soleIndex(host, CONFIGURE_FAILED);
  if (!sole.ok) return sole.outcome;

  const applied = await attempt(CONFIGURE_FAILED, () =>
    configure({
      target: sole.address,
      patch: { columns: settings.columns, runIn: settings.runIn },
    }),
  );
  if (!applied.ok) return applied.outcome;

  return failureOf(CONFIGURE_FAILED, applied.value) ?? { ok: true };
}
