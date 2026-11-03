/**
 * „תוכן עניינים” — הקבוצה הראשונה בלשונית „הפניות” של Word העברי, דרך
 * `doc.toc`.
 *
 * מה שכבר היה כאן הוא פקד יחיד מעל פקודת ה-registry `table-of-contents-insert`
 * (שרצה על `create.tableOfContents`). המודול הזה מוסיף את מה שהופך אותו
 * לקבוצה: עדכון, הסרה, התאמה אישית וסימון ערך ידני. הפקד הקיים לא נגע.
 *
 * ומה שנוסף אחרי דיווח משתמש: `layOutTocRows` — מוביל הנקודות, מספר העמוד
 * בקצה השורה, שני דברים שהמנוע אינו נותן בשום שכבה מפני
 * שהסגנונות `TOC1`…`TOC9` פשוט אינם קיימים. ההנמקה המלאה שם.
 *
 * ## הכול כאן נמדד בדפדפן. אלה התוצאות
 *
 * Chrome headless, `file://`, ה-dist האמיתי, מסמך שנוקה ב-`clearContent`
 * וזרוע בכותרות `Heading1`/`Heading2`, ביום 25.8.2026:
 *
 * 1. **`create.tableOfContents` כותב קוד Word תקני** — `TOC \h` — והטבלה
 *    נבנית **מלאה כבר ביצירה**: `toc.get().text` החזיר
 *    `"פרק ראשון1סעיף א1פרק שני1"`, ו-`entryCount` היה 3.
 * 2. **`toc.update` באמת ממלא.** זו הייתה השאלה שהגל נשען עליה. כותרת
 *    (`פרק שלישי`) נוספה למסמך **אחרי** יצירת הטבלה; לפני העדכון
 *    `toc.get().text` היה `"פרק ראשון1סעיף א1פרק שני1"`, ואחרי
 *    `update({mode:'all'})` הוא `"פרק ראשון1סעיף א1פרק שני1פרק שלישי1"`,
 *    ו-`getText()` של המסמך כולו קיבל שורה נוספת `"פרק שלישי\t1"`.
 *    **`entryCount` אינו העדות לכך.** נמדד שוב: הוא עלה מ-3 ל-4 כבר עם
 *    הוספת הכותרת, **לפני** ה-`update` ובלי קשר אליו — הוא נספר מהמקורות
 *    (הכותרות ושדות ה-`TC`) ולא מהשורות שרונדרו. מה שהעדכון שינה הוא
 *    ה-`text`, כלומר הטבלה שעל המסך.
 * 3. **`toc.markEntry` מייצר שדה `TC` תקני של Word** — `TC "ערך ידני" \l 1` —
 *    והערך נאסף אל הטבלה בעדכון הבא (הטקסט קיבל `"ערך ידני1"` בסופו).
 *
 * ## מה שנמדד ולכן **אינו** כאן
 *
 * `TocConfigurePatch` מציע עשר מתגים. שלושה מהם מתקבלים עם `success: true`
 * ואינם עושים כלום, ולכן אין להם פקד:
 *
 * - `tabLeader` (מנהיג הנקודות) — נשלח `'dot'` ואחריו `'none'`; אף אחד מהם
 *   לא הופיע ב-`displayConfig` שחזר, וה-`instruction` לא השתנה כלל. גם ערך
 *   שאינו בחוזה בכלל (`'zigzag'`) חזר `success: true`. **המוביל עצמו כן
 *   מושג** — לא דרך המתג הזה אלא בעצירת טאב על השורות, ראו
 *   `layOutTocRows`.
 * - `rightAlignPageNumbers` — אותה תוצאה בדיוק.
 * - `includePageNumbers` — זו „הקרנה” של המתג `\n` לקריאה בלבד, ולא מתג
 *   בפני עצמו: שליחתה נבלעת בשני הכיוונים.
 *
 * וחמור מזה: **הצגת מספרי העמודים היא מתג חד-כיווני.** אחרי
 * `omitPageNumberLevels: {from:1,to:9}` (שנכתב כ-`\n "1-9"` ובאמת מוחק את
 * המספרים מהטבלה) אין בחוזה שום ערך שמחזיר אותם — נוסו `includePageNumbers:
 * true`, `{from:0,to:0}`, `{from:10,to:10}`, `null` ו-`undefined`; כולם
 * חזרו `success: true`, ואף אחד מהם לא הסיר את `\n` מה-`instruction`.
 * תיבת סימון „הצג מספרי עמודים” הייתה לכן מלכודת: מכבים אותה פעם אחת ואי
 * אפשר להדליק בחזרה. לכן היא אינה כאן, ו-`configureTableOfContents` חושפת
 * בדיוק שני דברים שנמדדו כהפיכים לחלוטין — טווח רמות הכותרות (`\o`) והאם
 * הערכים הם קישורים (`\h`).
 *
 * ## הסרה ב-SuperDoc 2.23
 *
 * `toc.remove` מוחק את השדה כולו, כולל תוצאתו השמורה. אין לנקות אחריו
 * פסקאות לפי סגנון: גם פסקה רגילה יכולה לשאת TOC1, ו-`entryCount` נספר
 * מהכותרות החיות ולא מהשורות שנשמרו בעדכון האחרון. טבלה מיושנת עלולה
 * להכיל פחות שורות מכפי שהספירה מציעה, ומעקף כזה מוחק תוכן של המשתמש.
 *
 * ## למה „טבלה אחת” — כולל בעדכון
 *
 * להסרה ולהתאמה אישית אין דרך ציבורית לדעת על איזו טבלה המשתמש התכוון:
 * `doc.selection` מחזיר `blockId` של פסקה, ושורות הטבלה מלבד הראשונה הן
 * פסקאות שאינן נושאות שום זיהוי של הטבלה שהן שייכות לה. לכן שתי הפעולות
 * האלה פועלות על הטבלה היחידה שבמסמך, ובמסמך שיש בו יותר מאחת הן מסרבות
 * ואומרות למה.
 *
 * כאן נכתב קודם ש„עדכן טבלה” אינו דו-משמעי ולכן הוא רץ על כולן. **המדידה
 * סותרת.** שתי טבלאות שנוצרו באותו מסמך מקבלות את **אותו `nodeId` בדיוק**:
 * שתי הקריאות ל-`create.tableOfContents` החזירו `5D013B61`, `toc.list` דיווח
 * `total: 2` ושני items עם אותה כתובת, וגם `blocks.list` מציג את שני הבלוקים
 * (אורדינל 0 ואורדינל 2) תחת אותו מזהה. המזהה הוא hash של ה-`instruction`,
 * ושתי טבלאות שנוצרו כברירת מחדל נושאות `TOC \h` זהה. לולאה על ה-items
 * הייתה לכן בונה את הראשונה פעמיים, משאירה את השנייה מיושנת, ומדווחת „בוצע”.
 *
 * דרך ציבורית לפנות לטבלה השנייה **אין**, ונבדקה גם האפשרות המתבקשת: המזהה
 * שב-`blocks.list` הוא אותו מזהה עצמו. לכן העדכון רץ פעם אחת לכל כתובת
 * **שונה**, ואם היו כפילויות הוא אומר במפורש שהעדכון לא הושלם. בניגוד להסרה
 * הוא אינו מסרב מראש: בנייה מחדש אינה הרסנית, וטבלה אחת מעודכנת עדיף על
 * כלום — ובלבד שלא ייאמר למשתמש שהכול עודכן.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';
import { readDocSelection, type SelectionDocumentApi, type SelectionTarget } from './doc-selection';
import { readPageMargins, type PageSetupTarget } from './page-setup';

/** `TocAddress` — מה ש-`get`/`configure`/`update`/`remove` מקבלים כ-`target`. */
interface TocAddress {
  kind: 'block';
  nodeType: 'tableOfContents';
  nodeId: string;
}

/** `TocEntryAddress` — הכתובת של שדה `TC` בודד. */
interface TocEntryAddress {
  kind: 'inline';
  nodeType: 'tableOfContentsEntry';
  nodeId: string;
}

/** `TocConfigurePatch` בחלק שנמדד כמשפיע והפיך. ראו הערת הפתיחה. */
interface TocConfigurePatch {
  outlineLevels?: { from: number; to: number };
  hyperlinks?: boolean;
}

/**
 * `TocDomain` בחלק שנצרך כאן.
 *
 * `preserved` הוא חלק מהחוזה הציבורי (`TocPreservedSwitches`) ולא שדה runtime,
 * ונמדד שהוא באמת מגיע: טבלה שנוצרה עם `TOC \o "1-3" \h \t "MyToc1,1"` חזרה
 * מ-`toc.list` עם `preserved.customStyles: [{styleName:"MyToc1", level:1}]`,
 * גם אחרי `update`. זה המתג `\t` של Word — הוא ממפה סגנון פסקה משלו לרמה
 * בטבלה, ובמסמך שהגיע מ-Word שורות הטבלה נושאות אותו ולא `TOC1`…`TOC9`.
 * בלעדיו ניקוי השיירים היה עיוור בדיוק במסמכים שהמתג הזה קיים בהם.
 */
interface TocEntry {
  address?: TocAddress;
  sourceConfig?: { outlineLevels?: { from?: number; to?: number } };
  displayConfig?: { hyperlinks?: boolean };
  preserved?: { customStyles?: readonly { styleName?: string; level?: number }[] };
  entryCount?: number;
}

/** `TocEntryDomain` — ערך `TC` כפי ש-`listEntries` מחזיר אותו. */
interface MarkedEntry {
  address?: TocEntryAddress;
  text?: string;
  level?: number;
}

/** `BlockListEntry` בחלק שנצרך לניקוי השיירים. */
interface BlockEntry {
  ordinal?: number;
  nodeId?: string;
  nodeType?: string;
  styleId?: string | null;
}

export interface TocDocumentApi extends SelectionDocumentApi {
  toc?: {
    list?: (query?: {
      limit?: number;
      offset?: number;
    }) => MaybePromise<{ items?: readonly TocEntry[]; total?: number } | undefined>;
    update?: (input: {
      target: TocAddress;
      mode?: 'all' | 'pageNumbers';
    }) => MaybePromise<DocReceipt>;
    remove?: (input: { target: TocAddress }) => MaybePromise<DocReceipt>;
    configure?: (input: {
      target: TocAddress;
      patch: TocConfigurePatch;
    }) => MaybePromise<DocReceipt>;
    markEntry?: (input: {
      target: {
        kind: 'inline-insert';
        anchor: { nodeType: 'paragraph'; nodeId: string };
        position?: 'start' | 'end';
      };
      text: string;
      level?: number;
    }) => MaybePromise<DocReceipt>;
    unmarkEntry?: (input: { target: TocEntryAddress }) => MaybePromise<DocReceipt>;
    listEntries?: (query?: {
      limit?: number;
      offset?: number;
    }) => MaybePromise<{ items?: readonly MarkedEntry[]; total?: number } | undefined>;
  };
  format?: {
    paragraph?: {
      // `nodeType` הוא `'paragraph'` גם לבלוק ה-`tableOfContents` עצמו: נמדד
      // שהוא מתקבל תחתיו ומקבל את העצירה, בעוד `'tableOfContents'` אינו סוג
      // חוקי בחוזה של `format.paragraph.*`.
      setTabStop?: (input: {
        target: { kind: 'block'; nodeType: 'paragraph' | 'listItem'; nodeId: string };
        position: number;
        alignment: 'right';
        leader: 'dot';
      }) => MaybePromise<DocReceipt>;
      clearAllTabStops?: (input: {
        target: { kind: 'block'; nodeType: 'paragraph' | 'listItem'; nodeId: string };
      }) => MaybePromise<DocReceipt>;
    };
  };
  create?: {
    paragraph?: (input: {
      at: { kind: 'after'; target: { kind: 'block'; nodeType: string; nodeId: string } };
    }) => MaybePromise<DocReceipt>;
  };
  blocks?: {
    list?: (input?: {
      limit?: number;
      offset?: number;
    }) => MaybePromise<{ blocks?: readonly BlockEntry[]; total?: number } | undefined>;
  };
}

/** מה שנדרש מ-SuperDoc: רק הפאסדה של המסמך. ראו page-setup.ts. */
export interface TocHost {
  activeEditor?: { doc?: TocDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type TocTarget = SuperDoc | TocHost | null | undefined;

/**
 * הטיית הכשל בעברית תקנית. „תוכן העניינים” זכר, „הטבלה” נקבה — הביטוי נשמר
 * שלם ואינו נגזר ממזהה. ראו document-api.ts.
 */
const UPDATE_FAILED = 'עדכון תוכן העניינים נכשל';
const REMOVE_FAILED = 'הסרת תוכן העניינים נכשלה';
const CONFIGURE_FAILED = 'שינוי הגדרות תוכן העניינים נכשל';
const MARK_FAILED = 'סימון הערך נכשל';
const UNMARK_FAILED = 'ביטול סימון הערך נכשל';
const READ_FAILED = 'קריאת תוכן העניינים נכשלה';

/**
 * שלוש ההודעות שחוזרות על עצמן בשלוש פעולות. „אין במסמך תוכן עניינים” מוחזר
 * ככשל ולא כהצלחה שקטה — בשונה מ„עדכן שדות” ו„עדכן הפניות”, שהם רענון
 * גורף על אוכלוסייה בלתי נראית, תוכן עניינים הוא **עצם אחד גלוי**. לחיצה על
 * „עדכן טבלה” במסמך שאין בו טבלה היא טעות, והשתיקה עליה משאירה את המשתמש
 * בטוח שמשהו קרה.
 */
const NO_TOC_DETAIL = 'אין במסמך תוכן עניינים';
const AMBIGUOUS_DETAIL =
  'יש במסמך יותר מתוכן עניינים אחד, ואין דרך לדעת על איזה מהם הפעולה חלה';
const NO_ANCHOR_DETAIL = 'יש ללחוץ בגוף המסמך, על הפסקה שאליה יסומן הערך';

/**
 * מה שמוחזר כשהעדכון רץ אך לא על כל הטבלאות. „לא הושלם” ולא „נכשל”: משהו
 * כן עודכן, והודעה שאומרת „נכשל” הייתה שולחת את המשתמש לבדוק טבלה שדווקא
 * נבנתה מחדש.
 */
function partialUpdate(unreachable: number): CommandOutcome {
  const left = unreachable === 1 ? 'אחת מהן לא עודכנה' : `${unreachable} מהן לא עודכנו`;
  return {
    ok: false,
    message: `עדכון תוכן העניינים לא הושלם: יש במסמך כמה טבלאות תוכן עניינים שאינן ניתנות להבחנה זו מזו, ולכן ${left}`,
    reason: 'ambiguous-toc',
  };
}

/** התקרה של Word למתג `\l` בשדה `TC` ולמתג `\o` בשדה `TOC`. */
export const TOC_LEVEL_MIN = 1;
export const TOC_LEVEL_MAX = 9;

/**
 * מה שמוצג למשתמש כשהרמה נדחית. הבדיקה יושבת כאן ולא נשענת על המנוע מפני
 * שנמדד שהמנוע **מקבל** `level: 12` ומחזיר `success: true` — וכותב למסמך
 * `TC "…" \l 12`, מתג שאינו חוקי ב-Word.
 */
export const TOC_LEVEL_HINT = 'רמת הערך היא מספר שלם בין 1 ל-9';

/** מה שמוצג כשטווח הרמות נדחה. גם כאן המנוע מקבל הכול — ראו `normalizeTocLevels`. */
export const TOC_LEVELS_HINT = 'טווח הרמות מתחיל ברמה 1 לכל הפחות, מסתיים ברמה 9 לכל היותר, והרמה הראשונה אינה גדולה מהאחרונה';

/** טקסט הערך כפי שיישלח, או `null` כשהוא ריק. המנוע **זורק** על מחרוזת ריקה. */
export function normalizeTocEntryText(raw: string): string | null {
  const text = raw.trim();
  return text === '' ? null : text;
}

/** האם הרמה חוקית ל-Word. `Number.isInteger` ולא `>=`: `1.5` נשלח כמו שהוא. */
export function isValidTocLevel(level: number): boolean {
  return Number.isInteger(level) && level >= TOC_LEVEL_MIN && level <= TOC_LEVEL_MAX;
}

/**
 * טווח הרמות כפי שיישלח, או `null` כשהוא פסול.
 *
 * הבדיקה כאן ולא במנוע: נמדד ש-`configure` מקבל `{from:9,to:1}` בשמחה,
 * כותב `TOC \o "9-1"`, ומחזיר `success: true` — והטבלה נעשית **ריקה**.
 * כלומר המשתמש היה לוחץ „אישור” ורואה את תוכן העניינים שלו נמחק.
 */
export function normalizeTocLevels(from: number, to: number): { from: number; to: number } | null {
  if (!isValidTocLevel(from) || !isValidTocLevel(to) || from > to) return null;
  return { from, to };
}

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

function docOf(host: TocTarget): TocDocumentApi | null {
  return (host as TocHost | null | undefined)?.activeEditor?.doc ?? null;
}

/** גודל העמוד בכל שאיבה, ובלם מפני מנוע שיחזיר `total` שאינו יורד. ראו fields.ts. */
const PAGE_SIZE = 200;
const PAGE_GUARD = 1000;

/**
 * כל הפריטים של פעולת discovery, בשאיבת עמודים עד `total`.
 *
 * `total` ולא `items.length`: `TocListResult` ו-`TocListEntriesResult` הם
 * `DiscoveryOutput`, כלומר `items` הוא עמוד תחת `limit`/`offset`. רשימה
 * שמציגה עמוד אחד הייתה משאירה במסמך גדול טבלאות שהעדכון לא נגע בהן.
 *
 * כשל באמצע השאיבה מוחזר כמו שהוא ואינו נבלע: הקורא הוא שמחליט אם חצי
 * רשימה מספיקה לו (`readTocState`) או שהיא כשל (`updateTableOfContents`).
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
/* קריאה                                                               */
/* ------------------------------------------------------------------ */

/** ערך `TC` כפי שהדיאלוג מציג אותו. */
export interface TocEntrySummary {
  nodeId: string;
  text: string;
  level: number;
}

/** מה שהממשק צריך לדעת. תצלום ולא מנוי, כמו header-footer.ts. */
export interface TocState {
  /** מספר טבלאות התוכן במסמך. 0 = אין מה לעדכן, >1 = אי-אפשר להסיר. */
  count: number;
  /** טווח הרמות של הטבלה היחידה, או `null` כשהמנוע לא הצהיר עליו (`TOC` בלי `\o`). */
  levels: { from: number; to: number } | null;
  /** האם הערכים בטבלה הם קישורים (`\h`). */
  hyperlinks: boolean;
  /** הערכים הידניים (`TC`) שבמסמך. */
  entries: readonly TocEntrySummary[];
}

export function emptyTocState(): TocState {
  return { count: 0, levels: null, hyperlinks: false, entries: [] };
}

/**
 * קוראת את מצב תוכן העניינים במסמך. לעולם אינה זורקת: כשל של קריאה מחזיר
 * „אין תוכן עניינים”, כלומר ה-tooltip יאמר שאין מה לעדכן — ולא ימציא מספר.
 *
 * ההגדרות נקראות מהטבלה **הראשונה** בלבד, מפני שהדיאלוג ממילא פועל רק על
 * מסמך שיש בו טבלה אחת (ראו הערת הפתיחה); במסמך עם שתיים הוא מסרב לפני
 * שהוא מציג משהו, ולכן אין מצב שבו הערכים המוצגים שייכים לטבלה אחרת מזו
 * שתשונה.
 */
export async function readTocState(host: TocTarget): Promise<TocState> {
  const doc = docOf(host);
  const list = doc?.toc?.list;
  if (typeof list !== 'function') return emptyTocState();

  const listed = await collectAll<TocEntry>(READ_FAILED, (query) => list(query));
  const first = listed.items[0];

  const levels = first?.sourceConfig?.outlineLevels;
  const from = levels?.from;
  const to = levels?.to;

  const entries: TocEntrySummary[] = [];
  const listEntries = doc?.toc?.listEntries;
  if (typeof listEntries === 'function') {
    const marked = await collectAll<MarkedEntry>(READ_FAILED, (query) => listEntries(query));
    for (const entry of marked.items) {
      const nodeId = entry.address?.nodeId;
      if (typeof nodeId !== 'string' || nodeId === '') continue;
      entries.push({
        nodeId,
        text: typeof entry.text === 'string' ? entry.text : '',
        level: typeof entry.level === 'number' ? entry.level : TOC_LEVEL_MIN,
      });
    }
  }

  return {
    count: listed.items.length,
    levels: typeof from === 'number' && typeof to === 'number' ? { from, to } : null,
    hyperlinks: first?.displayConfig?.hyperlinks === true,
    entries,
  };
}

/**
 * הטבלה היחידה שבמסמך, או כשל מנומק.
 *
 * שלושת המצבים נפרדים בכוונה: „אין”, „יש כמה” ו„הקריאה נכשלה” הם שלוש
 * הודעות שונות למשתמש, ואיחוד שלהם לאחת היה שולח אותו לחפש את הבעיה
 * הלא נכונה.
 */
async function soleToc(
  host: TocTarget,
  failedAction: string,
): Promise<{ ok: true; address: TocAddress; rows: TocRowSignature } | { ok: false; outcome: CommandOutcome }> {
  const list = docOf(host)?.toc?.list;
  if (typeof list !== 'function') return { ok: false, outcome: unsupported(failedAction) };

  const listed = await collectAll<TocEntry>(failedAction, (query) => list(query));
  if (!listed.ok) return { ok: false, outcome: listed.outcome };

  if (listed.items.length === 0) {
    return { ok: false, outcome: unavailable(failedAction, NO_TOC_DETAIL, 'no-toc') };
  }
  if (listed.items.length > 1) {
    return { ok: false, outcome: unavailable(failedAction, AMBIGUOUS_DETAIL, 'ambiguous-toc') };
  }

  const address = listed.items[0].address;
  if (!address?.nodeId) {
    // טבלה בלי כתובת אינה יעד חוקי, ושליחתה הייתה חריגת `INVALID_TARGET`.
    return { ok: false, outcome: unavailable(failedAction, NO_TOC_DETAIL, 'no-toc') };
  }
  return { ok: true, address, rows: rowSignatureOf(listed.items[0]) };
}

/* ------------------------------------------------------------------ */
/* עדכון                                                               */
/* ------------------------------------------------------------------ */

/**
 * „עדכן טבלה” — בונה מחדש כל תוכן עניינים שאפשר לפנות אליו במסמך.
 *
 * `mode: 'all'` ולא `'pageNumbers'`: זה ההבדל בין שתי האפשרויות שהדיאלוג של
 * Word מציע, ו„עדכן את הטבלה כולה” הוא היחיד שבטוח נכון אחרי עריכה. עדכון
 * מספרי העמודים בלבד אינו נמדד כאן כמשמעותי — ב-headless אין עימוד וכל
 * המספרים הם 1 — ופקד שלא נמדד אינו נשלח.
 *
 * עצירה בכשל הראשון: מסמך שחציו עודכן וחציו לא הוא מצב שאי אפשר לתאר
 * למשתמש, ועדיף לעצור ולומר מה נכשל. אותה החלטה כמו `rebuildAllFields`.
 *
 * כתובת חוזרת נשלחת **פעם אחת**, ובסוף מדווחת: שתי טבלאות עם אותו
 * `instruction` חולקות `nodeId` (נמדד — ראו הערת הפתיחה), ולולאה תמימה הייתה
 * בונה את הראשונה פעמיים ומדווחת „בוצע” על שנייה שנשארה מיושנת.
 */
export async function updateTableOfContents(host: TocTarget): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(UPDATE_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const list = doc.toc?.list;
  const update = doc.toc?.update;
  if (typeof list !== 'function' || typeof update !== 'function') {
    return unsupported(UPDATE_FAILED);
  }

  const listed = await collectAll<TocEntry>(UPDATE_FAILED, (query) => list(query));
  if (!listed.ok) return listed.outcome;
  if (listed.items.length === 0) {
    return unavailable(UPDATE_FAILED, NO_TOC_DETAIL, 'no-toc');
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

    const updated = await attempt(UPDATE_FAILED, () => update({ target: address, mode: 'all' }));
    if (!updated.ok) return updated.outcome;
    const failure = failureOf(UPDATE_FAILED, updated.value);
    if (failure) return failure;
  }

  if (unreachable > 0) return partialUpdate(unreachable);

  return { ok: true };
}

/* ------------------------------------------------------------------ */
/* הסרה                                                                */
/* ------------------------------------------------------------------ */

/** סגנונות שורות הטבלה של Word. זה מה שהמנוע עצמו מרנדר — נמדד. */
const TOC_STYLE = /^TOC[1-9]$/;

/**
 * סוגי הבלוקים ששורת טבלה יכולה להופיע בהם.
 *
 * `listItem` ולא רק `paragraph`: הסיווג הוא של המנוע ולא שלנו, ותוכן עניינים
 * ממוספר היה מקבל שורה אחת מסווגת אחרת — הרצף היה נקטע באמצע, החצי השני היה
 * נשאר על המסך, וההסרה הייתה מדווחת „בוצע”. `heading` **אינו** כאן בכוונה:
 * הכותרת שאחרי הטבלה היא בדיוק מה שאסור לבלוע.
 */
const ROW_NODE_TYPES = new Set(['paragraph', 'listItem']);

/** סגנונות נוספים לזיהוי קצה טבלה, רק למעקף פסקת הסיום. */
interface TocRowSignature {
  styleNames: readonly string[];
}

function rowSignatureOf(entry: TocEntry | undefined): TocRowSignature {
  const styleNames = (entry?.preserved?.customStyles ?? [])
    .map((style) => style?.styleName)
    .filter((name): name is string => typeof name === 'string' && name !== '');
  return { styleNames };
}

/**
 * האם הבלוק הוא שורה של תוכן עניינים. `TOC1`…`TOC9` הם סגנונות Word, וסגנונות
 * ה-`\t` של הטבלה שהוסרה מצטרפים אליהם — ראו `TocEntry.preserved`.
 */
function isTocRow(block: BlockEntry | undefined, styleNames: readonly string[]): boolean {
  if (!ROW_NODE_TYPES.has(block?.nodeType ?? '')) return false;
  const styleId = block?.styleId ?? '';
  return TOC_STYLE.test(styleId) || styleNames.includes(styleId);
}

/**
 * ב-superdoc 2.17 `toc.remove` מסרב לטבלה שאין אחריה פסקה (superdoc/docx-editor#4033).
 * מוסיפה פסקה ריקה כשהמסמך נגמר בשורת טבלה; `true` אם נוספה.
 */
async function appendParagraphAfterTocEnd(host: TocTarget, signature: TocRowSignature): Promise<boolean> {
  const doc = docOf(host);
  const list = doc?.blocks?.list;
  const createParagraph = doc?.create?.paragraph;
  if (typeof list !== 'function' || typeof createParagraph !== 'function') return false;

  const head = await attempt(REMOVE_FAILED, () => list({ limit: 1 }));
  const total = head.ok ? head.value?.total : undefined;
  if (typeof total !== 'number' || total < 1) return false;
  const tail = await attempt(REMOVE_FAILED, () => list({ limit: 1, offset: total - 1 }));
  const last = tail.ok ? tail.value?.blocks?.[0] : undefined;
  if (typeof last?.nodeId !== 'string' || typeof last.nodeType !== 'string') return false;
  if (last.nodeType !== 'tableOfContents' && !isTocRow(last, signature.styleNames)) return false;

  const target = { kind: 'block' as const, nodeType: last.nodeType, nodeId: last.nodeId };
  const created = await attempt(REMOVE_FAILED, () => createParagraph({ at: { kind: 'after', target } }));
  return created.ok && created.value?.success !== false;
}

/** „הסר תוכן עניינים” — המנוע מוחק את השדה לפי גבולותיו ב-DOCX. */
export async function removeTableOfContents(host: TocTarget): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(REMOVE_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const remove = doc.toc?.remove;
  if (typeof remove !== 'function') return unsupported(REMOVE_FAILED);

  const sole = await soleToc(host, REMOVE_FAILED);
  if (!sole.ok) return sole.outcome;

  let removed = await attempt(REMOVE_FAILED, () => remove({ target: sole.address }));
  if (
    removed.ok &&
    removed.value?.failure?.code === 'CAPABILITY_UNAVAILABLE' &&
    (await appendParagraphAfterTocEnd(host, sole.rows))
  ) {
    removed = await attempt(REMOVE_FAILED, () => remove({ target: sole.address }));
  }
  if (!removed.ok) return removed.outcome;
  const failure = failureOf(REMOVE_FAILED, removed.value);
  if (failure) return failure;

  return { ok: true };
}

/* ------------------------------------------------------------------ */
/* עיצוב שורות הטבלה — מוביל ומספר עמוד                                 */
/* ------------------------------------------------------------------ */

/**
 * „כותרת……5” — מה שהופך רשימת שורות לתוכן עניינים.
 *
 * ## מה שהיה על המסך, ולמה
 *
 * המנוע כותב את השורה עם תו טאב יחיד בין הכותרת למספר:
 * `<w:r><w:tab/></w:r>` ואחריו `PAGEREF`. תו הטאב זקוק ל**עצירה** שתגיד לאן
 * הוא נדחף ואם מציירים משהו בדרך, ונמדד (scripts/qa/toc-direction-probe.mjs)
 * שאין אף אחת בשום שכבה:
 *
 * - ב-pPr של שורות ה-TOC אין `w:tabs` כלל;
 * - הסגנונות `TOC1`…`TOC9` — שם Word מחזיק את העצירה הזאת — **אינם מוגדרים
 *   ב-`word/styles.xml`** של התבנית (33 סגנונות, ואף אחד מהם אינו `TOC*`),
 *   ולמנוע אין API לחבר סגנון חדש;
 * - ולכן התו נופל על `w:defaultTabStop` — קפיצה אחת של חצי אינץ' ובלי מוביל.
 *
 * זה מה שנראה: הכותרת מימין (הכיוון נורש מ-`w:bidi` של המקטע), ומספר העמוד
 * 38 פיקסלים ממנה, בלי נקודות ובלי קשר לרוחב העמוד. בשורה ברוחב 602 פיקסלים
 * מספר העמוד יושב ב-x=593 — כלומר צמוד לכותרת ולא בקצה השורה.
 *
 * ## מה שנעשה, ולמה **בכל שורה** ולא במקום אחד
 *
 * `format.paragraph.setTabStop` עם `alignment: 'right'` ו-`leader: 'dot'`.
 * זו העצירה שהסגנון היה נותן, ומכיוון שאין דרך לכתוב סגנון היא נכתבת לכל
 * שורה בנפרד. נמדד על ה-dist הארוז: העצירה נכתבת קנונית
 * (`<w:tab w:val="right" w:pos="…" w:leader="dot"/>`), הצייר מצייר את המוביל
 * כ-`div.superdoc-leader` ברוחב 513 פיקסלים, ומספר העמוד עובר לקצה השורה.
 * העלות נמדדה גם היא: 26ms לשורה בממוצע (10 שורות ב-261ms).
 *
 * העצירה **שורדת `toc.update`**: אחרי בנייה מחדש ה-`w:tabs` נשאר על כל
 * השורות, וגם המוביל נשאר מצויר. כלומר זו אינה עבודה שמתאדה בלחיצה הבאה —
 * והיא כן נדרשת שוב אחרי עדכון שהוסיף כותרת, מפני שהשורה **החדשה** נולדת
 * בלי עצירה.
 *
 * ## למה המיקום אינו רוחב אזור הטקסט, אלא פחות ממנו
 *
 * כאן יש באג של המנוע, והמספר הזה הוא הקיזוז שלו. בפסקה עברית עצירה
 * ב„סוף” אמורה להצמיד את הטקסט שאחרי הטאב אל **הקצה השמאלי מבפנים**;
 * המנוע מציב את הקצה **הימני** של המספר על העצירה, כלומר הוא חורג משמאל
 * ברוחב עצמו. נמדד, A4 עם שוליים של אינץ' (אזור טקסט 9026 twips = 602px
 * שמתחיל ב-x=96 בעמוד):
 *
 *     עצירה            מספר העמוד על המסך
 *     right@9026        x=87  — 9 פיקסלים **מחוץ** לאזור הטקסט
 *     left@9026         x=87  — אותו מקום בדיוק
 *     right@8826        x=100 — בתוך אזור הטקסט
 *
 * (הדבר מדווח ב-docs/engine-gaps.md.) לכן העצירה נסוגה פנימה בשיעור שמכסה
 * את המספר הרחב ביותר: `TOC_PAGE_NUMBER_RESERVE_TWIPS`. הבחירה בקיזוז ולא
 * בהצמדה לרוחב המלא היא בחירה בין שני מצבים גלויים — מספר שיוצא אל השוליים,
 * או עמודת מספרים מיושרת שמתחילה סנטימטר מהשוליים — והשני הוא מה ש-Word
 * נראה כמו.
 *
 * ## מה שמנוקה לפני הכתיבה, ולמה
 *
 * `setTabStop` **מוסיף** עצירה ואינו מחליף אותה (נמדד: שתי קריאות השאירו
 * שתיים), ואי אפשר לקרוא את העצירות הקיימות — `doc.get()` אינו מחזיר `tabs`
 * בתכונות הפסקה (ראו page-ruler.ts). לכן `clearAllTabStops` לפני: מסמך
 * שגודל העמוד שלו שונה אחרי שהוכנס תוכן עניינים היה מקבל עצירה שנייה במיקום
 * חדש, ותו הטאב היה נדחף אל **הראשונה** שהוא פוגש — כלומר אל המיקום הישן.
 *
 * ## למה אין הזחה ישירה לפי רמה
 *
 * המנוע מודד עצירת טאב אחרי ההזחה, בעוד Word מודד `w:pos` משולי העמוד.
 * פיצוי שמחזיר את העמודה למקומה בעורך שומר בקובץ מיקום אחר לכל רמה, ולכן
 * פותח את ה-DOCX ב-Word עם עמודת מספרים שבורה. עד שיהיה תיקון במנוע או API
 * לסגנון `TOC*`, העדיפות היא למסמך תקני: אותה עצירה לכל רמה, בלי הזחה ישירה.
 * כך המוביל ומספר העמוד נכונים בשני העורכים.
 */
export const TOC_PAGE_NUMBER_RESERVE_TWIPS = 540;

/** מה שמוצג כשעיצוב השורות נכשל. הטבלה עצמה כבר במסמך — ראו `layOutTocRows`. */
const LAYOUT_FAILED = 'עיצוב שורות תוכן העניינים נכשל';

/**
 * שורות שכבר עוצבו במופע המנוע הנוכחי. `toc.update` משאיר את העצירות של
 * השורות הקיימות; לכן כתיבה חוזרת היא עבודה מיותרת ומצטברת במיוחד בספרים.
 * המפתח כולל סוג בלוק כי אותו `nodeId` אינו בהכרח כתובת זהה תחת שני סוגים.
 */
const tocLayoutCache = new WeakMap<TocDocumentApi, Map<string, number>>();

function tocLayoutKey(row: { nodeId: string; nodeType: 'paragraph' | 'listItem' }): string {
  return `${row.nodeType}:${row.nodeId}`;
}

/**
 * האם הבלוק הוא שורה של תוכן עניינים **קיים**.
 *
 * בשונה מ-`isTocRow`, שמזהה רק פסקאות, כאן הבלוק הראשון
 * של הטבלה עדיין במקומו — והוא השורה הראשונה שעל המסך. הוצאתו הייתה משאירה
 * את הכותרת הראשונה בלי מוביל, וזה בדיוק הפער הגלוי ביותר.
 */
function isTocLine(block: BlockEntry | undefined, levels: ReadonlyMap<string, number>): boolean {
  const styleId = block?.styleId ?? '';
  if (!levels.has(styleId)) return false;
  return ROW_NODE_TYPES.has(block?.nodeType ?? '') || block?.nodeType === 'tableOfContents';
}

/**
 * מעצבת את שורות תוכן העניינים: מוביל נקודות ומספר עמוד בקצה.
 *
 * רצה על **כל** שורות תוכן העניינים שבמסמך ולא על טבלה אחת: אין כאן שאלה של
 * „על איזו טבלה המשתמש התכוון” — עצירת טאב זהה בכל שורה של כל טבלה היא בדיוק
 * מה שסגנון היה עושה, ולכן גם מסמך עם שתי טבלאות שאינן ניתנות להבחנה
 * (ראו הערת הפתיחה) מקבל את שתיהן מיושרות.
 *
 * מסמך בלי שורות תוכן עניינים מחזיר `{ok:true}` ולא כשל: הפונקציה נקראת
 * **אחרי** פעולות אחרות, ו„אין מה ליישר” אינו דבר שיש לדווח עליו למשתמש.
 */
export async function layOutTocRows(host: TocTarget): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(LAYOUT_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const setTabStop = doc.format?.paragraph?.setTabStop;
  const list = doc.blocks?.list;
  if (typeof setTabStop !== 'function' || typeof list !== 'function') return unsupported(LAYOUT_FAILED);

  // מידות העמוד מהמסמך ולא קבוע: אזור הטקסט של A4 עם שוליים של אינץ' הוא
  // 9026 twips, ושל Letter עם שוליים צרים הוא אחר לגמרי. עצירה במיקום שאינו
  // של העמוד הזה דוחפת את המספר אל מחוץ לדף או אל אמצע השורה.
  const margins = await readPageMargins(host as PageSetupTarget);
  if (!margins) {
    return unavailable(LAYOUT_FAILED, 'מידות העמוד אינן זמינות', 'no-page-metrics');
  }
  const textWidth = margins.pageWidthTwips - margins.leftTwips - margins.rightTwips;
  const position = textWidth - TOC_PAGE_NUMBER_RESERVE_TWIPS;
  if (position <= 0) {
    // עמוד שרוחב הטקסט שלו קטן מהמקום שמספר עמוד תופס. `w:pos` שאינו חיובי
    // אינו חוקי ב-ECMA-376, והמנוע דווקא מקבל אותו (ראו paragraph-format.ts).
    return unavailable(LAYOUT_FAILED, 'אזור הטקסט בעמוד צר מדי', 'page-too-narrow');
  }

  // סגנון השורה → הרמה שלה. `TOC1`…`TOC9` הם המקרה הרגיל, ולצידם סגנונות
  // המתג `\t` של מסמך שהגיע מ-Word — שם הרמה מגיעה מהמנוע ואינה בשם הסגנון.
  // כשל בקריאה אינו עוצר: רשימה חלקית עדיפה על שום עיצוב.
  const levels = new Map<string, number>();
  for (let level = TOC_LEVEL_MIN; level <= TOC_LEVEL_MAX; level++) levels.set(`TOC${level}`, level);
  const tocList = doc.toc?.list;
  if (typeof tocList === 'function') {
    const listed = await collectAll<TocEntry>(LAYOUT_FAILED, (query) => tocList(query));
    for (const item of listed.items) {
      for (const style of item.preserved?.customStyles ?? []) {
        const name = style?.styleName;
        if (typeof name !== 'string' || name === '') continue;
        levels.set(name, isValidTocLevel(style?.level ?? NaN) ? (style.level as number) : TOC_LEVEL_MIN);
      }
    }
  }

  const rows: { nodeId: string; nodeType: 'paragraph' | 'listItem' }[] = [];
  let offset = 0;
  let guard = 0;
  for (;;) {
    const listed = await attempt(LAYOUT_FAILED, () => list({ limit: PAGE_SIZE, offset }));
    if (!listed.ok) return listed.outcome;

    const page = listed.value?.blocks ?? [];
    for (const block of page) {
      if (!isTocLine(block, levels) || typeof block.nodeId !== 'string') continue;
      rows.push({
        nodeId: block.nodeId,
        // הבלוק הראשון מסווג `tableOfContents`, שאינו סוג חוקי ליעד פסקה.
        nodeType: block.nodeType === 'listItem' ? 'listItem' : 'paragraph',
      });
    }
    if (page.length === 0) break;

    offset += page.length;
    const total = listed.value?.total;
    if (!Number.isFinite(total) || offset >= (total as number)) break;
    if (++guard > PAGE_GUARD) break;
  }

  if (rows.length === 0) {
    // אחרי בנייה שמחקה את כל השורות, מזהים ישנים אינם הוכחה שהעצירות
    // שרדו; אם הן ייווצרו מחדש הן יעוצבו מחדש.
    tocLayoutCache.delete(doc);
    return { ok: true };
  }

  const clearAll = doc.format?.paragraph?.clearAllTabStops;
  const cached = tocLayoutCache.get(doc) ?? new Map<string, number>();
  tocLayoutCache.set(doc, cached);
  const seen = new Set<string>();
  for (const row of rows) {
    const target = { kind: 'block' as const, nodeType: row.nodeType, nodeId: row.nodeId };
    const key = tocLayoutKey(row);
    seen.add(key);

    // עצירות טאב הן חלק מה-DOCX ומחזיקות מעמד ב-`toc.update`; אותה שורה
    // ובאותו רוחב אינה זקוקה לשתי פקודות מנוע נוספות בכל עדכון.
    if (cached.get(key) === position) continue;

    if (typeof clearAll === 'function') {
      const cleared = await attempt(LAYOUT_FAILED, () => clearAll({ target }));
      if (!cleared.ok) return cleared.outcome;
      const failure = failureOf(LAYOUT_FAILED, cleared.value);
      if (failure) return failure;
    }

    const applied = await attempt(LAYOUT_FAILED, () =>
      setTabStop({
        target,
        // `w:pos` נמדד משולי העמוד. אותו מיקום לכל רמה שומר על עמודת
        // מספרים אחת גם ב-SuperDoc וגם כאשר פותחים את ה-DOCX ב-Word.
        position,
        alignment: 'right',
        leader: 'dot',
      }),
    );
    if (!applied.ok) return applied.outcome;
    // עצירה שכבר קיימת חוזרת `NO_OP`, ו-`failureOf` מכריע אותה כהצלחה.
    const failure = failureOf(LAYOUT_FAILED, applied.value);
    if (failure) return failure;
    cached.set(key, position);
  }

  // שורה שנעלמה בבנייה מחדש לא תופיע שוב עם מזהה ישן במפת המטמון.
  for (const key of cached.keys()) if (!seen.has(key)) cached.delete(key);

  return { ok: true };
}

/* ------------------------------------------------------------------ */
/* התאמה אישית                                                         */
/* ------------------------------------------------------------------ */

/** מה שהדיאלוג שולח. שני השדות נמדדו כמשפיעים והפיכים — ראו הערת הפתיחה. */
export interface TocSettings {
  levels: { from: number; to: number };
  hyperlinks: boolean;
}

/**
 * „תוכן עניינים מותאם אישית” — טווח רמות הכותרות והאם הערכים קישורים.
 *
 * שינוי ההגדרות **אינו** בונה את הטבלה מחדש מהמסמך; הוא כותב את המתגים
 * ומרנדר את הטבלה מהמצב הקיים (נמדד: הטקסט השתנה מיד). לכן הפקד ברצועה
 * מריץ אחריו „עדכן טבלה”, וזו החלטה של הממשק ולא של המודול.
 */
export async function configureTableOfContents(
  host: TocTarget,
  settings: TocSettings,
): Promise<CommandOutcome> {
  const levels = normalizeTocLevels(settings.levels.from, settings.levels.to);
  // נדחה לפני שנוגעים במנוע: נמדד שהוא מקבל `{from:9,to:1}` ומרוקן את הטבלה.
  if (levels === null) {
    return { ok: false, message: `${CONFIGURE_FAILED}: ${TOC_LEVELS_HINT}`, reason: 'invalid-levels' };
  }

  const doc = docOf(host);
  if (!doc) return unavailable(CONFIGURE_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const configure = doc.toc?.configure;
  if (typeof configure !== 'function') return unsupported(CONFIGURE_FAILED);

  const sole = await soleToc(host, CONFIGURE_FAILED);
  if (!sole.ok) return sole.outcome;

  const applied = await attempt(CONFIGURE_FAILED, () =>
    configure({
      target: sole.address,
      patch: { outlineLevels: levels, hyperlinks: settings.hyperlinks },
    }),
  );
  if (!applied.ok) return applied.outcome;

  return failureOf(CONFIGURE_FAILED, applied.value) ?? { ok: true };
}

/* ------------------------------------------------------------------ */
/* ערך ידני (שדה TC)                                                   */
/* ------------------------------------------------------------------ */

/**
 * „סמן ערך” — מכניסה שדה `TC` בסוף הפסקה שהסמן בה.
 *
 * `position: 'end'` ולא `'start'`: השדה בלתי נראה במסמך, ובסוף הפסקה הוא
 * אינו דוחף את תחילת הטקסט. זו גם ההתנהגות שברירת המחדל בחוזה מצהירה עליה.
 *
 * העוגן הוא `blockId` של הבחירה ותו לא: החוזה דורש `nodeType: 'paragraph'`,
 * ונמדד שגם `nodeId` של כותרת (`heading`) מתקבל תחתיו — כלומר אין צורך
 * לחסום סימון ערך על כותרת, וחסימה כזאת הייתה מונעת בדיוק את השימוש הנפוץ.
 */
export async function markTocEntry(
  host: TocTarget,
  rawText: string,
  level: number,
): Promise<CommandOutcome> {
  const text = normalizeTocEntryText(rawText);
  // המנוע **זורק** על מחרוזת ריקה („requires a non-empty text string”), ולכן
  // הבדיקה כאן אינה כפילות אלא מה שהופך את הכשל להודעה בעברית.
  if (text === null) {
    return { ok: false, message: `${MARK_FAILED}: יש להקליד את טקסט הערך`, reason: 'invalid-text' };
  }
  if (!isValidTocLevel(level)) {
    return { ok: false, message: `${MARK_FAILED}: ${TOC_LEVEL_HINT}`, reason: 'invalid-level' };
  }

  const doc = docOf(host);
  if (!doc) return unavailable(MARK_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const markEntry = doc.toc?.markEntry;
  if (typeof markEntry !== 'function') return unsupported(MARK_FAILED);

  const selection = await readDocSelection(host as SelectionTarget);
  if (!selection.blockId) return unavailable(MARK_FAILED, NO_ANCHOR_DETAIL, 'no-selection');

  const marked = await attempt(MARK_FAILED, () =>
    markEntry({
      target: {
        kind: 'inline-insert',
        anchor: { nodeType: 'paragraph', nodeId: selection.blockId as string },
        position: 'end',
      },
      text,
      level,
    }),
  );
  if (!marked.ok) return marked.outcome;

  return failureOf(MARK_FAILED, marked.value) ?? { ok: true };
}

/**
 * „בטל סימון ערך” — מוחקת שדה `TC` שסומן קודם.
 *
 * המזהה מגיע מ-`readTocState` ואינו נבנה כאן. החוזה מתאר אותו כ-hash, אבל
 * המדידה מראה מזהה **מיקומי**: `toc-entry:33A6A9C2:0`, כלומר
 * `<blockId>:<index>`, ושני `markEntry` עוקבים על אותה פסקה החזירו את אותו
 * מזהה בדיוק. המסקנה המעשית היא שהמזהה תקף רק לתצלום שממנו נקרא: ביטול של
 * הערך הראשון בפסקה הופך את השני לאינדקס 0, ורשימה ישנה הייתה מצביעה על
 * הערך הלא נכון. לכן ReferencesTab קורא `readTocState` מחדש אחרי כל פעולה,
 * ודיאלוג הערכים מנקה את הבחירה ומאמת אותה מול הרשימה שהתקבלה.
 */
export async function unmarkTocEntry(host: TocTarget, nodeId: string): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(UNMARK_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const unmarkEntry = doc.toc?.unmarkEntry;
  if (typeof unmarkEntry !== 'function') return unsupported(UNMARK_FAILED);

  const unmarked = await attempt(UNMARK_FAILED, () =>
    unmarkEntry({ target: { kind: 'inline', nodeType: 'tableOfContentsEntry', nodeId } }),
  );
  if (!unmarked.ok) return unmarked.outcome;

  return failureOf(UNMARK_FAILED, unmarked.value) ?? { ok: true };
}
