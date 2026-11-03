/**
 * „סימנייה” — הפקד השני בקבוצה „קישורים” של לשונית „הוספה” ב-Word העברי,
 * דרך `doc.bookmarks`.
 *
 * ## שם בעברית מתקבל. זה נמדד, ולא הונח
 *
 * זו הייתה השאלה הפתוחה של הגל, והתשובה חד-משמעית: המנוע שומר, מחזיר,
 * משנה-שם ומוחק סימניות בשם עברי בלי שום עיוות. נמדד ב-Chrome על
 * `file://`, מסמך זרוע, `dist` אמיתי, ביום 25.8.2026:
 *
 *   insert({name:'פרק_ראשון'})           → success, ואותו שם חוזר ב-`list`
 *   get({target:{name:'פרק_ראשון'}})     → אותו שם, עם `range` תקין
 *   rename(→'פרק שני')                   → success, והשם החדש ב-`list`
 *   remove({target:{name:'שם עם רווח'}}) → success
 *
 * ## מה המנוע אוכף על השם, ומה **אינו** אוכף
 *
 * שכבת החוזה אוכפת דבר אחד בלבד: מחרוזת לא ריקה. מעבר לזה המימוש בודק רק
 * כפילות. כל אלה **התקבלו** במדידה, כולם עם `success: true`:
 *
 *   'שם עם רווח'   'סימנייה!סימן'   '1מספר'   '_קו_תחתון'   60 תווים
 *
 * ורק שני אלה נכשלו: `''` (נזרק `INVALID_INPUT`, „requires a non-empty name
 * string”), ושם שכבר קיים (קבלה עם `INVALID_INPUT`, „already exists”).
 *
 * כלומר המנוע יכתוב למסמך שם שהוא **פסול ב-Word**: Word דורש שם שמתחיל
 * באות, בלי רווחים, עד 40 תווים, והוא מסרב לפתוח את דיאלוג הסימניות על שם
 * כזה — וקוד שדה `REF שם עם רווח` נשבר בו לגמרי. מסמך שנשמר עם שם כזה נראה
 * תקין כאן ופגום שם, וזה בדיוק סוג הכשל השקט שאין לשתול במסמך של מישהו אחר.
 *
 * לכן הוולידציה כאן היא **של Word ולא של המנוע**, והיא מוצהרת ככזאת: היא
 * אינה המצאה שלנו אלא הכלל שהיעד — Word — אוכף. „אות” כוללת אות עברית:
 * הביטוי הוא `\p{L}` עם הדגל `u`, ולא `A-Za-z`, וזו ההבחנה שהופכת את הפקד
 * לשמיש בתוסף עברי.
 *
 * ## נקודה אחת שבה אנחנו מחמירים על Word, במודע
 *
 * Word **כן** מתיר קו תחתון פותח — כך הוא מסמן סימנייה מוסתרת, ובשמות
 * האלה הוא משתמש בעצמו (`_Toc…`, `_Ref…`, `_GoBack`). סימנייה שהמשתמש ייצור
 * בשם כזה לא תופיע בדיאלוג הסימניות של Word, והיא עלולה להתנגש בשם ש-Word
 * מייצר לעצמו בעדכון תוכן העניינים הבא. פקד משתמש אינו המקום לייצר סימניות
 * מוסתרות, ולכן קו תחתון פותח נדחה כאן — וההודעה למשתמש אינה מייחסת את
 * ההחמרה הזאת ל-Word.
 *
 * ## הסימנייה מסמנת את **הפסקה**, לא את הטווח שנבחר
 *
 * גם זה נמדד: `insert` על `range: {start:0, end:5}` בפסקה בת 15 תווים החזיר
 * סימנייה שה-`range` שלה `0..15` — כלומר המימוש נשען על `startBlockId` של
 * הכתובת בלבד ומתעלם מההיסטים. אין דרך ציבורית לצמצם אותה, ולכן ה-tooltip
 * אומר „הפסקה שבה הסמן” ולא „הטקסט המסומן”: הבטחה מדויקת עדיפה על הבטחה
 * נכונה-בערך.
 *
 * ## למה אין „עבור אל”
 *
 * ב-Word דיאלוג הסימניות מציע „עבור אל”. `superdoc.ui.selection.apply` אכן
 * קיים בחוזה הציבורי, אבל `activeEditor.view` הוא `null` ב-headless (נמדד
 * בגל קודם) ולכן אי אפשר למדוד שם כלום — ופקד שלא נמדד הוא בדיוק הכפתור המת
 * שהמאגר הזה נבנה כדי לא לייצר. הוא מועמד לגל נפרד, עם מדידה בדפדפן שיש בו
 * view.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';
import { readDocSelection, type SelectionDocumentApi, type SelectionTarget } from './doc-selection';

/**
 * הכלל של Word, לא של המנוע — ראו הערת הפתיחה. `\p{L}` ולא `A-Za-z`: אות
 * עברית היא אות, וסט לטיני היה הופך את הפקד לחסר שימוש בתוסף הזה.
 *
 * `\p{M}` בגוף השם הוא הכרעה מודעת לטובת עברית מנוקדת: במאגר תורני „שָׁלוֹם”
 * הוא הקלדה סבירה לגמרי, סימני הניקוד הם תווים משולבים נפרדים ב-Unicode,
 * ובלעדיהם בסט היה השם נדחה בעוד „שלום” עובר — הבדל שהמשתמש אינו רואה על
 * המסך ואינו יכול לנחש. הם מותרים **אחרי** אות ולא כתו פותח, מפני שסימן
 * משולב בלי אות לפניו אינו „מתחיל באות”.
 *
 * `0-9` ולא `\p{N}`: הכלל של Word הוא ספרות בפועל, וספרה ערבית-הודית כמו
 * `١` הייתה נכנסת דרך `\p{N}` לשם שיישבר ב-Word.
 */
const BOOKMARK_NAME_PATTERN = /^\p{L}[\p{L}\p{M}0-9_]*$/u;

/** התקרה של Word. שם ארוך יותר נחתך שם בשקט, ולכן הוא נדחה כאן. */
export const BOOKMARK_NAME_MAX = 40;

/**
 * מה שמוצג למשתמש כשהשם נדחה. נוסח אחד, גם בדיאלוג וגם בהודעת הכשל.
 *
 * הנוסח אומר מה מותר ואינו אומר „זה הכלל של Word”, ובכוונה: רוב הכלל אכן
 * של Word, אבל האיסור על קו תחתון פותח הוא החמרה שלנו (ראו הערת הפתיחה),
 * וייחוס שלה ל-Word היה אמירה לא נכונה למשתמש.
 */
export const BOOKMARK_NAME_HINT =
  'שם סימנייה מתחיל באות — עברית או לועזית — וממשיך באותיות, בספרות או בקו תחתון. בלי רווחים, עד 40 תווים';

/**
 * „שם כפול” — נחסם בדיאלוג ולא במנוע.
 *
 * המנוע כן מחזיר קבלה עם `INVALID_INPUT` ו-„already exists”, אבל
 * `receiptFailureText` מעדיף את התרגום הגנרי של הקוד על פני `failure.message`
 * (וזה נכון: הודעות המנוע אנגליות), ולכן המשתמש היה שומע „ערך שאינו חוקי” על
 * שם תקין לגמרי. הדיאלוג מחזיק את רשימת השמות ממילא, ולכן הוא זה שאומר את
 * האמת — לפני שנוגעים במסמך.
 */
export const BOOKMARK_NAME_TAKEN_HINT = 'סימנייה בשם הזה כבר קיימת במסמך';

/**
 * השם כפי שיישלח למנוע, או `null` כשהוא פסול.
 *
 * הפונקציה היחידה ששואלת את השאלה: הדיאלוג קורא לה כדי להחליט אם לאפשר
 * אישור, והמודול קורא לה שוב לפני השליחה. שני נוסחים לאותה שאלה היו
 * מאפשרים דיאלוג שמאשר שם שהמודול ידחה — כלומר כפתור שנלחץ ולא קורה כלום.
 * זו בדיוק התבנית של `normalizeLinkHref` ב-payloads.ts.
 */
export function normalizeBookmarkName(raw: string): string | null {
  const name = raw.trim();
  if (name === '' || name.length > BOOKMARK_NAME_MAX) return null;
  return BOOKMARK_NAME_PATTERN.test(name) ? name : null;
}

/** `BookmarkAddress` — מה ש-`get`/`rename`/`remove` מקבלים כ-`target`. */
interface BookmarkAddress {
  kind: 'entity';
  entityType: 'bookmark';
  name: string;
}

/** `BookmarkMutationResult` — הצלחה נושאת `bookmark`, כשל נושא `failure`. */
interface BookmarkReceipt extends DocReceipt {
  bookmark?: unknown;
}

/** `BookmarkDomain` בחלק שנצרך כאן. */
interface BookmarkEntry {
  name?: string;
}

export interface BookmarksDocumentApi extends SelectionDocumentApi {
  bookmarks?: {
    list?: (query?: {
      limit?: number;
      offset?: number;
    }) => MaybePromise<{ items?: readonly BookmarkEntry[]; total?: number } | undefined>;
    insert?: (input: { name: string; at: unknown }) => MaybePromise<BookmarkReceipt>;
    rename?: (input: {
      target: BookmarkAddress;
      newName: string;
    }) => MaybePromise<BookmarkReceipt>;
    remove?: (input: { target: BookmarkAddress }) => MaybePromise<BookmarkReceipt>;
  };
}

/** מה שנדרש מ-SuperDoc: רק הפאסדה של המסמך. ראו page-setup.ts. */
export interface BookmarksHost {
  activeEditor?: { doc?: BookmarksDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type BookmarksTarget = SuperDoc | BookmarksHost | null | undefined;

/**
 * הטיית הכשל בעברית תקנית. „הסימנייה” נקבה, „שם הסימנייה” זכר — הביטוי
 * נשמר שלם ואינו נגזר ממזהה. ראו document-api.ts.
 */
const ADD_FAILED = 'הוספת הסימנייה נכשלה';
const REMOVE_FAILED = 'מחיקת הסימנייה נכשלה';
const RENAME_FAILED = 'שינוי שם הסימנייה נכשל';
const READ_FAILED = 'קריאת הסימניות נכשלה';

/**
 * החוזה מונה שלוש סיבות ל-`target: null`. הנוסח זהה לזה שב-fields.ts ומאותו
 * טעם — המשתמש פוגש את אותו מצב בשני הפקדים, וקול שני לאותו מצב הוא באג.
 */
const NO_TARGET_DETAIL =
  'יש ללחוץ בגוף המסמך, על שורת טקסט שיש בה תו אחד לפחות, ואז להוסיף את הסימנייה';

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

function docOf(host: BookmarksTarget): BookmarksDocumentApi | null {
  return (host as BookmarksHost | null | undefined)?.activeEditor?.doc ?? null;
}

function addressOf(name: string): BookmarkAddress {
  return { kind: 'entity', entityType: 'bookmark', name };
}

/* ------------------------------------------------------------------ */
/* קריאה                                                               */
/* ------------------------------------------------------------------ */

/** מה שהדיאלוג צריך: שמות הסימניות שבמסמך. תצלום ולא מנוי, כמו header-footer.ts. */
export interface BookmarksState {
  names: readonly string[];
}

export function emptyBookmarksState(): BookmarksState {
  return { names: [] };
}

/**
 * שמות כל הסימניות במסמך. לעולם אינה זורקת: כשל של קריאה מחזיר רשימה ריקה,
 * כלומר הדיאלוג יאמר „אין סימניות” — ולא ימציא רשימה חלקית בלי לומר.
 *
 * שאיבת עמודים עד `total`, בדיוק כמו `rebuildAllFields`: `BookmarksListResult`
 * הוא `DiscoveryOutput`, כלומר `items` הוא עמוד תחת `limit`/`offset`. רשימה
 * שמציגה עמוד אחד הייתה מסתירה סימניות במסמך גדול — והמשתמש היה יוצר שם
 * כפול ומקבל „כבר קיים” על שם שאינו רואה.
 */
export async function readBookmarks(host: BookmarksTarget): Promise<BookmarksState> {
  const list = docOf(host)?.bookmarks?.list;
  if (typeof list !== 'function') return emptyBookmarksState();

  const PAGE_SIZE = 200;
  const names: string[] = [];
  let offset = 0;
  let guard = 0;

  for (;;) {
    const listed = await attempt(READ_FAILED, () => list({ limit: PAGE_SIZE, offset }));
    // כשל באמצע השאיבה מחזיר את מה שנאסף עד כה ולא רשימה ריקה: חצי רשימה
    // עדיפה על מסך ריק, וגם היא נכונה — כל שם בה באמת במסמך.
    if (!listed.ok) return { names };

    const items = listed.value?.items ?? [];
    for (const item of items) {
      if (typeof item.name === 'string' && item.name !== '') names.push(item.name);
    }
    if (items.length === 0) return { names };

    offset += items.length;

    const total = listed.value?.total;
    if (!Number.isFinite(total) || offset >= (total as number)) return { names };
    // בלם מפני מנוע שיחזיר `total` שאינו יורד לעולם. ראו rebuildAllFields.
    if (++guard > 1000) return { names };
  }
}

/* ------------------------------------------------------------------ */
/* מוטציות                                                             */
/* ------------------------------------------------------------------ */

/**
 * מוסיפה סימנייה על הפסקה שהסמן בה.
 *
 * `at` הוא פרמטר חובה בחוזה, ואין לו ברירת מחדל „במקום הסמן”. הבחירה נקראת
 * ונמסרת כמו שהיא, מאותו טעם בדיוק כמו ב-fields.ts: היא ה-`TextTarget`
 * שהמנוע עצמו הקרין, ובנייה מחדש שלה הייתה מקבעת אצלנו אחת מכמה צורות.
 */
export async function insertBookmark(
  host: BookmarksTarget,
  rawName: string,
): Promise<CommandOutcome> {
  const name = normalizeBookmarkName(rawName);
  // הכשל הזה מוחזר לפני שנוגעים במנוע ולא אחריו: המנוע היה מקבל את השם
  // בשמחה וכותב אותו למסמך. ראו הערת הפתיחה.
  if (name === null) {
    return { ok: false, message: `${ADD_FAILED}: ${BOOKMARK_NAME_HINT}`, reason: 'invalid-name' };
  }

  const doc = docOf(host);
  if (!doc) return unavailable(ADD_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const insert = doc.bookmarks?.insert;
  if (typeof insert !== 'function') return unsupported(ADD_FAILED);

  const selection = await readDocSelection(host as SelectionTarget);
  if (!selection.target) return unavailable(ADD_FAILED, NO_TARGET_DETAIL, 'no-selection');

  const inserted = await attempt(ADD_FAILED, () => insert({ name, at: selection.target }));
  if (!inserted.ok) return inserted.outcome;

  return failureOf(ADD_FAILED, inserted.value) ?? { ok: true };
}

/**
 * משנה שם של סימנייה קיימת.
 *
 * `newName` עובר את אותה ולידציה כמו שם חדש, ולא רק „אינו ריק”: שינוי שם
 * הוא בדיוק המסלול שבו שם פסול נכנס למסמך שהיה תקין.
 */
export async function renameBookmark(
  host: BookmarksTarget,
  currentName: string,
  rawNewName: string,
): Promise<CommandOutcome> {
  const newName = normalizeBookmarkName(rawNewName);
  if (newName === null) {
    return { ok: false, message: `${RENAME_FAILED}: ${BOOKMARK_NAME_HINT}`, reason: 'invalid-name' };
  }

  const doc = docOf(host);
  if (!doc) return unavailable(RENAME_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const rename = doc.bookmarks?.rename;
  if (typeof rename !== 'function') return unsupported(RENAME_FAILED);

  const renamed = await attempt(RENAME_FAILED, () =>
    rename({ target: addressOf(currentName), newName }),
  );
  if (!renamed.ok) return renamed.outcome;

  return failureOf(RENAME_FAILED, renamed.value) ?? { ok: true };
}

/**
 * מוחקת סימנייה.
 *
 * מוחקת את הסימון בלבד — הטקסט נשאר. זו ההתנהגות של Word, וזה מה שה-hint
 * בדיאלוג אומר.
 */
export async function removeBookmark(
  host: BookmarksTarget,
  name: string,
): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(REMOVE_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const remove = doc.bookmarks?.remove;
  if (typeof remove !== 'function') return unsupported(REMOVE_FAILED);

  const removed = await attempt(REMOVE_FAILED, () => remove({ target: addressOf(name) }));
  if (!removed.ok) return removed.outcome;

  return failureOf(REMOVE_FAILED, removed.value) ?? { ok: true };
}
