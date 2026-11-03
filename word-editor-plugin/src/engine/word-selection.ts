/**
 * לחיצה כפולה שבוחרת מילה — גם כשהמילה מנוקדת.
 *
 * ## מה נמדד, ולמה זה נדרש
 *
 * המנוע מטפל בלחיצה כפולה בעצמו, ובעברית שאינה מנוקדת הוא צודק. מדידה
 * ב-Chrome אמיתי מול superdoc@2.8.0, על אותו משפט בארבע צורות:
 *
 *   | הטקסט                          | מה שנבחר בלחיצה כפולה |
 *   |--------------------------------|-----------------------|
 *   | `שלום עולם גדול`               | `עולם` — תקין         |
 *   | `שלום, עולם! (גדול)`           | `עולם` — תקין         |
 *   | `בראשית ברא אלהים` עם ניקוד     | `ר` — אות בודדת       |
 *   | אותו טקסט עם ניקוד וטעמים       | `ר` — אות בודדת       |
 *
 * כלומר סימני הניקוד והטעמים נספרים אצל המנוע כמפרידי מילה, וכל מילה מנוקדת
 * מתפרקת לאותיות בודדות. במסמכי אוצריא — טקסט מנוקד ומוטעם — פירוש הדבר
 * שהלחיצה הכפולה כמעט לעולם אינה בוחרת מילה.
 *
 * ## למה שכבה מבחוץ ולא תיקון במנוע
 *
 * מנוע ה-DOCX קנייני: אסור לשנות אותו, לפרק אותו או להסיק ממנו מימוש (ראו
 * THIRD_PARTY_NOTICES.md). לכן התיקון אינו נוגע בו — הוא משתמש ב-API הציבורי
 * של `superdoc` בלבד, כמו כל שאר `src/engine/`:
 *
 *   1. `doc.selection.current()` — מה המנוע בחר בלחיצה הכפולה. זה **הזרע**:
 *      טווח שנמצא בתוך המילה שנלחצה, גם כשהוא אות אחת ממנה.
 *   2. `doc.ranges.resolve(...)` — חלון טקסט סביב הזרע, שממנו מחושב גבול
 *      המילה. חלון ולא הפסקה כולה: `preview.text` נחתך ב-200 תווים (נמדד),
 *      ופסקה באוצריא יכולה להיות פרק שלם.
 *   3. `ui.selection.apply(target)` — קביעת הבחירה המתוקנת. פעולה ציבורית
 *      של ה-controller, המתועדת כ„apply a public selection target”.
 *
 * שתי הקריאות יחד נמדדו ב-8ms על פסקה בת 3,899 תווים, ולכן אין כאן קפיצה
 * גלויה: המנוע בוחר אות, ואנחנו מרחיבים למילה בתוך פריים.
 *
 * ## למה הזרע ולא נקודת הלחיצה
 *
 * אפשר היה לקרוא את הסמן ב-`mousedown` השני ולחשב ממנו. הזרע עדיף: הוא מגיע
 * מהמנוע **אחרי** שכבר החליט מה נלחץ, ולכן אין מרוץ מול קריאה א-סינכרונית של
 * הבחירה ואין הסתמכות על `event.detail`. ואם המנוע יתקן את עצמו יום אחד,
 * הזרע כבר יהיה המילה השלמה, `wordBoundsIn` יחזיר את אותו טווח, ולא תישלח
 * שום `apply` — השכבה תיהפך לחסרת השפעה בלי שיהיה צורך לזכור להסיר אותה.
 */
import type { SuperDoc } from 'superdoc';
import type { MaybePromise } from './document-api';

/**
 * תו שמילה יכולה להתחיל בו: אות עברית (U+05D0–U+05EA, כולל הסופיות, ו-
 * U+05EF–U+05F2 שהם היו״ד המשולשת והדיגרפים), אות לטינית או ספרה. סימן ניקוד
 * אינו כאן — מילה אינה מתחילה בסגול.
 *
 * הטווחים כתובים כהיסטים ולא כתווים: תו צירוף בתוך מחלקת תווים נדבק חזותית
 * למקף שלפניו, ואי אפשר לבדוק בעין אם הטווח שנכתב הוא הטווח שהתכוונו לו.
 */
export const WORD_LETTER = /[0-9A-Za-z\u05D0-\u05EA\u05EF-\u05F2]/;

/**
 * ניקוד וטעמים: U+0591–U+05BD (הטעמים והתנועות עד המתג), U+05BF רפה,
 * U+05C1–U+05C2 נקודות שי״ן ושׂי״ן, U+05C4–U+05C5 נקודה עליונה ותחתונה,
 * U+05C7 קמץ קטן.
 *
 * ארבעה תווים בתוך אותו טווח נשארים בחוץ בכוונה, ולכן הטווח אינו רצף אחד:
 *   - U+05BE מקף — מפריד, בדיוק כמו מקף בלועזית. „על־כן” הן שתי מילים גם
 *     ב-Word, ולחיצה כפולה בוחרת אחת מהן.
 *   - U+05C0 פסק, U+05C3 סוף פסוק, U+05C6 נון הפוכה — סימני פיסוק.
 */
export const WORD_MARK = /[\u0591-\u05BD\u05BF\u05C1\u05C2\u05C4\u05C5\u05C7]/;

/**
 * גרש. חלק מהמילה גם בסופה — „ר׳ יוחנן”, „פ״ג ה׳”. שני התווים: העברי
 * (U+05F3) והלטיני, שהוא מה שיוצא ממקלדת עברית בפועל.
 */
export const WORD_GERESH = /['\u05F3]/;

/**
 * גרשיים. חלק מהמילה רק **בתוכה** — „רמב״ם”, „שו״ת” — ולעולם לא בקצה, שם
 * הוא מרכאות שסוגרות ציטוט. שוב שני התווים: העברי (U+05F4) והלטיני.
 */
export const WORD_GERSHAYIM = /["\u05F4]/;

/** כל מה שממשיך מילה מבפנים. */
export const WORD_INNER = new RegExp(
  `${WORD_LETTER.source}|${WORD_MARK.source}|${WORD_GERESH.source}|${WORD_GERSHAYIM.source}`,
);

/**
 * כמה תווים לקרוא לכל צד של הזרע.
 *
 * 90 ולא יותר: `ranges.resolve` חותך את `preview.text` ב-200 תווים ומסמן
 * `truncated`, וחלון של 180 תווים לכל היותר נשאר מתחת לגבול. 90 תווים הם גם
 * הרבה מעבר למילה עברית מנוקדת ארוכה — ניקוד מכפיל את מספר התווים, ומילה בת
 * 20 אותיות מגיעה לכ-45.
 */
export const WORD_WINDOW_RADIUS = 90;

/** נקודת קצה של בחירה, כפי שהמנוע מחזיר אותה. */
export interface SelectionPointLike {
  kind?: string;
  blockId?: string;
  offset?: number;
  story?: unknown;
}

/** `SelectionTarget` בחלק שנצרך כאן. */
export interface SelectionTargetLike {
  kind?: string;
  start?: SelectionPointLike;
  end?: SelectionPointLike;
  story?: unknown;
  coordinateSpace?: string;
}

/** התשובה של `doc.ranges.resolve`, בחלק שנצרך כאן. */
export interface ResolvedRangeLike {
  preview?: { text?: string; truncated?: boolean };
  target?: { start?: { offset?: number }; end?: { offset?: number } };
}

export interface WordSelectionDoc {
  selection?: {
    current?: (
      input?: { includeText?: boolean },
    ) => MaybePromise<{ selectionTarget?: SelectionTargetLike | null } | undefined>;
  };
  ranges?: { resolve?: (input: unknown) => MaybePromise<ResolvedRangeLike | undefined> };
}

export interface WordSelectionHost {
  activeEditor?: { doc?: WordSelectionDoc | null } | null;
  ui?: { selection?: { apply?: (target: unknown) => unknown } | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type WordSelectionTarget = SuperDoc | WordSelectionHost | null | undefined;

/** טווח תווים בתוך בלוק. */
export interface CharRange {
  start: number;
  end: number;
}

/** חלון טקסט סביב הזרע, כפי ש-`ranges.resolve` החזיר אותו. */
export interface TextWindow {
  /** הטקסט עצמו. */
  text: string;
  /** ההיסט בבלוק של התו הראשון בחלון. */
  base: number;
  /** האם קצה החלון הפותח הוא תחילת הבלוק, ולא גבול הרדיוס. */
  atBlockStart: boolean;
  /** האם קצה החלון הסוגר הוא סוף הבלוק, ולא גבול הרדיוס. */
  atBlockEnd: boolean;
}

/**
 * גבול המילה שההיסט `offset` נמצא בה, בהיסטים של הבלוק.
 *
 * `null` פירושו „אין כאן מילה שאפשר להבטיח עליה”, ואז הקורא אינו נוגע בבחירה
 * של המנוע. ארבעה מצבים כאלה, וכולם מכוונים:
 *
 *   - החלון ריק.
 *   - ההיסט אינו על תו של מילה — רווח, פסיק, סוף פסקה.
 *   - המילה נוגעת בקצה חלון שאינו קצה בלוק, כלומר ייתכן שהיא נמשכת מעבר למה
 *     שנקרא. בחירה חלקית של מילה גרועה מבחירת המנוע: היא נראית נכונה.
 *   - אחרי הקילוף לא נשארה אף אות — למשל לחיצה על מרכאות בודדות.
 */
export function wordBoundsIn(window: TextWindow, offset: number): CharRange | null {
  const { text, base } = window;
  if (text === '') return null;

  // הזרע הוא טווח, וקצהו הסוגר יושב אחרי התו האחרון שבו. לחיצה על התו האחרון
  // במילה מגיעה לכאן עם היסט שמצביע על מה שאחריו, ולכן נבדק גם התו שלפניו.
  const local = offset - base;
  let index = local;
  if (index < 0 || index >= text.length || !WORD_INNER.test(text[index])) {
    index = local - 1;
    if (index < 0 || index >= text.length || !WORD_INNER.test(text[index])) return null;
  }

  let start = index;
  while (start > 0 && WORD_INNER.test(text[start - 1])) start -= 1;
  let end = index + 1;
  while (end < text.length && WORD_INNER.test(text[end])) end += 1;

  if (start === 0 && !window.atBlockStart) return null;
  if (end === text.length && !window.atBlockEnd) return null;

  // קילוף הקצוות. בפתיחה נדרשת אות: ניקוד שנשאר שם שייך לאות שנחתכה ממנו,
  // ומרכאות פותחות אינן חלק מהמילה. בסגירה מותרים גם ניקוד — כל מילה מנוקדת
  // מסתיימת בו — וגם גרש, שהוא סימן קיצור ולא מרכאה.
  while (start < end && !WORD_LETTER.test(text[start])) start += 1;
  while (end > start) {
    const last = text[end - 1];
    if (WORD_LETTER.test(last) || WORD_MARK.test(last) || WORD_GERESH.test(last)) break;
    end -= 1;
  }
  if (start >= end) return null;

  return { start: base + start, end: base + end };
}

/** נקודת קצה לשליחה חזרה למנוע, בשמירת ה-story של הבחירה המקורית. */
function pointAt(blockId: string, offset: number, story: unknown): SelectionPointLike {
  const point: SelectionPointLike = { kind: 'text', blockId, offset };
  if (story !== undefined && story !== null) point.story = story;
  return point;
}

/** הזרע: הבחירה שהמנוע קבע, כשהיא בתוך בלוק טקסט אחד. */
interface Seed {
  blockId: string;
  range: CharRange;
  story: unknown;
}

/**
 * קריאת הזרע מ-`selectionTarget`, ולא מ-`target`: הראשון הוא „the public
 * selection-target model the write APIs consume directly”, כלומר בדיוק הצורה
 * שחוזרת ל-`apply`, והשני הוא רשימת קטעים שאין בה נקודות קצה.
 *
 * `coordinateSpace` שאינו `visible` נפסל: ההיסטים ש-`ranges.resolve` מחזיר הם
 * במרחב הנראה, וערבוב שני המרחבים היה בוחר טווח שגוי בדיוק במסמך שיש בו מעקב
 * שינויים — כלומר במקום שבו טעות היא היקרה ביותר.
 */
function readSeed(target: SelectionTargetLike | null | undefined): Seed | null {
  if (!target || target.kind !== 'selection') return null;
  if (target.coordinateSpace !== undefined && target.coordinateSpace !== 'visible') return null;

  const { start, end } = target;
  if (start?.kind !== 'text' || end?.kind !== 'text') return null;
  if (typeof start.blockId !== 'string' || start.blockId !== end.blockId) return null;
  if (typeof start.offset !== 'number' || typeof end.offset !== 'number') return null;

  return {
    blockId: start.blockId,
    range: { start: Math.min(start.offset, end.offset), end: Math.max(start.offset, end.offset) },
    story: target.story ?? start.story ?? null,
  };
}

export interface SelectWordOptions {
  /**
   * נבדק ברגע האחרון שלפני `apply`. `true` מבטל: המשתמש כבר המשיך ללחיצה
   * הבאה, והבחירה שהייתה נכונה לפני שתי קריאות א-סינכרוניות אינה נכונה עוד.
   */
  isStale?: () => boolean;
}

/**
 * מרחיבה את הבחירה הנוכחית למילה שלמה. מחזירה `true` רק כשהיא באמת שינתה את
 * הבחירה, כדי שבדיקה תמדוד תוצאה ולא כוונה.
 *
 * לעולם אינה זורקת: היא נקראת מתוך מאזין DOM, וחריגה שם מפילה את המאזין
 * ומשאירה את המשתמש בלי שום טיפול בלחיצה כפולה.
 */
export async function selectWordAtSelection(
  host: WordSelectionTarget,
  options: SelectWordOptions = {},
): Promise<boolean> {
  const target = host as WordSelectionHost | null | undefined;
  const doc = target?.activeEditor?.doc;
  const selection = target?.ui?.selection;
  if (!doc || typeof selection?.apply !== 'function') return false;
  if (typeof doc.selection?.current !== 'function') return false;
  if (typeof doc.ranges?.resolve !== 'function') return false;

  try {
    const info = await doc.selection.current();
    const seed = readSeed(info?.selectionTarget);
    if (!seed) return false;
    // זרע ארוך מהחלון אינו זרע של לחיצה כפולה: מילה אינה ארוכה ממנו, ולכן זו
    // בחירה קודמת שהמנוע עוד לא החליף. הרחבה ממנה הייתה קופצת למילה אחרת.
    if (seed.range.end - seed.range.start > WORD_WINDOW_RADIUS) return false;

    const from = Math.max(0, seed.range.start - WORD_WINDOW_RADIUS);
    const to = seed.range.end + WORD_WINDOW_RADIUS;
    const request: Record<string, unknown> = {
      start: { kind: 'point', point: pointAt(seed.blockId, from, seed.story) },
      end: { kind: 'point', point: pointAt(seed.blockId, to, seed.story) },
    };
    if (seed.story) request.in = seed.story;

    const resolved = await doc.ranges.resolve(request);
    const text = resolved?.preview?.text;
    if (typeof text !== 'string' || resolved?.preview?.truncated === true) return false;

    const base = resolved?.target?.start?.offset ?? from;
    // `to` נחתך לסוף הבלוק כשהוא חורג ממנו, וזה מה שמעיד שהחלון הגיע לסוף.
    const windowEnd = resolved?.target?.end?.offset ?? base + text.length;
    const bounds = wordBoundsIn(
      { text, base, atBlockStart: base === 0, atBlockEnd: windowEnd < to },
      seed.range.start,
    );
    if (!bounds) return false;
    if (bounds.start === seed.range.start && bounds.end === seed.range.end) return false;
    if (options.isStale?.() === true) return false;

    const applied: Record<string, unknown> = {
      kind: 'selection',
      start: pointAt(seed.blockId, bounds.start, seed.story),
      end: pointAt(seed.blockId, bounds.end, seed.story),
    };
    if (seed.story) applied.story = seed.story;

    selection.apply(applied);
    return true;
  } catch (error) {
    // הרחבה שנכשלה אינה סיבה להפיל מאזין. המשתמש נשאר עם הבחירה של המנוע.
    console.warn('[otzaria-word] הרחבת הבחירה למילה נכשלה', error);
    return false;
  }
}

/**
 * היסט שאין בלוק שמגיע אליו. `ranges.resolve` **חותך** היסט שחורג מאורך
 * הבלוק ומחזיר את האורך האמיתי — וזו הדרך הציבורית היחידה לשאול „איפה הבלוק
 * הזה נגמר”, שאין לה getter משלה.
 */
export const BLOCK_LENGTH_PROBE = 1_000_000;

/**
 * בוחרת את הפסקה כולה — מה ש-Word עושה בשלוש לחיצות.
 *
 * מחזירה `false` כשהפסקה כבר מסומנת במלואה, כך שבסביבה שבה המנוע כבר עשה את
 * זה בעצמו לא נשלחת בחירה שנייה מיותרת.
 */
export async function selectBlockAtSelection(
  host: WordSelectionTarget,
  options: SelectWordOptions = {},
): Promise<boolean> {
  const target = host as WordSelectionHost | null | undefined;
  const doc = target?.activeEditor?.doc;
  const selection = target?.ui?.selection;
  if (!doc || typeof selection?.apply !== 'function') return false;
  if (typeof doc.selection?.current !== 'function') return false;
  if (typeof doc.ranges?.resolve !== 'function') return false;

  try {
    const info = await doc.selection.current();
    const seed = readSeed(info?.selectionTarget);
    if (!seed) return false;

    const request: Record<string, unknown> = {
      start: { kind: 'point', point: pointAt(seed.blockId, 0, seed.story) },
      end: { kind: 'point', point: pointAt(seed.blockId, BLOCK_LENGTH_PROBE, seed.story) },
    };
    if (seed.story) request.in = seed.story;

    const resolved = await doc.ranges.resolve(request);
    const end = resolved?.target?.end?.offset;
    if (typeof end !== 'number' || end <= 0) return false;
    if (seed.range.start === 0 && seed.range.end === end) return false;
    if (options.isStale?.() === true) return false;

    const applied: Record<string, unknown> = {
      kind: 'selection',
      start: pointAt(seed.blockId, 0, seed.story),
      end: pointAt(seed.blockId, end, seed.story),
    };
    if (seed.story) applied.story = seed.story;

    selection.apply(applied);
    return true;
  } catch (error) {
    console.warn('[otzaria-word] בחירת הפסקה נכשלה', error);
    return false;
  }
}

/**
 * חלון הזמן שמחבר שתי לחיצות לרצף אחד. 500ms הוא ברירת המחדל של Windows
 * (`GetDoubleClickTime`), והוא גם מה ש-Chromium משתמש בו כשהוא סופר לחיצות
 * בעצמו — כלומר הספירה שלנו והספירה שלו אינן יכולות להיפרד.
 */
export const CLICK_SEQUENCE_MS = 500;

/** כמה מותר לסמן לזוז בין לחיצות שהן עדיין אותו רצף. */
export const CLICK_SEQUENCE_SLOP_PX = 5;

export interface WordSelectionHandle {
  /** מפסיק להאזין. הבחירה שכבר נקבעה נשארת. */
  dispose(): void;
}

/** שעון מונוטוני. `Date.now` הוא נפילה-לאחור לסביבה שאין בה `performance`. */
function now(): number {
  return typeof performance === 'object' && typeof performance.now === 'function'
    ? performance.now()
    : Date.now();
}

/**
 * מתקין את התיקון על ה-container של המנוע.
 *
 * ## למה אנחנו סופרים לחיצות בעצמנו, ולא מאזינים ל-`dblclick`
 *
 * ב-Windows אוצריא מריצה את התוסף ב-WebView2 במצב visual hosting, ושם **כל**
 * אירוע עכבר מועבר ידנית מ-Flutter דרך `SendMouseInput`
 * (`flutter_inappwebview_windows`). מה שנשלח שם הוא `MOVE`, `..._BUTTON_DOWN`,
 * `..._BUTTON_UP` ו-`LEAVE` בלבד — `LEFT_BUTTON_DOUBLE_CLICK` אינו קיים
 * בחבילה כלל. ספירת הלחיצות נשארת אפוא בידי Chromium, שמחשב אותה מזמן ומרחק,
 * ולכן היא **אינה מובטחת** ואינה זהה בין הסביבות.
 *
 * הספירה כאן היא באותם ספים בדיוק ({@link CLICK_SEQUENCE_MS},
 * {@link CLICK_SEQUENCE_SLOP_PX}), ולכן היא נותנת את אותה תשובה כשהמנוע צודק,
 * ותשובה נכונה כשהוא לא. שני האירועים שנדרשים לה — `mousedown` ו-`click` —
 * מגיעים בכל סביבה.
 *
 * ## ההחלטות שאינן קוסמטיות
 *
 *   - **שלב ה-capture.** האירוע נרשם על ה-container שלנו ולא על ה-DOM של
 *     המנוע (זה גם מה ש-tests/unit/engine-boundaries.test.ts אוכף), וב-capture
 *     המאזין רץ גם אם המנוע עוצר את ההפצה בדרך למעלה.
 *   - **`click` ולא `mousedown`.** הבחירה נקראת מהמנוע, והוא קובע אותה
 *     בלחיצה עצמה; `click` הוא האירוע הראשון שמובטח שהוא כבר אחריה.
 *   - **גרירה אינה לחיצה.** אם הסמן זז בין ה-`mousedown` ל-`click`, המשתמש
 *     סימן טווח בעצמו — והחלפתו במילה היא בדיוק ההפך ממה שביקש.
 *   - **מונה אינטראקציות.** העבודה א-סינכרונית, והלחיצה הבאה עלולה לקרות
 *     באמצע. בלי המונה ה-`apply` שלנו היה נוחת אחריה ומצמצם את הפסקה שנבחרה
 *     בשלוש לחיצות בחזרה למילה אחת.
 *   - **כפתור ראשי בלבד.** לחיצה כפולה בכפתור ימני פותחת תפריט הקשר, ובחירה
 *     שמשתנה תחת תפריט פתוח היא בדיוק מה שהמשתמש לא ביקש.
 */
/**
 * הסימון שמפריד בין לחיצה של אדם לבין לחיצה שהתוסף עצמו שיגר.
 *
 * תפריט ההקשר מזיז את הסמן בלחיצה שמאלית מסונתזת (composables/use-context-menu.ts,
 * שער ש9), מפני שהמנוע מתעלם מכפתור ימני. בלי ההפרדה הזאת שתי לחיצות ימניות
 * באותה נקודה היו נספרות כאן כלחיצה כפולה — כלומר בחירת מילה שאיש לא ביקש.
 *
 * **`isTrusted` אינו מתאים לתפקיד**, וזה נמדד: בדיקות משגרות אירועים מסונתזים
 * בדיוק כמו שהמשתמש היה לוחץ, וסינון לפיו היה הופך את כל
 * `tests/unit/word-selection.test.ts` לירוק על קוד מת. הסימון כאן אומר „אני
 * שיגרתי את זה”, ולא „זה לא מגיע ממכשיר”.
 */
const SYNTHETIC_POINTER = '__otzariaSyntheticPointer';

/** מסמנת אירוע כשלנו לפני שיגורו. מחזירה אותו, כדי שאפשר יהיה לשרשר. */
export function markSyntheticPointer<E extends Event>(event: E): E {
  Object.defineProperty(event, SYNTHETIC_POINTER, { value: true });
  return event;
}

function isSyntheticPointer(event: Event): boolean {
  return (event as unknown as Record<string, unknown>)[SYNTHETIC_POINTER] === true;
}

export function installWordSelection(
  container: HTMLElement,
  host: WordSelectionTarget,
): WordSelectionHandle {
  let interaction = 0;
  /** אורך הרצף הנוכחי: 1 לחיצה, 2 מילה, 3 פסקה. */
  let clicks = 0;
  let lastAt = 0;
  let lastX = Number.NaN;
  let lastY = Number.NaN;
  /** המקום שבו הכפתור נלחץ, כדי להבדיל לחיצה מגרירה. */
  let downX = Number.NaN;
  let downY = Number.NaN;

  const near = (x: number, y: number, toX: number, toY: number): boolean =>
    Math.abs(x - toX) <= CLICK_SEQUENCE_SLOP_PX && Math.abs(y - toY) <= CLICK_SEQUENCE_SLOP_PX;

  const onMouseDown = (event: MouseEvent): void => {
    if (isSyntheticPointer(event)) return;

    interaction += 1;
    if (event.button !== 0) {
      clicks = 0;
      return;
    }

    const at = now();
    const sequence = at - lastAt <= CLICK_SEQUENCE_MS && near(event.clientX, event.clientY, lastX, lastY);
    // מעל שלוש — Word מתחיל רצף חדש, ולא ממשיך לבחור יחידות גדולות יותר.
    clicks = sequence && clicks < 3 ? clicks + 1 : 1;
    lastAt = at;
    lastX = event.clientX;
    lastY = event.clientY;
    downX = event.clientX;
    downY = event.clientY;
  };

  const onClick = (event: MouseEvent): void => {
    if (isSyntheticPointer(event)) return;
    if (event.button !== 0) return;
    if (!near(event.clientX, event.clientY, downX, downY)) return;

    const token = interaction;
    const isStale = (): boolean => token !== interaction;
    if (clicks === 2) void selectWordAtSelection(host, { isStale });
    else if (clicks === 3) void selectBlockAtSelection(host, { isStale });
  };

  container.addEventListener('mousedown', onMouseDown, true);
  container.addEventListener('click', onClick, true);

  return {
    dispose() {
      container.removeEventListener('mousedown', onMouseDown, true);
      container.removeEventListener('click', onClick, true);
    },
  };
}
