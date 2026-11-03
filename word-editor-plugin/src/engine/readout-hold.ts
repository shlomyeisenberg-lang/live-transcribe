/**
 * החזקת החיווי של הרצועה בין קריאה לקריאה.
 *
 * ## התקלה שהמודול הזה נולד ממנה
 *
 * בזמן הקלדה החיוויים של „בית” — יישור, צבע גופן, צבע סימון — נכבים וחוזרים
 * שוב ושוב. נמדד ב-Chrome על ה-dist הארוז (scripts/ribbon-typing-probe.mjs):
 * ב-40 שניות של הקלדה רגילה כפתור „יישור לימין” איבד את החיווי הדלוק וקיבל
 * אותו בחזרה **34 פעמים**, כלומר הבהוב אחד לכל כשנייה של כתיבה.
 *
 * ## למה זה קורה
 *
 * `CommandState.value` הוא התשובה של המנוע לשאלה „מה מוחל כאן”. המנוע פותר
 * אותה מקריאות אסינכרוניות שמתאפסות בכל מוטציה במסמך — כלומר בכל תו שנקלד —
 * ובזמן שהן באוויר אין לו תשובה. עד כאן הכול תקין; מה שאינו תקין הוא
 * ש-`readToolbarParagraphAlignment` במנוע מקפל את **כל** המצבים שאינם
 * „ערך אחיד” לאותו `undefined`:
 *
 *     const resolved = resolution.status === "uniform" ? resolution.value : void 0;
 *
 * ל-`resolution.status` יש `'uniform' | 'mixed' | 'pending' | 'unavailable'`,
 * אבל ההבחנה בין „מעורב” לבין „עוד לא יודעים” נזרקת לפני שהיא מגיעה לצרכן.
 * הרצועה קיבלה `undefined` ועשתה את הדבר היחיד שאפשר לעשות עם `undefined` —
 * הציגה „אין יישור” — וזה נראה על המסך כמו כפתור שנכבה.
 *
 * המנוע עצמו מחזיק ערך אחרון (`projectInlineValuesWithSettledHold`), אבל רק
 * לערכי **תו** ורק בבחירת **טווח**: `if (selection.empty) return projection`.
 * כלומר בדיוק המצב שבו המשתמש מקליד — סמן מכווץ — אינו מכוסה, ולערכי **פסקה**
 * אין החזקה בכלל.
 *
 * ## הכלל שכאן, ולמה דווקא הוא
 *
 * `undefined` מוחזק כשידוע ש„אין ערך” אינו יכול להיות האמת:
 *
 * 1. **הבחירה מכווצת.** סמן יושב בפסקה אחת, ולפסקה אחת יש יישור אחד. לכן
 *    `undefined` על סמן הוא **תמיד** „עוד לא נפתר”, לעולם לא „מעורב”. זה הכלל
 *    שמכסה את ההקלדה, וגם את המקרה שהמשתמש תיאר בנפרד — מיקוד שיוצא מהמסמך
 *    מרוקן את כל הרצועה, בזמן ש-Word משאיר את הקריאה האחרונה על המסך.
 * 2. **הקריאה לא התיישבה** (`status !== 'ready'`). הערך האחרון שידענו טוב
 *    מריק, וזו בדיוק המשמעות שהמנוע נותן ל-`stale`: „best-known, not current”.
 *
 * ומה **לא** מוחזק: בחירת טווח שהתיישבה ואין לה ערך. שם `undefined` הוא
 * התשובה האמיתית — „מעורב” — ו-Word מציג בדיוק אותו דבר: אף כפתור יישור אינו
 * דלוק. החזקה שם הייתה משקרת.
 *
 * המדידה תומכת בכלל 1 לבדו על **כל** 15 ההיעלמויות שנמדדו (בכולן
 * `selection.empty === true`), וכלל 2 מוסיף כיסוי לטווח שנקרא מחדש. שני
 * הכללים יחד, ולא אחד מהם, מפני שהם מכסים שני מסלולים שונים: אחד מדבר על מה
 * שהבחירה **היא**, והשני על מה שידוע עליה **כרגע**.
 *
 * ## מה שאינו מוחזק בכוונה: `active`
 *
 * `CommandState.active` הוא בוליאני, ו-`false` בו הוא גם „לא מוחל” וגם „לא
 * ידוע” — אין בו את ההבחנה ש-`undefined` נותן. החזקה שלו הייתה מציגה חיווי
 * דלוק על עיצוב שכבר אינו במסמך, וזה כשל גרוע מהבהוב. הוא גם אינו זקוק לה:
 * המנוע גוזר אותו מ-`selection.activeMarks`, ונמדד שהוא יציב לאורך כל
 * ההקלדה (מודגש נשאר דלוק ברציפות בזמן ש„יישור לימין” הבהב 34 פעמים).
 *
 * ## מה שאינו מוחזק בכוונה: `enabled`
 *
 * זמינות אינה חיווי אלא הבטחה: כפתור שנראה לחיץ ואינו רץ הוא בדיוק התקלה
 * ש-doc-capabilities.ts נכתב בגללה. הוא עובר תמיד טרי.
 */
import type { CommandState } from 'superdoc/ui';

/**
 * מה שההחזקה צריכה לדעת על הבחירה. שני שדות ולא ה-slice כולו: אלה השניים
 * שהכלל נשען עליהם, וצורה צרה היא גם מה שאפשר לבנות בבדיקה בלי לזייף
 * `SelectionSlice` שלם.
 */
export interface ReadoutSelection {
  /** `true` = סמן מכווץ, או שאין בחירה במסמך כלל. */
  readonly empty: boolean;
  /** `true` = הקריאה של הבחירה התיישבה (`status === 'ready'`). */
  readonly settled: boolean;
}

/**
 * מה שמונח לפני שהמנוע דיווח בכלל.
 *
 * `settled: false` ולא `true` — לפני הדיווח הראשון שום דבר לא התיישב, וזו
 * בדיוק הנקודה: ברירת המחדל חייבת להיות זו שמחזיקה, לא זו שמרוקנת.
 */
export const UNSETTLED_SELECTION: ReadoutSelection = { empty: true, settled: false };

/** ה-slice של המנוע בחלק שנצרך כאן. שני השדות אופציונליים — נקרא בהגנה. */
export interface SelectionSliceLike {
  status?: string;
  empty?: boolean;
}

/** מה שנצרך מ-`superdoc.ui`. הכול אופציונלי: גרסה בלי `selection` נופלת בחן. */
export interface ReadoutSelectionSource {
  selection?: {
    getSnapshot?: () => SelectionSliceLike | undefined;
    observe?: (listener: (slice: SelectionSliceLike) => void) => () => void;
  };
}

/**
 * הצורה הצרה מתוך ה-slice של המנוע.
 *
 * `empty` נקרא ברירת מחדל `true` ולא `false`: slice בלי השדה הוא slice שאיננו
 * מבינים, ובמצב כזה עדיף להחזיק את הקריאה האחרונה מאשר לרוקן אותה.
 */
export function toReadoutSelection(slice: SelectionSliceLike | null | undefined): ReadoutSelection {
  return {
    empty: slice?.empty !== false,
    settled: slice?.status === 'ready',
  };
}

/**
 * האם מותר להחזיק את הערך האחרון. ראו את שני הכללים בהערת הראש.
 */
export function canHoldReadout(selection: ReadoutSelection): boolean {
  return selection.empty || !selection.settled;
}

/**
 * הערך שהפקד יציג.
 *
 * ערך טרי מנצח תמיד, גם באמצע החזקה — כך לחיצה על „מרכז” נדלקת מיד ואינה
 * ממתינה להתיישבות, ואין צורך במנגנון ביטול-החזקה נפרד.
 */
export function displayedValue(
  incoming: unknown,
  held: unknown,
  selection: ReadoutSelection,
): unknown {
  if (incoming !== undefined) return incoming;
  return canHoldReadout(selection) ? held : undefined;
}

/**
 * המצב שהפקד יציג: הכול טרי מהמנוע, מלבד הערך שעבר דרך ההחזקה.
 */
export function heldCommandState(
  incoming: CommandState,
  held: unknown,
  selection: ReadoutSelection,
): CommandState {
  const value = displayedValue(incoming.value, held, selection);
  // אותו אובייקט כשאין מה לשנות: `useCommand` מזין ממנו `computed`, ואובייקט
  // חדש בכל דיווח היה מרנדר מחדש כל פקד ברצועה על כל תו שנקלד.
  return value === incoming.value ? incoming : { ...incoming, value };
}

/** קריאה בהגנה: מתודה חסרה או זורקת אינה סיבה להפיל את הרצועה. */
function safeRead<T>(read: (() => T) | undefined): T | undefined {
  if (typeof read !== 'function') return undefined;
  try {
    return read();
  } catch (error) {
    console.warn('[otzaria-word] קריאת מצב הבחירה מהמנוע נכשלה', error);
    return undefined;
  }
}

/**
 * מאזינה למצב הבחירה. `observe` של המנוע יורה מיד עם ה-snapshot ואז על כל
 * שינוי, ולכן אין צורך בקריאה נפרדת לפניה — אותו דפוס בדיוק כמו
 * `observeFontOptions` ו-`observeStyleGallery`.
 *
 * מחזירה disposer גם כשאין `observe`, כדי שאתר הקריאה לא יצטרך להבחין.
 */
export function observeReadoutSelection(
  ui: ReadoutSelectionSource | null | undefined,
  listener: (selection: ReadoutSelection) => void,
): () => void {
  const handle = ui?.selection;
  const observe = handle?.observe;

  if (typeof observe !== 'function') {
    listener(toReadoutSelection(safeRead(handle?.getSnapshot?.bind(handle))));
    return () => {};
  }

  try {
    const off = observe.call(handle, (slice) => listener(toReadoutSelection(slice)));
    return typeof off === 'function' ? off : () => {};
  } catch (error) {
    console.warn('[otzaria-word] האזנה למצב הבחירה נכשלה', error);
    listener(toReadoutSelection(safeRead(handle?.getSnapshot?.bind(handle))));
    return () => {};
  }
}
