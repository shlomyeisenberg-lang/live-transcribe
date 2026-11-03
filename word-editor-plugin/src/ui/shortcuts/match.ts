/**
 * ההתאמה בין אירוע מקלדת לרשומה ברג'יסטרי. פונקציה טהורה, בקובץ נפרד, כדי
 * שכל מקרי הקצה של פריסת המקלדת יהיו נבדקים בלי להרים קומפוננטה.
 *
 * ## למה `code` ולא `key`
 *
 * `event.key` הוא **התו שהפריסה מייצרת**. בפריסה עברית `Ctrl+S` מדווח
 * `key: 'ד'`, `Ctrl+F` מדווח `key: 'כ'` ו-`Ctrl+B` מדווח `key: 'נ'`. הקוד
 * הקודם השווה `event.key === 's'`, ולכן בעורך שכל ייעודו כתיבת עברית כל
 * הקיצורים מתו ברגע שהמשתמש עבר לעברית — והבאג לא נראה באף בדיקה, כי כל
 * הבדיקות נכתבו בפריסה לטינית.
 *
 * `event.code` הוא המקש הפיזי (`KeyS`), והוא זהה בכל פריסה. לכן הוא ברירת
 * המחדל, ו-`key` נשמר לסימני פיסוק בלבד — שם ההפך נכון: המקש הפיזי של `]`
 * נודד בין פריסות, והתו הוא היציב.
 */
import type { Shortcut } from './registry';

/**
 * המינימום שההתאמה קוראת. לא `KeyboardEvent` מלא: בדיקה שבונה אירוע אמיתי רק
 * כדי לתאר „Ctrl+S בעברית” הייתה תלויה ב-jsdom, וזה בדיוק מה שצריך להיבדק כאן
 * בלי תלות.
 */
export interface KeyEventLike {
  key: string;
  code: string;
  ctrlKey: boolean;
  metaKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
  repeat?: boolean;
  isComposing?: boolean;
  keyCode?: number;
}

/**
 * האם האירוע הוא שלב ביניים של הקלדה עם IME. `keyCode === 229` הוא הדיווח של
 * דפדפנים שאינם מציבים `isComposing`, ובלעדיו קיצור היה נורה באמצע הרכבת תו.
 */
function isComposing(event: KeyEventLike): boolean {
  return event.isComposing === true || event.keyCode === 229;
}

/** Capture may reserve a chord on repeat without running its action again. */
interface MatchOptions {
  ignoreRepeat?: boolean;
}

export function matchShortcut(event: KeyEventLike, shortcut: Shortcut, options?: MatchOptions): boolean {
  if (isComposing(event)) return false;

  // מקש מוחזק חוזר עשרות פעמים בשנייה. „הגדל גופן” אמור לחזור, „שמור” לא.
  if (!options?.ignoreRepeat && event.repeat === true && shortcut.repeatable !== true) return false;

  // Meta שקול ל-Ctrl: מקלדת Mac שולחת Cmd, ואותה רשומה צריכה לתפוס את שניהם.
  if (event.ctrlKey || event.metaKey ? shortcut.ctrl !== true : shortcut.ctrl === true) return false;
  if (event.shiftKey !== (shortcut.shift === true)) return false;
  if (event.altKey !== (shortcut.alt === true)) return false;

  if (shortcut.code) {
    return typeof shortcut.code === 'string'
      ? event.code === shortcut.code
      : shortcut.code.includes(event.code);
  }
  if (shortcut.key) return event.key.toLowerCase() === shortcut.key.toLowerCase();

  // רשומה בלי מקש אינה מתאימה לכלום. בדיקת החוזה מונעת אותה מלכתחילה.
  return false;
}

/** הרשומה הראשונה שמתאימה. `undefined` פירושו „לא הקיצור שלנו — לא נוגעים”. */
export function matchAny(
  event: KeyEventLike,
  shortcuts: readonly Shortcut[],
  options?: MatchOptions,
): Shortcut | undefined {
  return shortcuts.find((shortcut) => matchShortcut(event, shortcut, options));
}
