/**
 * החזרת המיקוד למסמך, דרך `superdoc.focus()`.
 *
 * למה זה כאן ולא ב-`ui/`: זו הדרך היחידה להחזיר את הסמן לטקסט. מיקוד של
 * ה-`<main>` שמארח את המנוע מזיז את הפוקוס אבל **אינו** מחזיר את הסמן — כלומר
 * המשתמש היה מקבל F6 שמעביר אותו „למסמך” ואז הקלדה שלא מגיעה לשום מקום. אזור
 * העריכה עצמו שייך למנוע, וכל שאילתה עליו אסורה מחוץ לתיקייה הזאת
 * (tests/unit/engine-boundaries.test.ts).
 *
 * `restoreSelection` כדי שהחזרה תיפול על מה שהיה מסומן לפני שהמשתמש עבר
 * לרצועה, ולא על תחילת המסמך.
 */
import type { SuperDoc } from 'superdoc';

export interface FocusOptions {
  readonly restoreSelection?: boolean;
  readonly preventScroll?: boolean;
}

export interface FocusHost {
  focus?: (options?: FocusOptions) => unknown;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type FocusTarget = SuperDoc | FocusHost | null | undefined;

/**
 * ממקדת את המסמך. `false` פירושו „אין למי לפנות” — לפני שנפתח מסמך, או בגרסת
 * מנוע שאינה חושפת את הפעולה — ואז מי שקרא נופל למיקוד של אזור המסמך עצמו.
 *
 * לעולם אינה זורקת: היא נקראת מתוך טיפול במקש, וחריגה שם מפילה את המאזין
 * הגלובלי — כלומר את **כל** הקיצורים, ולא רק את זה.
 */
export function focusDocument(host: FocusTarget): boolean {
  const focus = (host as FocusHost | null | undefined)?.focus;
  if (typeof focus !== 'function') return false;

  try {
    focus.call(host, { restoreSelection: true });
    return true;
  } catch {
    return false;
  }
}
