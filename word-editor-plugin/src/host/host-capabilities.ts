/**
 * מה ה-Host יודע לעשות, לפי הגרסה שדיווח ב-`plugin.boot`.
 *
 * ## למה גרסה, ולא הרשאה ולא ניסיון
 *
 * שלוש דרכים לדעת אם קריאה זמינה, ורק אחת מהן קיימת כאן:
 *
 * 1. **הרשאה** — `BootPayload.permissions`. אינה עוזרת ל-`ui.exportPdf`:
 *    אין לה הרשאת מניפסט כלל, ודיאלוג „שמור בשם” של המערכת הוא שער ההסכמה
 *    (docs/plugin-sdk/API_REFERENCE.md).
 * 2. **גילוי יכולות** — אין. ל-SDK אין קריאה שמחזירה את רשימת המתודות
 *    הנתמכות; `app.getInfo` מחזיר גרסה, מספר build ופלטפורמה בלבד.
 * 3. **ניסיון** — לקרוא ולראות אם נכשל. נשלל: הכפתור היה נראה זמין, המשתמש
 *    היה לוחץ, ורק אז מגלה. כפתור מנוטרל שהטולטיפ שלו מסביר עדיף.
 *
 * נשארת הגרסה. הטבלה במקור היא `docs/plugin-sdk/spec.json` תחת
 * `minAppVersion`, ומה שמשוכפל לכאן הוא רק מה שהתוסף באמת קורא לו.
 */
import { computed, ref } from 'vue';

/** הגרסה שבה `ui.print` ו-`ui.exportPdf` נוספו (spec.json → minAppVersion). */
export const EXPORT_PDF_MIN_APP_VERSION = '0.9.97';

/** גרסת אוצריא כפי שדווחה ב-boot. `null` = טרם ידועה, או שאיננו באוצריא. */
const appVersion = ref<string | null>(null);

/**
 * משווה שתי גרסאות בנקודות. מחזירה שלילי / אפס / חיובי, כמו `Array.sort`.
 *
 * מתעלמת מכל מה שאחרי המספרים (`0.9.97+build3`, `1.0.0-rc1`): הסיומת אינה
 * משנה את הסדר של המספרים לפניה, וגרסת מועמד היא לצורך העניין הגרסה עצמה —
 * מי שמריץ `0.9.97-rc1` כן מקבל את הקריאות שנוספו ב-0.9.97.
 *
 * מקטע שאינו מספר נחשב `0`, ולא פוסל את ההשוואה כולה: גרסה משובשת אינה סיבה
 * להסתיר יכולת, והצד הקורא ממילא נכשל סגור כשאין גרסה בכלל.
 */
export function compareAppVersions(a: string, b: string): number {
  const parts = (value: string): number[] =>
    value
      .split('.')
      .map((piece) => Number.parseInt(piece, 10))
      .map((piece) => (Number.isFinite(piece) ? piece : 0));

  const left = parts(a);
  const right = parts(b);
  const length = Math.max(left.length, right.length);

  for (let i = 0; i < length; i += 1) {
    const diff = (left[i] ?? 0) - (right[i] ?? 0);
    if (diff !== 0) return diff;
  }
  return 0;
}

/** נקראת מ-main.ts ברגע שה-boot נפתר. */
export function setHostAppVersion(version: string | null | undefined): void {
  appVersion.value = typeof version === 'string' && version.trim() ? version.trim() : null;
}

/** הגרסה שנקבעה. מיוצאת לאבחון ולבדיקות. */
export function hostAppVersion(): string | null {
  return appVersion.value;
}

/**
 * האם `ui.exportPdf` זמינה.
 *
 * נכשלת סגור: גרסה שלא דווחה פירושה „לא”, ולא „ננסה בכל זאת”. מחוץ לאוצריא
 * (`host/dev-stub.ts`) אין גרסה, והכפתור מנוטרל עם הסבר — במקום לפתוח דיאלוג
 * מערכת שאינו קיים.
 */
export const supportsPdfExport = computed<boolean>(() => {
  const version = appVersion.value;
  if (!version) return false;
  return compareAppVersions(version, EXPORT_PDF_MIN_APP_VERSION) >= 0;
});
