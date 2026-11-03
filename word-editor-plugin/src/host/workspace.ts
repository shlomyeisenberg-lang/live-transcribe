/**
 * מרחב הקבצים הפרטי של התוסף (`fs.readFile` / `fs.writeFile` / `fs.stat` /
 * `fs.deleteEntry`, מ-0.9.97).
 *
 * ## למה הוא נדרש, ולמה לא ה-storage
 *
 * `storage.*` הוא KV של JSON, ומה שנשמר בו חייב להיות ערך JSON — כלומר
 * metadata ו-tokens, לא מסמך (ראו host/settings.ts). אבל „לחזור בדיוק למה
 * שהיה” כולל מסמך שהמשתמש **טרם שמר**: למסמך כזה אין קובץ בדיסק ואין token,
 * ולכן הזיכרון היחיד שיכול להחזיק אותו הוא מקום שמקבל בייטים. זה המקום:
 * תיקייה פרטית לתוסף, מכסה 100MB, ובלי הרשאה במניפסט — כל נתיב בה יחסי לשורש
 * של התוסף ואינו יכול לצאת ממנו.
 *
 * ## שלוש ההחלטות שבמימוש
 *
 * 1. **נתיב שטוח.** הקבצים יושבים בשורש המרחב ולא בתת-תיקייה, כדי שלא תידרש
 *    `fs.makeDir` — צעד נוסף שיכול להיכשל, ושכל כתיבה הייתה צריכה לחזור עליו.
 *    שני קבצים בסך הכול; תיקייה לא הייתה קונה כלום.
 *
 * 2. **המגבלה נמדדת על המחרוזת.** המכסה של אוצריא היא 10MB לקריאה או לכתיבה
 *    אחת, ומה שעובר בגשר הוא ה-base64 — שתופס שליש יותר מהבייטים. בדיקה
 *    שמשווה את גודל ה-Blob למגבלה מאשרת בטעות מסמך של 9MB, שהמחרוזת שלו היא
 *    12MB, והכתיבה נדחית אחרי שכבר שילמנו על ההמרה. לכן `MAX_CONTENT_BYTES`.
 *
 * 3. **לעולם לא זורק.** כל מה שהמודול הזה משרת הוא רשת ביטחון: כשל בכתיבת
 *    טיוטה אינו סיבה להפיל עריכה, וכשל בקריאתה אינו סיבה שהתוסף לא יעלה.
 *    התשובה היא ערך מוחזר — `WorkspaceWrite` בכתיבה, `null` בקריאה — והפירוט
 *    הולך ללוג של אוצריא.
 */
import { tryCall } from './otzaria-client';
import { base64ToBytes, bytesToBase64, base64Length } from './base64';

/** המכסה של אוצריא לקריאה או כתיבה אחת. */
export const MAX_PAYLOAD_BYTES = 10 * 1024 * 1024;

/**
 * הגודל המרבי של תוכן שאפשר לכתוב — הבייטים שה-base64 שלהם עדיין נכנס
 * למכסה. ראו החלטה 2 בראש הקובץ.
 */
export const MAX_CONTENT_BYTES = Math.floor((MAX_PAYLOAD_BYTES * 3) / 4);

/**
 * תוצאת כתיבה. שלושה ערכים ולא `boolean`, מפני ששני מצבי הכשל שונים במה
 * שראוי לעשות בהם:
 *
 * - **`too-large`** — המסמך גדול מהמכסה, וזה קבוע: כל ניסיון נוסף ייכשל
 *   באותו אופן עד שהמסמך יקטן. מי שהבטיח למשתמש „לכל היותר דקה של עבודה
 *   באוויר” חייב לומר לו שכאן ההבטחה אינה חלה.
 * - **`failed`** — אוצריא סירבה, ולרוב באופן חולף (מכסה מלאה, דיסק עסוק).
 *   הסבב הבא עשוי להצליח, ואין מה להטריד בו את המשתמש.
 *
 * ההבחנה יושבת כאן ולא אצל הקורא מפני שכאן היא ידועה — כאן נמדד הגודל.
 */
export type WorkspaceWrite = 'written' | 'too-large' | 'failed';

/**
 * כותבת בייטים לקובץ במרחב הפרטי.
 *
 * בייטים ולא `Blob`, אף שזה מה שהייצוא של המנוע מחזיר: המודול הזה מדבר עם
 * הגשר, וההמרה מ-`Blob` היא של מי שמחזיק את המנוע. כך גם אין כאן תלות
 * ב-API של `Blob` שסביבות שונות מממשות בחלקים שונים.
 */
export async function writeWorkspaceBytes(
  path: string,
  bytes: Uint8Array,
): Promise<WorkspaceWrite> {
  if (bytes.byteLength > MAX_CONTENT_BYTES) {
    console.warn(
      `[otzaria-word] ${path} אינו נכתב למרחב הפרטי: ${bytes.byteLength} בייטים` +
        ` (${base64Length(bytes.byteLength)} ב-base64) מעל המכסה של ${MAX_PAYLOAD_BYTES}`,
    );
    return 'too-large';
  }

  const result = await tryCall<{ path?: string }>('fs.writeFile', {
    path,
    content: bytesToBase64(bytes),
    encoding: 'base64',
  });
  if (result === null) {
    console.warn(`[otzaria-word] כתיבת ${path} למרחב הפרטי נכשלה`);
    return 'failed';
  }
  return 'written';
}

/**
 * קוראת בייטים מהמרחב הפרטי. `null` = הקובץ אינו קיים, אינו נקרא, או שתוכנו
 * אינו base64 תקין. שלושת המצבים זהים לצרכן: אין מה לשחזר.
 */
export async function readWorkspaceBytes(path: string): Promise<Uint8Array<ArrayBuffer> | null> {
  const result = await tryCall<{ content?: unknown }>('fs.readFile', {
    path,
    encoding: 'base64',
  });
  if (!result || typeof result.content !== 'string') return null;

  const bytes = base64ToBytes(result.content);
  if (!bytes) {
    console.warn(`[otzaria-word] ${path} במרחב הפרטי אינו base64 תקין`);
    return null;
  }
  return bytes;
}

/**
 * מוחקת קובץ מהמרחב הפרטי. אינה מדווחת כשל: מחיקה של קובץ שכבר אינו קיים היא
 * בדיוק התוצאה שהתבקשה, ואין לקורא מה לעשות עם ההבדל.
 */
export async function deleteWorkspaceEntry(path: string): Promise<void> {
  await tryCall('fs.deleteEntry', { path });
}
