/**
 * המרות base64 לבייטים ובחזרה.
 *
 * למה מודול משלהן: שני מסלולים בתוסף מעבירים בייטים דרך גשר ה-JS של אוצריא,
 * שהוא JSON בלבד — הטמעת תמונה (`host/files.ts`) וטיוטת השחזור
 * (`host/workspace.ts`) — ושניהם צריכים בדיוק את אותה המרה. עותק שני שלה הוא
 * עותק שני של המלכודת שלמטה.
 *
 * **המלכודת.** `String.fromCharCode(...bytes)` על מערך של מיליוני איברים חורג
 * ממגבלת הארגומנטים של המנוע וזורק `RangeError` — כלומר דווקא הקבצים הגדולים,
 * שבשבילם ההמרה נכתבה, היו נכשלים. לכן העבודה בגושים.
 */

/** גודל הגוש. 32KB בטוח בכל מנוע, וגם מהיר: קריאה אחת ל-`fromCharCode` לגוש. */
const CHUNK = 0x8000;

/** בייטים ל-base64. */
export function bytesToBase64(bytes: Uint8Array): string {
  let binary = '';
  for (let offset = 0; offset < bytes.byteLength; offset += CHUNK) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + CHUNK));
  }
  return btoa(binary);
}

/**
 * base64 לבייטים. `null` על מחרוזת שאינה base64 תקין — `atob` זורק עליה,
 * וקורא שקיבל טיוטה פגומה צריך להתייחס אליה כאילו לא הייתה, לא ליפול.
 */
export function base64ToBytes(text: string): Uint8Array<ArrayBuffer> | null {
  let binary: string;
  try {
    binary = atob(text);
  } catch {
    return null;
  }

  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/**
 * כמה תופסת המחרוזת שתיווצר מ-`byteLength` בייטים.
 *
 * נדרש מפני שהמגבלות של אוצריא הן על מה שעובר בגשר — כלומר על המחרוזת, לא על
 * הבייטים — ובדיקה שמשווה את גודל הבייטים למגבלה מפספסת שליש.
 */
export function base64Length(byteLength: number): number {
  return Math.ceil(byteLength / 3) * 4;
}
