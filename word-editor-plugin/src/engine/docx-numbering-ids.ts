/**
 * `w:nsid` ייחודי לכל הגדרת מספור מופשטת — בדרך **החוצה**.
 *
 * ## מה שנמדד
 *
 * רשימה עברית שנוצרה בעורך — גם מ„גימטריה” ברצועה וגם מזיהוי הרשימות בהקלדה —
 * מוצגת ב-Word כ-„1. 2.” ולא כ-„א) ב)”. ב-`numbering.xml` הכול נכון לכאורה:
 * `numId 4 → abstractNum 2 → numFmt hebrew1, lvlText %1)`. אבל המנוע יוצר את
 * `abstractNum 2` כשכפול של העשרוני, ומשאיר לו את אותו `w:nsid` (נמדד:
 * `587013BA` בשניהם). Word מזהה הגדרות מספור לפי ה-nsid, ומציג את שתיהן לפי
 * הראשונה.
 *
 * ההוכחה, ב-A/B על אותו קובץ: שינוי ה-nsid של `abstractNum 2` בלבד, וב-PDF
 * מ-Word הסמנים הם „א) ב) ג) ד)”; בלעדיו — „1. 2. 3. 4.”. זה קיים גם ב-main,
 * ולא נוצר בזיהוי הרשימות, אבל הוא מבטל את התוצאה שלו בכל קובץ שיוצא.
 *
 * ## הכלל
 *
 * ECMA-376 §17.9.12: ה-nsid הוא **המזהה** של ההגדרה המופשטת. שתי הגדרות עם
 * אותו מזהה הן שגיאה, ולא „אותה רשימה” — רשימות שאמורות להמשיך זו את זו
 * חולקות `abstractNumId`, לא nsid. לכן כל הופעה שנייה של nsid מקבלת ערך חדש,
 * והראשונה נשארת כפי שהיא.
 *
 * הערך החדש **דטרמיניסטי** — נגזר מה-nsid המקורי וממזהה ההגדרה — כדי ששתי
 * שמירות של אותו מסמך ייתנו אותם בייטים.
 */
import { SKIPPED_SPANS, TOKEN_SOURCE, valueOf } from './docx-parts';

/** FNV-1a של 32 סיביות, כשמונה ספרות הקסדצימליות גדולות — הצורה של nsid. */
function hex32(input: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < input.length; i += 1) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash.toString(16).toUpperCase().padStart(8, '0');
}

interface NsidTag {
  /** מזהה ההגדרה שהתג בתוכה. */
  abstractId: string;
  value: string;
  at: number;
  end: number;
  tag: string;
}

/**
 * ‏`numbering.xml` שבו לכל `abstractNum` יש nsid משלו, או `null` כשאין מה
 * לשנות.
 */
export function uniqueNumberingIds(xml: string): string | null {
  const tags: NsidTag[] = [];
  let abstractId: string | null = null;
  let depth = 0;

  const token = new RegExp(TOKEN_SOURCE.source, 'g');
  for (let match = token.exec(xml); match; match = token.exec(xml)) {
    const closer = SKIPPED_SPANS.get(match[0]);
    if (closer !== undefined) {
      const end = xml.indexOf(closer, token.lastIndex);
      if (end < 0) break;
      token.lastIndex = end + closer.length;
      continue;
    }

    const [, closing, , name, attributes] = match;
    const selfClosing = attributes.endsWith('/');

    if (name === 'abstractNum' && !selfClosing) {
      if (closing) {
        abstractId = null;
      } else {
        const id = /\s(?:[\w.-]+:)?abstractNumId\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(attributes);
        abstractId = id ? (id[1] ?? id[2] ?? '') : '';
        depth = 0;
      }
      continue;
    }
    if (abstractId === null) continue;

    if (name === 'nsid' && !closing && depth === 0) {
      const value = valueOf(attributes);
      // `w:val=""` נופל כאן, וגם היעדר המאפיין: שתי הגדרות עם nsid ריק לא
      // ייחשבו כפילות. הערך הוא `ST_LongHexNumber` — ארבעה בייטים
      // הקסדצימליים — ו-Word אינו כותב אותו ריק; בלי מקרה נמדד אין כאן מה
      // להוסיף, ולתקן ערך שאינו חוקי מלכתחילה זה לנחש מה הוא היה אמור להיות.
      if (value) tags.push({ abstractId, value, at: match.index, end: token.lastIndex, tag: match[0] });
      continue;
    }
    if (!selfClosing) depth = Math.max(0, depth + (closing ? -1 : 1));
  }

  const used = new Set(tags.map((tag) => tag.value.toUpperCase()));
  const seen = new Set<string>();
  const edits: { at: number; end: number; text: string }[] = [];
  for (const tag of tags) {
    const key = tag.value.toUpperCase();
    if (!seen.has(key)) {
      seen.add(key);
      continue;
    }
    let fresh = hex32(`${key}:${tag.abstractId}`);
    // הלולאה נכנסת כשההאש פוגע בערך שכבר קיים בקובץ, ומה שמגיע לשם אינו
    // התנגשות של FNV-1a: **שלוש הגדרות שחולקות גם nsid וגם `abstractNumId`**
    // מקבלות כולן את אותו קלט להאש, `hex32("<nsid>:<abstractNumId>")`, ולכן
    // השנייה והשלישית נגזרות לאותו ערך בדיוק. בלי הלולאה הן יוצאות מכאן עם
    // nsid כפול — כלומר בדיוק הבאג שהמודול בא לתקן, ובלי שום סימן. גם זה
    // OOXML לא חוקי — `abstractNumId` הוא מזהה — אבל הוא רחוק בסדרי גודל
    // מהתנגשות האש, ולכן הוא הנימוק. המלח משתתף בקלט, ולכן הרביעית ממשיכה
    // משם הלאה.
    //
    // התנגשות אמיתית של ההאש היא אפשרות, ולא מצב שהפלט שלנו מגיע אליו:
    // אחרי שמירה אחת הערך הפוגע כבר הוחלף, ושמירה חוזרת אינה משחזרת אותו.
    // שתי הבדיקות עומדות זו לצד זו: אחת שותלת את התנגשות ההאש בידיים,
    // והשנייה היא שלוש ההגדרות.
    for (let salt = 1; used.has(fresh); salt += 1) fresh = hex32(`${key}:${tag.abstractId}:${salt}`);
    used.add(fresh);
    const text = tag.tag.replace(/(\s(?:[\w.-]+:)?val\s*=\s*["'])[^"']*/, `$1${fresh}`);
    edits.push({ at: tag.at, end: tag.end, text });
  }

  if (edits.length === 0) return null;
  const parts: string[] = [];
  let at = 0;
  for (const edit of edits) {
    parts.push(xml.slice(at, edit.at), edit.text);
    at = edit.end;
  }
  parts.push(xml.slice(at));
  return parts.join('');
}
