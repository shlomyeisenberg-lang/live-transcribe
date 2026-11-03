/**
 * ראשי תיבות: מיפוי `"א"א" → "אי אפשר"`. קובץ המקור הוא מיזוג של מילון
 * אוצריא הרשמי ושל KleiKodeshProject — 25,363 ערכים יחד (ר'
 * THIRD_PARTY_NOTICES.md), ומתוכם 17,840 נכנסים לנכס (ר' scripts/acronyms-asset.ts).
 *
 * בניגוד ל-book-completion/static-completion, כאן אין "השלמת המשך" — הראשי
 * תיבות כבר הוקלד **במלואו** (עם הגרשיים), וההשלמה היא הפירוש שלו. לכן
 * ההתאמה היא שווה-ערך מדויק על המילה האחרונה שהוקלדה, לא prefix.
 *
 * ## הנכס הוא אובייקט של פירוש יחיד
 *
 * `scripts/acronyms-asset.ts` מציב ב-`window` ליטרל אובייקט — ולכן הפענוח כאן
 * אינו `JSON.parse` אלא שימוש ישיר. הגרסה הקודמת ציפתה למחרוזת, קיבלה
 * אובייקט, והחזירה `null` בכל טעינה — כלומר ההשלמה כולה לא עבדה.
 *
 * והערך הוא מחרוזת ולא מערך: ה-UI מציג את הפירוש הראשון בלבד, וקובץ המקור
 * מחזיק את כל המערכים (8,877 ערכים עם יותר מפירוש אחד). נשיאת כולם בזיכרון
 * הכפילה את הנכס בלי רווח — מחזור בין פירושים, כשימומש, יטען נכס שני מלא.
 */
import { normalizeWord } from './spellcheck';

/** מה שהנכס מציב ב-`window`: ר"ת → הפירוש שלו. */
export type PackedAcronyms = Readonly<Record<string, string>>;

/**
 * שתי צורות ר"ת, ושתיהן קיימות בנתונים:
 *
 *   - גרשיים בתוך המילה — „רמב"ם”, „חז"ל”. 17,523 מפתחות.
 *   - גרש בסופה — „ר'” (רבי), „וכו'” (וכולי). 31 מפתחות, אך מהשכיחים שבשימוש.
 *
 * זה גם השער שלפני הטעינה, ולכן הוא מוגדר בצמצום: מפתחות KleiKodesh כוללים
 * גם 3,862 מילים ארמיות בלי שום גרש (מילון מונחים שנכנס לאותו קובץ), וקבלתן
 * הייתה מזריקה 0.7MB אחרי כל מילה עברית שהוקלדה — בדיוק העצלות שהנכס קיים
 * בשבילה. הן אינן נכנסות לנכס.
 */
const GERSHAYIM_SHAPE = /^[א-ת]+"[א-ת]+$/;
const GERESH_SHAPE = /^[א-ת]+'$/;

/**
 * הצורה שהמילון מחזיק: בלי ניקוד, וגרשיים ישרים.
 *
 * `normalizeWord` (בדיקת האיות) מטפלת בניקוד ובגרשיים הטיפוגרפיים ״/׳ שמקלדת
 * עברית מייצרת. מעליה כאן שקילות שנוגעת לראשי תיבות בלבד: **שני גרשים
 * נפרדים** הם צורת כתיב מקובלת לגרשיים, ו-1,306 ממפתחות KleiKodesh כתובים
 * כך („א''א”). בלי האיחוד הם מפתחות מקבילים שאיש אינו מגיע אליהם.
 */
export function normalizeAcronym(word: string): string {
  return normalizeWord(word).replace(/''/g, '"');
}

export interface AcronymDictionary {
  /** הפירוש של `acronym`, או `null` אם אינו ר"ת מוכר. */
  lookup(acronym: string): string | null;
}

/**
 * האם המילה נראית ר"ת — מבחן זול שאינו נוגע במילון.
 *
 * זה השער שמונע הזרקת נכס של 0.7MB אחרי כל מילה רגילה שההשלמה מהספר לא מצאה
 * לה המשך. `acronyms-asset` מאמת שכל מפתח בנכס עונה עליו, ולכן אין ערך
 * שהשער מסתיר.
 */
export function looksLikeAcronym(word: string): boolean {
  const normalized = normalizeAcronym(word);
  return GERSHAYIM_SHAPE.test(normalized) || GERESH_SHAPE.test(normalized);
}

export function createAcronymDictionary(packed: PackedAcronyms): AcronymDictionary {
  return {
    lookup(acronym: string): string | null {
      const expansion = packed[normalizeAcronym(acronym)];
      return typeof expansion === 'string' && expansion !== '' ? expansion : null;
    },
  };
}
