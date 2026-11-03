/**
 * הדרך של פקדי הממשק להגיע ל-Document API הציבורי (`superdoc.activeEditor.doc`).
 *
 * למה נדרש מודול נפרד ולא `useCommand`: ה-`CommandAdapter` עוטף את
 * `superdoc.ui.commands` בלבד, וה-registry שלנו (engine/capabilities.ts) מכיל
 * רק פקודות שה-controller מנתב. לפעולות שאין להן פקודה בקטלוג — שוליים, כיוון
 * דף, עמודות, הערות שוליים — אין שום מסלול דרך ה-controller, והמסלול הציבורי
 * היחיד הוא ה-Document API. הוא יושב על המופע, לא על ה-controller, ולכן הוא
 * צריך הזרקה משלו.
 *
 * ההזרקה מחזיקה את המופע עצמו ולא את `doc`: הפאסדה נוצרת מחדש בכל החלפת מסמך,
 * וקריאה של `activeEditor.doc` ברגע השימוש היא מה שמונע החזקה של פאסדה של
 * מסמך שנסגר.
 *
 * המפתח יושב כאן ולא ב-composables/keys.ts (המקום הטבעי שלו, ליד
 * `COMMAND_ADAPTER`) מפני שהקובץ ההוא בבעלות גל אחר בזמן כתיבת הקומיט. ראו
 * הדיווח בסוף הגל.
 */
import type { InjectionKey, Ref } from 'vue';
import type { SuperDoc } from 'superdoc';
import { FAILURE_TEXT } from './command-adapter';

/**
 * המופע הפעיל, או `null` כשאין מסמך פתוח. `shallowRef` בצד המספק: המופע הוא
 * אובייקט זר עם גרפים פנימיים גדולים, ואין שום סיבה לעטוף אותו ב-proxy עמוק.
 */
export const ACTIVE_SUPERDOC: InjectionKey<Ref<SuperDoc | null>> = Symbol('activeSuperdoc');

/** הקבלה שהמנוע מחזיר עשויה להיות סינכרונית או הבטחה — הפאסדה בדפדפן א-סינכרונית. */
export type MaybePromise<T> = T | Promise<T>;

/**
 * הצורה המשותפת לכל קבלה של ה-Document API. מוגדרת כאן ולא מיובאת מהמנוע:
 * הטיפוסים שלו יושבים תחת `superdoc/dist/document-api/...`, ו-import מנתיב
 * פנימי אסור (tests/unit/engine-boundaries.test.ts).
 */
export interface DocReceipt {
  success?: boolean;
  failure?: { code?: string; message?: string };
}

/**
 * קודי כשל של ה-Document API בעברית.
 *
 * הבסיס הוא `FAILURE_TEXT` של command-adapter.ts ולא העתק שלו: המשתמש פוגש את
 * אותו כשל בשני מסלולים (פקודה מנותבת מול Document API ישיר), וקול שני לאותו
 * כשל הוא באג בממשק.
 */
const RECEIPT_FAILURE_TEXT: Record<string, string> = {
  ...FAILURE_TEXT,
  // שני אלה אינם במסלול הפקודות: פעולות ה-Document API **זורקות** על קלט פסול
  // במקום להחזיר קבלה, ולכן הן מגיעות לכאן דרך ה-catch של הקורא.
  INVALID_INPUT: 'הפעולה קיבלה ערך שאינו חוקי',
  PRECONDITION_FAILED: 'המסמך אינו במצב שמאפשר את הפעולה',
};

/**
 * הודעה בעברית מקבלה שנכשלה.
 *
 * `failedAction` הוא ביטוי שלם שכולל את הטיה הכשל („שינוי השוליים נכשל”,
 * „הוספת הערת שוליים נכשלה”) ולא שם עצם: מין דקדוקי אינו נגזר ממזהה, וניסוח
 * גנרי אחד היה מייצר עברית שגויה בחצי מהמקרים.
 *
 * קוד שאין לו תרגום מוצג עם ההסבר של המנוע ועם הקוד עצמו — באנגלית, אבל אפשר
 * לדווח עליו. בלעדיו נשארת הודעה גנרית שאין מה לעשות איתה.
 */
export function receiptFailureText(failedAction: string, receipt: DocReceipt | undefined): string {
  const code = receipt?.failure?.code;
  const known = code ? RECEIPT_FAILURE_TEXT[code] : undefined;
  if (known) return `${failedAction}: ${known}`;

  const message = receipt?.failure?.message;
  if (code && message) return `${failedAction}: ${message} (${code})`;
  if (code) return `${failedAction} (${code})`;
  return failedAction;
}

/** הודעה בעברית מחריגה שנזרקה. */
export function thrownText(failedAction: string, error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  return message ? `${failedAction}: ${message}` : failedAction;
}
