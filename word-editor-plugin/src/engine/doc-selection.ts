/**
 * קריאת הבחירה במסמך דרך `doc.selection.current()`.
 *
 * ## למה זה נדרש בכלל, ולא מסתמכים על הבחירה החיה
 *
 * פקד ברצועה שפותח דיאלוג גוזל את המיקוד מהעורך. פקודות ה-registry שתלויות
 * בבחירה נשענות על `state.selection` של ה-controller, ו-`commandSelectionIsReady`
 * מחזיר `false` כשהיא אינה `ready` — כלומר הפקודה עשויה לרוץ על בחירה שכבר
 * אינה קיימת, או לא לרוץ בכלל. `@pointerdown.prevent` על הכפתור מונע את גזילת
 * המיקוד בלחיצה עצמה, אבל לא ברגע שהמשתמש מקליד בשדה בתוך הדיאלוג.
 *
 * הפתרון שהמנוע עצמו מציע: `executeLinkCommand` בודק `readLinkPayloadTarget`
 * **לפני** הבחירה החיה, ו-`linkPayloadHasExplicitTarget` הוא מה שמכריז על
 * הפקודה כמוכנה גם בלי בחירה. כלומר הדרך הנכונה היא לתפוס את היעד ברגע
 * הלחיצה ולמסור אותו חזרה — וזה מה שהמודול הזה נותן.
 *
 * ## תצלום, לא מנוי
 *
 * המודול מחזיר מצב ברגע הקריאה ואינו מחזיק מנוי על המנוע. זה מה שהצרכן צריך:
 * הוא שואל פעם אחת, בלחיצה, ומחזיק את התשובה עד שהמשתמש מאשר או מבטל.
 *
 * הצורות מוגדרות כאן ואינן מיובאות מהמנוע — הטיפוסים שלו יושבים תחת
 * `superdoc/dist/document-api/...` ו-import מנתיב פנימי אסור
 * (tests/unit/engine-boundaries.test.ts). ההסבר המלא ב-document-api.ts.
 */
import type { SuperDoc } from 'superdoc';
import type { MaybePromise } from './document-api';

/** קטע טקסט בבחירה. כל שדה אופציונלי: זו תשובה של המנוע, לא מבנה שלנו. */
export interface SelectionSegment {
  blockId?: string;
  range?: { start?: number; end?: number };
}

/** `SelectionInfo` בחלק שנצרך כאן. */
export interface SelectionInfoLike {
  empty?: boolean;
  target?: {
    kind?: string;
    segments?: readonly SelectionSegment[];
    story?: unknown;
  } | null;
  selectionTarget?: unknown;
  text?: string;
}

export interface SelectionDocumentApi {
  selection?: {
    current?: (input?: { includeText?: boolean }) => MaybePromise<SelectionInfoLike | undefined>;
  };
}

export interface SelectionHost {
  activeEditor?: { doc?: SelectionDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type SelectionTarget = SuperDoc | SelectionHost | null | undefined;

/** תצלום הבחירה ברגע הלחיצה. */
export interface DocSelectionSnapshot {
  /**
   * ה-`TextTarget` כפי שהמנוע החזיר, למסירה חזרה ב-payload. מועבר **כמו
   * שהוא** ולא נבנה מחדש: `textAddressesFromTarget` מכיר גם `{blockId, range}`
   * וגם `{segments}` וגם `{kind:'selection', start, end}`, ובנייה מחדש הייתה
   * מקבעת אצלנו צורה אחת מהשלוש.
   */
  target: unknown | null;
  /**
   * `SelectionTarget` — הצורה ש**פעולות הכתיבה** צורכות ישירות, ולא `target`.
   *
   * שני שדות ולא אחד, כי המנוע מחזיר שניים והם אינם מחליפים זה את זה:
   * `target` הוא רשימת קטעים (`{segments}`) ואין בו נקודות קצה, ו-
   * `selectionTarget` הוא `{kind:'selection', start, end}`. `hyperlinks.wrap`
   * מקבל את הראשון (`textAddressesFromTarget` מכיר את שלוש הצורות שלו),
   * ו-`doc.insert` מקבל **רק** את השני.
   *
   * מסירת `target` ל-`insert` נכשלת סגור עם `target must be a SelectionTarget
   * object.` — נמדד בשער המעטפת על „ציטוט מהקורא”, שם המשתמש קיבל הודעת
   * שגיאה ושום דבר לא נכתב. הפרדת השדות היא מה שמונע מכל קורא חדש ליפול לשם
   * שוב, במקום לזכור מי מהשניים מתאים לאיזו פעולה.
   */
  selectionTarget: unknown | null;
  /** הטקסט המסומן. `''` כשאין בחירה או כשלא התבקש. */
  text: string;
  /** האם יש טווח ולא רק סמן. קובע אם הקישור יעטוף טקסט קיים או יכניס חדש. */
  hasRange: boolean;
  /** מזהה הפסקה שהבחירה מתחילה בה, או `null`. */
  blockId: string | null;
  /**
   * **כל** הפסקאות שהבחירה נוגעת בהן, לפי סדרן, בלי כפילויות.
   *
   * `blockId` לידו הוא הראשון מהן, והוא נשאר מפני שרוב הצרכנים פועלים על
   * פסקה אחת (דיאלוג הפסקה, מעבר עמוד) — ושינוי שלו לרשימה היה מכריח כל
   * אחד מהם להכריע מחדש על מה הוא פועל.
   *
   * הרשימה קיימת בשביל ההפך: פעולות שב-Word חלות על **כל** הבחירה, כמו
   * „הוסף רווח לפני הפסקה”. `format.paragraph.setSpacing` מקבל כתובת של
   * בלוק יחיד (`ParagraphTarget` בחוזה), ולכן החלה על בחירה מרובה היא לולאה
   * כאן — ואי אפשר לכתוב אותה בלי לדעת מי כל הפסקאות.
   *
   * המנוע מחזיר קטע לכל פסקה גם כשהבחירה רציפה, ולכן אין צורך „למלא” בלוקים
   * שבאמצע: הם שם. קטע בלי `blockId` מדולג.
   */
  blockIds: readonly string[];
  /** ה-story, כשהבחירה אינה בגוף המסמך (כותרת עליונה/תחתונה). */
  story: unknown | null;
  /**
   * `info.empty` כמו שהמנוע דיווח אותו — לא נגזר מ-`hasRange`.
   *
   * ה-JSDoc של `SelectionInfo` במנוע מגדיר אותו כ„True when the selection is
   * empty (cursor only, no range)”, ומבהיר ש-`target` יכול להיות `null` **גם**
   * כשיש בחירה אמיתית שאינה טקסט (בחירת אובייקט — תא טבלה, תמונה): "node
   * selection". `hasRange` לעומת זאת נגזר מ-`target.segments` ונועד לשאלה
   * „אפשר לעטוף בקישור”, ולכן `false` גם כשיש בחירה כזאת. שדה זה הוא מה
   * שמבחין בין השניים — ראו השימוש ב-`mayMoveCaret` (use-context-menu.ts).
   *
   * כשהמנוע אינו מדווח את השדה (`undefined`) נופלים לאחור לחוזה הקודם —
   * `hasRange || text` — כדי לא לשנות התנהגות בגרסת מנוע שאינה חושפת אותו.
   */
  empty: boolean;
}

/** התצלום כשאין בחירה שאפשר לפעול עליה. לא קבוע משותף — כדי שקורא לא ישנה אותו לכולם. */
export function emptySelectionSnapshot(): DocSelectionSnapshot {
  return {
    target: null,
    selectionTarget: null,
    text: '',
    hasRange: false,
    blockId: null,
    blockIds: [],
    story: null,
    empty: true,
  };
}

function segmentsOf(info: SelectionInfoLike): readonly SelectionSegment[] {
  const segments = info.target?.segments;
  return Array.isArray(segments) ? segments : [];
}

/**
 * האם המנוע יתרגם את הקטע לכתובת טקסט שאפשר לעטוף. אותו תנאי בדיוק כמו ב-
 * `textAddressesFromTarget`, כדי ששני הצדדים לא יגיעו לשתי תשובות שונות.
 */
function isWrappableSegment(segment: SelectionSegment | undefined): boolean {
  if (typeof segment?.blockId !== 'string') return false;
  const { start, end } = segment.range ?? {};
  return typeof start === 'number' && typeof end === 'number' && start !== end;
}

/**
 * קוראת את הבחירה. לעולם אינה זורקת ולעולם אינה מחזירה `null`: קורא בממשק
 * צריך תשובה אחת שהוא יכול לבדוק, ולא שלושה מצבים (`null`, זריקה, ריק) שכל
 * אחד מהם דורש טיפול משלו. „אין בחירה” הוא `hasRange: false` עם `target: null`,
 * וזה בדיוק מה שהצרכן שואל.
 *
 * `includeText` נשלח רק כשהוא מבוקש: החוזה מציין במפורש שחילוץ הטקסט עולה
 * בביצועים, ולפקד שאינו מציג את הטקסט אין סיבה לשלם עליו.
 */
export async function readDocSelection(
  host: SelectionTarget,
  options: { includeText?: boolean } = {},
): Promise<DocSelectionSnapshot> {
  const current = (host as SelectionHost | null | undefined)?.activeEditor?.doc?.selection?.current;
  if (typeof current !== 'function') return emptySelectionSnapshot();

  let info: SelectionInfoLike | undefined;
  try {
    info = await current(options.includeText ? { includeText: true } : undefined);
  } catch (error) {
    // קריאת בחירה שנכשלה אינה סיבה להפיל פקד. ללוג, וממשיכים עם „אין בחירה”.
    console.warn('[otzaria-word] קריאת הבחירה במסמך נכשלה', error);
    return emptySelectionSnapshot();
  }

  if (!info || typeof info !== 'object') return emptySelectionSnapshot();

  const segments = segmentsOf(info);
  const first = segments.find((segment) => typeof segment?.blockId === 'string');
  const hasWrappableRange = segments.some((segment) => isWrappableSegment(segment));
  const text = typeof info.text === 'string' ? info.text : '';
  // `Set` ולא סינון: המנוע מחזיר קטע לכל ריצה, ופסקה עם כמה ריצות מסומנות
  // מופיעה יותר מפעם אחת. כתיבה כפולה לאותה פסקה היא NO_OP בשנייה, אבל גם
  // ספירה שגויה בכל מי שיציג „N פסקאות”.
  const blockIds = [
    ...new Set(
      segments
        .map((segment) => segment?.blockId)
        .filter((blockId): blockId is string => typeof blockId === 'string' && blockId !== ''),
    ),
  ];

  return {
    // `target` נמסר רק כשיש בו קטע שאפשר לפעול עליו. `{segments: []}` היה
    // עובר את `linkPayloadHasExplicitTarget` (הוא אובייקט) ואז מייצר רשימת
    // כתובות ריקה — כלומר פקודה שנראית מוכנה ונכשלת סגור.
    target: first ? (info.target ?? null) : null,
    // אובייקט בלבד: המנוע מחזיר `null` כשאין בחירה, ומסירת כל דבר אחר
    // ל-`insert` הייתה זריקת `INVALID_INPUT` במקום נפילה חזרה לסוף המסמך.
    selectionTarget:
      info.selectionTarget !== null && typeof info.selectionTarget === 'object'
        ? info.selectionTarget
        : null,
    text,
    // טווח, ולא `empty` של המנוע: `empty` מתייחס לבחירה כולה, ומה שקובע את
    // המסלול (`hyperlinks.wrap` מול `hyperlinks.insert`) הוא האם יש קטע
    // שהמנוע יתרגם לכתובת טקסט — כלומר **בדיוק** התנאי של
    // `textAddressesFromTarget`: `blockId` שהוא string, ו-`start !== end`.
    //
    // ה-`blockId` הוא חלק מהתנאי ולא קישוט. בלעדיו קטע עם טווח ובלי `blockId`
    // היה מדווח `hasRange: true` לצד `target: null` — הדיאלוג היה מסתיר את
    // שדה הטקסט („הקישור יוחל על הטקסט המסומן”), הפקודה הייתה נשלחת בלי יעד,
    // ולא היה נכתב כלום. בדיקה תפסה את זה.
    hasRange: hasWrappableRange,
    blockId: first?.blockId ?? null,
    blockIds,
    story: info.target?.story ?? null,
    empty: typeof info.empty === 'boolean' ? info.empty : !(hasWrappableRange || text.length > 0),
  };
}
