/**
 * „כתב עליון” ו„כתב תחתי”, דרך `doc.format.vertAlign`.
 *
 * ## מה היה כאן
 *
 * שני הפקדים היו `:disabled="true"` קשיח, וה-tooltip אמר „אינו נתמך במנוע
 * הנוכחי”. הטענה הזאת אינה נכונה: `vertAlign` הוא מפתח ב-`InlineRunPatch`,
 * `format.vertAlign` הוא `OperationId` אמיתי בקטלוג, והוא ברשימת הפעולות
 * שהרַנטַיים v2 (זה שרץ בדפדפן אצלנו) מנתב. מה שחסר היה **פקודה ברצועה**: אין
 * מזהה כזה ב-`COMMAND_CATALOG` של ה-controller, ולכן `useCommand` נטרל בצדק,
 * והתווית האשימה את המנוע בחוסר שאינו קיים.
 *
 * המסלול הנכון הוא זה שנבחר לשוליים ולהערות שוליים: פקד בלי פקודה ב-registry,
 * שעובד דרך ה-Document API על `ACTIVE_SUPERDOC`. ראו engine/footnotes.ts
 * ו-engine/page-break.ts.
 *
 * ## למה אין מצב „דלוק” על הכפתור
 *
 * זו החלטה, לא שכחה. `vertAlign` נשמר במנוע כתכונה של mark מסוג `textStyle`
 * (`markTextStyleValue('vertAlign', 'string', 'w:vertAlign', …)`), ואף אחד
 * משלושת המשטחים שהרצועה קוראת מהם אינו מדווח את **הערך**:
 *
 *   1. `ui.commands` — אין פקודת superscript/subscript בקטלוג בכלל.
 *   2. `selection.current().activeMarks` — שמות marks בלבד (`'textStyle'`),
 *      ולא תכונות שלהם.
 *   3. `query.match()` — `MatchStyle` מדווח bold/italic/underline/strike,
 *      color, highlight, fontFamily ו-fontSizePt. אין בו `vertAlign`.
 *
 * מסלול רביעי כן קיים — `getNode` מחזיר `SDRunProps.verticalAlign` — אבל הוא
 * דורש למפות את היסט הסמן לריצה הנכונה בתוך הבלוק, כלומר לשחזר בקוד שלנו את
 * מה שהמנוע כבר עושה. זה נדחה מאותו טעם שמתועד ב-doc-selection.ts. כפתור בלי
 * חיווי „דלוק” הוא פקד חסר-תכונה; כפתור שמחזיק ניחוש מקומי על מצב המסמך הוא
 * פקד שמשקר, וזו התקלה שכל הגל הזה נכתב בגללה.
 *
 * ## הכיבוי, ומה מדווח מצב
 *
 * `NO_OP` הוא הדיווח היחיד על מצב שיש לנו, והוא מדויק: המנוע מחזיר אותו כשה-XML
 * לא השתנה, כלומר כשהבחירה **כבר** בכתב שביקשנו. לכן הלחיצה היא toggle אמיתי
 * בשני שלבים: מחילים את הכתב המבוקש, וכשהתשובה היא `NO_OP` — כלומר „זה כבר
 * המצב” — שולחים `baseline` ומורידים אותו חזרה. אין כאן state מקומי: המנוע הוא
 * שאומר מה המצב, בכל לחיצה מחדש.
 *
 * `'baseline'` ולא `null` לכיבוי, בכוונה: `null` מסיר את העיצוב הישיר ומחזיר
 * את הריצה לירושה מהסגנון — ובטקסט שסגנון התו שלו מגדיר כתב עליון הכיבוי לא
 * היה נראה בכלל. `'baseline'` הוא מה ש-Word עצמו כותב במקרה הזה, והוא היחיד
 * שמבטיח שמה שהמשתמש ביקש יקרה על המסך.
 *
 * ## דרוש טקסט מסומן
 *
 * ה-Document API עובד על טווח ב-XML, ולא על „מה שיוקלד הלאה”. סמן בלי בחירה
 * הוא בקשה שאין לה יעד, ולכן היא נעצרת כאן עם הודעה מדויקת במקום להישלח
 * ולחזור כ-`INVALID_TARGET`. זה שונה מ-Word, שם הכפתור חמוש להקלדה הבאה —
 * וההבדל מתועד ואינו מוסתר.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';
import { readDocCapabilities } from './doc-capabilities';

/** מזהה הפעולה בקטלוג של המנוע. מיוצא כדי שהבדיקה תקבע אותו מול החבילה. */
export const VERT_ALIGN_OPERATION = 'format.vertAlign';

/** הערכים שהסכימה של המנוע מתירה. `null` אינו בשימוש — ראו הערת הפתיחה. */
export type VertAlign = 'superscript' | 'subscript' | 'baseline';

export type VertAlignKind = 'superscript' | 'subscript';

/** הנוסח שהתכנית קובעת ב-§12 לפקד שאין לו API זמין. */
const UNAVAILABLE_TEXT = 'אינו זמין בגרסה זו';

/** ההסבר כשאין בכלל מה לשאול. זהה לנוסח ב-doc-capabilities.ts. */
const LOADING_TEXT = 'המסמך עדיין נטען';

const FAILED_ACTION: Record<VertAlignKind, string> = {
  superscript: 'המעבר לכתב עליון נכשל',
  subscript: 'המעבר לכתב תחתי נכשל',
};

interface VertAlignInput {
  target: unknown;
  value: VertAlign;
}

/** `SelectionInfo` בחלק שנצרך כאן. ראו למטה למה לא דרך doc-selection.ts. */
interface SelectionInfoLike {
  empty?: boolean;
  selectionTarget?: unknown;
}

/** הצורה שנצרכת מ-`doc`. ראו ההסבר ב-document-defaults.ts למה מוגדרת ולא מיובאת. */
export interface VertAlignDocumentApi {
  /**
   * מוצהר כאן כדי ש-`readDocCapabilities` יקבל את ה-host הזה: החוזה המצומצם
   * נבדק מבנית מול הצורה ששם, ובלי השדה הזה אין להם שדה משותף. אותה תבנית
   * בדיוק כמו ב-page-break.ts.
   */
  capabilities?: {
    get?: () => MaybePromise<
      { operations?: Partial<Record<string, { available?: boolean } | undefined>> } | undefined
    >;
  };
  selection?: {
    current?: (input?: { includeText?: boolean }) => MaybePromise<SelectionInfoLike | undefined>;
  };
  format?: {
    vertAlign?: (input: VertAlignInput) => MaybePromise<DocReceipt>;
  };
}

export interface VertAlignHost {
  activeEditor?: { doc?: VertAlignDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type VertAlignTarget = SuperDoc | VertAlignHost | null | undefined;

/** האם הפקד זמין, ומה להציג ב-tooltip כשלא. זהה בצורתה ל-PageBreakSupport. */
export interface VertAlignSupport {
  available: boolean;
  /** הסבר בעברית, מוכן ל-tooltip. מחרוזת ריקה כשזמין. */
  explanation: string;
}

/**
 * קוראת את זמינות הפעולה.
 *
 * לעולם אינה זורקת: כשל של קריאת יכולות אינו סיבה להפיל את רינדור הרצועה.
 * נכשלת **סגור** — „אולי כן” הוא בדיוק הכפתור המת.
 */
export async function readVertAlignSupport(host: VertAlignTarget): Promise<VertAlignSupport> {
  const doc = (host as VertAlignHost | null | undefined)?.activeEditor?.doc;
  if (!doc) return { available: false, explanation: LOADING_TEXT };

  // נוכחות הפונקציה נבדקת **לפני** היכולות, ולא רק בגללן: מפת ה-`operations`
  // נבנית מקטלוג הפעולות, ולכן גרסה שהסירה את המימוש ועוד מכריזה על הפעולה
  // בקטלוג הייתה מחזירה „זמין” לפקד שאין לו למה לקרוא. אותו שיקול כמו
  // ב-page-break.ts, ושם בדיקה תפסה בדיוק את המקרה הזה.
  if (typeof doc.format?.vertAlign !== 'function') {
    return { available: false, explanation: UNAVAILABLE_TEXT };
  }

  const report = await readDocCapabilities(host);
  if (!report.available) return { available: false, explanation: LOADING_TEXT };
  if (report.can('canSetVertAlign')) return { available: true, explanation: '' };

  return {
    available: false,
    explanation: report.explain('canSetVertAlign') || UNAVAILABLE_TEXT,
  };
}

/**
 * ה-`SelectionTarget` של הבחירה החיה, או `null`.
 *
 * `selectionTarget` ולא `target`: החוזה מציין במפורש ש-`target` הוא `TextTarget`
 * לצריכה של comments, ואילו `selectionTarget` הוא „the public selection-target
 * model the write APIs consume directly” — כלומר בדיוק מה ש-`format.*` מקבל.
 *
 * למה לא דרך `readDocSelection` (engine/doc-selection.ts), שזה מקומו הנכון:
 * התצלום שם אינו מחזיר את `selectionTarget` אלא את ה-`TextTarget` בלבד,
 * והמודול ההוא היה בבעלות גל אחר בזמן כתיבת הקומיט הזה. התוספת שנדרשת שם היא
 * שדה אחד; עד שהיא תיעשה, הקריאה כאן היא הקריאה היחידה ולא כפילות שלה.
 */
async function readSelectionTarget(
  doc: VertAlignDocumentApi,
): Promise<{ target: unknown } | { failure: CommandOutcome }> {
  const current = doc.selection?.current;
  if (typeof current !== 'function') {
    return { failure: { ok: false, message: LOADING_TEXT, reason: 'not-ready' } };
  }

  let info: SelectionInfoLike | undefined;
  try {
    info = await current();
  } catch (error) {
    console.warn('[otzaria-word] קריאת הבחירה לכתב עליון/תחתי נכשלה', error);
    info = undefined;
  }

  // בחירה ריקה או יעד שהמנוע לא הצליח להקרין — שני מצבים שונים עם אותה הוראה
  // למשתמש, ולכן אותה הודעה: יש מה לעשות, וזה לסמן טקסט.
  if (!info || info.empty === true || !info.selectionTarget) {
    return {
      failure: {
        ok: false,
        message: 'יש לסמן טקסט תחילה',
        reason: 'range-selection-required',
      },
    };
  }

  return { target: info.selectionTarget };
}

/**
 * מחילה כתב עליון/תחתי על הטקסט המסומן, או מכבה אותו אם הוא כבר במצב הזה.
 *
 * לעולם אינה זורקת: הוולידטורים של ה-Document API **זורקים** `INVALID_INPUT`
 * על קלט פסול במקום להחזיר קבלה, וחריגה מפקד ב-Ribbon מפילה את רינדור הרצועה.
 */
export async function toggleVertAlign(
  host: VertAlignTarget,
  kind: VertAlignKind,
): Promise<CommandOutcome> {
  const failedAction = FAILED_ACTION[kind];
  const doc = (host as VertAlignHost | null | undefined)?.activeEditor?.doc;
  const vertAlign = doc?.format?.vertAlign;

  if (!doc || typeof vertAlign !== 'function') {
    // אותו נוסח שהיכולת מחזירה, כדי שהמשתמש יראה את אותו הסבר בין אם הפקד
    // מנוטרל ובין אם הוא נלחץ לפני שהיכולות נקראו.
    return {
      ok: false,
      message: `${failedAction}: ${UNAVAILABLE_TEXT}`,
      reason: 'command-unsupported',
    };
  }

  const selection = await readSelectionTarget(doc);
  if ('failure' in selection) return selection.failure;

  const apply = async (value: VertAlign): Promise<DocReceipt | CommandOutcome> => {
    try {
      return await vertAlign({ target: selection.target, value });
    } catch (error) {
      return { ok: false, message: thrownText(failedAction, error), reason: 'threw' };
    }
  };

  const first = await apply(kind);
  if ('ok' in first) return first;

  // `NO_OP` = הבחירה כבר בכתב הזה. זהו הדיווח היחיד על מצב שהמנוע נותן, וזה
  // מה שהופך את הלחיצה ל-toggle: מכאן הדרך היא חזרה לבסיס.
  if (first?.success === false && first.failure?.code === 'NO_OP') {
    const second = await apply('baseline');
    if ('ok' in second) return second;

    // `NO_OP` גם בכיבוי: אין מה לשנות בשני הכיוונים (בחירה בלי ריצות טקסט,
    // למשל). מבחינת המשתמש שום דבר לא נשבר, ולכן זו אינה שגיאה.
    if (second?.success === false && second.failure?.code === 'NO_OP') return { ok: true };

    if (second?.success === false) {
      return {
        ok: false,
        message: receiptFailureText('הכיבוי של הכתב העליון/התחתי נכשל', second),
        reason: second.failure?.code,
      };
    }

    return { ok: true };
  }

  if (first?.success === false) {
    return {
      ok: false,
      message: receiptFailureText(failedAction, first),
      reason: first.failure?.code,
    };
  }

  return { ok: true };
}
