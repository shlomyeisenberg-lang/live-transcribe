/**
 * „עדכן הפניות” — חישוב מחדש של שדות ההפניה במסמך, דרך `doc.crossRefs`.
 *
 * ## למה אין כאן „הוספת הפניה מקושרת”, והמדידה שהכריעה
 *
 * זו הייתה המשימה המרכזית של הגל, והיא **בוטלה על סמך מדידה**. `crossRefs.insert`
 * קיים, מוצהר `available: true`, ומחזיר `success: true` — ומכניס למסמך שדה
 * שאינו עובד לא כאן ולא ב-Word.
 *
 * מה שנמדד ב-Chrome על `file://`, `dist` אמיתי, מסמך זרוע, 25.8.2026:
 *
 * 1. **קוד השדה שנכתב אינו קוד Word.** `crossRefs.list().items[].instruction`
 *    אחרי הכנסה מחזיר:
 *
 *        REF SDXREF kind=bookmark;value=%7B%22kind%22%3A%22bookmark%22%2C…;display=pageNumber
 *
 *    כלומר המנוע מקודד את היעד ואת סוג התצוגה כ-JSON מקודד-URL בתוך פרמטר
 *    בשם `SDXREF`. ב-Word שם הסימנייה של שדה `REF` הוא האסימון שאחרי
 *    `REF` — כאן זה `SDXREF` — ולכן Word יציג „שגיאה! מקור ההפניה לא נמצא”.
 *
 * 2. **גם המנוע עצמו אינו יודע לפתור אותו.** אחרי `crossRefs.rebuild` על כל
 *    אחד מהשדות, `resolvedText` נשאר `''` — בכל חמשת סוגי התצוגה שנוסו
 *    (`content`, `pageNumber`, `aboveBelow`, `numberOnly`, `labelAndNumber`),
 *    ועל סימנייה **קיימת**. כלומר המשתמש רואה כלום.
 *
 * 3. **אין ולידציה על היעד.** הכנסה שמפנה לסימנייה שאינה קיימת מחזירה
 *    `success: true` בדיוק כמו הכנסה תקינה.
 *
 * ניסיון החלופה — לכתוב קוד `REF` תקני של Word דרך `fields.insert`, כמו
 * שגל 2 עושה ל-`PAGE` — נמדד גם הוא, ונכשל משתי סיבות בלתי תלויות:
 * `fields.insert` שני על אותה כתובת מוחזר עם `INVALID_CONTEXT` („does not
 * support replacing text inside an existing field”), ומעל זה `bookmarks.insert`
 * מסמן את **כל הפסקה** (ראו bookmarks.ts) — ולכן `REF` על „תוכן הסימנייה”
 * החזיר את הפסקה כולה כולל השדה עצמו, כלומר טקסט שמשכפל את עצמו
 * (`"1פרק ראשוןפרק ראשון"` במדידה).
 *
 * לכן אין כאן פונקציית הכנסה, ואין פקד „הפניה מקושרת” ברצועה. פקד שמכניס
 * למסמך שדה בלתי נראה, שנשבר בפתיחה ב-Word, ומדווח „בוצע” — הוא בדיוק
 * הכפתור המת שהמאגר הזה נבנה כדי לא לייצר.
 *
 * ## מה כן עובד, ולמה זה שווה פקד
 *
 * `crossRefs.list` אינו מוגבל לשדות שהמנוע יצר: הוא מחזיר כל שדה שקוד השדה
 * שלו מתחיל ב-`REF`, `NOTEREF` או `STYLEREF` — כלומר את ההפניות המקושרות
 * שנוצרו **ב-Word**, במסמך שנפתח כאן. `crossRefs.rebuild` על אלה מחשב אותן
 * מחדש כהלכה, כי הן קוד Word תקני. זה המסלול שהפקד משרת: מסמך שהגיע מ-Word,
 * שההפניות בו התיישנו אחרי עריכה.
 *
 * ## החפיפה עם „עדכן שדות”, וההחלטה לשמור על שני הפקדים
 *
 * `rebuildAllFields` (גל 2) מחשב מחדש את **כל** השדות, כולל אלה. הפקד כאן
 * צר יותר בכוונה: מסמך תורני גדול עשוי לשאת מאות שדות `SEQ`, `TOC` ו-`INDEX`
 * שחישובם מחדש יקר, ובלשונית „הפניות” המשתמש מבקש דבר אחד — שההפניות
 * שלו יהיו נכונות. Word עצמו עושה את אותה הבחנה בין F9 על הכול לבין עדכון
 * של טבלה או של שדה בודד.
 */
import type { SuperDoc } from 'superdoc';
import type { CommandOutcome } from './command-adapter';
import { receiptFailureText, thrownText, type DocReceipt, type MaybePromise } from './document-api';

/** `CrossRefDomain` בחלק שנצרך כאן. `address` הוא מה ש-`rebuild` מקבל כ-`target`. */
interface CrossRefEntry {
  address?: unknown;
}

export interface CrossRefsDocumentApi {
  crossRefs?: {
    list?: (query?: {
      limit?: number;
      offset?: number;
    }) => MaybePromise<{ items?: readonly CrossRefEntry[]; total?: number } | undefined>;
    rebuild?: (input: { target: unknown }) => MaybePromise<DocReceipt>;
  };
}

/** מה שנדרש מ-SuperDoc: רק הפאסדה של המסמך. ראו page-setup.ts. */
export interface CrossRefsHost {
  activeEditor?: { doc?: CrossRefsDocumentApi | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type CrossRefsTarget = SuperDoc | CrossRefsHost | null | undefined;

const REBUILD_FAILED = 'עדכון ההפניות נכשל';
const READ_FAILED = 'קריאת ההפניות במסמך נכשלה';

function unavailable(failedAction: string, detail: string, reason: string): CommandOutcome {
  return { ok: false, message: `${failedAction}: ${detail}`, reason };
}

/** הנוסח שהתכנית קובעת ב-§12, וזהה לזה שהיכולת מחזירה. ראו footnotes.ts. */
function unsupported(failedAction: string): CommandOutcome {
  return {
    ok: false,
    message: `${failedAction}: אינו זמין בגרסה זו`,
    reason: 'command-unsupported',
  };
}

/** קריאה למנוע שאינה זורקת החוצה. ראו הערת ה„לעולם לא זורקת” ב-footnotes.ts. */
async function attempt<T>(
  failedAction: string,
  call: () => MaybePromise<T>,
): Promise<{ ok: true; value: T } | { ok: false; outcome: CommandOutcome }> {
  try {
    return { ok: true, value: await call() };
  } catch (error) {
    return {
      ok: false,
      outcome: { ok: false, message: thrownText(failedAction, error), reason: 'threw' },
    };
  }
}

/**
 * כשל הקבלה, או `null` כשהיא הצליחה. `NO_OP` נחשב הצלחה: הפניה שהמנוע כבר
 * חישב מחדש אינה טעונה עדכון. ראו fields.ts.
 */
function failureOf(failedAction: string, receipt: DocReceipt | undefined): CommandOutcome | null {
  const code = receipt?.failure?.code;
  if (receipt?.success !== false || code === 'NO_OP') return null;
  return { ok: false, message: receiptFailureText(failedAction, receipt), reason: code };
}

function docOf(host: CrossRefsTarget): CrossRefsDocumentApi | null {
  return (host as CrossRefsHost | null | undefined)?.activeEditor?.doc ?? null;
}

/** מה שהממשק צריך לדעת: האם יש בכלל מה לעדכן. תצלום ולא מנוי, כמו fields.ts. */
export interface CrossRefsState {
  count: number;
}

export function emptyCrossRefsState(): CrossRefsState {
  return { count: 0 };
}

/**
 * מונה את ההפניות במסמך. לעולם אינה זורקת: כשל של קריאה מחזיר „אין הפניות”,
 * כלומר ה-tooltip יאמר שאין מה לעדכן — ולא ימציא מספר.
 */
export async function readCrossRefsState(host: CrossRefsTarget): Promise<CrossRefsState> {
  const list = docOf(host)?.crossRefs?.list;
  if (typeof list !== 'function') return emptyCrossRefsState();

  const listed = await attempt(READ_FAILED, () => list());
  if (!listed.ok) return emptyCrossRefsState();

  // `total` ולא `items.length`: זה `DiscoveryOutput`, ו-`items` הוא עמוד
  // תחת `limit`/`offset`. ההסבר המלא ב-fields.ts.
  const { items, total } = listed.value ?? {};
  if (typeof total === 'number' && Number.isFinite(total)) return { count: total };
  return { count: Array.isArray(items) ? items.length : 0 };
}

/**
 * „עדכן הפניות” — מחשבת מחדש כל שדה הפניה במסמך.
 *
 * שאיבת עמודים עד `total`, עצירה בכשל הראשון, ומסמך בלי הפניות שמחזיר הצלחה
 * שקטה — שלוש ההחלטות זהות ל-`rebuildAllFields`, ומאותם טעמים בדיוק. ההסבר
 * המלא שם.
 */
export async function rebuildAllCrossRefs(host: CrossRefsTarget): Promise<CommandOutcome> {
  const doc = docOf(host);
  if (!doc) return unavailable(REBUILD_FAILED, 'המסמך עדיין נטען', 'document-api-unavailable');

  const list = doc.crossRefs?.list;
  const rebuild = doc.crossRefs?.rebuild;
  if (typeof list !== 'function' || typeof rebuild !== 'function') {
    return unsupported(REBUILD_FAILED);
  }

  const PAGE_SIZE = 200;
  let offset = 0;
  let guard = 0;

  for (;;) {
    const listed = await attempt(REBUILD_FAILED, () => list({ limit: PAGE_SIZE, offset }));
    if (!listed.ok) return listed.outcome;

    const items = listed.value?.items ?? [];
    if (items.length === 0) return { ok: true };

    for (const entry of items) {
      // הפניה בלי כתובת אינה יעד חוקי ל-`rebuild`, ושליחתה הייתה חריגת
      // `INVALID_TARGET` על שדה שאיש לא ביקש במיוחד.
      if (entry.address === undefined || entry.address === null) continue;

      const rebuilt = await attempt(REBUILD_FAILED, () => rebuild({ target: entry.address }));
      if (!rebuilt.ok) return rebuilt.outcome;
      const failure = failureOf(REBUILD_FAILED, rebuilt.value);
      if (failure) return failure;
    }

    offset += items.length;

    const total = listed.value?.total;
    if (!Number.isFinite(total) || offset >= (total as number)) return { ok: true };
    if (++guard > 1000) return { ok: true };
  }
}
