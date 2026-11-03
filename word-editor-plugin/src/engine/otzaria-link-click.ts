/**
 * לחיצה על קישור `otzaria://` בתוך העורך.
 *
 * ## למה `hyperlinks.onActivate` לבדו אינו מספיק
 *
 * ה-handler ב-`otzaria-link-activation.ts` מחווט נכון ל-`new SuperDoc(...)`,
 * אבל הוא **לעולם אינו נקרא** על קישורי אוצריא. נמדד על 2.15.0, ב-dist הארוז
 * ובדפדפן אמיתי: ה-DomPainter מריץ כל `href` דרך סניטציה מול רשימת סכימות
 * קבועה (`http`, `https`, `mailto`, `tel`, `sms`), ומה שנפסל אינו מצויר
 * כקישור בכלל. שני הצדדים של ההבדל נמדדו זה מול זה באותו מסמך:
 *
 * ```
 * https://example.com/x → <a class="superdoc-text-run superdoc-link"
 *                            data-link-rid="rId7" href="…" target="_blank"
 *                            role="link">          cursor: pointer
 * otzaria://open/book/42 → <span class="superdoc-text-run"
 *                            data-link-rid="rId7" data-link-blocked="true"
 *                            role="text"
 *                            aria-label="Invalid link - not clickable">
 *                                                  cursor: text
 * ```
 *
 * לחיצה על ה-`<span>` — רגילה או עם Ctrl — הפיקה **אפס** קריאות למאחז. אין
 * `<a>`, אין ניווט, ואין גם הודעת שגיאה: מבחינת המנוע לא נלחץ שום קישור.
 *
 * ## הגשר
 *
 * מה שכן נמצא על ה-`<span>` הוא `data-link-rid` — מזהה ה-Relationship של
 * אותו קישור. `hyperlinks.list()` מחזיר לכל קישור את אותו `rId` לצד
 * ‏`externalTarget`, ולכן לחיצה ניתנת לתרגום ליעד בלי לגעת בפנימיות המנוע
 * ובלי להישען על ה-href המצויר (שאינו קיים).
 *
 * ## המחווה: לחיצה רגילה, לא Ctrl בלבד
 *
 * זו אינה העדפה אלא התאמה למה שהמנוע עצמו עושה. נמדד על קישור `https` באותו
 * עורך: **לחיצה רגילה** מפעילה אותו (ה-`<a>` נלחץ), והסמן נקבע במקום הלחיצה
 * באותה פעולה. קישור אוצריא מתנהג עכשיו כמוהו — ומכיוון שאיננו קוראים
 * ל-`preventDefault`, המנוע ממשיך להציב את הסמן בדיוק כמו בכל לחיצה אחרת.
 * Ctrl+לחיצה עוברת באותו מסלול, כי היא אותו אירוע `click`.
 *
 * ## למה `<a>` מדולג
 *
 * אם הסכימה תעבור אי-פעם במעלה הזרם, הקישור ייצבע כ-`<a>` ו-`onActivate`
 * יתחיל להיקרא. הדילוג על `<a>` כאן הוא מה שמונע פתיחה כפולה באותו יום.
 */
import { parseOtzariaLink, type OtzariaLinkTarget } from './otzaria-link';
import type { MaybePromise } from './document-api';

export interface LinkClickDoc {
  hyperlinks?: {
    list?: (input?: Record<string, unknown>) => MaybePromise<unknown>;
  } | null;
}

export interface LinkClickHost {
  activeEditor?: { doc?: LinkClickDoc | null } | null;
}

/** תוצאת הניווט, בצורה ש-`host/otzaria-reader.ts:openOtzariaLink` מחזיר. */
export interface LinkNavigationResult {
  ok: boolean;
  message?: string;
}

export interface LinkClickOptions {
  /** מבצע את הניווט. מוזרק כדי שהבדיקות לא ידרשו גשר. */
  navigate: (target: OtzariaLinkTarget) => MaybePromise<LinkNavigationResult>;
  /** דיווח כשל למשתמש. */
  onStatus?: (message: string, isError: boolean) => void;
}

export interface LinkClickHandle {
  dispose(): void;
}

/** המאפיין שהמנוע מציב על כל ריצה שהיא חלק מקישור, חסום או לא. */
const RID_ATTRIBUTE = 'data-link-rid';

/**
 * ה-`externalTarget` של ה-`rId` המבוקש, מתוך תשובת `hyperlinks.list`.
 *
 * `stories[].hyperlinks[]` ולא `items[]`: רק הראשונה נושאת את ה-`rId`, וזה
 * המפתח היחיד שיש ביד אחרי לחיצה. `rId` ממופה ל-Relationship אחד בדיוק, ולכן
 * ההתאמה הראשונה היא גם היחידה.
 */
function targetOfRid(raw: unknown, rid: string): string | null {
  const stories = (raw as { stories?: unknown } | null)?.stories;
  if (!Array.isArray(stories)) return null;

  for (const story of stories) {
    const links = (story as { hyperlinks?: unknown })?.hyperlinks;
    if (!Array.isArray(links)) continue;
    for (const link of links) {
      const item = link as { rId?: unknown; externalTarget?: unknown };
      if (item?.rId === rid && typeof item.externalTarget === 'string') {
        return item.externalTarget;
      }
    }
  }
  return null;
}

/** מתקינה את הגשר על ה-container של מסמך יחיד, כמו שאר השכבות החיצוניות. */
export function installOtzariaLinkClicks(
  container: HTMLElement,
  host: LinkClickHost | null | undefined,
  options: LinkClickOptions,
): LinkClickHandle {
  let disposed = false;

  async function activate(rid: string): Promise<void> {
    const list = host?.activeEditor?.doc?.hyperlinks?.list;
    if (typeof list !== 'function') return;

    let raw: unknown;
    try {
      raw = await list();
    } catch (error) {
      console.warn('[otzaria-word] קריאת הקישורים במסמך נכשלה', error);
      return;
    }
    if (disposed) return;

    const href = targetOfRid(raw, rid);
    if (!href) {
      // „לחצתי ולא קרה כלום ולא נאמר לי למה” הוא בדיוק התסמין שהגשר נכתב
      // כדי לסגור; אם הצורה שהמנוע מחזיר תשתנה, שיישאר לפחות עקבות.
      console.warn(`[otzaria-word] לא נמצא יעד לקישור ${rid}`);
      return;
    }

    // כל href שאינו אוצריא נשאר באחריות המנוע — הוא מצייר אותו כ-`<a>` אמיתי.
    const target = parseOtzariaLink(href);
    if (!target) return;

    try {
      const result = await options.navigate(target);
      if (disposed) return;
      if (!result?.ok) {
        options.onStatus?.(result?.message ?? 'פתיחת הקישור נכשלה', true);
      }
    } catch (error) {
      console.warn('[otzaria-word] פתיחת הקישור נכשלה', error);
      options.onStatus?.('פתיחת הקישור נכשלה', true);
    }
  }

  /**
   * מאיפה התחילה הלחיצה. גרירה לסימון טקסט מסתיימת גם היא ב-`click`, ובלי
   * ההבחנה הזאת אי אפשר היה לסמן או לתקן את הטקסט של הקישור בלי לצאת
   * מהמסמך — הסימון היה מנווט לקורא.
   */
  let downAt: { x: number; y: number } | null = null;
  const onMouseDown = (event: MouseEvent): void => {
    downAt = event.button === 0 ? { x: event.clientX, y: event.clientY } : null;
  };

  /** מרחק שמעליו זו גרירה ולא לחיצה. רווח ליד יציבה, לא יותר. */
  const DRAG_SLOP_PX = 4;

  const onClick = (event: MouseEvent): void => {
    if (disposed || event.button !== 0 || event.defaultPrevented) return;
    // לחיצה כפולה היא בחירת מילה, ושלוש — בחירת פסקה. אף אחת אינה „פתח”.
    if (event.detail > 1) return;
    if (downAt && Math.hypot(event.clientX - downAt.x, event.clientY - downAt.y) > DRAG_SLOP_PX) {
      return;
    }

    const from = event.target as Element | null;
    if (!from || typeof from.closest !== 'function') return;
    const el = from.closest(`[${RID_ATTRIBUTE}]`);
    // ‏`<a>` = הסכימה עברה את הסניטציה, והמנוע מפעיל אותו בעצמו.
    if (!el || el.tagName === 'A') return;

    const rid = el.getAttribute(RID_ATTRIBUTE);
    if (rid) void activate(rid);
  };

  container.addEventListener('mousedown', onMouseDown);
  container.addEventListener('click', onClick);

  return {
    dispose() {
      disposed = true;
      container.removeEventListener('mousedown', onMouseDown);
      container.removeEventListener('click', onClick);
    },
  };
}
