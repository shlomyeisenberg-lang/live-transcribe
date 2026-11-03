/**
 * גרירת הדיאלוגים הצפים בכותרת שלהם.
 *
 * ## מה זה פותר
 *
 * עשרים וארבעה דיאלוגים במאגר נפתחים על אותו עוגן קבוע —
 * `top: 140px; inset-inline-start: 40px` — כלומר תמיד באותה פינה, בדיוק מעל
 * הפסקה שהמשתמש עומד עליה כשהוא עובד בראש העמוד. הם אינם חוסמים את המסמך
 * (אין רקע מאחוריהם, ואפשר להקליד תוך כדי), אבל הם **מסתירים** אותו, ולא
 * הייתה שום דרך להזיז אותם. ב-Word כל דיאלוג כזה נגרר בכותרת, וזה מה שיש כאן.
 *
 * ## למה `left`/`top` ולא `inset-inline-start`
 *
 * הכיוון באפליקציה הוא ימין-לשמאל (`<html dir="rtl">`), ולכן
 * `inset-inline-start` של גיליון הסגנון הוא **`right`** בפועל. אילו הגרירה
 * הייתה כותבת `left` בלבד היו שלושתם — `left`, `right` ו-`width` — מוגדרים
 * יחד, וזה מצב over-constrained: הדפדפן מתעלם אז מ-`right` בכיוון ltr,
 * ומ-**`left`** בכיוון rtl. כלומר הדיאלוג פשוט לא היה זז. לכן סגנון הגרירה
 * מנטרל תחילה את שני הקצוות הלוגיים, ורק אחריהם כותב `left` — הסדר בתוך
 * האובייקט הוא הסדר בהצהרה, ומה שנכתב אחרון מנצח.
 *
 * ## קואורדינטות: שתי מערכות, לא אחת
 *
 * רוב הדיאלוגים הם `position: fixed` (קואורדינטות חלון), ו-`FindReplaceDialog`
 * הוא `position: absolute` (קואורדינטות המכל). החישוב אינו מניח אף אחת מהן:
 * `origin` הוא מה שה-CSS כותב, `rect` הוא היכן שזה נחת בחלון, וההצמדה לקצוות
 * נעשית על המלבן — כך שאותו קוד נכון לשתי המשפחות.
 */
import { onUnmounted, ref, type CSSProperties, type Ref } from 'vue';
import { clamp } from './popover-position';

export interface DragRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface DragPoint {
  left: number;
  top: number;
}

export interface DragSize {
  width: number;
  height: number;
}

/**
 * לאן הדיאלוג עובר, בהינתן היכן הוא היה וכמה המצביע זז.
 *
 * @param origin   המיקום כפי שה-CSS כותב אותו — לזה מוסיפים את התזוזה.
 * @param rect     מלבן הדיאלוג בקואורדינטות חלון בתחילת הגרירה.
 * @param delta    תזוזת המצביע מאז הלחיצה.
 * @param viewport גודל החלון.
 *
 * פונקציה טהורה בכוונה: ההצמדה לקצוות היא כל ההיגיון כאן, ו-jsdom אינו מודד
 * פריסה כלל — כלומר בדיקת רכיב לבדה הייתה מאשרת בירוק כל נוסחה שהיא.
 */
export function dialogDragPosition(
  origin: DragPoint,
  rect: DragRect,
  delta: { x: number; y: number },
  viewport: DragSize,
): DragPoint {
  // הדיאלוג נשאר **שלם** בתוך החלון, ולא „רצועה נשארת לתפוס בה”: הכפתורים
  // („אישור” / „ביטול”) יושבים בשורה התחתונה שלו, ודיאלוג שקצהו התחתון יצא
  // מהמסך הוא בדיוק הבאג שהפוטר הדביק בא לתקן — בלבוש של גרירה.
  //
  // `Math.max(0, …)` הוא בשביל דיאלוג שגבוה מהחלון (חלון נמוך במיוחד): שם
  // ההצמדה היא לקצה העליון, כלומר הכותרת נראית ואפשר להמשיך לגרור.
  const left = clamp(rect.left + delta.x, 0, Math.max(0, viewport.width - rect.width));
  const top = clamp(rect.top + delta.y, 0, Math.max(0, viewport.height - rect.height));

  // ההפרש בין המלבן לבין מה שה-CSS כותב נשמר: הוא אפס ב-`fixed`, ואינו אפס
  // ב-`absolute` בתוך מכל ממוקם.
  return { left: origin.left + (left - rect.left), top: origin.top + (top - rect.top) };
}

/**
 * מה שלחיצה עליו אינה מתחילה גרירה.
 *
 * הכותרת אינה רק ידית: יש בה כפתור סגירה, ובחיפוש והחלפה גם לשוניות („חפש” /
 * „החלף”). גרירה שהייתה מתחילה מהם הייתה בולעת את ה-`click` שלהם.
 */
const INTERACTIVE = 'button, a, input, select, textarea, [role="tab"], [role="button"]';

export interface DialogDrag {
  /** נקשר ב-`:style` על שורש הדיאלוג. `undefined` עד הגרירה הראשונה. */
  dragStyle: Ref<CSSProperties | undefined>;
  /** נקשר ב-`@pointerdown` על הכותרת. */
  startDialogDrag: (event: PointerEvent) => void;
}

/**
 * המיקום הנוכחי במערכת שבה ה-CSS כותב.
 *
 * `getComputedStyle().left` על אלמנט ממוקם מחזיר את הערך **בשימוש**
 * בפיקסלים — גם כשהגיליון הצהיר `inset-inline-start` בלבד, וגם כשהערך נגזר
 * מהצד השני. הנפילה למלבן היא בשביל jsdom, שאינו מחשב פריסה ומחזיר `'auto'`;
 * שם הוא גם הערך הנכון, מפני ש-`position: fixed` נמדד בקואורדינטות חלון.
 */
function cssOrigin(root: HTMLElement, rect: DragRect): DragPoint {
  const style = getComputedStyle(root);
  const left = Number.parseFloat(style.left);
  const top = Number.parseFloat(style.top);
  return {
    left: Number.isFinite(left) ? left : rect.left,
    top: Number.isFinite(top) ? top : rect.top,
  };
}

export function useDialogDrag(): DialogDrag {
  const dragStyle = ref<CSSProperties | undefined>(undefined);

  /** המיקום שנקבע בגרירה האחרונה, במערכת של ה-CSS. */
  let placed: DragPoint | null = null;
  /**
   * מה שנדרש כדי להצמיד מחדש בלי לגעת ב-DOM.
   *
   * החלון יכול להשתנות **בזמן שהדיאלוג סגור** (`v-if` מוחק את האלמנט), ואז
   * הפתיחה הבאה הייתה מציבה אותו מחוץ למסך — כלומר דיאלוג בלי כותרת שאפשר
   * לתפוס. המידות של הגרירה האחרונה מספיקות לחשב את ההצמדה מחדש בלי אלמנט.
   */
  let frame: { offsetX: number; offsetY: number; width: number; height: number } | null = null;
  /** קיים רק בזמן גרירה — גם „כבר גוררים” וגם הניקוי בפירוק. */
  let endDrag: (() => void) | null = null;

  function viewport(): DragSize {
    return { width: window.innerWidth, height: window.innerHeight };
  }

  function place(next: DragPoint): void {
    placed = next;
    dragStyle.value = {
      insetInlineStart: 'auto',
      insetInlineEnd: 'auto',
      left: `${next.left}px`,
      top: `${next.top}px`,
    };
  }

  function startDialogDrag(event: PointerEvent): void {
    if (event.button !== 0 || endDrag) return;
    const handle = event.currentTarget as HTMLElement | null;
    const target = event.target as Element | null;
    if (!handle || !target || target.closest(INTERACTIVE)) return;
    const root = handle.closest<HTMLElement>('[role="dialog"]');
    if (!root) return;

    const measured = root.getBoundingClientRect();
    const rect = {
      left: measured.left,
      top: measured.top,
      width: measured.width,
      height: measured.height,
    };
    const origin = cssOrigin(root, rect);
    const startX = event.clientX;
    const startY = event.clientY;
    frame = {
      offsetX: rect.left - origin.left,
      offsetY: rect.top - origin.top,
      width: rect.width,
      height: rect.height,
    };

    // בלי זה הגרירה מסמנת את כותרת הדיאלוג כטקסט.
    event.preventDefault();
    // לכידה: הדיאלוג צף מעל העורך, וסמן שמקדים את הדיאלוג ונכנס לתוכו היה
    // מאבד את ה-`pointermove`. אירוע נלכד ממשיך לעלות אל `window`, ולכן
    // המאזינים נשארים שם.
    try {
      handle.setPointerCapture(event.pointerId);
    } catch {
      /* עכבר בלי לכידה, או jsdom שאינו מממש אותה כלל — הגרירה עובדת בלעדיה */
    }

    const onMove = (moving: PointerEvent): void => {
      place(
        dialogDragPosition(
          origin,
          rect,
          { x: moving.clientX - startX, y: moving.clientY - startY },
          viewport(),
        ),
      );
    };
    const onEnd = (): void => {
      window.removeEventListener('pointermove', onMove);
      window.removeEventListener('pointerup', onEnd);
      window.removeEventListener('pointercancel', onEnd);
      endDrag = null;
      try {
        handle.releasePointerCapture(event.pointerId);
      } catch {
        /* לא נלכד מלכתחילה */
      }
    };

    endDrag = onEnd;
    window.addEventListener('pointermove', onMove);
    window.addEventListener('pointerup', onEnd);
    window.addEventListener('pointercancel', onEnd);
  }

  /** החלון הוקטן — מה שנגרר לקצה מוצמד חזרה פנימה, בלי לגעת באלמנט. */
  function onResize(): void {
    if (!placed || !frame) return;
    place(
      dialogDragPosition(
        placed,
        {
          left: placed.left + frame.offsetX,
          top: placed.top + frame.offsetY,
          width: frame.width,
          height: frame.height,
        },
        { x: 0, y: 0 },
        viewport(),
      ),
    );
  }

  window.addEventListener('resize', onResize);
  onUnmounted(() => {
    window.removeEventListener('resize', onResize);
    endDrag?.();
  });

  return { dragStyle, startDialogDrag };
}
