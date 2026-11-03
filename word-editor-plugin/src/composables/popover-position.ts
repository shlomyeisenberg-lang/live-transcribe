/**
 * מיקום הפופאוברים שנפתחים מכפתור ברצועה.
 *
 * ## הבאג שזה מתקן
 *
 * שלושת הפופאוברים של הרצועה (פלטת הצבעים, תפריט הכפתור ובורר הטבלה) מוקמו
 * `position: absolute; top: 100%` בתוך המכל שלהם — כלומר בתוך
 * `.word-ribbon-body`, שמוגדר `overflow-x: auto; overflow-y: hidden`
 * (styles/ribbon.css). ההורה חותך אנכית בגובה הרצועה, ולכן פופאובר שגובהו
 * גדול מהמרווח שנשאר מתחת לכפתור פשוט לא נראה: מפלטת הצבעים, שגובהה ~150px,
 * הוצגה בפועל רק שורת „ללא צבע”.
 *
 * `overflow-y: visible` על הרצועה אינו פתרון — `overflow-x: auto` דורש
 * שהציר השני יהיה `hidden` או `auto`, והדפדפן מתרגם `visible` ל-`auto` שם.
 * כלומר הגלילה האופקית של הרצועה והחיתוך האנכי הם אותה החלטה, ואי אפשר לבטל
 * את השני בלי לוותר על הראשונה.
 *
 * ## למה קוד ולא CSS
 *
 * `position: fixed` יוצא מהחיתוך של האבות — במעטפת אין שום `transform` /
 * `filter` / `contain` שהיה יוצר containing block חדש ומחזיר אותו. אבל הוא גם
 * מנתק את הפופאובר מהכפתור: `top: 100%` מתייחס לחלון ולא למכל, ולכן
 * הקואורדינטות חייבות להימדד. זה מה שיש כאן.
 *
 * החישוב מופרד מה-DOM בכוונה: `popoverPlacement` היא פונקציה טהורה שמקבלת
 * מלבנים ומחזירה מספרים, ולכן ההיפוך למעלה, ההצמדה לקצה והגבלת הגובה נבדקים
 * בלי jsdom — שאינו מודד פריסה כלל ומחזיר אפסים מכל `getBoundingClientRect`.
 */
import { nextTick, onUnmounted, ref, watch, type CSSProperties, type Ref } from 'vue';

/** המרווח בין הכפתור לפופאובר. היה `margin-top: 2px` בשלושת הקבצים. */
export const POPOVER_GAP_PX = 2;

/** כמה הפופאובר נשמר מקצה החלון, בכל הכיוונים. */
export const POPOVER_MARGIN_PX = 8;

/**
 * הגובה המינימלי שמוקצה גם כשאין מקום.
 *
 * בלעדיו חלון נמוך במיוחד היה מקבל `max-height: 0` — כלומר פופאובר פתוח שאינו
 * מציג דבר, וזה בדיוק הבאג המקורי בלבוש אחר. עם המינימום הוא גולל בתוך עצמו.
 */
export const POPOVER_MIN_HEIGHT_PX = 96;

/** מה שנדרש מ-`DOMRect` — כדי שבדיקה תוכל למסור מלבן ולא לזייף DOM. */
export interface AnchorRect {
  top: number;
  bottom: number;
  left: number;
  right: number;
}

export interface Size {
  width: number;
  height: number;
}

export interface Placement {
  /** קואורדינטות חלון, לשימוש ישיר ב-`position: fixed`. */
  top: number;
  left: number;
  /** הגובה שהפופאובר רשאי לתפוס. מעבר לו הוא גולל בתוך עצמו. */
  maxHeight: number;
  /** לאיזה צד של הכפתור נפתח בפועל. */
  side: 'below' | 'above';
}

export interface PlacementOptions {
  /** בכיוון ימין-לשמאל הפופאובר מיושר לקצה **הימני** של הכפתור. */
  rtl?: boolean;
  gap?: number;
  margin?: number;
  /**
   * היישור האופקי ביחס לכפתור.
   *
   * `'start'` — קצה ההתחלה של הכפתור, וזה מה שתפריט או פלטה שנפתחים ממנו
   * עושים: הם רחבים ממנו, והקצה המשותף הוא מה שקושר ביניהם.
   *
   * `'center'` — מרכז מול מרכז, וזה מה שטולטיפ צריך: הוא לרוב *צר* מהמרווח
   * שהוא מכסה ואינו נפתח מהכפתור אלא מסביר אותו, ולכן יישור לקצה היה מסיט
   * אותו הצדה מהאייקון שאליו הוא מתייחס.
   */
  align?: 'start' | 'center';
}

export function clamp(value: number, low: number, high: number): number {
  return Math.min(Math.max(value, low), high);
}

/**
 * לאן הפופאובר הולך, בהינתן הכפתור, הגודל הטבעי של הפופאובר וגודל החלון.
 *
 * כל המידות בקואורדינטות חלון (`clientX`-style), כמו ש-`getBoundingClientRect`
 * מחזיר.
 */
export function popoverPlacement(
  anchor: AnchorRect,
  size: Size,
  viewport: Size,
  options: PlacementOptions = {},
): Placement {
  const gap = options.gap ?? POPOVER_GAP_PX;
  const margin = options.margin ?? POPOVER_MARGIN_PX;

  const roomBelow = viewport.height - anchor.bottom - gap - margin;
  const roomAbove = anchor.top - gap - margin;

  // למטה הוא ברירת המחדל — זה הכיוון שהמשתמש מצפה לו מכפתור ברצועה — ורק
  // כשהפופאובר אינו נכנס שם בוחנים את הצד השני. כששני הצדדים קטנים מדי בוחרים
  // את הגדול מביניהם ומגבילים גובה, כלומר גלילה בתוך הפופאובר ולא חיתוך.
  const side: Placement['side'] = size.height <= roomBelow || roomBelow >= roomAbove ? 'below' : 'above';
  const maxHeight = Math.max(side === 'below' ? roomBelow : roomAbove, POPOVER_MIN_HEIGHT_PX);
  const height = Math.min(size.height, maxHeight);

  const wantedTop = side === 'below' ? anchor.bottom + gap : anchor.top - gap - height;
  const top = clamp(wantedTop, margin, Math.max(margin, viewport.height - margin - height));

  // ההצמדה היא לקצה ההתחלה של הכפתור: בעברית הימני, ובלטינית השמאלי. פופאובר
  // רחב מהכפתור ליד קצה החלון היה יוצא ממנו, ולכן ההצמדה נכנעת לחלון.
  //
  // ביישור למרכז הכיווניות אינה משנה — מרכז הוא מרכז — ולכן `rtl` אינו נבדק
  // שם, וההיכנעות לחלון היא אותה היכנעות.
  const wantedLeft =
    options.align === 'center'
      ? (anchor.left + anchor.right) / 2 - size.width / 2
      : options.rtl
        ? anchor.right - size.width
        : anchor.left;
  const left = clamp(wantedLeft, margin, Math.max(margin, viewport.width - margin - size.width));

  return { top, left, maxHeight, side };
}

/**
 * הצורה שנמדדת לפני שיש מה למדוד.
 *
 * הפופאובר חייב להיות ב-DOM כדי שנדע את גודלו הטבעי, אבל אסור שייראה במקום
 * הלא נכון ולו לפריים אחד — ולכן `visibility: hidden` עד שהמדידה הסתיימה.
 */
const UNMEASURED: CSSProperties = {
  position: 'fixed',
  top: '0px',
  left: '0px',
  visibility: 'hidden',
};

export interface PopoverPosition {
  /** נקשר ב-`:style` על אלמנט הפופאובר. */
  popoverStyle: Ref<CSSProperties>;
}

/**
 * לאיזה כיוון האלמנט נמצא.
 *
 * הנפילה לתכונת `dir` אינה קוסמטית: jsdom אינו מחשב `direction` ומחזיר מחרוזת
 * ריקה מ-`getComputedStyle`, ובלעדיה כל בדיקת מיקום הייתה נמדדת בכיוון ההפוך
 * מזה שהאפליקציה רצה בו (`<html dir="rtl">` ב-index.html) — כלומר בדיקה ירוקה
 * שמאשרת יישור לקצה הלא נכון.
 */
export function isRightToLeft(element: HTMLElement): boolean {
  const computed = getComputedStyle(element).direction;
  if (computed) return computed === 'rtl';
  const declared = element.closest('[dir]')?.getAttribute('dir') ?? document.documentElement.dir;
  return declared.toLowerCase() === 'rtl';
}

/**
 * מצמידה פופאובר לכפתור שפתח אותו, בקואורדינטות חלון.
 *
 * @param anchor  המכל של הכפתור — ממנו נמדד המלבן שאליו מיישרים.
 * @param popover אלמנט הפופאובר עצמו, לאחר שנוצר.
 * @param isOpen  אותו `ref` שמפעיל את ה-`v-if`.
 */
export function usePopoverPosition(
  anchor: Ref<HTMLElement | null>,
  popover: Ref<HTMLElement | null>,
  isOpen: Ref<boolean>,
): PopoverPosition {
  const popoverStyle = ref<CSSProperties>({ ...UNMEASURED });

  /**
   * הגודל הטבעי, כפי שנמדד בפתיחה.
   *
   * הוא נשמר ואינו נמדד מחדש בכל מיקום: מרגע שהוחל `max-height`, מדידה חוזרת
   * מחזירה את הגובה **החתוך** — ואז „נכנס למטה” היה נענה בחיוב תמיד, והפופאובר
   * לא היה חוזר לגובהו המלא גם כשהתפנה מקום.
   */
  let natural: Size | null = null;

  function place(): void {
    const anchorEl = anchor.value;
    const popoverEl = popover.value;
    if (!anchorEl || !popoverEl) return;

    const measured = popoverEl.getBoundingClientRect();
    natural ??= { width: measured.width, height: measured.height };

    const placement = popoverPlacement(
      anchorEl.getBoundingClientRect(),
      natural,
      { width: window.innerWidth, height: window.innerHeight },
      { rtl: isRightToLeft(anchorEl) },
    );

    popoverStyle.value = {
      position: 'fixed',
      top: `${placement.top}px`,
      left: `${placement.left}px`,
      maxHeight: `${placement.maxHeight}px`,
    };
  }

  /**
   * גלילה **בתוך** הפופאובר אינה מזיזה אותו ביחס לכפתור, ומיקום מחדש בכל אירוע
   * כזה היה עבודה מיותרת בכל תנועת גלגלת. גלילה של כל דבר אחר — הרצועה עצמה
   * נגללת אופקית — כן מזיזה.
   */
  function onViewportChange(event?: Event): void {
    if (!isOpen.value) return;
    const target = event?.target;
    if (target instanceof Node && popover.value?.contains(target)) return;
    place();
  }

  watch(isOpen, async (open) => {
    if (!open) {
      natural = null;
      popoverStyle.value = { ...UNMEASURED };
      return;
    }
    // ה-watcher רץ לפני שהרינדור יצר את האלמנט, ולכן המצב הלא-מדוד נקבע כאן
    // וה-`nextTick` הוא מה שנותן משהו למדוד.
    popoverStyle.value = { ...UNMEASURED };
    await nextTick();
    place();
  });

  // `capture` ולא `bubble`: אירוע `scroll` של אלמנט אינו עולה למעלה, ובלעדיו
  // גלילת הרצועה עצמה לא הייתה מגיעה לכאן כלל.
  window.addEventListener('resize', onViewportChange);
  document.addEventListener('scroll', onViewportChange, true);

  onUnmounted(() => {
    window.removeEventListener('resize', onViewportChange);
    document.removeEventListener('scroll', onViewportChange, true);
  });

  return { popoverStyle };
}
