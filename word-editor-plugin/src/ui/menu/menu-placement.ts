/**
 * לאן נפתח תפריט ההקשר, בהינתן הנקודה שנלחצה.
 *
 * ## למה לא `usePopoverPosition` כמו שהוא
 *
 * הפופאוברים של הרצועה נפתחים מ**כפתור**: יש להם מלבן עוגן, והם מיושרים לקצה
 * שלו. תפריט הקשר נפתח מ**נקודה**, ושני הבדלים נובעים מזה:
 *
 * 1. **אין קצה להיצמד אליו, ולכן יש כיוון להתהפך אליו.** `popoverPlacement`
 *    ב-RTL מיישר לקצה הימני של העוגן ונכנע לחלון בהצמדה בלבד. עם עוגן-נקודה
 *    זה שובר: לחיצה ב-x=20 עם כרטיס בן 264px נותנת `left = -244`, ההצמדה
 *    מרימה אותו ל-8, והתפריט נפתח **מתחת לסמן** — כלומר מכסה בדיוק את מה
 *    שנלחץ. בעברית לחיצה בשוליים השמאליים אינה מקרה קצה, היא סוף כל שורה.
 *    לכן כאן יש היפוך אופקי אמיתי: לא נכנס מימין לסמן — נפתח משמאלו.
 * 2. **המרווח הוא אפס.** 2px בין כפתור לפופאובר הם הפרדה נכונה; בין קצה
 *    הסמן לפינת התפריט הם רק סטייה.
 *
 * הציר האנכי זהה לחלוטין — היפוך למעלה כשאין מקום למטה, הצמדה לחלון והגבלת
 * גובה — ולכן הוא **נלקח** מ-`popoverPlacement` ואינו משוכפל כאן.
 */
import {
  clamp,
  POPOVER_MARGIN_PX,
  popoverPlacement,
  type Placement,
  type Size,
} from '../../composables/popover-position';

export interface MenuPoint {
  readonly x: number;
  readonly y: number;
}

export interface MenuPlacementOptions {
  /** ברירת המחדל היא ימין-לשמאל: זו הכיווניות של הממשק. */
  readonly rtl?: boolean;
  readonly margin?: number;
}

/**
 * `top` ו-`maxHeight` מגיעים מהחישוב המשותף; `left` נקבע כאן.
 *
 * הסדר הוא: הצד המועדף (בעברית — התפריט משמאל לסמן, כלומר קצהו הימני נוגע
 * בו), ואם אינו נכנס — הצד השני. רק אם **שניהם** אינם נכנסים, כלומר התפריט
 * רחב מהחלון, נופלים להצמדה.
 */
export function contextMenuPlacement(
  point: MenuPoint,
  size: Size,
  viewport: Size,
  options: MenuPlacementOptions = {},
): Placement {
  const rtl = options.rtl ?? true;
  const margin = options.margin ?? POPOVER_MARGIN_PX;

  const vertical = popoverPlacement(
    { top: point.y, bottom: point.y, left: point.x, right: point.x },
    size,
    viewport,
    { rtl, gap: 0, margin },
  );

  const preferred = rtl ? point.x - size.width : point.x;
  const flipped = rtl ? point.x : point.x - size.width;
  const fits = (left: number): boolean =>
    left >= margin && left + size.width <= viewport.width - margin;

  const left = fits(preferred)
    ? preferred
    : fits(flipped)
      ? flipped
      : clamp(preferred, margin, Math.max(margin, viewport.width - margin - size.width));

  return { top: vertical.top, left, maxHeight: vertical.maxHeight, side: vertical.side };
}
