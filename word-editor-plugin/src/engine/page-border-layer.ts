/**
 * גיאומטריית „גבולות עמוד” — ממיר את מה שנקרא מהמסמך (`PageBordersReading`,
 * engine/page-setup.ts) ואת מלבני העמודים המצוירים (`IndexedPageRect`,
 * engine/page-ruler.ts) לתיבות CSS מוכנות לציור. אין כאן קריאת DOM ואין
 * כתיבת DOM — רק חשבון. ui/shell/PageBorderOverlay.vue הוא היחיד שמרכיב את
 * שני המקורות (host, viewportSource) וקורא לפונקציות כאן.
 *
 * ## למה יחידה חדשה, ולא בתוך page-setup.ts או page-ruler.ts
 *
 * `page-setup.ts` קורא ומאמת מה שהמסמך מחזיק (`sections.list()`), ו-
 * `page-ruler.ts` מודד גיאומטריה גולמית מה-DOM (`data-page-index`) —
 * tests/unit/engine-boundaries.test.ts אוכף ששניהם נשארים המקום היחיד לכל
 * אחד מהעיגונים שלהם. ההמרה בין twips/eighth-points/points ל-CSS px, ובחירת
 * מה מצייר Word בכל `w:val`, אינה שייכת לאף אחד מהם — היא צריכה את שניהם
 * יחד, ולכן היא כאן.
 *
 * ## היחידות (OOXML → CSS)
 *
 *   - `w:sz` (עובי הקו) הוא `ST_EighthPointMeasure` — שמיניות נקודה.
 *     נקודה = 96/72 פיקסל, ולכן `px = eighths / 8 * 96/72 = eighths / 6`.
 *   - `w:space` (המרחק מקצה הדף, כש-`offsetFrom: 'page'`) הוא נקודות שלמות:
 *     `px = points * 96/72`.
 *   - שני אלה נמדדו ב-docs/engine-gaps.md מול ה-docx המיוצא, לא הונחו.
 *
 * ## `offsetFrom: 'text'`
 *
 * ששת פריטי התפריט (engine/page-setup.ts, `applyPageBorders`) כותבים תמיד
 * `offsetFrom: 'page'` — זו ברירת המחדל של Word לגלריית „גבולות עמוד”, וזה
 * מה ש-`docs/button-audit.md` ו-scripts/qa/layout-qa.mjs מודדים. מסמך
 * שהגיע מ-Word יכול לשאת `offsetFrom: 'text'` (מרחק משולי הטקסט ולא מקצה
 * הדף), ושכבת הציור מקרבת אותו כאילו היה `'page'` — אותו חשבון `space`,
 * בלי תוספת השוליים. זו קירוב מכוון ולא טעות: הפער היחיד הוא מסמכים
 * שנוצרו ב-Word עצמו עם הבחירה הזאת, ולא משהו שהתפריט שלנו יכול לייצר.
 */
import type { IndexedPageRect } from './page-ruler';
import type {
  PageBorderDisplay,
  PageBorderSideReading,
  PageBordersReading,
} from './page-setup';

/** נקודה = 96/72 פיקסל CSS (96 dpi, נמדד יחד עם page-setup.ts — TWIPS_PER_INCH/20). */
const PX_PER_POINT = 96 / 72;

export function pointsToPx(points: number): number {
  return points * PX_PER_POINT;
}

/** `w:sz` (`ST_EighthPointMeasure`) → פיקסלים. */
export function eighthPointsToPx(eighths: number): number {
  return (eighths / 8) * PX_PER_POINT;
}

/**
 * `w:val` של `CT_Border` → `border-style` של CSS.
 *
 * המנוע אינו מאמת `style` בכתיבה — כל מחרוזת נכתבת כמות שהיא (נמדד:
 * docs/engine-gaps.md). ששת פריטי התפריט כותבים רק `single`/`double`/
 * `dashed`/`dotted`, וכל ערך אחר (כולל ריק, וכולל סגנונות ECMA-376 שה-UI
 * שלנו אינו מציע) נופל ל-`solid` — קו בודד הוא הבחירה הכי פחות מפתיעה כשיש
 * גבול אבל אין דרך לצייר בדיוק את הסגנון שהתבקש.
 */
export function cssBorderStyle(wordStyle: string): 'solid' | 'double' | 'dashed' | 'dotted' {
  switch (wordStyle) {
    case 'double':
      return 'double';
    case 'dashed':
    case 'dashSmallGap':
    case 'dotDash':
      return 'dashed';
    case 'dotted':
    case 'dotDotDash':
      return 'dotted';
    default:
      return 'solid';
  }
}

/**
 * `w:color` (`ST_HexColor`: `auto` או שש ספרות הקסה) → צבע CSS.
 *
 * `auto` הוא מה שהתפריט שלנו כותב תמיד (`PAGE_BORDER_COLOR`), וזה גם מה
 * ש-Word עצמו כותב כברירת מחדל: „צבע הטקסט של הנושא”. `currentColor` הוא
 * התרגום המדויק ביותר בלי לקרוא את ערכת הנושא של המסמך. כל מחרוזת שאינה
 * הקסה תקין (כולל `''`, שהמנוע כותב כשקלט ריק נבלע — נמדד) נופלת לאותו מקום.
 */
export function cssBorderColor(wordColor: string): string {
  if (wordColor === 'auto') return 'currentColor';
  const hex = wordColor.replace(/^#/, '');
  return /^[0-9a-fA-F]{6}$/.test(hex) ? `#${hex}` : 'currentColor';
}

export interface CssBorderSide {
  widthPx: number;
  style: 'solid' | 'double' | 'dashed' | 'dotted';
  color: string;
  /** `w:space` המקורי, בפיקסלים — המרחק שהמסגרת שוקעת מקצה העמוד. */
  insetPx: number;
}

export function toCssBorderSide(side: PageBorderSideReading): CssBorderSide {
  return {
    widthPx: Math.max(0.5, eighthPointsToPx(side.sizeEighthPoints)),
    style: cssBorderStyle(side.style),
    color: cssBorderColor(side.color),
    insetPx: pointsToPx(side.spacePoints),
  };
}

/**
 * האם עמוד מסוים מקבל גבול, לפי `w:display`.
 *
 * „העמוד הראשון” נקבע לפי **המינימום** בין האינדקסים שנמדדו על המסך ולא לפי
 * `0`/`1` קשיח: `data-page-index` אינו חוזה מתועד (ראו page-ruler.ts), ואין
 * ערובה לבסיס המספור. מסמך שגלל חלקית ואינו מציג את העמוד הראשון בכלל פשוט
 * אינו מצייר עליו גבול באותו רגע — אין לו על מה.
 */
export function shouldDrawBorder(
  display: PageBorderDisplay,
  pageIndex: number,
  firstPageIndex: number,
): boolean {
  if (display === 'firstPage') return pageIndex === firstPageIndex;
  if (display === 'notFirstPage') return pageIndex !== firstPageIndex;
  return true;
}

/** תיבת גבול מוכנה לציור — מלבן + ארבעת הצדדים, ב-CSS. */
export interface PageBorderBox {
  pageIndex: number;
  leftPx: number;
  topPx: number;
  widthPx: number;
  heightPx: number;
  top: CssBorderSide;
  right: CssBorderSide;
  bottom: CssBorderSide;
  left: CssBorderSide;
}

/**
 * ממיר את מלבני כל העמודים ומצב „גבולות עמוד” לתיבות ציור — אחת לכל עמוד
 * שאמור לקבל גבול, בלי לגעת ב-DOM.
 *
 * המלבן של כל תיבה **שוקע** מהמלבן של העמוד לפי `space` של כל צד (ולא רק
 * צד אחד לכולם): זו בדיוק הצורה של `<w:pgBorders offsetFrom="page">` —
 * המסגרת נמדדת מקצה הנייר פנימה, וכל צד יכול לשאת `space` משלו (גם אם
 * ששת פריטי התפריט שלנו כותבים תמיד את אותו ערך לארבעתם).
 */
export function buildPageBorderBoxes(
  rects: readonly IndexedPageRect[],
  reading: PageBordersReading | null,
): PageBorderBox[] {
  if (!reading || rects.length === 0) return [];

  const firstPageIndex = rects.reduce((min, r) => Math.min(min, r.pageIndex), Infinity);
  const top = toCssBorderSide(reading.top);
  const right = toCssBorderSide(reading.right);
  const bottom = toCssBorderSide(reading.bottom);
  const left = toCssBorderSide(reading.left);

  const boxes: PageBorderBox[] = [];
  for (const rect of rects) {
    if (!shouldDrawBorder(reading.display, rect.pageIndex, firstPageIndex)) continue;
    const widthPx = rect.widthPx - left.insetPx - right.insetPx;
    const heightPx = rect.heightPx - top.insetPx - bottom.insetPx;
    if (!(widthPx > 0) || !(heightPx > 0)) continue;
    boxes.push({
      pageIndex: rect.pageIndex,
      leftPx: rect.leftPx + left.insetPx,
      topPx: rect.topPx + top.insetPx,
      widthPx,
      heightPx,
      top,
      right,
      bottom,
      left,
    });
  }
  return boxes;
}
