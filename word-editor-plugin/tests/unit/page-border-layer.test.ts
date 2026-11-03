/**
 * engine/page-border-layer.ts — ההמרה הטהורה בין מה שנקרא מהמסמך
 * (`PageBordersReading`, engine/page-setup.ts) ומלבני העמודים המצוירים
 * (`IndexedPageRect`, engine/page-ruler.ts) לבין תיבות CSS. אין כאן DOM —
 * כל הבדיקות הן על נתונים רגילים, ולכן jsdom אפילו אינו נדרש.
 *
 * שלוש המשפחות שנבדקות:
 *   1. **יחידות** — `w:sz`/`w:space` (OOXML) → פיקסלים.
 *   2. **מיפוי ל-CSS** — `w:val`/`w:color` → `border-style`/`color`, כולל
 *      נפילה אחורה על ערך שהמנוע כותב בלי ולידציה (docs/engine-gaps.md).
 *   3. **הרכבה** — `buildPageBorderBoxes`: אילו עמודים מקבלים גבול
 *      (`display`), והשוקע לכל צד.
 */
import { describe, expect, it } from 'vitest';
import {
  buildPageBorderBoxes,
  cssBorderColor,
  cssBorderStyle,
  eighthPointsToPx,
  pointsToPx,
  shouldDrawBorder,
  toCssBorderSide,
  type PageBorderBox,
} from '../../src/engine/page-border-layer';
import type { IndexedPageRect } from '../../src/engine/page-ruler';
import type { PageBorderSideReading, PageBordersReading } from '../../src/engine/page-setup';

/* ------------------------------------------------------------------ */
/* יחידות                                                              */
/* ------------------------------------------------------------------ */

describe('pointsToPx', () => {
  it('96/72 פיקסל לנקודה — `PAGE_BORDER_SPACE_POINTS` (24) הופך ל-32px', () => {
    expect(pointsToPx(24)).toBeCloseTo(32, 5);
  });

  it('0 נקודות = 0 פיקסלים', () => {
    expect(pointsToPx(0)).toBe(0);
  });
});

describe('eighthPointsToPx', () => {
  it('`w:sz="24"` (3 נקודות, „קו עבה”) → 4px בדיוק', () => {
    expect(eighthPointsToPx(24)).toBeCloseTo(4, 5);
  });

  it('`w:sz="4"` (חצי נקודה, ברירת המחדל) → 2/3 פיקסל', () => {
    expect(eighthPointsToPx(4)).toBeCloseTo(2 / 3, 5);
  });

  it('`w:sz="6"` (קו כפול) → 1px בדיוק', () => {
    expect(eighthPointsToPx(6)).toBeCloseTo(1, 5);
  });
});

/* ------------------------------------------------------------------ */
/* מיפוי ל-CSS                                                         */
/* ------------------------------------------------------------------ */

describe('cssBorderStyle', () => {
  it('ארבעת הסגנונות שנכתבים מהתפריט שלנו', () => {
    expect(cssBorderStyle('single')).toBe('solid');
    expect(cssBorderStyle('double')).toBe('double');
    expect(cssBorderStyle('dashed')).toBe('dashed');
    expect(cssBorderStyle('dotted')).toBe('dotted');
  });

  it('סגנונות ECMA-376 נוספים שה-UI שלנו אינו מציע נופלים לקבוצה הקרובה', () => {
    expect(cssBorderStyle('dashSmallGap')).toBe('dashed');
    expect(cssBorderStyle('dotDash')).toBe('dashed');
    expect(cssBorderStyle('dotDotDash')).toBe('dotted');
  });

  it('ערך שאינו סגנון Word בכלל — כולל ריק — נופל ל„solid”, לא נעלם', () => {
    // המנוע אינו מאמת `style` בכתיבה (נמדד: `'zigzag'` נכתב `w:val="zigzag"`,
    // ו-`''` מייצר גבול בלי `w:val` בכלל). שכבת הציור עדיין חייבת לצייר קו.
    expect(cssBorderStyle('zigzag')).toBe('solid');
    expect(cssBorderStyle('')).toBe('solid');
    expect(cssBorderStyle('wave')).toBe('solid');
  });
});

describe('cssBorderColor', () => {
  it('`auto` — צבע הטקסט של הנושא — הופך ל-`currentColor`', () => {
    expect(cssBorderColor('auto')).toBe('currentColor');
  });

  it('הקסה תקין, עם סולמית או בלעדיה', () => {
    expect(cssBorderColor('FF0000')).toBe('#FF0000');
    expect(cssBorderColor('#00ff00')).toBe('#00ff00');
  });

  it('כל מחרוזת שאינה `ST_HexColor` — כולל ריק — נופלת ל-`currentColor`', () => {
    expect(cssBorderColor('zigzag')).toBe('currentColor');
    expect(cssBorderColor('')).toBe('currentColor');
    expect(cssBorderColor('#FF00')).toBe('currentColor'); // פחות משש ספרות
    expect(cssBorderColor('#GGGGGG')).toBe('currentColor'); // לא הקסדצימלי
  });
});

describe('toCssBorderSide', () => {
  it('מרכיבה widthPx/style/color/insetPx יחד', () => {
    const side: PageBorderSideReading = {
      style: 'double',
      sizeEighthPoints: 6,
      spacePoints: 24,
      color: 'auto',
    };
    expect(toCssBorderSide(side)).toEqual({
      widthPx: 1,
      style: 'double',
      color: 'currentColor',
      insetPx: 32,
    });
  });

  it('קו שאמור לצאת דק מ-0.5px עדיין מצייר — לא נעלם לגמרי', () => {
    const side: PageBorderSideReading = {
      style: 'single',
      sizeEighthPoints: 1, // ~0.17px גולמי
      spacePoints: 0,
      color: 'auto',
    };
    expect(toCssBorderSide(side).widthPx).toBe(0.5);
  });
});

/* ------------------------------------------------------------------ */
/* `shouldDrawBorder` — `w:display`                                    */
/* ------------------------------------------------------------------ */

describe('shouldDrawBorder', () => {
  it('`allPages` — כל עמוד, כולל הראשון', () => {
    expect(shouldDrawBorder('allPages', 0, 0)).toBe(true);
    expect(shouldDrawBorder('allPages', 5, 0)).toBe(true);
  });

  it('`firstPage` — רק העמוד עם האינדקס המינימלי שנמדד', () => {
    expect(shouldDrawBorder('firstPage', 0, 0)).toBe(true);
    expect(shouldDrawBorder('firstPage', 1, 0)).toBe(false);
    // `data-page-index` אינו בהכרח מתחיל מ-0 (אינו חוזה מתועד) — „ראשון” הוא
    // המינימום שנמדד, לא `0` קשיח.
    expect(shouldDrawBorder('firstPage', 3, 3)).toBe(true);
  });

  it('`notFirstPage` — כל עמוד חוץ מהראשון', () => {
    expect(shouldDrawBorder('notFirstPage', 0, 0)).toBe(false);
    expect(shouldDrawBorder('notFirstPage', 1, 0)).toBe(true);
    expect(shouldDrawBorder('notFirstPage', 2, 0)).toBe(true);
  });
});

/* ------------------------------------------------------------------ */
/* `buildPageBorderBoxes` — ההרכבה השלמה                               */
/* ------------------------------------------------------------------ */

const UNIFORM_SIDE: PageBorderSideReading = {
  style: 'single',
  sizeEighthPoints: 4,
  spacePoints: 24, // → 32px, PAGE_BORDER_SPACE_POINTS
  color: 'auto',
};

const ALL_PAGES: PageBordersReading = {
  display: 'allPages',
  offsetFrom: 'page',
  top: UNIFORM_SIDE,
  right: UNIFORM_SIDE,
  bottom: UNIFORM_SIDE,
  left: UNIFORM_SIDE,
};

function rect(pageIndex: number, leftPx: number, topPx: number, widthPx = 794, heightPx = 1123): IndexedPageRect {
  return { pageIndex, leftPx, topPx, widthPx, heightPx };
}

describe('buildPageBorderBoxes', () => {
  it('בלי גבול (`null`) — אין קופסאות, גם כשיש עמודים', () => {
    expect(buildPageBorderBoxes([rect(0, 0, 0)], null)).toEqual([]);
  });

  it('בלי עמודים מצוירים — אין קופסאות, גם כשיש גבול', () => {
    expect(buildPageBorderBoxes([], ALL_PAGES)).toEqual([]);
  });

  it('עמוד יחיד: הקופסה שוקעת מכל צד לפי `space`, לא צמודה לקצה', () => {
    const boxes = buildPageBorderBoxes([rect(0, 100, 50, 794, 1123)], ALL_PAGES);

    expect(boxes).toHaveLength(1);
    expect(boxes[0]).toMatchObject({
      pageIndex: 0,
      leftPx: 100 + 32,
      topPx: 50 + 32,
      widthPx: 794 - 32 - 32,
      heightPx: 1123 - 32 - 32,
    });
  });

  it('כל ארבעת הצדדים מגיעים בתיבה, עם הסגנון/צבע/עובי הנכונים', () => {
    const boxes = buildPageBorderBoxes([rect(0, 0, 0)], ALL_PAGES);
    const box = boxes[0]!;

    for (const side of [box.top, box.right, box.bottom, box.left] as const) {
      expect(side.style).toBe('solid');
      expect(side.color).toBe('currentColor');
      expect(side.widthPx).toBeCloseTo(2 / 3, 5);
    }
  });

  it('שוקע שונה לכל צד — לא רק ערך אחיד', () => {
    const reading: PageBordersReading = {
      display: 'allPages',
      offsetFrom: 'page',
      top: { ...UNIFORM_SIDE, spacePoints: 12 }, // → 16px
      right: { ...UNIFORM_SIDE, spacePoints: 24 }, // → 32px
      bottom: { ...UNIFORM_SIDE, spacePoints: 6 }, // → 8px
      left: { ...UNIFORM_SIDE, spacePoints: 0 }, // → 0px
    };
    const box = buildPageBorderBoxes([rect(0, 100, 50, 800, 1000)], reading)[0]!;

    expect(box.topPx).toBe(50 + 16);
    expect(box.leftPx).toBe(100 + 0);
    expect(box.widthPx).toBe(800 - 0 - 32);
    expect(box.heightPx).toBe(1000 - 16 - 8);
  });

  it('`firstPage` — רק העמוד עם האינדקס המינימלי מקבל קופסה', () => {
    const reading: PageBordersReading = { ...ALL_PAGES, display: 'firstPage' };
    const boxes = buildPageBorderBoxes([rect(0, 0, 0), rect(1, 0, 1200), rect(2, 0, 2400)], reading);

    expect(boxes.map((b) => b.pageIndex)).toEqual([0]);
  });

  it('`notFirstPage` — כל עמוד חוץ מהראשון', () => {
    const reading: PageBordersReading = { ...ALL_PAGES, display: 'notFirstPage' };
    const boxes = buildPageBorderBoxes([rect(0, 0, 0), rect(1, 0, 1200), rect(2, 0, 2400)], reading);

    expect(boxes.map((b) => b.pageIndex)).toEqual([1, 2]);
  });

  it('מסמך רב-עמודי: קופסה אחת לכל עמוד, בסדר שבו הם נמדדו', () => {
    const rects = [rect(0, 0, 0), rect(1, 0, 1200), rect(2, 0, 2400)];
    const boxes = buildPageBorderBoxes(rects, ALL_PAGES);

    expect(boxes.map((b) => b.pageIndex)).toEqual([0, 1, 2]);
    expect(boxes[1]!.topPx).toBe(1200 + 32);
    expect(boxes[2]!.topPx).toBe(2400 + 32);
  });

  it('שוקע שבולע את כל העמוד מדלג על אותו עמוד, לא מייצר קופסה שלילית', () => {
    // עמוד קטן משמעותית מהשוקע ההדוק (32px מכל צד): רוחב/גובה סופיים היו
    // יוצאים שליליים, ותיבה עם רוחב שלילי מציירת גבול הפוך במקום שום גבול.
    const tiny = rect(0, 0, 0, 40, 40);
    expect(buildPageBorderBoxes([tiny], ALL_PAGES)).toEqual([]);
  });

  it('עמוד תקין לצד עמוד שנבלע — רק התקין מקבל קופסה', () => {
    const boxes = buildPageBorderBoxes([rect(0, 0, 0, 40, 40), rect(1, 0, 1200)], ALL_PAGES);
    expect(boxes.map((b) => b.pageIndex)).toEqual([1]);
  });
});

/** בדיקת-תקינות טיפוסית: `PageBorderBox` נושא בדיוק את מה שהרכיב צורך. */
describe('PageBorderBox — הצורה שהרכיב מצייר', () => {
  it('כולל את ארבעת הצדדים כ-CssBorderSide שלם', () => {
    const box: PageBorderBox = buildPageBorderBoxes([rect(0, 0, 0)], ALL_PAGES)[0]!;
    expect(Object.keys(box).sort()).toEqual(
      ['bottom', 'heightPx', 'left', 'leftPx', 'pageIndex', 'right', 'top', 'topPx', 'widthPx'].sort(),
    );
  });
});
