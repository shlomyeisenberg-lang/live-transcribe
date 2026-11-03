/**
 * engine/line-number-layer.ts — ההמרה הטהורה בין מלבנים גולמיים
 * (`RawTextRect`, מ-`Range.getClientRects()` על עמוד שלם — engine/page-ruler.ts)
 * ומה שנקרא מהמסמך (`LineNumberingReading`, engine/page-setup.ts) לבין רשימת
 * (מיקום, מספר) לציור. אין כאן DOM — כל הבדיקות הן על נתונים רגילים.
 *
 * שלוש המשפחות שנבדקות, תואמות לשלושת השלבים בהערת הפתיחה של המודול:
 *   1. **`groupLinesFromRects`** — סינון קונטיינרים וקיבוץ ריצות-טקסט לשורה.
 *   2. **`filterBodyLines`**, `bodyBandPx`, `marginGutterPx` — גיאומטריית העמוד.
 *   3. **`buildLineNumberBoxes`** — ההרכבה השלמה, כולל countBy/start/restart.
 */
import { describe, expect, it } from 'vitest';
import {
  bodyBandPx,
  buildLineNumberBoxes,
  filterBodyLines,
  groupLinesFromRects,
  marginGutterPx,
  type LineBox,
  type PageGeometryTwips,
  type PageLineSource,
} from '../../src/engine/line-number-layer';
import type { IndexedPageRect, RawTextRect } from '../../src/engine/page-ruler';
import type { LineNumberingReading } from '../../src/engine/page-setup';

function rect(topPx: number, heightPx: number, leftPx = 500, widthPx = 100): RawTextRect {
  return { topPx, heightPx, leftPx, widthPx };
}

/* ------------------------------------------------------------------ */
/* groupLinesFromRects                                                 */
/* ------------------------------------------------------------------ */

describe('groupLinesFromRects', () => {
  it('בלי קלט — בלי שורות', () => {
    expect(groupLinesFromRects([])).toEqual([]);
  });

  it('מלבן יחיד בעל שטח — שורה אחת', () => {
    expect(groupLinesFromRects([rect(100, 17)])).toEqual([{ topPx: 100, heightPx: 17, leftPx: 500, rightPx: 600 }]);
  });

  it('מלבנים ברוחב/גובה 0 מסוננים — קונטיינרים ריקים אינם שורה', () => {
    const rects = [rect(100, 17), { topPx: 50, heightPx: 0, leftPx: 0, widthPx: 100 }, { topPx: 50, heightPx: 20, leftPx: 0, widthPx: 0 }];
    expect(groupLinesFromRects(rects)).toHaveLength(1);
  });

  it('שתי ריצות-טקסט על אותה שורה (`top` זהה) מתמזגות למלבן אחד', () => {
    // בדיוק המקרה שנמדד: RTL עם כמה `span` על אותה שורה, לא בהכרח ברצף שמאל-ימין.
    const groups = groupLinesFromRects([rect(200, 17, 600, 50), rect(200, 17, 500, 90)]);
    expect(groups).toEqual([{ topPx: 200, heightPx: 17, leftPx: 500, rightPx: 650 }]);
  });

  it('הפרש `top` זניח (תת-פיקסל) עדיין נחשב אותה שורה', () => {
    const groups = groupLinesFromRects([rect(200, 17), rect(200.3, 17, 700)]);
    expect(groups).toHaveLength(1);
  });

  it('הפרש `top` אמיתי — שתי שורות נפרדות', () => {
    const groups = groupLinesFromRects([rect(100, 17), rect(120, 17)]);
    expect(groups).toHaveLength(2);
  });

  it('מלבן-קונטיינר (פסקה שלמה, גבוה משמעותית מהחציון) מסונן ואינו שורה בפני עצמו', () => {
    // בדיוק מה שנמדד: מלבן בגובה 552 לצד עשרות מלבני שורה בגובה 17.
    const rects = [rect(115, 552, 499, 601), ...Array.from({ length: 10 }, (_, i) => rect(115 + i * 18, 17))];
    const groups = groupLinesFromRects(rects);
    expect(groups).toHaveLength(10);
    expect(groups.every((g) => g.heightPx === 17)).toBe(true);
  });

  it('קונטיינר שחולק `top` עם השורה הראשונה שלו — השורה לא נבלעת בקונטיינר', () => {
    // המקרה העדין שנמדד: שורה ראשונה של פסקה מתחילה **באותו** top כמו
    // מלבן-הקונטיינר של הפסקה כולה. בלי הסינון-לפי-גובה, הקיבוץ לפי top
    // היה ממזג את שניהם ומחזיר גובה 552 לשורה הראשונה.
    const groups = groupLinesFromRects([rect(115, 552, 499, 601), rect(115, 17, 519, 4), rect(115, 17, 523, 577)]);
    expect(groups).toEqual([{ topPx: 115, heightPx: 17, leftPx: 519, rightPx: 1100 }]);
  });

  it('לא רגיש לסדר הקלט — התוצאה ממוינת ומאוחדת בכל סדר', () => {
    const a = groupLinesFromRects([rect(300, 17), rect(100, 17), rect(200, 17)]);
    const b = groupLinesFromRects([rect(200, 17), rect(300, 17), rect(100, 17)]);
    expect(a.map((g) => g.topPx)).toEqual([100, 200, 300]);
    expect(b.map((g) => g.topPx)).toEqual([100, 200, 300]);
  });
});

/* ------------------------------------------------------------------ */
/* bodyBandPx / filterBodyLines / marginGutterPx                       */
/* ------------------------------------------------------------------ */

const PAGE: IndexedPageRect = { pageIndex: 0, leftPx: 100, topPx: 50, widthPx: 800, heightPx: 1200 };

const GEOMETRY: PageGeometryTwips = {
  pageWidthTwips: 12240, // Letter
  pageHeightTwips: 15840,
  leftTwips: 1440, // 1"
  rightTwips: 1440,
  effectiveTopTwips: 1440,
  effectiveBottomTwips: 1440,
};

describe('bodyBandPx', () => {
  it('שוליים סימטריים של 1/11 מהגובה — פס הגוף מתחיל ומסתיים באותו יחס', () => {
    const band = bodyBandPx(PAGE, GEOMETRY);
    const fraction = 1440 / 15840;
    expect(band.topPx).toBeCloseTo(PAGE.topPx + fraction * PAGE.heightPx, 5);
    expect(band.bottomPx).toBeCloseTo(PAGE.topPx + PAGE.heightPx - fraction * PAGE.heightPx, 5);
  });

  it('כותרת עליונה שהרימה את השוליים האפקטיביים — פס הגוף מתחיל יותר נמוך', () => {
    // effectiveTopTwips גבוה מהשוליים הרגילים — בדיוק מה ש-readPageMargins
    // מחזיר כשכותרת עליונה דוחקת את הטקסט מטה (page-setup.ts, readEffectiveMargins).
    const withHeader: PageGeometryTwips = { ...GEOMETRY, effectiveTopTwips: 2000 };
    const band = bodyBandPx(PAGE, withHeader);
    const normal = bodyBandPx(PAGE, GEOMETRY);
    expect(band.topPx).toBeGreaterThan(normal.topPx);
  });

  it('`pageHeightTwips` לא תקין — נופל לכל גובה העמוד (בלי חלוקה באפס)', () => {
    const band = bodyBandPx(PAGE, { ...GEOMETRY, pageHeightTwips: 0 });
    expect(band).toEqual({ topPx: PAGE.topPx, bottomPx: PAGE.topPx + PAGE.heightPx });
  });
});

describe('filterBodyLines', () => {
  it('שורה שמרכזה בתוך הפס נשארת, כותרת/שוליים מחוץ לו מוסרים', () => {
    const band = bodyBandPx(PAGE, GEOMETRY); // ~159px מלמעלה ומלמטה (1"/11 * 1200)
    const headerLine: LineBox = { topPx: PAGE.topPx + 10, heightPx: 18, leftPx: 0, rightPx: 100 };
    const bodyLine: LineBox = { topPx: band.topPx + 20, heightPx: 18, leftPx: 0, rightPx: 100 };
    const footerLine: LineBox = { topPx: band.bottomPx + 5, heightPx: 18, leftPx: 0, rightPx: 100 };

    const kept = filterBodyLines([headerLine, bodyLine, footerLine], PAGE, GEOMETRY);
    expect(kept).toEqual([bodyLine]);
  });
});

describe('marginGutterPx', () => {
  const margins = { pageWidthTwips: 12240, leftTwips: 1440, rightTwips: 1440 };

  it('צד שמאל — הרוחב תואם ליחס leftTwips/pageWidthTwips, מוצמד לקצה השמאלי', () => {
    const gutter = marginGutterPx(PAGE, margins, 'left');
    expect(gutter.leftPx).toBe(PAGE.leftPx);
    expect(gutter.widthPx).toBeCloseTo((1440 / 12240) * PAGE.widthPx, 5);
  });

  it('צד ימין — מוצמד לקצה הימני של העמוד, לא לקצה השמאלי', () => {
    const gutter = marginGutterPx(PAGE, margins, 'right');
    expect(gutter.leftPx).toBeCloseTo(PAGE.leftPx + PAGE.widthPx - gutter.widthPx, 5);
  });

  it('`pageWidthTwips` לא תקין — פס ברוחב 0, בלי חלוקה באפס', () => {
    expect(marginGutterPx(PAGE, { ...margins, pageWidthTwips: 0 }, 'left')).toEqual({
      leftPx: PAGE.leftPx,
      widthPx: 0,
    });
  });
});

/* ------------------------------------------------------------------ */
/* buildLineNumberBoxes — ההרכבה השלמה                                 */
/* ------------------------------------------------------------------ */

/** שורה בודדת בתוך פס גוף הטקסט, ברוחב/גובה קבועים — נוח לבנות ריצת שורות. */
function bodyLineRect(index: number): RawTextRect {
  const band = bodyBandPx(PAGE, GEOMETRY);
  return rect(band.topPx + 5 + index * 20, 17);
}

function reading(overrides: Partial<LineNumberingReading> = {}): LineNumberingReading {
  return { countBy: 1, start: 1, restart: 'continuous', page: GEOMETRY as never, ...overrides };
}

describe('buildLineNumberBoxes', () => {
  it('`reading` null — בלי תוויות, גם כשיש שורות ועמודים', () => {
    const sources: PageLineSource[] = [{ pageIndex: 0, rects: [bodyLineRect(0)] }];
    expect(buildLineNumberBoxes(sources, [PAGE], null)).toEqual([]);
  });

  it('בלי עמודים, או בלי מקורות — בלי תוויות', () => {
    expect(buildLineNumberBoxes([], [PAGE], reading())).toEqual([]);
    expect(buildLineNumberBoxes([{ pageIndex: 0, rects: [bodyLineRect(0)] }], [], reading())).toEqual([]);
  });

  it('`countBy: 1` — כל שורה מקבלת מספר, ברצף מ-`start`', () => {
    const rects = [bodyLineRect(0), bodyLineRect(1), bodyLineRect(2)];
    const out = buildLineNumberBoxes([{ pageIndex: 0, rects }], [PAGE], reading({ start: 1 }));
    expect(out.map((b) => b.value)).toEqual([1, 2, 3]);
  });

  it('`start` שונה מ-1 — הרצף מתחיל ממנו, לא מ-1', () => {
    const rects = [bodyLineRect(0), bodyLineRect(1)];
    const out = buildLineNumberBoxes([{ pageIndex: 0, rects }], [PAGE], reading({ start: 5 }));
    expect(out.map((b) => b.value)).toEqual([5, 6]);
  });

  it('`countBy: 5` — רק כפולות מוצגות, שאר השורות עדיין נספרות', () => {
    const rects = Array.from({ length: 12 }, (_, i) => bodyLineRect(i));
    const out = buildLineNumberBoxes([{ pageIndex: 0, rects }], [PAGE], reading({ countBy: 5, start: 1 }));
    // (n-1) % 5 === 0 → 1, 6, 11
    expect(out.map((b) => b.value)).toEqual([1, 6, 11]);
  });

  it('`restart: "newPage"` — המונה מתאפס ל-start בתחילת כל עמוד', () => {
    const page1: IndexedPageRect = { ...PAGE, pageIndex: 1, topPx: PAGE.topPx + PAGE.heightPx + 20 };
    const rects0 = [bodyLineRect(0), bodyLineRect(1)];
    const rects1 = [
      // שורות "עמוד 2" — אותו מבנה יחסי לתוך band של עמוד 1, בהזזה מתאימה.
      { ...bodyLineRect(0), topPx: bodyLineRect(0).topPx + page1.topPx - PAGE.topPx },
      { ...bodyLineRect(1), topPx: bodyLineRect(1).topPx + page1.topPx - PAGE.topPx },
    ];
    const out = buildLineNumberBoxes(
      [
        { pageIndex: 0, rects: rects0 },
        { pageIndex: 1, rects: rects1 },
      ],
      [PAGE, page1],
      reading({ restart: 'newPage', start: 1 }),
    );
    expect(out.filter((b) => b.pageIndex === 0).map((b) => b.value)).toEqual([1, 2]);
    expect(out.filter((b) => b.pageIndex === 1).map((b) => b.value)).toEqual([1, 2]);
  });

  it('`restart: "continuous"` — המונה ממשיך בין עמודים, לא מתאפס', () => {
    const page1: IndexedPageRect = { ...PAGE, pageIndex: 1, topPx: PAGE.topPx + PAGE.heightPx + 20 };
    const rects0 = [bodyLineRect(0), bodyLineRect(1)];
    const rects1 = [{ ...bodyLineRect(0), topPx: bodyLineRect(0).topPx + page1.topPx - PAGE.topPx }];
    const out = buildLineNumberBoxes(
      [
        { pageIndex: 0, rects: rects0 },
        { pageIndex: 1, rects: rects1 },
      ],
      [PAGE, page1],
      reading({ restart: 'continuous', start: 1 }),
    );
    expect(out.map((b) => b.value)).toEqual([1, 2, 3]);
  });

  it('`restart: "newSection"` — קירוב מכוון: מתנהג כמו „continuous” (ר׳ הערת הפתיחה)', () => {
    const page1: IndexedPageRect = { ...PAGE, pageIndex: 1, topPx: PAGE.topPx + PAGE.heightPx + 20 };
    const rects0 = [bodyLineRect(0)];
    const rects1 = [{ ...bodyLineRect(0), topPx: bodyLineRect(0).topPx + page1.topPx - PAGE.topPx }];
    const out = buildLineNumberBoxes(
      [
        { pageIndex: 0, rects: rects0 },
        { pageIndex: 1, rects: rects1 },
      ],
      [PAGE, page1],
      reading({ restart: 'newSection', start: 1 }),
    );
    expect(out.map((b) => b.value)).toEqual([1, 2]);
  });

  it('כותרת עליונה אינה מקבלת מספר, גם כשהיא נמצאת ברשימת המלבנים הגולמיים', () => {
    const band = bodyBandPx(PAGE, GEOMETRY);
    const headerRect = rect(PAGE.topPx + 5, 18); // בתוך השוליים, לא בתוך פס הגוף
    const bodyRect = rect(band.topPx + 5, 17);
    const out = buildLineNumberBoxes([{ pageIndex: 0, rects: [headerRect, bodyRect] }], [PAGE], reading());
    expect(out).toHaveLength(1);
    expect(out[0]!.topPx).toBe(bodyRect.topPx);
  });

  it('עמודים שלא נמדדו (חסרים ב-`pages`) מדולגים בלי לשבור את הרצף', () => {
    const out = buildLineNumberBoxes(
      [{ pageIndex: 99, rects: [bodyLineRect(0)] }],
      [PAGE], // אין עמוד 99 ב-pages
      reading(),
    );
    expect(out).toEqual([]);
  });

  it('סדר המקורות אינו קובע — הרצף תמיד לפי `pageIndex` עולה', () => {
    const page1: IndexedPageRect = { ...PAGE, pageIndex: 1, topPx: PAGE.topPx + PAGE.heightPx + 20 };
    const rects0 = [bodyLineRect(0)];
    const rects1 = [{ ...bodyLineRect(0), topPx: bodyLineRect(0).topPx + page1.topPx - PAGE.topPx }];
    const out = buildLineNumberBoxes(
      [
        { pageIndex: 1, rects: rects1 },
        { pageIndex: 0, rects: rects0 },
      ],
      [PAGE, page1],
      reading({ restart: 'continuous' }),
    );
    expect(out[0]!.pageIndex).toBe(0);
    expect(out[0]!.value).toBe(1);
    expect(out[1]!.pageIndex).toBe(1);
    expect(out[1]!.value).toBe(2);
  });
});

/* ------------------------------------------------------------------ */
/* פער ידוע ומתועד: טקסט בתוך תא טבלה מקבל מספר שורה משלו              */
/* ------------------------------------------------------------------ */

/**
 * **זה לא מקרה שאמור לעבור נכון — זו נעיצה (pin) של פער שנחקר ולא נסגר.**
 * ראו הערת הפתיחה של line-number-layer.ts, „פער שנחקר ולא נסגר”, לניתוח
 * המלא של למה: אין API ציבורי שממפה בלוק-מסמך למלבן-מסך, ושלושה סימנים
 * גיאומטריים חלופיים (יחס-גובה, הכלה, שני מלבנים נפרדים על אותו `top`)
 * נבדקו ונפסלו — כל אחד מסיבה שנמדדה, לא ניחוש.
 *
 * הנתונים כאן **אינם מומצאים**: הם `Range.getClientRects()` שנמדד בפועל
 * (Chrome headless, ה-`dist` הארוז) על עמוד A4 עם פסקה, אחריה טבלה 1×2
 * (תא אחד עם טקסט, אחד ריק), ואחריה עוד פסקה — ממש התרחיש שדווח. הערכים
 * הומרו ל-page-relative (הפחתת מלבן העמוד שנמדד יחד איתם: `left: 403.14,
 * top: 192, width: 793.72, height: 1122.53` — A4).
 *
 * **אם הבדיקה הזאת ניפלה כי `buildLineNumberBoxes` השתנה** — זה עשוי
 * להיות תיקון אמיתי (מישהו מצא דרך לסנן תוכן טבלה) או רגרסיה. בשני
 * המקרים: תעדכנו את הבדיקה **במודע**, ותעדכנו גם את הערת הפתיחה של
 * line-number-layer.ts — אל תשנו את המספר הצפוי כאן בלי לקרוא אותה.
 */
describe('טקסט בתוך תא טבלה — פער ידוע, לא מסונן', () => {
  const TABLE_PAGE: IndexedPageRect = { pageIndex: 0, leftPx: 0, topPx: 0, widthPx: 793.71875, heightPx: 1122.53125 };

  /** A4, שוליים רגילים (1440 טוויפס = 1" מכל צד) — בלי כותרת. */
  const TABLE_GEOMETRY: PageGeometryTwips = {
    pageWidthTwips: 11906,
    pageHeightTwips: 16838,
    leftTwips: 1440,
    rightTwips: 1440,
    effectiveTopTwips: 1440,
    effectiveBottomTwips: 1440,
  };

  // page-relative: raw.top/left פחות מלבן העמוד שנמדד יחד איתם (403.140625, 192).
  const rectsAroundTable: RawTextRect[] = [
    // "פסקה ראשונה לפני הטבלה" — שתי ריצות על אותה שורה (בידי־RTL).
    { topPx: 96, leftPx: 96, widthPx: 601.71875, heightPx: 18.390625 },
    { topPx: 96, leftPx: 530.015625, widthPx: 167.703125, heightPx: 17 },
    // מלבן-קונטיינר של שורת/תא הטבלה — גובה (19.72) קרוב מדי לגובה שורה
    // אמיתית כדי להיתפס כחריג (ראו „מה שנבדק ונדחה”, סעיף 1).
    { topPx: 114.390625, leftPx: 96, widthPx: 602, heightPx: 19.71875 },
    // תוכן התא (טקסט אמיתי) — `top` שונה ב-1px בדיוק מהקונטיינר שמעליו,
    // ולכן אינו מתמזג איתו (סף האיחוד הוא 0.75px).
    { topPx: 115.390625, leftPx: 331.515625, widthPx: 58.15625, heightPx: 17 },
    // התא השני (ריק) — סמן-סמן במיקום שונה לגמרי, אותו `top` בדיוק.
    { topPx: 115.390625, leftPx: 685.078125, widthPx: 4.453125, heightPx: 17 },
    // "פסקה אחרי הטבלה" — שוב שתי ריצות על אותה שורה.
    { topPx: 134.125, leftPx: 96, widthPx: 601.71875, heightPx: 18.390625 },
    { topPx: 134.125, leftPx: 572.5625, widthPx: 125.15625, heightPx: 17 },
  ];

  it('groupLinesFromRects סופר 4 "שורות" — 2 אמיתיות ו-2 מהטבלה', () => {
    const groups = groupLinesFromRects(rectsAroundTable);
    expect(groups.map((g) => g.topPx)).toEqual([96, 114.390625, 115.390625, 134.125]);
  });

  it('buildLineNumberBoxes ממספר את כל 4 השורות — הטבלה לא מוסרת, ודוחפת את הפסקה שאחריה', () => {
    const out = buildLineNumberBoxes(
      [{ pageIndex: 0, rects: rectsAroundTable }],
      [TABLE_PAGE],
      { countBy: 1, start: 1, restart: 'continuous', page: TABLE_GEOMETRY as never },
    );
    // הצפוי מ-Word: [1 (פסקה ראשונה), 2 (פסקה אחרי הטבלה)] — 2 בלבד.
    // מה שקורה בפועל: 4, כולל שתי "שורות" מתוך הטבלה.
    expect(out.map((b) => b.value)).toEqual([1, 2, 3, 4]);
  });
});
