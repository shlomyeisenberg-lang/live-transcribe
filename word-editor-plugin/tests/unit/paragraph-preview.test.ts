/**
 * הגאומטריה של פס התצוגה המקדימה ב„פסקה” — engine/paragraph-preview.ts.
 *
 * מה שנבדק כאן הוא **המכנה המשותף**: כל שש ההגדרות שהפס מצייר נמסרות ל-CSS
 * כאחוז מרוחב עמודת הטקסט או כיחס לגודל הגופן, ולכן די בקנה מידה אחד שגוי
 * כדי שהפס יראה יחסים שאינם במסמך. שלושת הכללים של `line-height` נבדקים
 * בנפרד: הם שלושה חישובים שונים, ולא שלושה מצבים של אחד.
 *
 * A4 עם שוליים של 2.54 ס״מ הוא הבסיס: 11906 twips פחות 2×1440 = 9026, כלומר
 * עמודה של 15.93 ס״מ. כל האחוזים למטה נגזרים ממנה בחשבון יד.
 */
import { describe, expect, it } from 'vitest';
import {
  PREVIEW_NEIGHBOUR_LINES,
  PREVIEW_TARGET_LINES,
  paragraphPreviewGeometry,
  type ParagraphPreviewSource,
} from '../../src/engine/paragraph-preview';
import { MIN_TEXT_WIDTH_TWIPS } from '../../src/engine/ruler-geometry';

/** A4 פחות שוליים של אינץ' לכל צד. */
const COLUMN_TWIPS = 11906 - 1440 - 1440;

function source(patch: Partial<ParagraphPreviewSource> = {}): ParagraphPreviewSource {
  return {
    textWidthTwips: COLUMN_TWIPS,
    fontSizePt: 11,
    startTwips: 0,
    endTwips: 0,
    special: 'none',
    amountTwips: 0,
    beforeTwips: 0,
    afterTwips: 0,
    lineTwips: 240,
    lineRule: 'auto',
    ...patch,
  };
}

describe('paragraphPreviewGeometry — קנה המידה', () => {
  it('הכניסות הן אחוז מרוחב עמודת הטקסט, ולא מרוחב הדף', () => {
    // 567 twips = 1 ס"מ. מול העמודה (9026) זה 6.28%; מול הדף (11906) 4.76%.
    const geometry = paragraphPreviewGeometry(source({ startTwips: 567, endTwips: 1134 }));

    expect(geometry?.startPct).toBe(6.28);
    expect(geometry?.endPct).toBe(12.56);
  });

  it('הריווח לפני ואחרי נמסר באותו מכנה — כך CSS מודד שוליים באחוזים בכל הכיוונים', () => {
    // 12 נק' = 240 twips → 2.66% מהעמודה. אותו מכנה של הכניסות בדיוק, וזה
    // מה שהופך את הפס למיניאטורה ולא לשתי סכמות זו לצד זו.
    const geometry = paragraphPreviewGeometry(source({ beforeTwips: 240, afterTwips: 120 }));

    expect(geometry?.beforePct).toBe(2.66);
    expect(geometry?.afterPct).toBe(1.33);
  });

  it('גודל הגופן ב-cqw הוא אותו מכנה — 11 נק׳ הם 2.44% מעמודה של 15.93 ס״מ', () => {
    // 11 נק' = 220 twips; 220/9026 = 2.437%.
    expect(paragraphPreviewGeometry(source())?.fontCqw).toBe(2.44);
  });

  it('גופן כפול הוא אחוז כפול — בלעדיו מרווח השורות נמדד בסרגל אחר מהכניסות', () => {
    // עמודה של 10000 twips, כדי שהיחס ייראה בלי העיגול לשתי ספרות.
    const ten = { textWidthTwips: 10000 };

    expect(paragraphPreviewGeometry(source({ ...ten, fontSizePt: 11 }))?.fontCqw).toBe(2.2);
    expect(paragraphPreviewGeometry(source({ ...ten, fontSizePt: 22 }))?.fontCqw).toBe(4.4);
  });

  it('עמודה צרה מגדילה את אותה כניסה — 1 ס״מ הוא 25% מעמודה של 4 ס״מ', () => {
    const geometry = paragraphPreviewGeometry(
      source({ textWidthTwips: 4 * 567, startTwips: 567 }),
    );

    expect(geometry?.startPct).toBe(25);
  });
});

describe('paragraphPreviewGeometry — „מיוחד”', () => {
  it('שורה ראשונה היא text-indent חיובי', () => {
    const geometry = paragraphPreviewGeometry(
      source({ special: 'firstLine', amountTwips: 567 }),
    );

    expect(geometry?.firstLinePct).toBe(6.28);
  });

  it('תלויה היא אותו מרחק בסימן הפוך — כך OOXML מגדיר `w:hanging`', () => {
    const geometry = paragraphPreviewGeometry(
      source({ special: 'hanging', amountTwips: 567 }),
    );

    expect(geometry?.firstLinePct).toBe(-6.28);
  });

  it('„ללא” מתעלם מהמידה שנשארה בשדה', () => {
    // השדה אינו מתאפס כשעוברים ל„ללא” (הוא רק מנוטרל), ומידה שהייתה נקראת
    // משם הייתה מזיזה את השורה הראשונה בפס בלי שההגדרה תחול במסמך.
    const geometry = paragraphPreviewGeometry(source({ special: 'none', amountTwips: 567 }));

    expect(geometry?.firstLinePct).toBe(0);
  });
});

describe('paragraphPreviewGeometry — שלושת כללי מרווח השורות', () => {
  it('`auto`: `w:line` הוא כפולה×240, ולכן החלוקה מדויקת ואינה תלויה בגופן', () => {
    expect(paragraphPreviewGeometry(source({ lineTwips: 240 }))?.lineHeight).toBe(1);
    expect(paragraphPreviewGeometry(source({ lineTwips: 360 }))?.lineHeight).toBe(1.5);
    expect(paragraphPreviewGeometry(source({ lineTwips: 480 }))?.lineHeight).toBe(2);
    // וגופן אחר אינו משנה דבר בכלל הזה.
    expect(
      paragraphPreviewGeometry(source({ lineTwips: 360, fontSizePt: 22 }))?.lineHeight,
    ).toBe(1.5);
  });

  it('`exact`: הגובה הוא היחס למרחק הגופן — 22 נק׳ על גופן 11 הם פי שניים', () => {
    const geometry = paragraphPreviewGeometry(
      source({ lineRule: 'exact', lineTwips: 22 * 20 }),
    );

    expect(geometry?.lineHeight).toBe(2);
  });

  it('`exact` מתחת לגופן מצטמצם באמת — זה מה ש-Word עושה', () => {
    // 6 נק' על גופן 11: השורות נחתכות זו על זו, וזו ההתנהגות במסמך.
    const geometry = paragraphPreviewGeometry(
      source({ lineRule: 'exact', lineTwips: 6 * 20 }),
    );

    expect(geometry?.lineHeight).toBe(0.55);
  });

  it('`atLeast` הוא רצפה: מעל תיבת ה-em הוא כמו `exact`, ומתחתיה נעצר בה', () => {
    expect(
      paragraphPreviewGeometry(source({ lineRule: 'atLeast', lineTwips: 22 * 20 }))?.lineHeight,
    ).toBe(2);
    // ‏6 נק' על גופן 11 הם „לפחות 6” — כלומר השורה נשארת בגובהה הטבעי, ולא
    // מצטמצמת. ראו „מה מקורב” בהערת הראש של המודול.
    expect(
      paragraphPreviewGeometry(source({ lineRule: 'atLeast', lineTwips: 6 * 20 }))?.lineHeight,
    ).toBe(1);
  });

  it('`w:line` אפס נופל לשורה בודדת ולא ל-0 — אחרת כל הפסים מתקפלים לאחד', () => {
    expect(paragraphPreviewGeometry(source({ lineTwips: 0 }))?.lineHeight).toBe(1);
    expect(
      paragraphPreviewGeometry(source({ lineRule: 'exact', lineTwips: 0 }))?.lineHeight,
    ).toBe(1);
  });
});

describe('paragraphPreviewGeometry — מה שאין לו קנה מידה', () => {
  it('בלי רוחב עמודה אין פס בכלל', () => {
    expect(paragraphPreviewGeometry(source({ textWidthTwips: 0 }))).toBeNull();
    expect(paragraphPreviewGeometry(source({ textWidthTwips: -1 }))).toBeNull();
    expect(paragraphPreviewGeometry(source({ textWidthTwips: Number.NaN }))).toBeNull();
  });

  it('בלי גודל גופן אין פס — לציר האנכי אין סרגל', () => {
    expect(paragraphPreviewGeometry(source({ fontSizePt: 0 }))).toBeNull();
    expect(paragraphPreviewGeometry(source({ fontSizePt: Number.NaN }))).toBeNull();
  });
});

describe('paragraphPreviewGeometry — כניסות שאינן משאירות מקום', () => {
  it('הרצפה היא זו של הסרגל, ולא מספר חדש', () => {
    // בדיוק על הרצפה — עובר.
    const onFloor = paragraphPreviewGeometry(
      source({ startTwips: COLUMN_TWIPS - MIN_TEXT_WIDTH_TWIPS, endTwips: 0 }),
    );
    expect(onFloor?.leavesRoom).toBe(true);

    // twip אחד מתחתיה — נופל.
    const belowFloor = paragraphPreviewGeometry(
      source({ startTwips: COLUMN_TWIPS - MIN_TEXT_WIDTH_TWIPS + 1, endTwips: 0 }),
    );
    expect(belowFloor?.leavesRoom).toBe(false);
  });

  it('שתי הכניסות נמדדות יחד — כל אחת לבדה חוקית', () => {
    // 25 ס"מ כל אחת: הדיאלוג מתיר 55.87 לכל שדה בנפרד, ואינו מצליב ביניהם.
    const geometry = paragraphPreviewGeometry(
      source({ startTwips: 25 * 567, endTwips: 25 * 567 }),
    );

    expect(geometry).not.toBeNull();
    expect(geometry?.leavesRoom).toBe(false);
  });
});

describe('מספר השורות שהפס מצייר', () => {
  it('לפסקה שנערכת יש יותר מרחקים פנימיים מאשר חיצוניים', () => {
    // זה כל הנימוק למספר: עם שתי שורות יש מרחק פנימי אחד, וריווח „לפני”
    // נראה בדיוק כמו מרווח שורות. ראו PREVIEW_TARGET_LINES.
    expect(PREVIEW_TARGET_LINES - 1).toBeGreaterThan(2);
    expect(PREVIEW_NEIGHBOUR_LINES).toBeGreaterThan(0);
  });
});
