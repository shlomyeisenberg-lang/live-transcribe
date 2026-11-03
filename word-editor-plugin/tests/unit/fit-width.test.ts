/**
 * „רוחב עמוד”: החישוב שלנו מול הגיאומטריה.
 *
 * הרקע בהערת הפתיחה של engine/fit-width.ts. כאן נמדדות שלוש ההבטחות:
 * היחס עצמו (כולל גבולות), סינון מדידות שאינן מדידה, והקריאה מהמנוע —
 * שנשענת על אותו מסלול `sections.list` מסונן של ההדפסה.
 */
import { describe, expect, it } from 'vitest';
import {
  CSS_PX_PER_INCH,
  computeFitPercent,
  editorStackWidth,
  fitWidthPercent,
} from '../../src/engine/fit-width';

/** A4: 8.268 אינץ' ≈ 793.7px ב-96dpi — המידה שנמדדה ברינדור של המנוע. */
const A4_INCHES = 8.268;
const BOUNDS = { min: 50, max: 200 };

describe('computeFitPercent', () => {
  it('חלון צר מהעמוד נותן אחוז מתחת ל-100', () => {
    // 740 / (8.268 × 96) = 93.2… → עיגול של clampZoom.
    expect(computeFitPercent(740, A4_INCHES, BOUNDS)).toBe(93);
  });

  it('חלון רחב נותן הגדלה, ולא נעול ל-100%', () => {
    const max = { min: 50, max: 400 };
    expect(computeFitPercent(1480, A4_INCHES, max)).toBe(186);
  });

  it('עיגול כלפי מטה: לעולם לא עמוד רחב מהמאגס', () => {
    // 1425 / (8.268 × 96) = 179.53… — עיגול לשלם הקרוב היה נותן 180%, כלומר
    // עמוד רחב ב-4px מהמקום שיש לו וגלילה אופקית מיד אחרי הלחיצה.
    expect(computeFitPercent(1425, A4_INCHES, { min: 10, max: 500 })).toBe(179);
  });

  it('התוצאה נצמדת לגבולות שהמנוע מתיר', () => {
    // 1620px מול A4 → 204% → התקרה; 300px → 38% → הרצפה.
    expect(computeFitPercent(1620, A4_INCHES, BOUNDS)).toBe(200);
    expect(computeFitPercent(300, A4_INCHES, BOUNDS)).toBe(50);
  });

  it('מדידה שאינה מדידה מוחזרת כ-null ולא כמספר מומצא', () => {
    expect(computeFitPercent(0, A4_INCHES, BOUNDS)).toBeNull();
    expect(computeFitPercent(-5, A4_INCHES, BOUNDS)).toBeNull();
    expect(computeFitPercent(Number.NaN, A4_INCHES, BOUNDS)).toBeNull();
    expect(computeFitPercent(740, 0, BOUNDS)).toBeNull();
    expect(computeFitPercent(740, -1, BOUNDS)).toBeNull();
    expect(computeFitPercent(740, Number.NaN, BOUNDS)).toBeNull();
  });

  it('קבוע ההמרה הוא פיקסלי CSS לאינץ', () => {
    expect(CSS_PX_PER_INCH).toBe(96);
  });
});

describe('editorStackWidth', () => {
  it('מודד את המאגס שלנו ולא כל אלמנט', () => {
    const host = document.createElement('div');
    const stack = document.createElement('main');
    stack.className = 'editor-stack';
    Object.defineProperty(stack, 'clientWidth', { value: 1234 });
    const other = document.createElement('div');
    other.className = 'something-else';
    Object.defineProperty(other, 'clientWidth', { value: 999 });
    host.append(stack, other);

    expect(editorStackWidth({ querySelector: (sel: string) => host.querySelector(sel) } as unknown as Document)).toBe(1234);
  });

  it('מיכל הגלילה קודם ל-<main>: פס הגלילה גורע רק ממנו', () => {
    // בלעדיו „רוחב עמוד” מכוון לעמוד רחב ב-15px מהמקום שיש לו, ומייצר בדיוק
    // את הגלילה האופקית שהחישוב בא למנוע.
    const root = document.createElement('div');
    const stack = document.createElement('main');
    stack.className = 'editor-stack';
    Object.defineProperty(stack, 'clientWidth', { value: 1440 });
    const scroller = document.createElement('div');
    scroller.className = 'editor-stack__host';
    Object.defineProperty(scroller, 'clientWidth', { value: 1425 });
    stack.append(scroller);
    root.append(stack);

    expect(editorStackWidth({ querySelector: (sel: string) => root.querySelector(sel) } as unknown as Document)).toBe(1425);
  });

  it('בלי מאגס, או עם רוחב אפס — מוחזר 0 ולא נזרקת חריגה', () => {
    expect(editorStackWidth(document)).toBe(0);
    const host = document.createElement('div');
    expect(editorStackWidth({ querySelector: (sel: string) => host.querySelector(sel) } as unknown as Document)).toBe(0);
  });
});

describe('fitWidthPercent', () => {
  function hostWith(width: number | undefined) {
    return {
      activeEditor: {
        doc: {
          sections: {
            list: async () => ({
              items: [{ pageSetup: width === undefined ? {} : { width, height: 11.694 } }],
            }),
          },
        },
      },
    };
  }

  it('קורא את רוחב העמוד מ-sections.list ומחשב את היחס', async () => {
    await expect(fitWidthPercent(hostWith(A4_INCHES), 740, BOUNDS)).resolves.toBe(93);
  });

  it('רוחב שאינו תשובה (twips גולמיים, למשל) מוחזר כ-null', async () => {
    // 11906 "אינץ'" הוא מעבר לגבול השפיות — כלומר יחידות שנשכחו בדרך.
    await expect(fitWidthPercent(hostWith(11906), 740, BOUNDS)).resolves.toBeNull();
  });

  it('כשל קריאת המקטעים הוא null ולא חריגה', async () => {
    const broken = {
      activeEditor: {
        doc: {
          sections: {
            list: async () => {
              throw new Error('boom');
            },
          },
        },
      },
    };
    await expect(fitWidthPercent(broken, 740, BOUNDS)).resolves.toBeNull();
  });

  it('בלי מאגס אין חישוב בכלל — ולא קריאת מקטעים סרק', async () => {
    let called = 0;
    const counting = {
      activeEditor: {
        doc: {
          sections: {
            list: async () => {
              called += 1;
              return { items: [{ pageSetup: { width: A4_INCHES } }] };
            },
          },
        },
      },
    };
    await expect(fitWidthPercent(counting, 0, BOUNDS)).resolves.toBeNull();
    expect(called).toBe(0);
  });
});