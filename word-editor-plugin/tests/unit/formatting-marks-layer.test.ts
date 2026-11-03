/**
 * engine/formatting-marks-layer.ts — ההתאמה הטהורה בין ריצות-טקסט (`Text
 * node`, מ-`TreeWalker`/`Range` על עמוד שלם — engine/page-ruler.ts) ובין
 * הטקסט הקנוני של כל פסקה (`doc.blocks.list({includeText:true})`) לרשימת
 * מיקומי ¶ לציור. אין כאן DOM — כל הבדיקות הן על נתונים רגילים, בדיוק כמו
 * tests/unit/line-number-layer.test.ts.
 *
 * התרחישים תואמים למה שנמדד בפועל (docs/superdoc-2.10-review.md) ולמה
 * שהתבקש: מסמך רב-פסקאות רגיל, פסקה שגולשת לכמה שורות (ריצה עם כמה
 * מלבנים), פסקה ריקה עם/בלי placeholder, טאב באמצע פסקה (מוסר לפני השוואה
 * כי הוא ממילא לא קיים בצד ה-DOM), כיוון RTL/LTR לפי הריצה בפועל ולא הנחה
 * גורפת, ובלוק לא-זכאי (כותרת/פריט-רשימה) שלא שובר את מה שאחריו.
 */
import { describe, expect, it } from 'vitest';
import {
  buildPilcrowMarks,
  RESYNC_LOOKAHEAD_RUNS,
  type FormattingMarksBlock,
  type FormattingMarksRun,
} from '../../src/engine/formatting-marks-layer';
import type { RawTextRect } from '../../src/engine/page-ruler';

function rect(topPx: number, heightPx: number, leftPx: number, widthPx: number): RawTextRect {
  return { topPx, heightPx, leftPx, widthPx };
}

function run(text: string, rects: readonly RawTextRect[], direction: 'ltr' | 'rtl' = 'ltr'): FormattingMarksRun {
  return { text, rects, direction };
}

function block(nodeId: string, text: string, opts: { nodeType?: string; isEmpty?: boolean } = {}): FormattingMarksBlock {
  return {
    nodeId,
    nodeType: opts.nodeType ?? 'paragraph',
    text,
    isEmpty: opts.isEmpty ?? text.length === 0,
  };
}

describe('buildPilcrowMarks — קלט ריק', () => {
  it('בלי בלוקים ובלי ריצות — בלי סימנים', () => {
    expect(buildPilcrowMarks([], [])).toEqual([]);
  });

  it('בלוקים בלי שום ריצה תואמת — לא נמצא, לא מנוחש', () => {
    const blocks = [block('p1', 'hello')];
    expect(buildPilcrowMarks(blocks, [])).toEqual([]);
  });
});

describe('buildPilcrowMarks — מסמך רב-פסקאות רגיל', () => {
  it('שתי פסקאות, כל אחת ריצה אחת — ¶ בסוף כל אחת', () => {
    const blocks = [block('p1', 'alpha beta'), block('p2', 'gamma')];
    const runs = [
      run('alpha beta', [rect(100, 17, 96, 140)]),
      run('gamma', [rect(120, 17, 96, 60)]),
    ];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks).toEqual([
      { nodeId: 'p1', topPx: 100, heightPx: 17, anchorXPx: 236, direction: 'ltr', approximate: false },
      { nodeId: 'p2', topPx: 120, heightPx: 17, anchorXPx: 156, direction: 'ltr', approximate: false },
    ]);
  });

  it('פסקה שמפוצלת לכמה ריצות (הדגשה/עיצוב שבור לכמה span) — נצרכות כולן, ¶ אחרי האחרונה', () => {
    const blocks = [block('p1', 'hello world')];
    const runs = [
      run('hello ', [rect(100, 17, 96, 60)]),
      run('world', [rect(100, 17, 156, 50)]),
    ];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks).toEqual([
      { nodeId: 'p1', topPx: 100, heightPx: 17, anchorXPx: 206, direction: 'ltr', approximate: false },
    ]);
  });
});

describe('buildPilcrowMarks — פסקה שגולשת לכמה שורות', () => {
  it('ריצה אחת עם כמה מלבנים (עטיפה) — העוגן הוא המלבן האחרון, לא הראשון', () => {
    const blocks = [block('p1', 'a very long paragraph that wraps to two lines')];
    const runs = [
      run('a very long paragraph that wraps to two lines', [
        rect(100, 17, 96, 500), // שורה ראשונה — נגמרת בקצה הדף
        rect(120, 17, 96, 80), // שורה שנייה — קצרה, כאן ה-¶ צריך לשבת
      ]),
    ];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks).toEqual([
      { nodeId: 'p1', topPx: 120, heightPx: 17, anchorXPx: 176, direction: 'ltr', approximate: false },
    ]);
  });
});

describe('buildPilcrowMarks — כיוון RTL/LTR לפי הריצה בפועל', () => {
  it('פסקה RTL — העוגן בקצה השמאלי של הריצה (לא הימני)', () => {
    const blocks = [block('p1', 'שלום עולם')];
    const runs = [run('שלום עולם', [rect(100, 17, 500, 140)], 'rtl')];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks).toEqual([
      { nodeId: 'p1', topPx: 100, heightPx: 17, anchorXPx: 500, direction: 'rtl', approximate: false },
    ]);
  });

  it('מסמך מעורב — כל פסקה מקבלת את הכיוון שנמדד לה, לא כיוון-מסמך גורף', () => {
    const blocks = [block('p1', 'hello'), block('p2', 'שלום')];
    const runs = [run('hello', [rect(100, 17, 96, 60)], 'ltr'), run('שלום', [rect(120, 17, 500, 60)], 'rtl')];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks.map((m) => m.direction)).toEqual(['ltr', 'rtl']);
    expect(marks[0]!.anchorXPx).toBe(156); // ltr: left+width
    expect(marks[1]!.anchorXPx).toBe(500); // rtl: left
  });
});

describe('buildPilcrowMarks — טאב באמצע פסקה', () => {
  it('טאב אינו קיים כתו ב-DOM (נמדד) — מוסר לפני השוואה, וההתאמה עדיין מוצאת את הריצה האחרונה', () => {
    const blocks = [block('p1', 'delta\tepsilon')];
    const runs = [run('delta', [rect(100, 17, 96, 40)]), run('epsilon', [rect(100, 17, 160, 60)])];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks).toEqual([
      { nodeId: 'p1', topPx: 100, heightPx: 17, anchorXPx: 220, direction: 'ltr', approximate: false },
    ]);
  });
});

describe('buildPilcrowMarks — פסקה ריקה', () => {
  it('עם placeholder (רווח בודד ב-DOM, כפי שנמדד בפועל) — עוגן אמיתי, לא נפילה-לאחור', () => {
    const blocks = [block('p1', 'before'), block('p2', '', { isEmpty: true }), block('p3', 'after')];
    const runs = [
      run('before', [rect(100, 17, 96, 60)]),
      run(' ', [rect(120, 17, 96, 4)]), // ה-placeholder שנמדד בפועל
      run('after', [rect(140, 17, 96, 60)]),
    ];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks).toEqual([
      { nodeId: 'p1', topPx: 100, heightPx: 17, anchorXPx: 156, direction: 'ltr', approximate: false },
      { nodeId: 'p2', topPx: 120, heightPx: 17, anchorXPx: 100, direction: 'ltr', approximate: false },
      { nodeId: 'p3', topPx: 140, heightPx: 17, anchorXPx: 156, direction: 'ltr', approximate: false },
    ]);
    expect(marks.every((m) => !m.approximate)).toBe(true);
  });

  it('בלי placeholder (הריצה הבאה אינה whitespace-בלבד) — נפילה-לאחור גיאומטרית, מסומנת approximate', () => {
    const blocks = [block('p1', 'before'), block('p2', '', { isEmpty: true }), block('p3', 'after')];
    const runs = [run('before', [rect(100, 17, 96, 60)]), run('after', [rect(140, 17, 96, 60)])];
    const marks = buildPilcrowMarks(blocks, runs);
    // p2 נופלת-לאחור: שורה אחת (heightPx) מתחת לעוגן המוצלח הקודם (p1).
    expect(marks[0]).toEqual({ nodeId: 'p1', topPx: 100, heightPx: 17, anchorXPx: 156, direction: 'ltr', approximate: false });
    expect(marks[1]).toEqual({ nodeId: 'p2', topPx: 117, heightPx: 17, anchorXPx: 156, direction: 'ltr', approximate: true });
    // p3 עדיין ממוקמת נכון על הריצה שלה עצמה — הנפילה-לאחור של p2 לא צרכה כלום.
    expect(marks[2]).toEqual({ nodeId: 'p3', topPx: 140, heightPx: 17, anchorXPx: 156, direction: 'ltr', approximate: false });
  });

  it('פסקה ריקה ראשונה במסמך, בלי עוגן קודם ובלי placeholder — מדולגת (פער קטן ומתועד, לא ניחוש)', () => {
    const blocks = [block('p1', '', { isEmpty: true }), block('p2', 'after')];
    const runs = [run('after', [rect(100, 17, 96, 60)])];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks).toEqual([
      { nodeId: 'p2', topPx: 100, heightPx: 17, anchorXPx: 156, direction: 'ltr', approximate: false },
    ]);
  });

  it('שתי פסקאות ריקות רצופות, שתיהן עם placeholder — כל אחת מקבלת את ה-placeholder שלה', () => {
    const blocks = [block('p1', '', { isEmpty: true }), block('p2', '', { isEmpty: true }), block('p3', 'after')];
    const runs = [run(' ', [rect(100, 17, 96, 4)]), run(' ', [rect(120, 17, 96, 4)]), run('after', [rect(140, 17, 96, 60)])];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks.map((m) => m.topPx)).toEqual([100, 120, 140]);
    expect(marks.every((m) => !m.approximate)).toBe(true);
  });
});

describe('buildPilcrowMarks — היקף: רק nodeType === paragraph', () => {
  it('כותרת/פריט-רשימה מדולגים — בלי סימון, ובלי לצרוך את הריצות שלהם', () => {
    const blocks = [block('h1', 'כותרת', { nodeType: 'heading' }), block('p1', 'גוף הפסקה')];
    const runs = [run('כותרת', [rect(100, 17, 500, 60)]), run('גוף הפסקה', [rect(120, 17, 400, 80)])];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks).toEqual([
      { nodeId: 'p1', topPx: 120, heightPx: 17, anchorXPx: 480, direction: 'ltr', approximate: false },
    ]);
  });

  it('בלוק לא-זכאי עם תוכן DOM לא-מתועד (סמן רשימה) לא שובר את הפסקה שאחריו — סנכרון מחדש', () => {
    const blocks = [block('li1', 'פריט ראשון', { nodeType: 'listItem' }), block('p1', 'פסקה רגילה')];
    const runs = [
      // בדיוק התבנית שנמדדה: סמן + רווח מפריד, שאינם בטקסט הקנוני של li1.
      run('•', [rect(100, 17, 650, 12)]),
      run(' ', [rect(100, 17, 640, 4)]),
      run('פריט ראשון', [rect(100, 17, 550, 80)]),
      run('פסקה רגילה', [rect(120, 17, 400, 90)]),
    ];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks).toEqual([
      { nodeId: 'p1', topPx: 120, heightPx: 17, anchorXPx: 490, direction: 'ltr', approximate: false },
    ]);
  });
});

describe('buildPilcrowMarks — כשל בטוח, לא ניחוש', () => {
  it('שני טאבים רצופים (ריצה חסרה גבול פנימי) — עדיין נמצא סוף הפסקה, לא מנסים למקם משהו באמצע', () => {
    const blocks = [block('p1', 'a\t\tb')];
    const runs = [run('a', [rect(100, 17, 96, 10)]), run('b', [rect(100, 17, 200, 10)])];
    const marks = buildPilcrowMarks(blocks, runs);
    expect(marks).toEqual([{ nodeId: 'p1', topPx: 100, heightPx: 17, anchorXPx: 210, direction: 'ltr', approximate: false }]);
  });

  it('אין התאמה בכלל בחלון החיפוש — הבלוק מדולג, לא זורק ולא מנחש', () => {
    const blocks = [block('p1', 'טקסט שלא קיים בכלל ב-DOM')];
    const runs = [run('משהו אחר לגמרי', [rect(100, 17, 96, 60)])];
    expect(() => buildPilcrowMarks(blocks, runs)).not.toThrow();
    expect(buildPilcrowMarks(blocks, runs)).toEqual([]);
  });

  it('חלון החיפוש חסום — התאמה שנמצאת רק מעבר ל-RESYNC_LOOKAHEAD_RUNS לא מסתנכרנת', () => {
    const filler = Array.from({ length: RESYNC_LOOKAHEAD_RUNS + 5 }, (_, i) => run(`filler${i}`, [rect(100, 17, 96, 10)]));
    const blocks = [block('p1', 'target-text')];
    const runs = [...filler, run('target-text', [rect(200, 17, 96, 60)])];
    expect(buildPilcrowMarks(blocks, runs)).toEqual([]);
  });
});
