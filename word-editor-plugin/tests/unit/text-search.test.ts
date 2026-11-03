/**
 * בדיקות הלוגיקה הטהורה: מציאת מופעים בטקסט הקנוני של בלוקים, ניווט
 * בין מופעים, ומיפוי מופע ← יעד-מנוע. זו בדיוק הלוגיקה שעוקפת את ה-
 * projection השבור של `ui.search` (ראו הראש של engine/text-search.ts) —
 * ולכן הבדיקה המרכזית כאן היא בדיוק התרחיש שהיה שבור: מסמך רב-פסקאות
 * שבו לכל פסקה מופע אחד, וסך הכול נמצאים כולם.
 */
import { describe, it, expect } from 'vitest';
import {
  findAllMatches,
  matchToTarget,
  advanceActiveIndex,
  activeIndexAfterReplace,
  matchesForReplacement,
  groupMatchesByBlock,
  type SearchableBlock,
  type TextMatch,
} from '../../src/engine/text-search';

describe('findAllMatches', () => {
  it('מוצאת מופע אחד בכל אחת משמונה פסקאות — התרחיש שהיה שבור', () => {
    // המנוע (ui.search) מדד כאן 4 מתוך 8 בגלל projection-incomplete.
    // הלוגיקה כאן אינה תלויה בו כלל, ולכן חייבת למצוא את כל השמונה.
    const blocks: SearchableBlock[] = Array.from({ length: 8 }, (_, i) => ({
      blockId: `p${i}`,
      text: `zzq paragraph ${i}`,
    }));

    const matches = findAllMatches(blocks, 'zzq');

    expect(matches).toHaveLength(8);
    expect(matches.map((m) => m.blockId)).toEqual(blocks.map((b) => b.blockId));
    for (const match of matches) {
      expect(match.start).toBe(0);
      expect(match.end).toBe(3);
    }
  });

  it('מוצאת כמה מופעים לא-חופפים באותו בלוק', () => {
    const blocks: SearchableBlock[] = [{ blockId: 'p0', text: 'cat cat cat' }];
    const matches = findAllMatches(blocks, 'cat');
    expect(matches).toEqual([
      { blockId: 'p0', start: 0, end: 3 },
      { blockId: 'p0', start: 4, end: 7 },
      { blockId: 'p0', start: 8, end: 11 },
    ]);
  });

  it('מופעים חופפים אינם נספרים פעמיים — "aaa" בתוך "aaaa" נותן שניים לא שלושה', () => {
    const blocks: SearchableBlock[] = [{ blockId: 'p0', text: 'aaaa' }];
    const matches = findAllMatches(blocks, 'aa');
    expect(matches).toEqual([
      { blockId: 'p0', start: 0, end: 2 },
      { blockId: 'p0', start: 2, end: 4 },
    ]);
  });

  it('ברירת המחדל אינה רגישה לרישיות — תואם את ברירת המחדל שנמדדה במנוע', () => {
    const blocks: SearchableBlock[] = [{ blockId: 'p0', text: 'Hello hello HELLO' }];
    expect(findAllMatches(blocks, 'hello')).toHaveLength(3);
  });

  it('caseSensitive:true מגביל להתאמה מדויקת', () => {
    const blocks: SearchableBlock[] = [{ blockId: 'p0', text: 'Hello hello HELLO' }];
    const matches = findAllMatches(blocks, 'hello', { caseSensitive: true });
    expect(matches).toEqual([{ blockId: 'p0', start: 6, end: 11 }]);
  });

  it('שאילתה ריקה מחזירה מערך ריק, לא כשל', () => {
    expect(findAllMatches([{ blockId: 'p0', text: 'טקסט' }], '')).toEqual([]);
  });

  it('רשימת בלוקים ריקה מחזירה מערך ריק', () => {
    expect(findAllMatches([], 'שאילתה')).toEqual([]);
  });

  it('בלוק בלי התאמה נעדר מהתוצאה, והבא אחריו כן נמצא', () => {
    const blocks: SearchableBlock[] = [
      { blockId: 'p0', text: 'אין כלום פה' },
      { blockId: 'p1', text: 'zzq נמצא כאן' },
    ];
    const matches = findAllMatches(blocks, 'zzq');
    expect(matches).toEqual([{ blockId: 'p1', start: 0, end: 3 }]);
  });
});

describe('matchToTarget', () => {
  it('בונה SelectionTarget מנקודת-טקסט לנקודת-טקסט, באותו בלוק', () => {
    const match: TextMatch = { blockId: 'p3', start: 5, end: 9 };
    expect(matchToTarget(match)).toEqual({
      kind: 'selection',
      start: { kind: 'text', blockId: 'p3', offset: 5 },
      end: { kind: 'text', blockId: 'p3', offset: 9 },
    });
  });
});

describe('advanceActiveIndex', () => {
  it('total=0 תמיד מחזיר -1', () => {
    expect(advanceActiveIndex(-1, 0, 'next')).toBe(-1);
    expect(advanceActiveIndex(3, 0, 'prev')).toBe(-1);
  });

  it('אין התאמה פעילה, next → הראשונה', () => {
    expect(advanceActiveIndex(-1, 5, 'next')).toBe(0);
  });

  it('אין התאמה פעילה, prev → האחרונה', () => {
    expect(advanceActiveIndex(-1, 5, 'prev')).toBe(4);
  });

  it('next מתקדם ועוטף מהאחרונה לראשונה', () => {
    expect(advanceActiveIndex(0, 3, 'next')).toBe(1);
    expect(advanceActiveIndex(2, 3, 'next')).toBe(0);
  });

  it('prev נסוג ועוטף מהראשונה לאחרונה', () => {
    expect(advanceActiveIndex(1, 3, 'prev')).toBe(0);
    expect(advanceActiveIndex(0, 3, 'prev')).toBe(2);
  });

  it('אינדקס מחוץ לתחום מטופל כמו "אין התאמה פעילה"', () => {
    expect(advanceActiveIndex(9, 3, 'next')).toBe(0);
  });
});

describe('activeIndexAfterReplace', () => {
  it('הרשימה החדשה ריקה → -1', () => {
    expect(activeIndexAfterReplace(2, 0)).toBe(-1);
  });

  it('האינדקס הקודם עדיין בתחום → נשאר — זו ההתקדמות ל"הבא" בחינם', () => {
    expect(activeIndexAfterReplace(3, 7)).toBe(3);
  });

  it('הוחלף המופע האחרון → נצמד לאחרון החדש', () => {
    expect(activeIndexAfterReplace(7, 7)).toBe(6);
  });

  it('אינדקס שלילי נצמד לאפס', () => {
    expect(activeIndexAfterReplace(-1, 4)).toBe(0);
  });
});

describe('matchesForReplacement', () => {
  it('בתוך בלוק אחד — מהאחרון לראשון', () => {
    const matches: TextMatch[] = [
      { blockId: 'p0', start: 0, end: 3 },
      { blockId: 'p0', start: 4, end: 7 },
      { blockId: 'p0', start: 8, end: 11 },
    ];
    expect(matchesForReplacement(matches)).toEqual([
      { blockId: 'p0', start: 8, end: 11 },
      { blockId: 'p0', start: 4, end: 7 },
      { blockId: 'p0', start: 0, end: 3 },
    ]);
  });

  it('בין בלוקים — סדר ההופעה הראשונה, וכל בלוק פנימית מהסוף להתחלה', () => {
    const matches: TextMatch[] = [
      { blockId: 'p0', start: 0, end: 3 },
      { blockId: 'p1', start: 0, end: 3 },
      { blockId: 'p0', start: 10, end: 13 },
      { blockId: 'p1', start: 10, end: 13 },
    ];
    expect(matchesForReplacement(matches)).toEqual([
      { blockId: 'p0', start: 10, end: 13 },
      { blockId: 'p0', start: 0, end: 3 },
      { blockId: 'p1', start: 10, end: 13 },
      { blockId: 'p1', start: 0, end: 3 },
    ]);
  });

  it('שמונה מופעים, בלוק נפרד לכל אחד — כל אחד "ראשון ואחרון" בבלוק שלו', () => {
    const matches: TextMatch[] = Array.from({ length: 8 }, (_, i) => ({
      blockId: `p${i}`,
      start: 0,
      end: 3,
    }));
    expect(matchesForReplacement(matches)).toEqual(matches);
  });

  it('אינה משנה את המערך המקורי', () => {
    const matches: TextMatch[] = [
      { blockId: 'p0', start: 0, end: 3 },
      { blockId: 'p0', start: 4, end: 7 },
    ];
    const copy = [...matches];
    matchesForReplacement(matches);
    expect(matches).toEqual(copy);
  });
});

describe('groupMatchesByBlock', () => {
  it('מונה מופעים לכל בלוק, בסדר ההופעה הראשונה', () => {
    const matches: TextMatch[] = [
      { blockId: 'p0', start: 0, end: 3 },
      { blockId: 'p1', start: 0, end: 3 },
      { blockId: 'p0', start: 10, end: 13 },
      { blockId: 'p0', start: 20, end: 23 },
    ];
    expect(groupMatchesByBlock(matches)).toEqual([
      { blockId: 'p0', count: 3 },
      { blockId: 'p1', count: 1 },
    ]);
  });

  it('שמונה בלוקים נפרדים, מופע אחד בכל אחד — קבוצה לכל אחד', () => {
    const matches: TextMatch[] = Array.from({ length: 8 }, (_, i) => ({
      blockId: `p${i}`,
      start: 0,
      end: 3,
    }));
    expect(groupMatchesByBlock(matches)).toEqual(
      Array.from({ length: 8 }, (_, i) => ({ blockId: `p${i}`, count: 1 })),
    );
  });

  it('מערך ריק מחזיר רשימת קבוצות ריקה', () => {
    expect(groupMatchesByBlock([])).toEqual([]);
  });
});
