/**
 * „חלונית ניווט” — engine/navigation-pane.ts: מה נחשב כותרת, העץ, הפרק שהסמן
 * בו, והקפיצה.
 */
import { describe, expect, it, vi } from 'vitest';
import {
  goToNavHeading,
  navHasChildren,
  navHeadingLevel,
  navHidden,
  navParents,
  navSectionIndex,
  readCaretBlockId,
  readNavOutline,
  type NavigationTarget,
} from '../../src/engine/navigation-pane';

interface FakeBlock {
  nodeId: string;
  text?: string;
  styleId?: string | null;
  headingLevel?: number;
}

function fakeHost(blocks: FakeBlock[], options: { throwOnList?: boolean } = {}): NavigationTarget {
  return {
    activeEditor: {
      doc: {
        blocks: {
          list: ({ offset, limit }: { offset: number; limit: number }) => {
            if (options.throwOnList) throw new Error('engine gone');
            return { blocks: blocks.slice(offset, offset + limit) };
          },
        },
      },
    },
  } as NavigationTarget;
}

describe('navHeadingLevel', () => {
  it('headingLevel של המנוע קודם לסגנון', () => {
    expect(navHeadingLevel('Normal', 2)).toBe(2);
    expect(navHeadingLevel('Heading1', 3)).toBe(3);
  });

  it('תשע רמות, בלי תקרת ה-6 של אוצריא', () => {
    expect(navHeadingLevel('Heading8', undefined)).toBe(8);
    expect(navHeadingLevel(null, 9)).toBe(9);
    expect(navHeadingLevel(null, 12)).toBe(9);
  });

  it('שמות סגנון לא עקביים ועבריים', () => {
    expect(navHeadingLevel('heading 2', undefined)).toBe(2);
    expect(navHeadingLevel('Heading-3', undefined)).toBe(3);
    expect(navHeadingLevel('כותרת 1', undefined)).toBe(1);
    expect(navHeadingLevel('Title', undefined)).toBe(1);
  });

  it('פסקת גוף אינה כותרת', () => {
    expect(navHeadingLevel('Normal', undefined)).toBeNull();
    expect(navHeadingLevel(null, undefined)).toBeNull();
    expect(navHeadingLevel(null, 0)).toBeNull();
  });
});

describe('readNavOutline', () => {
  it('מחזירה רק כותרות לא-ריקות, עם המיקום בין כל הבלוקים', async () => {
    const outline = await readNavOutline(
      fakeHost([
        { nodeId: 'a', text: 'פתיחה', styleId: 'Normal' },
        { nodeId: 'b', text: 'פרק א', styleId: 'Heading1' },
        { nodeId: 'c', text: '', styleId: 'Heading2' },
        { nodeId: 'd', text: 'סעיף\nראשון  ', styleId: 'Heading2' },
      ]),
    );
    expect(outline?.headings).toEqual([
      { blockId: 'b', text: 'פרק א', level: 1, ordinal: 1 },
      { blockId: 'd', text: 'סעיף ראשון', level: 2, ordinal: 3 },
    ]);
    expect(outline?.ordinals.get('a')).toBe(0);
    expect(outline?.ordinals.get('c')).toBe(2);
  });

  it('מדפדפת מעבר לעמוד הראשון', async () => {
    const blocks: FakeBlock[] = Array.from({ length: 1203 }, (_, i) => ({
      nodeId: `n${i}`,
      text: `שורה ${i}`,
      styleId: i % 400 === 0 ? 'Heading1' : 'Normal',
    }));
    const outline = await readNavOutline(fakeHost(blocks));
    expect(outline?.headings.map((h) => h.ordinal)).toEqual([0, 400, 800, 1200]);
  });

  it('marks a bounded outline partial rather than hiding omitted headings', async () => {
    const blocks = Array.from({ length: 25001 }, (_, i) => ({ nodeId: `b${i}`, text: `p${i}`, styleId: i === 25000 ? 'Heading1' : 'Normal' }));
    const outline = await readNavOutline(fakeHost(blocks));
    expect(outline?.partial).toBe(true);
    expect(outline?.ordinals.size).toBe(25000);
    expect(outline?.headings).toEqual([]);
    expect((await readNavOutline(fakeHost(blocks.slice(0, 25000))))?.partial).toBe(false);
  });

  it('stops paginating as soon as the reader becomes stale', async () => {
    const list = vi.fn(async () => ({ blocks: Array.from({ length: 500 }, (_, i) => ({ nodeId: String(i) })) }));
    const host = { activeEditor: { doc: { blocks: { list } } } } as NavigationTarget;
    expect(await readNavOutline(host, () => false)).toBeNull();
    expect(list).toHaveBeenCalledTimes(1);
  });

  it('null כשאין API או שהקריאה נזרקה', async () => {
    expect(await readNavOutline(null)).toBeNull();
    expect(await readNavOutline(fakeHost([], { throwOnList: true }))).toBeNull();
  });
});

describe('העץ', () => {
  const levels = [1, 2, 3, 2, 1, 3].map((level) => ({ level }));

  it('הורה הוא הכותרת הקודמת עם רמה נמוכה יותר, גם בקפיצת רמות', () => {
    expect(navParents(levels)).toEqual([-1, 0, 1, 0, -1, 4]);
  });

  it('מי שיש לו ילדים', () => {
    expect(navHasChildren(navParents(levels))).toEqual([true, true, false, false, true, false]);
  });

  it('קיפול מסתיר את כל הצאצאים, לא רק את הילדים הישירים', () => {
    const parents = navParents(levels);
    const ids = ['a', 'b', 'c', 'd', 'e', 'f'];
    expect(navHidden(parents, new Set(['a']), ids)).toEqual([false, true, true, true, false, false]);
    expect(navHidden(parents, new Set(['b']), ids)).toEqual([false, false, true, false, false, false]);
    expect(navHidden(parents, new Set(), ids)).toEqual([false, false, false, false, false, false]);
  });
});

describe('navSectionIndex', () => {
  const headings = [{ ordinal: 2 }, { ordinal: 5 }, { ordinal: 9 }];

  it('הכותרת האחרונה שאינה אחרי הבלוק', () => {
    expect(navSectionIndex(headings, 2)).toBe(0);
    expect(navSectionIndex(headings, 4)).toBe(0);
    expect(navSectionIndex(headings, 5)).toBe(1);
    expect(navSectionIndex(headings, 100)).toBe(2);
  });

  it('-1 לפני הכותרת הראשונה או בלי כותרות', () => {
    expect(navSectionIndex(headings, 0)).toBe(-1);
    expect(navSectionIndex([], 3)).toBe(-1);
  });
});

describe('readCaretBlockId', () => {
  it('הבלוק של הקטע הראשון בבחירה', async () => {
    const host = {
      activeEditor: { doc: { selection: { current: () => ({ target: { segments: [{ blockId: 'x' }, { blockId: 'y' }] } }) } } },
    } as NavigationTarget;
    expect(await readCaretBlockId(host)).toBe('x');
  });

  it('null כשאין בחירה או API', async () => {
    expect(await readCaretBlockId(null)).toBeNull();
    const host = { activeEditor: { doc: { selection: { current: () => ({ target: null }) } } } } as NavigationTarget;
    expect(await readCaretBlockId(host)).toBeNull();
  });
});

describe('goToNavHeading', () => {
  it('מציבה סמן בתחילת הכותרת וגוללת אליה', async () => {
    const apply = vi.fn();
    const scrollIntoView = vi.fn();
    const host = { ui: { selection: { apply }, viewport: { scrollIntoView } } } as NavigationTarget;

    expect(await goToNavHeading(host, 'h1')).toBe(true);
    expect(apply).toHaveBeenCalledWith({
      kind: 'selection',
      start: { kind: 'text', blockId: 'h1', offset: 0 },
      end: { kind: 'text', blockId: 'h1', offset: 0 },
    });
    expect(scrollIntoView).toHaveBeenCalledWith(
      expect.objectContaining({ target: { kind: 'text', blockId: 'h1', range: { start: 0, end: 0 } }, block: 'start' }),
    );
  });

  it('כשל בגלילה אינו כשל — הסמן כבר זז', async () => {
    const host = {
      ui: {
        selection: { apply: vi.fn() },
        viewport: { scrollIntoView: () => Promise.reject(new Error('no layout')) },
      },
    } as NavigationTarget;
    expect(await goToNavHeading(host, 'h1')).toBe(true);
  });

  it('false כשאין API לבחירה או שהוא נזרק', async () => {
    expect(await goToNavHeading(null, 'h1')).toBe(false);
    const host = { ui: { selection: { apply: () => { throw new Error('bad target'); } } } } as NavigationTarget;
    expect(await goToNavHeading(host, 'h1')).toBe(false);
  });
});
