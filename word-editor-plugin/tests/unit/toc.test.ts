/**
 * תוכן עניינים — עדכון, הסרה, התאמה אישית וסימון ערך.
 *
 * ארבע הטענות שנמדדות כאן, וכולן מקבעות ממצא שנמדד בדפדפן ולא הנחה
 * (ההנמקה המלאה ב-engine/toc.ts):
 *
 * 1. **ההסרה נשענת על גבולות השדה של המנוע ב-2.23.** פסקאות רגילות
 *    בסגנון TOC1 חייבות לשרוד גם כשתוכן העניינים מיושן וספירת המקורות שונה.
 * 2. **הוולידציה יושבת אצלנו, כי המנוע מקבל הכול.** `level: 12` ו-
 *    `{from:9,to:1}` חוזרים ממנו עם `success: true` — הראשון כותב מתג פסול
 *    ל-Word, השני מרוקן את הטבלה. שניהם נדחים כאן לפני שנוגעים במסמך.
 * 3. **דו-משמעות מוחזרת כהודעה ולא כניחוש.** במסמך עם שתי טבלאות אין דרך
 *    ציבורית לדעת על מי הלחיצה חלה, ומימוש שבוחר את הראשונה היה מוחק את
 *    הטבלה הלא נכונה.
 * 4. **המודול לעולם אינו זורק.** חריגה, קבלה שנכשלה, ופעולה שאינה קיימת
 *    בגרסת המנוע — שלושתן `CommandOutcome`.
 */
import { describe, expect, it } from 'vitest';
import {
  TOC_PAGE_NUMBER_RESERVE_TWIPS,
  layOutTocRows,
  configureTableOfContents,
  emptyTocState,
  isValidTocLevel,
  markTocEntry,
  normalizeTocEntryText,
  normalizeTocLevels,
  readTocState,
  removeTableOfContents,
  unmarkTocEntry,
  updateTableOfContents,
  type TocHost,
} from '../../src/engine/toc';

interface Call {
  op: string;
  input?: unknown;
}

/** בלוק כפי ש-`blocks.list` מחזיר אותו. */
interface Block {
  ordinal: number;
  nodeId: string;
  nodeType: string;
  styleId?: string | null;
}

interface FakeOptions {
  /** הטבלאות שבמסמך, לפי מזהה הבלוק שלהן. */
  tocs?: readonly string[];
  /** מה שכל טבלה מצהירה עליו. נקרא רק מהראשונה. */
  sourceConfig?: unknown;
  displayConfig?: unknown;
  /** מתגים ששרדו round-trip, כפי ש-`toc.list` מחזיר אותם. `\t` הוא מה שנצרך כאן. */
  preserved?: unknown;
  /** מספר המקורות החיים — עשוי להיות שונה ממספר השורות שנשמרו. */
  entryCount?: number;
  total?: number;
  /** ערכי `TC` שסומנו במסמך. */
  entries?: readonly { nodeId: string; text?: string; level?: number }[];
  entriesTotal?: number;
  /** הבלוקים שבמסמך, כולל שורות הטבלה. */
  blocks?: readonly Block[];
  /** מידות המקטע כפי ש-`sections.list` מחזיר אותן, באינצ'ים. `null` = אין מקטעים. */
  page?: { widthInches: number; leftInches: number; rightInches: number } | null;
  /** מה `selection.current` מדווח. */
  blockId?: string | null;
  selectionText?: string;
  failures?: Record<string, { code: string; message?: string }>;
  throws?: readonly string[];
  missing?: readonly string[];
  /** כמו superdoc 2.17: `toc.remove` מסרב כשהמסמך נגמר בשורת טבלה. */
  refuseTocAtEnd?: boolean;
  /** כמו superdoc 2.23: `toc.remove` מוחק גם את השורות שאחרי הבלוק. */
  removeTakesRows?: number;
}

function fakeEngine(options: FakeOptions = {}) {
  const calls: Call[] = [];
  const missing = new Set(options.missing ?? []);
  const throwing = new Set(options.throws ?? []);
  const failures = options.failures ?? {};
  /** המסמך משתנה תוך כדי: ההסרה מוחקת בלוקים, והבדיקה בודקת את מה שנשאר. */
  let blocks: Block[] = [...(options.blocks ?? [])];

  function route<T>(op: string, impl: (input: unknown) => T): ((input: unknown) => T) | undefined {
    if (missing.has(op)) return undefined;
    return (input: unknown) => {
      calls.push({ op, input });
      if (throwing.has(op)) throw new Error('boom');
      return impl(input);
    };
  }

  const receipt = (op: string): { success: boolean; failure?: { code: string; message?: string } } =>
    failures[op] ? { success: false, failure: failures[op] } : { success: true };

  /** עמוד מתוך רשימה, בדיוק כמו `DiscoveryOutput` של המנוע. */
  const page = <T>(all: readonly T[], input: unknown, total: number | undefined) => {
    const query = (input ?? {}) as { limit?: number; offset?: number };
    const offset = query.offset ?? 0;
    const end = query.limit === undefined ? undefined : offset + query.limit;
    return { items: all.slice(offset, end), ...(total === undefined ? {} : { total }) };
  };

  const doc = {
    toc: {
      list: route('toc.list', (input) =>
        page(
          (options.tocs ?? []).map((nodeId) => ({
            address: { kind: 'block', nodeType: 'tableOfContents', nodeId },
            sourceConfig: options.sourceConfig ?? {},
            displayConfig: options.displayConfig ?? {},
            ...(options.preserved === undefined ? {} : { preserved: options.preserved }),
            ...(options.entryCount === undefined ? {} : { entryCount: options.entryCount }),
          })),
          input,
          options.total,
        ),
      ),
      update: route('toc.update', () => receipt('toc.update')),
      remove: route('toc.remove', (input) => {
        if (failures['toc.remove']) return receipt('toc.remove');
        const target = (input as { target: { nodeId: string } }).target;
        const last = blocks[blocks.length - 1];
        if (options.refuseTocAtEnd && /^TOC/.test(last?.styleId ?? '')) {
          return { success: false, failure: { code: 'CAPABILITY_UNAVAILABLE' } };
        }
        const at = blocks.findIndex((block) => block.nodeId === target.nodeId);
        if (at !== -1) blocks.splice(at, 1 + (options.removeTakesRows ?? 0));
        return receipt('toc.remove');
      }),
      configure: route('toc.configure', () => receipt('toc.configure')),
      markEntry: route('toc.markEntry', () => receipt('toc.markEntry')),
      unmarkEntry: route('toc.unmarkEntry', () => receipt('toc.unmarkEntry')),
      listEntries: route('toc.listEntries', (input) =>
        page(
          (options.entries ?? []).map((entry) => ({
            address: { kind: 'inline', nodeType: 'tableOfContentsEntry', nodeId: entry.nodeId },
            text: entry.text,
            level: entry.level,
          })),
          input,
          options.entriesTotal,
        ),
      ),
    },
    create: {
      paragraph: route('create.paragraph', (input) => {
        const target = (input as { at: { target: { nodeId: string } } }).at.target;
        const at = blocks.findIndex((block) => block.nodeId === target.nodeId);
        blocks.splice(at + 1, 0, { ordinal: at + 1, nodeId: 'trail', nodeType: 'paragraph', styleId: 'Normal' });
        return receipt('create.paragraph');
      }),
    },
    blocks: {
      list: route('blocks.list', (input) => {
        const query = (input ?? {}) as { limit?: number; offset?: number };
        const offset = query.offset ?? 0;
        const end = query.limit === undefined ? undefined : offset + query.limit;
        return { blocks: blocks.slice(offset, end), total: blocks.length };
      }),
      deleteRange: route('blocks.deleteRange', (input) => {
        const { start, end } = input as { start: { nodeId: string }; end: { nodeId: string } };
        const from = blocks.findIndex((block) => block.nodeId === start.nodeId);
        const to = blocks.findIndex((block) => block.nodeId === end.nodeId);
        if (from !== -1 && to !== -1) blocks.splice(from, to - from + 1);
        return receipt('blocks.deleteRange');
      }),
    },
    format: {
      paragraph: {
        setTabStop: route('setTabStop', () => receipt('setTabStop')),
        clearAllTabStops: route('clearAllTabStops', () => receipt('clearAllTabStops')),
        setIndentation: route('setIndentation', () => receipt('setIndentation')),
      },
    },
    sections: {
      list: route('sections.list', () => {
        const page = options.page === undefined ? A4_WITH_INCH_MARGINS : options.page;
        if (page === null) return { items: [] };
        return {
          items: [
            {
              pageSetup: { width: page.widthInches, height: 11.6929 },
              margins: {
                left: page.leftInches,
                right: page.rightInches,
                top: 1,
                bottom: 1,
              },
              sectionDirection: 'rtl',
            },
          ],
        };
      }),
    },
    selection: {
      current: route('selection.current', () => ({
        empty: true,
        target:
          options.blockId === null
            ? null
            : {
                kind: 'text',
                segments: [
                  { blockId: options.blockId ?? 'block-1', range: { start: 0, end: 0 } },
                ],
              },
        text: options.selectionText ?? '',
      })),
    },
  };

  const host = { activeEditor: { doc } } as unknown as TocHost;
  const ops = (): string[] => calls.map((call) => call.op);
  const inputs = (op: string): unknown[] =>
    calls.filter((call) => call.op === op).map((call) => call.input);

  return { host, calls, ops, inputs, remaining: () => blocks.map((block) => block.nodeId) };
}

/**
 * A4 עם שוליים של אינץ' — 11906 twips רוחב, 1440 מכל צד. באינצ'ים, מפני שזו
 * היחידה ש-`sections.list` מחזיר בה (ראו page-setup.ts).
 */
const A4_WITH_INCH_MARGINS = {
  widthInches: 11906 / 1440,
  leftInches: 1,
  rightInches: 1,
};

/** אזור הטקסט של A4 כזה, פחות המקום שנשמר למספר העמוד. */
const A4_TAB_POSITION = 11906 - 1440 - 1440 - TOC_PAGE_NUMBER_RESERVE_TWIPS;

/** כתובת טבלה, בצורה ש-`toc.list` מחזיר וש-`update`/`remove` מקבלים. */
const tocAddress = (nodeId: string) => ({
  kind: 'block',
  nodeType: 'tableOfContents',
  nodeId,
});

/** מסמך אמיתי: הטבלה היא בלוק אחד ואחריו פסקאות `TOC*`, ואז גוף המסמך. */
const documentWithToc = (): Block[] => [
  { ordinal: 0, nodeId: 'toc-1', nodeType: 'tableOfContents', styleId: 'TOC1' },
  { ordinal: 1, nodeId: 'row-1', nodeType: 'paragraph', styleId: 'TOC2' },
  { ordinal: 2, nodeId: 'row-2', nodeType: 'paragraph', styleId: 'TOC1' },
  { ordinal: 3, nodeId: 'body-1', nodeType: 'heading', styleId: 'Heading1' },
  { ordinal: 4, nodeId: 'body-2', nodeType: 'paragraph', styleId: 'Normal' },
];

describe('עדכון תוכן העניינים', () => {
  it('מעדכן כל טבלה במצב „הכול”, ומעביר את הכתובת שהרשימה החזירה', async () => {
    const engine = fakeEngine({ tocs: ['toc-1', 'toc-2'], total: 2 });

    expect(await updateTableOfContents(engine.host)).toEqual({ ok: true });

    expect(engine.inputs('toc.update')).toEqual([
      { target: tocAddress('toc-1'), mode: 'all' },
      { target: tocAddress('toc-2'), mode: 'all' },
    ]);
  });

  it('שואב עמודים עד `total` ולא עוצר בעמוד הראשון', async () => {
    const tocs = Array.from({ length: 320 }, (_, index) => `toc-${index}`);
    const engine = fakeEngine({ tocs, total: 320 });

    expect(await updateTableOfContents(engine.host)).toEqual({ ok: true });

    expect(engine.inputs('toc.update')).toHaveLength(320);
    expect(engine.inputs('toc.list')).toEqual([
      { limit: 200, offset: 0 },
      { limit: 200, offset: 200 },
    ]);
  });

  it('שתי טבלאות עם אותה כתובת — קריאה אחת, ודיווח שהעדכון לא הושלם', async () => {
    // נמדד בדפדפן: שתי טבלאות שנוצרו באותו מסמך חולקות `nodeId` (hash של
    // ה-`instruction`), וגם `blocks.list` מציג את שתיהן תחת אותו מזהה. לולאה
    // על ה-items בונה את הראשונה פעמיים ומשאירה את השנייה מיושנת — ומדווחת
    // „בוצע”. זה בדיוק הכפתור שמשקר.
    const engine = fakeEngine({ tocs: ['toc-1', 'toc-1'], total: 2 });

    expect(await updateTableOfContents(engine.host)).toEqual({
      ok: false,
      message:
        'עדכון תוכן העניינים לא הושלם: יש במסמך כמה טבלאות תוכן עניינים שאינן ניתנות להבחנה זו מזו, ולכן אחת מהן לא עודכנה',
      reason: 'ambiguous-toc',
    });
    expect(engine.inputs('toc.update')).toEqual([{ target: tocAddress('toc-1'), mode: 'all' }]);
  });

  it('שלוש טבלאות בשתי כתובות — מונה את מה שלא עודכן', async () => {
    const engine = fakeEngine({ tocs: ['toc-1', 'toc-1', 'toc-1', 'toc-2'], total: 4 });

    const outcome = await updateTableOfContents(engine.host);

    expect(outcome.ok === false && outcome.message).toContain('2 מהן לא עודכנו');
    expect(engine.inputs('toc.update')).toHaveLength(2);
  });

  it('מסמך בלי תוכן עניינים — כשל מוסבר, ולא שתיקה', async () => {
    // בשונה מ„עדכן שדות”, תוכן עניינים הוא עצם אחד גלוי: לחיצה על „עדכן
    // טבלה” במסמך שאין בו טבלה היא טעות שכדאי לומר עליה.
    const engine = fakeEngine();

    expect(await updateTableOfContents(engine.host)).toEqual({
      ok: false,
      message: 'עדכון תוכן העניינים נכשל: אין במסמך תוכן עניינים',
      reason: 'no-toc',
    });
    expect(engine.ops()).not.toContain('toc.update');
  });

  it('עוצר בכשל הראשון', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1', 'toc-2'],
      total: 2,
      failures: { 'toc.update': { code: 'PRECONDITION_FAILED' } },
    });

    const outcome = await updateTableOfContents(engine.host);

    expect(outcome.ok === false && outcome.reason).toBe('PRECONDITION_FAILED');
    expect(engine.inputs('toc.update')).toHaveLength(1);
  });

  it('`NO_OP` היא הצלחה — טבלה מעודכנת אינה טעונה עדכון', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'],
      total: 1,
      failures: { 'toc.update': { code: 'NO_OP' } },
    });

    expect(await updateTableOfContents(engine.host)).toEqual({ ok: true });
  });

  it('חריגה אינה יוצאת החוצה', async () => {
    const engine = fakeEngine({ tocs: ['toc-1'], total: 1, throws: ['toc.update'] });

    const outcome = await updateTableOfContents(engine.host);

    expect(outcome.ok === false && outcome.reason).toBe('threw');
  });

  it('גרסה בלי `toc.update` — הנוסח שהתכנית קובעת', async () => {
    expect(await updateTableOfContents(fakeEngine({ missing: ['toc.update'] }).host)).toEqual({
      ok: false,
      message: 'עדכון תוכן העניינים נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });

  it('בלי Document API כלל — כשל מוסבר ולא זריקה', async () => {
    expect(await updateTableOfContents(null)).toEqual({
      ok: false,
      message: 'עדכון תוכן העניינים נכשל: המסמך עדיין נטען',
      reason: 'document-api-unavailable',
    });
  });
});

describe('הסרת תוכן העניינים', () => {
  it('המנוע מסיר את כל השדה; האדפטר אינו מוחק טווח נוסף', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'], total: 1, blocks: documentWithToc(), removeTakesRows: 2,
    });
    expect(await removeTableOfContents(engine.host)).toEqual({ ok: true });
    expect(engine.inputs('toc.remove')).toEqual([{ target: tocAddress('toc-1') }]);
    expect(engine.remaining()).toEqual(['body-1', 'body-2']);
    expect(engine.ops()).not.toContain('blocks.deleteRange');
    expect(engine.ops()).not.toContain('blocks.list');
    expect(engine.ops()).not.toContain('create.paragraph');
  });

  it.each([undefined, 0, 1, 2, 4, 9])(
    'פסקת TOC1 של המשתמש שורדת טבלה מיושנת עם entryCount=%s', async (entryCount) => {
      const engine = fakeEngine({
        tocs: ['toc-1'], total: 1, entryCount, removeTakesRows: 1,
        blocks: [
          { ordinal: 0, nodeId: 'toc-1', nodeType: 'tableOfContents', styleId: 'TOC1' },
          { ordinal: 1, nodeId: 'row-1', nodeType: 'paragraph', styleId: 'TOC2' },
          { ordinal: 2, nodeId: 'mine', nodeType: 'paragraph', styleId: 'TOC1' },
          { ordinal: 3, nodeId: 'body-1', nodeType: 'paragraph', styleId: 'Normal' },
        ],
      });
      expect(await removeTableOfContents(engine.host)).toEqual({ ok: true });
      expect(engine.remaining()).toEqual(['mine', 'body-1']);
      expect(engine.ops()).not.toContain('blocks.deleteRange');
    },
  );

  it('פסקה בסגנון מותאם אישית אחרי השדה נשמרת גם כשאין מקורות חיים', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'], total: 1, entryCount: 0, removeTakesRows: 0,
      preserved: { customStyles: [{ styleName: 'MyToc1', level: 1 }] },
      blocks: [
        { ordinal: 0, nodeId: 'toc-1', nodeType: 'tableOfContents', styleId: 'MyToc1' },
        { ordinal: 1, nodeId: 'mine', nodeType: 'paragraph', styleId: 'MyToc1' },
      ],
    });
    expect(await removeTableOfContents(engine.host)).toEqual({ ok: true });
    expect(engine.remaining()).toEqual(['mine']);
  });

  it('תוכן עניינים בן 250 שורות מוסר בקריאה אחת למנוע', async () => {
    const rows = Array.from({ length: 249 }, (_, index) => ({
      ordinal: index + 1, nodeId: `row-${index}`, nodeType: 'paragraph', styleId: 'TOC1',
    }));
    const engine = fakeEngine({
      tocs: ['toc-1'], total: 1, entryCount: 250, removeTakesRows: rows.length,
      blocks: [
        { ordinal: 0, nodeId: 'toc-1', nodeType: 'tableOfContents', styleId: 'TOC1' },
        ...rows,
        { ordinal: 250, nodeId: 'body-1', nodeType: 'paragraph', styleId: 'Normal' },
      ],
    });
    expect(await removeTableOfContents(engine.host)).toEqual({ ok: true });
    expect(engine.inputs('toc.remove')).toHaveLength(1);
    expect(engine.remaining()).toEqual(['body-1']);
    expect(engine.ops()).not.toContain('blocks.deleteRange');
  });

  it('טבלה בסוף המסמך שהמנוע מסרב להסיר — פסקה נוספת וניסיון חוזר', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'], total: 1, blocks: documentWithToc().slice(0, 3),
      refuseTocAtEnd: true, removeTakesRows: 2,
    });
    expect(await removeTableOfContents(engine.host)).toEqual({ ok: true });
    expect(engine.inputs('create.paragraph')).toEqual([
      { at: { kind: 'after', target: { kind: 'block', nodeType: 'paragraph', nodeId: 'row-2' } } },
    ]);
    expect(engine.inputs('toc.remove')).toHaveLength(2);
    expect(engine.remaining()).toEqual(['trail']);
  });

  it('סירוב שאינו בגלל סוף המסמך מוחזר בלי הוספת פסקה', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'], total: 1, blocks: documentWithToc(),
      failures: { 'toc.remove': { code: 'CAPABILITY_UNAVAILABLE' } },
    });
    expect((await removeTableOfContents(engine.host)).ok).toBe(false);
    expect(engine.ops()).not.toContain('create.paragraph');
    expect(engine.inputs('toc.remove')).toHaveLength(1);
  });

  it('הסרה אינה דורשת blocks.list או blocks.deleteRange', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'], total: 1, blocks: documentWithToc(), removeTakesRows: 2,
      missing: ['blocks.list', 'blocks.deleteRange'],
    });
    expect(await removeTableOfContents(engine.host)).toEqual({ ok: true });
    expect(engine.remaining()).toEqual(['body-1', 'body-2']);
  });

  it('שתי טבלאות — מסרבת ואינה מנחשת', async () => {
    const engine = fakeEngine({ tocs: ['toc-1', 'toc-2'], total: 2, blocks: documentWithToc() });
    const outcome = await removeTableOfContents(engine.host);
    expect(outcome.ok === false && outcome.reason).toBe('ambiguous-toc');
    expect(engine.ops()).not.toContain('toc.remove');
  });

  it('מסמך בלי תוכן עניינים — כשל מוסבר', async () => {
    expect(await removeTableOfContents(fakeEngine().host)).toEqual({
      ok: false,
      message: 'הסרת תוכן העניינים נכשלה: אין במסמך תוכן עניינים',
      reason: 'no-toc',
    });
  });

  it('כשל של toc.remove משאיר את הבלוקים ומוחזר למשתמש', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'], total: 1, blocks: documentWithToc(),
      failures: { 'toc.remove': { code: 'TARGET_NOT_FOUND' } },
    });
    const outcome = await removeTableOfContents(engine.host);
    expect(outcome.ok === false && outcome.reason).toBe('TARGET_NOT_FOUND');
    expect(engine.remaining()).toEqual(documentWithToc().map((block) => block.nodeId));
    expect(engine.ops()).not.toContain('blocks.deleteRange');
  });

  it('חריגה אינה יוצאת החוצה', async () => {
    const engine = fakeEngine({ tocs: ['toc-1'], total: 1, throws: ['toc.remove'] });
    expect((await removeTableOfContents(engine.host)).ok).toBe(false);
  });
});

/* ------------------------------------------------------------------ */
/* מוביל הנקודות ומספר העמוד                                            */
/* ------------------------------------------------------------------ */

/**
 * הבאג שנסגר כאן: בשורת תוכן עניינים היה מספר העמוד צמוד לכותרת ובלי נקודות
 * מפרידות. נמדד שאין עצירת טאב בשום שכבה — לא ב-pPr, ולא בסגנונות
 * `TOC1`…`TOC9` שאינם מוגדרים ב-`styles.xml` של התבנית כלל — ולכן תו הטאב
 * שהמנוע כותב נופל על `w:defaultTabStop`. ההנמקה המלאה, כולל המדידות,
 * ב-engine/toc.ts.
 */
describe('מוביל הנקודות ומספר העמוד', () => {
  it('מיישרת כל שורה, כולל הבלוק הראשון של הטבלה, ולא את גוף המסמך', async () => {
    const engine = fakeEngine({ tocs: ['toc-1'], blocks: documentWithToc() });

    expect(await layOutTocRows(engine.host)).toEqual({ ok: true });

    // המיקום אחיד בכל רמה: `w:pos` נמדד משולי העמוד ב-Word, ולכן פיצוי לפי
    // הזחה היה יוצר מסמך שמספרי העמוד שלו אינם בעמודה אחת מחוץ ל-SuperDoc.
    expect(engine.inputs('setTabStop')).toEqual([
      {
        target: { kind: 'block', nodeType: 'paragraph', nodeId: 'toc-1' },
        position: A4_TAB_POSITION,
        alignment: 'right',
        leader: 'dot',
      },
      {
        target: { kind: 'block', nodeType: 'paragraph', nodeId: 'row-1' },
        position: A4_TAB_POSITION,
        alignment: 'right',
        leader: 'dot',
      },
      {
        target: { kind: 'block', nodeType: 'paragraph', nodeId: 'row-2' },
        position: A4_TAB_POSITION,
        alignment: 'right',
        leader: 'dot',
      },
    ]);
  });

  it('רמות עמוקות נשארות באותה עמודה', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'],
      blocks: [
        { ordinal: 0, nodeId: 'toc-1', nodeType: 'tableOfContents', styleId: 'TOC1' },
        { ordinal: 1, nodeId: 'row-1', nodeType: 'paragraph', styleId: 'TOC3' },
        { ordinal: 2, nodeId: 'row-2', nodeType: 'paragraph', styleId: 'TOC9' },
      ],
    });

    await layOutTocRows(engine.host);

    expect(
      engine.inputs('setTabStop').map((input) => (input as { position: number }).position),
    ).toEqual([A4_TAB_POSITION, A4_TAB_POSITION, A4_TAB_POSITION]);
  });

  it('עמוד צר — העצירה נשארת בערך חיובי', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'],
      // אזור טקסט 1440 twips, גדול מההסתייגות.
      page: { widthInches: 2, leftInches: 0.25, rightInches: 0.25 },
      blocks: [{ ordinal: 0, nodeId: 'row-1', nodeType: 'paragraph', styleId: 'TOC9' }],
    });

    await layOutTocRows(engine.host);

    const [input] = engine.inputs('setTabStop') as { position: number }[];
    expect(input.position).toBeGreaterThan(0);
  });

  it('המיקום נגזר ממידות העמוד שבמסמך, ולא מקבוע', async () => {
    // Letter (12240) עם שוליים צרים (720) — אזור טקסט 10800.
    const engine = fakeEngine({
      tocs: ['toc-1'],
      blocks: documentWithToc(),
      page: { widthInches: 12240 / 1440, leftInches: 0.5, rightInches: 0.5 },
    });

    expect(await layOutTocRows(engine.host)).toEqual({ ok: true });

    const positions = engine
      .inputs('setTabStop')
      .map((input) => (input as { position: number }).position);
    expect(positions).toEqual([
      10800 - TOC_PAGE_NUMBER_RESERVE_TWIPS,
      10800 - TOC_PAGE_NUMBER_RESERVE_TWIPS,
      10800 - TOC_PAGE_NUMBER_RESERVE_TWIPS,
    ]);
  });

  it('מנקה את העצירות של השורה לפני שהיא כותבת, ובסדר הזה', async () => {
    // `setTabStop` מוסיף ואינו מחליף, ואי אפשר לקרוא עצירות קיימות: בלי
    // הניקוי, מסמך שגודל העמוד שלו שונה היה נושא שתי עצירות — ותו הטאב היה
    // נדחף אל הישנה שבהן.
    const engine = fakeEngine({
      tocs: ['toc-1'],
      blocks: [{ ordinal: 0, nodeId: 'toc-1', nodeType: 'tableOfContents', styleId: 'TOC1' }],
    });

    await layOutTocRows(engine.host);

    const order = engine
      .ops()
      .filter((op) => op === 'clearAllTabStops' || op === 'setTabStop');
    expect(order).toEqual(['clearAllTabStops', 'setTabStop']);
    expect(engine.inputs('clearAllTabStops')).toEqual([
      { target: { kind: 'block', nodeType: 'paragraph', nodeId: 'toc-1' } },
    ]);
  });

  it('עדכון חוזר של אותן שורות אינו כותב להן עצירות מחדש', async () => {
    const engine = fakeEngine({ tocs: ['toc-1'], blocks: documentWithToc() });

    await layOutTocRows(engine.host);
    expect(engine.inputs('setTabStop')).toHaveLength(3);
    expect(engine.inputs('clearAllTabStops')).toHaveLength(3);

    await layOutTocRows(engine.host);
    expect(engine.inputs('setTabStop')).toHaveLength(3);
    expect(engine.inputs('clearAllTabStops')).toHaveLength(3);
  });

  it('שורה שסווגה `listItem` נשלחת בסוג שלה, ולא כפסקה', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'],
      blocks: [{ ordinal: 0, nodeId: 'row-1', nodeType: 'listItem', styleId: 'TOC1' }],
    });

    await layOutTocRows(engine.host);

    expect(engine.inputs('setTabStop')).toEqual([
      {
        target: { kind: 'block', nodeType: 'listItem', nodeId: 'row-1' },
        position: A4_TAB_POSITION,
        alignment: 'right',
        leader: 'dot',
      },
    ]);
  });

  it('סגנון שורה של מסמך Word (מתג t) מיושר גם הוא', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'],
      preserved: { customStyles: [{ styleName: 'MyToc1', level: 1 }] },
      blocks: [
        { ordinal: 0, nodeId: 'toc-1', nodeType: 'tableOfContents', styleId: 'MyToc1' },
        { ordinal: 1, nodeId: 'row-1', nodeType: 'paragraph', styleId: 'MyToc1' },
        { ordinal: 2, nodeId: 'body-1', nodeType: 'paragraph', styleId: 'Normal' },
      ],
    });

    await layOutTocRows(engine.host);

    expect(
      engine
        .inputs('setTabStop')
        .map((input) => (input as { target: { nodeId: string } }).target.nodeId),
    ).toEqual(['toc-1', 'row-1']);
  });

  it('מסמך בלי שורות תוכן עניינים — הצלחה שקטה ובלי מוטציה', async () => {
    const engine = fakeEngine({
      tocs: [],
      blocks: [{ ordinal: 0, nodeId: 'body-1', nodeType: 'paragraph', styleId: 'Normal' }],
    });

    expect(await layOutTocRows(engine.host)).toEqual({ ok: true });
    expect(engine.ops()).not.toContain('setTabStop');
  });

  it('מידות עמוד שאינן זמינות — כשל מנומק ולא ניחוש של רוחב', async () => {
    const engine = fakeEngine({ tocs: ['toc-1'], blocks: documentWithToc(), page: null });

    expect(await layOutTocRows(engine.host)).toEqual({
      ok: false,
      message: 'עיצוב שורות תוכן העניינים נכשל: מידות העמוד אינן זמינות',
      reason: 'no-page-metrics',
    });
    expect(engine.ops()).not.toContain('setTabStop');
  });

  it('אזור טקסט צר מהמקום שנשמר למספר — כשל, ולא `w:pos` פסול', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'],
      blocks: documentWithToc(),
      // 720 twips רוחב, פחות 144 מכל צד — אזור טקסט 432, צר מהמקום השמור.
      page: { widthInches: 0.5, leftInches: 0.1, rightInches: 0.1 },
    });

    expect(await layOutTocRows(engine.host)).toEqual({
      ok: false,
      message: 'עיצוב שורות תוכן העניינים נכשל: אזור הטקסט בעמוד צר מדי',
      reason: 'page-too-narrow',
    });
  });

  it('גרסת מנוע בלי `setTabStop` — „אינו זמין בגרסה זו”', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'],
      blocks: documentWithToc(),
      missing: ['setTabStop'],
    });

    expect(await layOutTocRows(engine.host)).toEqual({
      ok: false,
      message: 'עיצוב שורות תוכן העניינים נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });

  it('קבלה שנכשלה נעצרת ומדווחת, ואינה ממשיכה לשורה הבאה', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'],
      blocks: documentWithToc(),
      failures: { setTabStop: { code: 'INVALID_TARGET', message: 'no' } },
    });

    const outcome = await layOutTocRows(engine.host);
    expect(outcome.ok === false && outcome.reason).toBe('INVALID_TARGET');
    expect(engine.inputs('setTabStop')).toHaveLength(1);
  });

  it('חריגה מהמנוע מוחזרת כ-`CommandOutcome` ואינה מתפשטת', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'],
      blocks: documentWithToc(),
      throws: ['setTabStop'],
    });

    const outcome = await layOutTocRows(engine.host);
    expect(outcome.ok === false && outcome.reason).toBe('threw');
  });

  it('`NO_OP` — עצירה שכבר קיימת — הוא הצלחה', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'],
      blocks: documentWithToc(),
      failures: { setTabStop: { code: 'NO_OP' }, clearAllTabStops: { code: 'NO_OP' } },
    });

    expect(await layOutTocRows(engine.host)).toEqual({ ok: true });
    expect(engine.inputs('setTabStop')).toHaveLength(3);
  });

  it('שורות שמעבר לעמוד הראשון של `blocks.list` מיושרות גם הן', async () => {
    // PAGE_SIZE הוא 200, ותוכן עניינים של ספר ארוך ממנו. קריאה יחידה הייתה
    // משאירה את השורות שאחריה בלי מוביל.
    const blocks = Array.from({ length: 250 }, (_, index) => ({
      ordinal: index,
      nodeId: `row-${index}`,
      nodeType: index === 0 ? 'tableOfContents' : 'paragraph',
      styleId: 'TOC1',
    }));
    const engine = fakeEngine({ tocs: ['toc-1'], blocks });

    expect(await layOutTocRows(engine.host)).toEqual({ ok: true });
    expect(engine.inputs('setTabStop')).toHaveLength(250);
  });
});

describe('התאמה אישית', () => {
  it('שולחת את טווח הרמות ואת דגל הקישורים, ותו לא', async () => {
    // ההגבלה לשני שדות אינה קיצור דרך: `tabLeader`, `rightAlignPageNumbers`
    // ו-`includePageNumbers` נמדדו כנבלעים בשקט. ראו engine/toc.ts.
    const engine = fakeEngine({ tocs: ['toc-1'], total: 1 });

    expect(
      await configureTableOfContents(engine.host, { levels: { from: 1, to: 3 }, hyperlinks: true }),
    ).toEqual({ ok: true });

    expect(engine.inputs('toc.configure')).toEqual([
      {
        target: tocAddress('toc-1'),
        patch: { outlineLevels: { from: 1, to: 3 }, hyperlinks: true },
      },
    ]);
  });

  it('טווח הפוך נדחה לפני שנוגעים במנוע', async () => {
    // המנוע מקבל `{from:9,to:1}` עם `success: true` וכותב `TOC \\o "9-1"` —
    // והטבלה נעשית ריקה. נמדד.
    const engine = fakeEngine({ tocs: ['toc-1'], total: 1 });

    const outcome = await configureTableOfContents(engine.host, {
      levels: { from: 9, to: 1 },
      hyperlinks: false,
    });

    expect(outcome.ok === false && outcome.reason).toBe('invalid-levels');
    expect(engine.ops()).not.toContain('toc.configure');
  });

  it('רמה שאינה בטווח 1–9 נדחית גם היא', async () => {
    const engine = fakeEngine({ tocs: ['toc-1'], total: 1 });

    for (const levels of [
      { from: 0, to: 3 },
      { from: 1, to: 10 },
      { from: 1.5, to: 3 },
    ]) {
      const outcome = await configureTableOfContents(engine.host, { levels, hyperlinks: false });
      expect(outcome.ok, JSON.stringify(levels)).toBe(false);
    }
    expect(engine.ops()).not.toContain('toc.configure');
  });

  it('שתי טבלאות — מסרבת, כמו בהסרה ומאותו טעם', async () => {
    const engine = fakeEngine({ tocs: ['toc-1', 'toc-2'], total: 2 });

    const outcome = await configureTableOfContents(engine.host, {
      levels: { from: 1, to: 3 },
      hyperlinks: false,
    });

    expect(outcome.ok === false && outcome.reason).toBe('ambiguous-toc');
  });

  it('גרסה בלי `toc.configure` — הנוסח שהתכנית קובעת', async () => {
    const engine = fakeEngine({ tocs: ['toc-1'], total: 1, missing: ['toc.configure'] });

    expect(
      await configureTableOfContents(engine.host, { levels: { from: 1, to: 3 }, hyperlinks: true }),
    ).toEqual({
      ok: false,
      message: 'שינוי הגדרות תוכן העניינים נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });
});

describe('סימון ערך', () => {
  it('מכניס שדה `TC` בסוף הפסקה שהסמן בה', async () => {
    const engine = fakeEngine({ blockId: 'block-7' });

    expect(await markTocEntry(engine.host, 'הלכות שבת', 2)).toEqual({ ok: true });

    expect(engine.inputs('toc.markEntry')).toEqual([
      {
        target: {
          kind: 'inline-insert',
          anchor: { nodeType: 'paragraph', nodeId: 'block-7' },
          position: 'end',
        },
        text: 'הלכות שבת',
        level: 2,
      },
    ]);
  });

  it('גוזם רווחים, ודוחה טקסט ריק לפני שנוגע במנוע', async () => {
    // המנוע **זורק** על מחרוזת ריקה; הבדיקה כאן היא מה שהופך את זה להודעה.
    const engine = fakeEngine({ blockId: 'block-7' });

    expect(await markTocEntry(engine.host, '  הקדמה  ', 1)).toEqual({ ok: true });
    expect((engine.inputs('toc.markEntry')[0] as { text: string }).text).toBe('הקדמה');

    const outcome = await markTocEntry(engine.host, '   ', 1);
    expect(outcome.ok === false && outcome.reason).toBe('invalid-text');
    expect(engine.inputs('toc.markEntry')).toHaveLength(1);
  });

  it('רמה מחוץ ל-1–9 נדחית — המנוע מקבל אותה וכותב מתג פסול ל-Word', async () => {
    const engine = fakeEngine({ blockId: 'block-7' });

    for (const level of [0, 10, 12, 1.5]) {
      const outcome = await markTocEntry(engine.host, 'ערך', level);
      expect(outcome.ok === false && outcome.reason, String(level)).toBe('invalid-level');
    }
    expect(engine.ops()).not.toContain('toc.markEntry');
  });

  it('בלי סמן במסמך — מסביר מה לעשות', async () => {
    const engine = fakeEngine({ blockId: null });

    const outcome = await markTocEntry(engine.host, 'ערך', 1);

    expect(outcome.ok === false && outcome.reason).toBe('no-selection');
    expect(outcome.ok === false && outcome.message).toContain('הפסקה שאליה יסומן הערך');
  });

  it('קבלה שנכשלה מתורגמת לעברית', async () => {
    const engine = fakeEngine({
      blockId: 'block-7',
      failures: { 'toc.markEntry': { code: 'TARGET_NOT_FOUND' } },
    });

    const outcome = await markTocEntry(engine.host, 'ערך', 1);

    expect(outcome.ok === false && outcome.reason).toBe('TARGET_NOT_FOUND');
    expect(outcome.ok === false && outcome.message).toContain('סימון הערך נכשל');
  });

  it('חריגה אינה יוצאת החוצה', async () => {
    const engine = fakeEngine({ blockId: 'block-7', throws: ['toc.markEntry'] });

    expect((await markTocEntry(engine.host, 'ערך', 1)).ok).toBe(false);
  });

  it('גרסה בלי `toc.markEntry` — הנוסח שהתכנית קובעת', async () => {
    expect(await markTocEntry(fakeEngine({ missing: ['toc.markEntry'] }).host, 'ערך', 1)).toEqual({
      ok: false,
      message: 'סימון הערך נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });
});

describe('ביטול סימון ערך', () => {
  it('שולח את הכתובת שהרשימה החזירה', async () => {
    const engine = fakeEngine({ entries: [{ nodeId: 'entry-1' }] });

    expect(await unmarkTocEntry(engine.host, 'entry-1')).toEqual({ ok: true });
    expect(engine.inputs('toc.unmarkEntry')).toEqual([
      { target: { kind: 'inline', nodeType: 'tableOfContentsEntry', nodeId: 'entry-1' } },
    ]);
  });

  it('ערך שאינו קיים — קבלה שנכשלה, לא זריקה', async () => {
    const engine = fakeEngine({ failures: { 'toc.unmarkEntry': { code: 'TARGET_NOT_FOUND' } } });

    const outcome = await unmarkTocEntry(engine.host, 'entry-9');

    expect(outcome.ok === false && outcome.reason).toBe('TARGET_NOT_FOUND');
    expect(outcome.ok === false && outcome.message).toContain('ביטול סימון הערך נכשל');
  });

  it('בלי Document API כלל — כשל מוסבר', async () => {
    expect(await unmarkTocEntry(null, 'entry-1')).toEqual({
      ok: false,
      message: 'ביטול סימון הערך נכשל: המסמך עדיין נטען',
      reason: 'document-api-unavailable',
    });
  });
});

describe('מצב תוכן העניינים', () => {
  it('סופר טבלאות, קורא הגדרות מהראשונה, ומונה את הערכים המסומנים', async () => {
    const engine = fakeEngine({
      tocs: ['toc-1'],
      total: 1,
      sourceConfig: { outlineLevels: { from: 1, to: 4 } },
      displayConfig: { hyperlinks: true },
      entries: [{ nodeId: 'entry-1', text: 'הקדמה', level: 2 }],
      entriesTotal: 1,
    });

    expect(await readTocState(engine.host)).toEqual({
      count: 1,
      levels: { from: 1, to: 4 },
      hyperlinks: true,
      entries: [{ nodeId: 'entry-1', text: 'הקדמה', level: 2 }],
    });
  });

  it('שואב עמודים עד `total` — רשימה חלקית הייתה מסתירה ערכים', async () => {
    const entries = Array.from({ length: 250 }, (_, index) => ({ nodeId: `entry-${index}` }));
    const engine = fakeEngine({ entries, entriesTotal: 250 });

    const state = await readTocState(engine.host);

    expect(state.entries).toHaveLength(250);
    expect(engine.inputs('toc.listEntries')).toEqual([
      { limit: 200, offset: 0 },
      { limit: 200, offset: 200 },
    ]);
  });

  it('טבלה בלי `\\o` מדווחת `levels: null` ולא ממציאה טווח', async () => {
    // זה המצב הרגיל: `create.tableOfContents` כותב `TOC \\h` בלי `\\o` בכלל.
    const engine = fakeEngine({ tocs: ['toc-1'], total: 1 });

    expect((await readTocState(engine.host)).levels).toBeNull();
  });

  it('ערך בלי כתובת מדולג ואינו מגיע לדיאלוג כמזהה ריק', async () => {
    const engine = fakeEngine({ entries: [{ nodeId: '' }, { nodeId: 'entry-2' }], entriesTotal: 2 });

    expect((await readTocState(engine.host)).entries).toEqual([
      { nodeId: 'entry-2', text: '', level: 1 },
    ]);
  });

  it('חריגה או פעולה חסרה — „אין תוכן עניינים”, ולא זריקה', async () => {
    expect(await readTocState(fakeEngine({ throws: ['toc.list'] }).host)).toEqual(emptyTocState());
    expect(await readTocState(fakeEngine({ missing: ['toc.list'] }).host)).toEqual(emptyTocState());
    expect(await readTocState(null)).toEqual(emptyTocState());
  });
});

describe('הוולידציה שהדיאלוגים והמודול חולקים', () => {
  it('`normalizeTocEntryText` גוזם, ודוחה ריק', () => {
    expect(normalizeTocEntryText('  ערך  ')).toBe('ערך');
    expect(normalizeTocEntryText('   ')).toBeNull();
    expect(normalizeTocEntryText('')).toBeNull();
  });

  it('`isValidTocLevel` — שלם בין 1 ל-9 בלבד', () => {
    expect([1, 5, 9].every(isValidTocLevel)).toBe(true);
    expect([0, 10, -1, 1.5, Number.NaN].some(isValidTocLevel)).toBe(false);
  });

  it('`normalizeTocLevels` מחזיר את הטווח, או `null`', () => {
    expect(normalizeTocLevels(1, 3)).toEqual({ from: 1, to: 3 });
    expect(normalizeTocLevels(3, 3)).toEqual({ from: 3, to: 3 });
    expect(normalizeTocLevels(3, 1)).toBeNull();
    expect(normalizeTocLevels(0, 3)).toBeNull();
  });
});
