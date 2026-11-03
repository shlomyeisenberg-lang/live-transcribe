/**
 * `createSearchAdapter` הוחלף כליל: הוא כבר אינו קורא ל-`ui.search` (ראו
 * הראש של `engine/search.ts` להסבר המלא — `ui.search` נמדד שאינו מכסה מסמך
 * רב-פסקאות, `projection-incomplete`). המימוש החדש עובר דרך `doc.blocks
 * .list`/`doc.replace`/`doc.mutations.apply` בלבד, ולכן הכפיל כאן ("strict
 * double") אוכף בדיוק את אלה — לא `ui.search` ולא שום מתודה אחרת של
 * Document API — באותה רוח שהכפיל הישן אכף את `SearchHandle`.
 *
 * הכפיל מדמה מסמך אמיתי (מערך בלוקים עם טקסט קנוני, ו-`replace`/
 * `mutations.apply` שמבצעים את הגזירה בפועל), ולא רק "מחזיר true ומקליט
 * ארגומנטים" — כי הבדיקה המרכזית כאן היא בדיוק התרחיש שהיה שבור: מסמך עם
 * מופע אחד בכל אחת משמונה פסקאות, ו"החלף הכל" שמחליף את כולן.
 *
 * `mutations.apply` קיים כאן בגלל רגרסיה שנייה, שנמדדה **אחרי** שהכיסוי
 * תוקן: הקוד הישן קרא ל-`handle.replaceAll()` — batch יחיד במנוע, עטוף
 * ב-`undo`-step אחד. המעבר ל-`doc.replace` הנקודתי (מופע-מופע) פתר את
 * הכיסוי אבל פתח N שלבי ביטול נפרדים ל-N מופעים. הפתרון — ראו התיעוד המלא
 * ב-`engine/search.ts` (`applyAtomicRewriteChunks`) — הוא `doc.mutations
 * .apply({atomic:true, steps:[...text.rewrite]})`, שנותן `undo`-step יחיד
 * לכל הקבוצה. הבדיקות כאן מוודאות ששני הנתיבים תקינים: הנתיב האטומי כשהוא
 * זמין ומדווח נכון, והנפילה לנתיב הנקודתי כשהוא לא (או מדווח ספירה שגויה)
 * — בלי לאבד ולו מופע אחד באף מסלול.
 */
import { describe, it, expect, vi, afterEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  createSearchAdapter,
  searchCounterText,
  idleSearchState,
  replaceControlsVisible,
  NO_MATCHES_TEXT,
  NO_QUERY_TEXT,
  SEARCH_DEBOUNCE_MS,
  type SearchState,
  type SearchOutcome,
} from '../../src/engine/search';
import { REASON_TEXT } from '../../src/engine/command-adapter';

/** זורק על גישה לכל שם שאינו ברשימה — אותה טכניקה בדיוק כמו הכפיל הישן. */
function guard<T extends object>(label: string, obj: T): T {
  return new Proxy(obj, {
    get(target, prop, receiver) {
      if (typeof prop === 'string' && !(prop in target)) {
        throw new Error(`${label}.${prop} אינו קיים בחוזה`);
      }
      return Reflect.get(target, prop, receiver);
    },
  });
}

interface FakeBlock {
  nodeId: string;
  text: string;
  nodeType?: string;
}

interface Call {
  name: string;
  args: unknown[];
}

type ReplaceReceipt = { success: true } | { success: false; failure?: { code?: string; message?: string } };

interface FakeHostOptions {
  /** מותאם אישית: מה `replace` מחזיר על מופע נתון, *לפני* שהגזירה בפועל מתבצעת. */
  onReplace?: (input: { blockId: string; start: number; end: number; text: string }) => ReplaceReceipt | undefined;
  /** `replace` מחזיר Promise שנפתר רק כשקוראים ל-`settle()` — לבדיקת מרוצים. */
  deferReplace?: boolean;
  listThrows?: boolean;
  noBlocksApi?: boolean;
  noReplaceApi?: boolean;
  /** בלי `doc.mutations` בכלל — "החלף הכל" נופל ישירות לנתיב הנקודתי. */
  noMutationsApi?: boolean;
  mutationsApplyThrows?: boolean;
  mutationsApplyFailure?: { code?: string; message?: string };
  /**
   * בלוקים שעבורם `mutations.apply` "משקר" — מדווח `success` עם ספירה
   * שגויה, ובלי לגעת בבלוק בפועל. מדמה מנוע שמחזיר תוצאה חלקית לא-אמינה,
   * ובודק שהאדפטר לא בוטח בה ונופל לנתיב הנקודתי בשביל הבלוק הזה בלבד.
   */
  mutationsMismatchBlockIds?: ReadonlySet<string>;
}

function createFakeHost(initialBlocks: readonly FakeBlock[], options: FakeHostOptions = {}) {
  const blocks: FakeBlock[] = initialBlocks.map((b) => ({ ...b }));
  const calls: Call[] = [];
  const selectionCalls: unknown[] = [];
  const scrollCalls: unknown[] = [];
  let pendingSettle: ((value: ReplaceReceipt) => void) | undefined;

  function record(name: string, args: unknown[]): void {
    calls.push({ name, args });
  }

  function listFn(input?: { offset?: number; limit?: number }) {
    record('blocks.list', [input]);
    if (options.listThrows) throw new Error('קריאת הבלוקים נכשלה');
    const offset = input?.offset ?? 0;
    const limit = input?.limit ?? (blocks.length || 1);
    const page = blocks.slice(offset, offset + limit);
    return {
      total: blocks.length,
      blocks: page.map((b) => ({ nodeId: b.nodeId, text: b.text, nodeType: b.nodeType ?? 'paragraph' })),
    };
  }

  function replaceFn(input: {
    target: { start: { blockId: string; offset: number }; end: { blockId: string; offset: number } };
    text: string;
  }) {
    record('replace', [input]);
    const { start, end } = input.target;

    function apply(): ReplaceReceipt {
      const block = blocks.find((b) => b.nodeId === start.blockId);
      if (!block) return { success: false, failure: { code: 'TARGET_NOT_FOUND', message: 'לא נמצא' } };
      const custom = options.onReplace?.({
        blockId: start.blockId,
        start: start.offset,
        end: end.offset,
        text: input.text,
      });
      if (custom && !custom.success) return custom;
      block.text = block.text.slice(0, start.offset) + input.text + block.text.slice(end.offset);
      return { success: true };
    }

    if (options.deferReplace) {
      return new Promise<ReplaceReceipt>((resolve) => {
        pendingSettle = resolve;
      }).then((forced) => forced ?? apply());
    }
    return apply();
  }

  /**
   * מבצעת בפועל כל צעד `text.rewrite` (`where.by === 'select'`) על הבלוק
   * הנקוב ב-`within.nodeId`: מחליפה כל מופע לא-חופף של `pattern` (כמו
   * `findAllMatches` — כדי שהבדיקות יוכלו לצפות לאותה ספירה), וסופרת
   * אותם ב-`matchCount`. בלוקים ב-`mutationsMismatchBlockIds` "משקרים":
   * מדווחים ספירה שגויה ואינם נוגעים בבלוק בפועל — מדמים מנוע שהתוצאה
   * החלקית שלו אינה אמינה.
   */
  function mutationsApplyFn(input: {
    atomic: true;
    changeMode: string;
    steps: readonly {
      id: string;
      op: 'text.rewrite';
      where: {
        by: 'select';
        select: { type: 'text'; pattern: string; mode: 'contains'; caseSensitive: boolean };
        within: { kind: 'block'; nodeType: string; nodeId: string };
        require: 'all';
      };
      args: { replacement: { text: string } };
    }[];
  }) {
    record('mutations.apply', [input]);
    if (options.mutationsApplyThrows) throw new Error('mutations.apply נכשל');
    if (options.mutationsApplyFailure) return { success: false, failure: options.mutationsApplyFailure };

    const steps = input.steps.map((step) => {
      const blockId = step.where.within.nodeId;
      const block = blocks.find((b) => b.nodeId === blockId);
      if (!block) return { matchCount: 0 };

      if (options.mutationsMismatchBlockIds?.has(blockId)) {
        return { matchCount: -1 }; // ספירה שגויה בכוונה; הבלוק עצמו נשאר כפי שהיה.
      }

      const { pattern, caseSensitive } = step.where.select;
      const needle = caseSensitive ? pattern : pattern.toLowerCase();
      const haystack = caseSensitive ? block.text : block.text.toLowerCase();
      let count = 0;
      let rebuilt = '';
      let cursor = 0;
      for (;;) {
        const at = haystack.indexOf(needle, cursor);
        if (at < 0) break;
        rebuilt += block.text.slice(cursor, at) + step.args.replacement.text;
        cursor = at + needle.length;
        count += 1;
      }
      rebuilt += block.text.slice(cursor);
      block.text = rebuilt;
      return { matchCount: count };
    });

    return { success: true, steps };
  }

  // המפתחות נוכחים תמיד, גם כשהיכולת "כבויה" (`undefined` ולא השמטה): זה
  // ההבדל בין "היכולת הזאת פשוט אינה קיימת במסמך הזה" (מה ש-`noBlocksApi`/
  // `noReplaceApi`/`noMutationsApi` מדמים — בדיוק מה ש-`doc?.blocks?.list`
  // הייצור בודק) לבין "קריאה לשם שגוי שאינו בחוזה בכלל" (מה שה-guard על
  // doc/blocks/ui נועד לתפוס). guard שהיה זורק גם על הראשון היה הופך כל
  // בדיקת-יכולת לתקועה.
  const docShape: Record<string, unknown> = {
    blocks: options.noBlocksApi ? undefined : guard('doc.blocks', { list: listFn }),
    replace: options.noReplaceApi ? undefined : replaceFn,
    mutations: options.noMutationsApi
      ? undefined
      : guard('doc.mutations', { apply: mutationsApplyFn }),
  };

  const doc = guard('doc', docShape);
  const ui = guard('ui', {
    selection: guard('ui.selection', {
      apply: (target: unknown) => {
        selectionCalls.push(target);
        return { ok: true };
      },
    }),
    viewport: guard('ui.viewport', {
      scrollIntoView: (input: unknown) => {
        scrollCalls.push(input);
        return { success: true };
      },
    }),
  });

  return {
    host: { activeEditor: { doc }, ui },
    calls,
    selectionCalls,
    scrollCalls,
    names: () => calls.map((c) => c.name),
    currentTexts: () => Object.fromEntries(blocks.map((b) => [b.nodeId, b.text])),
    /** משחררת קריאת `replace` שממתינה (`deferReplace:true`). `undefined` = ההחלפה האמיתית. */
    settlePendingReplace: (forced?: ReplaceReceipt) => {
      pendingSettle?.(forced as ReplaceReceipt);
      pendingSettle = undefined;
    },
    /**
     * ממתינה עד ש-`replace` בפועל נקרא על ה-doc המדומה. נחוצה כי המימוש
     * החדש קורא ל-`refresh()` (א-סינכרוני — `blocks.list`) *לפני* שהוא מגיע
     * ל-`doc.replace` בכלל, ולכן `pendingSettle` אינו מוקצה מיד כשקוראים
     * ל-`adapter.replace()` — בניגוד למימוש הישן, שקרא ל-`handle.replace`
     * באופן סינכרוני. לולאת מיקרו-טסקים ולא ספירת `await`ים: מספר הטיקים
     * המדויק הוא פרט מימוש של V8/Node ולא חוזה שכדאי להישען עליו.
     */
    waitForPendingReplace: async () => {
      for (let i = 0; i < 50 && pendingSettle === undefined; i++) {
        await Promise.resolve();
      }
    },
  };
}

/** שמונה פסקאות, מופע אחד ("zzq") בכל אחת — בדיוק התרחיש שהיה שבור ב-ui.search. */
function eightParagraphBlocks(): FakeBlock[] {
  return Array.from({ length: 8 }, (_, i) => ({ nodeId: `p${i}`, text: `zzq paragraph ${i}` }));
}

afterEach(() => {
  vi.useRealTimers();
});

describe('הכפיל הקפדני של Document API', () => {
  it('זורק על גישה לשם שאינו בחוזה (blocks/replace/mutations בלבד)', () => {
    const { host } = createFakeHost([]);
    const doc = host.activeEditor.doc as unknown as Record<string, unknown>;
    expect(() => doc.find).toThrow(/find/);
    expect(() => doc.getText).toThrow();
    const blocksApi = doc.blocks as unknown as Record<string, unknown>;
    expect(() => blocksApi.search).toThrow();
    const mutationsApi = doc.mutations as unknown as Record<string, unknown>;
    expect(() => mutationsApi.preview).toThrow();
  });

  it('ui — רק selection.apply ו-viewport.scrollIntoView נגישים', () => {
    const { host } = createFakeHost([]);
    const ui = host.ui as unknown as Record<string, unknown>;
    expect(() => ui.commands).toThrow();
    const selection = ui.selection as unknown as Record<string, unknown>;
    expect(() => selection.restore).toThrow();
  });
});

describe('createSearchAdapter — כיסוי מלא של מסמך רב-פסקאות (התרחיש שהיה שבור)', () => {
  it('מוצא את כל שמונת המופעים, לא רק ארבעה', async () => {
    const { host } = createFakeHost(eightParagraphBlocks());
    const adapter = createSearchAdapter(host as never);

    const outcome = await adapter.find('zzq', 'next');

    expect(outcome.ok).toBe(true);
    expect(outcome.ok && outcome.snapshot.total).toBe(8);
    expect(outcome.ok && outcome.snapshot.activeIndex).toBe(0);
  });

  it('"החלף הכל" מחליף את כל שמונת המופעים, לא ארבעה', async () => {
    const { host, currentTexts } = createFakeHost(eightParagraphBlocks());
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const outcome = await adapter.replaceAll('YYY');

    expect(outcome.ok).toBe(true);
    expect(outcome.ok && outcome.snapshot.total).toBe(0);
    const texts = currentTexts();
    for (let i = 0; i < 8; i++) {
      expect(texts[`p${i}`]).toBe(`YYY paragraph ${i}`);
    }
  });

  /**
   * הרגרסייה השנייה: הקוד שקדם ל-`doc.mutations.apply` קרא ל-`doc.replace`
   * פעם לכל מופע — 8 קריאות נפרדות, שכל אחת שלב-ביטול משלה. `mutations
   * .apply` היחיד כאן הוא בדיוק מה שנותן `undo`-step יחיד (ראו QA: Ctrl+Z
   * בודד מבטל את כל השמונה בדפדפן אמיתי, `scripts/qa/replace-all-multiparagraph-qa.mjs`).
   */
  it('"החלף הכל" על שמונה פסקאות משתמש בקריאת mutations.apply אחת — לא בשמונה קריאות replace', async () => {
    const { host, calls } = createFakeHost(eightParagraphBlocks());
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const outcome = await adapter.replaceAll('YYY');

    expect(outcome.ok).toBe(true);
    expect(calls.filter((c) => c.name === 'mutations.apply')).toHaveLength(1);
    expect(calls.filter((c) => c.name === 'replace')).toHaveLength(0);

    const applyCall = calls.find((c) => c.name === 'mutations.apply');
    const input = applyCall?.args[0] as { atomic: boolean; steps: unknown[] };
    expect(input.atomic).toBe(true);
    expect(input.steps).toHaveLength(8);
  });
});

describe('createSearchAdapter — חיפוש', () => {
  it('שאילתה חדשה מסמנת את המופע הראשון, ולא מקדמת הלאה', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'a b a b a' }]);
    const adapter = createSearchAdapter(host as never);

    const outcome = await adapter.find('a', 'next');

    expect(outcome.ok && outcome.snapshot.activeIndex).toBe(0);
    expect(outcome.ok && outcome.snapshot.total).toBe(3);
  });

  it('"מצא הבא" על אותה שאילתה מתקדם, ו"מצא קודם" נסוג', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'a b a b a' }]);
    const adapter = createSearchAdapter(host as never);

    await adapter.find('a', 'next');
    const second = await adapter.find('a', 'next');
    expect(second.ok && second.snapshot.activeIndex).toBe(1);

    const back = await adapter.find('a', 'prev');
    expect(back.ok && back.snapshot.activeIndex).toBe(0);
  });

  it('"מצא הבא" עוטף מהאחרון לראשון', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'a a' }]);
    const adapter = createSearchAdapter(host as never);

    await adapter.find('a', 'next');
    const second = await adapter.find('a', 'next');
    expect(second.ok && second.snapshot.activeIndex).toBe(1);
    const wrapped = await adapter.find('a', 'next');
    expect(wrapped.ok && wrapped.snapshot.activeIndex).toBe(0);
  });

  it('לא רגיש לרישיות כברירת מחדל — תואם את ברירת המחדל שנמדדה במנוע', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'Hello hello HELLO' }]);
    const adapter = createSearchAdapter(host as never);
    const outcome = await adapter.find('hello', 'next');
    expect(outcome.ok && outcome.snapshot.total).toBe(3);
  });

  it('שאילתה בלי התאמות אינה כשל — "אין תוצאות"', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'שלום עולם' }]);
    const adapter = createSearchAdapter(host as never);

    const outcome = await adapter.find('אשכנז', 'next');

    expect(outcome.ok).toBe(true);
    expect(outcome.ok && searchCounterText(outcome.snapshot)).toBe('אין תוצאות');
  });

  it('שאילתה ריקה מנקה במקום לחפש — ואינה קוראת ל-blocks.list', async () => {
    const { host, calls } = createFakeHost([{ nodeId: 'p0', text: 'טקסט' }]);
    const adapter = createSearchAdapter(host as never);

    await adapter.find('טקסט', 'next');
    calls.length = 0;
    const outcome = await adapter.find('', 'next');

    expect(outcome.ok).toBe(true);
    expect(outcome.ok && outcome.snapshot.total).toBe(0);
    expect(calls).toEqual([]);
  });

  it('הדגשת המופע הפעיל: selection.apply ו-viewport.scrollIntoView נקראים על היעד הנכון', async () => {
    const { host, selectionCalls, scrollCalls } = createFakeHost([{ nodeId: 'p0', text: 'xx zzq xx' }]);
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');

    expect(selectionCalls).toEqual([
      { kind: 'selection', start: { kind: 'text', blockId: 'p0', offset: 3 }, end: { kind: 'text', blockId: 'p0', offset: 6 } },
    ]);
    expect(scrollCalls).toEqual([
      { target: { kind: 'text', blockId: 'p0', range: { start: 3, end: 6 } }, block: 'center', behavior: 'smooth' },
    ]);
  });
});

describe('createSearchAdapter — חיפוש שאינו זמין', () => {
  it('בלי doc.blocks בכלל — available:false, ולא נקרא replace', async () => {
    const { host, calls } = createFakeHost([{ nodeId: 'p0', text: 'טקסט' }], { noBlocksApi: true });
    const adapter = createSearchAdapter(host as never);

    const outcome = await adapter.find('טקסט', 'next');

    expect(outcome).toEqual({
      ok: false,
      message: REASON_TEXT['search-unavailable'],
      reason: 'search-unavailable',
    });
    expect(calls).toEqual([]);
  });

  it('open() נכשל סגור כש-available:false', () => {
    const { host } = createFakeHost([], { noBlocksApi: true });
    const adapter = createSearchAdapter(host as never);

    expect(adapter.open()).toEqual({
      ok: false,
      message: REASON_TEXT['search-unavailable'],
      reason: 'search-unavailable',
    });
  });

  it('blocks.list שזורק אינו מדווח כיסוי חלקי בשקט — כשל גלוי', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'טקסט' }], { listThrows: true });
    const adapter = createSearchAdapter(host as never);

    const outcome = await adapter.find('טקסט', 'next');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok || outcome.reason).toBe('threw');
  });

  it('דפדוף על פני 501 בלוקים — מופע יחיד בבלוק ה-501 עדיין נמצא', async () => {
    const blocks: FakeBlock[] = Array.from({ length: 501 }, (_, i) => ({
      nodeId: `p${i}`,
      text: i === 500 ? 'zzq' : 'אין כאן כלום',
    }));
    const { host } = createFakeHost(blocks);
    const adapter = createSearchAdapter(host as never);

    const outcome = await adapter.find('zzq', 'next');

    expect(outcome.ok).toBe(true);
    expect(outcome.ok && outcome.snapshot.total).toBe(1);
  });
});

describe('createSearchAdapter — החלפה', () => {
  it('replace מחליף רק את המופע הפעיל, בבלוק הנכון', async () => {
    const { host, currentTexts } = createFakeHost([
      { nodeId: 'p0', text: 'zzq אחד' },
      { nodeId: 'p1', text: 'zzq שתיים' },
    ]);
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const outcome = await adapter.replace('YYY');

    expect(outcome.ok).toBe(true);
    const texts = currentTexts();
    expect(texts.p0).toBe('YYY אחד');
    expect(texts.p1).toBe('zzq שתיים');
  });

  it('אחרי replace, ההתאמה הבאה תופסת את אותו מספור — התקדמות חינם', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'zzq zzq zzq' }]);
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const outcome = await adapter.replace('YYYYYYY');

    expect(outcome.ok).toBe(true);
    expect(outcome.ok && outcome.snapshot.total).toBe(2);
    expect(outcome.ok && outcome.snapshot.activeIndex).toBe(0);
  });

  it('בלי שאילתה אין החלפה, ואין קריאה למנוע', async () => {
    const { host, calls } = createFakeHost([{ nodeId: 'p0', text: 'טקסט' }]);
    const adapter = createSearchAdapter(host as never);

    const outcome = await adapter.replace('אחרי');

    expect(outcome).toEqual({ ok: false, message: NO_QUERY_TEXT, reason: 'no-query' });
    expect(calls).toEqual([]);
  });

  it('שאילתה בלי התאמות מדווחת "אין התאמות"', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'שלום' }]);
    const adapter = createSearchAdapter(host as never);

    await adapter.find('אשכנז', 'next');
    const outcome = await adapter.replace('אחרי');

    expect(outcome).toEqual({ ok: false, message: NO_MATCHES_TEXT, reason: 'no-matches' });
  });

  it('כשל אמיתי של replace (מסמך לקריאה בלבד וכו׳) מגיע כהודעה בעברית', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'zzq' }], {
      onReplace: () => ({ success: false, failure: { code: 'PERMISSION_DENIED' } }),
    });
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const outcome = await adapter.replace('אחרי');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok || outcome.reason).toBe('PERMISSION_DENIED');
    expect(outcome.ok || outcome.message).toContain('הרשאה');
  });

  it('replaceAll מחליף כמה מופעים באותו בלוק (נתיב אטומי), בלי לפגוע זה בזה', async () => {
    const { host, currentTexts, calls } = createFakeHost([{ nodeId: 'p0', text: 'cat cat cat' }]);
    const adapter = createSearchAdapter(host as never);

    await adapter.find('cat', 'next');
    const outcome = await adapter.replaceAll('dog');

    expect(outcome.ok).toBe(true);
    expect(currentTexts().p0).toBe('dog dog dog');
    expect(calls.filter((c) => c.name === 'mutations.apply')).toHaveLength(1);
  });

  it('replaceAll — כשאין mutations.apply, הנתיב הנקודתי מחליף מהסוף להתחלה בתוך כל בלוק', async () => {
    const { host, currentTexts, calls } = createFakeHost([{ nodeId: 'p0', text: 'cat cat cat' }], {
      noMutationsApi: true,
    });
    const adapter = createSearchAdapter(host as never);

    await adapter.find('cat', 'next');
    const outcome = await adapter.replaceAll('dog');

    expect(outcome.ok).toBe(true);
    expect(currentTexts().p0).toBe('dog dog dog');
    expect(calls.filter((c) => c.name === 'replace')).toHaveLength(3);
  });

  it('replaceAll — כשאין mutations.apply בכלל, שמונה הפסקאות עדיין מוחלפות כולן (נתיב נקודתי)', async () => {
    const { host, currentTexts, calls } = createFakeHost(eightParagraphBlocks(), { noMutationsApi: true });
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const outcome = await adapter.replaceAll('YYY');

    expect(outcome.ok).toBe(true);
    expect(outcome.ok && outcome.snapshot.total).toBe(0);
    expect(calls.filter((c) => c.name === 'replace')).toHaveLength(8);
    const texts = currentTexts();
    for (let i = 0; i < 8; i++) {
      expect(texts[`p${i}`]).toBe(`YYY paragraph ${i}`);
    }
  });

  it('replaceAll — mutations.apply שזורק נופל לנתיב הנקודתי, בלי לאבד אף מופע', async () => {
    const { host, currentTexts, calls } = createFakeHost(eightParagraphBlocks(), {
      mutationsApplyThrows: true,
    });
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const outcome = await adapter.replaceAll('YYY');

    expect(outcome.ok).toBe(true);
    expect(calls.filter((c) => c.name === 'mutations.apply')).toHaveLength(1);
    expect(calls.filter((c) => c.name === 'replace')).toHaveLength(8);
    const texts = currentTexts();
    for (let i = 0; i < 8; i++) {
      expect(texts[`p${i}`]).toBe(`YYY paragraph ${i}`);
    }
  });

  it('replaceAll — mutations.apply שמדווח כשל (success:false) נופל לנתיב הנקודתי', async () => {
    const { host, currentTexts, calls } = createFakeHost(eightParagraphBlocks(), {
      mutationsApplyFailure: { code: 'PRECONDITION_FAILED' },
    });
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const outcome = await adapter.replaceAll('YYY');

    expect(outcome.ok).toBe(true);
    expect(calls.filter((c) => c.name === 'replace')).toHaveLength(8);
    expect(currentTexts().p3).toBe('YYY paragraph 3');
  });

  it('replaceAll — בלוק בודד עם matchCount שגוי מדווח: הצ׳אנק כולו לא "מטופל", אבל שום מופע לא אבד', async () => {
    // חוסר-התאמה בצעד אחד מספיק כדי שכל הצ'אנק לא יסומן כ"טופל" (ראו
    // applyAtomicRewriteChunks: `allMatch` הוא כולל, לא לכל צעד בנפרד) —
    // אבל ה-doc המדומה כבר ביצע בפועל את שאר הצעדים בתוך אותה קריאה, ולכן
    // הרענון הטרי שאחרי רואה את p3 בלבד כ"עדיין עם zzq", והנתיב הנקודתי
    // מתקן רק אותו. זו בדיוק הנקודה: לא לבטוח בדיווח, אבל גם לא לאבד או
    // לשכפל עבודה שכבר קרתה במציאות.
    const { host, currentTexts, calls } = createFakeHost(eightParagraphBlocks(), {
      mutationsMismatchBlockIds: new Set(['p3']),
    });
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const outcome = await adapter.replaceAll('YYY');

    expect(outcome.ok).toBe(true);
    expect(calls.filter((c) => c.name === 'mutations.apply')).toHaveLength(1);
    expect(calls.filter((c) => c.name === 'replace')).toHaveLength(1);
    const texts = currentTexts();
    for (let i = 0; i < 8; i++) {
      expect(texts[`p${i}`]).toBe(`YYY paragraph ${i}`);
    }
  });

  it('replaceAll בלי התאמות מדווח "אין התאמות" ולא מריץ אף replace', async () => {
    const { host, calls } = createFakeHost([{ nodeId: 'p0', text: 'שלום' }]);
    const adapter = createSearchAdapter(host as never);

    await adapter.find('אשכנז', 'next');
    calls.length = 0;
    const outcome = await adapter.replaceAll('אחרי');

    expect(outcome).toEqual({ ok: false, message: NO_MATCHES_TEXT, reason: 'no-matches' });
    expect(calls.filter((c) => c.name === 'replace')).toEqual([]);
  });

  it('replace שממתין ל-Promise: המצב מוחזק isReplacing עד שההחלפה מתיישבת', async () => {
    const { host, settlePendingReplace, waitForPendingReplace } = createFakeHost(
      [{ nodeId: 'p0', text: 'zzq' }],
      { deferReplace: true },
    );
    const adapter = createSearchAdapter(host as never);
    const states: boolean[] = [];
    adapter.subscribe((state) => states.push(state.isReplacing));

    await adapter.find('zzq', 'next');
    const pending = adapter.replace('YYY');

    // מסומן מיד — לפני שהקריאה בכלל הגיעה ל-doc.replace. ראו ההערה על
    // מוקד ה-race ב-mutateOne/mutateAll (engine/search.ts).
    expect(adapter.getState().isReplacing).toBe(true);
    await waitForPendingReplace();
    settlePendingReplace();
    const outcome = await pending;

    expect(outcome.ok).toBe(true);
    expect(adapter.getState().isReplacing).toBe(false);
    expect(states).toContain(true);
    expect(states[states.length - 1]).toBe(false);
    // ה-outcome עצמו, לא רק adapter.getState() (שכבר "טרי" בכל מקרה
    // אחרי ש-ה-Promise התיישב) — זה בדיוק השדה שקפא על `true` בבאג
    // שנמדד: snapshot() חושב בתוך ה-try, לפני שה-finally כיבה את
    // isReplacing, וה-outcome המוחזר (ש-App.vue דורס איתו את
    // searchState.value בהצלחה) נשא ערך מיושן.
    expect(outcome.ok && outcome.snapshot.isReplacing).toBe(false);
  });

  /**
   * הרגרסייה שנמדדה בדפדפן אמיתי: "החלף"/"החלף הכל" מוצלח אחד השאיר את
   * כפתורי ההחלפה מנוטרלים לצמיתות — עד לפעולה לא-קשורה (הקלדה, "מצא
   * הבא") שהזדמן ל-emit() טרי שתיקן את זה בטעות. השורש: `snapshot()`
   * חושב *בתוך* ה-try, לפני שה-`finally` הופך את `isReplacing` ל-false
   * — ה-outcome המוחזר קפא עם `isReplacing:true`, ו-App.vue (`reportReplace`)
   * דורס את `searchState.value` הנכון (שה-emit() הפנימי כבר עדכן) בחזרה
   * לערך הקפוא. הבדיקות כאן קוראות את ה-outcome המוחזר עצמו — לא
   * `adapter.getState()` — כי זה בדיוק מה שהחמיץ את זה קודם.
   */
  it('replace מוצלח (בלי עיכוב מלאכותי) מחזיר outcome עם isReplacing:false — לא רק state.getState()', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'zzq' }]);
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const outcome = await adapter.replace('YYY');

    expect(outcome.ok).toBe(true);
    expect(outcome.ok && outcome.snapshot.isReplacing).toBe(false);
  });

  it('replaceAll מוצלח מחזיר outcome עם isReplacing:false — הכפתורים לא נשארים מנוטרלים', async () => {
    const { host } = createFakeHost(eightParagraphBlocks());
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const outcome = await adapter.replaceAll('YYY');

    expect(outcome.ok).toBe(true);
    expect(outcome.ok && outcome.snapshot.isReplacing).toBe(false);
  });

  it('החלפה שנייה בזמן שהראשונה רצה אינה נשלחת למנוע', async () => {
    const { host, calls, settlePendingReplace, waitForPendingReplace } = createFakeHost(
      [{ nodeId: 'p0', text: 'zzq' }],
      { deferReplace: true },
    );
    const adapter = createSearchAdapter(host as never);

    await adapter.find('zzq', 'next');
    const first = adapter.replace('YYY');
    // השנייה נחסמת סינכרונית על `isReplacing` — עוד לפני שהראשונה הגיעה
    // ל-doc.replace בכלל (ראו ההערה ב-mutateOne). אין כאן race לחכות לו.
    const second = await adapter.replace('YYY');
    expect(second.ok).toBe(false);

    await waitForPendingReplace();
    expect(calls.filter((c) => c.name === 'replace')).toHaveLength(1);

    settlePendingReplace();
    await first;
  });

  describe('replaceControlsVisible — מה שהדיאלוג מרנדר', () => {
    const state = (patch: Partial<SearchState>): SearchState => ({
      ...idleSearchState(),
      available: true,
      ...patch,
    });

    it('אין doc.blocks במסמך הזה → אין פקדי החלפה', () => {
      expect(replaceControlsVisible(state({ available: false, total: 5, canReplace: true }))).toBe(false);
    });

    it('אין התאמות → הפקדים נשארים', () => {
      expect(replaceControlsVisible(state({ total: 0, canReplace: false }))).toBe(true);
    });

    it('יש התאמות → canReplace מכריע', () => {
      expect(replaceControlsVisible(state({ total: 5, canReplace: true }))).toBe(true);
      expect(replaceControlsVisible(state({ total: 5, canReplace: false }))).toBe(false);
    });
  });
});

describe('createSearchAdapter — יכולת', () => {
  it('canReplace דורש גם blocks.list וגם replace', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'x' }], { noReplaceApi: true });
    const adapter = createSearchAdapter(host as never);

    const state = adapter.getState();
    expect(state.available).toBe(true);
    expect(state.canReplace).toBe(false);
  });
});

describe('createSearchAdapter — השקטה של חיפוש בזמן הקלדה', () => {
  it('כמה הקשות רצופות מתלכדות לקריאת blocks.list אחת', async () => {
    vi.useFakeTimers();
    const { host, calls } = createFakeHost([{ nodeId: 'p0', text: 'בראשית' }]);
    const adapter = createSearchAdapter(host as never);
    const outcomes: SearchOutcome[] = [];

    adapter.findDebounced('ב', (o) => outcomes.push(o));
    adapter.findDebounced('בר', (o) => outcomes.push(o));
    adapter.findDebounced('ברא', (o) => outcomes.push(o));
    adapter.findDebounced('בראשית', (o) => outcomes.push(o));

    await vi.advanceTimersByTimeAsync(SEARCH_DEBOUNCE_MS - 1);
    expect(calls).toEqual([]);

    await vi.advanceTimersByTimeAsync(1);
    expect(calls.filter((c) => c.name === 'blocks.list')).toHaveLength(1);
    expect(outcomes).toHaveLength(1);
    expect(outcomes[0].ok && outcomes[0].snapshot.query).toBe('בראשית');
  });

  it('סגירת הדיאלוג מבטלת הקלדה שממתינה', async () => {
    vi.useFakeTimers();
    const { host, calls } = createFakeHost([{ nodeId: 'p0', text: 'בראשית' }]);
    const adapter = createSearchAdapter(host as never);

    adapter.findDebounced('בראשית', () => {});
    adapter.close();
    await vi.advanceTimersByTimeAsync(SEARCH_DEBOUNCE_MS * 2);

    expect(calls).toEqual([]);
  });

  it('פירוק האדפטר מבטל הקלדה שממתינה', async () => {
    vi.useFakeTimers();
    const { host, calls } = createFakeHost([{ nodeId: 'p0', text: 'בראשית' }]);
    const adapter = createSearchAdapter(host as never);

    adapter.findDebounced('בראשית', () => {});
    adapter.dispose();
    await vi.advanceTimersByTimeAsync(SEARCH_DEBOUNCE_MS * 2);

    expect(calls).toEqual([]);
  });
});

describe('createSearchAdapter — מצב והרשמה', () => {
  it('subscribe יורה מיד עם המצב הנוכחי, ואז על כל שינוי', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'zzq zzq' }]);
    const adapter = createSearchAdapter(host as never);
    const seen: number[] = [];

    const off = adapter.subscribe((state) => seen.push(state.total));
    await adapter.find('zzq', 'next');
    off();
    await adapter.find('אשכנז', 'next');

    expect(seen).toEqual([0, 2]);
  });

  it('open()/close() — open מצליח כש-available, close מנקה שאילתה והתאמות', async () => {
    const { host } = createFakeHost([{ nodeId: 'p0', text: 'zzq' }]);
    const adapter = createSearchAdapter(host as never);

    const opened = adapter.open();
    expect(opened.ok).toBe(true);
    expect(opened.ok && opened.snapshot.open).toBe(true);

    await adapter.find('zzq', 'next');
    adapter.close();

    const state = adapter.getState();
    expect(state.open).toBe(false);
    expect(state.query).toBe('');
    expect(state.total).toBe(0);
  });
});

describe('searchCounterText', () => {
  it('מונה התוצאות נגזר מ-total ומ-activeIndex', () => {
    expect(
      searchCounterText({ ...idleSearchState(), query: 'בראשית', total: 12, activeIndex: 2 }),
    ).toBe('3 מתוך 12');
  });

  it('אין התאמות — "אין תוצאות" ולא מונה ריק', () => {
    expect(searchCounterText({ ...idleSearchState(), query: 'אשכנז', total: 0 })).toBe('אין תוצאות');
  });

  it('בלי שאילתה אין מונה', () => {
    expect(searchCounterText(idleSearchState())).toBe('');
    expect(searchCounterText({ ...idleSearchState(), total: 12, activeIndex: 0 })).toBe('');
  });

  it('התאמות בלי התאמה פעילה מציגות את המספר, ולא "1 מתוך"', () => {
    expect(
      searchCounterText({ ...idleSearchState(), query: 'בראשית', total: 12, activeIndex: -1 }),
    ).toBe('12 תוצאות');
  });
});

/**
 * הבאג המקורי לא היה במודול הזה אלא בחיווט: `(ui as any).search?.find?.()`
 * בתוך `catch` ריק. הבדיקה הזאת שומרת על החיווט — מקום שאין בו כרגע תשתית
 * לבדיקת קומפוננטות.
 */
describe('החיווט ב-App.vue', () => {
  const app = readFileSync(join(process.cwd(), 'src', 'App.vue'), 'utf8');

  it('אין `as any` על ה-controller', () => {
    expect(app).not.toMatch(/as any/);
  });

  it('אין קריאה ל-search.find ואין catch ריק', () => {
    expect(app).not.toMatch(/\.find\?\.\(/);
    expect(app).not.toMatch(/catch\s*\{\s*(\/\/[^\n]*\n\s*)?\}/);
  });

  it('החיפוש עובר דרך האדפטר', () => {
    expect(app).toMatch(/createSearchAdapter/);
    expect(app).toMatch(/searchAdapter\?\.find\(/);
  });

  it('הדיאלוג מקבל את `replaceControlsVisible` ולא את `canReplace` של המנוע', () => {
    expect(app).toContain('replaceControlsVisible(searchState.value)');
    expect(app).toContain(':can-replace="canShowReplace"');
    expect(app).not.toContain(':can-replace="searchState.canReplace"');
  });

  it('"אין התאמות" אינו מגיע לשורת המצב כשגיאה', () => {
    expect(app).toContain("REPLACE_NOT_AN_ERROR = new Set(['no-matches', 'no-query'])");
    expect(app).toContain('REPLACE_NOT_AN_ERROR.has(outcome.reason');
  });
});
