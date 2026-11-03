/**
 * הפניות מקושרות — עדכון בלבד.
 *
 * אין כאן בדיקת הכנסה, ולא מפני שנשכחה: `crossRefs.insert` כותב קוד שדה שאינו
 * קוד Word ושהמנוע עצמו מחזיר עליו טקסט ריק אחרי `rebuild`. המדידה המלאה
 * וההנמקה ב-engine/cross-refs.ts, ומה שנבדק כאן הוא הצד שכן עובד.
 *
 * שתי הטענות שנמדדות:
 *
 * 1. **העדכון שואב עמודים עד `total`.** `CrossRefsListResult` הוא
 *    `DiscoveryOutput`; מימוש שרץ על העמוד הראשון היה משאיר הפניות מיושנות
 *    במסמך גדול — ומדווח „בוצע”.
 * 2. **המודול לעולם אינו זורק.** חריגה, קבלה שנכשלה, ופעולה שאינה קיימת
 *    בגרסת המנוע — שלושתן `CommandOutcome`.
 */
import { describe, expect, it } from 'vitest';
import {
  emptyCrossRefsState,
  readCrossRefsState,
  rebuildAllCrossRefs,
  type CrossRefsHost,
} from '../../src/engine/cross-refs';

interface Call {
  op: string;
  input?: unknown;
}

interface FakeOptions {
  /** מה `crossRefs.list` מחזיק — כל המסמך, לא עמוד. */
  crossRefs?: readonly { address?: unknown }[];
  total?: number;
  failures?: Record<string, { code: string; message?: string }>;
  throws?: readonly string[];
  missing?: readonly string[];
}

function fakeEngine(options: FakeOptions = {}) {
  const calls: Call[] = [];
  const missing = new Set(options.missing ?? []);
  const throwing = new Set(options.throws ?? []);
  const failures = options.failures ?? {};

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

  const doc = {
    crossRefs: {
      list: route('crossRefs.list', (input) => {
        const all = options.crossRefs ?? [];
        const query = (input ?? {}) as { limit?: number; offset?: number };
        const offset = query.offset ?? 0;
        const end = query.limit === undefined ? undefined : offset + query.limit;
        return {
          items: all.slice(offset, end),
          ...(options.total === undefined ? {} : { total: options.total }),
        };
      }),
      rebuild: route('crossRefs.rebuild', () => receipt('crossRefs.rebuild')),
    },
  };

  const host = { activeEditor: { doc } } as unknown as CrossRefsHost;
  const ops = (): string[] => calls.map((call) => call.op);
  const inputs = (op: string): unknown[] =>
    calls.filter((call) => call.op === op).map((call) => call.input);

  return { host, calls, ops, inputs };
}

/** כתובת הפניה, בצורה ש-`crossRefs.list` מחזיר וש-`rebuild` מקבל. */
const address = (offset: number) => ({
  kind: 'inline',
  nodeType: 'crossRef',
  anchor: { start: { blockId: 'block-1', offset }, end: { blockId: 'block-1', offset: offset + 1 } },
});

describe('עדכון ההפניות', () => {
  it('מעדכן כל הפניה, ומעביר את הכתובת שהרשימה החזירה', async () => {
    const engine = fakeEngine({ crossRefs: [{ address: address(0) }, { address: address(1) }] });

    expect(await rebuildAllCrossRefs(engine.host)).toEqual({ ok: true });

    expect(engine.inputs('crossRefs.rebuild')).toEqual([
      { target: address(0) },
      { target: address(1) },
    ]);
  });

  it('שואב עמודים עד `total` ולא עוצר בעמוד הראשון', async () => {
    const crossRefs = Array.from({ length: 320 }, (_, index) => ({ address: address(index) }));
    const engine = fakeEngine({ crossRefs, total: 320 });

    expect(await rebuildAllCrossRefs(engine.host)).toEqual({ ok: true });

    expect(engine.inputs('crossRefs.rebuild')).toHaveLength(320);
    expect(engine.inputs('crossRefs.list')).toEqual([
      { limit: 200, offset: 0 },
      { limit: 200, offset: 200 },
    ]);
  });

  it('הפניה בלי כתובת מדולגת ואינה נשלחת כ-`undefined`', async () => {
    const engine = fakeEngine({ crossRefs: [{}, { address: address(1) }], total: 2 });

    expect(await rebuildAllCrossRefs(engine.host)).toEqual({ ok: true });
    expect(engine.inputs('crossRefs.rebuild')).toEqual([{ target: address(1) }]);
  });

  it('מסמך בלי הפניות — הצלחה שקטה, בלי קריאת עדכון', async () => {
    const engine = fakeEngine();

    expect(await rebuildAllCrossRefs(engine.host)).toEqual({ ok: true });
    expect(engine.ops()).not.toContain('crossRefs.rebuild');
  });

  it('עוצר בכשל הראשון ואינו חוזר על אותה הודעה', async () => {
    const engine = fakeEngine({
      crossRefs: [{ address: address(0) }, { address: address(1) }],
      total: 2,
      failures: { 'crossRefs.rebuild': { code: 'PRECONDITION_FAILED' } },
    });

    const outcome = await rebuildAllCrossRefs(engine.host);

    expect(outcome.ok === false && outcome.reason).toBe('PRECONDITION_FAILED');
    expect(outcome.ok === false && outcome.message).toContain('עדכון ההפניות נכשל');
    expect(engine.inputs('crossRefs.rebuild')).toHaveLength(1);
  });

  it('`NO_OP` היא הצלחה — הפניה מעודכנת אינה טעונה עדכון', async () => {
    const engine = fakeEngine({
      crossRefs: [{ address: address(0) }],
      total: 1,
      failures: { 'crossRefs.rebuild': { code: 'NO_OP' } },
    });

    expect(await rebuildAllCrossRefs(engine.host)).toEqual({ ok: true });
  });

  it('חריגה שנזרקה אינה יוצאת החוצה', async () => {
    const engine = fakeEngine({ throws: ['crossRefs.list'] });

    const outcome = await rebuildAllCrossRefs(engine.host);

    expect(outcome.ok === false && outcome.reason).toBe('threw');
  });

  it('חריגה מ-`rebuild` עצמו עוצרת את הלולאה ואינה נבלעת', async () => {
    // הזריקה כאן היא מ-`rebuild` ולא מ-`list`, וזה מסלול אחר לגמרי: כשל
    // שנבלע היה מדלג על ההפניה, ממשיך לזו שאחריה, ומחזיר „בוצע” על מסמך
    // שנשארו בו הפניות מיושנות — הכשל השקט שהמודול הזה נבנה כדי לא לייצר.
    const engine = fakeEngine({
      crossRefs: [{ address: address(0) }, { address: address(1) }],
      total: 2,
      throws: ['crossRefs.rebuild'],
    });

    const outcome = await rebuildAllCrossRefs(engine.host);

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('threw');
    expect(outcome.ok === false && outcome.message).toContain('עדכון ההפניות נכשל');
    expect(engine.inputs('crossRefs.rebuild')).toHaveLength(1);
  });

  it('גרסה בלי `crossRefs.rebuild` — הנוסח שהתכנית קובעת', async () => {
    expect(await rebuildAllCrossRefs(fakeEngine({ missing: ['crossRefs.rebuild'] }).host)).toEqual({
      ok: false,
      message: 'עדכון ההפניות נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });

  it('בלי Document API כלל — כשל מוסבר ולא זריקה', async () => {
    expect(await rebuildAllCrossRefs(null)).toEqual({
      ok: false,
      message: 'עדכון ההפניות נכשל: המסמך עדיין נטען',
      reason: 'document-api-unavailable',
    });
  });
});

describe('מונה ההפניות', () => {
  it('סופר לפי `total` ולא לפי אורך העמוד', async () => {
    const crossRefs = Array.from({ length: 5 }, (_, index) => ({ address: address(index) }));
    const engine = fakeEngine({ crossRefs, total: 91 });

    expect(await readCrossRefsState(engine.host)).toEqual({ count: 91 });
  });

  it('בלי `total` נופל לאורך העמוד — מספר חלקי עדיף על אפס שקרי', async () => {
    const engine = fakeEngine({ crossRefs: [{ address: address(0) }, { address: address(1) }] });

    expect(await readCrossRefsState(engine.host)).toEqual({ count: 2 });
  });

  it('חריגה או פעולה חסרה — „אין הפניות”, ולא זריקה', async () => {
    expect(await readCrossRefsState(fakeEngine({ throws: ['crossRefs.list'] }).host)).toEqual(
      emptyCrossRefsState(),
    );
    expect(await readCrossRefsState(fakeEngine({ missing: ['crossRefs.list'] }).host)).toEqual({
      count: 0,
    });
    expect(await readCrossRefsState(null)).toEqual({ count: 0 });
  });
});
