/**
 * כותרת עליונה ותחתונה.
 *
 * שלוש טענות שהמודול עומד או נופל עליהן, ולכן הן עיקר מה שנמדד כאן:
 *
 * 1. **נוצר חלק אחד וכל המקטעים מפנים אליו.** יצירת חלק לכל מקטע הייתה
 *    מייצרת כותרות שונות במסמך שהמשתמש חושב שיש בו אחת.
 * 2. **הסרה מנקה לפני שהיא מוחקת.** מחיקת חלק שמקטע עדיין מפנה אליו משאירה
 *    הפניה שבורה ב-`sectPr`, וזה קובץ שלא ייפתח.
 * 3. **„עריכה” על מסמך שכבר יש בו כותרת אינה יוצרת שנייה.** זו ההצלחה
 *    השקטה — אין מוטציה, ואין גם הודעת כשל.
 */
import { describe, expect, it } from 'vitest';
import {
  emptyHeaderFooterState,
  ensureHeaderFooter,
  readHeaderFooterState,
  removeHeaderFooter,
  setDifferentFirstPage,
  setDifferentOddEvenPages,
  setLinkedToPrevious,
  type HeaderFooterHost,
} from '../../src/engine/header-footer';

interface Call {
  op: string;
  input?: unknown;
}

interface FakeOptions {
  /** כמה מקטעים במסמך. ברירת המחדל: אחד, כמו מסמך רגיל. */
  sections?: number;
  /** מה `resolve` מדווח. ברירת המחדל: אין כותרת. */
  status?: 'explicit' | 'inherited' | 'none';
  /** קבלה שנכשלת, לפי שם הפעולה. */
  failures?: Record<string, { code: string; message?: string }>;
  /** פעולות שזורקות. */
  throws?: readonly string[];
  /** מסלולים שאינם קיימים בפאסדה — גרסת מנוע שאין לה את היכולת. */
  missing?: readonly string[];
  /** מה `sections.list` מדווח על כל מקטע, מעבר לכתובת. */
  sectionState?: (index: number) => Record<string, unknown>;
  /** מה `parts.list` מחזיר. */
  parts?: readonly { refId: string; referencedBySections?: unknown[] }[];
  /** הקבלה של `parts.create` תחזיר את המזהה הזה. `null` = בלי מזהה כלל. */
  refId?: string | null;
  /**
   * `total` ש-`sections.list`/`headerFooters.parts.list` מדווחים, בדיוק כמו
   * `DiscoveryOutput` של המנוע. כשהוא מוגדר, הרשימה מדפדפת לפי `limit`/`offset`
   * שנשלחים בפועל — ובלעדיו (ברירת המחדל של כל שאר הבדיקות) היא מחזירה הכול
   * בעמוד הראשון, כמו שהמסמכים הקטנים שרוב הבדיקות משתמשות בהם מתנהגים.
   */
  sectionsTotal?: number;
  partsTotal?: number;
}

/** עמוד מתוך רשימה, בדיוק כמו `DiscoveryOutput` של המנוע. ראו toc.test.ts. */
function page<T>(all: readonly T[], input: unknown, total: number | undefined) {
  const query = (input ?? {}) as { limit?: number; offset?: number };
  const offset = query.offset ?? 0;
  const end = query.limit === undefined ? undefined : offset + query.limit;
  return { items: all.slice(offset, end), ...(total === undefined ? {} : { total }) };
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

  const count = options.sections ?? 1;
  const items = Array.from({ length: count }, (_, index) => ({
    address: { kind: 'section', sectionId: `s${index}` },
    ...(options.sectionState?.(index) ?? {}),
  }));

  const doc = {
    sections: {
      list: route('sections.list', (input) => page(items, input, options.sectionsTotal)),
      setTitlePage: route('sections.setTitlePage', () => receipt('sections.setTitlePage')),
      setOddEvenHeadersFooters: route('sections.setOddEvenHeadersFooters', () =>
        receipt('sections.setOddEvenHeadersFooters'),
      ),
    },
    headerFooters: {
      resolve: route('headerFooters.resolve', () => ({ status: options.status ?? 'none' })),
      refs: {
        set: route('headerFooters.refs.set', () => receipt('headerFooters.refs.set')),
        clear: route('headerFooters.refs.clear', () => receipt('headerFooters.refs.clear')),
        setLinkedToPrevious: route('headerFooters.refs.setLinkedToPrevious', () =>
          receipt('headerFooters.refs.setLinkedToPrevious'),
        ),
      },
      parts: {
        list: route('headerFooters.parts.list', (input) =>
          page(options.parts ?? [], input, options.partsTotal),
        ),
        create: route('headerFooters.parts.create', () => {
          const result = receipt('headerFooters.parts.create');
          if (!result.success) return result;
          const refId = options.refId === undefined ? 'rId9' : options.refId;
          return refId === null ? result : { ...result, refId };
        }),
        delete: route('headerFooters.parts.delete', () => receipt('headerFooters.parts.delete')),
      },
    },
  };

  const host = { activeEditor: { doc } } as unknown as HeaderFooterHost;
  const ops = (): string[] => calls.map((call) => call.op);
  const inputs = (op: string): unknown[] =>
    calls.filter((call) => call.op === op).map((call) => call.input);

  return { host, calls, ops, inputs };
}

describe('ensureHeaderFooter', () => {
  it('יוצרת חלק ריק אחד, ומפנה אליו כל מקטע בווריאנט default', async () => {
    const engine = fakeEngine({ sections: 2 });

    expect(await ensureHeaderFooter(engine.host, 'header')).toEqual({ ok: true });

    // חלק אחד בלבד — לא אחד לכל מקטע.
    expect(engine.inputs('headerFooters.parts.create')).toEqual([{ kind: 'header' }]);
    expect(engine.inputs('headerFooters.refs.set')).toEqual([
      {
        target: {
          kind: 'headerFooterSlot',
          section: { kind: 'section', sectionId: 's0' },
          headerFooterKind: 'header',
          variant: 'default',
        },
        refId: 'rId9',
      },
      {
        target: {
          kind: 'headerFooterSlot',
          section: { kind: 'section', sectionId: 's1' },
          headerFooterKind: 'header',
          variant: 'default',
        },
        refId: 'rId9',
      },
    ]);
  });

  it('אינה שולחת תוכן התחלתי — הכותרת נוצרת ריקה', async () => {
    const engine = fakeEngine();
    await ensureHeaderFooter(engine.host, 'footer');

    const [input] = engine.inputs('headerFooters.parts.create') as Record<string, unknown>[];
    // `sourceRefId` בחוץ = חלק ריק. שום טקסט אינו נשתל במסמך של מישהו אחר.
    expect(Object.keys(input!)).toEqual(['kind']);
    expect(input!.kind).toBe('footer');
  });

  it('כותרת שכבר קיימת — הצלחה בלי יצירה שנייה', async () => {
    for (const status of ['explicit', 'inherited'] as const) {
      const engine = fakeEngine({ status });

      expect(await ensureHeaderFooter(engine.host, 'header'), status).toEqual({ ok: true });
      expect(engine.ops(), status).not.toContain('headerFooters.parts.create');
    }
  });

  it('קבלה שהצליחה אך אין בה מזהה אינה מייצרת הפניה ריקה', async () => {
    const engine = fakeEngine({ refId: null });

    expect(await ensureHeaderFooter(engine.host, 'header')).toEqual({
      ok: false,
      message: 'הוספת הכותרת העליונה נכשלה: המנוע לא החזיר מזהה לכותרת שנוצרה',
      reason: 'missing-ref-id',
    });
    expect(engine.ops()).not.toContain('headerFooters.refs.set');
  });

  it('פעולה חסרה מדווחת „אינו זמין בגרסה זו” ואינה מנסה מסלול אחר', async () => {
    for (const op of ['headerFooters.parts.create', 'headerFooters.refs.set']) {
      const engine = fakeEngine({ missing: [op] });

      expect(await ensureHeaderFooter(engine.host, 'footer'), op).toEqual({
        ok: false,
        message: 'הוספת הכותרת התחתונה נכשלה: אינו זמין בגרסה זו',
        reason: 'command-unsupported',
      });
    }
  });

  it('אין Document API — „המסמך עדיין נטען”, ובלי חריגה', async () => {
    for (const host of [null, undefined, {}, { activeEditor: null }, { activeEditor: { doc: null } }]) {
      expect(await ensureHeaderFooter(host as HeaderFooterHost, 'header')).toEqual({
        ok: false,
        message: 'הוספת הכותרת העליונה נכשלה: המסמך עדיין נטען',
        reason: 'document-api-unavailable',
      });
    }
  });

  it('קבלה שנכשלה מתורגמת לעברית', async () => {
    const engine = fakeEngine({
      failures: { 'headerFooters.parts.create': { code: 'DOCUMENT_READONLY' } },
    });

    expect(await ensureHeaderFooter(engine.host, 'header')).toEqual({
      ok: false,
      message: 'הוספת הכותרת העליונה נכשלה: המסמך פתוח לקריאה בלבד',
      reason: 'DOCUMENT_READONLY',
    });
  });

  it('פעולה שזורקת מדווחת ואינה מפילה את התוסף', async () => {
    const engine = fakeEngine({ throws: ['headerFooters.parts.create'] });

    expect(await ensureHeaderFooter(engine.host, 'footer')).toEqual({
      ok: false,
      message: 'הוספת הכותרת התחתונה נכשלה: boom',
      reason: 'threw',
    });
  });

  it('מסמך בלי מקטעים אינו יעד לפעולה', async () => {
    const engine = fakeEngine({ sections: 0 });

    expect(await ensureHeaderFooter(engine.host, 'header')).toEqual({
      ok: false,
      message: 'הוספת הכותרת העליונה נכשלה: לא נמצא מקטע במסמך',
      reason: 'target-unresolved',
    });
  });
});

describe('removeHeaderFooter', () => {
  it('מנקה את שלושת הווריאנטים בכל מקטע, ורק אז מוחקת את החלק היתום', async () => {
    const engine = fakeEngine({
      sections: 2,
      status: 'explicit',
      parts: [{ refId: 'rId9', referencedBySections: [] }],
    });

    expect(await removeHeaderFooter(engine.host, 'header')).toEqual({ ok: true });

    const cleared = engine.inputs('headerFooters.refs.clear') as { target: { variant: string } }[];
    expect(cleared.map((call) => call.target.variant)).toEqual([
      'default',
      'first',
      'even',
      'default',
      'first',
      'even',
    ]);

    // הסדר הוא הטענה: מחיקה של חלק שעדיין מופנה אליו משאירה הפניה שבורה.
    const ops = engine.ops();
    expect(ops.lastIndexOf('headerFooters.refs.clear')).toBeLessThan(
      ops.indexOf('headerFooters.parts.delete'),
    );
    expect(engine.inputs('headerFooters.parts.delete')).toEqual([
      { target: { kind: 'headerFooterPart', refId: 'rId9' } },
    ]);
  });

  it('חלק שמקטע אחר עדיין מפנה אליו אינו נמחק', async () => {
    const engine = fakeEngine({
      parts: [{ refId: 'rId9', referencedBySections: [{ kind: 'section', sectionId: 's7' }] }],
    });

    expect(await removeHeaderFooter(engine.host, 'footer')).toEqual({ ok: true });
    expect(engine.ops()).not.toContain('headerFooters.parts.delete');
  });

  it('NO_OP על סלוט ריק הוא הצלחה — המשתמש ביקש שלא תהיה כותרת, ואין', async () => {
    const engine = fakeEngine({ failures: { 'headerFooters.refs.clear': { code: 'NO_OP' } } });

    expect(await removeHeaderFooter(engine.host, 'header')).toEqual({ ok: true });
  });

  it('גרסה בלי parts.delete עדיין מסירה את ההפניות', async () => {
    // ההפניות נוקו, וזה מה שהמשתמש ביקש. חלק יתום אינו שובר את המסמך.
    const engine = fakeEngine({ missing: ['headerFooters.parts.delete'] });

    expect(await removeHeaderFooter(engine.host, 'footer')).toEqual({ ok: true });
    expect(engine.ops()).toContain('headerFooters.refs.clear');
  });

  it('פעולה חסרה, קבלה שנכשלה וחריגה — כולן מדווחות ואינן מפילות', async () => {
    expect(await removeHeaderFooter(fakeEngine({ missing: ['headerFooters.refs.clear'] }).host, 'header')).toEqual({
      ok: false,
      message: 'הסרת הכותרת העליונה נכשלה: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });

    expect(
      await removeHeaderFooter(
        fakeEngine({ failures: { 'headerFooters.refs.clear': { code: 'LOCK_VIOLATION' } } }).host,
        'header',
      ),
    ).toEqual({
      ok: false,
      message: 'הסרת הכותרת העליונה נכשלה: החלק הזה במסמך מוגן מפני שינוי',
      reason: 'LOCK_VIOLATION',
    });

    expect(
      await removeHeaderFooter(fakeEngine({ throws: ['headerFooters.refs.clear'] }).host, 'footer'),
    ).toEqual({
      ok: false,
      message: 'הסרת הכותרת התחתונה נכשלה: boom',
      reason: 'threw',
    });
  });
});

describe('שונה בעמוד ראשון', () => {
  it('מוחל על כל מקטע, עם הכתובת שלו', async () => {
    const engine = fakeEngine({ sections: 2 });

    expect(await setDifferentFirstPage(engine.host, true)).toEqual({ ok: true });
    expect(engine.inputs('sections.setTitlePage')).toEqual([
      { target: { kind: 'section', sectionId: 's0' }, enabled: true },
      { target: { kind: 'section', sectionId: 's1' }, enabled: true },
    ]);
  });

  it('כיבוי הוא אותה פעולה עם false', async () => {
    const engine = fakeEngine();

    expect(await setDifferentFirstPage(engine.host, false)).toEqual({ ok: true });
    expect(engine.inputs('sections.setTitlePage')).toEqual([
      { target: { kind: 'section', sectionId: 's0' }, enabled: false },
    ]);
  });

  it('פעולה חסרה, קבלה שנכשלה וחריגה', async () => {
    expect(await setDifferentFirstPage(fakeEngine({ missing: ['sections.setTitlePage'] }).host, true)).toEqual({
      ok: false,
      message: 'שינוי „שונה בעמוד ראשון” נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });

    expect(
      await setDifferentFirstPage(
        fakeEngine({ failures: { 'sections.setTitlePage': { code: 'PERMISSION_DENIED' } } }).host,
        true,
      ),
    ).toEqual({
      ok: false,
      message: 'שינוי „שונה בעמוד ראשון” נכשל: אין הרשאה לבצע את הפעולה',
      reason: 'PERMISSION_DENIED',
    });

    expect(
      await setDifferentFirstPage(fakeEngine({ throws: ['sections.setTitlePage'] }).host, true),
    ).toEqual({
      ok: false,
      message: 'שינוי „שונה בעמוד ראשון” נכשל: boom',
      reason: 'threw',
    });
  });
});

describe('שונה בעמודים זוגיים ואי-זוגיים', () => {
  it('נשלחת פעם אחת בלי `target` — הדגל הוא ברמת המסמך', async () => {
    const engine = fakeEngine({ sections: 3 });

    expect(await setDifferentOddEvenPages(engine.host, true)).toEqual({ ok: true });
    expect(engine.inputs('sections.setOddEvenHeadersFooters')).toEqual([{ enabled: true }]);
    // אין קריאה למקטעים בכלל: אין על מה לעבור.
    expect(engine.ops()).not.toContain('sections.list');
  });

  it('פעולה חסרה, קבלה שנכשלה וחריגה', async () => {
    expect(
      await setDifferentOddEvenPages(
        fakeEngine({ missing: ['sections.setOddEvenHeadersFooters'] }).host,
        true,
      ),
    ).toEqual({
      ok: false,
      message: 'שינוי „שונה בעמודים זוגיים ואי-זוגיים” נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });

    expect(
      await setDifferentOddEvenPages(
        fakeEngine({
          failures: { 'sections.setOddEvenHeadersFooters': { code: 'DOCUMENT_READONLY' } },
        }).host,
        false,
      ),
    ).toEqual({
      ok: false,
      message: 'שינוי „שונה בעמודים זוגיים ואי-זוגיים” נכשל: המסמך פתוח לקריאה בלבד',
      reason: 'DOCUMENT_READONLY',
    });

    expect(
      await setDifferentOddEvenPages(
        fakeEngine({ throws: ['sections.setOddEvenHeadersFooters'] }).host,
        true,
      ),
    ).toEqual({
      ok: false,
      message: 'שינוי „שונה בעמודים זוגיים ואי-זוגיים” נכשל: boom',
      reason: 'threw',
    });
  });
});

describe('קשר לקודם', () => {
  it('מוחל על המקטעים שאחרי הראשון בלבד, ועל שני הסוגים', async () => {
    const engine = fakeEngine({ sections: 3 });

    expect(await setLinkedToPrevious(engine.host, true)).toEqual({ ok: true });

    const inputs = engine.inputs('headerFooters.refs.setLinkedToPrevious') as {
      target: { section: { sectionId: string }; headerFooterKind: string; variant: string };
      linked: boolean;
    }[];
    expect(
      inputs.map((call) => `${call.target.section.sectionId}:${call.target.headerFooterKind}`),
    ).toEqual(['s1:header', 's1:footer', 's2:header', 's2:footer']);
    // למקטע הראשון אין קודם, ולכן הוא אינו יעד.
    expect(inputs.some((call) => call.target.section.sectionId === 's0')).toBe(false);
    expect(inputs.every((call) => call.target.variant === 'default' && call.linked)).toBe(true);
  });

  it('מסמך בעל מקטע יחיד מקבל הסבר, ולא כשל של המנוע', async () => {
    const engine = fakeEngine({ sections: 1 });

    expect(await setLinkedToPrevious(engine.host, false)).toEqual({
      ok: false,
      message: 'שינוי הקישור לכותרת של המקטע הקודם נכשל: אין במסמך מקטע קודם לקשר אליו',
      reason: 'no-previous-section',
    });
    expect(engine.ops()).not.toContain('headerFooters.refs.setLinkedToPrevious');
  });

  it('פעולה חסרה, קבלה שנכשלה וחריגה', async () => {
    expect(
      await setLinkedToPrevious(
        fakeEngine({ missing: ['headerFooters.refs.setLinkedToPrevious'] }).host,
        true,
      ),
    ).toEqual({
      ok: false,
      message: 'שינוי הקישור לכותרת של המקטע הקודם נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });

    expect(
      await setLinkedToPrevious(
        fakeEngine({
          sections: 2,
          failures: { 'headerFooters.refs.setLinkedToPrevious': { code: 'INVALID_TARGET' } },
        }).host,
        true,
      ),
    ).toEqual({
      ok: false,
      message: 'שינוי הקישור לכותרת של המקטע הקודם נכשל: לא ניתן לבצע את הפעולה במקום הזה במסמך',
      reason: 'INVALID_TARGET',
    });

    expect(
      await setLinkedToPrevious(
        fakeEngine({ sections: 2, throws: ['headerFooters.refs.setLinkedToPrevious'] }).host,
        true,
      ),
    ).toEqual({
      ok: false,
      message: 'שינוי הקישור לכותרת של המקטע הקודם נכשל: boom',
      reason: 'threw',
    });
  });
});

describe('readHeaderFooterState', () => {
  it('קוראת את שני הסוגים דרך resolve, ואת המתגים מהמקטע', async () => {
    const engine = fakeEngine({
      status: 'inherited',
      sectionState: () => ({ titlePage: true, oddEvenHeadersFooters: true }),
    });

    expect(await readHeaderFooterState(engine.host)).toEqual({
      hasHeader: true,
      hasFooter: true,
      titlePage: true,
      oddEven: true,
      linkedToPrevious: true,
      sectionCount: 1,
    });
  });

  it('`status: none` = אין כותרת', async () => {
    const engine = fakeEngine();
    const state = await readHeaderFooterState(engine.host);

    expect(state.hasHeader).toBe(false);
    expect(state.hasFooter).toBe(false);
  });

  it('הפניה מפורשת במקטע שאינו הראשון = המקטע נותק מהקודם לו', async () => {
    const engine = fakeEngine({
      sections: 2,
      sectionState: (index) => (index === 1 ? { headerRefs: { default: 'rId5' } } : {}),
    });

    const state = await readHeaderFooterState(engine.host);

    expect(state.linkedToPrevious).toBe(false);
    expect(state.sectionCount).toBe(2);
  });

  it('כשל של הקריאה מחזיר את המצב הריק ולא „אולי”', async () => {
    for (const engine of [
      fakeEngine({ throws: ['sections.list'] }),
      fakeEngine({ missing: ['sections.list'] }),
      fakeEngine({ sections: 0 }),
    ]) {
      expect(await readHeaderFooterState(engine.host)).toEqual(emptyHeaderFooterState());
    }

    expect(await readHeaderFooterState(null)).toEqual(emptyHeaderFooterState());
  });

  it('גרסה בלי `resolve` אינה מפילה — המתגים נקראים, והכותרות מדווחות כחסרות', async () => {
    const engine = fakeEngine({
      missing: ['headerFooters.resolve'],
      status: 'explicit',
      sectionState: () => ({ titlePage: true }),
    });

    const state = await readHeaderFooterState(engine.host);

    expect(state.titlePage).toBe(true);
    expect(state.hasHeader).toBe(false);
  });

  /**
   * הבדיקה שמגנה על המסמך של המשתמש: בלי `resolve` אין דרך לדעת שכבר יש
   * כותרת, ולכן `ensureHeaderFooter` היה יוצר חלק ריק חדש ומפנה אליו את כל
   * המקטעים — מחיקה שקטה של הכותרת הקיימת. נכשל סגור.
   */
  it('בלי `resolve` היצירה אינה נוגעת במסמך אלא מדווחת „אינו זמין”', async () => {
    for (const kind of ['header', 'footer'] as const) {
      const engine = fakeEngine({ missing: ['headerFooters.resolve'], status: 'explicit' });

      const outcome = await ensureHeaderFooter(engine.host, kind);

      expect(outcome.ok).toBe(false);
      expect(outcome.ok === false && outcome.reason).toBe('command-unsupported');
      expect(engine.ops()).not.toContain('headerFooters.parts.create');
      expect(engine.ops()).not.toContain('headerFooters.refs.set');
    }
  });
});

describe('דפדוף על מקטעים וחלקים', () => {
  // באג 4: `sections.list` ו-`headerFooters.parts.list` הם `DiscoveryOutput`
  // (נמדד `limit: 250` — docs/engine-gaps.md), כלומר `items` הוא עמוד בלבד.
  // קריאה יחידה הייתה מחילה פעולה על 250 המקטעים/החלקים הראשונים ומדווחת
  // „בוצע” על מסמך גדול מזה.

  it('שואבת עמודי מקטעים עד `total`, ומחילה על כולם ולא רק על העמוד הראשון', async () => {
    const engine = fakeEngine({ sections: 320, sectionsTotal: 320 });

    expect(await setDifferentFirstPage(engine.host, true)).toEqual({ ok: true });

    expect(engine.inputs('sections.list')).toEqual([
      { limit: 200, offset: 0 },
      { limit: 200, offset: 200 },
    ]);
    expect(engine.inputs('sections.setTitlePage')).toHaveLength(320);
  });

  it('שואבת עמודי חלקים יתומים עד `total`, ומוחקת את כולם', async () => {
    const parts = Array.from({ length: 320 }, (_, index) => ({
      refId: `rId${index}`,
      referencedBySections: [],
    }));
    const engine = fakeEngine({ parts, partsTotal: 320 });

    expect(await removeHeaderFooter(engine.host, 'header')).toEqual({ ok: true });

    expect(engine.inputs('headerFooters.parts.list')).toEqual([
      { kind: 'header', limit: 200, offset: 0 },
      { kind: 'header', limit: 200, offset: 200 },
    ]);
    expect(engine.inputs('headerFooters.parts.delete')).toHaveLength(320);
  });
});
