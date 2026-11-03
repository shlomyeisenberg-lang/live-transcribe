/**
 * כיתובים — הוספה, עריכה והסרה של פסקת הכיתוב.
 *
 * שבע הטענות שנמדדות כאן, וכולן מקבעות ממצא שנמדד בדפדפן ולא הנחה (ההנמקה
 * המלאה, כולל הפירוק של ה-docx המיוצא, ב-engine/captions.ts):
 *
 * 1. **התווית נשלחת כמות שהיא, בעברית.** `SEQ איור \* ARABIC` נמדד בקובץ
 *    המיוצא, ולכן המודול אינו מתרגם, אינו ממפה לרשימה סגורה, ואינו מחליף
 *    ב-`Figure`. זו הנקודה שהפילה את טבלת המקורות בגל 6.
 * 2. **`captions.update` אינו נקרא לעולם.** נמדד שהוא מוסיף את הטקסט החדש
 *    על הישן („אלף” → „אלף: בית”), ולכן עריכה כאן היא `remove`+`insert`.
 *    הבדיקה דורשת שהעוגן נקרא **לפני** ההסרה, ושהמיקום נשמר.
 * 3. **כיתוב שאין לו שכן אינו נערך.** הסרה שהצליחה והוספה שנכשלה אחריה היא
 *    טקסט שנמחק בלי דרך חזרה, ולכן הסירוב קודם להסרה.
 * 4. **הוולידציה יושבת אצלנו, כי המנוע בולע.** תווית של רווחים בלבד נכתבת
 *    `SEQ "   "`, וירידת שורה נכתבת **גולמית לתוך קוד השדה** — שניהם
 *    `success: true`.
 * 5. **הפסקה מסובבת לימין-לשמאל.** המנוע כותב פסקת `Caption` בלי
 *    `<w:bidi/>` (נמדד ב-docx), והצעד הזה הוא מה שמשלים אותו — בלי להפיל
 *    את הפעולה כשהוא עצמו נכשל.
 * 6. **הרשימה שואבת עמודים עד `total`.** ספר עם מאות לוחות הוא התרחיש, לא
 *    מקרה קצה.
 * 7. **המודול לעולם אינו זורק.** חריגה, קבלה שנכשלה, ופעולה שאינה קיימת
 *    בגרסת המנוע — שלושתן `CommandOutcome`.
 */
import { describe, expect, it } from 'vitest';
import {
  CAPTION_LABELS,
  CAPTION_LABEL_HINT,
  DEFAULT_CAPTION_LABEL,
  captionDisplay,
  emptyCaptionDraft,
  emptyCaptionsState,
  insertCaption,
  listCaptions,
  normalizeCaptionLabel,
  normalizeCaptionText,
  readCaptionsState,
  removeCaption,
  updateCaption,
  type CaptionDraft,
  type CaptionsHost,
} from '../../src/engine/captions';

interface Call {
  op: string;
  input?: unknown;
}

interface FakeCaption {
  nodeId: string;
  label: string;
  text?: string;
  /** כיתוב שהמנוע מחזיר בלי כתובת — מה שאסור להציג. */
  addressless?: boolean;
  /** כתובת עם `nodeId` ריק. אותה מלכודת, בצורה שנייה. */
  emptyId?: boolean;
}

interface FakeOptions {
  /** הכיתובים שבמסמך, בסדר הופעתם. */
  captions?: readonly FakeCaption[];
  /**
   * מזהי הבלוקים של המסמך, בסדרם. ברירת המחדל: `block-1` ואחריו הכיתובים —
   * כלומר לכל כיתוב יש שכן קודם.
   *
   * מזהה שמתחיל ב-`tbl:` מדווח `nodeType: 'table'` ומזהה שמתחיל ב-`toc:`
   * מדווח `tableOfContents`, כמו במנוע האמיתי. הראשון הוא המסלול של הכיתוב
   * שמתחת ללוח — הצורה השכיחה של כיתוב — והשני הוא הסוג שהמנוע עדיין דוחה.
   * בלעדיהם הכפיל היה מסמך של פסקאות בלבד, ולא היה מודד אף אחד מהשניים.
   */
  blocks?: readonly string[];
  /** `total` שהמנוע מדווח, כשהוא שונה מאורך העמוד — כלומר שאיבת עמודים. */
  captionsTotal?: number;
  /** מה `selection.current` מדווח. `null` = אין בחירה. */
  caret?: { blockId: string } | null;
  /** בחירה רב-מקטעית, כמו שהמנוע מדווח על בחירה שחוצה פסקאות. גובר על `caret`. */
  segments?: readonly { blockId: string }[];
  failures?: Record<string, { code: string; message?: string }>;
  /**
   * כשל **בקריאה הראשונה בלבד** לאותה פעולה. זה מה שמפריד בין „ההוספה
   * נכשלה והשחזור הצליח” ל„גם השחזור נכשל”, ובלעדיו שתי ההתנהגויות היו
   * נראות אותו דבר.
   */
  failOnce?: Record<string, { code: string; message?: string }>;
  throws?: readonly string[];
  missing?: readonly string[];
}

function fakeEngine(options: FakeOptions = {}) {
  const calls: Call[] = [];
  const missing = new Set(options.missing ?? []);
  const throwing = new Set(options.throws ?? []);
  const failures = options.failures ?? {};
  const pendingOnce = new Map(Object.entries(options.failOnce ?? {}));

  /** הכשל שנשלף פעם אחת, או `undefined`. */
  function onceFailure(op: string): { code: string; message?: string } | undefined {
    const failure = pendingOnce.get(op);
    if (failure) pendingOnce.delete(op);
    return failure;
  }

  function route<T>(op: string, impl: (input: unknown) => T): ((input: unknown) => T) | undefined {
    if (missing.has(op)) return undefined;
    return (input: unknown) => {
      calls.push({ op, input });
      if (throwing.has(op)) throw new Error(`${op} התפוצץ`);
      return impl(input);
    };
  }

  function receipt(op: string): unknown {
    const failure = onceFailure(op) ?? failures[op];
    return failure ? { success: false, failure } : { success: true };
  }

  const captions = [...(options.captions ?? [])];
  const blockIds = [...(options.blocks ?? ['block-1', ...captions.map((item) => item.nodeId)])];
  const caret = options.caret === undefined ? { blockId: 'block-1' } : options.caret;

  /**
   * הסוג שהמנוע מדווח על בלוק: `tbl:*` טבלה, `toc:*` תוכן עניינים, וכל
   * השאר פסקה. מזהה שאינו במסמך מחזיר `null` — כך „לא נמצא” נשאר „לא נמצא”
   * ואינו הופך בשקט לפסקה.
   */
  function blockTypeOf(nodeId: string): string | null {
    if (!blockIds.includes(nodeId)) return null;
    if (nodeId.startsWith('tbl:')) return 'table';
    if (nodeId.startsWith('toc:')) return 'tableOfContents';
    return 'paragraph';
  }

  /** עמוד תחת `limit`/`offset`, כמו `DiscoveryOutput` האמיתי. */
  function page<T>(items: readonly T[], total: number, input: unknown) {
    const query = (input ?? {}) as { limit?: number; offset?: number };
    const offset = query.offset ?? 0;
    const end = query.limit === undefined ? undefined : offset + query.limit;
    return { items: items.slice(offset, end), total };
  }

  const doc = {
    selection: {
      current: route('selection.current', () => {
        const segments = options.segments ?? (caret ? [{ blockId: caret.blockId }] : []);
        return { target: segments.length > 0 ? { segments } : null };
      }),
    },
    blocks: {
      list: route('blocks.list', (input) => {
        const query = (input ?? {}) as { limit?: number; offset?: number };
        const offset = query.offset ?? 0;
        const end = query.limit === undefined ? undefined : offset + query.limit;
        return {
          blocks: blockIds.slice(offset, end).map((nodeId) => ({ nodeId, nodeType: blockTypeOf(nodeId) })),
          total: blockIds.length,
        };
      }),
    },
    paragraphs: {
      setDirection: route('paragraphs.setDirection', () => receipt('paragraphs.setDirection')),
    },
    captions: {
      list: route('captions.list', (input) => {
        const counters = new Map<string, number>();
        const items = captions.map((item) => {
          const next = (counters.get(item.label) ?? 0) + 1;
          counters.set(item.label, next);
          return {
            id: item.nodeId,
            address: item.addressless
              ? {}
              : {
                  kind: 'block',
                  nodeType: 'paragraph',
                  nodeId: item.emptyId ? '' : item.nodeId,
                },
            label: item.label,
            number: next,
            text: item.text ?? '',
            instruction: `SEQ ${item.label} \\* ARABIC`,
          };
        });
        return page(items, options.captionsTotal ?? items.length, input);
      }),
      insert: route('captions.insert', (input) => {
        const failure = onceFailure('captions.insert') ?? failures['captions.insert'];
        if (failure) return { success: false, failure };
        /*
          הכתובת נבדקת כאן, ולא נבלעת: נמדד על 2.15.0-next.15 שהמנוע דורש את
          הסוג **האמיתי** של הבלוק — טבלה שנשלחה כ-`paragraph` ופסקה שנשלחה
          כ-`table` הוחזרו שתיהן `TARGET_NOT_FOUND` — ושבלוק `tableOfContents`
          נדחה `INVALID_TARGET` גם עם סוגו הנכון. כפיל שמקבל כל כתובת היה
          מאשר בירוק בדיוק את השגיאה שהבדיקות האלה נכתבו בשבילה.
        */
        const { adjacentTo } = (input ?? {}) as { adjacentTo?: { nodeId?: string; nodeType?: string } };
        const nodeId = adjacentTo?.nodeId ?? '';
        const actual = blockTypeOf(nodeId);
        if (actual === null || actual !== adjacentTo?.nodeType) {
          return {
            success: false,
            failure: {
              code: 'TARGET_NOT_FOUND',
              message: `target ${adjacentTo?.nodeType ?? 'paragraph'} ${nodeId} was not found`,
            },
          };
        }
        if (actual !== 'paragraph' && actual !== 'table') {
          return {
            success: false,
            failure: { code: 'INVALID_TARGET', message: `received ${actual}.` },
          };
        }
        return { success: true, caption: { kind: 'block', nodeType: 'paragraph', nodeId: 'caption-new' } };
      }),
      remove: route('captions.remove', () => receipt('captions.remove')),
      /**
       * קיים בכפיל **בכוונה**, ואינו אמור להיקרא לעולם: זו הפעולה שמוסיפה
       * את הטקסט החדש על הישן. כפיל בלעדיה היה הופך „לא נקרא” ל„לא קיים”,
       * ואת הבדיקה לחסרת ערך.
       */
      update: route('captions.update', () => receipt('captions.update')),
      configure: route('captions.configure', () => receipt('captions.configure')),
    },
  };

  return {
    host: { activeEditor: { doc } } as CaptionsHost,
    calls,
    ops: () => calls.map((call) => call.op),
    inputs: (op: string) => calls.filter((call) => call.op === op).map((call) => call.input),
  };
}

const DRAFT: CaptionDraft = { label: 'איור', text: 'שרטוט המשכן', position: 'below' };

describe('הוספת כיתוב', () => {
  it('שולחת את התווית העברית כמות שהיא, ואת הפסקה שבסמן כעוגן', async () => {
    const engine = fakeEngine();

    expect(await insertCaption(engine.host, DRAFT)).toEqual({ ok: true });
    expect(engine.inputs('captions.insert')).toEqual([
      {
        adjacentTo: { kind: 'block', nodeType: 'paragraph', nodeId: 'block-1' },
        position: 'below',
        label: 'איור',
        text: 'שרטוט המשכן',
      },
    ]);
  });

  it('„מעל” נשלח כ-`above`, ואינו נופל לברירת המחדל', async () => {
    const engine = fakeEngine();
    await insertCaption(engine.host, { ...DRAFT, position: 'above' });
    expect((engine.inputs('captions.insert')[0] as { position: string }).position).toBe('above');
  });

  it('מסובבת את פסקת הכיתוב לימין-לשמאל, לפי הכתובת שהקבלה החזירה', async () => {
    const engine = fakeEngine();
    await insertCaption(engine.host, DRAFT);
    expect(engine.inputs('paragraphs.setDirection')).toEqual([
      { target: { kind: 'block', nodeType: 'paragraph', nodeId: 'caption-new' }, direction: 'rtl' },
    ]);
  });

  it('כשל של סיבוב הפסקה אינו מפיל את הכיתוב שכבר נכתב', async () => {
    const engine = fakeEngine({ failures: { 'paragraphs.setDirection': { code: 'NO_OP' } } });
    expect(await insertCaption(engine.host, DRAFT)).toEqual({ ok: true });

    const missing = fakeEngine({ missing: ['paragraphs.setDirection'] });
    expect(await insertCaption(missing.host, DRAFT)).toEqual({ ok: true });

    const thrown = fakeEngine({ throws: ['paragraphs.setDirection'] });
    expect(await insertCaption(thrown.host, DRAFT)).toEqual({ ok: true });
  });

  it('לוקחת את המקטע האחרון בבחירה שחוצה פסקאות, ולא את הראשון', async () => {
    const engine = fakeEngine({ segments: [{ blockId: 'block-1' }, { blockId: 'block-9' }] });
    await insertCaption(engine.host, DRAFT);
    expect((engine.inputs('captions.insert')[0] as { adjacentTo: { nodeId: string } }).adjacentTo.nodeId).toBe(
      'block-9',
    );
  });

  it('בלי סמן — מסבירה, ואינה נוגעת במסמך', async () => {
    const engine = fakeEngine({ caret: null });
    const outcome = await insertCaption(engine.host, DRAFT);

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.message).toContain('יש למקם את הסמן');
    expect(engine.ops()).not.toContain('captions.insert');
  });
});

describe('הוולידציה שהמנוע אינו עושה', () => {
  it('תווית של רווחים בלבד נדחית לפני הקריאה', async () => {
    const engine = fakeEngine();
    const outcome = await insertCaption(engine.host, { ...DRAFT, label: '   ' });

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.message).toContain(CAPTION_LABEL_HINT);
    expect(engine.ops()).not.toContain('captions.insert');
  });

  it('ירידת שורה בתווית נדחית — היא נכתבת גולמית לתוך קוד השדה', async () => {
    const engine = fakeEngine();
    const outcome = await insertCaption(engine.host, { ...DRAFT, label: 'אי\nור' });

    expect(outcome.ok).toBe(false);
    expect(engine.ops()).not.toContain('captions.insert');
  });

  it('ירידת שורה בטקסט מכווצת לרווח, ואינה חוסמת', async () => {
    const engine = fakeEngine();
    expect(await insertCaption(engine.host, { ...DRAFT, text: ' שורה\nשנייה ' })).toEqual({ ok: true });
    expect((engine.inputs('captions.insert')[0] as { text: string }).text).toBe('שורה שנייה');
  });

  it('טקסט ריק הוא ערך חוקי — „איור 1” בלי תיאור', async () => {
    const engine = fakeEngine();
    expect(await insertCaption(engine.host, { ...DRAFT, text: '   ' })).toEqual({ ok: true });
    expect((engine.inputs('captions.insert')[0] as { text: string }).text).toBe('');
  });

  it('`normalizeCaptionLabel` ו-`normalizeCaptionText` הן ההכרעה, ולא הדיאלוג', () => {
    expect(normalizeCaptionLabel('  איור  ')).toBe('איור');
    expect(normalizeCaptionLabel('   ')).toBeNull();
    expect(normalizeCaptionLabel('')).toBeNull();
    expect(normalizeCaptionLabel('א\tב')).toBeNull();
    expect(normalizeCaptionLabel(7 as unknown as string)).toBeNull();
    // גרשיים ולוכסן **אינם** נדחים: נמדד שהמנוע מבריח אותם כהלכה
    // (`SEQ "א\"ב" \* ARABIC`), ודחייה כאן הייתה חוסמת תווית תקינה בעברית.
    expect(normalizeCaptionLabel('א"ב')).toBe('א"ב');
    expect(normalizeCaptionLabel('איור \\* MERGEFORMAT')).toBe('איור \\* MERGEFORMAT');

    expect(normalizeCaptionText(' א\r\nב ')).toBe('א ב');
    expect(normalizeCaptionText(5 as unknown as string)).toBe('');
  });
});

describe('עריכת כיתוב', () => {
  const EXISTING = [
    { nodeId: 'cap-1', label: 'איור', text: 'ראשון' },
    { nodeId: 'cap-2', label: 'איור', text: 'שני' },
  ];

  it('היא הסרה והוספה מחדש, ולעולם אינה קוראת ל-`captions.update`', async () => {
    const engine = fakeEngine({ captions: EXISTING });

    expect(
      await updateCaption(engine.host, 'cap-2', { label: 'איור', text: 'מתוקן', position: 'below' }),
    ).toEqual({ ok: true });

    expect(engine.ops()).not.toContain('captions.update');
    expect(engine.inputs('captions.remove')).toEqual([
      { target: { kind: 'block', nodeType: 'paragraph', nodeId: 'cap-2' } },
    ]);
    expect(engine.inputs('captions.insert')).toEqual([
      {
        adjacentTo: { kind: 'block', nodeType: 'paragraph', nodeId: 'cap-1' },
        position: 'below',
        label: 'איור',
        text: 'מתוקן',
      },
    ]);
  });

  it('קוראת את סדר הבלוקים **לפני** ההסרה — אחריה כבר אין ממה לגזור מיקום', async () => {
    const engine = fakeEngine({ captions: EXISTING });
    await updateCaption(engine.host, 'cap-1', { label: 'איור', text: 'x', position: 'below' });

    const ops = engine.ops();
    expect(ops.indexOf('blocks.list')).toBeLessThan(ops.indexOf('captions.remove'));
  });

  it('כיתוב שהוא הבלוק הראשון נצמד חזרה אל הבלוק שאחריו, מלמעלה', async () => {
    const engine = fakeEngine({
      captions: [{ nodeId: 'cap-1', label: 'איור' }],
      blocks: ['cap-1', 'block-1'],
    });
    await updateCaption(engine.host, 'cap-1', { label: 'איור', text: 'x', position: 'below' });

    expect(engine.inputs('captions.insert')).toEqual([
      {
        adjacentTo: { kind: 'block', nodeType: 'paragraph', nodeId: 'block-1' },
        position: 'above',
        label: 'איור',
        text: 'x',
      },
    ]);
  });

  it('כיתוב שהוא הבלוק היחיד — מסרבת, ואינה מוחקת מה שלא תוכל להחזיר', async () => {
    const engine = fakeEngine({ captions: [{ nodeId: 'cap-1', label: 'איור' }], blocks: ['cap-1'] });
    const outcome = await updateCaption(engine.host, 'cap-1', {
      label: 'איור',
      text: 'x',
      position: 'below',
    });

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('no-neighbour');
    expect(engine.ops()).not.toContain('captions.remove');
  });

  it('שכן שהוא טבלה — העוגן הוא הטבלה עצמה, עם הסוג האמיתי שלה', async () => {
    // הצורה שנמדדה בדפדפן, והשכיחה ביותר: פסקה, טבלה, הכיתוב שמתחתיה,
    // ופסקה. עד 2.14.0-next.5 `tbl:…` כעוגן הוחזר `TARGET_NOT_FOUND` והעוגן
    // היה הפסקה שאחרי; ב-2.15.0-next.15 נמדד שהעוגן מתקבל — ובלבד
    // ש-`nodeType` הוא `table` ולא `paragraph`.
    const engine = fakeEngine({
      captions: [{ nodeId: 'cap-1', label: 'טבלה', text: 'סדר הדורות' }],
      blocks: ['block-1', 'tbl:4122CC21', 'cap-1', 'block-2'],
    });
    const outcome = await updateCaption(engine.host, 'cap-1', {
      label: 'טבלה',
      text: 'סדר הדורות המתוקן',
      position: 'below',
    });

    expect(outcome).toEqual({ ok: true });
    expect(engine.inputs('captions.insert')).toEqual([
      {
        adjacentTo: { kind: 'block', nodeType: 'table', nodeId: 'tbl:4122CC21' },
        position: 'below',
        label: 'טבלה',
        text: 'סדר הדורות המתוקן',
      },
    ]);
  });

  it('שכן שהמנוע דוחה — נופלת לאחור אל הבלוק שאחרי, ומחזירה לאותו רווח', async () => {
    // תוכן עניינים כשכן: נמדד `INVALID_TARGET`, ושליחתו כ„פסקה” כותבת את
    // הכיתוב בתוך ה-sdt. לכן העוגן הוא הבלוק שאחרי עם „מעל” — אותו רווח.
    const engine = fakeEngine({
      captions: [{ nodeId: 'cap-1', label: 'איור' }],
      blocks: ['block-1', 'toc:4122CC30', 'cap-1', 'block-2'],
    });
    const outcome = await updateCaption(engine.host, 'cap-1', {
      label: 'איור',
      text: 'x',
      position: 'below',
    });

    expect(outcome).toEqual({ ok: true });
    expect(engine.inputs('captions.insert')).toEqual([
      {
        adjacentTo: { kind: 'block', nodeType: 'paragraph', nodeId: 'block-2' },
        position: 'above',
        label: 'איור',
        text: 'x',
      },
    ]);
  });

  it('הנפילה-לאחור אינה גוברת על פסקה שלפני הכיתוב', async () => {
    // המסלול הרגיל נשאר הרגיל: כששני השכנים פסקאות, העוגן הוא זה שלפני,
    // אחרת כיתוב באמצע מסמך היה נודד רווח אחד קדימה בכל עריכה.
    const engine = fakeEngine({
      captions: [{ nodeId: 'cap-1', label: 'איור' }],
      blocks: ['block-1', 'cap-1', 'block-2'],
    });
    await updateCaption(engine.host, 'cap-1', { label: 'איור', text: 'x', position: 'below' });

    expect(engine.inputs('captions.insert')).toEqual([
      {
        adjacentTo: { kind: 'block', nodeType: 'paragraph', nodeId: 'block-1' },
        position: 'below',
        label: 'איור',
        text: 'x',
      },
    ]);
  });

  it('טבלה משני הצדדים — נערך, ונשאר בין שתי הטבלאות', async () => {
    // המצב שסורב עד 2.14.0-next.5. נמדד בדפדפן על 2.15.0-next.15 במסמך
    // `טבלה │ כיתוב │ טבלה`: ההסרה וההוספה מחדש בעוגן הטבלה שלפני החזירו
    // את הכיתוב לאותו רווח, עם אותו מספר.
    const engine = fakeEngine({
      captions: [{ nodeId: 'cap-1', label: 'טבלה', text: 'סדר הדורות' }],
      blocks: ['tbl:4122CC21', 'cap-1', 'tbl:4122CC22'],
    });
    const outcome = await updateCaption(engine.host, 'cap-1', {
      label: 'טבלה',
      text: 'סדר הדורות המתוקן',
      position: 'below',
    });

    expect(outcome).toEqual({ ok: true });
    expect(engine.inputs('captions.insert')).toEqual([
      {
        adjacentTo: { kind: 'block', nodeType: 'table', nodeId: 'tbl:4122CC21' },
        position: 'below',
        label: 'טבלה',
        text: 'סדר הדורות המתוקן',
      },
    ]);
  });

  it('שני שכנים שהמנוע דוחה — מסרבת לפני שנגעה במסמך, ואינה מוחקת את הכיתוב', async () => {
    const engine = fakeEngine({
      captions: [{ nodeId: 'cap-1', label: 'איור', text: 'שרטוט המשכן' }],
      blocks: ['toc:4122CC30', 'cap-1', 'toc:4122CC31'],
    });
    const outcome = await updateCaption(engine.host, 'cap-1', {
      label: 'איור',
      text: 'שרטוט המשכן המתוקן',
      position: 'below',
    });

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('anchor-kind-unsupported');
    expect(outcome.ok === false && outcome.message).toContain('אינו פסקה ואינו טבלה');
    expect(engine.ops()).not.toContain('captions.remove');
    expect(engine.ops()).not.toContain('captions.insert');
  });

  it('הוספה שנכשלה אחרי הסרה שהצליחה מחזירה את הכיתוב הישן למקומו', async () => {
    const engine = fakeEngine({
      captions: EXISTING,
      failOnce: { 'captions.insert': { code: 'TARGET_NOT_FOUND' } },
    });
    const outcome = await updateCaption(engine.host, 'cap-2', {
      label: 'איור',
      text: 'מתוקן',
      position: 'below',
    });

    // ההוספה נכשלה, אבל המסמך חזר למה שהיה — ולכן ההודעה היא „נכשלה”
    // בלבד, בלי מילה על אובדן.
    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.message).not.toContain('הוסר מהמסמך');
    const inserts = engine.inputs('captions.insert') as { label: string; text: string }[];
    expect(inserts).toHaveLength(2);
    expect(inserts[1]).toMatchObject({ label: 'איור', text: 'שני' });
  });

  it('כשגם השחזור נכשל — ההודעה אומרת שהכיתוב הוסר, ומפנה לביטול', async () => {
    const engine = fakeEngine({
      captions: EXISTING,
      failures: { 'captions.insert': { code: 'TARGET_NOT_FOUND' } },
    });
    const outcome = await updateCaption(engine.host, 'cap-2', {
      label: 'איור',
      text: 'מתוקן',
      position: 'below',
    });

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('caption-lost');
    expect(outcome.ok === false && outcome.message).toContain('הוסר מהמסמך');
    expect(outcome.ok === false && outcome.message).toContain('Ctrl+Z');
  });

  it('מזהה שאינו במסמך — מסרבת לפני ההסרה, עם הודעה אחרת', async () => {
    const engine = fakeEngine({ captions: EXISTING });
    const outcome = await updateCaption(engine.host, 'cap-9', {
      label: 'איור',
      text: 'x',
      position: 'below',
    });

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('caption-not-found');
    expect(engine.ops()).not.toContain('captions.remove');
  });

  it('הסרה שנכשלה עוצרת — ואינה ממשיכה להוסיף כיתוב שני', async () => {
    const engine = fakeEngine({
      captions: EXISTING,
      failures: { 'captions.remove': { code: 'TARGET_NOT_FOUND' } },
    });
    const outcome = await updateCaption(engine.host, 'cap-2', {
      label: 'איור',
      text: 'x',
      position: 'below',
    });

    expect(outcome.ok).toBe(false);
    expect(engine.ops()).not.toContain('captions.insert');
  });

  it('תווית פסולה נדחית לפני שנגעו במסמך', async () => {
    const engine = fakeEngine({ captions: EXISTING });
    const outcome = await updateCaption(engine.host, 'cap-2', {
      label: '  ',
      text: 'x',
      position: 'below',
    });

    expect(outcome.ok).toBe(false);
    expect(engine.ops()).not.toContain('blocks.list');
    expect(engine.ops()).not.toContain('captions.remove');
  });
});

describe('הסרת כיתוב', () => {
  it('צעד אחד, בלי ניקוי שיירים דרך `blocks.*`', async () => {
    const engine = fakeEngine({ captions: [{ nodeId: 'cap-1', label: 'איור' }] });

    expect(await removeCaption(engine.host, 'cap-1')).toEqual({ ok: true });
    expect(engine.inputs('captions.remove')).toEqual([
      { target: { kind: 'block', nodeType: 'paragraph', nodeId: 'cap-1' } },
    ]);
    expect(engine.ops()).not.toContain('blocks.deleteRange');
  });

  it('בלי מזהה — מסבירה, ואינה שולחת כתובת ריקה למנוע', async () => {
    const engine = fakeEngine();
    const outcome = await removeCaption(engine.host, '');

    expect(outcome.ok).toBe(false);
    expect(engine.ops()).not.toContain('captions.remove');
  });
});

describe('קריאת הכיתובים', () => {
  it('שואבת עמודים עד `total`, ואינה עוצרת בעמוד הראשון', async () => {
    // 250 > `PAGE_SIZE` של המודול (200), ולכן העמוד הראשון אינו הכול.
    const many = Array.from({ length: 250 }, (_, index) => ({
      nodeId: `cap-${index}`,
      label: 'איור',
      text: `לוח ${index}`,
    }));
    const engine = fakeEngine({ captions: many });

    const captions = await listCaptions(engine.host);
    expect(captions).toHaveLength(250);
    expect(captions[249].id).toBe('cap-249');
    // שתי קריאות ולא אחת — 200 ואז 50. זו העדות לשאיבה עצמה.
    expect(engine.inputs('captions.list')).toEqual([
      { limit: 200, offset: 0 },
      { limit: 200, offset: 200 },
    ]);
  });

  it('מחשבת את התצוגה כמו שהיא נראית במסמך, ואוספת את התוויות שבשימוש', async () => {
    const engine = fakeEngine({
      captions: [
        { nodeId: 'cap-1', label: 'איור', text: 'שרטוט המשכן' },
        { nodeId: 'cap-2', label: 'טבלה', text: 'סדר הדורות' },
        { nodeId: 'cap-3', label: 'איור' },
      ],
    });

    const state = await readCaptionsState(engine.host);
    expect(state.captions.map((caption) => caption.display)).toEqual([
      'איור 1: שרטוט המשכן',
      'טבלה 1: סדר הדורות',
      'איור 2',
    ]);
    expect(state.labels).toEqual(['איור', 'טבלה']);
  });

  it('כיתוב בלי כתובת אינו מוצג — לחיצה עליו הייתה שולחת `undefined` להסרה', async () => {
    const engine = fakeEngine({
      captions: [
        { nodeId: 'cap-1', label: 'איור', addressless: true },
        { nodeId: 'cap-0', label: 'איור', emptyId: true },
        { nodeId: 'cap-2', label: 'איור' },
      ],
    });

    const captions = await listCaptions(engine.host);
    expect(captions.map((caption) => caption.id)).toEqual(['cap-2']);
  });

  it('כשל קריאה מחזיר „אין”, ואינו ממציא רשומות', async () => {
    const thrown = fakeEngine({ throws: ['captions.list'] });
    expect(await readCaptionsState(thrown.host)).toEqual(emptyCaptionsState());

    const missing = fakeEngine({ missing: ['captions.list'] });
    expect(await readCaptionsState(missing.host)).toEqual(emptyCaptionsState());

    expect(await readCaptionsState(null)).toEqual(emptyCaptionsState());
    expect(await listCaptions(undefined)).toEqual([]);
  });
});

describe('לעולם אינו זורק', () => {
  it('קבלה שנכשלה מתורגמת להודעה בעברית', async () => {
    const engine = fakeEngine({
      failures: { 'captions.insert': { code: 'INVALID_TARGET', message: 'nope' } },
    });
    const outcome = await insertCaption(engine.host, DRAFT);

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.message).toContain('הוספת הכיתוב נכשלה');
    expect(outcome.ok === false && outcome.reason).toBe('INVALID_TARGET');
  });

  it('חריגה שנזרקה מתורגמת ואינה יוצאת החוצה', async () => {
    const engine = fakeEngine({ throws: ['captions.insert'] });
    const outcome = await insertCaption(engine.host, DRAFT);

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('threw');
  });

  it('פעולה שאינה בגרסת המנוע מוחזרת בנוסח של §12', async () => {
    for (const [target, run] of [
      ['captions.insert', () => insertCaption(fakeEngine({ missing: ['captions.insert'] }).host, DRAFT)],
      ['captions.remove', () => removeCaption(fakeEngine({ missing: ['captions.remove'] }).host, 'cap-1')],
      ['blocks.list', () => updateCaption(fakeEngine({ missing: ['blocks.list'] }).host, 'cap-1', DRAFT)],
    ] as const) {
      const outcome = await run();
      expect(outcome.ok, target).toBe(false);
      expect(outcome.ok === false && outcome.message).toContain('אינו זמין בגרסה זו');
      expect(outcome.ok === false && outcome.reason).toBe('command-unsupported');
    }
  });

  it('מסמך שעדיין נטען אינו כשל של המנוע', async () => {
    for (const outcome of [
      await insertCaption(null, DRAFT),
      await updateCaption(undefined, 'cap-1', DRAFT),
      await removeCaption({} as CaptionsHost, 'cap-1'),
    ]) {
      expect(outcome.ok).toBe(false);
      expect(outcome.ok === false && outcome.reason).toBe('document-api-unavailable');
    }
  });

  it('`NO_OP` נחשב הצלחה, כמו בשאר המודולים', async () => {
    const engine = fakeEngine({ failures: { 'captions.remove': { code: 'NO_OP' } } });
    expect(await removeCaption(engine.host, 'cap-1')).toEqual({ ok: true });
  });
});

describe('המודל שהממשק עובד מולו', () => {
  it('ברירת המחדל היא התווית הראשונה של Word העברי, ומתחת לפסקה', () => {
    expect(CAPTION_LABELS).toEqual(['איור', 'טבלה', 'משוואה']);
    expect(DEFAULT_CAPTION_LABEL).toBe('איור');
    expect(emptyCaptionDraft()).toEqual({ label: 'איור', text: '', position: 'below' });
  });

  it('`captionDisplay` מוותר על הנקודתיים כשאין תיאור', () => {
    expect(captionDisplay('איור', 2, 'שרטוט')).toBe('איור 2: שרטוט');
    expect(captionDisplay('איור', 2, '')).toBe('איור 2');
  });
});
