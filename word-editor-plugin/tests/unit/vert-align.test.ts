/**
 * „כתב עליון” ו„כתב תחתי”. הבדיקה היא על **מה נשלח למנוע**, על ה-toggle ועל
 * הדיווח — ההחלה עצמה נבדקת בדפדפן.
 *
 * מה שנבדק כאן במיוחד:
 *   1. ה-toggle: `NO_OP` על ההחלה פירושו „זה כבר המצב”, ולכן נשלח `baseline`.
 *      זהו הדיווח היחיד על מצב שהמנוע נותן, ואין כאן state מקומי.
 *   2. היעד הוא `selectionTarget` ולא `target` — האחד הוא מודל הכתיבה, השני
 *      הוא `TextTarget` של תגובות.
 *   3. כל מסלול כשל מחזיר תוצאה מטופסת ולא זריקה: הוולידטורים של ה-Document
 *      API זורקים `INVALID_INPUT` על קלט פסול, וחריגה מפקד ב-Ribbon מפילה את
 *      רינדור הרצועה כולה.
 *
 * התבנית — כפיל שמקליט קלטים, ואימות מזהה הפעולה — היא זו של
 * tests/unit/page-break.test.ts.
 */
import { describe, expect, it, vi } from 'vitest';
import {
  VERT_ALIGN_OPERATION,
  readVertAlignSupport,
  toggleVertAlign,
  type VertAlignDocumentApi,
} from '../../src/engine/vert-align';

/** ה-`SelectionTarget` שהמנוע מקרין לבחירה חיה. */
const RANGE = {
  kind: 'selection',
  start: { kind: 'text', blockId: 'p1', offset: 2 },
  end: { kind: 'text', blockId: 'p1', offset: 6 },
};

type Receipt = { success?: boolean; failure?: { code: string; message?: string } };

function fakeDoc(
  options: {
    /** התשובה לכל קריאה, לפי הסדר. הערך האחרון חוזר על עצמו. */
    receipts?: Array<Receipt | (() => never)>;
    available?: boolean;
    capabilities?: VertAlignDocumentApi['capabilities'];
    selection?: unknown;
    omitFormat?: boolean;
    omitSelection?: boolean;
  } = {},
) {
  const calls: unknown[] = [];
  const receipts = options.receipts ?? [{ success: true }];
  let index = 0;

  const doc = {
    ...(options.omitSelection
      ? {}
      : {
          selection: {
            current: vi.fn(async () =>
              options.selection === undefined
                ? { empty: false, selectionTarget: RANGE }
                : options.selection,
            ),
          },
        }),
    capabilities:
      options.capabilities === undefined
        ? {
            get: async () => ({
              operations: { [VERT_ALIGN_OPERATION]: { available: options.available ?? true } },
            }),
          }
        : options.capabilities,
    ...(options.omitFormat
      ? {}
      : {
          format: {
            vertAlign: (input: unknown) => {
              calls.push(input);
              const next = receipts[Math.min(index, receipts.length - 1)];
              index += 1;
              return typeof next === 'function' ? next() : (next as never);
            },
          },
        }),
  } as unknown as VertAlignDocumentApi;

  return { doc, calls, host: { activeEditor: { doc } } };
}

const NO_OP: Receipt = { success: false, failure: { code: 'NO_OP', message: 'no changes' } };

describe('readVertAlignSupport', () => {
  it('זמין כשהמנוע מדווח שהפעולה זמינה', async () => {
    const { host } = fakeDoc();

    await expect(readVertAlignSupport(host)).resolves.toEqual({
      available: true,
      explanation: '',
    });
  });

  it('פעולה שאינה זמינה מקבלת את ההסבר של המנוע — ולא „אינו נתמך במנוע”', async () => {
    // זו כל הנקודה: הכפתור היה מנוטרל קשיח עם tooltip שהאשים את המנוע בחוסר
    // שאינו קיים. מעכשיו ההסבר בא מהמנוע, ורק כשהוא באמת חוסם.
    const { host } = fakeDoc({ available: false });

    await expect(readVertAlignSupport(host)).resolves.toEqual({
      available: false,
      explanation: 'הפעולה אינה זמינה בגרסה הזאת של המנוע',
    });
  });

  it('גרסה שאינה חושפת `format.vertAlign` מסומנת „אינו זמין בגרסה זו”', async () => {
    // נבדק לפני היכולות: מפת ה-operations נבנית מהקטלוג, ולכן גרסה שהסירה את
    // המימוש ועוד מכריזה על הפעולה הייתה מחזירה „זמין” לפקד שאין לו למה לקרוא.
    const { host } = fakeDoc({ omitFormat: true });

    await expect(readVertAlignSupport(host)).resolves.toEqual({
      available: false,
      explanation: 'אינו זמין בגרסה זו',
    });
  });

  it('אין Document API — „המסמך עדיין נטען”, ולא זריקה', async () => {
    for (const host of [null, undefined, { activeEditor: null }, { activeEditor: { doc: null } }]) {
      await expect(readVertAlignSupport(host)).resolves.toEqual({
        available: false,
        explanation: 'המסמך עדיין נטען',
      });
    }
  });

  it('אין capabilities לשאול, או קריאה שזורקת — נכשל סגור', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});

    await expect(readVertAlignSupport(fakeDoc({ capabilities: {} }).host)).resolves.toMatchObject({
      available: false,
    });
    await expect(
      readVertAlignSupport(
        fakeDoc({
          capabilities: {
            get: () => {
              throw new Error('boom');
            },
          },
        }).host,
      ),
    ).resolves.toMatchObject({ available: false });

    warn.mockRestore();
  });
});

describe('toggleVertAlign', () => {
  it('שולחת את הכתב המבוקש על ה-selectionTarget של הבחירה', async () => {
    const { host, calls } = fakeDoc();

    await expect(toggleVertAlign(host, 'superscript')).resolves.toEqual({ ok: true });
    expect(calls).toEqual([{ target: RANGE, value: 'superscript' }]);
  });

  it('כתב תחתי הוא אותו מסלול עם ערך אחר', async () => {
    const { host, calls } = fakeDoc();

    await expect(toggleVertAlign(host, 'subscript')).resolves.toEqual({ ok: true });
    expect(calls).toEqual([{ target: RANGE, value: 'subscript' }]);
  });

  it('NO_OP = „זה כבר המצב”, ולכן נשלח baseline — וזה ה-toggle', async () => {
    // אין state מקומי: המנוע הוא שאומר שהבחירה כבר בכתב הזה.
    const { host, calls } = fakeDoc({ receipts: [NO_OP, { success: true }] });

    await expect(toggleVertAlign(host, 'superscript')).resolves.toEqual({ ok: true });
    expect(calls).toEqual([
      { target: RANGE, value: 'superscript' },
      { target: RANGE, value: 'baseline' },
    ]);
  });

  it('`baseline` ולא `null`: כיבוי שנשען על ירושה מהסגנון לא היה נראה', async () => {
    const { host, calls } = fakeDoc({ receipts: [NO_OP, { success: true }] });

    await toggleVertAlign(host, 'subscript');

    expect(calls[1]).toEqual({ target: RANGE, value: 'baseline' });
  });

  it('NO_OP בשני הכיוונים אינו שגיאה — לא היה מה לשנות', async () => {
    const { host, calls } = fakeDoc({ receipts: [NO_OP] });

    await expect(toggleVertAlign(host, 'superscript')).resolves.toEqual({ ok: true });
    expect(calls).toHaveLength(2);
  });

  it('הצלחה בהחלה אינה שולחת קריאה שנייה', async () => {
    const { host, calls } = fakeDoc();

    await toggleVertAlign(host, 'superscript');

    expect(calls).toHaveLength(1);
  });

  it('בלי טקסט מסומן — הודעה מדויקת, ובלי קריאה למנוע', async () => {
    // ה-Document API עובד על טווח ב-XML ולא על „מה שיוקלד הלאה”, ולכן סמן
    // בלי בחירה הוא בקשה בלי יעד.
    for (const selection of [
      { empty: true, selectionTarget: RANGE },
      { empty: false, selectionTarget: null },
      {},
      null,
    ]) {
      const { host, calls } = fakeDoc({ selection });

      await expect(toggleVertAlign(host, 'superscript')).resolves.toEqual({
        ok: false,
        message: 'יש לסמן טקסט תחילה',
        reason: 'range-selection-required',
      });
      expect(calls).toEqual([]);
    }
  });

  it('קריאת בחירה שזורקת אינה מפילה את הרצועה', async () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const doc = {
      selection: {
        current: () => {
          throw new Error('boom');
        },
      },
      format: { vertAlign: () => ({ success: true }) },
    } as unknown as VertAlignDocumentApi;

    await expect(toggleVertAlign({ activeEditor: { doc } }, 'superscript')).resolves.toMatchObject({
      ok: false,
      reason: 'range-selection-required',
    });
    expect(warn).toHaveBeenCalled();
    warn.mockRestore();
  });

  it('קבלה שנכשלה מתורגמת לעברית עם הקוד', async () => {
    const { host } = fakeDoc({
      receipts: [{ success: false, failure: { code: 'DOCUMENT_READONLY', message: 'readonly' } }],
    });

    const outcome = await toggleVertAlign(host, 'superscript');

    expect(outcome).toMatchObject({ ok: false, reason: 'DOCUMENT_READONLY' });
    if (!outcome.ok) expect(outcome.message).toContain('לקריאה בלבד');
  });

  it('כשל בכיבוי מדווח כשל של כיבוי, ולא של החלה', async () => {
    const { host } = fakeDoc({
      receipts: [NO_OP, { success: false, failure: { code: 'DOCUMENT_READONLY' } }],
    });

    const outcome = await toggleVertAlign(host, 'superscript');

    expect(outcome).toMatchObject({ ok: false, reason: 'DOCUMENT_READONLY' });
    if (!outcome.ok) expect(outcome.message).toContain('הכיבוי');
  });

  it('קוד שאין לו תרגום מוצג עם ההסבר של המנוע — ולא נעלם', async () => {
    const { host } = fakeDoc({
      receipts: [{ success: false, failure: { code: 'WEIRD_CODE', message: 'something odd' } }],
    });

    const outcome = await toggleVertAlign(host, 'superscript');

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.message).toContain('something odd');
      expect(outcome.message).toContain('WEIRD_CODE');
    }
  });

  it('זריקה הופכת להודעה ולא מפילה את הרצועה', async () => {
    const { host } = fakeDoc({
      receipts: [
        () => {
          throw new Error('INVALID_INPUT: target is required');
        },
      ],
    });

    const outcome = await toggleVertAlign(host, 'superscript');

    expect(outcome).toMatchObject({ ok: false, reason: 'threw' });
    if (!outcome.ok) expect(outcome.message).toContain('INVALID_INPUT');
  });

  it('זריקה בכיבוי מטופלת גם היא', async () => {
    const { host } = fakeDoc({
      receipts: [
        NO_OP,
        () => {
          throw new Error('boom');
        },
      ],
    });

    await expect(toggleVertAlign(host, 'superscript')).resolves.toMatchObject({
      ok: false,
      reason: 'threw',
    });
  });

  it('גרסה שאינה חושפת את הפעולה מדווחת „אינו זמין בגרסה זו”', async () => {
    const { host, calls } = fakeDoc({ omitFormat: true });

    const outcome = await toggleVertAlign(host, 'superscript');

    expect(outcome).toMatchObject({ ok: false, reason: 'command-unsupported' });
    if (!outcome.ok) expect(outcome.message).toContain('אינו זמין בגרסה זו');
    expect(calls).toEqual([]);
  });

  it('אין Document API — תוצאה מטופסת, לא זריקה', async () => {
    await expect(toggleVertAlign(null, 'superscript')).resolves.toMatchObject({ ok: false });
    await expect(
      toggleVertAlign({ activeEditor: { doc: null } }, 'subscript'),
    ).resolves.toMatchObject({ ok: false });
  });

  it('גרסה בלי `selection.current` אינה נופלת', async () => {
    const { host } = fakeDoc({ omitSelection: true });

    await expect(toggleVertAlign(host, 'superscript')).resolves.toMatchObject({
      ok: false,
      reason: 'not-ready',
    });
  });

  it('Promise נתמך בדיוק כמו קבלה סינכרונית', async () => {
    const { host } = fakeDoc({
      receipts: [Promise.resolve({ success: true }) as unknown as Receipt],
    });

    await expect(toggleVertAlign(host, 'superscript')).resolves.toEqual({ ok: true });
  });
});
