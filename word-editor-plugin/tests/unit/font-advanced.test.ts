/**
 * „גופן מתקדם". הבדיקה על **מה נשלח למנוע**, על השערים שהמנוע אינו מפעיל,
 * ועל מסלולי הכשל; שההחלה עצמה עובדת נבדק במדידת הדפדפן (ראו הערת הפתיחה
 * במודול).
 *
 * השערים שנבדקים כאן הם בדיוק אלה שהמנוע בולע ב-`success:true` (נמדד):
 * charScale 9999 נכתב `w:w="9999"` למרות שתחום Word הוא 1..600, ו-kerning
 * שלילי נכתב `w:kern` שלילי למרות ש-ST_HpsMeasure אינו חתום.
 */
import { describe, expect, it, vi } from 'vitest';
import {
  applyFontAdvanced,
  buildInlinePatch,
  CHAR_SCALE_MAX,
  hasRangeSelection,
  type FontAdvancedPatch,
} from '../../src/engine/font-advanced';

/** מה ש-`selection.current` מחזיר בפועל: ה-selectionTarget ולא TextTarget. */
const SELECTION_INFO = {
  empty: false,
  selectionTarget: { kind: 'selection', start: { kind: 'text', blockId: 'b1', offset: 0 }, end: { kind: 'text', blockId: 'b1', offset: 5 } },
};

function fakeDoc(options: { apply?: (input: unknown) => unknown; selection?: unknown } = {}) {
  const calls: unknown[] = [];
  const impl = options.apply;
  const doc = {
    selection: { current: vi.fn(async () => options.selection ?? SELECTION_INFO) },
    ...(impl === undefined
      ? {}
      : {
          format: {
            apply: (input: unknown) => {
              calls.push(input);
              return impl(input) as never;
            },
          },
        }),
  } as never;
  return { doc, calls, host: { activeEditor: { doc } } };
}

describe('buildInlinePatch', () => {
  it('patch מלא נבנה עם המפתחות של החוזה', () => {
    const patch: FontAdvancedPatch = {
      charScale: 150,
      letterSpacingPt: -20,
      positionPt: 10,
      kerningPt: 12,
      dstrike: true,
      vanish: false,
      boldCs: true,
      complexScript: true,
      rtl: true,
      fontSizeCsPt: 12.5,
      complexFontName: ' David ',
      proofingLangBidi: 'he-IL',
    };

    const result = buildInlinePatch(patch);

    expect('error' in result).toBe(false);
    if (!('error' in result)) {
      expect(result.inline).toEqual({
        charScale: 150,
        letterSpacing: -20,
        position: 10,
        kerning: 12,
        dstrike: true,
        vanish: false,
        bCs: true,
        cs: true,
        rtl: true,
        fontSizeCs: 12.5,
        rFonts: { cs: 'David' },
        lang: { bidi: 'he-IL' },
      });
    }
  });

  it(`charScale מעל ${CHAR_SCALE_MAX} נעצר — המנוע קיבל 9999 בחיוב וכתב w:w פסול`, () => {
    const result = buildInlinePatch({ charScale: 9999 });

    expect('error' in result).toBe(true);
    if ('error' in result) expect(result.error).toMatchObject({ ok: false, reason: 'invalid-char-scale' });
  });

  it('kerning שלילי נעצר — המנוע קיבלו בחיוב ו-ST_HpsMeasure אינו חתום', () => {
    const result = buildInlinePatch({ kerningPt: -5 });

    expect('error' in result).toBe(true);
    if ('error' in result) expect(result.error).toMatchObject({ ok: false, reason: 'invalid-kerning' });
  });

  it('letterSpacing שלילי דווקא חוקי — מכווץ (נמדד: -20 → w:spacing=-400)', () => {
    const result = buildInlinePatch({ letterSpacingPt: -20 });

    expect('error' in result).toBe(false);
    if (!('error' in result)) expect(result.inline.letterSpacing).toBe(-20);
  });

  it('fontSizeCs מקבל חצאי נקודות (12.5 → szCs 25) ודוחה שאר שברים', () => {
    expect(buildInlinePatch({ fontSizeCsPt: 12.5 })).not.toHaveProperty('error');
    const result = buildInlinePatch({ fontSizeCsPt: 12.3 });
    expect('error' in result).toBe(true);
  });

  it('קוד שפה פסול נעצר', () => {
    const result = buildInlinePatch({ proofingLangBidi: 'not a lang!' });

    expect('error' in result).toBe(true);
    if ('error' in result) expect(result.error).toMatchObject({ ok: false, reason: 'invalid-lang' });
  });

  it('patch ריק מחזיר inline ריק — הקריאה תדלג על המנוע', () => {
    const result = buildInlinePatch({});

    expect('error' in result).toBe(false);
    if (!('error' in result)) expect(result.inline).toEqual({});
  });
});

describe('applyFontAdvanced', () => {
  it('שולחת { target: selectionTarget, inline } ל-format.apply', async () => {
    const { host, calls } = fakeDoc({ apply: () => ({ success: true }) });

    const outcome = await applyFontAdvanced(host, { charScale: 150, dstrike: true });

    expect(outcome).toEqual({ ok: true });
    expect(calls[0]).toMatchObject({
      target: SELECTION_INFO.selectionTarget,
      inline: { charScale: 150, dstrike: true },
    });
  });

  it('NO_OP הוא הצלחה (נמדד: vanish null על טקסט גלוי)', async () => {
    const { host } = fakeDoc({
      apply: () => ({ success: false, failure: { code: 'NO_OP', message: 'produced no change.' } }),
    });

    await expect(applyFontAdvanced(host, { vanish: false })).resolves.toEqual({ ok: true });
  });

  it('קבלה שנכשלה מתורגמת לעברית עם הקוד', async () => {
    const { host } = fakeDoc({
      apply: () => ({ success: false, failure: { code: 'DOCUMENT_READONLY', message: 'readonly' } }),
    });

    const outcome = await applyFontAdvanced(host, { outline: true });

    expect(outcome.ok).toBe(false);
    if (!outcome.ok) {
      expect(outcome.message).toContain('החלת עיצוב הגופן נכשלה');
      expect(outcome.reason).toBe('DOCUMENT_READONLY');
    }
  });

  it('זריקה של הוולידטור הופכת להודעה ולא מפילה את הרצועה', async () => {
    const { host } = fakeDoc({
      apply: () => {
        throw new Error('INVALID_INPUT: bad patch');
      },
    });

    const outcome = await applyFontAdvanced(host, { shadow: true });

    expect(outcome).toMatchObject({ ok: false, reason: 'threw' });
    if (!outcome.ok) expect(outcome.message).toContain('INVALID_INPUT');
  });

  it('בלי טווח מסומן — „יש לסמן טקסט תחילה" ולא קריאה למנוע', async () => {
    const { host, calls } = fakeDoc({
      apply: () => ({ success: true }),
      selection: { empty: true },
    });

    const outcome = await applyFontAdvanced(host, { emboss: true });

    expect(outcome).toMatchObject({ ok: false, reason: 'range-selection-required' });
    expect(calls).toHaveLength(0);
  });

  it('patch ריק מסתיים בהצלחה בלי לגעת במנוע', async () => {
    const { host, calls } = fakeDoc({ apply: () => ({ success: true }) });

    await expect(applyFontAdvanced(host, {})).resolves.toEqual({ ok: true });
    expect(calls).toHaveLength(0);
  });

  it('אין Document API — תוצאה מטופלת עם הנוסח של §12', async () => {
    for (const host of [null, undefined, { activeEditor: null }] as never[]) {
      const outcome = await applyFontAdvanced(host, { imprint: true });
      expect(outcome).toMatchObject({ ok: false, reason: 'command-unsupported' });
      if (!outcome.ok) expect(outcome.message).toContain('אינו זמין בגרסה זו');
    }
  });

  it('ערך פסול נעצר בשער שלנו ולא מגיע למנוע בכלל', async () => {
    const { host, calls } = fakeDoc({ apply: () => ({ success: true }) });

    const outcome = await applyFontAdvanced(host, { charScale: 9999 });

    expect(outcome.ok).toBe(false);
    expect(calls).toHaveLength(0);
  });
});

/**
 * `hasRangeSelection` היא אותה קריאה של `applyFontAdvanced`, מוקדמת.
 *
 * מה שהיא **אינה** עושה חשוב לא פחות ממה שהיא כן: היא מבדילה בין „אין בחירה”
 * לבין „אין לדעת”. הדיאלוג נועל את „אישור” על הראשון בלבד — נעילה על מסמך
 * שעדיין נטען הייתה חוסמת פעולה תקינה בלי שום דרך לפתוח אותה.
 */
describe('hasRangeSelection', () => {
  it('טווח מסומן → range', async () => {
    const { host } = fakeDoc({ apply: () => ({ success: true }) });

    await expect(hasRangeSelection(host)).resolves.toBe('range');
  });

  it('סמן מכווץ → none, וזה מה שנועל את „אישור”', async () => {
    const { host } = fakeDoc({ apply: () => ({ success: true }), selection: { empty: true } });

    await expect(hasRangeSelection(host)).resolves.toBe('none');
  });

  it('בחירה בלי selectionTarget → none', async () => {
    const { host } = fakeDoc({ apply: () => ({ success: true }), selection: { empty: false } });

    await expect(hasRangeSelection(host)).resolves.toBe('none');
  });

  /**
   * „אין מסמך” אינו „אין יכולת”, וזו אינה קפדנות לשונית: בין מסמכים, ובזמן
   * טעינה, המסמך הפעיל הוא `null` — והודעה שאומרת שם „אינו זמין בגרסה זו של
   * המנוע” היא טענה שקרית על הבניין, שאין למשתמש שום דרך להפריך.
   */
  it('אין מסמך פעיל → unknown, ואינו נועל דבר', async () => {
    for (const host of [null, undefined, { activeEditor: null }] as never[]) {
      await expect(hasRangeSelection(host)).resolves.toBe('unknown');
    }
  });

  it('יש מסמך ואין format.apply → unavailable: זה הפער שההודעה מתארת', async () => {
    const { host } = fakeDoc();
    await expect(hasRangeSelection(host)).resolves.toBe('unavailable');
  });

  it('אין selection.current → unknown: „לא מוכן” אינו „אין בחירה”, ואינו נועל', async () => {
    const doc = { format: { apply: () => ({ success: true }) } } as never;
    await expect(hasRangeSelection({ activeEditor: { doc } } as never)).resolves.toBe('unknown');
  });

  /**
   * קריאה שזרקה נספרת כ„אין בחירה” ולא כ„אין לדעת”, וזה מכוון: `applyFontAdvanced`
   * עובר דרך אותה `readSelectionTarget` ויענה בדיוק אותו דבר. דיאלוג שהיה
   * נפתח כאן והופך את הכפתור לזמין היה מבטיח החלה שתיכשל מיד.
   */
  it('קריאת הבחירה זרקה → נספר כאין בחירה, ולעולם אינה זורקת הלאה', async () => {
    const doc = {
      selection: { current: () => { throw new Error('boom'); } },
      format: { apply: () => ({ success: true }) },
    } as never;

    await expect(hasRangeSelection({ activeEditor: { doc } } as never)).resolves.toBe('none');
  });
});
