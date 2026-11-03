/**
 * החזקת החיווי ברצועה.
 *
 * מה שנבדק כאן הוא הכלל שמפריד „מעורב” מ„עוד לא נפתר”, ולא הנוחות שהוא נותן:
 * המנוע מקפל את שניהם לאותו `undefined` (ראו את ההנמקה ב-readout-hold.ts),
 * והמחיר של טעות לכל כיוון שונה לגמרי — החזקה במקום הלא נכון משקרת למשתמש על
 * מה שיש במסמך, ואי-החזקה במקום הנכון היא ההבהוב שנמדד בכפתור „יישור לימין”.
 */
import { describe, expect, it, vi } from 'vitest';
import type { CommandState } from 'superdoc/ui';
import {
  UNSETTLED_SELECTION,
  canHoldReadout,
  displayedValue,
  heldCommandState,
  observeReadoutSelection,
  toReadoutSelection,
  type SelectionSliceLike,
} from '../../src/engine/readout-hold';

const CARET_SETTLED = { empty: true, settled: true };
const CARET_UNSETTLED = { empty: true, settled: false };
const RANGE_SETTLED = { empty: false, settled: true };
const RANGE_UNSETTLED = { empty: false, settled: false };

function state(patch: Partial<CommandState> = {}): CommandState {
  return { supported: true, enabled: true, active: false, value: undefined, ...patch };
}

describe('מתי מותר להחזיק', () => {
  it('סמן מכווץ — תמיד', () => {
    // לפסקה אחת יש יישור אחד, ולכן `undefined` על סמן אינו יכול להיות „מעורב”.
    expect(canHoldReadout(CARET_SETTLED)).toBe(true);
    expect(canHoldReadout(CARET_UNSETTLED)).toBe(true);
  });

  it('טווח שהקריאה שלו לא התיישבה — כן', () => {
    expect(canHoldReadout(RANGE_UNSETTLED)).toBe(true);
  });

  it('טווח שהתיישב — לא', () => {
    // כאן `undefined` הוא התשובה האמיתית: „מעורב”. Word מציג בדיוק אותו דבר.
    expect(canHoldReadout(RANGE_SETTLED)).toBe(false);
  });
});

describe('הערך המוצג', () => {
  it('ערך טרי מנצח תמיד, גם באמצע החזקה', () => {
    // זה מה שמייתר מנגנון ביטול-החזקה: לחיצה על „מרכז” נדלקת מיד.
    expect(displayedValue('center', 'right', CARET_UNSETTLED)).toBe('center');
    expect(displayedValue('center', 'right', RANGE_SETTLED)).toBe('center');
  });

  it('אין ערך ומותר להחזיק — האחרון שידענו', () => {
    expect(displayedValue(undefined, 'right', CARET_SETTLED)).toBe('right');
    expect(displayedValue(undefined, 'right', RANGE_UNSETTLED)).toBe('right');
  });

  it('אין ערך ואסור להחזיק — ריק', () => {
    expect(displayedValue(undefined, 'right', RANGE_SETTLED)).toBeUndefined();
  });

  it('אין מה להחזיק — ריק, ולא „undefined מוחזק”', () => {
    expect(displayedValue(undefined, undefined, CARET_SETTLED)).toBeUndefined();
  });

  it('`null` הוא ערך ולא היעדר ערך', () => {
    // „ללא צבע” הוא תשובה של המנוע, ולא „עוד לא יודעים”.
    expect(displayedValue(null, '#ff0000', CARET_SETTLED)).toBeNull();
  });
});

describe('המצב שהפקד מקבל', () => {
  it('זמינות וחיווי עוברים טריים — רק הערך מוחזק', () => {
    const incoming = state({ enabled: false, active: true, reason: 'selection-required' });
    const held = heldCommandState(incoming, 'right', CARET_SETTLED);

    expect(held.value).toBe('right');
    // זמינות אינה חיווי אלא הבטחה, ו-`active` בוליאני שאין בו „לא ידוע”.
    expect(held.enabled).toBe(false);
    expect(held.active).toBe(true);
    expect(held.reason).toBe('selection-required');
  });

  it('אותו אובייקט כשאין מה לשנות', () => {
    // `useCommand` מזין מכאן `computed`; אובייקט חדש בכל דיווח היה מרנדר
    // מחדש כל פקד ברצועה על כל תו שנקלד.
    const incoming = state({ value: 'right' });
    expect(heldCommandState(incoming, 'left', CARET_SETTLED)).toBe(incoming);

    const blank = state();
    expect(heldCommandState(blank, undefined, CARET_SETTLED)).toBe(blank);
    expect(heldCommandState(blank, 'right', RANGE_SETTLED)).toBe(blank);
  });
});

describe('קריאת ה-slice של המנוע', () => {
  it('`status: ready` הוא היחיד שנחשב מיושב', () => {
    expect(toReadoutSelection({ status: 'ready', empty: false }).settled).toBe(true);
    expect(toReadoutSelection({ status: 'stale', empty: false }).settled).toBe(false);
    expect(toReadoutSelection({ status: 'pending', empty: false }).settled).toBe(false);
  });

  it('slice שאיננו מבינים נקרא כסמן שלא התיישב', () => {
    // ברירת המחדל חייבת להיות זו שמחזיקה, לא זו שמרוקנת.
    expect(toReadoutSelection(undefined)).toEqual(UNSETTLED_SELECTION);
    expect(toReadoutSelection({})).toEqual(UNSETTLED_SELECTION);
  });

  it('`empty: false` מפורש הוא היחיד שנקרא כטווח', () => {
    expect(toReadoutSelection({ status: 'ready', empty: false }).empty).toBe(false);
    expect(toReadoutSelection({ status: 'ready' }).empty).toBe(true);
  });
});

describe('ההאזנה', () => {
  it('מעבירה כל דיווח של המנוע בצורה הצרה', () => {
    const seen: unknown[] = [];
    const captured: { emit: ((slice: SelectionSliceLike) => void) | null } = { emit: null };
    const off = vi.fn();

    const dispose = observeReadoutSelection(
      {
        selection: {
          observe: (listener) => {
            captured.emit = listener;
            return off;
          },
        },
      },
      (selection) => seen.push(selection),
    );

    captured.emit?.({ status: 'ready', empty: false });
    captured.emit?.({ status: 'stale', empty: true });

    expect(seen).toEqual([
      { empty: false, settled: true },
      { empty: true, settled: false },
    ]);

    dispose();
    expect(off).toHaveBeenCalled();
  });

  it('גרסה בלי `observe` נופלת ל-snapshot ומחזירה disposer', () => {
    const seen: unknown[] = [];
    const dispose = observeReadoutSelection(
      { selection: { getSnapshot: () => ({ status: 'ready', empty: false }) } },
      (selection) => seen.push(selection),
    );

    expect(seen).toEqual([{ empty: false, settled: true }]);
    expect(() => dispose()).not.toThrow();
  });

  it('`observe` שזורק אינו מפיל את הרצועה', () => {
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const seen: unknown[] = [];

    const dispose = observeReadoutSelection(
      {
        selection: {
          observe: () => {
            throw new Error('boom');
          },
          getSnapshot: () => ({ status: 'ready', empty: true }),
        },
      },
      (selection) => seen.push(selection),
    );

    expect(seen).toEqual([{ empty: true, settled: true }]);
    expect(() => dispose()).not.toThrow();
    warn.mockRestore();
  });

  it('אין `selection` בכלל — דיווח אחד של „לא התיישב”, ולא חריגה', () => {
    const seen: unknown[] = [];
    const dispose = observeReadoutSelection({}, (selection) => seen.push(selection));

    expect(seen).toEqual([UNSETTLED_SELECTION]);
    expect(() => dispose()).not.toThrow();
  });
});
