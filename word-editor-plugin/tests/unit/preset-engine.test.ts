/**
 * שני הצדדים שנוגעים במנוע: קריאת העיצוב הנוכחי, והחלת ערכה.
 *
 * שתי ההכרעות שנבדקות כאן הן אלה שאי אפשר להסיק מהקוד:
 *
 * 1. **„המנוע לא דיווח צבע” אינו דבר אחד.** על סמן מיושב זה „אין צבע ישיר”,
 *    ועל טווח זה „מעורב”. בלי ההבחנה הזאת המקרה הנפוץ ביותר נשבר — טקסט בצבע
 *    ברירת המחדל, קיצור שצובע אותו, ולחיצה שנייה שאינה יודעת לאן לחזור.
 * 2. **מתג שכבר במצב המבוקש אינו נלחץ.** `bold` היא „הפוך הדגשה”, ולכן ערכה
 *    שמבקשת „מודגש” על טקסט מודגש הייתה **מבטלת** את ההדגשה.
 */
import { describe, expect, it, vi } from 'vitest';
import type { CommandState } from 'superdoc/ui';
import { readFormat } from '../../src/engine/format-reading';
import { applyPreset } from '../../src/engine/apply-preset';
import type { CommandOutcome } from '../../src/engine/command-adapter';

const CARET = { empty: true, settled: true };
const RANGE = { empty: false, settled: true };

function state(over: Partial<CommandState> = {}): CommandState {
  return { supported: true, enabled: true, active: false, value: undefined, ...over };
}

/** מתאם מזויף: מפה ממזהה פקודה למצב. פקודה שאינה במפה אינה „מוכרת”. */
function fakeSource(states: Record<string, CommandState>) {
  return {
    has: (id: string) => id in states,
    getState: (id: string) => states[id] ?? state({ supported: false, enabled: false }),
  };
}

describe('קריאת העיצוב', () => {
  it('גופן, גודל וצבע שדווחו נקראים ומנורמלים', () => {
    const reading = readFormat(
      fakeSource({
        'font-family': state({ value: ' David ' }),
        'font-size': state({ value: '14' }),
        'text-color': state({ value: 'ff0000' }),
      }),
      CARET,
    );
    expect(reading.fontFamily).toBe('David');
    expect(reading.fontSizePt).toBe(14);
    expect(reading.color).toBe('#FF0000');
  });

  it('צבע שלא דווח על **סמן מיושב** נקרא כ„אין צבע”', () => {
    const reading = readFormat(fakeSource({ 'text-color': state() }), CARET);
    expect(reading.color).toBeNull();
  });

  it('צבע שלא דווח על **טווח** נשאר „לא נודע”', () => {
    // על טווח מעורב אין צבע אחד לשחזר, וניחוש „אין צבע” היה מנקה את הצבע של
    // החלק הצבוע.
    const reading = readFormat(fakeSource({ 'text-color': state() }), RANGE);
    expect('color' in reading).toBe(false);
  });

  it('בחירה שטרם התיישבה אינה מזמינה הסקה, גם כשהיא סמן', () => {
    const reading = readFormat(fakeSource({ 'text-color': state() }), {
      empty: true,
      settled: false,
    });
    expect('color' in reading).toBe(false);
  });

  it('גופן וגודל שלא דווחו נשארים „לא נודע” בשתי צורות הבחירה', () => {
    for (const selection of [CARET, RANGE]) {
      const reading = readFormat(fakeSource({}), selection);
      expect('fontFamily' in reading).toBe(false);
      expect('fontSizePt' in reading).toBe(false);
    }
  });

  it('המתגים תמיד נודעים, ופקודה שאינה מוכרת נקראת ככבויה', () => {
    const reading = readFormat(fakeSource({ bold: state({ active: true }) }), CARET);
    expect(reading.bold).toBe(true);
    expect(reading.italic).toBe(false);
    expect(reading.underline).toBe(false);
    expect(reading.strikethrough).toBe(false);
  });
});

describe('החלת ערכה', () => {
  const ok: CommandOutcome = { ok: true };

  function target(states: Record<string, CommandState>) {
    const run = vi.fn(async () => ok);
    return { ...fakeSource(states), run, calls: run.mock.calls };
  }

  it('פקודות ערך מורצות עם ה-payload שהמנוע מחלץ', async () => {
    const engine = target({ 'font-size': state(), 'text-color': state() });
    const ran = await applyPreset(engine, { fontSizePt: 14, color: '#FF0000' }, () => {});
    expect(ran).toBe(2);
    expect(engine.calls).toEqual([
      ['font-size', 14],
      ['text-color', { value: '#FF0000' }],
    ]);
  });

  it('**מתג שכבר במצב המבוקש אינו מורץ**', async () => {
    const engine = target({ bold: state({ active: true }) });
    const ran = await applyPreset(engine, { bold: true }, () => {});
    expect(ran).toBe(0);
    expect(engine.calls).toEqual([]);
  });

  it('מתג שאינו במצב המבוקש מורץ בלי payload', async () => {
    const engine = target({ bold: state({ active: false }) });
    await applyPreset(engine, { bold: true }, () => {});
    // בלי payload: המתג אינו מקבל ערך, הוא הופך את מה שיש.
    expect(engine.calls).toEqual([['bold']]);
  });

  it('כיבוי מתג דלוק מורץ, וכיבוי מתג כבוי — לא', async () => {
    const on = target({ italic: state({ active: true }) });
    await applyPreset(on, { italic: false }, () => {});
    expect(on.calls).toEqual([['italic']]);

    const off = target({ italic: state({ active: false }) });
    await applyPreset(off, { italic: false }, () => {});
    expect(off.calls).toEqual([]);
  });

  it('הצעדים רצים בזה אחר זה, ולא במקביל', async () => {
    // `Promise.all` היה שולח את כולן לאותה בחירה בו-זמנית, ומצב הבחירה שכל
    // אחת קוראת היה תלוי בסדר שה-controller הספיק לנתב.
    const order: string[] = [];
    const engine = {
      ...fakeSource({ 'font-family': state(), 'font-size': state() }),
      run: async (id: string): Promise<CommandOutcome> => {
        order.push(`start:${id}`);
        await Promise.resolve();
        order.push(`end:${id}`);
        return ok;
      },
    };
    await applyPreset(engine, { fontFamily: 'David', fontSizePt: 14 }, () => {});
    expect(order).toEqual([
      'start:font-family',
      'end:font-family',
      'start:font-size',
      'end:font-size',
    ]);
  });

  it('פקודה שהמנוע אינו מכיר מדווחת כלא-נתמכת, ואינה עוצרת את השאר', async () => {
    const engine = target({ 'font-size': state() });
    const reported: { id: string; outcome: CommandOutcome }[] = [];
    const ran = await applyPreset(
      engine,
      { highlight: '#FFFF00', fontSizePt: 14 },
      (outcome, id) => reported.push({ id, outcome }),
    );

    expect(ran).toBe(1);
    expect(engine.calls).toEqual([['font-size', 14]]);
    // הגודל רץ לפני ההדגשה (סדר `presetSteps`), ולכן הדיווח על מה שאינו
    // נתמך הוא השני — ובכל מקרה הוא מגיע, וזה העיקר.
    expect(reported.map((entry) => entry.id)).toEqual(['font-size', 'highlight-color']);
    expect(reported[1]).toMatchObject({
      id: 'highlight-color',
      outcome: { ok: false, reason: 'command-unsupported' },
    });
  });

  it('כשל של צעד אחד אינו מבטל את הבאים אחריו', async () => {
    // ערכה שהצבע שלה נדחה עדיין אמורה להחיל את הגופן — בדיוק כמו שלחיצה על
    // שני כפתורים ברצועה אינה נזנחת באמצע.
    const failures: string[] = [];
    const engine = {
      ...fakeSource({ 'text-color': state(), 'font-size': state() }),
      run: async (id: string): Promise<CommandOutcome> =>
        id === 'text-color' ? { ok: false, message: 'נכשל' } : ok,
    };
    const ran = await applyPreset(engine, { color: '#FF0000', fontSizePt: 14 }, (outcome, id) => {
      if (!outcome.ok) failures.push(id);
    });
    expect(ran).toBe(2);
    expect(failures).toEqual(['text-color']);
  });

  it('ערכה ריקה אינה מריצה דבר', async () => {
    const engine = target({ bold: state() });
    expect(await applyPreset(engine, {}, () => {})).toBe(0);
    expect(engine.calls).toEqual([]);
  });
});
