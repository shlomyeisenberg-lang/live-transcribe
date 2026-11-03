/**
 * המתג — ההכרעה שכל הפיצ'ר תלוי בה.
 *
 * הבדיקה המרכזית כאן היא **לא** „לחיצה מחילה, לחיצה מחזירה”. זה החלק הקל.
 * מה שנבדק הוא מה קורה כשהמשתמש עשה משהו בין הלחיצות: זז לטקסט אחר, שינה
 * גופן ביד, או סימן טווח מעורב. דגל בוליאני היה מחזיר במצבים האלה „עיצוב
 * קודם” שאינו קודם לשום דבר — כלומר משכתב טקסט שלא נגעו בו.
 */
import { describe, expect, it } from 'vitest';
import { createPresetToggles, decideToggle } from '../../src/ui/shortcuts/preset-toggle';
import type { FormatPreset, FormatReading } from '../../src/ui/shortcuts/format-preset';

const PRESET: FormatPreset = { fontFamily: 'David', fontSizePt: 14 };

const PLAIN: FormatReading = {
  fontFamily: 'Arial',
  fontSizePt: 10,
  color: null,
  highlight: null,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
};

/** הקריאה שמתארת טקסט שהערכה כבר הוחלה עליו. */
const STYLED: FormatReading = { ...PLAIN, fontFamily: 'David', fontSizePt: 14 };

function input(over: Partial<Parameters<typeof decideToggle>[0]> = {}) {
  return {
    name: 'כותרת קטע',
    label: 'Ctrl+Alt+K',
    preset: PRESET,
    reading: PLAIN,
    previous: null,
    ...over,
  };
}

describe('לחיצה ראשונה', () => {
  it('מחילה את הערכה ולוכדת את מה שהיה', () => {
    const decision = decideToggle(input());
    expect(decision.restoring).toBe(false);
    expect(decision.apply).toEqual(PRESET);
    expect(decision.remember).toEqual({ fontFamily: 'Arial', fontSizePt: 10 });
  });

  it('ההודעה נוקבת בשם ובצירוף — היא זו שמלמדת שיש חזרה', () => {
    expect(decideToggle(input()).status).toBe(
      '„כותרת קטע” הוחל — Ctrl+Alt+K להחזרת העיצוב הקודם',
    );
  });

  it('שדה שלא נקרא מדווח בהודעה, ולא נעלם בשקט', () => {
    const decision = decideToggle(
      input({ reading: { ...PLAIN, fontFamily: undefined, fontSizePt: undefined } }),
    );
    expect(decision.remember).toEqual({});
    expect(decision.status).toContain('גופן וגודל לא נקראו מהבחירה');
  });
});

describe('לחיצה שנייה', () => {
  it('מחזירה את מה שנלכד, כשהטקסט אכן נראה כמו הערכה', () => {
    const decision = decideToggle(
      input({ reading: STYLED, previous: { fontFamily: 'Arial', fontSizePt: 10 } }),
    );
    expect(decision.restoring).toBe(true);
    expect(decision.apply).toEqual({ fontFamily: 'Arial', fontSizePt: 10 });
    // הזיכרון נמחק: לחיצה שלישית מחילה מחדש.
    expect(decision.remember).toBeNull();
    expect(decision.status).toContain('העיצוב הקודם הוחזר');
  });

  it('**מחילה מחדש** כשהמשתמש זז לטקסט שאינו נראה כמו הערכה', () => {
    // זה הלב. עם דגל בוליאני הלחיצה הזאת הייתה מחילה על הפסקה החדשה את
    // „Arial 10” של הפסקה הקודמת — עיצוב שלטקסט הזה לא היה מעולם.
    const decision = decideToggle(
      input({ reading: PLAIN, previous: { fontFamily: 'Arial', fontSizePt: 10 } }),
    );
    expect(decision.restoring).toBe(false);
    expect(decision.apply).toEqual(PRESET);
    expect(decision.remember).toEqual({ fontFamily: 'Arial', fontSizePt: 10 });
  });

  it('מחילה מחדש גם כשרק חלק מהערכה עוד בתוקף', () => {
    // המשתמש הגדיל ל-16 ביד: הגופן עוד „דוד” אבל הגודל אינו 14, ולכן הערכה
    // אינה בתוקף — והלחיצה מחזירה אותה במלואה.
    const partly: FormatReading = { ...PLAIN, fontFamily: 'David', fontSizePt: 16 };
    const decision = decideToggle(input({ reading: partly, previous: { fontSizePt: 10 } }));
    expect(decision.apply).toEqual(PRESET);
  });

  it('בחירה מעורבת אינה נחשבת „דלוקה”', () => {
    const mixed: FormatReading = { ...PLAIN, fontFamily: undefined, fontSizePt: undefined };
    expect(decideToggle(input({ reading: mixed, previous: { fontSizePt: 10 } })).restoring).toBe(
      false,
    );
  });

  it('„קודם” ריק אינו מחיל דבר, ומשחרר את הזיכרון', () => {
    // הלחיצה הראשונה נעשתה על בחירה שלא נקרא ממנה כלום. אין מה להחזיר, ואין
    // להיתקע: הלחיצה הבאה חייבת להחיל מחדש.
    const decision = decideToggle(input({ reading: STYLED, previous: {} }));
    expect(decision.apply).toEqual({});
    expect(decision.remember).toBeNull();
    expect(decision.status).toContain('לא נשמר עיצוב קודם');
  });
});

describe('הזיכרון', () => {
  it('שתי לחיצות עוברות הלוך וחזור, והשלישית מחילה שוב', () => {
    const toggles = createPresetToggles();
    const base = { id: 'cs-1', name: 'כותרת', label: 'Ctrl+Alt+K', preset: PRESET };

    const first = toggles.decide({ ...base, reading: PLAIN });
    expect(first.apply).toEqual(PRESET);
    expect(toggles.isOn('cs-1')).toBe(true);

    const second = toggles.decide({ ...base, reading: STYLED });
    expect(second.apply).toEqual({ fontFamily: 'Arial', fontSizePt: 10 });
    expect(toggles.isOn('cs-1')).toBe(false);

    const third = toggles.decide({ ...base, reading: PLAIN });
    expect(third.apply).toEqual(PRESET);
  });

  it('`forget` מאפס קיצור אחד, ו-`forgetAll` את כולם', () => {
    const toggles = createPresetToggles();
    const base = { name: 'כותרת', label: 'Ctrl+Alt+K', preset: PRESET, reading: PLAIN };

    toggles.decide({ ...base, id: 'cs-1' });
    toggles.decide({ ...base, id: 'cs-2' });
    toggles.forget('cs-1');
    expect(toggles.isOn('cs-1')).toBe(false);
    expect(toggles.isOn('cs-2')).toBe(true);

    toggles.forgetAll();
    expect(toggles.isOn('cs-2')).toBe(false);
  });

  it('אחרי שכחה, לחיצה על טקסט שנראה כמו הערכה מחילה ולא מחזירה', () => {
    // התרחיש של החלפת טאב: הזיכרון של טאב א' נשכח, ולכן לחיצה בטאב ב' —
    // שהטקסט בו ממילא נראה כמו הערכה — אינה מחילה עליו את העיצוב של טאב א'.
    const toggles = createPresetToggles();
    const base = { id: 'cs-1', name: 'כותרת', label: 'Ctrl+Alt+K', preset: PRESET };

    toggles.decide({ ...base, reading: PLAIN });
    toggles.forgetAll();

    const afterSwitch = toggles.decide({ ...base, reading: STYLED });
    expect(afterSwitch.restoring).toBe(false);
    expect(afterSwitch.apply).toEqual(PRESET);
  });
});
