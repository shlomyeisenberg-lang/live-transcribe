/**
 * ערכת העיצוב: מה נשמר, מה מוצג, ומה מורץ.
 *
 * המבחן שחוזר בכל הקבוצות כאן הוא ההבחנה בין „אל תיגע” (שדה חסר) ל„נקה”
 * (`null`). שתיהן היו נופלות לאותו `undefined` בכל ייצוג נאיבי, וההתמוטטות
 * הזאת היא בדיוק ערכה שמגדילה גופן ומוחקת בדרך את הצבע.
 */
import { describe, expect, it } from 'vitest';
import {
  fieldText,
  isEmptyPreset,
  normalizePreset,
  presetFields,
  presetSteps,
  presetSummary,
  readingMatchesPreset,
  reverseOf,
  unreadableFields,
  type FormatPreset,
  type FormatReading,
} from '../../src/ui/shortcuts/format-preset';

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

describe('נורמליזציה של ערכה', () => {
  it('שדות תקינים עוברים, ומנורמלים כמו שהמנוע דורש', () => {
    expect(normalizePreset({ fontFamily: '  David  ', fontSizePt: '14', color: 'ff0000' })).toEqual({
      fontFamily: 'David',
      fontSizePt: 14,
      color: '#FF0000',
    });
  });

  it('שדה שאינו תקין נושר — הוא לא היה מגיע למנוע בכל מקרה', () => {
    // `fontSizePt: 0` ו-`color: 'אדום'` נדחים סגור בוולידטור של המנוע, ושם
    // המשתמש היה רואה קיצור ששותק בלי הסבר.
    expect(normalizePreset({ fontSizePt: 0, color: 'אדום', fontFamily: '   ' })).toEqual({});
  });

  it('`null` מפורש שורד כ„נקה”, והיעדר השדה שורד כ„אל תיגע”', () => {
    expect(normalizePreset({ color: null })).toEqual({ color: null });
    expect(normalizePreset({})).toEqual({});
    expect('color' in normalizePreset({})).toBe(false);
  });

  it('מתג שאינו בוליאני נושר, ו-`false` נשמר', () => {
    expect(normalizePreset({ bold: 'yes', italic: false })).toEqual({ italic: false });
  });

  it('ערך שאינו אובייקט הוא ערכה ריקה, ולא קריסה', () => {
    for (const raw of [null, undefined, 7, 'x', []]) {
      expect(isEmptyPreset(normalizePreset(raw))).toBe(true);
    }
  });
});

describe('תצוגה', () => {
  it('כל שדה מקבל נוסח שאומר מה יקרה', () => {
    const preset: FormatPreset = {
      fontFamily: 'David',
      fontSizePt: 14,
      color: '#FF0000',
      highlight: null,
      bold: true,
      italic: false,
    };
    expect(fieldText(preset, 'fontFamily')).toBe('David');
    expect(fieldText(preset, 'fontSizePt')).toBe("14 נק'");
    expect(fieldText(preset, 'color')).toBe('צבע #FF0000');
    expect(fieldText(preset, 'highlight')).toBe('ללא הדגשה');
    expect(fieldText(preset, 'bold')).toBe('מודגש');
    // מתג שכובה אינו „מודגש”: ערכה יכולה לבקש את שני הכיוונים.
    expect(fieldText(preset, 'italic')).toBe('ללא נטוי');
  });

  it('הסיכום מונה רק את השדות שנקבעו, בסדר קבוע', () => {
    expect(presetSummary({ bold: true, fontFamily: 'David' })).toBe('David, מודגש');
    expect(presetSummary({})).toBe('');
  });

  it('`presetFields` אינו כולל שדה שלא נקבע', () => {
    expect(presetFields({ color: null, bold: false })).toEqual(['color', 'bold']);
  });
});

describe('הקריאה של המנוע והמתג', () => {
  const preset: FormatPreset = { fontFamily: 'David', fontSizePt: 14, color: '#FF0000' };

  it('„הקודם” מוגבל לשדות שהערכה נוגעת בהם', () => {
    // הערכה אינה נוגעת בהדגשה, ולכן היא גם לא תשחזר אותה — גם אם המשתמש
    // שינה אותה בינתיים בעצמו.
    const reading: FormatReading = { ...PLAIN, highlight: '#FFFF00' };
    expect(reverseOf(reading, preset)).toEqual({
      fontFamily: 'Arial',
      fontSizePt: 10,
      color: null,
    });
  });

  it('שדה שהמנוע לא ידע לדווח אינו נכנס ל„קודם”, ומדווח כלא-משוחזר', () => {
    const mixed: FormatReading = { ...PLAIN, fontFamily: undefined, color: undefined };
    expect(reverseOf(mixed, preset)).toEqual({ fontSizePt: 10 });
    expect(unreadableFields(mixed, preset)).toEqual(['fontFamily', 'color']);
  });

  it('התאמה חסינה לרישיות ולרווחים בשם הגופן', () => {
    expect(readingMatchesPreset({ ...PLAIN, fontFamily: ' david ', fontSizePt: 14, color: '#FF0000' }, preset)).toBe(
      true,
    );
  });

  it('שדה שאינו נודע נחשב כלא מתאים — אין להתייחס לערכה כפעילה בלי לאמת', () => {
    const unknown: FormatReading = { ...PLAIN, fontFamily: 'David', fontSizePt: 14, color: undefined };
    expect(readingMatchesPreset(unknown, preset)).toBe(false);
  });

  it('„ללא צבע” בערכה מתאים ל„ללא צבע” בקריאה, ולא ל„לא נודע”', () => {
    expect(readingMatchesPreset({ ...PLAIN, color: null }, { color: null })).toBe(true);
    expect(readingMatchesPreset({ ...PLAIN, color: undefined }, { color: null })).toBe(false);
  });

  it('מתג נכבה מתאים לערכה שמבקשת לכבות אותו', () => {
    expect(readingMatchesPreset(PLAIN, { bold: false })).toBe(true);
    expect(readingMatchesPreset(PLAIN, { bold: true })).toBe(false);
  });
});

describe('הצעדים שמורצים', () => {
  it('ערך ומתג מתורגמים לשני סוגי צעדים שונים', () => {
    // ההפרדה אינה קוסמטית: `bold` היא „הפוך הדגשה”, ולכן היא אינה יכולה
    // להיות צעד עם payload.
    expect(presetSteps({ fontSizePt: 14, bold: true })).toEqual([
      { kind: 'value', command: 'font-size', payload: 14 },
      { kind: 'toggle', command: 'bold', to: true },
    ]);
  });

  it('צבע נשלח בצורה שהמנוע מחלץ, ו„ניקוי” הוא `null` מפורש', () => {
    expect(presetSteps({ color: '#FF0000' })).toEqual([
      { kind: 'value', command: 'text-color', payload: { value: '#FF0000' } },
    ]);
    expect(presetSteps({ highlight: null })).toEqual([
      { kind: 'value', command: 'highlight-color', payload: { value: null } },
    ]);
  });

  it('ערכה ריקה אינה מייצרת צעדים', () => {
    expect(presetSteps({})).toEqual([]);
  });
});
