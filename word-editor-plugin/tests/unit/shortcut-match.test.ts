/**
 * ההתאמה בין אירוע מקלדת לרשומה. הבדיקה המרכזית כאן היא פריסת מקלדת עברית:
 * הקוד שקדם לרג'יסטרי השווה `event.key === 's'`, ולכן כל ששת הקיצורים
 * שהתוסף החזיק מתו ברגע שהמשתמש עבר לעברית — והם מתו בשקט, כי כל הבדיקות
 * שהיו נכתבו בפריסה לטינית בלבד.
 */
import { describe, it, expect } from 'vitest';
import { matchShortcut, matchAny, type KeyEventLike } from '../../src/ui/shortcuts/match';
import { SHORTCUTS, findShortcut, type Shortcut } from '../../src/ui/shortcuts/registry';

function key(over: Partial<KeyEventLike> = {}): KeyEventLike {
  return {
    key: '',
    code: '',
    ctrlKey: false,
    metaKey: false,
    shiftKey: false,
    altKey: false,
    ...over,
  };
}

/** הרשומה מהרג'יסטרי, כדי שהבדיקה תרוץ על מה שבאמת רץ. */
function shortcut(id: string): Shortcut {
  const found = findShortcut(id);
  if (!found) throw new Error(`אין רשומה בשם ${id}`);
  return found;
}

describe('matchShortcut — פריסת מקלדת', () => {
  it('Ctrl+S בפריסה לטינית', () => {
    expect(matchShortcut(key({ key: 's', code: 'KeyS', ctrlKey: true }), shortcut('save'))).toBe(
      true,
    );
  });

  it('Ctrl+S בפריסה עברית — key הוא „ד” וההתאמה עדיין תופסת', () => {
    // זה הבאג שהתיקון הזה נועד לו: בפריסה עברית הדפדפן מדווח את תו הפריסה.
    expect(matchShortcut(key({ key: 'ד', code: 'KeyS', ctrlKey: true }), shortcut('save'))).toBe(
      true,
    );
  });

  it('Ctrl+F בפריסה עברית — key הוא „כ”', () => {
    expect(matchShortcut(key({ key: 'כ', code: 'KeyF', ctrlKey: true }), shortcut('find'))).toBe(
      true,
    );
  });

  it('Ctrl+H בפריסה עברית — key הוא „י”', () => {
    expect(matchShortcut(key({ key: 'י', code: 'KeyH', ctrlKey: true }), shortcut('replace'))).toBe(
      true,
    );
  });

  it('Ctrl+P בפריסה עברית — key הוא „פ”', () => {
    expect(matchShortcut(key({ key: 'פ', code: 'KeyP', ctrlKey: true }), shortcut('print'))).toBe(
      true,
    );
  });
});

describe('matchShortcut — מודיפיירים', () => {
  it('אות בלי מודיפייר אינה קיצור', () => {
    expect(matchShortcut(key({ key: 's', code: 'KeyS' }), shortcut('save'))).toBe(false);
  });

  it('Ctrl+Alt+S אינו Ctrl+S', () => {
    expect(
      matchShortcut(key({ code: 'KeyS', ctrlKey: true, altKey: true }), shortcut('save')),
    ).toBe(false);
  });

  it('Ctrl+Shift+S אינו Ctrl+S, ולהפך', () => {
    const event = key({ code: 'KeyS', ctrlKey: true, shiftKey: true });
    expect(matchShortcut(event, shortcut('save'))).toBe(false);
    expect(matchShortcut(event, shortcut('save-as'))).toBe(true);
    expect(matchShortcut(key({ code: 'KeyS', ctrlKey: true }), shortcut('save-as'))).toBe(false);
  });

  it('Meta שקול ל-Ctrl (macOS)', () => {
    expect(matchShortcut(key({ code: 'KeyS', metaKey: true }), shortcut('save'))).toBe(true);
    expect(
      matchShortcut(key({ code: 'KeyS', metaKey: true, shiftKey: true }), shortcut('save-as')),
    ).toBe(true);
  });

  it('Escape עם מודיפייר אינו Escape', () => {
    expect(matchShortcut(key({ code: 'Escape' }), shortcut('escape'))).toBe(true);
    expect(matchShortcut(key({ code: 'Escape', ctrlKey: true }), shortcut('escape'))).toBe(false);
  });
});

describe('matchShortcut — מקרי קצה', () => {
  const punctuation: Shortcut = {
    id: 'grow-font',
    label: 'Ctrl+]',
    description: 'הגדלת גופן',
    group: 'font',
    key: ']',
    ctrl: true,
    command: 'font-size',
    repeatable: true,
  };

  it('רשומת פיסוק מותאמת לפי key ולא לפי code', () => {
    // ה-code של „]” נודד בין פריסות פיזיות, ולכן שם ההשוואה הפוכה.
    expect(matchShortcut(key({ key: ']', code: 'BracketRight', ctrlKey: true }), punctuation)).toBe(
      true,
    );
    expect(matchShortcut(key({ key: ']', code: 'Digit9', ctrlKey: true }), punctuation)).toBe(true);
    expect(matchShortcut(key({ key: '[', code: 'BracketRight', ctrlKey: true }), punctuation)).toBe(
      false,
    );
  });

  it('הקלדה עם IME אינה מפעילה קיצור', () => {
    expect(
      matchShortcut(key({ code: 'KeyS', ctrlKey: true, isComposing: true }), shortcut('save')),
    ).toBe(false);
    // דפדפן שאינו מציב isComposing מדווח keyCode 229.
    expect(
      matchShortcut(key({ code: 'KeyS', ctrlKey: true, keyCode: 229 }), shortcut('save')),
    ).toBe(false);
  });

  it('מקש מוחזק חוזר רק ברשומה שמתירה זאת', () => {
    expect(
      matchShortcut(key({ code: 'KeyS', ctrlKey: true, repeat: true }), shortcut('save')),
    ).toBe(false);
    expect(matchShortcut(key({ key: ']', ctrlKey: true, repeat: true }), punctuation)).toBe(true);
  });

  it('רשומה בלי מקש אינה מתאימה לכלום', () => {
    const broken = { ...punctuation, key: undefined, code: undefined } as Shortcut;
    expect(matchShortcut(key({ key: 'x', code: 'KeyX', ctrlKey: true }), broken)).toBe(false);
  });
});

describe('matchAny', () => {
  it('מחזירה את הרשומה המתאימה מתוך הרשימה', () => {
    expect(matchAny(key({ key: 'ד', code: 'KeyS', ctrlKey: true }), SHORTCUTS)?.id).toBe('save');
  });

  it('צירוף שאינו ברשימה מחזיר undefined', () => {
    expect(matchAny(key({ code: 'KeyG', ctrlKey: true }), SHORTCUTS)).toBeUndefined();
  });
});
