/**
 * כיווניות פסקה — הקיצור היחיד בתוסף שנופל על **שחרור** מודיפייר.
 *
 * הוא גם הקיצור שהכי קל לשבור בו משהו אחר: כל `Ctrl+Shift+משהו` בממשק נגמר
 * בשחרור של אותו `Shift` בדיוק, ולכן זיהוי רשלני היה הופך את כיוון הפסקה בכל
 * פעם שמשתמש לוחץ „קו חוצה”.
 */
import { describe, it, expect, vi } from 'vitest';
import {
  createDirectionDetector,
  createDirectionShortcut,
  DIRECTION_COMMANDS,
  type ModifierEventLike,
} from '../../src/ui/shortcuts/direction';
import { SHORTCUTS, type Shortcut } from '../../src/ui/shortcuts/registry';

function shift(side: 'ShiftLeft' | 'ShiftRight', over: Partial<ModifierEventLike> = {}) {
  return { code: side, key: 'Shift', ctrlKey: true, metaKey: false, ...over };
}

function letter(code: string, over: Partial<ModifierEventLike> = {}) {
  return { code, key: code.slice(3).toLowerCase(), ctrlKey: true, metaKey: false, ...over };
}

describe('זיהוי הכיווניות', () => {
  it('Ctrl + Shift ימני — מימין לשמאל', () => {
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight'));

    expect(detector.keyup(shift('ShiftRight'))).toBe('rtl');
  });

  it('Ctrl + Shift שמאלי — משמאל לימין', () => {
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftLeft'));

    expect(detector.keyup(shift('ShiftLeft'))).toBe('ltr');
  });

  it('Shift בלי Ctrl אינו כיווניות', () => {
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight', { ctrlKey: false }));

    expect(detector.keyup(shift('ShiftRight', { ctrlKey: false }))).toBeNull();
  });

  it('Ctrl+Shift+X אינו הופך את כיוון הפסקה', () => {
    // הבאג שהמכונה הזאת קיימת בשבילו: „קו חוצה” נגמר בשחרור אותו Shift.
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftLeft'));
    detector.keydown(letter('KeyX', { ctrlKey: true }));

    expect(detector.keyup(shift('ShiftLeft'))).toBeNull();
  });

  it('Shift נלחץ לפני Ctrl — עדיין נחשב', () => {
    // סדר הלחיצה אינו משהו שמשתמש חושב עליו. מה שקובע הוא ש-Ctrl לחוץ ברגע
    // השחרור.
    //
    // **הרצף כאן הוא המלא, כולל ה-keydown של Control עצמו.** הגרסה הראשונה
    // של הבדיקה דילגה עליו, ולכן אישרה בירוק התנהגות שאינה קיימת: לחיצת
    // Control היא אירוע keydown ככל אחר, ו-„כל מקש אחר מבטל” ניקה בגללה את
    // החימוש.
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight', { ctrlKey: false }));
    detector.keydown({ code: 'ControlLeft', key: 'Control', ctrlKey: true, metaKey: false });

    expect(detector.keyup(shift('ShiftRight'))).toBe('rtl');
  });

  it('Meta שנלחץ אחרי ה-Shift אינו מבטל', () => {
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight', { ctrlKey: false }));
    detector.keydown({ code: 'MetaLeft', key: 'Meta', ctrlKey: false, metaKey: true });

    expect(detector.keyup(shift('ShiftRight', { ctrlKey: false, metaKey: true }))).toBe('rtl');
  });

  it('Alt שנלחץ אחרי ה-Shift כן מבטל — זה AltGr', () => {
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight'));
    detector.keydown({ code: 'AltRight', key: 'Alt', ctrlKey: true, metaKey: false, altKey: true });

    expect(detector.keyup(shift('ShiftRight'))).toBeNull();
  });

  it('מקש אמיתי אחרי ה-Shift עדיין מבטל', () => {
    // ההגנה שהמכונה קיימת בשבילה לא נחלשה: Ctrl+Shift+X הוא קו חוצה.
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight'));
    detector.keydown({ code: 'ControlLeft', key: 'Control', ctrlKey: true, metaKey: false });
    detector.keydown(letter('KeyX'));

    expect(detector.keyup(shift('ShiftRight'))).toBeNull();
  });

  it('AltGr אינו מחמש כיווניות', () => {
    // AltGr מדווח ctrl+alt, והוא שכבת תווים בפריסה העברית הסטנדרטית.
    const detector = createDirectionDetector();
    const withAlt = shift('ShiftRight', { altKey: true });

    detector.keydown(withAlt);

    expect(detector.keyup(withAlt)).toBeNull();
  });

  it('שני Shift-ים: שחרור הראשון מבטל, ואין פעולה כפולה', () => {
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight'));
    detector.keydown(shift('ShiftLeft'));

    // הלחיצה השנייה החליפה את החימוש: זה מה שמשוחרר, וזה מה שנספר.
    expect(detector.keyup(shift('ShiftLeft'))).toBe('ltr');
    expect(detector.keyup(shift('ShiftRight'))).toBeNull();
  });

  it('reset מבטל חימוש שנתקע', () => {
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight'));
    detector.reset();

    expect(detector.keyup(shift('ShiftRight'))).toBeNull();
  });

  it('Ctrl ששוחרר לפני ה-Shift מבטל', () => {
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight'));

    expect(detector.keyup(shift('ShiftRight', { ctrlKey: false }))).toBeNull();
  });

  it('שחרור של Shift אחר מזה שנלחץ אינו נחשב', () => {
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight'));

    expect(detector.keyup(shift('ShiftLeft'))).toBeNull();
  });

  it('בלי code — נופל ל-location', () => {
    const detector = createDirectionDetector();
    const right = { key: 'Shift', location: 2, ctrlKey: true, metaKey: false };

    detector.keydown(right);

    expect(detector.keyup(right)).toBe('rtl');
  });

  it('Shift בלי צד ידוע אינו מנחש כיוון', () => {
    const detector = createDirectionDetector();
    const unknown = { key: 'Shift', ctrlKey: true, metaKey: false };

    detector.keydown(unknown);

    expect(detector.keyup(unknown)).toBeNull();
  });

  it('שחרור חוזר בלי לחיצה חדשה אינו יורה שוב', () => {
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight'));

    expect(detector.keyup(shift('ShiftRight'))).toBe('rtl');
    expect(detector.keyup(shift('ShiftRight'))).toBeNull();
  });

  it('שתי לחיצות רצופות הן שתי פעולות נפרדות', () => {
    const detector = createDirectionDetector();

    detector.keydown(shift('ShiftRight'));
    expect(detector.keyup(shift('ShiftRight'))).toBe('rtl');

    detector.keydown(shift('ShiftLeft'));
    expect(detector.keyup(shift('ShiftLeft'))).toBe('ltr');
  });

  it('Meta שקול ל-Ctrl', () => {
    const detector = createDirectionDetector();
    const withMeta = shift('ShiftRight', { ctrlKey: false, metaKey: true });

    detector.keydown(withMeta);

    expect(detector.keyup(withMeta)).toBe('rtl');
  });
});

/** יעד מזויף, כדי לוודא שהמאזינים נרשמים ומנותקים. */
function fakeTarget() {
  const listeners: Record<string, EventListener[]> = {};
  return {
    total: () => Object.values(listeners).flat().length,
    fire: (type: string, event: unknown) =>
      (listeners[type] ?? []).forEach((listener) => listener(event as Event)),
    addEventListener: (type: string, listener: unknown) => {
      (listeners[type] ??= []).push(listener as EventListener);
    },
    removeEventListener: (type: string, listener: unknown) => {
      listeners[type] = (listeners[type] ?? []).filter((item) => item !== listener);
    },
  };
}

describe('החיבור לאירועים', () => {
  it('מריץ את הפקודה של הכיוון', () => {
    const runCommand = vi.fn();
    const target = fakeTarget();
    createDirectionShortcut({ runCommand, target });

    target.fire('keydown', shift('ShiftRight'));
    target.fire('keyup', shift('ShiftRight'));

    expect(runCommand).toHaveBeenCalledExactlyOnceWith(DIRECTION_COMMANDS.rtl);
  });

  it('נחסם כשהפוקוס בשדה טקסט של הממשק', () => {
    const runCommand = vi.fn();
    const target = fakeTarget();
    createDirectionShortcut({ runCommand, target, isBlocked: () => true });

    target.fire('keydown', shift('ShiftRight'));
    target.fire('keyup', shift('ShiftRight'));

    expect(runCommand).not.toHaveBeenCalled();
  });

  it('dispose מנתק את כל המאזינים', () => {
    const runCommand = vi.fn();
    const target = fakeTarget();
    const handle = createDirectionShortcut({ runCommand, target });
    // keydown, keyup, blur
    expect(target.total()).toBe(3);

    handle.dispose();
    expect(target.total()).toBe(0);

    target.fire('keydown', shift('ShiftRight'));
    target.fire('keyup', shift('ShiftRight'));
    expect(runCommand).not.toHaveBeenCalled();
  });

  it('אובדן פוקוס מבטל חימוש תלוי באוויר', () => {
    // ה-keyup מגיע לחלון האחר. בלי האיפוס, החזרה ושחרור ה-Shift היו הופכים
    // את כיוון הפסקה בלי שהמשתמש ביקש.
    const runCommand = vi.fn();
    const target = fakeTarget();
    createDirectionShortcut({ runCommand, target });

    target.fire('keydown', shift('ShiftRight'));
    target.fire('blur', {});
    target.fire('keyup', shift('ShiftRight'));

    expect(runCommand).not.toHaveBeenCalled();
  });

  it('מזהי הפקודות הם אלה שב-registry של היכולות', () => {
    expect(DIRECTION_COMMANDS).toEqual({ rtl: 'direction-rtl', ltr: 'direction-ltr' });
  });
});

describe('הרשומות ברשימה', () => {
  it('שתי הכיווניות מתועדות, עם תווית ועם סימון keyup', () => {
    // הן חייבות להיות ברשימה כדי שהתווית ברצועה ובדיאלוג העזרה תבוא ממנה,
    // ומסומנות `onKeyUp` כדי שהמנתב הרגיל לא יירה אותן בלחיצת ה-Shift.
    const entries: readonly Shortcut[] = SHORTCUTS;
    const rtl = entries.find((entry) => entry.id === 'direction-rtl');
    const ltr = entries.find((entry) => entry.id === 'direction-ltr');

    expect(rtl?.command).toBe('direction-rtl');
    expect(ltr?.command).toBe('direction-ltr');
    expect(rtl?.onKeyUp).toBe(true);
    expect(ltr?.onKeyUp).toBe(true);
    expect(rtl?.label).toContain('ימני');
    expect(ltr?.label).toContain('שמאלי');
    expect(rtl?.group).toBe('direction');
  });
});
