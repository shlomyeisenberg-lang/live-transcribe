/**
 * מברשת עיצוב — הבאג היה שהחימוש (`copy-format` דולק) לא מלווה בהחלה,
 * כי `ui.formatPainter` הוא "DOM listener coordination" בלבד: מישהו חייב
 * לקרוא ל-`notifyPointerUp`/`notifyKeyUp` כדי שההחלה תרוץ, וכש-`ui: false`
 * אין `SuperToolbar` שעושה את זה.
 *
 * אבל קריאה ל-`notify*Up` על **כל** `pointerup`/`keyup` (כמו ש-`SuperToolbar`
 * עצמו עושה) מדידה הראתה שתי בעיות: (1) לחיצה שרק ממקמת סמן, או מקש חץ בלי
 * Shift שרק מזיז אותו, מפעילים במנוע מסלול "צביעת פסקה על הסמן" שמכבה את
 * המצב החמוש גם כשאין שום שינוי לצייר — עוד לפני שבחירת היעד האמיתית
 * מתחילה; (2) גם כשמסננים את אלה, הרחבת בחירה בכמה לחיצות מקש *נפרדות*
 * (לא החזקה אחת) מייצרת כמה `keyup` נפרדים, וההחלה שמופעלת מהראשון שבהם
 * מסיימת את המצב החמוש אחרי תו אחד בלבד — לפני שהלחיצות הבאות מרחיבות את
 * הטווח. ראו את ההסבר המלא בראש format-painter.ts.
 *
 * הבדיקות כאן בודקות את שלוש השכבות: שהמאזינים נרשמים ומשתחררים כהלכה,
 * שהסייג עצמו עובד (קליק/ניווט תמים לא קורא ל-`notify*Up`, ורק גרירה
 * אמיתית או Shift כן), ושהדבאונס מאחד רצף לחיצות להחלה אחת בסוף.
 */
import { describe, it, expect, afterEach, beforeEach, vi } from 'vitest';
import {
  installFormatPainter,
  type FormatPainterHandle,
  type FormatPainterHost,
} from '../../src/engine/format-painter';

/** אחרי הדבאונס (250ms) — נדיב כדי לא להיות שביר לערך המדויק. */
const AFTER_DEBOUNCE_MS = 400;

function container(): HTMLElement {
  const element = document.createElement('div');
  document.body.append(element);
  return element;
}

interface Calls {
  setPointerSelecting: boolean[];
  notifyPointerUp: number;
  setKeyboardSelecting: boolean[];
  notifyKeyUp: number;
  cancel: number;
}

/** כפיל של `ui.formatPainter`, עם יומן קריאות. */
function painterDouble(): { host: FormatPainterHost; calls: Calls } {
  const calls: Calls = {
    setPointerSelecting: [],
    notifyPointerUp: 0,
    setKeyboardSelecting: [],
    notifyKeyUp: 0,
    cancel: 0,
  };

  const host: FormatPainterHost = {
    ui: {
      formatPainter: {
        setPointerSelecting: (flag: boolean) => calls.setPointerSelecting.push(flag),
        notifyPointerUp: () => {
          calls.notifyPointerUp += 1;
        },
        setKeyboardSelecting: (flag: boolean) => calls.setKeyboardSelecting.push(flag),
        notifyKeyUp: () => {
          calls.notifyKeyUp += 1;
        },
        cancel: () => {
          calls.cancel += 1;
        },
      },
    },
  };

  return { host, calls };
}

let installed: FormatPainterHandle | null = null;

beforeEach(() => {
  vi.useFakeTimers();
});

afterEach(() => {
  installed?.dispose();
  installed = null;
  document.body.innerHTML = '';
  vi.useRealTimers();
});

/** לחיצה בנקודה אחת — מיקום סמן, לא גרירה. */
function click(target: HTMLElement, x = 10, y = 10, extra: MouseEventInit = {}): void {
  target.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, clientX: x, clientY: y }));
  target.dispatchEvent(new MouseEvent('pointerup', { bubbles: true, clientX: x, clientY: y, ...extra }));
}

/** גרירה אמיתית — נקודת שחרור רחוקה מנקודת הלחיצה. */
function drag(target: HTMLElement, fromX = 10, fromY = 10, toX = 200, toY = 10): void {
  target.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, clientX: fromX, clientY: fromY }));
  target.dispatchEvent(new MouseEvent('pointerup', { bubbles: true, clientX: toX, clientY: toY }));
}

/** Shift+מקש ניווט: לחיצה ושחרור בודדים. */
function shiftKey(target: HTMLElement, key: string): void {
  target.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key, shiftKey: true }));
  target.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key, shiftKey: true }));
}

describe('חיווט מברשת העיצוב', () => {
  it('pointerdown מזין setPointerSelecting(true)', () => {
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    root.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, clientX: 10, clientY: 10 }));

    expect(calls.setPointerSelecting).toEqual([true]);
    expect(calls.notifyPointerUp).toBe(0);
  });

  it('קליק במקום אחד (בלי גרירה, בלי Shift) מכבה pointerSelecting אבל לא מחיל', () => {
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    // בדיוק המקרה שנמדד כבאג: caretPara ואז Home בלי Shift, לפני בחירת היעד.
    click(root);
    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);

    expect(calls.setPointerSelecting).toEqual([true, false]);
    expect(calls.notifyPointerUp).toBe(0);
  });

  it('גרירה אמיתית (המצביע זז מעבר לסף) כן מחילה דרך notifyPointerUp, אחרי דבאונס', () => {
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    drag(root);
    expect(calls.notifyPointerUp).toBe(0); // עוד לא — הדבאונס טרם חלף
    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);

    expect(calls.setPointerSelecting).toEqual([true, false]);
    expect(calls.notifyPointerUp).toBe(1);
  });

  it('Shift+קליק (בלי גרירה) כן מחיל — מרחיב בחירה קיימת', () => {
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    click(root, 10, 10, { shiftKey: true });
    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);

    expect(calls.notifyPointerUp).toBe(1);
  });

  it('זעזוע קטן מתחת לסף עדיין נחשב קליק ולא גרירה', () => {
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    root.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true, clientX: 10, clientY: 10 }));
    root.dispatchEvent(new MouseEvent('pointerup', { bubbles: true, clientX: 12, clientY: 9 }));
    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);

    expect(calls.notifyPointerUp).toBe(0);
  });

  it('Shift+מקש חיצים מזין setKeyboardSelecting(true), ו-keyup מכבה ומחיל אחרי דבאונס', () => {
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    root.dispatchEvent(
      new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight', shiftKey: true }),
    );
    expect(calls.setKeyboardSelecting).toEqual([true]);
    expect(calls.notifyKeyUp).toBe(0);

    root.dispatchEvent(
      new KeyboardEvent('keyup', { bubbles: true, key: 'ArrowRight', shiftKey: true }),
    );
    expect(calls.setKeyboardSelecting).toEqual([true, false]);
    expect(calls.notifyKeyUp).toBe(0); // עוד לא — הדבאונס טרם חלף

    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);
    expect(calls.notifyKeyUp).toBe(1);
  });

  it('מקש חיצים בלי Shift מזיז סמן — לא מחיל, בדיוק המקרה שנמדד כבאג', () => {
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    root.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight' }));
    root.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'ArrowRight' }));
    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);

    expect(calls.setKeyboardSelecting).toEqual([true, false]);
    expect(calls.notifyKeyUp).toBe(0);
  });

  it('Home בלי Shift (מיקום סמן לפני בחירת יעד) לא מחיל', () => {
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    root.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Home' }));
    root.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Home' }));
    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);

    expect(calls.notifyKeyUp).toBe(0);
  });

  it('רצף Shift+חץ נפרדים (לחיצה-שחרור, לחיצה-שחרור...) מחיל פעם אחת בלבד, בסוף', () => {
    // בדיוק איך selectRange בשער ה-QA מרחיבה בחירה: כמה לחיצות בודדות, לא
    // החזקה רציפה אחת. בלי הדבאונס — ה-keyup הראשון היה כבר מסיים את המצב
    // החמוש אחרי תו אחד.
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    for (let i = 0; i < 4; i += 1) {
      shiftKey(root, 'ArrowRight');
      vi.advanceTimersByTime(22); // בדיוק הקצב שנמדד בשער ה-QA
    }
    expect(calls.notifyKeyUp).toBe(0); // כל לחיצה איפסה את הטיימר

    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);
    expect(calls.notifyKeyUp).toBe(1); // פעם אחת בלבד, אחרי שהרצף נגמר
  });

  it('Home/End/PageUp/PageDown נחשבים מקשי בחירה ב-keydown', () => {
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    for (const key of ['Home', 'End', 'PageUp', 'PageDown', 'ArrowLeft', 'ArrowUp', 'ArrowDown']) {
      root.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key }));
    }

    expect(calls.setKeyboardSelecting).toEqual(Array(7).fill(true));
  });

  it('מקש שאינו מקש בחירה לא מפעיל setKeyboardSelecting וגם לא מחיל', () => {
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    root.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'a' }));
    expect(calls.setKeyboardSelecting).toEqual([]);

    root.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'a', shiftKey: true }));
    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);
    expect(calls.setKeyboardSelecting).toEqual([false]);
    expect(calls.notifyKeyUp).toBe(0);
  });

  it('Escape מבטל דרך cancel, מבטל דבאונס ממתין, ואינו נוגע ב-setKeyboardSelecting', () => {
    const { host, calls } = painterDouble();
    const root = container();
    installed = installFormatPainter(root, host);

    shiftKey(root, 'ArrowRight'); // דבאונס ממתין
    root.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Escape' }));
    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);

    expect(calls.cancel).toBe(1);
    expect(calls.notifyKeyUp).toBe(0); // הדבאונס לא הופעל אחרי הביטול
  });

  it('dispose מסיר את כל המאזינים ומבטל דבאונס ממתין', () => {
    const { host, calls } = painterDouble();
    const root = container();
    const handle = installFormatPainter(root, host);

    drag(root); // מתזמן דבאונס
    handle.dispose();
    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);

    root.dispatchEvent(
      new KeyboardEvent('keydown', { bubbles: true, key: 'ArrowRight', shiftKey: true }),
    );
    root.dispatchEvent(
      new KeyboardEvent('keyup', { bubbles: true, key: 'ArrowRight', shiftKey: true }),
    );
    root.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Escape' }));
    vi.advanceTimersByTime(AFTER_DEBOUNCE_MS);

    expect(calls.setPointerSelecting).toEqual([true, false]); // מה-drag, לפני ה-dispose
    expect(calls.notifyPointerUp).toBe(0); // הדבאונס בוטל ב-dispose
    expect(calls.setKeyboardSelecting).toEqual([]);
    expect(calls.notifyKeyUp).toBe(0);
    expect(calls.cancel).toBe(0);
  });

  it('מארח שאין בו את משטח ה-formatPainter המלא לא רושם מאזין וגם לא זורק', () => {
    const root = container();
    installed = installFormatPainter(root, null);
    installFormatPainter(root, {}).dispose();
    installFormatPainter(root, { ui: { formatPainter: { setPointerSelecting: () => undefined } } }).dispose();

    // dispose לא זורק גם כשלא נרשם כלום
    expect(() => installed?.dispose()).not.toThrow();
  });
});
