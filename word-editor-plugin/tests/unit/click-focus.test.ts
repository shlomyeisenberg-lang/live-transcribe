/**
 * לחיצה מחוץ לעמוד והמיקוד שנשאר במסמך — engine/click-focus.ts.
 *
 * שתי שכבות: ההכרעה הטהורה, וההחלה על אירועים ב-jsdom מול מודד-כפיל. מה
 * שנמדד בדפדפן האמיתי ומה שהתיקון מבטיח — בהערת הפתיחה של המודול.
 *
 * ‏`cancelable: true` בכל אירוע: `preventDefault` על אירוע שאינו ניתן לביטול
 * אינו עושה דבר, ובלי הדגל הבדיקה הייתה עוברת ירוק גם על קוד שאינו מבטל.
 */
import { describe, it, expect, afterEach } from 'vitest';
import {
  clickFocusAction,
  installClickFocus,
  type ClickFocusHandle,
} from '../../src/engine/click-focus';

describe('clickFocusAction', () => {
  it('המיקוד במסמך ולחיצה מחוץ לעמוד — לבטל, כדי שלא יאבד', () => {
    expect(clickFocusAction({ onPage: false, focused: true })).toBe('keep');
  });

  it('המיקוד במסמך ולחיצה על העמוד — לא לגעת', () => {
    expect(clickFocusAction({ onPage: true, focused: true })).toBe('none');
  });

  it('המיקוד מחוץ למסמך — לוודא שהוא חוזר, בשני המקרים', () => {
    expect(clickFocusAction({ onPage: false, focused: false })).toBe('ensure');
    expect(clickFocusAction({ onPage: true, focused: false })).toBe('ensure');
  });

  it('אין עמודים בעץ — אין הכרעה, ולא „מחוץ לעמוד”', () => {
    expect(clickFocusAction({ onPage: null, focused: true })).toBe('none');
  });
});

describe('installClickFocus', () => {
  let handle: ClickFocusHandle | null = null;
  let container: HTMLElement;
  let inside: HTMLElement;
  let outsideButton: HTMLButtonElement;
  let focusCalls: Array<{ restoreSelection?: boolean }>;
  /** מה שהשומר נשאל: הנקודות שנמדדו. */
  let asked: Array<{ x: number; y: number }>;
  /** הבדיקות שנדחו ל-task הבא, ומה שמריץ אותן. */
  let pending: Array<() => void>;

  /** העמוד המצויר: 100..900 אופקית, 200..1300 אנכית. */
  const PAGE = { left: 100, top: 200, right: 900, bottom: 1300 };
  const onPageAt = (x: number, y: number): boolean | null => {
    asked.push({ x, y });
    return x >= PAGE.left && x <= PAGE.right && y >= PAGE.top && y <= PAGE.bottom;
  };

  const host = {
    focus(options?: { restoreSelection?: boolean }) {
      focusCalls.push(options ?? {});
      /* המנוע האמיתי מחזיר את המיקוד למשטח ההקלדה שבתוך ה-container. */
      inside.focus();
    },
  };

  function setup(options: Parameters<typeof installClickFocus>[2] = {}): void {
    asked = [];
    focusCalls = [];
    pending = [];
    container = document.createElement('div');
    /* משטח ההקלדה של המנוע: בתוך ה-container, וממוקד. */
    inside = document.createElement('textarea');
    container.appendChild(inside);
    outsideButton = document.createElement('button');
    document.body.append(container, outsideButton);
    handle = installClickFocus(container, host, {
      onPageAt,
      /* ה-task שאחרי האירוע: הדפדפן והמנוע כבר סיימו להזיז מיקוד. `flush`
         הוא הגבול הזה, ולכן בדיקה סינכרונית אינה מדווחת „לא חזר” על מיקוד
         שכן חזר. */
      schedule: (run) => void pending.push(run),
      ...options,
    });
  }

  /** מריץ את מה שנדחה ל-task הבא. */
  function flush(): void {
    for (const run of pending.splice(0)) run();
  }

  afterEach(() => {
    handle?.dispose();
    handle = null;
    document.body.innerHTML = '';
  });

  /** לחיצה אמיתית: `pointerdown` שניתן לביטול, על המטרה שנמסרה. */
  function down(target: Element, x: number, y: number, button = 0): MouseEvent {
    const event = new MouseEvent('pointerdown', {
      bubbles: true,
      cancelable: true,
      button,
      clientX: x,
      clientY: y,
    });
    target.dispatchEvent(event);
    return event;
  }

  it('המיקוד במסמך: לחיצה מחוץ לעמוד מבוטלת, ולכן המיקוד נשאר', () => {
    setup();
    inside.focus();
    const event = down(container, 40, 500);
    expect(event.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(inside);
    expect(focusCalls).toEqual([]);
  });

  it('המיקוד במסמך: לחיצה על העמוד אינה מבוטלת', () => {
    setup();
    inside.focus();
    const event = down(inside, 500, 700);
    expect(event.defaultPrevented).toBe(false);
    expect(focusCalls).toEqual([]);
  });

  it('המיקוד ברצועה: הלחיצה אינה מבוטלת, והמיקוד מוחזר למסמך', () => {
    setup();
    outsideButton.focus();
    const event = down(container, 40, 500);
    /* לא מבוטלת — אחרת רשימת הגופנים, שנסגרת על blur, הייתה נשארת פתוחה. */
    expect(event.defaultPrevented).toBe(false);
    flush();
    expect(focusCalls).toEqual([{ restoreSelection: true }]);
    expect(document.activeElement).toBe(inside);
  });

  it('המיקוד ברצועה ולחיצה על העמוד שלא הזיזה מיקוד — מוחזר גם אז', () => {
    // זה מה שנמדד על רצועת הכותרת של העמוד: לחיצה שאינה מזיזה מיקוד בכלל.
    setup();
    outsideButton.focus();
    down(container, 500, 700);
    flush();
    expect(focusCalls).toEqual([{ restoreSelection: true }]);
  });

  it('לחיצה על העמוד שכן מיקדה את המסמך — אין החזרה מיותרת', () => {
    setup();
    outsideButton.focus();
    const event = new MouseEvent('pointerdown', {
      bubbles: true,
      cancelable: true,
      button: 0,
      clientX: 500,
      clientY: 700,
    });
    /* כמו המנוע: הוא ממקד את משטח ההקלדה בתוך הטיפול באירוע. */
    container.addEventListener('pointerdown', () => inside.focus(), { once: true });
    container.dispatchEvent(event);
    flush();
    expect(focusCalls).toEqual([]);
    expect(document.activeElement).toBe(inside);
  });

  it('כפתור שאינו ראשי אינו נוגע', () => {
    setup();
    inside.focus();
    const event = down(container, 40, 500, 2);
    expect(event.defaultPrevented).toBe(false);
    expect(asked).toEqual([]);
  });

  it('לחיצה מחוץ ל-container אינה נוגעת', () => {
    setup();
    inside.focus();
    const event = down(outsideButton, 40, 500);
    expect(event.defaultPrevented).toBe(false);
    expect(asked).toEqual([]);
  });

  it('אין עמודים מצוירים: הלחיצה נשארת כפי שהיא', () => {
    setup({ onPageAt: () => null });
    inside.focus();
    const event = down(container, 40, 500);
    expect(event.defaultPrevented).toBe(false);
    flush();
    expect(focusCalls).toEqual([]);
  });

  it('אחרי dispose שום דבר אינו מבוטל', () => {
    setup();
    inside.focus();
    handle?.dispose();
    handle = null;
    const event = down(container, 40, 500);
    expect(event.defaultPrevented).toBe(false);
    expect(asked).toEqual([]);
  });
});
