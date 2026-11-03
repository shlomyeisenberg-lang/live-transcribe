/**
 * לחיצה על רקע של פס מעטפת שאינה לוקחת את הסמן — ui/shell/caret-keeper.ts.
 *
 * ‏`cancelable: true` בכל אירוע: `preventDefault` על אירוע שאינו ניתן לביטול
 * אינו עושה דבר, ובלעדיו הבדיקה עוברת ירוק גם על קוד שאינו מבטל.
 */
import { describe, it, expect, afterEach } from 'vitest';
import {
  createCaretKeeper,
  shouldKeepCaret,
  CARET_KEEPER_CONTROL_SELECTOR,
} from '../../src/ui/shell/caret-keeper';

describe('shouldKeepCaret', () => {
  it('רקע של מעטפת כשהמיקוד במסמך — לבטל', () => {
    expect(shouldKeepCaret({ onControl: false, inDocument: false, documentFocused: true })).toBe(true);
  });

  it('פקד — שייך לו המיקוד', () => {
    expect(shouldKeepCaret({ onControl: true, inDocument: false, documentFocused: true })).toBe(false);
  });

  it('אזור המסמך — שם לחיצה מזיזה סמן', () => {
    expect(shouldKeepCaret({ onControl: false, inDocument: true, documentFocused: true })).toBe(false);
  });

  it('המיקוד אינו במסמך — אין מה לשמור, ושכבה שנסגרת על blur צריכה להיסגר', () => {
    expect(shouldKeepCaret({ onControl: false, inDocument: false, documentFocused: false })).toBe(false);
  });
});

describe('createCaretKeeper', () => {
  let shell: HTMLElement;
  let bar: HTMLElement;
  let barButton: HTMLButtonElement;
  let area: HTMLElement;
  let surface: HTMLTextAreaElement;
  let keep: (event: PointerEvent) => void;

  function setup({ withArea = true } = {}): void {
    shell = document.createElement('div');
    bar = document.createElement('div');
    barButton = document.createElement('button');
    bar.appendChild(barButton);
    area = document.createElement('main');
    /* משטח ההקלדה של המנוע: בתוך אזור המסמך. */
    surface = document.createElement('textarea');
    area.appendChild(surface);
    shell.append(bar, area);
    document.body.appendChild(shell);
    keep = createCaretKeeper({ documentArea: () => (withArea ? area : null) });
    shell.addEventListener('pointerdown', (event) => keep(event as PointerEvent), true);
  }

  afterEach(() => {
    document.body.innerHTML = '';
  });

  /** `MouseEvent` ולא `PointerEvent`: jsdom אינו מגדיר את השני, והמאזין קורא
      רק `button` ו-`target`. */
  function down(target: Element): MouseEvent {
    const event = new MouseEvent('pointerdown', { bubbles: true, cancelable: true, button: 0 });
    target.dispatchEvent(event);
    return event;
  }

  it('רקע הפס: מבוטל, והמיקוד נשאר במסמך', () => {
    setup();
    surface.focus();
    const event = down(bar);
    expect(event.defaultPrevented).toBe(true);
    expect(document.activeElement).toBe(surface);
  });

  it('כפתור בתוך הפס: אינו מבוטל', () => {
    setup();
    surface.focus();
    expect(down(barButton).defaultPrevented).toBe(false);
  });

  it('אזור המסמך: אינו מבוטל', () => {
    setup();
    surface.focus();
    expect(down(area).defaultPrevented).toBe(false);
  });

  it('המיקוד מחוץ למסמך: אינו מבוטל', () => {
    setup();
    barButton.focus();
    expect(down(bar).defaultPrevented).toBe(false);
  });

  it('אין מסמך פתוח: אינו מבוטל', () => {
    setup({ withArea: false });
    expect(down(bar).defaultPrevented).toBe(false);
  });

  it('כפתור שאינו ראשי אינו נוגע', () => {
    setup();
    surface.focus();
    const event = new MouseEvent('pointerdown', { bubbles: true, cancelable: true, button: 2 });
    bar.dispatchEvent(event);
    expect(event.defaultPrevented).toBe(false);
  });

  it('הבורר מכסה את הפקדים שבפסים', () => {
    // אם מישהו יוסיף פקד מסוג חדש לפס, זו השורה שתזכיר לעדכן את הבורר.
    for (const html of [
      '<button></button>',
      '<input>',
      '<select></select>',
      '<textarea></textarea>',
      '<a href="#"></a>',
      '<div role="button"></div>',
      '<div role="tab"></div>',
      '<div role="combobox"></div>',
      '<div tabindex="-1"></div>',
    ]) {
      const holder = document.createElement('div');
      holder.innerHTML = html;
      const el = holder.firstElementChild!;
      expect(el.matches(CARET_KEEPER_CONTROL_SELECTOR), html).toBe(true);
    }
  });
});
