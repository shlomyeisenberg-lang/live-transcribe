/**
 * ‏„אין מיקוד — אין סמן” — ui/shell/caret-focus.ts.
 *
 * ‏`document.hasFocus` מוחלף בכל בדיקה, וזה לא נוחות: בסביבת הבדיקה אין חלון
 * אמיתי שאפשר להעביר ממנו מיקוד, וזה **בדיוק** האות היחיד שמבדיל בין „החלון
 * ממוקד” ל„החלון איבד מיקוד” — המדידה בדפדפן אמיתי הראתה ש-`activeElement`
 * אינו זז כשעוברים לחלון אחר. בדיקה שנשענת על מיקוד אלמנטים בלבד הייתה עוברת
 * ירוק על קוד שאינו מסתכל על החלון כלל.
 */
import { describe, it, expect, afterEach, vi } from 'vitest';
import { shouldPaintCaret, watchCaretFocus, type CaretFocusHandle } from '../../src/ui/shell/caret-focus';

describe('shouldPaintCaret', () => {
  it('חלון ממוקד ומיקוד במסמך — מהבהב', () => {
    expect(shouldPaintCaret({ windowFocused: true, documentFocused: true })).toBe(true);
  });

  it('החלון איבד מיקוד — לא, גם אם המיקוד עדיין על המסמך', () => {
    expect(shouldPaintCaret({ windowFocused: false, documentFocused: true })).toBe(false);
  });

  it('המיקוד בפקד של הממשק — לא, גם אם החלון ממוקד', () => {
    expect(shouldPaintCaret({ windowFocused: true, documentFocused: false })).toBe(false);
  });

  it('שניהם כבויים — לא', () => {
    expect(shouldPaintCaret({ windowFocused: false, documentFocused: false })).toBe(false);
  });
});

describe('watchCaretFocus', () => {
  let area: HTMLElement;
  let surface: HTMLTextAreaElement;
  let control: HTMLButtonElement;
  let painted: boolean[];
  let handle: CaretFocusHandle | null = null;
  let hasFocus: ReturnType<typeof vi.spyOn>;

  function setup({ withArea = true, windowFocused = true } = {}): void {
    area = document.createElement('main');
    /* משטח ההקלדה של המנוע: בתוך אזור המסמך. */
    surface = document.createElement('textarea');
    area.appendChild(surface);
    control = document.createElement('button');
    document.body.append(area, control);

    hasFocus = vi.spyOn(document, 'hasFocus').mockReturnValue(windowFocused);
    painted = [];
    handle = watchCaretFocus({
      documentArea: () => (withArea ? area : null),
      onChange: (paint) => painted.push(paint),
    });
  }

  /** „המשתמש עבר לחלון אחר”: `activeElement` אינו זז, רק החלון. */
  function windowBlur(): void {
    hasFocus.mockReturnValue(false);
    window.dispatchEvent(new Event('blur'));
  }

  function windowFocus(): void {
    hasFocus.mockReturnValue(true);
    window.dispatchEvent(new Event('focus'));
  }

  afterEach(() => {
    handle?.dispose();
    handle = null;
    vi.restoreAllMocks();
    document.body.innerHTML = '';
  });

  it('ההתקנה מודדת מיד', () => {
    setup();
    surface.focus();
    expect(painted).toEqual([false, true]);
  });

  it('מעבר לחלון אחר מכבה — והמיקוד לא זז', () => {
    setup();
    surface.focus();
    painted.length = 0;

    windowBlur();

    expect(document.activeElement).toBe(surface);
    expect(painted).toEqual([false]);
  });

  it('חזרה לחלון מדליקה', () => {
    setup();
    surface.focus();
    windowBlur();
    painted.length = 0;

    windowFocus();

    expect(painted).toEqual([true]);
  });

  it('מיקוד שעובר לפקד של הממשק מכבה', () => {
    setup();
    surface.focus();
    painted.length = 0;

    control.focus();

    expect(painted[painted.length - 1]).toBe(false);
  });

  it('מיקוד שנופל על body מכבה — ואין עליו focusin', () => {
    setup();
    surface.focus();
    painted.length = 0;

    surface.blur();

    expect(document.activeElement).toBe(document.body);
    expect(painted[painted.length - 1]).toBe(false);
  });

  it('אין מסמך פתוח — אין סמן', () => {
    setup({ withArea: false });
    surface.focus();
    expect(painted).toEqual([false]);
  });

  it('אירוע שאינו משנה דבר אינו מדווח פעמיים', () => {
    setup();
    surface.focus();
    painted.length = 0;

    windowBlur();
    window.dispatchEvent(new Event('blur'));
    control.focus();

    expect(painted).toEqual([false]);
  });

  it('dispose מנתק — חלון שאיבד מיקוד אחריו אינו מדווח', () => {
    setup();
    surface.focus();
    handle?.dispose();
    handle = null;
    painted.length = 0;

    windowBlur();
    control.focus();

    expect(painted).toEqual([]);
  });
});
