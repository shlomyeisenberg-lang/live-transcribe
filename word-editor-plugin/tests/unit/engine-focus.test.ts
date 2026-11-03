/**
 * `focusDocument` — הדרך היחידה להחזיר את הסמן לטקסט.
 *
 * היא נקראת מתוך טיפול במקש, ולכן החוזה החשוב שלה אינו „היא ממקדת” אלא „היא
 * לעולם אינה זורקת”: חריגה שם מפילה את המאזין הגלובלי, כלומר את כל הקיצורים
 * ולא רק את זה שנכשל.
 */
import { describe, it, expect, vi } from 'vitest';
import { focusDocument } from '../../src/engine/focus';

describe('focusDocument', () => {
  it('מבקשת מהמנוע למקד, עם שחזור הבחירה', () => {
    // בלי `restoreSelection` החזרה נופלת על תחילת המסמך ולא על מה שהמשתמש סימן.
    const focus = vi.fn();

    expect(focusDocument({ focus })).toBe(true);
    expect(focus).toHaveBeenCalledExactlyOnceWith({ restoreSelection: true });
  });

  it('בלי מסמך פתוח — מדווחת שלא הצליחה, ואינה נופלת', () => {
    expect(focusDocument(null)).toBe(false);
    expect(focusDocument(undefined)).toBe(false);
    expect(focusDocument({})).toBe(false);
  });

  it('גרסת מנוע שאינה חושפת מיקוד', () => {
    expect(focusDocument({ focus: 'not-a-function' } as never)).toBe(false);
  });

  it('מנוע שזורק אינו מפיל את מי שקרא', () => {
    const focus = vi.fn(() => {
      throw new Error('runtime not ready');
    });

    expect(focusDocument({ focus })).toBe(false);
  });

  it('נקראת על המופע עצמו — `this` הוא המנוע', () => {
    // `superdoc.focus()` נשען על המצב הפנימי של המופע. קריאה מנותקת הייתה
    // זורקת בפועל, ולא בבדיקה שמחזיקה כפיל.
    const host = {
      seen: null as unknown,
      focus(this: unknown) {
        (host as { seen: unknown }).seen = this;
      },
    };

    focusDocument(host);

    expect(host.seen).toBe(host);
  });
});
