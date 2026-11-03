/**
 * שליפת גופן לפי דרישה — `ensureFamilyDrawable`.
 *
 * הכול מוזרק: `isFamilyAvailable` האמיתי מחזיר `true` בלי canvas („בלי מדידה
 * אין לנו מה לומר”), ובדיקה שתסמוך עליו ב-jsdom לא תגיע לשום מסלול — היא
 * תחזיר „זמין” בשורה הראשונה ותעבור ירוק על קוד שבור.
 */
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  ensureFamilyDrawable,
  onPickerFontsChanged,
  PICKER_FONT_STYLE_ID,
  resetPickerFonts,
} from '../../src/engine/picker-fonts';

/** מכונה שפותרת רק את מה שברשימה, ומתעדכנת כשמזריקים. */
function machine(resolvable: Set<string>) {
  return (name: string) => resolvable.has(name.toLowerCase());
}

/** מארח שמגיש בייטים רק למה שברשימה. מחזיר גם את מה שנשאל. */
function host(servable: readonly string[]) {
  const asked: string[] = [];
  const call = vi.fn(async (method: string, params?: unknown) => {
    if (method !== 'fonts.resolveFamilies') return null;
    const families = (params as { families?: { name: string; substitutes?: string[] }[] })?.families ?? [];
    const name = families[0]?.name ?? '';
    asked.push(name);
    if (!servable.includes(name)) return { css: '', resolved: [] };
    return {
      css: `@font-face{font-family:${JSON.stringify(name)};src:url(data:font/ttf;base64,AAAA);}`,
      resolved: [name],
    };
  });
  return { call: call as never, asked, params: () => call.mock.calls.map((c) => c[1]) };
}

function styleText(): string {
  return document.getElementById(PICKER_FONT_STYLE_ID)?.textContent ?? '';
}

beforeEach(() => {
  resetPickerFonts();
  document.getElementById(PICKER_FONT_STYLE_ID)?.remove();
});

afterEach(() => {
  resetPickerFonts();
  document.getElementById(PICKER_FONT_STYLE_ID)?.remove();
});

describe('ensureFamilyDrawable', () => {
  it('גופן שהדפדפן כבר פותר אינו נשלף — ואוצריא אינה נשאלת', async () => {
    const h = host(['David']);
    const ok = await ensureFamilyDrawable('David', {
      call: h.call,
      available: machine(new Set(['david'])),
    });

    expect(ok).toBe(true);
    expect(h.asked).toEqual([]);
    expect(styleText()).toBe('');
  });

  it('גופן שאינו נפתר נשלף, מוזרק, ונמדד שוב', async () => {
    const resolvable = new Set<string>();
    const h = host(['Guttman Kav-Light']);
    // ההזרקה היא מה שהופך את התשובה: `forget` מדמה את שכחת המדידה הישנה,
    // ואחריה המכונה כבר פותרת.
    const ok = await ensureFamilyDrawable('Guttman Kav-Light', {
      call: h.call,
      available: machine(resolvable),
      forget: (name) => resolvable.add(name.toLowerCase()),
      load: async () => {},
    });

    expect(ok).toBe(true);
    expect(h.asked).toEqual(['Guttman Kav-Light']);
    expect(styleText()).toContain('Guttman Kav-Light');
  });

  it('נשלח השם כתחליף של עצמו — אין הגשה של גופן אחר תחת שם המשתמש', async () => {
    const h = host(['Segoe UI Semilight']);
    await ensureFamilyDrawable('Segoe UI Semilight', {
      call: h.call,
      available: machine(new Set()),
      forget: () => {},
      load: async () => {},
    });

    expect(h.params()[0]).toEqual({
      families: [{ name: 'Segoe UI Semilight', substitutes: ['Segoe UI Semilight'] }],
    });
  });

  it('מארח בלי בייטים — נשאל פעם אחת בלבד, גם על בחירה חוזרת', async () => {
    const h = host([]);
    const deps = { call: h.call, available: machine(new Set()), forget: () => {}, load: async () => {} };

    expect(await ensureFamilyDrawable('אין כזה', deps)).toBe(false);
    expect(await ensureFamilyDrawable('אין כזה', deps)).toBe(false);
    // הסירוב נגזר ממה שמותקן בדיסק ואינו משתנה תוך ההפעלה, ולכן נזכר.
    expect(h.asked).toEqual(['אין כזה']);
  });

  it('שתי בחירות רצופות של אותו שם הן בקשה אחת', async () => {
    const resolvable = new Set<string>();
    const h = host(['Rashi']);
    const deps = {
      call: h.call,
      available: machine(resolvable),
      forget: (name: string) => resolvable.add(name.toLowerCase()),
      load: async () => {},
    };

    const [a, b] = await Promise.all([
      ensureFamilyDrawable('Rashi', deps),
      ensureFamilyDrawable('Rashi', deps),
    ]);

    expect([a, b]).toEqual([true, true]);
    expect(h.asked).toEqual(['Rashi']);
  });

  it('הכללים נצברים — בחירה שנייה אינה מוחקת את הגופן של הראשונה', async () => {
    const resolvable = new Set<string>();
    const h = host(['First', 'Second']);
    const deps = {
      call: h.call,
      available: machine(resolvable),
      forget: (name: string) => resolvable.add(name.toLowerCase()),
      load: async () => {},
    };

    await ensureFamilyDrawable('First', deps);
    await ensureFamilyDrawable('Second', deps);

    expect(styleText()).toContain('First');
    expect(styleText()).toContain('Second');
  });

  it('בייטים שהדפדפן לא קיבל — לא מדווח הצלחה ולא מודיע', async () => {
    const h = host(['Broken']);
    const listener = vi.fn();
    onPickerFontsChanged(listener);

    // `forget` שאינו משנה דבר: המדידה שאחרי ההזרקה עדיין שלילית.
    const ok = await ensureFamilyDrawable('Broken', {
      call: h.call,
      available: machine(new Set()),
      forget: () => {},
      load: async () => {},
    });

    expect(ok).toBe(false);
    expect(listener).not.toHaveBeenCalled();
  });

  it('המנוי נקרא כשמשפחה הפכה ניתנת לציור, ומבוטל כשמסירים אותו', async () => {
    const resolvable = new Set<string>();
    const h = host(['One', 'Two']);
    const listener = vi.fn();
    const stop = onPickerFontsChanged(listener);
    const deps = {
      call: h.call,
      available: machine(resolvable),
      forget: (name: string) => resolvable.add(name.toLowerCase()),
      load: async () => {},
    };

    await ensureFamilyDrawable('One', deps);
    expect(listener).toHaveBeenCalledTimes(1);

    stop();
    await ensureFamilyDrawable('Two', deps);
    expect(listener).toHaveBeenCalledTimes(1);
  });

  it('כשל של המארח אינו זורק', async () => {
    const ok = await ensureFamilyDrawable('Whatever', {
      call: (async () => {
        throw new Error('DB unavailable');
      }) as never,
      available: machine(new Set()),
      forget: () => {},
      load: async () => {},
    });

    expect(ok).toBe(false);
  });

  it('שם ריק אינו מגיע לאוצריא', async () => {
    const h = host(['x']);
    expect(await ensureFamilyDrawable('   ', { call: h.call, available: machine(new Set()) })).toBe(
      false,
    );
    expect(h.asked).toEqual([]);
  });
});
