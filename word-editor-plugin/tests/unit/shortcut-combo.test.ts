/**
 * צירוף המקשים כנתון: מה נלכד, איך הוא מוצג, ומה נחשב תפוס.
 *
 * הבדיקה החשובה כאן היא ההיפך מהמובן מאליו — ש**הרג'יסטרי מוגן**. כל טעמה של
 * המערכת האישית הוא שהיא נוספת ואינה דורסת, וההגנה הזאת אינה נשענת על תוויות
 * שאפשר לפרסר אלא על השדות המובנים של הרשומות.
 */
import { describe, expect, it } from 'vitest';
import {
  comboAsShortcut,
  comboFromEvent,
  comboLabel,
  comboSignature,
  isBindableCombo,
  keyName,
  registryOwners,
  type KeyCombo,
} from '../../src/ui/shortcuts/combo';
import { matchShortcut } from '../../src/ui/shortcuts/match';
import { SHORTCUTS } from '../../src/ui/shortcuts/registry';

function event(over: Partial<Parameters<typeof comboFromEvent>[0]> = {}) {
  return { code: 'KeyK', ctrlKey: false, metaKey: false, shiftKey: false, altKey: false, ...over };
}

describe('לכידת צירוף', () => {
  it('צירוף עם Ctrl נלכד', () => {
    expect(comboFromEvent(event({ ctrlKey: true }))).toEqual({
      code: 'KeyK',
      ctrl: true,
      shift: false,
      alt: false,
    });
  });

  it('Meta נלכד כ-Ctrl — מקלדת Mac מקבלת את אותו צירוף', () => {
    expect(comboFromEvent(event({ metaKey: true }))?.ctrl).toBe(true);
  });

  it('אות בלי מודיפייר אינה צירוף — היא הקלדה', () => {
    expect(comboFromEvent(event())).toBeNull();
  });

  it('Shift לבדו אינו מודיפייר לקשירה', () => {
    expect(comboFromEvent(event({ shiftKey: true }))).toBeNull();
  });

  it('מקש פונקציה נלכד גם בלי מודיפייר', () => {
    expect(comboFromEvent(event({ code: 'F9' }))).toEqual({
      code: 'F9',
      ctrl: false,
      shift: false,
      alt: false,
    });
  });

  it('הקשה על המודיפייר עצמו אינה נלכדת', () => {
    // בלי זה הלוכד היה „נסגר” על Ctrl ברגע שהמשתמש מתחיל להקיש את הצירוף.
    for (const code of ['ControlLeft', 'ShiftRight', 'AltLeft', 'MetaLeft']) {
      expect(comboFromEvent(event({ code, ctrlKey: true })), code).toBeNull();
    }
  });

  it('F13 אינו מקש פונקציה לצורך קשירה בלי מודיפייר', () => {
    // התחום הוא F1–F12: מקלדות שמדווחות F13 ומעלה עושות זאת לרוב כמיפוי של
    // מקש אחר, ואין להתיר קשירה בלי מודיפייר על מה שאינו מקש פיזי מוכר.
    expect(isBindableCombo({ code: 'F13', ctrl: false, shift: false, alt: false })).toBe(false);
    expect(isBindableCombo({ code: 'F13', ctrl: true, shift: false, alt: false })).toBe(true);
  });
});

describe('תווית הצירוף', () => {
  it('הסדר הוא Ctrl→Shift→Alt, כמו כל תוויות הרג׳יסטרי', () => {
    expect(comboLabel({ code: 'KeyK', ctrl: true, shift: true, alt: true })).toBe(
      'Ctrl+Shift+Alt+K',
    );
  });

  it('אות, ספרה ומקש מיוחד מוצגים בשם שכתוב על המקש', () => {
    expect(keyName('KeyQ')).toBe('Q');
    expect(keyName('Digit7')).toBe('7');
    expect(keyName('Numpad3')).toBe('Numpad 3');
    expect(keyName('F8')).toBe('F8');
    expect(keyName('BracketLeft')).toBe('[');
    expect(keyName('ArrowUp')).toBe('↑');
  });

  it('קוד שאינו במפה מוצג כמו שהוא, ולא כ„לא ידוע”', () => {
    expect(keyName('IntlBackslash')).toBe('IntlBackslash');
  });
});

describe('חתימה והתאמה', () => {
  it('שני צירופים זהים מקבלים אותה חתימה, ושונים — לא', () => {
    const a: KeyCombo = { code: 'KeyK', ctrl: true, shift: false, alt: false };
    const b: KeyCombo = { code: 'KeyK', ctrl: true, shift: false, alt: false };
    const c: KeyCombo = { code: 'KeyK', ctrl: true, shift: true, alt: false };
    expect(comboSignature(a)).toBe(comboSignature(b));
    expect(comboSignature(a)).not.toBe(comboSignature(c));
  });

  it('הצירוף עובר את `matchShortcut` — אין מסלול התאמה שני', () => {
    const shortcut = comboAsShortcut(
      { code: 'KeyK', ctrl: true, shift: false, alt: false },
      { id: 'cs-1', description: 'כותרת' },
    );

    // פריסה עברית: `key` הוא התו העברי, `code` הוא המקש הפיזי. זה הבאג שהרג
    // פעם את כל הקיצורים, וההמרה חייבת להיות חסינה לו.
    expect(
      matchShortcut(
        { key: 'ל', code: 'KeyK', ctrlKey: true, metaKey: false, shiftKey: false, altKey: false },
        shortcut,
      ),
    ).toBe(true);

    // Shift מיותר אינו מתאים: ההתאמה מדויקת בכל המודיפיירים.
    expect(
      matchShortcut(
        { key: 'ל', code: 'KeyK', ctrlKey: true, metaKey: false, shiftKey: true, altKey: false },
        shortcut,
      ),
    ).toBe(false);
  });

  it('מקש מוחזק אינו חוזר על קיצור אישי', () => {
    const shortcut = comboAsShortcut(
      { code: 'KeyK', ctrl: true, shift: false, alt: false },
      { id: 'cs-1', description: 'כותרת' },
    );
    expect(
      matchShortcut(
        {
          key: 'k',
          code: 'KeyK',
          ctrlKey: true,
          metaKey: false,
          shiftKey: false,
          altKey: false,
          repeat: true,
        },
        shortcut,
      ),
    ).toBe(false);
  });
});

describe('הרשומות השמורות', () => {
  const owners = registryOwners();

  it('הרשימה אינה ריקה, ומכסה כל רשומה שיש לה מקש', () => {
    const withKeys = (SHORTCUTS as readonly { code?: string | readonly string[] }[]).filter(
      (shortcut) => shortcut.code !== undefined,
    );
    expect(withKeys.length).toBeGreaterThan(40);
    // רשומה עם כמה מקשים פיזיים תורמת חתימה לכל אחד מהם, ולכן הרשימה גדולה
    // או שווה למספר הרשומות — ובשום מצב לא קטנה ממנו.
    expect(owners.size).toBeGreaterThanOrEqual(withKeys.length);
  });

  it('Ctrl+S ו-Ctrl+B תפוסים — הצירופים המובנים אינם זמינים להצמדה', () => {
    expect(owners.has(comboSignature({ code: 'KeyS', ctrl: true, shift: false, alt: false }))).toBe(
      true,
    );
    expect(owners.has(comboSignature({ code: 'KeyB', ctrl: true, shift: false, alt: false }))).toBe(
      true,
    );
  });

  it('החתימה של רשומה מובנית תופסת בדיוק את מה שההתאמה שלה תופסת', () => {
    // זו ההצלבה שההגנה נשענת עליה: אם החתימה נבנתה מהתווית ולא מהשדות, רשומה
    // שתוויתה אינה ניתנת לפירוק הייתה חסרה כאן — והמשתמש היה יכול להצמיד לה
    // צירוף ולהשתיק אותה בשקט.
    for (const [signature, shortcut] of owners) {
      const [ctrl, shift, alt, code] = signature.split('|');
      expect(
        matchShortcut(
          {
            key: '',
            code: code!,
            ctrlKey: ctrl === 'Ctrl',
            metaKey: false,
            shiftKey: shift === 'Shift',
            altKey: alt === 'Alt',
          },
          shortcut,
        ),
        `${shortcut.id} (${signature})`,
      ).toBe(true);
    }
  });

  it('כיווניות הפסקה מוגנת מעצם צורתה — אין דרך להצמיד צירוף למודיפייר', () => {
    // שתי הרשומות היחידות שהמקש שלהן הוא **מודיפייר** (`ShiftLeft`/
    // `ShiftRight`, מזוהות בשחרור). ההגנה עליהן אינה החתימה אלא `isBindableCombo`:
    // הלוכד אינו מסוגל לייצר צירוף שהמקש שלו הוא מודיפייר, ולכן אין מה
    // להתנגש. בלי הכלל הזה משתמש היה יכול להצמיד ל-Ctrl+Shift ולהשתיק את
    // כיווניות הפסקה — הפעולה הבסיסית ביותר במסמך עברי.
    const directional = (SHORTCUTS as readonly { id: string; code?: unknown }[]).filter(
      (shortcut) => shortcut.code === 'ShiftLeft' || shortcut.code === 'ShiftRight',
    );
    expect(directional.map((shortcut) => shortcut.id)).toEqual([
      'direction-rtl',
      'direction-ltr',
    ]);

    for (const code of ['ShiftLeft', 'ShiftRight']) {
      expect(isBindableCombo({ code, ctrl: true, shift: true, alt: false }), code).toBe(false);
    }
  });
});
