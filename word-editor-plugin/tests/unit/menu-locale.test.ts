/**
 * שפת התפריטים כסימן שגם CSS יכול לקרוא.
 *
 * `menuString` לבדו הספיק כל עוד כל המחרוזות עברו דרך תבנית של Vue. שכבת
 * הכותרות שהמנוע מצייר שברה את ההנחה: ארבע תוויות שם מתורגמות ב-CSS ולא
 * ב-JS (styles/engine-chrome.css מסביר למה), ו-CSS אינו יכול לקרוא `ref`.
 * לכן `setMenuLocale` כותב את השפה גם על שורש ה-HTML.
 *
 * מה שנמדד כאן הוא בדיוק החוזה שהגיליון נשען עליו: מתי התכונה קיימת, מה
 * ערכה, ומה קורה לפני שהאתחול קבע שפה. בדיקת החוזה
 * (tests/contract/engine-hf-chrome.test.ts) מאמתת מהצד השני שכל כלל בגיליון
 * אכן מסתייג בתכונה הזאת.
 */
import { describe, it, expect, afterEach } from 'vitest';
import { MENU_LOCALE_ATTRIBUTE, menuLocale, menuString, setMenuLocale } from '../../src/ui/ribbon/i18n';

afterEach(() => {
  setMenuLocale('he');
  document.documentElement.removeAttribute(MENU_LOCALE_ATTRIBUTE);
});

describe('סימן השפה על שורש ה-HTML', () => {
  it('היעדר התכונה הוא עברית — לפני האתחול הכללים בגיליון חלים', () => {
    // `:root:not([data-menu-locale='en'])` מתאים גם כשאין תכונה בכלל, וזו
    // ההתנהגות הרצויה: עברית היא שפת המקור של התוסף.
    expect(document.documentElement.hasAttribute(MENU_LOCALE_ATTRIBUTE)).toBe(false);
    expect(menuLocale()).toBe('he');
  });

  it('אנגלית נכתבת כתכונה, ולא רק ל-ref', () => {
    setMenuLocale('en');

    expect(document.documentElement.getAttribute(MENU_LOCALE_ATTRIBUTE)).toBe('en');
    expect(menuLocale()).toBe('en');
    expect(menuString('בית')).toBe('Home');
  });

  it('חזרה לעברית מוחקת את השער — התכונה אינה נדבקת', () => {
    setMenuLocale('en');
    setMenuLocale('he');

    expect(document.documentElement.getAttribute(MENU_LOCALE_ATTRIBUTE)).toBe('he');
    expect(menuString('בית')).toBe('בית');
  });

  it('תג מלא ושפה שלא דווחה הם עברית', () => {
    setMenuLocale('he-IL');
    expect(document.documentElement.getAttribute(MENU_LOCALE_ATTRIBUTE)).toBe('he');

    setMenuLocale(undefined);
    expect(document.documentElement.getAttribute(MENU_LOCALE_ATTRIBUTE)).toBe('he');

    setMenuLocale('en-US');
    expect(document.documentElement.getAttribute(MENU_LOCALE_ATTRIBUTE)).toBe('en');
  });
});
