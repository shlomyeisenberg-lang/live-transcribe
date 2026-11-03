/**
 * החשבון של גלילת הגלגלת האופקית.
 *
 * המספרים כאן נמדדו בכרום על הרצועה עצמה (חלון 900px, לשונית „בית”):
 * `scrollWidth` 1227, `clientWidth` 900, ולכן טווח `scrollLeft` של
 * `[-327, 0]` — שלילי, מפני שהמכולה `dir="rtl"`. הבדיקה נשענת עליהם ולא על
 * מספרים עגולים, כי הסימן הוא כל מה שיש כאן לטעות בו.
 *
 * בדיקת הרכיב (tests/component/wheel-scroll.test.ts) מודדת את החיווט; היא רצה
 * ב-jsdom, שאין בו פריסה, ולכן אין בכוחה לפסול שום נוסחה.
 */
import { describe, expect, it } from 'vitest';
import { wheelScrollDelta, type WheelScrollMetrics } from '../../src/composables/wheel-scroll';

/** הרצועה כפי שנמדדה, בתחילת הגלילה. */
const RIBBON: WheelScrollMetrics = { scrollLeft: 0, scrollWidth: 1227, clientWidth: 900 };
/** אותה רצועה בסופה: `-(1227 - 900)`. */
const RIBBON_END: WheelScrollMetrics = { ...RIBBON, scrollLeft: -327 };

/** גלגול רגיל: מטה חיובי, מעלה שלילי, בלי מקש מלווה. */
const down = { deltaX: 0, deltaY: 100, ctrlKey: false };
const up = { deltaX: 0, deltaY: -100, ctrlKey: false };

describe('wheelScrollDelta', () => {
  it('גלגלת מטה ב-RTL מתקדמת שמאלה — אל המשך הרצועה', () => {
    expect(wheelScrollDelta(down, RIBBON, true)).toBe(-100);
  });

  it('גלגלת מעלה ב-RTL חוזרת ימינה', () => {
    expect(wheelScrollDelta(up, { ...RIBBON, scrollLeft: -200 }, true)).toBe(100);
  });

  it('ב-LTR אותה מחווה הפוכה בסימן', () => {
    expect(wheelScrollDelta(down, { ...RIBBON, scrollLeft: 0 }, false)).toBe(100);
    expect(wheelScrollDelta(up, { ...RIBBON, scrollLeft: 200 }, false)).toBe(-100);
  });

  /*
   * שני הקצוות מחזירים `null` ולא 0: ההבדל אינו סגנוני — `null` הוא „לא
   * טיפלתי”, והוא מה שמשאיר את האירוע לגלול את המכולה שמעל. גלריית הסגנונות
   * יושבת בתוך הרצועה, וזו השרשרת שמאפשרת לה להמשיך אליה.
   */
  it('בסוף הרצועה אין להמשיך קדימה', () => {
    expect(wheelScrollDelta(down, RIBBON_END, true)).toBeNull();
    expect(wheelScrollDelta(up, RIBBON_END, true)).toBe(100);
  });

  it('בתחילת הרצועה אין לחזור אחורה', () => {
    expect(wheelScrollDelta(up, RIBBON, true)).toBeNull();
    expect(wheelScrollDelta(down, RIBBON, true)).toBe(-100);
  });

  it('קצה נמדד בסבילות תת-פיקסלית', () => {
    expect(wheelScrollDelta(up, { ...RIBBON, scrollLeft: -0.4 }, true)).toBeNull();
    expect(wheelScrollDelta(down, { ...RIBBON, scrollLeft: -326.6 }, true)).toBeNull();
  });

  it('רצועה שנכנסת כולה אינה נוגעת בגלגלת', () => {
    const fits = { scrollLeft: 0, scrollWidth: 900, clientWidth: 900 };
    expect(wheelScrollDelta(down, fits, true)).toBeNull();
    expect(wheelScrollDelta(up, fits, true)).toBeNull();
  });

  /* Ctrl+גלגלת הוא זום, ולקיחתו כאן הייתה מבטלת אותו על כל שטח הרצועה. */
  it('Ctrl+גלגלת נשאר של הדפדפן', () => {
    expect(wheelScrollDelta({ ...down, ctrlKey: true }, RIBBON, true)).toBeNull();
  });

  /*
   * מחווה אופקית אמיתית — משטח מגע, או Shift+גלגלת שכרום ממיר לציר X — כבר
   * נגללת מאליה. תרגום שלה כאן היה מכפיל אותה.
   */
  it('מחווה שהציר האופקי שלה שולט עוברת הלאה', () => {
    expect(wheelScrollDelta({ deltaX: -120, deltaY: 0, ctrlKey: false }, RIBBON, true)).toBeNull();
    expect(wheelScrollDelta({ deltaX: -120, deltaY: 40, ctrlKey: false }, RIBBON, true)).toBeNull();
    // אלכסון שהאנכי שלו שולט הוא עדיין גלגול גלגלת, ונלקח.
    expect(wheelScrollDelta({ deltaX: -40, deltaY: 120, ctrlKey: false }, RIBBON, true)).toBe(-120);
  });

  it('גלגול ריק או לא-מספרי אינו גלילה', () => {
    expect(wheelScrollDelta({ deltaX: 0, deltaY: 0, ctrlKey: false }, RIBBON, true)).toBeNull();
    expect(wheelScrollDelta({ deltaX: 0, deltaY: NaN, ctrlKey: false }, RIBBON, true)).toBeNull();
  });

  /* המכולה מדווחת NaN רק כשהיא מנותקת מהמסמך — אבל אז `scrollLeft +=` היה
     כותב NaN, וזה מצב שאין ממנו חזרה בלי גלילה ידנית. */
  it('מידות שאינן מספר אינן מזיזות דבר', () => {
    const broken = { scrollLeft: 0, scrollWidth: NaN, clientWidth: 900 };
    expect(wheelScrollDelta(down, broken, true)).toBeNull();
  });
});
