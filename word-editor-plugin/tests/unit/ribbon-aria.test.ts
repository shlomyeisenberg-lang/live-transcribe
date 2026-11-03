/**
 * הנגישות של הרצועה — הלוגיקה, לא ה-markup.
 *
 * שני דברים נמדדים כאן:
 *   1. ניווט החצים בין הלשוניות. בסרגל RTL הלשונית הבאה נמצאת שמאלה מהפעילה,
 *      ולכן ArrowLeft מתקדם — ההיפוך הזה הוא מה שנשבר בלי שרואים.
 *   2. המכניקה שמאפשרת ל-RibbonButton לדעת אם הוא מתג: prop עם ברירת מחדל
 *      false אינו מבדיל בין „מתג כבוי” ל„כפתור פעולה”, ו-`vnode.props` כן.
 *
 * מה שהיה כאן ואינו כאן יותר: קומפוננטה שקולה שנבנתה מ-vue עצמו כדי למדוד את
 * התנהגות Vue שהפתרון נשען עליה — תחליף שנדרש כל עוד לא היה במאגר במה להרכיב
 * קומפוננטה. היא הוחלפה במדידה על `RibbonButton` האמיתי בתוך הרצועה
 * (tests/component/ribbon-shell.test.ts): „שמור” ו„מסמך חדש” יוצאים בלי
 * aria-pressed, ומתג מדווח את מצבו ומתעדכן כשהמנוע מזיז אותו.
 */
import { describe, expect, it } from 'vitest';
import { RIBBON_PANEL_ID, isToggleButton, nextTabIndex, ribbonTabId } from '../../src/ui/ribbon/aria';

describe('nextTabIndex', () => {
  it('ב-RTL ArrowLeft מתקדם ו-ArrowRight חוזר', () => {
    expect(nextTabIndex('ArrowLeft', 3, 8)).toBe(4);
    expect(nextTabIndex('ArrowRight', 3, 8)).toBe(2);
  });

  it('ב-LTR הכיוונים מתהפכים', () => {
    expect(nextTabIndex('ArrowRight', 3, 8, 'ltr')).toBe(4);
    expect(nextTabIndex('ArrowLeft', 3, 8, 'ltr')).toBe(2);
  });

  it('עוטף משני הקצוות', () => {
    expect(nextTabIndex('ArrowLeft', 7, 8)).toBe(0);
    expect(nextTabIndex('ArrowRight', 0, 8)).toBe(7);
  });

  it('Home ו-End קופצים לקצוות', () => {
    expect(nextTabIndex('Home', 5, 8)).toBe(0);
    expect(nextTabIndex('End', 5, 8)).toBe(7);
  });

  it('מקש שאינו ניווט מחזיר null, כדי שהאירוע לא ייחטף', () => {
    for (const key of ['ArrowUp', 'ArrowDown', 'Tab', 'Enter', ' ', 'a', 'Escape']) {
      expect(nextTabIndex(key, 3, 8)).toBeNull();
    }
  });

  it('אינו מתפוצץ על סרגל ריק או על לשונית שאינה ברשימה', () => {
    expect(nextTabIndex('ArrowLeft', 0, 0)).toBeNull();
    expect(nextTabIndex('ArrowLeft', -1, 8)).toBe(0);
    expect(nextTabIndex('ArrowRight', -1, 8)).toBe(6);
  });
});

describe('isToggleButton', () => {
  it('מזהה כפתור שההורה קשר לו active', () => {
    expect(isToggleButton({ active: false })).toBe(true);
    expect(isToggleButton({ active: true })).toBe(true);
  });

  it('כפתור פעולה — שההורה העביר לו רק תווית ואייקון — אינו מתג', () => {
    expect(isToggleButton({ label: 'שמור', icon: 'save', variant: 'large' })).toBe(false);
    expect(isToggleButton({})).toBe(false);
    expect(isToggleButton(null)).toBe(false);
    expect(isToggleButton(undefined)).toBe(false);
  });
});

describe('מזהי ה-DOM שמקשרים לשונית לפאנל', () => {
  it('יציבים ומבוססים על מזהה הלשונית', () => {
    expect(ribbonTabId('home')).toBe('word-ribbon-tab-home');
    expect(RIBBON_PANEL_ID).toBe('word-ribbon-panel');
  });
});
