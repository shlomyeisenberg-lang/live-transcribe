/**
 * לוגיקת הנגישות של הרצועה, במנותק מהקומפוננטות.
 *
 * למה מודול נפרד: אין במאגר תשתית לבדיקות קומפוננטות, ושתי ההחלטות כאן הן
 * בדיוק מה שנוטים לשבור בלי לשים לב — הכיוון שאליו חץ מתקדם בסרגל RTL, והשאלה
 * מתי כפתור הוא מתג. כפונקציות טהורות אפשר למדוד אותן ישירות
 * (tests/unit/ribbon-aria.test.ts), והקומפוננטה נשארת חיווט.
 */

/** מזהי ה-DOM שמקשרים לשונית לפאנל. חייבים להיות משותפים לשני הצדדים. */
export const RIBBON_PANEL_ID = 'word-ribbon-panel';

export function ribbonTabId(tabId: string): string {
  return `word-ribbon-tab-${tabId}`;
}

/**
 * האינדקס שאליו ניווט המקלדת עובר, או null אם המקש אינו מקש ניווט.
 *
 * הכיווניות אינה קוסמטית: WAI-ARIA קובע שהחצים נעים לפי הכיוון **החזותי**,
 * ובסרגל RTL הלשונית הבאה נמצאת שמאלה מהפעילה — כלומר ArrowLeft מתקדם
 * ו-ArrowRight חוזר, הפוך מ-LTR. עטיפה מסוף לתחילה, כמו ברצועה של Word.
 */
export function nextTabIndex(
  key: string,
  current: number,
  count: number,
  direction: 'rtl' | 'ltr' = 'rtl',
): number | null {
  if (count <= 0) return null;

  const forward = direction === 'rtl' ? 'ArrowLeft' : 'ArrowRight';
  const back = direction === 'rtl' ? 'ArrowRight' : 'ArrowLeft';
  // לשונית פעילה שאינה ברשימה (current = -1) לא תפיל את החישוב.
  const from = ((current % count) + count) % count;

  switch (key) {
    case forward:
      return (from + 1) % count;
    case back:
      return (from - 1 + count) % count;
    case 'Home':
      return 0;
    case 'End':
      return count - 1;
    default:
      return null;
  }
}

/**
 * האם להוציא aria-pressed על כפתור הרצועה — כלומר האם הכפתור הוא מתג.
 *
 * למה לא props.active: ל-prop יש ברירת מחדל false, ולכן אחרי withDefaults
 * „מתג כבוי” ו„כפתור פעולה” נראים זהים לחלוטין מבפנים. התוצאה הייתה
 * aria-pressed="false" גם על „שמור”, „מסמך חדש” ו„הדפסה” — קורא מסך הכריז
 * אותם כמתג כבוי. `vnode.props` הוא מה שההורה **כתב בפועל**, ושם ההבחנה
 * קיימת: אתר קריאה שלא העביר active לא יופיע בו כלל. הפתרון עובד בלי לגעת
 * בלשוניות, ולכן אין אתר קריאה שצריך לשנות.
 */
export function isToggleButton(rawProps: Record<string, unknown> | null | undefined): boolean {
  if (!rawProps) return false;
  return Object.prototype.hasOwnProperty.call(rawProps, 'active');
}
