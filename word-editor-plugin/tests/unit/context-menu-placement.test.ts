/**
 * לאן נפתח תפריט ההקשר.
 *
 * הבדיקה המרכזית כאן היא ההיפוך האופקי, והיא נכתבה מהתקלה עצמה: עוגן-נקודה
 * ב-RTL מבקש `left = x - width`, ולחיצה בשוליים השמאליים נותנת מספר שלילי.
 * ההצמדה של `popoverPlacement` מרימה אותו לשוליים — ואז הכרטיס נפתח **מתחת
 * לסמן**, כלומר מכסה בדיוק את מה שנלחץ עליו. בעברית זה סוף כל שורה, לא מקרה
 * קצה.
 */
import { describe, it, expect } from 'vitest';
import { contextMenuPlacement } from '../../src/ui/menu/menu-placement';

const SIZE = { width: 264, height: 300 };
const VIEWPORT = { width: 1200, height: 800 };

describe('contextMenuPlacement', () => {
  it('בעברית הכרטיס נפתח משמאל לנקודה — קצהו הימני נוגע בה', () => {
    const placement = contextMenuPlacement({ x: 900, y: 200 }, SIZE, VIEWPORT);

    expect(placement.left).toBe(900 - SIZE.width);
    expect(placement.top).toBe(200);
    expect(placement.side).toBe('below');
  });

  it('בקצה השמאלי הוא מתהפך ונפתח ימינה, ולא נצמד מתחת לסמן', () => {
    const placement = contextMenuPlacement({ x: 20, y: 200 }, SIZE, VIEWPORT);

    expect(placement.left).toBe(20);
  });

  it('בכיוון שמאל-לימין הכיוונים הפוכים', () => {
    const right = contextMenuPlacement({ x: 100, y: 200 }, SIZE, VIEWPORT, { rtl: false });
    expect(right.left).toBe(100);

    const flipped = contextMenuPlacement({ x: 1190, y: 200 }, SIZE, VIEWPORT, { rtl: false });
    expect(flipped.left).toBe(1190 - SIZE.width);
  });

  /**
   * כרטיס רחב מהחלון אינו יכול „להיכנס”, ולכן מה שנדרש ממנו הוא להיצמד לשוליים
   * ולא לברוח שמאלה. הניסוח הראשון של הבדיקה הזאת טען „אינו יוצא מהחלון”
   * ואישר `left = 8` עם רוחב 400 בחלון 320 — כלומר 88 פיקסלים בחוץ.
   */
  it('כרטיס רחב מהחלון נצמד לשוליים ולא נדחף אל מחוץ להם', () => {
    const placement = contextMenuPlacement({ x: 150, y: 200 }, { width: 400, height: 200 }, { width: 320, height: 800 });

    expect(placement.left).toBe(8);
  });

  it('בקצה התחתון הוא נפתח כלפי מעלה', () => {
    const placement = contextMenuPlacement({ x: 900, y: 780 }, SIZE, VIEWPORT);

    expect(placement.side).toBe('above');
    expect(placement.top + SIZE.height).toBeLessThanOrEqual(780);
  });

  it('בחלון נמוך הגובה מוגבל במקום שהכרטיס ייחתך', () => {
    const placement = contextMenuPlacement({ x: 900, y: 100 }, SIZE, { width: 1200, height: 220 });

    expect(placement.maxHeight).toBeLessThan(SIZE.height);
    expect(placement.maxHeight).toBeGreaterThan(0);
  });

  it('אין מרווח בין הסמן לפינת הכרטיס', () => {
    const placement = contextMenuPlacement({ x: 600, y: 300 }, SIZE, VIEWPORT);

    expect(placement.top).toBe(300);
  });
});
