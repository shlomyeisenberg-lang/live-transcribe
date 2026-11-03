/**
 * הגאומטריה של גרירת הדיאלוגים.
 *
 * כל מה שנשאלת כאן הוא השאלה שאין לה קיום ב-jsdom: לאן הדיאלוג נוחת כשגוררים
 * אותו, וכמה החלון מרשה. בדיקת הרכיב (tests/component/dialog-drag.test.ts)
 * מודדת רק שהחיווט קיים — היא רצה על מלבנים של אפס, ולכן הייתה מאשרת בירוק
 * גם נוסחה שמתעלמת מקצוות החלון לחלוטין.
 *
 * שתי המערכות (`origin` מול `rect`) אינן קפריזה: `FindReplaceDialog` הוא
 * `position: absolute` והשאר `fixed`, ולכן „מה שה-CSS כותב” ו„היכן זה נחת
 * בחלון” אינם אותו מספר.
 */
import { describe, expect, it } from 'vitest';
import { dialogDragPosition } from '../../src/composables/dialog-drag';

const VIEWPORT = { width: 1280, height: 800 };
/** דיאלוג במידות אמיתיות: `.fontadv-dialog` הוא 560 רוחב (שתי עמודות). */
const DIALOG = { left: 40, top: 140, width: 560, height: 420 };

describe('dialogDragPosition', () => {
  it('גרירה בתוך החלון מזיזה בדיוק כמו המצביע', () => {
    const at = dialogDragPosition({ left: 40, top: 140 }, DIALOG, { x: 200, y: 60 }, VIEWPORT);
    expect(at).toEqual({ left: 240, top: 200 });
  });

  it('הקצה התחתון אינו נחצה — שם יושבים „אישור” ו„ביטול”', () => {
    const at = dialogDragPosition({ left: 40, top: 140 }, DIALOG, { x: 0, y: 900 }, VIEWPORT);
    expect(at.top).toBe(VIEWPORT.height - DIALOG.height);
    expect(at.top + DIALOG.height).toBeLessThanOrEqual(VIEWPORT.height);
  });

  it('הקצה העליון אינו נחצה — הכותרת היא הידית היחידה', () => {
    const at = dialogDragPosition({ left: 40, top: 140 }, DIALOG, { x: 0, y: -900 }, VIEWPORT);
    expect(at.top).toBe(0);
  });

  it('שני הקצוות האופקיים עוצרים את הדיאלוג בשלמותו', () => {
    const right = dialogDragPosition({ left: 40, top: 140 }, DIALOG, { x: 5000, y: 0 }, VIEWPORT);
    expect(right.left).toBe(VIEWPORT.width - DIALOG.width);

    const left = dialogDragPosition({ left: 40, top: 140 }, DIALOG, { x: -5000, y: 0 }, VIEWPORT);
    expect(left.left).toBe(0);
  });

  /**
   * חלון נמוך מהדיאלוג: אין מיקום שמכיל אותו, וההכרעה היא לטובת הראש. הצמדה
   * לקצה התחתון הייתה דוחפת את הכותרת מעל המסך — כלומר דיאלוג שאי אפשר לגרור
   * בחזרה ואי אפשר לסגור בעכבר.
   */
  it('דיאלוג גבוה מהחלון נצמד לראש ולא לתחתית', () => {
    const at = dialogDragPosition({ left: 40, top: 140 }, DIALOG, { x: 0, y: 300 }, { width: 1280, height: 300 });
    expect(at.top).toBe(0);
  });

  /**
   * `position: absolute` בתוך מכל שאינו בראש החלון: המלבן בקואורדינטות חלון,
   * וה-CSS נכתב ביחס למכל. ההפרש בין השתיים חייב לשרוד את ההצמדה, אחרת
   * הדיאלוג היה קופץ בגובה הפס העליון ברגע שנוגעים בו.
   */
  it('הפרש הקואורדינטות נשמר כשה-CSS אינו נכתב בקואורדינטות חלון', () => {
    const rect = { left: 40, top: 188, width: 560, height: 300 };
    const at = dialogDragPosition({ left: 40, top: 140 }, rect, { x: 10, y: 20 }, VIEWPORT);
    expect(at).toEqual({ left: 50, top: 160 });
  });

  it('גרירה בלי תזוזה אינה משנה דבר', () => {
    const at = dialogDragPosition({ left: 40, top: 140 }, DIALOG, { x: 0, y: 0 }, VIEWPORT);
    expect(at).toEqual({ left: 40, top: 140 });
  });
});
