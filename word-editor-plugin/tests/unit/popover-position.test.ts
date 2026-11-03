/**
 * מיקום הפופאוברים של הרצועה.
 *
 * הבאג שהוליד את המודול: שלושת הפופאוברים היו `position: absolute; top: 100%`
 * בתוך `.word-ribbon-body`, שמוגדר `overflow-x: auto; overflow-y: hidden` —
 * כלומר ההורה חותך אנכית בגובה הרצועה, ומפלטת הצבעים נראתה בפועל רק השורה
 * הראשונה. התיקון הוא `position: fixed` בקואורדינטות שנמדדות, וזה מה שנבדק
 * כאן.
 *
 * למה פונקציה טהורה ולא רק בדיקת קומפוננטה: jsdom אינו מודד פריסה ומחזיר אפס
 * מכל `getBoundingClientRect`, ולכן ההיפוך למעלה, ההצמדה לקצה החלון והגבלת
 * הגובה אינם ניתנים למדידה שם בכלל. הפרדת החישוב היא מה שהופך אותם לבדיקים.
 */
import { describe, expect, it } from 'vitest';
import {
  POPOVER_GAP_PX,
  POPOVER_MARGIN_PX,
  POPOVER_MIN_HEIGHT_PX,
  popoverPlacement,
  type AnchorRect,
} from '../../src/composables/popover-position';

/** כפתור ברצועה: 24px גובה, בערך ברוחב של בורר צבע. */
function anchorAt(top: number, right: number): AnchorRect {
  return { top, bottom: top + 24, left: right - 40, right };
}

const PALETTE = { width: 200, height: 150 };
const DESKTOP = { width: 1000, height: 800 };

describe('popoverPlacement', () => {
  it('כשיש מקום מתחת לכפתור — הפופאובר נפתח למטה, במרווח של גדלים', () => {
    const placement = popoverPlacement(anchorAt(80, 960), PALETTE, DESKTOP, { rtl: true });

    expect(placement.side).toBe('below');
    expect(placement.top).toBe(80 + 24 + POPOVER_GAP_PX);
  });

  it('כשאין מקום מתחת ויש מעל — הפופאובר מתהפך למעלה', () => {
    // זה המצב שהמשתמש ראה: כפתור בשליש התחתון וחלון נמוך. קודם הפופאובר נחתך;
    // עכשיו הוא נפתח כלפי מעלה, כמו ב-Word.
    const placement = popoverPlacement(anchorAt(400, 960), PALETTE, { width: 1000, height: 500 }, {
      rtl: true,
    });

    expect(placement.side).toBe('above');
    // הקצה התחתון של הפופאובר נוגע בקצה העליון של הכפתור, פחות המרווח.
    expect(placement.top + PALETTE.height + POPOVER_GAP_PX).toBe(400);
  });

  it('כששני הצדדים צרים — נבחר הרחב, והגובה מוגבל במקום שנחתך', () => {
    // ההבדל בין „גולל בתוך עצמו” ל„נחתך”: `maxHeight` הוא מה שהופך את הראשון
    // לאפשרי, ובלעדיו הבחירה בצד הרחב לבדה עוד הייתה משאירה תוכן בלתי נגיש.
    const placement = popoverPlacement(anchorAt(100, 960), { width: 200, height: 400 }, {
      width: 1000,
      height: 300,
    });

    expect(placement.side).toBe('below');
    expect(placement.maxHeight).toBe(300 - 124 - POPOVER_GAP_PX - POPOVER_MARGIN_PX);
    expect(placement.maxHeight).toBeLessThan(400);
  });

  it('חלון שאין בו מקום בכלל — גובה מינימלי, והפופאובר נשאר בתוך החלון', () => {
    // חלון בגובה 120px: 36px מתחת לכפתור ו-40px מעליו. בלי רצפת גובה זה היה
    // `max-height: 40px` — פופאובר פתוח שאינו מציג דבר, כלומר הבאג המקורי שוב.
    const viewport = { width: 1000, height: 120 };
    const placement = popoverPlacement(anchorAt(50, 960), { width: 200, height: 400 }, viewport);

    expect(placement.maxHeight).toBe(POPOVER_MIN_HEIGHT_PX);
    expect(placement.top).toBeGreaterThanOrEqual(POPOVER_MARGIN_PX);
    expect(placement.top + POPOVER_MIN_HEIGHT_PX).toBeLessThanOrEqual(
      viewport.height - POPOVER_MARGIN_PX,
    );
  });

  it('בעברית הפופאובר מיושר לקצה הימני של הכפתור, ובלטינית לשמאלי', () => {
    // כפתור באמצע החלון, כדי שההצמדה לקצה לא תסתיר את מה שנמדד כאן.
    const anchor = anchorAt(80, 500);

    expect(popoverPlacement(anchor, PALETTE, DESKTOP, { rtl: true }).left).toBe(
      anchor.right - PALETTE.width,
    );
    expect(popoverPlacement(anchor, PALETTE, DESKTOP, { rtl: false }).left).toBe(anchor.left);
  });

  it('כפתור בקצה החלון — ההצמדה לכפתור נכנעת לחלון', () => {
    // בלי ההצמדה הזאת פופאובר רחב מהכפתור היה יוצא מהחלון, וזה אותו באג בציר
    // האחר בדיוק.
    const nearStart = popoverPlacement(anchorAt(80, 60), PALETTE, DESKTOP, { rtl: true });
    expect(nearStart.left).toBe(POPOVER_MARGIN_PX);

    const nearEnd = popoverPlacement(anchorAt(80, 990), PALETTE, DESKTOP, { rtl: false });
    expect(nearEnd.left).toBe(DESKTOP.width - POPOVER_MARGIN_PX - PALETTE.width);
  });

  it('פופאובר רחב מהחלון נצמד לקצה ההתחלה ואינו נדחק לשלילי', () => {
    const placement = popoverPlacement(anchorAt(80, 500), { width: 1200, height: 100 }, DESKTOP, {
      rtl: true,
    });

    expect(placement.left).toBe(POPOVER_MARGIN_PX);
  });

  /**
   * היישור למרכז נוסף בעבור הטולטיפ (ui/tooltip/TooltipLayer.vue): כרטיס ההסבר
   * אינו „נפתח מ”הכפתור אלא מסביר אותו, ולעתים צר ממנו — יישור לקצה היה מסיט
   * אותו מהאייקון.
   */
  describe('align: center', () => {
    const TIP = { width: 240, height: 60 };

    it('מרכז הפופאובר נופל על מרכז הכפתור, בשני הכיוונים', () => {
      const anchor = anchorAt(80, 500);
      const middle = (anchor.left + anchor.right) / 2;

      for (const rtl of [true, false]) {
        const placement = popoverPlacement(anchor, TIP, DESKTOP, { align: 'center', rtl });
        expect(placement.left + TIP.width / 2).toBe(middle);
      }
    });

    it('כפתור בקצה — המרכוז נכנע לחלון, כמו ההצמדה לקצה', () => {
      const nearEnd = popoverPlacement(anchorAt(80, 995), TIP, DESKTOP, { align: 'center' });
      expect(nearEnd.left).toBe(DESKTOP.width - POPOVER_MARGIN_PX - TIP.width);

      const nearStart = popoverPlacement(anchorAt(80, 40), TIP, DESKTOP, { align: 'center' });
      expect(nearStart.left).toBe(POPOVER_MARGIN_PX);
    });

    it('ההיפוך האנכי אינו תלוי ביישור האופקי', () => {
      const placement = popoverPlacement(anchorAt(760, 500), TIP, DESKTOP, {
        align: 'center',
        gap: 8,
      });

      expect(placement.side).toBe('above');
      expect(placement.top).toBe(760 - 8 - TIP.height);
    });
  });
});
