/**
 * מרכוז העמוד בזום: המשתנים שמפורסמים ל-CSS, וההאזנה שמחזיקה אותם עדכניים.
 *
 * הגזירה והמדידות בהערת הפתיחה של engine/zoom-center.ts. כאן נמדדות שלוש
 * ההבטחות שאפשר למדוד בלי דפדפן: הערכים עצמם, הנפילה בחן על דיווח פגום, ומי
 * מפעיל חישוב מחדש — הזום, שינוי הגודל, והחלפת מיכל הגלילה בין מסמכים.
 * הגיאומטריה עצמה נמדדת חי ב-scripts/zoom-center-probe.mjs.
 */
import { describe, expect, it, vi } from 'vitest';
import {
  DOCUMENT_DIRECTION_ATTRIBUTE,
  HOST_SELECTOR,
  VIEWPORT_WIDTH_VAR,
  ZOOM_INVERSE_VAR,
  ZOOM_PAGE_CLASS,
  ZOOM_VAR,
  createZoomCenter,
  zoomCenterVars,
} from '../../src/engine/zoom-center';

describe('zoomCenterVars', () => {
  it('אחוז הופך ליחס, ולצידו ההופכי — כדי ש-CSS לא יחלק במשתנה', () => {
    expect(zoomCenterVars(150, 1425)).toEqual({
      [ZOOM_VAR]: '1.5',
      [ZOOM_INVERSE_VAR]: '0.6667',
      [VIEWPORT_WIDTH_VAR]: '1425px',
    });
  });

  it('100% הוא הזהות: הכלל אינו מזיז דבר', () => {
    const vars = zoomCenterVars(100, 1425);
    expect(vars[ZOOM_VAR]).toBe('1');
    expect(vars[ZOOM_INVERSE_VAR]).toBe('1');
  });

  it('רוחב המאגס מעוגל לפיקסל — משתנה CSS אינו זקוק ל-17 ספרות', () => {
    expect(zoomCenterVars(60, 1424.6)[VIEWPORT_WIDTH_VAR]).toBe('1425px');
  });

  it('אחוז שאינו מספר חיובי נופל ל-100% ולא מייצר NaN בגיליון', () => {
    for (const bad of [0, -50, Number.NaN, Number.POSITIVE_INFINITY]) {
      expect(zoomCenterVars(bad, 1425)[ZOOM_VAR]).toBe('1');
    }
  });

  it('בלי מדידת רוחב — 100vw, כלומר מרכוז מלא ולא הצמדה לקצה', () => {
    // רחב לפחות כמו המאגס, ולכן ענף ה-min בוחר תמיד את המרכוז.
    for (const bad of [0, -1, Number.NaN]) {
      expect(zoomCenterVars(200, bad)[VIEWPORT_WIDTH_VAR]).toBe('100vw');
    }
  });
});

describe('createZoomCenter', () => {
  /** `clientWidth` הוא getter על האב-טיפוס, ולכן נקבע בהגדרת מאפיין. */
  function withWidth<T extends Element>(element: T, width: number): T {
    Object.defineProperty(element, 'clientWidth', { value: width, configurable: true });
    return element;
  }

  /** מאגס מזויף עם מיכל גלילה בפנים — הצורה שהפקד מצפה לה. */
  function stackWith(hostWidth: number | null, stackWidth = 1440) {
    const setProperty = vi.fn();
    const host =
      hostWidth === null ? null : withWidth(document.createElement('div'), hostWidth);

    return {
      setProperty,
      host,
      stack: {
        clientWidth: stackWidth,
        querySelector: (selector: string) => (selector === HOST_SELECTOR ? host : null),
        style: { setProperty },
      },
    };
  }

  function published(setProperty: ReturnType<typeof vi.fn>): Record<string, string> {
    return Object.fromEntries(setProperty.mock.calls);
  }

  it('מפרסם את שלושת המשתנים לפי מיכל הגלילה ולא לפי ה-<main>', () => {
    const { stack, setProperty } = stackWith(1425);
    createZoomCenter(stack as never).setZoom(200);

    expect(published(setProperty)).toEqual({
      [ZOOM_VAR]: '2',
      [ZOOM_INVERSE_VAR]: '0.5',
      [VIEWPORT_WIDTH_VAR]: '1425px',
    });
  });

  it('בלי מיכל גלילה נופל ל-<main>, ולא לאפס', () => {
    const { stack, setProperty } = stackWith(null, 1440);
    createZoomCenter(stack as never);

    expect(published(setProperty)[VIEWPORT_WIDTH_VAR]).toBe('1440px');
  });

  it('מפרסם כבר בהקמה — הדיווח הראשון של observeZoom לא מגיע לפניו', () => {
    const { stack, setProperty } = stackWith(1425);
    createZoomCenter(stack as never);

    expect(published(setProperty)[ZOOM_VAR]).toBe('1');
  });

  it('שינוי גודל מחשב מחדש עם הזום האחרון', () => {
    const { stack, setProperty } = stackWith(1425);
    const center = createZoomCenter(stack as never);
    center.setZoom(150);

    stack.clientWidth = 900;
    withWidth(stack.querySelector(HOST_SELECTOR) as Element, 885);
    center.refresh();

    expect(published(setProperty)).toEqual({
      [ZOOM_VAR]: '1.5',
      [ZOOM_INVERSE_VAR]: '0.6667',
      [VIEWPORT_WIDTH_VAR]: '885px',
    });
  });

  /**
   * מיכל שגולש אופקית. `scrollWidth` הוא getter על האב-טיפוס, כמו
   * `clientWidth`, ולכן נקבע בהגדרת מאפיין; `scrollLeft` הוא מאפיין רגיל
   * ב-jsdom ולכן ההשמה עליו נמדדת ישירות.
   */
  function overflowing(host: Element, content: number): Element {
    Object.defineProperty(host, 'scrollWidth', { value: content, configurable: true });
    return host;
  }

  it('שינוי זום מחזיר את הגלילה לתחילת השורה במסמך RTL', () => {
    // מיכל הגלילה הוא ltr (הצהרה על צד פס הגלילה, ראו shell.css), ולכן הוא
    // נח ב-scrollLeft=0 — כלומר בסופי השורות העבריות. הנוסחה שב-CSS מבטיחה
    // שכל העמוד נגיש; מה שמחזיר את הגלילה לתחילת השורה הוא כאן.
    document.documentElement.setAttribute(DOCUMENT_DIRECTION_ATTRIBUTE, 'rtl');
    const { stack } = stackWith(1425);
    const host = overflowing(stack.querySelector(HOST_SELECTOR) as Element, 2400);

    createZoomCenter(stack as never).setZoom(300);

    expect(host.scrollLeft).toBe(2400 - 1425);
    document.documentElement.removeAttribute(DOCUMENT_DIRECTION_ATTRIBUTE);
  });

  it('במסמך שאינו RTL הגלילה אינה נגעת — שם 0 הוא תחילת השורה', () => {
    const { stack } = stackWith(1425);
    const host = overflowing(stack.querySelector(HOST_SELECTOR) as Element, 2400);

    createZoomCenter(stack as never).setZoom(300);

    expect(host.scrollLeft).toBe(0);
  });

  it('מיכל שאינו גולש אופקית אינו מקבל השמה', () => {
    document.documentElement.setAttribute(DOCUMENT_DIRECTION_ATTRIBUTE, 'rtl');
    const { stack } = stackWith(1425);
    const host = overflowing(stack.querySelector(HOST_SELECTOR) as Element, 1200);

    createZoomCenter(stack as never).setZoom(60);

    expect(host.scrollLeft).toBe(0);
    document.documentElement.removeAttribute(DOCUMENT_DIRECTION_ATTRIBUTE);
  });

  it('שינוי גודל אינו מזיז את הגלילה — גרירת חלון אינה בקשה לזוז', () => {
    // ההבחנה הזאת היא כל הטעם בכך שההצמדה יושבת ב-`setZoom` ולא ב-`apply`:
    // משתמש שגלל לאמצע השורה וגרר את החלון אינו מבקש לחזור להתחלה.
    document.documentElement.setAttribute(DOCUMENT_DIRECTION_ATTRIBUTE, 'rtl');
    const { stack } = stackWith(1425);
    const host = overflowing(stack.querySelector(HOST_SELECTOR) as Element, 2400);
    const center = createZoomCenter(stack as never);
    center.setZoom(300);
    host.scrollLeft = 500;

    center.refresh();

    expect(host.scrollLeft).toBe(500);
    document.documentElement.removeAttribute(DOCUMENT_DIRECTION_ATTRIBUTE);
  });

  it('dispose מנתק את ההאזנה', () => {
    const disconnect = vi.fn();
    const observe = vi.fn();
    vi.stubGlobal(
      'ResizeObserver',
      class {
        observe = observe;
        unobserve = vi.fn();
        disconnect = disconnect;
      },
    );

    const { stack } = stackWith(1425);
    createZoomCenter(stack as never).dispose();

    expect(disconnect).toHaveBeenCalledTimes(1);
    vi.unstubAllGlobals();
  });
});

describe('שם המחלקה של העמוד', () => {
  it('הוא מה שהסלקטור ב-shell.css מכוון אליו', () => {
    // הקבוע קיים כדי ששער ההיגיינה יוכיח שהמחלקה אינה שם שנכתב בטעות.
    expect(ZOOM_PAGE_CLASS).toBe('superdoc-page');
  });
});
