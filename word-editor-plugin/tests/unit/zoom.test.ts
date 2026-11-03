/**
 * גודל התצוגה.
 *
 * מה שהיה שבור: הסרגל ב-StatusBar.vue קידד `min="50" max="200"` בתבנית, ואיש
 * לא שאל את המנוע. הגבולות האמיתיים יושבים ב-`ui.zoom.getSnapshot()`, ומסמך
 * שהמנוע מגביל אחרת היה מקבל סרגל שנע לערכים שהמנוע דוחה — הסרגל זז והמסמך
 * לא. הבדיקות כאן מקבעות שהגבולות מגיעים מהמנוע, ושדיווח פגום שלו נופל בחן
 * לברירת מחדל שאפשר לעבוד איתה ולא משתיק את הסרגל.
 */
import { describe, it, expect, vi } from 'vitest';
import {
  FALLBACK_ZOOM,
  ZOOM_PERCENT_MAX,
  clampZoom,
  normalizeZoomState,
  observeZoom,
  readZoom,
  zoomBounds,
} from '../../src/engine/zoom';

describe('clampZoom', () => {
  it('מגביל לטווח ומעגל', () => {
    expect(clampZoom(150, 50, 200)).toBe(150);
    expect(clampZoom(320, 50, 200)).toBe(200);
    expect(clampZoom(10, 50, 200)).toBe(50);
    expect(clampZoom(99.6, 50, 200)).toBe(100);
  });

  it('ערך שאינו מספר חוזר לברירת המחדל', () => {
    expect(clampZoom(Number.NaN, 50, 200)).toBe(FALLBACK_ZOOM.value);
  });
});

describe('zoomBounds', () => {
  it('מרחיב את התקרה המדווחת אל היקף Word — ה-max של המנוע הוא גבול fit-width ולא מגבלת זום', () => {
    // זה בדיוק מה שהמנוע מדווח בפועל (נמדד ב-bundle): 10–100.
    expect(zoomBounds({ min: 10, max: 100 })).toEqual({ min: 10, max: ZOOM_PERCENT_MAX });
  });

  it('תקרה מדווחת נמוכה מ-500 מורחבת אליו', () => {
    // התקרה שהמנוע מדווח היא גבול fit-width שלו ולא מגבלת זום — לכן מרחיבים.
    expect(zoomBounds({ min: 25, max: 400 })).toEqual({ min: 25, max: ZOOM_PERCENT_MAX });
  });

  it('דיווח תקרה גבוהה מ-500 מנצח', () => {
    expect(zoomBounds({ min: 10, max: 800 })).toEqual({ min: 10, max: 800 });
  });

  it('דיווח פגום נופל לגבולות ברירת המחדל', () => {
    expect(zoomBounds(null)).toEqual({ min: FALLBACK_ZOOM.min, max: FALLBACK_ZOOM.max });
    expect(zoomBounds({})).toEqual({ min: FALLBACK_ZOOM.min, max: FALLBACK_ZOOM.max });
    expect(zoomBounds({ min: '10', max: 100 })).toEqual({ min: FALLBACK_ZOOM.min, max: FALLBACK_ZOOM.max });
    expect(zoomBounds({ min: 400, max: 25 })).toEqual({ min: FALLBACK_ZOOM.min, max: FALLBACK_ZOOM.max });
  });
});

describe('normalizeZoomState', () => {
  it('לוקח את הגבולות מהמנוע, עם תקרה מורחבת לפחות להיקף Word', () => {
    // 400 < 500: התקרה המדווחת מורחבת אל היקף Word, המינימום נשמר מהמנוע.
    expect(normalizeZoomState({ value: 120, min: 25, max: 400 })).toEqual({
      value: 120,
      min: 25,
      max: ZOOM_PERCENT_MAX,
    });
    // התקרה שהמנוע מדווח (גבול ה-fit-width שלו) אינה מגבלה אמיתית —
    // setZoom אינו מצמצם; הסרגל חייב להציע עד 500.
    expect(normalizeZoomState({ value: 100, min: 10, max: 100 })).toEqual({
      value: 100,
      min: 10,
      max: ZOOM_PERCENT_MAX,
    });
  });

  it('מגביל את הערך המדווח לגבולות האפקטיביים', () => {
    expect(normalizeZoomState({ value: 700, min: 25, max: 400 }).value).toBe(ZOOM_PERCENT_MAX);
  });

  it('טווח הפוך נדחה — הוא היה מקפיא כל ערך על אותו מספר', () => {
    expect(normalizeZoomState({ value: 100, min: 400, max: 25 })).toEqual(FALLBACK_ZOOM);
  });

  it('גבולות שאינם מספרים חיוביים נופלים לברירת המחדל', () => {
    expect(normalizeZoomState({ value: 100, min: 0, max: 0 })).toEqual(FALLBACK_ZOOM);
    expect(normalizeZoomState({ value: 100, min: '50', max: '200' })).toEqual(FALLBACK_ZOOM);
    expect(normalizeZoomState(null)).toEqual(FALLBACK_ZOOM);
    expect(normalizeZoomState({})).toEqual(FALLBACK_ZOOM);
  });
});

describe('readZoom', () => {
  it('קורא את ה-snapshot של המנוע', () => {
    const ui = { zoom: { getSnapshot: () => ({ value: 75, min: 10, max: 500 }) } };
    expect(readZoom(ui)).toEqual({ value: 75, min: 10, max: 500 });
  });

  it('גרסת מנוע בלי zoom אינה מפילה', () => {
    expect(readZoom({})).toEqual(FALLBACK_ZOOM);
    expect(readZoom(null)).toEqual(FALLBACK_ZOOM);
  });

  it('snapshot שזורק מוחזר כברירת מחדל ולא כחריגה', () => {
    const ui = {
      zoom: {
        getSnapshot: () => {
          throw new Error('controller מפורק');
        },
      },
    };
    expect(readZoom(ui)).toEqual(FALLBACK_ZOOM);
  });
});

describe('observeZoom', () => {
  it('מדווח על כל שינוי, מנורמל', () => {
    const listeners: Array<(slice: unknown) => void> = [];
    const ui = {
      zoom: {
        getSnapshot: () => ({ value: 100, min: 50, max: 200 }),
        observe: (listener: (slice: { value?: unknown }) => void) => {
          listeners.push(listener as (slice: unknown) => void);
          return () => {};
        },
      },
    };
    const seen: unknown[] = [];

    observeZoom(ui, (state) => seen.push(state));
    listeners[0]({ value: 180, min: 50, max: 200 });

    expect(seen).toEqual([{ value: 180, min: 50, max: ZOOM_PERCENT_MAX }]);
  });

  it('מחזיר disposer, ומעביר את זה של המנוע', () => {
    const unsubscribe = vi.fn();
    const ui = { zoom: { observe: () => unsubscribe } };

    observeZoom(ui, () => {})();

    expect(unsubscribe).toHaveBeenCalledTimes(1);
  });

  it('בלי observe — מדווח פעם אחת ומחזיר disposer שאינו נופל', () => {
    const seen: unknown[] = [];
    const dispose = observeZoom({ zoom: { getSnapshot: () => ({ value: 90, min: 50, max: 200 }) } }, (state) =>
      seen.push(state),
    );

    expect(seen).toEqual([{ value: 90, min: 50, max: ZOOM_PERCENT_MAX }]);
    expect(() => dispose()).not.toThrow();
  });

  it('observe שזורק אינו מפיל את פתיחת המסמך', () => {
    const seen: unknown[] = [];
    const dispose = observeZoom(
      {
        zoom: {
          getSnapshot: () => ({ value: 100, min: 50, max: 200 }),
          observe: () => {
            throw new Error('boom');
          },
        },
      },
      (state) => seen.push(state),
    );

    // observe נפל — אך קריאת snapshot ישירה עדיין תקפה, והיא זו שמגיעה למאזין
    // (עם הרחבת התקרה כרגיל), ולא FALLBACK החוזה.
    expect(seen).toEqual([{ value: 100, min: 50, max: ZOOM_PERCENT_MAX }]);
    expect(() => dispose()).not.toThrow();
  });
});
