/**
 * גודל התצוגה, כפי שהמנוע מדווח אותו.
 *
 * למה מודול ולא קריאה ישירה מהקומפוננטה: הסרגל ב-StatusBar.vue קידד 50–200
 * כמספרים בתבנית, ואיש לא בדק מול המנוע. הגבולות האמיתיים יושבים
 * ב-`ui.zoom.getSnapshot()` (`min`/`max`), ומסמך שהמנוע מגביל אחרת היה מקבל
 * סרגל שנע לערכים שהמנוע דוחה — כלומר הסרגל זז והמסמך לא.
 *
 * **התקרה שאנחנו קובעים.** ה-max שהמנוע מדווח ב-snapshot אינו מגבלה על זום
 * ידני אלא גבולות ה-fit-width שלו (ברירת מחדל 10–100; נמדד ב-bundle של
 * superdoc@2.8.0: הפסאדה מעתיקה את `getZoomState().min/max`, שהם גבולות
 * ההתאמה לרוחב). `setZoom` עצמו מקבל כל מספר חיובי בלי clamp — אותה מדידה.
 * לכן `normalizeZoomState` מרחיב את התקרה המדווחת אל `ZOOM_PERCENT_MAX`
 * (ההיקף של Word: 10%–500%), וכל הפקדים — הסרגל, לחצני ±, „גודל אמיתי”
 * ו„רוחב עמוד” — חולקים אותו דרך `zoomBounds`. אם גרסת מנוע עתידית תדווח
 * תקרה גבוהה מ-500, היא מנצחת; נמוכה ממנו — אנחנו מנצחים, כי הדיווח אינו
 * מגבלה אמיתית של מסלול הכתיבה.
 *
 * **הכתיבה אינה כאן.** שינוי זום נעשה דרך הפקודה `zoom` של ה-controller
 * (`engine/payloads.ts` → `zoomPayload`), ולא דרך `ui.zoom.set`: כל פעולה
 * שהמשתמש מפעיל עוברת ב-command-adapter, ושם היא מקבלת הודעת כשל בעברית ומצב
 * `enabled` לפקד. שני מסלולי כתיבה לאותו דבר היו שני חוזים לתחזק.
 *
 * הצורה מוגדרת כאן מבנית ואינה מיובאת מ-`superdoc/ui`: הפאסדה אינה מייצאת
 * `ZoomHandle`/`ZoomSlice`, ו-import מנתיב פנימי של החבילה אסור
 * (tests/unit/engine-boundaries.test.ts). הכול אופציונלי — גרסת מנוע בלי
 * `zoom` נופלת בחן לגבולות ברירת המחדל ולא בחריגה.
 */

/** גודל התצוגה והגבולות שהמנוע מתיר, באחוזים. */
export interface ZoomState {
  value: number;
  min: number;
  max: number;
}

/**
 * תקרת הזום שמוצעת למשתמש, באחוזים — ההיקף של Word (שם הסליידר עד 500%).
 * ראו הערת הפתיחה: זו החלטה שלנו ולא דיווח של המנוע, שכן מסלול הכתיבה
 * (`setZoom`) אינו מצמצם כלל.
 */
export const ZOOM_PERCENT_MAX = 500;

/**
 * מה שמוצג כשאין מסמך פתוח, וכשהמנוע אינו מדווח גבולות. ההיקף של Word
 * (10%–500%) — מכאן והלאה ברירת מחדל אחרונה ולא החוזה.
 */
export const FALLBACK_ZOOM: ZoomState = { value: 100, min: 10, max: ZOOM_PERCENT_MAX };

/** מה שנצרך מ-`superdoc.ui`. ראו הערת הפתיחה. */
export interface ZoomSource {
  zoom?: {
    getSnapshot?: () => { value?: unknown; min?: unknown; max?: unknown } | null | undefined;
    observe?: (listener: (slice: { value?: unknown; min?: unknown; max?: unknown }) => void) => () => void;
  };
}

function positive(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) && value > 0 ? value : null;
}

/** מגביל ערך לטווח. הסרגל, הצעדים ולחצני ± כולם עוברים דרכה. */
export function clampZoom(value: number, min: number, max: number): number {
  if (!Number.isFinite(value)) return FALLBACK_ZOOM.value;
  return Math.min(max, Math.max(min, Math.round(value)));
}

/**
 * הגבולות האפקטיביים מתוך דיווח גולמי של המנוע.
 *
 * המנוע מדווח את גבולות ה-fit-width שלו ולא מגבלת זום ידני (ראו הערת
 * הפתיחה), ולכן התקרה מורחבת לפחות אל `ZOOM_PERCENT_MAX`; דיווח תקרה גבוהה
 * יותר מנצח. דיווח פגום — שאינו מספרים חיוביים, או טווח הפוך (`min >= max`)
 * שהיה מקפיא כל ערך על אותו מספר — נופל בחן לגבולות ברירת המחדל.
 */
export function zoomBounds(reported: { min?: unknown; max?: unknown } | null | undefined): Pick<ZoomState, 'min' | 'max'> {
  const min = positive(reported?.min);
  const max = positive(reported?.max);
  if (min === null || max === null || min >= max) {
    return { min: FALLBACK_ZOOM.min, max: FALLBACK_ZOOM.max };
  }
  return { min, max: Math.max(max, ZOOM_PERCENT_MAX) };
}

/**
 * מנרמלת slice של המנוע ל-`ZoomState` שאפשר לסמוך עליו.
 *
 * הגבולות דרך `zoomBounds` (כולל הרחבת התקרה), והערך המדווח מצומצם אליהם —
 * כך התווית אף פעם לא מציגה מספר שמחוץ לטווח שהפקדים מציעים.
 */
export function normalizeZoomState(slice: {
  value?: unknown;
  min?: unknown;
  max?: unknown;
} | null | undefined): ZoomState {
  const range = zoomBounds(slice);

  const value = positive(slice?.value) ?? FALLBACK_ZOOM.value;
  return { ...range, value: clampZoom(value, range.min, range.max) };
}

/** הערכים ברגע זה. כשל קריאה מוחזר כברירת המחדל ולא כחריגה. */
export function readZoom(ui: ZoomSource | null | undefined): ZoomState {
  const zoom = ui?.zoom;
  if (typeof zoom?.getSnapshot !== 'function') return { ...FALLBACK_ZOOM };
  try {
    return normalizeZoomState(zoom.getSnapshot());
  } catch (error) {
    console.warn('[otzaria-word] קריאת גודל התצוגה מהמנוע נכשלה', error);
    return { ...FALLBACK_ZOOM };
  }
}

/**
 * מאזינה לגודל התצוגה. `observe` של המנוע יורה מיד עם ה-snapshot ואז על כל
 * שינוי — כולל שינויים שלא באו מאיתנו (`zoom-fit-width`, התאמה לרוחב חלון) —
 * ולכן ההאזנה היא מה שמונע מהתווית בשורת המצב להראות ערך שהמסמך כבר לא בו.
 *
 * מחזירה disposer גם כשאין `observe`, כדי שאתר הקריאה לא יצטרך להבחין.
 */
export function observeZoom(
  ui: ZoomSource | null | undefined,
  listener: (state: ZoomState) => void,
): () => void {
  const zoom = ui?.zoom;
  const observe = zoom?.observe;

  if (typeof observe !== 'function') {
    listener(readZoom(ui));
    return () => {};
  }

  try {
    return observe.call(zoom, (slice) => listener(normalizeZoomState(slice)));
  } catch (error) {
    console.warn('[otzaria-word] האזנה לגודל התצוגה נכשלה', error);
    listener(readZoom(ui));
    return () => {};
  }
}
