/**
 * מברשת עיצוב (Format Painter) — חיווט המאזינים שהמנוע דורש כדי **להחיל**.
 *
 * ## מה נמדד, ולמה המודול הזה קיים
 *
 * לחיצה על „מברשת עיצוב” מריצה את הפקודה `copy-format`, וזו מזיינת את
 * הפקד כהלכה: `ui.commands.get('copy-format').getState().active` עובר
 * ל-`true`. אבל סימון היעד אחר כך לא מחיל דבר — נמדד ב-Chrome חי ש-`rPr`
 * של הטקסט ביעד נשאר ריק בכל מסלול. קריאה ידנית ל-
 * `ui.formatPainter.notifyPointerUp()` מיד אחרי סימון היעד החילה בפועל
 * (`<w:sz.../><w:rFonts.../><w:b/>` הופיעו ב-XML).
 *
 * הסיבה: `ui.formatPainter` הוא, לפי התיעוד שלו ב-`superdoc/ui`, "Format-
 * painter surface (**DOM listener coordination**)" — כלומר הוא לא צופה על
 * העכבר או המקלדת בעצמו; מישהו אחר צריך להגיד לו "עכשיו נגררה בחירה" /
 * "המקש שוחרר". ב-SuperDoc הרגיל זה `SuperToolbar` (`arm/disarmFormatPainter
 * Listeners`), אבל האפליקציה בונה את המנוע עם `ui: false`
 * (create-editor.ts), ולכן `SuperToolbar` אף פעם לא קם ואף אחד לא קורא
 * ל-`notifyPointerUp` / `notifyKeyUp`. חימוש בלי מי שיפעיל את ההחלה.
 *
 * ## המשטח, לפי `superdoc/ui` (`FormatPainterHandle`)
 *
 *   - `setPointerSelecting(flag)` — "Notify the controller that a
 *     pointer-drag selection has started."
 *   - `notifyPointerUp()` — "Notify the controller that the pointer was
 *     released; **triggers apply** if a non-source selection exists."
 *   - `setKeyboardSelecting(flag)` — "Notify the controller that a
 *     keyboard selection key is held."
 *   - `notifyKeyUp()` — "Notify the controller that the keyboard selection
 *     key was released; **triggers apply**."
 *   - `cancel()` — "Cancel an active painter (Esc or programmatic
 *     cancel)."
 *
 * ## איך מחווטים את זה
 *
 * הבסיס הוא בדיוק מה ש-`SuperToolbar` עצמו עושה על ה-container של העורך
 * כש-`ui` אינו מכובה (`#armFormatPainterListeners` ב-`superdoc.es.js`):
 * `pointerdown`/`keydown` מזינים `setPointerSelecting`/`setKeyboardSelecting`,
 * ו-`pointerup`/`keyup` קוראים ל-`notifyPointerUp`/`notifyKeyUp`; `Escape`
 * קורא `cancel()`. אלה בדיוק ארבעת סוגי האירועים שה-DOM מבטיח שיגיעו בכל
 * סביבה (כולל WebView2 ב-Windows, ראו word-selection.ts).
 *
 * המאזינים נרשמים על ה-container של המנוע (כמו `installWordSelection`),
 * לא על ה-DOM הפנימי שלו: זה מה ש-tests/unit/engine-boundaries.test.ts
 * אוכף, וזה גם בדיוק מה ש-`SuperToolbar` עצמו עושה (`activeEditor.container`
 * — קונטיינר של המנוע, לא selector לתוכו).
 *
 * ## הסייג היחיד: מתי *לא* לקרוא ל-`notify*Up`
 *
 * `SuperToolbar` עצמו קורא ל-`notifyPointerUp`/`notifyKeyUp` על **כל**
 * `pointerup`/`keyup`, בלי תנאי. קריאה בקובץ המקור של `ui.formatPainter`
 * (`create-super-doc-ui-*.js`, `maybeApply`/`applyFormatPainter`) מגלה
 * שכשהבחירה החיה בזמן ה-`notify` היא **ריקה** (סמן ללא טווח) — למשל לחיצה
 * שרק ממקמת את הסמן, או מקש חץ בלי Shift שרק מזיז אותו — המנוע נופל למסלול
 * "צביעת פסקה על הסמן": הוא **מסיים** את המצב החמוש (`exitFormatPainter`)
 * גם כשאין שום שינוי עיצוב לצייר. כלומר לחיצה או ניווט תמימים, שנועדו רק
 * למקם את הסמן *לפני* בחירת היעד, מכבים את המברשת *לפני* שהבחירה עצמה
 * מתחילה — נמדד: `caretPara` (קליק) ואז `Home` בלי Shift (למקם סמן), שני אלה
 * לפני לולאת ה-Shift+חץ שבוחרת את היעד, כיבו את המברשת מוקדם מדי, וה-`rPr`
 * של היעד יצא ריק לגמרי.
 *
 * לכן `notifyPointerUp`/`notifyKeyUp` נקראים כאן רק כשיש סיבה טובה להאמין
 * שזו בחירה אמיתית ולא מיקום סמן: גרירה אמיתית (המצביע זז מעבר לסף) או
 * `Shift` לחוץ. זה עדיין אך ורק שימוש ב-API הציבורי המתועד — הבחירה *מתי*
 * לקרוא לו נאמנה למילים המדויקות של התיעוד עצמו ("pointer**-drag** selection",
 * "keyboard selection key... **held**"), ולא ניחוש. `setPointerSelecting`/
 * `setKeyboardSelecting` עצמם עדיין נקראים בלי תנאי בכל `pointerdown`/מקש
 * ניווט, בדיוק כמו ב-`SuperToolbar` — רק ה"מפעיל" (`notify*Up`) מסויג.
 *
 * ## דבאונס: הרחבה בכמה לחיצות מקש נפרדות, לא החזקה אחת
 *
 * `applyFormatPainter` (אותו קובץ מקור) מסיים את המצב החמוש בסוף **כל**
 * החלה מוצלחת — גם כשההחלה כיסתה רק תו אחד. הרחבת בחירה במקלדת ע"י החזקת
 * מקש חץ מיוצרת ע"י ה-OS כרצף `keydown` חוזרים ו**שחרור אחד בלבד** בסוף,
 * ואז `notifyKeyUp` יחיד רואה את הטווח המלא. אבל הרחבה ע"י כמה לחיצות
 * *נפרדות* על אותו מקש (לחיצה-שחרור, לחיצה-שחרור...) — נמדד: בדיוק איך
 * `selectRange` בשער ה-QA מרחיבה בחירה — מייצרת `keyup` נפרד על כל לחיצה,
 * וקריאה ל-`notifyKeyUp` על הראשון שבהם מחילה על תו אחד בלבד ומסיימת את
 * המצב החמוש, לפני שהלחיצות הבאות מספיקות להרחיב את הטווח. התוצאה הנמדדת:
 * ריצה מפוצלת (rPr על "d" בלבד, "stx" נשאר בלי עיצוב) — לא ריצה שלמה.
 *
 * לכן `notifyKeyUp` נדחה ב-{@link KEY_APPLY_DEBOUNCE_MS}: כל Shift+מקש
 * ניווט נוסף מאפס את הטיימר, כך שההחלה בפועל קורית רק אחרי שקט קצר —
 * כשההרחבה כולה כבר הסתיימה, בין אם ברצף לחיצות בודדות ובין אם בהחזקה אחת.
 * אותו דבר, מאותה סיבה, על `notifyPointerUp`: כמה קליקים/Shift-קליקים
 * רצופים (הרחבת בחירה בעכבר בלי גרירה רצופה) לא יסיימו את המצב החמוש
 * לפני שהמשתמש הפסיק לבחור.
 */
import type { SuperDoc } from 'superdoc';

/** מקשי מקלדת שמזיזים בחירה — בדיוק הרשימה ש-`SuperToolbar` בודק. */
const SELECTION_KEYS = new Set([
  'ArrowLeft',
  'ArrowRight',
  'ArrowUp',
  'ArrowDown',
  'Home',
  'End',
  'PageUp',
  'PageDown',
]);

/**
 * כמה מותר לסמן לזוז בין `pointerdown` ל-`pointerup` ועדיין להיחשב לחיצה
 * (מיקום סמן) ולא גרירה (בחירת יעד). אותו סף בדיוק כמו
 * `CLICK_SEQUENCE_SLOP_PX` ב-word-selection.ts, ומאותה סיבה: זעזוע יד קטן
 * בין לחיצה לשחרור לא הופך מיקום סמן לגרירת בחירה.
 */
const DRAG_SLOP_PX = 5;

/**
 * שקט נדרש אחרי Shift+ניווט/קליק אחרון לפני שקוראים בפועל ל-`notify*Up`.
 * ראו ה-JSDoc בראש הקובץ ("דבאונס"). 250ms נדיב ביחס לקצב לחיצה אנושי בין
 * תו לתו (נמדד ב-QA: כ-22ms בין לחיצות סקריפט, אנושי תמיד איטי מזה יותר),
 * ומספיק קטן כדי שהצביעה עדיין תיראה מיידית למשתמש.
 */
const KEY_APPLY_DEBOUNCE_MS = 250;

/** `FormatPainterHandle`, בחלק שנצרך כאן. */
export interface FormatPainterHandleLike {
  setPointerSelecting?: (flag: boolean) => void;
  notifyPointerUp?: () => void;
  setKeyboardSelecting?: (flag: boolean) => void;
  notifyKeyUp?: () => void;
  cancel?: () => void;
}

export interface FormatPainterHost {
  ui?: { formatPainter?: FormatPainterHandleLike | null } | null;
}

/** ה-union מאפשר גם את המופע האמיתי וגם כפיל. ההסבר המלא ב-page-setup.ts. */
export type FormatPainterTarget = SuperDoc | FormatPainterHost | null | undefined;

export interface FormatPainterHandle {
  /** מפסיק להאזין. */
  dispose(): void;
}

/**
 * שולפת את `formatPainter` מהמארח, ומוודאת שכל חמש הפעולות שנצרכות כאן
 * קיימות. `null` פירושו „גרסת מנוע שאינה חושפת את המשטח” — ואז אין מה
 * לחווט, ולא נרשם אף מאזין.
 */
function readPainter(host: FormatPainterTarget): Required<FormatPainterHandleLike> | null {
  const painter = (host as FormatPainterHost | null | undefined)?.ui?.formatPainter;
  if (
    !painter ||
    typeof painter.setPointerSelecting !== 'function' ||
    typeof painter.notifyPointerUp !== 'function' ||
    typeof painter.setKeyboardSelecting !== 'function' ||
    typeof painter.notifyKeyUp !== 'function' ||
    typeof painter.cancel !== 'function'
  ) {
    return null;
  }
  return painter as Required<FormatPainterHandleLike>;
}

/**
 * מתקינה את מאזיני המברשת על ה-container של המנוע.
 *
 * מחזירה handle עם `dispose` תמיד — גם כשהמשטח חסר (`readPainter` מחזיר
 * `null`) — כדי שהקורא (create-editor.ts) יוכל להתקין ולשחרר באחידות בלי
 * ענף מיוחד.
 */
export function installFormatPainter(
  container: HTMLElement,
  host: FormatPainterTarget,
): FormatPainterHandle {
  const painter = readPainter(host);
  if (!painter) return { dispose() {} };

  /** נקודת ה-`pointerdown` — כדי למדוד אם ה-`pointerup` הוא גרירה. */
  let downX = Number.NaN;
  let downY = Number.NaN;

  /** הדבאונס המשותף: ראו „דבאונס” בראש הקובץ. */
  let notifyTimer: ReturnType<typeof setTimeout> | undefined;
  const scheduleNotify = (notify: () => void): void => {
    if (notifyTimer !== undefined) clearTimeout(notifyTimer);
    notifyTimer = setTimeout(() => {
      notifyTimer = undefined;
      notify();
    }, KEY_APPLY_DEBOUNCE_MS);
  };
  const cancelPendingNotify = (): void => {
    if (notifyTimer === undefined) return;
    clearTimeout(notifyTimer);
    notifyTimer = undefined;
  };

  const onPointerDown = (event: PointerEvent): void => {
    downX = event.clientX;
    downY = event.clientY;
    painter.setPointerSelecting(true);
  };

  const onPointerUp = (event: PointerEvent): void => {
    painter.setPointerSelecting(false);
    // "יש סיבה טובה להאמין שזו בחירה": גרירה אמיתית, או Shift שמרחיב בחירה
    // קיימת בלי גרירה (למשל Shift+קליק). ראו ההסבר בראש הקובץ.
    const dragged =
      Math.abs(event.clientX - downX) > DRAG_SLOP_PX || Math.abs(event.clientY - downY) > DRAG_SLOP_PX;
    if (dragged || event.shiftKey) scheduleNotify(() => painter.notifyPointerUp());
  };

  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') {
      cancelPendingNotify();
      painter.cancel();
      return;
    }
    if (SELECTION_KEYS.has(event.key)) painter.setKeyboardSelecting(true);
  };

  const onKeyUp = (event: KeyboardEvent): void => {
    painter.setKeyboardSelecting(false);
    // רק Shift+מקש ניווט מרחיב בחירה. אותו מקש בלי Shift רק מזיז את הסמן —
    // ראו ההסבר בראש הקובץ למה קריאה כאן הייתה מכבה את המברשת מוקדם מדי.
    if (SELECTION_KEYS.has(event.key) && event.shiftKey) scheduleNotify(() => painter.notifyKeyUp());
  };

  container.addEventListener('pointerdown', onPointerDown);
  container.addEventListener('pointerup', onPointerUp);
  container.addEventListener('keydown', onKeyDown);
  container.addEventListener('keyup', onKeyUp);

  return {
    dispose() {
      cancelPendingNotify();
      container.removeEventListener('pointerdown', onPointerDown);
      container.removeEventListener('pointerup', onPointerUp);
      container.removeEventListener('keydown', onKeyDown);
      container.removeEventListener('keyup', onKeyUp);
    },
  };
}
