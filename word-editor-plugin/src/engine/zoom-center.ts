/**
 * מרכוז העמוד בכל אחוז זום — הגיאומטריה שהמנוע משאיר פתוחה.
 *
 * ## מה נמדד (CDP על ה-dist הארוז, superdoc@2.8.0)
 *
 * המנוע מקנה את הזום בשני חלקים על מיכל העימוד (`.presentation-editor`):
 * `transform: scale(z)` ולצידו `width: (100/z)%` — פיצוי כדי שקופסת הפריסה
 * כפול ה-scale תחזיר את רוחב העמוד המקורי. בתוך המיכל יושב העמוד
 * (`.superdoc-page`) ברוחב קבוע `P` ועם `margin: 0 auto` של המנוע.
 *
 * מכאן שתי התנהגויות מדודות:
 *
 * - **z < 1** — קופסת הפריסה רחבה מהעמוד, ה-`auto` מתחלק לשני הצדדים
 *   והעמוד ממורכז בתוכה. עם `transform-origin` בקצה ההתחלה של הקופסה — פינה
 *   שהמנוע עצמו מעגן בה (ראו shell.css) — התוצאה ממורכזת גם על המסך. זה המצב
 *   שתוקן קודם, והוא נשאר נכון.
 *
 * - **z > 1** — קופסת הפריסה **צרה** מהעמוד, המרווח הפנוי שלילי, ושני
 *   ה-`auto` קורסים לאפס. העמוד נדבק לקצה ההתחלה של הקופסה (בממשק עברי —
 *   הימני) וכל הגדילה יוצאת שמאלה בלבד. נמדד בחלון 1440: ב-150% העמוד
 *   (1191px) עדיין נכנס למאגס (1425px) אבל מרכזו ב-529 מול מרכז מאגס 727 —
 *   סטייה של 199px, עם רצועה אפורה מימין. ב-200% הסטייה 397px. זה בדיוק
 *   „ההגדלה לא מתמרכזת, הדף זז הצידה”.
 *
 * ## מה שאי אפשר לעשות במקום
 *
 * מרכוז „לא בטוח” בלבד — לאפס את ה-`auto` ולתת ל-`align-items: center` של
 * המנוע לעבוד — אכן ממרכז בכל אחוז, אבל **מוציא את תחילת השורה מהישג יד**:
 * גלישה מעבר לקצה ההתחלה של מיכל הגלילה אינה נכנסת לאזור הגלילה. נמדד
 * ב-300% בחלון 1440: הקצה הימני של העמוד נשאר ב-1918 גם בגלילה מלאה ימינה,
 * כלומר 493px מתחילת כל שורה בעברית — בלתי נגישים. לכן הכלל חייב להכיר את
 * רוחב המאגס ולא רק את הזום.
 *
 * ## הנוסחה
 *
 * הפקד היחיד שנדרש הוא שולי ההתחלה של העמוד בתוך קופסת הפריסה, `a`, במידות
 * לפני ה-scale. עם origin בקצה ההתחלה של הקופסה, הקצה המרונדר של העמוד הוא
 * `R - a·z` (R = קצה הקופסה, קבוע). מכאן:
 *
 *   - כשהעמוד המרונדר נכנס למאגס (`P·z ≤ V`) רוצים מרכוז: `a = P/(2z) - P/2`
 *   - כשהוא גדול ממנו רוצים הצמדה לקצה ההתחלה של המאגס, כדי שכל העמוד יהיה
 *     נגיש בגלילה: `a = P/(2z) - V/(2z)`
 *
 * שתי הצורות הן אותה נוסחה: `a = P/(2z) - min(P, V/z)/2`. וזו בדיוק
 * הסמנטיקה של `safe center`, אלא שהיא נמדדת מול רוחב המאגס ולא מול קופסת
 * הפריסה — הבחנה ש-CSS לבדו אינו יכול לעשות, מפני שה-scale אינו קיים בפריסה.
 *
 * ב-CSS `100%` בשולי העמוד נפתר מול רוחב קופסת הפריסה, שהוא `P/z` — ולכן
 * `P = 100%·z`, `P/(2z) = 50%`, ו-`V/z = V·(1/z)`. הכלל ב-shell.css מחשב את
 * הנוסחה בדיוק כך, וכל מה שהמודול הזה עושה הוא לפרסם את שלושת הנעלמים
 * שאינם ידועים ל-CSS: `z`, `1/z` ורוחב המאגס.
 *
 * הרווח מהצורה הזאת הוא ש-`P` לעולם אינו נמדד ב-JS: הוא נכנס דרך האחוזים.
 * מסמך שמתחלק לעמודים אחרת, מקטע לרוחב, או ריבוי עמודים — כולם מקבלים את
 * המרכוז הנכון בלי שנקרא ולו אלמנט אחד של המנוע.
 *
 * ## הנוסחה אינה תלויה בכיווניות — נקודת הפתיחה כן
 *
 * `margin-inline-start` ופינת העיגון של ה-scale מתהפכים יחד עם כיווניות מיכל
 * הגלילה, ולכן הנוסחה נשארה נכונה כשהמיכל עבר ל-`direction: ltr` (הצהרה על
 * צד סרגל הגלילה, ראו shell.css). מה שאינו סימטרי הוא **איפה מיכל הגלילה נח**:
 * מיכל ltr נח ב-`scrollLeft = 0`, כלומר בקצה השמאלי — ובמסמך עברי שגולש זהו
 * צד **סופי** השורות. ההצמדה שבנוסחה מבטיחה שכל העמוד נגיש בגלילה, אבל היא
 * אינה קובעת לאן הגלילה מצביעה.
 *
 * לכן `setZoom` מצמיד את הגלילה בחזרה לתחילת השורה — ורק בשלושה תנאים יחד:
 * המסמך מוצהר RTL (`data-document-direction` על שורש ה-HTML, שנקבע
 * ב-engine/document-defaults.ts רק כשכל שכבות הכיווניות הסכימו), המיכל אכן
 * גולש אופקית, והשינוי בא משינוי זום ולא משינוי גודל — `refresh` אינו מצמיד,
 * מפני שגרירת החלון אינה בקשה לזוז.
 */

/**
 * העמוד שבתוך מיכל העימוד — האלמנט שהכלל ב-shell.css מזיז את שוליו.
 *
 * הקבוע כאן ולא ליטרל ב-CSS בלבד, מאותה סיבה של `ZOOM_LAYOUT_CLASS`
 * (engine/fit-width.ts): שער ההיגיינה (tests/unit/css-hygiene.test.ts) מחייב
 * שכל מחלקה בסלקטור גלובלי תופיע בקוד, והגדרה מפורשת היא מה שמוכיח שהמחלקה
 * קיימת ולא נכתבה בטעות. שינוי שם המחלקה במנוע ינטרל את הכלל בשקט — ושער
 * scripts/zoom-center-probe.mjs נופל על כך.
 */
export const ZOOM_PAGE_CLASS = 'superdoc-page';

/** אחוז הזום כמספר יחס (`1.5` ל-150%). */
export const ZOOM_VAR = '--otzaria-zoom';

/** ההופכי, `1/z`. מפורסם כדי ש-CSS לא יידרש לחלק במשתנה. */
export const ZOOM_INVERSE_VAR = '--otzaria-zoom-inverse';

/** רוחב התוכן של מיכל הגלילה, בפיקסלי CSS. */
export const VIEWPORT_WIDTH_VAR = '--otzaria-viewport-width';

/**
 * מיכל הגלילה שבתוכו המנוע מרנדר. הקלס הוא שלנו (sessions/editor-swap.ts),
 * ולכן המדידה כאן אינה נגיעה ב-DOM הפנימי של המנוע.
 */
export const HOST_SELECTOR = '.editor-stack__host';

/**
 * הכיווניות שהאפליקציה מצהירה עליה על שורש ה-HTML. נקראת ולא נגזרת מחדש:
 * engine/document-defaults.ts כותב אותה רק כששלוש שכבות הכיווניות של המנוע
 * הסכימו, ושער check:rtl מודד אותה. גזירה עצמאית כאן הייתה מקור אמת שני.
 */
export const DOCUMENT_DIRECTION_ATTRIBUTE = 'data-document-direction';

/** הערכים שהכלל ב-shell.css צורך. */
export interface ZoomCenterVars {
  [ZOOM_VAR]: string;
  [ZOOM_INVERSE_VAR]: string;
  [VIEWPORT_WIDTH_VAR]: string;
}

/** מעגל ערך למספר ספרות סבירות — משתנה CSS אינו זקוק ל-17 ספרות. */
function round(value: number): string {
  return String(Math.round(value * 1e4) / 1e4);
}

/**
 * הערכים לפי אחוז הזום ורוחב המאגס.
 *
 * דיווח פגום — אחוז שאינו מספר חיובי — נופל ל-100%, ורוחב מאגס שאינו נמדד
 * נופל ל-`100vw`: רחב לפחות כמו המאגס, ולכן הכלל מתנהג כמרכוז מלא. זו
 * הנפילה הנכונה, שהרי מרכוז שגוי ב-15px עדיף על הצמדה לקצה בלי סיבה.
 */
export function zoomCenterVars(percent: number, viewportPx: number): ZoomCenterVars {
  const zoom = Number.isFinite(percent) && percent > 0 ? percent / 100 : 1;
  const viewport = Number.isFinite(viewportPx) && viewportPx > 0 ? `${Math.round(viewportPx)}px` : '100vw';

  return {
    [ZOOM_VAR]: round(zoom),
    [ZOOM_INVERSE_VAR]: round(1 / zoom),
    [VIEWPORT_WIDTH_VAR]: viewport,
  };
}

/** מה שהמודול צריך מהמאגס. אלמנט אמיתי מקיים את זה מאליו. */
type StackElement = Pick<HTMLElement, 'clientWidth'> & {
  querySelector: (selector: string) => Element | null;
  style: Pick<CSSStyleDeclaration, 'setProperty'>;
};

/**
 * האם המסמך המוצג מוצהר RTL. בסביבה בלי `document` — אין, ולכן אין הצמדה.
 */
function isRightToLeftDocument(): boolean {
  if (typeof document === 'undefined') return false;
  return document.documentElement.getAttribute(DOCUMENT_DIRECTION_ATTRIBUTE) === 'rtl';
}

/**
 * מצמיד את הגלילה האופקית לתחילת השורה, במיכל ltr שמציג מסמך RTL.
 *
 * `scrollWidth - clientWidth` הוא ה-`scrollLeft` המקסימלי, ובמיכל ltr הוא
 * הקצה הימני — כלומר תחילת השורה בעברית. מיכל שאינו גולש אופקית מחזיר אפס,
 * וההשמה הופכת ל-no-op.
 */
function anchorToLineStart(host: HTMLElement): void {
  if (!isRightToLeftDocument()) return;
  const max = host.scrollWidth - host.clientWidth;
  if (max > 0) host.scrollLeft = max;
}

/** הפקד שהאפליקציה מחזיקה. */
export interface ZoomCenter {
  /** מעדכן את הזום ומחיל מחדש. נקרא מכל דיווח של `observeZoom`. */
  setZoom: (percent: number) => void;
  /** מחיל מחדש עם הזום האחרון — למשל אחרי שינוי גודל. */
  refresh: () => void;
  dispose: () => void;
}

/**
 * מחבר את המרכוז למאגס.
 *
 * רוחב המאגס נמדד מ-`clientWidth` של מיכל הגלילה ולא של ה-`<main>`: פס
 * הגלילה האנכי גורע 15px מהראשון בלבד, ומרכוז שמתעלם ממנו סוטה בחצי מזה
 * וההצמדה בכולו.
 *
 * `ResizeObserver` על שני האלמנטים, ובכוונה: על ה-`<main>` נתפס שינוי גודל
 * החלון, ועל מיכל הגלילה — ב-`content-box` — נתפסת גם הופעת פס הגלילה, שאינה
 * משנה כלל את קופסת הגבול של המיכל. מסמך שגדל לעמוד שני היה אחרת מזיז את
 * המרכוז ב-7.5px בלי שדבר יאזין.
 */
export function createZoomCenter(stack: StackElement): ZoomCenter {
  let percent = 100;
  let observed: Element | null = null;

  const observer =
    typeof ResizeObserver === 'function'
      ? new ResizeObserver(() => {
          apply();
        })
      : null;

  /**
   * ההצמדה נדחית גם לפריים הבא, ולא רק מיד: המנוע מחיל את ה-`scale()` שלו
   * בעצמו, ומדידה של `scrollWidth` באותו tick יכולה לקרוא את הרוחב שלפני
   * השינוי. שתי הקריאות אידמפוטנטיות, ולכן הכפילות אינה עולה דבר.
   */
  function anchorSoon(host: HTMLElement): void {
    anchorToLineStart(host);
    if (typeof requestAnimationFrame === 'function') {
      requestAnimationFrame(() => anchorToLineStart(host));
    }
  }

  function apply(): void {
    const host = stack.querySelector(HOST_SELECTOR);

    if (observer && host !== observed) {
      if (observed) observer.unobserve(observed);
      if (host) observer.observe(host, { box: 'content-box' });
      observed = host;
    }

    const viewport = host instanceof HTMLElement ? host.clientWidth : stack.clientWidth;
    const vars = zoomCenterVars(percent, viewport);
    for (const [name, value] of Object.entries(vars)) stack.style.setProperty(name, value);
  }

  if (observer && stack instanceof Element) observer.observe(stack);
  apply();

  return {
    setZoom(next) {
      percent = next;
      apply();
      // רק כאן, ולא ב-`refresh`: שינוי זום הוא בקשה לזוז, שינוי גודל אינו.
      const host = stack.querySelector(HOST_SELECTOR);
      if (host instanceof HTMLElement) anchorSoon(host);
    },
    refresh: apply,
    dispose() {
      observer?.disconnect();
      observed = null;
    },
  };
}
