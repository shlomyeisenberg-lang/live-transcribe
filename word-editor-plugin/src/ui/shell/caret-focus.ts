/**
 * סמן שמהבהב בלי מיקוד הוא הבטחה שקרית.
 *
 * ## מה נמדד
 *
 * הסמן של המנוע אינו סמן מקורי של הדפדפן אלא `div` בשכבת ציור
 * (`.sd-v2-local-selection-caret`), וההבהוב שלו הוא אנימציית CSS. אנימציה אינה
 * יודעת דבר על מיקוד, ולכן היא ממשיכה לרוץ גם כשהדף איבד אותו. נמדד ב-Chrome
 * אמיתי, על המסמך הריק שנפתח בהפעלה:
 *
 *   | מצב                     | `document.hasFocus()` | `activeElement` | האנימציה |
 *   |-------------------------|-----------------------|-----------------|----------|
 *   | בסיס                    | `true`                | `TEXTAREA` באזור המסמך | רצה |
 *   | **לשונית/חלון אחר**     | **`false`**           | `TEXTAREA` — לא זז     | **עדיין רצה** |
 *   | חזרה                    | `true`                | `TEXTAREA`             | רצה |
 *
 * השורה האמצעית היא הדיווח: „אני על חלון אחר והסמן ממשיך להבהב”. והיא גם
 * מסבירה למה `:focus-within` ב-CSS לא היה עוזר — **המיקוד לא זז**. הדפדפן אינו
 * מזיז `activeElement` כשהחלון מאבד מיקוד, הוא רק מפסיק לספק אותו: `blur`
 * על ה-`window` נורה פעם אחת, `document.hasFocus()` מתהפך, וזה כל האות שיש.
 *
 * ## הכלל
 *
 * ‏„מהבהב” פירושו „ההקלדה תיכנס לכאן”, ולכן שני התנאים חייבים להתקיים יחד:
 *
 *   1. **החלון ממוקד** — `document.hasFocus()`.
 *   2. **המיקוד בתוך אזור המסמך** — ולא ברצועה, בדיאלוג או על `<body>`.
 *
 * התנאי השני אינו הרחבה מיותרת: הוא בדיוק אותו כשל מהצד הפנימי. תיבת הרצועה
 * לוקחת מיקוד אמיתי (`RibbonCombo`), המשתמש מקליד לתוכה, והסמן במסמך ממשיך
 * להבהב כאילו הטקסט נכנס לשם. ב-Word הסמן נעלם בשני המקרים.
 *
 * מה שנשאר מהבהב הוא בדיוק מה ששומרי המיקוד של התוסף מגנים עליו: לחיצה על רקע
 * של פס מעטפת (`caret-keeper.ts`) ולחיצה בשטח האפור (`engine/click-focus.ts`)
 * אינן מוציאות את המיקוד מהמסמך מלכתחילה, ולכן אינן מכבות כאן שום דבר.
 *
 * ## מה זה **אינו**
 *
 * אין כאן שום ניסיון להחזיר מיקוד, ואין כאן תיקון לאיבוד מיקוד. מיקוד שיצא
 * יצא, ולפעמים מסיבה לגיטימית לגמרי; הכלל כאן רק דואג שהמסך יגיד את האמת על
 * המצב. ההחלה עצמה היא CSS בלבד (`styles/shell.css`), ואינה נוגעת בשכבת הציור
 * של המנוע.
 */

export interface CaretFocusInput {
  /** החלון עצמו ממוקד. */
  readonly windowFocused: boolean;
  /** המיקוד כרגע בתוך אזור המסמך. */
  readonly documentFocused: boolean;
}

/** ההכרעה עצמה, בלי DOM. `true` = הסמן מצויר ומהבהב. */
export function shouldPaintCaret({ windowFocused, documentFocused }: CaretFocusInput): boolean {
  return windowFocused && documentFocused;
}

export interface CaretFocusOptions {
  /** אזור המסמך. `null` = אין עוד מה למקד, ולכן גם אין סמן. */
  readonly documentArea: () => HTMLElement | null;
  /** נקראת רק כשהתשובה **משתנה**, ומיד עם ההתקנה. */
  readonly onChange: (paint: boolean) => void;
  /** החלון שעליו מאזינים. הפרמטר קיים כדי שבדיקה תמסור כפיל. */
  readonly view?: Window;
}

/**
 * אין כאן `refresh()`.
 *
 * הייתה כזאת, ואיש לא קרא לה. ידית שמחזירה פעולה שאיש אינו מפעיל נראית כמו
 * כיסוי לחור שנשאר פתוח, ובמקרה הזה גם החור עצמו נמדד כסגור — ראו הטבלה
 * ב-`watchCaretFocus`.
 */
export interface CaretFocusHandle {
  dispose(): void;
}

/**
 * ארבעת האירועים, וכל אחד מהם נמדד כנחוץ:
 *
 *   - `focus`/`blur` על ה-`window` — המעבר לחלון אחר וממנו, השורה האמצעית
 *     בטבלה שלמעלה. הם אינם מבעבעים, ולכן `blur` של אלמנט אינו מגיע לכאן.
 *   - `focusin` — מיקוד שנכנס לפקד כלשהו, ובכלל זה חזרה למסמך.
 *   - `focusout` — **ובלעדיו נשאר חור**: מיקוד שנופל ל-`<body>` (לחיצה על שטח
 *     שאינו פקד) אינו מייצר `focusin` על אף אלמנט, ורק `focusout` מדווח עליו.
 *
 * `focusout` מתפרסם לפני שהפקד הבא קיבל מיקוד, כלומר כש-`activeElement` הוא
 * כבר `<body>`. זה אינו הבהוב: שני האירועים של אותה לחיצה מתפרסמים באותו
 * task, והדפדפן אינו מצייר ביניהם.
 *
 * מקור האמת הוא תמיד `document.hasFocus()` בזמן הקריאה ולא תוכן האירוע, ולכן
 * כל אירוע שנוחת מסנכרן את המצב המלא — גם אם מישהו הגיע מכיוון שלא נצפה.
 *
 * ## והמקרה שנראה כמו חור ואינו
 *
 * מיקוד יכול להיעלם גם בלי שהמשתמש הזיז אותו: הפאנל שבתוכו יושב משטח ההקלדה
 * של המנוע יוצא מתחת לרגליו. **ושתי הדרכים שונות**: סגירת טאב מסירה אותו
 * מה-DOM, אבל החלפת טאב — המסלול השכיח בהרבה — אינה מסירה דבר, אלא מציבה
 * `display: none` על הפאנל של הקודם (App.vue, `activateTab`). השאלה היא אם
 * הדפדפן מדווח, ולכן היא נמדדה — Chrome 152, ארבעת המבנים, ב-
 * `scripts/qa/caret-focus-qa.mjs`:
 *
 *   | מה קרה לפאנל | האירועים שנורו | `activeElement` אחרי |
 *   |---|---|---|
 *   | האלמנט הממוקד עצמו הוסר | `blur`, `focusout` | `BODY` |
 *   | **אב** של האלמנט הממוקד הוסר | `blur`, `focusout` | `BODY` |
 *   | **אב הוסתר ב-`display: none`** | `blur`, `focusout` | `BODY` |
 *   | אב הוסתר ב-`visibility: hidden` | `blur`, `focusout` | `BODY` |
 *
 * כלומר `focusout` מגיע בכל אחד מהם, וארבעת המאזינים מספיקים. אין
 * `MutationObserver` ואין `refresh()` לקורא — שניהם היו קוד שמכסה על חור
 * שנמדד כסגור. שימו לב ש-jsdom **אינו** יורה את האירועים האלה, ולכן בדיקת
 * יחידה שתיבנה על ההתנהגות שלו תמדוד את סביבת הבדיקה ולא את הדפדפן; השורות
 * שבטבלה נמדדות בשער בלבד.
 *
 * ומה שהשער **אינו** מודד: הוא מריץ Chrome עצמאי, בעוד התוסף רץ בתוך ה-WebView
 * של אוצריא — ושם יש פער מיקוד/מקלדת מתועד (מיזעור והחזרה משאירים את ה-WebView
 * בלי מקלדת עד לחיצה; Otzaria#1356). כלומר הטבלה תקפה להתנהגות הדפדפן, ולא
 * הבטחה שכל מסלול מיקוד במארח מתנהג כך.
 */
export function watchCaretFocus({
  documentArea,
  onChange,
  view = window,
}: CaretFocusOptions): CaretFocusHandle {
  const doc = view.document;
  let painted: boolean | null = null;

  const refresh = (): void => {
    const area = documentArea();
    const active = doc.activeElement;
    const paint = shouldPaintCaret({
      windowFocused: doc.hasFocus(),
      documentFocused: area !== null && active !== null && area.contains(active),
    });
    if (paint === painted) return;
    painted = paint;
    onChange(paint);
  };

  const listener = refresh as EventListener;
  view.addEventListener('focus', listener);
  view.addEventListener('blur', listener);
  doc.addEventListener('focusin', listener);
  doc.addEventListener('focusout', listener);
  refresh();

  return {
    dispose() {
      view.removeEventListener('focus', listener);
      view.removeEventListener('blur', listener);
      doc.removeEventListener('focusin', listener);
      doc.removeEventListener('focusout', listener);
    },
  };
}
