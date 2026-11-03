/**
 * „מקלידים ולא רואים” — הדף שאינו זז מצד לצד בחלון צר.
 *
 * ## הדיווח
 *
 * „יש בעיה בהקלדה כשהחלון קטן — לא מזיזה את הדף מצד לצד.”
 *
 * ## מה נמדד (CDP על ה-dist הארוז, superdoc 2.14.0-next.5, מסמך עברי)
 *
 * הגרסה היא זו שהייתה **מותקנת** בזמן המדידה, ולא זו ש-package.json מכריז
 * עליה (2.15.0) — `node_modules` פיגר אחריה, וכל מה שנמדד כאן רץ על המותקן.
 * נמדד מחדש ב-17.9.2026 על 2.15.0 המותקנת (Chrome 152): אותם מספרים בדיוק,
 * ואותה תשובה מ-`scrollIntoView`. הטבלה למטה תקפה לשתי הגרסאות.
 *
 * עמוד A4 הוא 794px, ומיכל הגלילה שלנו (`.editor-stack__host`) צר ממנו ברגע
 * שהחלון יורד מתחת ל-‎~800px. המיכל הוא `direction: ltr` — הצהרה על צד פס
 * הגלילה, ראו styles/shell.css — ולכן הוא **נח על `scrollLeft = 0`, שהוא
 * הקצה השמאלי**. במסמך עברי הקצה השמאלי הוא *סוף* השורה, וכל צד תחילת-השורה
 * נשאר מחוץ לתצוגה.
 *
 * בחלון של 600px מיכל הגלילה מדווח `clientWidth` 585 — פס הגלילה האנכי
 * גורע את ההפרש — מול `scrollWidth` 794, כלומר 209px של גלילה זמינה:
 *
 * | מה נעשה | `scrollLeft` | x של הסמן | תצוגה נגמרת ב- |
 * |---|---|---|---|
 * | הבחירה הועברה לתחילת השורה | 0 | 698 | 585 — **מחוץ למסך** |
 * | `ui.viewport.scrollIntoView`, שלושת ה-`block` | **0** | 698 | לא זז |
 * | הוקלדו 3 תווים | **0** | 672 | נכנסו ל-OOXML, לא נראים |
 * | בקרה: `scrollLeft = max` ביד | 209 | 463 | גלוי |
 *
 * שתי השורות האמצעיות הן הפער עצמו: הטקסט **כן** נכנס למסמך (ההיסט עלה
 * מ-0 ל-3 וה-OOXML מאשר), והסמן פשוט מצויר במקום שאי אפשר לראות. שורת
 * הבקרה היא שמוכיחה שאין כאן מניעה — הגלילה זמינה, ואיש אינו מבצע אותה.
 *
 * ## למה לא `ui.viewport.scrollIntoView`
 *
 * מפני שהוא **מחזיר `{ success: true }` ואינו מזיז את הציר האופקי** — נמדד
 * על שלושת ערכי ה-`block` (`nearest`, `center`, `start`), עם הסמן מחוץ
 * לתצוגה, `scrollLeft` 0 לפני ו-0 אחרי. הציר האנכי שלו תקין ואינו נוגע כאן.
 * זה גם מה שמונע „סתם לקרוא ל-API של המנוע”: קריאה כזאת עוברת בהצלחה
 * מדווחת ואינה עושה דבר, כלומר היא תיקון שנראה כתוב ואינו קיים.
 *
 * ## מה כן, ולמה זה לגיטימי
 *
 * מיכל הגלילה הוא **שלנו** (`sessions/editor-swap.ts` יוצר אותו, ו-shell.css
 * מגדיר אותו), ולכן גלילה שלו אינה נגיעה בפנימיות המנוע. אותו שיקול בדיוק
 * כמו ב-engine/rtl-line-end.ts.
 *
 * ## האות: `ui.selection.subscribe`, ולא ה-DOM
 *
 * נמדד שאלמנט הסמן (`.sd-v2-local-selection-caret`) **נהרס ונוצר מחדש בכל
 * הקשה** — המזהה שהוצמד לו משתנה, ואין עליו מוטציית `style` להאזין לה —
 * והוא חוזר למקומו רק ‎~200ms אחרי ההקשה (נמדד בדגימה פר-פריים: הסמן נעלם
 * ב-68ms וחוזר ב-201ms). כלומר `keydown` ועוד `requestAnimationFrame` הוא
 * תזמון שנשען על מזל.
 *
 * `ui.selection.subscribe` הוא API ציבורי, והוא יורה בדיוק כשהבחירה זזה.
 * לצדו `ui.selection.getAnchorRect()` מחזיר את מלבן הסמן **בקואורדינטות
 * חלון** — נמדד `left: 777.47` מול `777` שה-DOM דיווח על אותו סמן, ובחלון
 * הצר `615` מול `615`. הוא גם מקדים את ה-DOM: באירוע השני של הקשה הוא כבר
 * מחזיר מלבן בעוד `document.querySelector` על הסמן מחזיר `null`. כלומר
 * הגלילה יוצאת לדרך יחד עם ציור התו, ולא אחריו.
 *
 * ## מה מכוון בכוונה
 *
 * **הציר האופקי בלבד.** הציר האנכי של המנוע נמדד תקין, ומי שיגלול גם אותו
 * רק ייכנס איתו להתגוששות.
 *
 * **בלי אנימציה.** הגלילה היא המשך של ההקלדה, לא ניווט שהמשתמש ביקש; אנימציה
 * בכל תו היא בדיוק התחושה שהמודול נועד להסיר. אותו שיקול כמו ב-`scrollTo`
 * של engine/caret-anchor.ts.
 *
 * **שינוי גודל חלון אינו מפעיל כאן דבר.** גרירת החלון אינה בקשה לזוז — אותה
 * הכרעה שכבר נשמרה ב-`refresh` של engine/zoom-center.ts — והקלדה ראשונה
 * ממילא מחזירה את הסמן לתצוגה.
 *
 * ## מי עוד מאזין לאותו מיכל
 *
 * נמדד: כתיבה ל-`host.scrollLeft` יורה אירוע `scroll` אחד על
 * `.editor-stack__host` עצמו. זה בדיוק המיכל ששני מודולי ההשלמה מאזינים לו —
 * נמדד ש-`editor.container`, שהם מקבלים, ו-`ui.viewport.getHost()`, שהמודול
 * כאן גולל, הם **אותו אלמנט** (`same: true`, שניהם `.editor-stack__host`, ויש
 * אחד כזה ב-DOM). `engine/at-mention-overlay.ts` סוגר עליו את המושב, ו-
 * `engine/book-completion-overlay.ts` מסתיר את ה-ghost — שניהם במשמעות
 * „המשתמש גלל משם”. כאן זה לא נכון: הגלילה היא המשך של ההקלדה, והיא מתרחשת
 * תוך כדי הקלדת השאילתה בחלון צר. `isCaretFollowScroll` מאפשר להם להבחין,
 * ושניהם כבר קוראים לה.
 */

/** מלבן בקואורדינטות חלון, בחלק שנצרך כאן. */
export interface EdgeSpan {
  left: number;
  right: number;
}

/**
 * המרווח שמעבר לסמן, בפיקסלי CSS.
 *
 * סמן שמוצמד בדיוק לקצה התצוגה נראה כאילו הוא עומד ליפול ממנה, ובעברית הוא
 * גם מסתיר את מה שנכתב לפניו. 48px הוא בערך רוחב מילה, כלומר „רואים לאן
 * ממשיכים” — וגם מה שמונע גלילה מחדש בכל תו בודד.
 */
export const CARET_MARGIN_PX = 48;

/**
 * כמה להוסיף ל-`scrollLeft` כדי שהסמן ייכנס לתצוגה, או 0 כשאין מה לעשות.
 *
 * המיכל הוא `ltr`, ולכן `scrollLeft` חיובי בטווח `[0, max]` ודלתא חיובית
 * חושפת ימינה. זה נכון גם למסמך עברי: הכיווניות משפיעה על **היכן** הסמן
 * נמצא, לא על מודל הגלילה של המיכל שמחזיק אותו.
 *
 * המרווח נחתך לרבע מרוחב התצוגה. בלי החיתוך, מאגס צר מ-‎192px היה מקבל שני
 * גבולות מוחלפים (`min > max`) — כלומר כל מיקום סמן מפר את שניהם, וכל תיקון
 * של האחד מפר את השני. עם החיתוך נשאר תמיד חצי מרוחב התצוגה בין הגבולות.
 *
 * מלבן שאינו מספרים אינו נבדק כאן בשורה משלו: `edgeSpanOf` כבר סינן אותו לפני
 * הקריאה, וגם אילו הגיע — `NaN` נכשל בשתי ההשוואות ומחזיר 0 ממילא. נמדד: הסרת
 * שורת הבדיקה שהייתה כאן השאירה 0 מתוך 25 הבדיקות אדומות. בדיקת `width <= 0`
 * שמעליה **כן** מחזיקה אחת מהן, והיא נשארה.
 */
export function horizontalScrollDelta(caret: EdgeSpan, view: EdgeSpan, margin: number): number {
  const width = view.right - view.left;
  if (!Number.isFinite(width) || width <= 0) return 0;

  const pad = Math.max(0, Math.min(margin, width / 4));
  const min = view.left + pad;
  const max = view.right - pad;

  // הקצה הימני קודם: בעברית שם מתחילה השורה, ולכן שם הסמן נמצא אחרי כל
  // `Enter` — המקרה השכיח, ולא הקצה.
  if (caret.right > max) return caret.right - max;
  if (caret.left < min) return caret.left - min;
  return 0;
}

/**
 * המלבן שהמנוע מדווח, בחלק שנצרך כאן.
 *
 * נמדד על superdoc 2.15.0 (17.9.2026): המלבן מגיע עם
 * `{pageIndex, left, right, top, bottom, width, height}` — כלומר `right` הוא
 * שדה אמיתי, ואין צורך לגזור אותו מ-`left + width`. שני מודולי ההשלמה
 * (`engine/at-mention-overlay.ts`, `engine/book-completion-overlay.ts`)
 * מכריזים על אותו מלבן כ-`{left, top, width, height}` — זו הכרזה על מה
 * ש**הם** צורכים, ולא עדות לצורה אחרת.
 */
interface AnchorRectLike {
  left?: unknown;
  right?: unknown;
}

/**
 * מה שנצרך מ-`superdoc.ui`. מוגדר כאן ולא מיובא — ההסבר ב-document-api.ts.
 * גרסת מנוע בלי ההידיות האלה מקבלת ידית ריקה, ולא שגיאה.
 *
 * ‏`placement` בחתימה: זו החתימה שכל ארבעת אתרי הקריאה האחרים בריפו מכריזים
 * עליה — `composables/use-context-menu.ts` ושני מודולי ההשלמה שולחים `'end'`,
 * ו-`engine/doc-metrics.ts` שולח `'start'`. מה שהמנוע עושה איתה נמדד, וכתוב
 * ליד הקריאה עצמה.
 */
export interface CaretVisibilityUi {
  selection?: {
    subscribe?: (listener: () => void) => unknown;
    getAnchorRect?: (input?: {
      placement?: 'start' | 'end' | 'center';
    }) => AnchorRectLike | null | undefined;
  } | null;
}

/**
 * מה שנצרך ממיכל הגלילה.
 *
 * מוגדר מבנית ולא כ-`HTMLElement`, כדי שבדיקה תוכל למסור כפיל: ב-jsdom אין
 * פריסה, ולכן `scrollWidth` של אלמנט אמיתי הוא 0 ו-`getBoundingClientRect`
 * מחזיר אפסים — כלומר כל התנהגות שהיא עוברת שם ירוק.
 */
export interface CaretVisibilityHost {
  scrollWidth: number;
  clientWidth: number;
  clientLeft: number;
  scrollLeft: number;
  getBoundingClientRect: () => { left: number };
}

export interface CaretVisibilityHandle {
  /** מביאה את הסמן לתצוגה עכשיו. מוחזרת כדי שהשער יוכל למדוד קריאה יזומה. */
  follow(): void;
  dispose(): void;
}

export interface CaretVisibilityOptions {
  /** מיכל הגלילה שהמנוע מצייר בתוכו — `paintedHost(editor.ui)`. */
  host: CaretVisibilityHost | null;
  /** ה-`ui` של אותו מסמך. */
  ui: CaretVisibilityUi | null | undefined;
  /** דריסה לבדיקות; בפועל תמיד `CARET_MARGIN_PX`. */
  margin?: number;
}

/** מספר, או `null` כשהמנוע דיווח משהו שאינו מספר. */
function edge(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

/** שני הקצוות של המלבן, או `null` כשאין מהם מספרים. */
function edgeSpanOf(rect: AnchorRectLike | null | undefined): EdgeSpan | null {
  const left = edge(rect?.left);
  const right = edge(rect?.right);
  return left === null || right === null ? null : { left, right };
}

/**
 * המיכלים שהמודול הזיז **בפריים הנוכחי**. ראו `isCaretFollowScroll`.
 *
 * אוסף ולא שדה על הידית: מי ששואל הוא מאזין `scroll` של מודול אחר, ומה שיש לו
 * ביד הוא ה-`target` של האירוע — לא הידית שהותקנה על אותו אלמנט.
 */
const followedNow = new WeakSet<object>();

/** מסמנת שהמודול כתב לכאן עכשיו, ומורידה את הסימן בפריים הבא. */
function markFollowed(host: CaretVisibilityHost): void {
  followedNow.add(host);
  requestAnimationFrame(() => followedNow.delete(host));
}

/**
 * מה שנצרך מאירוע ה-`scroll`: ה-`target` בלבד.
 *
 * מוגדר מבנית ולא כ-`Event`, כדי שבדיקה תוכל למסור כפיל — ב-jsdom אין פריסה
 * ולכן אין שם גלילה אמיתית שתייצר אירוע.
 */
export interface ScrollEventLike {
  target: unknown;
}

/**
 * האם האירוע הזה הוא הגלילה שהמודול כאן ביצע.
 *
 * מאזיני ה-`scroll` על אותו מיכל — `engine/at-mention-overlay.ts:onScroll`
 * ו-`engine/book-completion-overlay.ts:onScroll` — סוגרים את המושב שלהם על כל
 * אירוע, במשמעות „המשתמש גלל משם”. הגלילה כאן היא **המשך של ההקלדה** ולא
 * ניווט: היא מתרחשת בדיוק ברגע שהתו נכנס, ובחלון צר היא מתרחשת תוך כדי
 * הקלדת השאילתה. בלי ההבחנה, רשימת המקורות של „@” נסגרת באמצע ההקלדה.
 *
 * ## למה לא המיקום, למרות שזה מה שנכתב כאן קודם
 *
 * הניסיון הראשון רשם את `scrollLeft` שנכתב והשווה אליו. נמדד ב-Chrome 152
 * שהדפדפן **מעגל** את מה שנכתב, בעוד המודול מחזיק שבר:
 *
 *   | נכתב | נקרא בחזרה | שווה |
 *   |---|---|---|
 *   | 161.47 | 161 | ✗ |
 *   | 176.46875 | 176 | ✗ |
 *   | 176.6 | 177 | ✗ |
 *   | 0.5 | 1 | ✗ |
 *   | 224 | 224 | ✓ |
 *
 * ‏`next` נגזר מ-`caret.right` של `getAnchorRect`, והוא שבר בפועל (נמדד
 * `left: 777.47`, ומלבן בחירה `453.5..539.1`). כלומר ההשוואה נתפסה רק כשהכתיבה
 * נחתכה ל-0 או ל-`overflow` — שני מספרים שלמים — ובתרחיש שבגללו היא נכתבה היא
 * לא נתפסה כלל. וגם כשכן נתפסה, הסימן לא פג לעולם: כל עוד המיכל נח על 0, כל
 * גלילה אנכית בגלגל, כל גלילה של צאצא וכל `resize` נבלעו, והרשימה הפסיקה
 * להיסגר בכלל.
 *
 * ## שני התנאים שהחליפו אותו, ושניהם נמדדו
 *
 * 1. **פריים אחד.** הסימן נדלק בכתיבה ונמחק ב-`requestAnimationFrame` שאחריה.
 *    נמדד שאירוע ה-`scroll` שהכתיבה יוצרת מגיע **לפני** אותו rAF — גם כשהכתיבה
 *    ממשימה רגילה וגם כשהיא מתוך rAF, ובכל הערכים שנוסו (161.47, 88.5, 240):
 *    ‏`scroll → raf1 → raf2`. שני מאזינים שקראו את אותו אירוע קיבלו את אותה
 *    תשובה ובאותו חותם זמן, כלומר הקריאה הראשונה אינה צורכת את הסימן.
 * 2. **ה-`target`.** נמדד שגלילה של **צאצא** נשמעת על שלב ה-capture של המיכל
 *    עם `target` של הצאצא, ושה-`resize` (‏at-mention רושם לו את אותו מאזין)
 *    מגיע עם `target` של ה-`window`. בלי התנאי הזה שניהם היו נבלעים אילו נפלו
 *    באותו פריים.
 *
 * בקרה: גלילת גלגל אמיתית (`Input.dispatchMouseEvent`) נמדדה עם הסימן **כבוי**,
 * גם כשנורתה 23ms אחרי גלילת מעקב — כלומר בפריים שאחריה.
 *
 * וזליגה בין מסמכים אינה אפשרית: הסימן חי פריים אחד ואינו ממתין לכך שמישהו
 * ידרוס אותו, ולכן מיכל שמוסתר ב-`display: none` וחוזר — או ש-`applyPaneScroll`
 * כותב לו `scrollLeft` בעצמו — אינו נושא סימן של המסמך הקודם.
 *
 * ## ומי מודד את כל זה
 *
 * ‏`scripts/qa/caret-visibility-qa.mjs`, בשתי שורות מקצה לקצה: רשימת „@”
 * שנשארת פתוחה בזמן שגלילת המעקב רצה, ובקרה שגלגל עכבר אמיתי **כן** סוגר
 * אותה. שתיהן נבדקו במוטציה — החזרת ההשוואה למיקום מאדימה את הראשונה, והסרת
 * הפקיעה מאדימה את השנייה. שאר השורות בשער נשארו ירוקות בשני המקרים, וזו
 * הסיבה שהן לא הספיקו.
 */
export function isCaretFollowScroll(event: ScrollEventLike): boolean {
  const { target } = event;
  // הצרה בלבד: `WeakSet.has` מקבל אובייקט, ו-`target` מוכרז `unknown`. הפרמטר
  // עצמו אינו מקבל `null`: שני אתרי הקריאה הם מאזיני אירוע, ומאזין מקבל אירוע.
  return typeof target === 'object' && target !== null && followedNow.has(target);
}

/**
 * מחברת את מעקב הסמן למסמך אחד.
 *
 * הביטול מגיע מ-`subscribe` עצמו: נמדד שהוא מחזיר פונקציה. ערך אחר פירושו
 * גרסת מנוע שאיננו יודעים לפרק ממנה, והידית מוותרת על הביטול במקום לנחש.
 *
 * ## ואין כאן `try/catch` סביב `getAnchorRect`
 *
 * ‏`follow` רץ בתוך לולאת המנויים של `ui.selection.subscribe`, ולכן זריקה ממנו
 * הייתה מגיעה למנוע. הייתה כאן עטיפה, והוסרה — אבל לא „כי זה לא יכול לקרות”.
 * הנתיב היחיד שבו `ui.selection` שנתפס בהתקנה חי אחרי שהמסמך מת הוא פירוק
 * המסמך, והוא נמדד (Chrome 152, על הידיות שנתפסו לפני הפירוק):
 *
 *   | מתי | `getAnchorRect.call(sel, {placement:'end'})` | `sel.subscribe(fn)` |
 *   |---|---|---|
 *   | לפני `destroy()` | מלבן | פונקציה |
 *   | **אחרי `destroy()`** | **`null`, בלי זריקה** | פונקציה, בלי זריקה |
 *
 * ‏`null` עובר דרך `edgeSpanOf` ויוצא ביציאה השקטה שכבר קיימת. אילו היה זורק,
 * העטיפה הייתה חוזרת — עם השורה הזאת כטריגר.
 */
export function installCaretVisibility({ host, ui, margin = CARET_MARGIN_PX }: CaretVisibilityOptions): CaretVisibilityHandle {
  const subscribe = ui?.selection?.subscribe;
  const getAnchorRect = ui?.selection?.getAnchorRect;

  const follow = (): void => {
    if (!host || typeof getAnchorRect !== 'function') return;

    // `placement: 'end'` — הקצה **האחרון בסדר המסמך**, וזו החתימה של ארבעת
    // אתרי הקריאה האחרים (שלושה מהם שולחים `'end'`). זה אינו „ראש הבחירה”:
    // בבחירה שנפתחה אחורה הקצה האחרון בסדר המסמך הוא דווקא העוגן הקבוע.
    // נמדד על 2.15.0 שהמנוע **מתעלם** מהארגומנט לגמרי: על בחירה בת 12 תווים
    // (מלבן ברוחב 85.7px) ארבע הקריאות — בלי ארגומנט, `start`, `end`,
    // `center` — החזירו את אותו מלבן בדיוק, זה שעוטף את הבחירה כולה. כלומר
    // הוא אינו משנה התנהגות היום, והוא נשלח כדי שהחתימה תישאר אחת בכל הריפו.
    //
    // הקריאה לפני מדידת הגלישה, וזה **נמדד** ולא הונח (Chrome 152, ~200 אירועי
    // בחירה אמיתיים בכל רוחב, הסדר נבחר אקראית בכל אירוע כדי שלא ייצמד
    // לאירועים שאין בהם מלבן):
    //
    //   | רוחב חלון | `scrollWidth` תחילה | `getAnchorRect` תחילה |
    //   |---|---|---|
    //   | 1100, בלי גלישה | 0.169ms לאירוע | **0.068ms** |
    //   | 600, עם גלישה   | 0.281ms לאירוע | **0.089ms** |
    //
    // מי שקורא ראשון הוא זה שמשלם על הפריסה שההקשה השאירה פתוחה, ו-
    // `getAnchorRect` משלם פחות: אחריו קריאת `scrollWidth` כבר כמעט חינם
    // (0.006ms באירועים שאין בהם מלבן, שהם כמחצית). כלומר הסדר הזה זול יותר
    // גם בחלון רחב, שבו אין לנו מה לעשות בכלל.
    //
    // ונמדד גם שבאירוע הראשון של כל הקשה המנוע עדיין מחזיר `null` (ראו ראש
    // הקובץ). יציאה שקטה, והאירוע הבא הוא זה שיגלול.
    const caret = edgeSpanOf(getAnchorRect.call(ui!.selection, { placement: 'end' }));
    if (caret === null) return;

    // אין גלישה — אין מה לגלול. בחלון רגיל העמוד נכנס במלואו, וזו היציאה של
    // רוב הקריאות.
    const overflow = host.scrollWidth - host.clientWidth;
    if (overflow <= 0) return;

    // תיבת התוכן, ולא קופסת הגבול: פס הגלילה האנכי גורע 15px מהקצה הימני
    // (המיכל `ltr`), וסמן שיושב מתחתיו אינו נראה. `clientWidth` מכיר בו,
    // `getBoundingClientRect().width` אינו.
    const box = host.getBoundingClientRect();
    const viewLeft = box.left + host.clientLeft;
    const delta = horizontalScrollDelta(caret, { left: viewLeft, right: viewLeft + host.clientWidth }, margin);
    if (delta === 0) return;

    host.scrollLeft = Math.min(Math.max(host.scrollLeft + delta, 0), overflow);
    markFollowed(host);
  };

  let unsubscribe: (() => void) | null = null;
  if (typeof subscribe === 'function') {
    const result = subscribe.call(ui!.selection, follow);
    if (typeof result === 'function') unsubscribe = result as () => void;
  }

  return {
    follow,
    dispose() {
      unsubscribe?.();
      unsubscribe = null;
    },
  };
}
