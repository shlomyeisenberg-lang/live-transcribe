/**
 * „רוחב עמוד”: חישוב אחוז ההתאמה כאן, והחלתו דרך פקודת `zoom` הרגילה.
 *
 * ## למה איננו שולחים את `zoom-fit-width` של המנוע
 *
 * נמדד חי (CDP על ה-dist הארוז, superdoc@2.8.0): מסלול `setZoomMode('fit-width')`
 * מפעיל מדידה חוזרת של רוחב המסמך **אחרי** שהזום הוחל, ורוחב העמוד באותה מדידה
 * מדורג לפי הזום הנוכחי — ב-bundle של המנוע: "page metrics viewport coords
 * scale with the same zoom value". התוצאה היא לולאת משוב: כל הערכה מכפילה את
 * הזום ביחס ההתאמה שוב. נמדד: לחיצה אחת על „רוחב עמוד” בחלון שבו ההתאמה
 * האמיתית היא ~93% הורידה את המסמך לרצפה (10%) תוך פחות משנייה, והעמוד
 * המרונדר לא נע בגבולות שהמשתמש מזהה.
 *
 * לכן הכפתור מחשב את האחוז בעצמו — רוחב המאגס שלנו חלקי רוחב עמוד המסמך
 * ב-100%, מהמידות הקנוניות של `sections.list` (אינצ'ים; ראו engine/print.ts) —
 * ומחיל אותו דרך פקודת `zoom`. `setZoom` של המנוע מעביר ל-zoomMode `manual`
 * מטעמי עיצוב שלו, ובכך מנטרל את הלולאה. מסלול הכתיבה נשאר יחיד: האדפטר.
 */
import type { ZoomState } from './zoom';
import { clampZoom } from './zoom';
import { readPrintPageSize, type PrintTarget } from './print';

/**
 * המאגס שלנו שבתוכו המנוע מרנדר.
 *
 * זהו **קלס שלנו** ולא של המנוע — הכלל שאוסר `querySelector` אל DOM פנימי
 * של SuperDoc (`.sd-*`) אינו נפגע: אנחנו מודדים את המקום שאנחנו עצמנו נתנו
 * למנוע, שהוא בדיוק מה ש„רוחב עמוד” אמור להתאים אליו.
 */
export const EDITOR_STACK_SELECTOR = '.editor-stack';

/** פיקסלי CSS לאינץ' — ההמרה ממידות ה-`sections.list` לגיאומטריית המסך. */
export const CSS_PX_PER_INCH = 96;

/**
 * מיכל העימוד של המנוע, שנושא את `scale()` של הזום (נמדד ב-bundle של
 * superdoc@2.8.0). הקבוע כאן ולא ליטרל ב-CSS בלבד: styles/shell.css מכוון
 * אל המחלקה הזאת כדי לתקן את מרכוז הזום ב-RTL (`transform-origin`), ושער
 * ההיגיינה (tests/unit/css-hygiene.test.ts) מחייב שכל מחלקה בסלקטור גלובלי
 * תופיע בקוד — הגדרה מפורשת היא מה שמוכיח שהמחלקה קיימת ולא נכתבה בטעות.
 */
export const ZOOM_LAYOUT_CLASS = 'presentation-editor';

/**
 * מיכל הגלילה שבתוך המאגס — האלמנט שנושא בפועל את פס הגלילה האנכי.
 * הקלס הוא שלנו (sessions/editor-swap.ts).
 */
export const EDITOR_HOST_SELECTOR = '.editor-stack__host';

/**
 * הרוחב הפנוי לעמוד בפיקסלי CSS, או `0` כשאין מה למדוד.
 *
 * `clientWidth` ולא `getBoundingClientRect`: הראשון מוציא את פס הגלילה,
 * והתאמה שמתעלמת ממנו הייתה משאירה גלילה אופקית של בדיוק רוחב פס הגלילה.
 *
 * ונמדד על מיכל הגלילה ולא על ה-`<main>`: פס הגלילה יושב על הראשון בלבד
 * (`.editor-stack` הוא `position: relative` והמיכל הוא `inset: 0` שגולל).
 * נמדד בחלון 1440: `.editor-stack` מדווח 1440 והמיכל 1425 — כלומר „רוחב
 * עמוד” היה מכוון לעמוד רחב ב-15px מהמקום שיש לו, ומקבל בדיוק את הגלילה
 * האופקית שההערה שמעל באה למנוע. נפילה בחן ל-`<main>` כשאין עדיין מיכל.
 */
export function editorStackWidth(root: Document = document): number {
  const el = root.querySelector(EDITOR_HOST_SELECTOR) ?? root.querySelector(EDITOR_STACK_SELECTOR);
  return el && el.clientWidth > 0 ? Math.round(el.clientWidth) : 0;
}

/**
 * אחוז ההתאמה, מוגבל לגבולות האפקטיביים, או `null` כשאין מדידה תקינה.
 *
 * החישוב הוא יחס גלם ולא ניחוש: חלון 1480px על A4 (8.268 אינץ' ≈ 794px)
 * נותן 186% — כלומר בחלון רחב החישוב **מציע הגדלה**, בדיוק כמו ב-Word.
 * הגבולות מגיעים מ-`zoomBounds` (engine/zoom.ts) — אותו מקור של הסליידר
 * בשורת המצב — וכוללים את הרחבת התקרה להיקף Word (500%): ה-max שהמנוע
 * מדווח הוא גבול ה-fit-width שלו ולא מגבלת זום ידני, ולכן חלון רחב מגדיל
 * מעבר ל-100% במקום להיעצר שם.
 *
 * עיגול **כלפי מטה** ולא לשלם הקרוב: האחוז הוא מספר שלם, וחצי אחוז מעל
 * ההתאמה המדויקת הוא עמוד רחב מהמאגס — כלומר גלילה אופקית של כמה פיקסלים
 * מיד אחרי לחיצה על „רוחב עמוד”. נמדד בחלון 1440: ההתאמה היא 179.5%,
 * ועיגול לשלם הקרוב נתן 180% ועמוד רחב ב-4px מהמקום שיש לו. רצועה אפורה
 * של 4px היא המחיר, והיא הצד הנכון לטעות בו.
 */
export function computeFitPercent(
  containerPx: number,
  pageInches: number,
  bounds: Pick<ZoomState, 'min' | 'max'>,
): number | null {
  if (!(containerPx > 0)) return null;
  if (!(pageInches > 0) || !Number.isFinite(pageInches)) return null;
  const exact = (containerPx / (pageInches * CSS_PX_PER_INCH)) * 100;
  return clampZoom(Math.floor(exact), bounds.min, bounds.max);
}

/**
 * אחוז ההתאמה לפי המסמך הפתוח, או `null` כשרוחב העמוד לא נקרא.
 *
 * קורא את מידות הדף מ-`sections.list` דרך אותו מסלול מנוטר ומסונן של
 * ההדפסה (`readPrintPageSize`): אינץ' עם סינון תשובות שאינן תשובה, ולעולם
 * בלי זריקה — כשל קריאה הוא `null`, והכפתור ידווח על כך ולא ינחש.
 */
export async function fitWidthPercent(
  host: PrintTarget,
  containerPx: number,
  bounds: Pick<ZoomState, 'min' | 'max'>,
): Promise<number | null> {
  if (!(containerPx > 0)) return null;
  const size = await readPrintPageSize(host);
  if (!size) return null;
  return computeFitPercent(containerPx, size.widthIn, bounds);
}