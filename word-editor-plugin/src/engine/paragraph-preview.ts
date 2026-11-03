/**
 * הגאומטריה של פס התצוגה המקדימה בדיאלוג „פסקה” — מיניאטורה של עמודת הטקסט.
 *
 * הפס עונה על השאלה שהמספרים בדיאלוג אינם עונים עליה: „1.25 ס״מ” הוא מספר,
 * ו„רבע מרוחב השורה” הוא מה שהמשתמש בא לדעת. שש ההגדרות שהדיאלוג מחיל —
 * כניסה בהתחלה, כניסה בסוף, שורה ראשונה/תלויה, ריווח לפני, ריווח אחרי ומרווח
 * השורות — כולן יחסים, ולא אחת מהן נראית מהמספר לבדו.
 *
 * ## קנה מידה אחד לשני הצירים, ומאיפה הוא בא
 *
 * הפס אינו סכמה: הוא מיניאטורה **אמיתית** של עמודת הטקסט, בקנה מידה אחד
 * לרוחב ולגובה. קנה המידה נגזר משני מספרים שכבר נמדדים במקום אחר בתוסף, ואף
 * אחד מהם אינו מומצא כאן:
 *
 * 1. **רוחב עמודת הטקסט** — `readPageMargins` (engine/page-setup.ts), אותה
 *    קריאה שהסרגל מצייר לפיה: רוחב הדף פחות שני השוליים של המקטע.
 * 2. **גודל הגופן שבסמן** — `size` של `use-font-controls`, אותו מספר שמוצג
 *    בבורר הגודל ברצועה.
 *
 * ולמה שניהם נדרשים: כניסה של 1 ס״מ היא 6% מעמודה של 16 ס״מ ו-25% מעמודה של
 * 4 ס״מ, ומרווח שורות של „1.5” הוא פי וחצי מגובה הגופן — כלומר בלי רוחב
 * העמודה אין לכניסות סרגל, ובלי גודל הגופן אין למרווח סרגל. פס שמצייר את
 * שניהם בקנה מידה מומצא מראה יחסים שאינם במסמך, וזה גרוע מלא להראות דבר.
 *
 * ## למה אחוזים ו-`cqw`, ולא פיקסלים
 *
 * זו ההכרעה המבנית היחידה כאן, והיא מה שמייתר מדידה של ה-DOM.
 *
 * הרוחב שהפס מקבל בפועל אינו ידוע למודול הזה — הדיאלוג הוא 560px אבל
 * `max-width: calc(100vw - 32px)` מכווץ אותו בחלון צר. מודול שמחשב פיקסלים
 * היה חייב `ResizeObserver` ומדידה חוזרת בכל שינוי גודל.
 *
 * שלוש התכונות שהפס נשען עליהן פותרות את זה בלי מדידה בכלל, מפני שכולן
 * נמדדות ב-CSS מול **הרוחב הפנימי של המכל**:
 *
 * | תכונה | מול מה אחוז נמדד |
 * |---|---|
 * | `margin-inline-start` / `-end` | הרוחב הפנימי של המכל |
 * | `text-indent` | הרוחב הפנימי של המכל |
 * | `margin-block-start` / `-end` | **גם הן** הרוחב הפנימי — כך CSS מגדיר שוליים באחוזים בכל ארבעת הכיוונים, ולא הגובה |
 *
 * כלומר חמש מתוך שש ההגדרות הן אחוז נקי, וכולן באותו מכנה. השישית — גודל
 * הגופן — היא היחידה שאין לה צורת אחוז (`font-size: 40%` הוא אחוז מהגופן
 * של האב ולא מהרוחב), ולכן היא נמסרת ב-`cqw`: אחוז מרוחב מכל השאילתה, אותו
 * מכנה בדיוק. מכאן שקנה המידה נשמר בכל רוחב, וגם בזמן שינוי גודל החלון.
 *
 * מרווח השורות אינו נמסר כלל ביחידות מרחק אלא כמספר חסר יחידה — `line-height`
 * חסר יחידה הוא כפולה של גודל הגופן, וזה בדיוק מה שהוא ב-OOXML כשהכלל הוא
 * `auto` (240 = שורה בודדת). כלומר בציר הזה אין המרה שיכולה לשקר.
 *
 * ## מה מקורב, ומה לא
 *
 * נאמן במדויק: הכניסות, השורה הראשונה/התלויה, הריווח לפני ואחרי, ומרווח
 * שורות בכלל `auto` או `exact`.
 *
 * מקורב אחד בלבד: הכלל `atLeast` (‏„לפחות”). Word משווה את הגובה המבוקש לגובה
 * השורה הטבעי של הגופן ולוקח את הגדול, והגובה הטבעי הזה אינו ידוע מכאן —
 * הוא תלוי בגופן, ובפס אין את הגופן של המסמך. הפס מצייר את **הרצפה
 * המבוקשת** כשהיא גדולה מתיבת ה-em, ואת תיבת ה-em כשלא; כלומר בשורה צפופה
 * במיוחד הוא עשוי להראות מרווח קטן ממה ש-Word יצייר, ולעולם לא גדול ממנו.
 *
 * ## ולמה פסי שורה ולא אותיות
 *
 * זו אינה בחירה גרפית אלא מה שהפס נמדד לפיו. **מספר השורות הנראות חייב
 * להיות קבוע**, מפני שכל מה שהפס מראה הוא מרחקים בין שורות ובין פסקות:
 * פסקה של שלוש מילים הייתה מצטמצמת לשורה אחת, ואז מרווח השורות — אחת משש
 * ההגדרות — לא היה נראה בכלל, בדיוק בפסקות שבהן משנים אותו. Word עצמו מצייר
 * בתצוגה המקדימה שלו גוף טקסט בגודל אחיד ולא את הפסקה כפי שהיא, ומאותו טעם.
 *
 * ומכיוון שהמספר קבוע, טקסט אמיתי היה ממילא נחתך או מרופד. שתי הדרכים לקבל
 * אותו גם אינן חינם: `readSelectionText` (engine/font-preview.ts) מחזיר את
 * ה**סימון**, והדיאלוג נפתח כמעט תמיד על סמן מכווץ — כלומר מחרוזת ריקה;
 * ו-`paragraph.inlines` שב-`doc.get()` הוא מודל שצורתו לא נמדדה כאן. פסי
 * שורה אינם מתיימרים על התוכן, ומראים את כל שש ההגדרות במדויק.
 */
import { MIN_TEXT_WIDTH_TWIPS } from './ruler-geometry';
import { TWIPS_PER_PT, type LineSpacingRule } from './paragraph-format';

/**
 * מספר השורות שהפס מצייר לפסקה שנערכת, ולכל אחת מהשכנות.
 *
 * ארבע לפסקה שנערכת ולא שתיים: מרווח שורות נראה במרחק **בין** השורות, ושתי
 * שורות מציגות מרחק אחד — כלומר גם ריווח „לפני” של הפסקה נראה בדיוק כמוהו.
 * עם ארבע יש שלושה מרחקים פנימיים מול שני מרחקים חיצוניים, וההבדל בין
 * „מרווח שורות” לבין „ריווח לפני/אחרי” מובחן.
 *
 * שתיים לשכנות: הן קיימות רק כדי שיהיה למרחק „לפני” ו„אחרי” גוף להיפרד ממנו,
 * ואין להן שום הגדרה משלהן להראות.
 */
export const PREVIEW_TARGET_LINES = 4;
export const PREVIEW_NEIGHBOUR_LINES = 2;

/** מצב הדיאלוג החי — לא התצלום. הפס מראה את מה שהמשתמש ערך עכשיו. */
export interface ParagraphPreviewSource {
  /** רוחב עמודת הטקסט של המקטע. `0` = אין גיאומטריה, ואז אין פס. */
  textWidthTwips: number;
  /** גודל הגופן שבסמן, בנקודות — הסרגל של הציר האנכי. */
  fontSizePt: number;
  /** הכניסה בצד ההתחלה של הפסקה (‏„לפני טקסט”). */
  startTwips: number;
  /** הכניסה בצד הסוף (‏„אחרי טקסט”). */
  endTwips: number;
  /** „מיוחד” — בדיוק הבורר שבדיאלוג, ולכן שני המצבים אינם יכולים להתקיים יחד. */
  special: 'none' | 'firstLine' | 'hanging';
  /** מידת ה„מיוחד”. נקראת רק כש-`special` אינו `'none'`. */
  amountTwips: number;
  beforeTwips: number;
  afterTwips: number;
  /** ב-`auto` זו כפולה×240; ב-`exact`/`atLeast` זה גובה שורה במרחק. */
  lineTwips: number;
  lineRule: LineSpacingRule;
}

/** מה שה-CSS של הפס מקבל. כל האחוזים מול אותו מכנה — ראו הערת הראש. */
export interface ParagraphPreviewGeometry {
  /** `font-size`, ביחידות `cqw` — אחוז מרוחב מכל השאילתה. */
  fontCqw: number;
  /** `line-height` חסר יחידה — כפולה של גודל הגופן. */
  lineHeight: number;
  /** `margin-inline-start` באחוזים. */
  startPct: number;
  /** `margin-inline-end` באחוזים. */
  endPct: number;
  /** `text-indent` באחוזים: חיובי לשורה ראשונה, שלילי לתלויה, `0` ל„ללא”. */
  firstLinePct: number;
  /** `margin-block-start` באחוזים — של הרוחב, כך CSS מגדיר. */
  beforePct: number;
  /** `margin-block-end` באחוזים. */
  afterPct: number;
  /**
   * האם שתי הכניסות משאירות רוחב שאפשר לכתוב בו.
   *
   * הרצפה היא `MIN_TEXT_WIDTH_TWIPS` של הסרגל — ס״מ אחד — ולא מספר חדש:
   * `clampIndent` (engine/ruler-geometry.ts) כבר אוסר לגרור סמן כניסה מעבר
   * לה, מאותו נימוק בדיוק („כניסות שנפגשות יוצרות פסקה שאי אפשר לכתוב בה”).
   * הדיאלוג עצמו **אינו** אוסר את זה — שני השדות מוגבלים ל-55.87 ס״מ כל אחד
   * ואינם מוצלבים — ולכן זה מצב שאפשר להגיע אליו דרכו, והפס אומר אותו במקום
   * לצייר פסקה ברוחב אפס בשקט.
   */
  leavesRoom: boolean;
}

/** שתי ספרות מעבר לאחוז השלם. מעבר לזה המספר מספר על שגיאות עיגול ולא על המסמך. */
function round(value: number): number {
  return Math.round(value * 100) / 100;
}

/**
 * הגאומטריה, או `null` כשאין ממה לגזור קנה מידה.
 *
 * `null` הוא מצב רגיל ולא כשל, בדיוק כמו `readPageMargins` שמזין אותו: מסמך
 * שעדיין נטען אינו מחזיר מקטע, ואז אין רוחב עמודה. הדיאלוג פשוט אינו מצייר
 * פס — כל שאר הפקדים בו עובדים.
 */
export function paragraphPreviewGeometry(
  source: ParagraphPreviewSource,
): ParagraphPreviewGeometry | null {
  const { textWidthTwips, fontSizePt } = source;
  if (!Number.isFinite(textWidthTwips) || textWidthTwips <= 0) return null;
  if (!Number.isFinite(fontSizePt) || fontSizePt <= 0) return null;

  const fontTwips = fontSizePt * TWIPS_PER_PT;
  const pct = (twips: number): number =>
    round(((Number.isFinite(twips) ? twips : 0) / textWidthTwips) * 100);

  const amount = Number.isFinite(source.amountTwips) ? source.amountTwips : 0;
  const firstLineTwips =
    source.special === 'firstLine' ? amount : source.special === 'hanging' ? -amount : 0;

  const startTwips = Number.isFinite(source.startTwips) ? source.startTwips : 0;
  const endTwips = Number.isFinite(source.endTwips) ? source.endTwips : 0;

  return {
    fontCqw: round((fontTwips / textWidthTwips) * 100),
    lineHeight: previewLineHeight(source.lineTwips, source.lineRule, fontTwips),
    startPct: pct(startTwips),
    endPct: pct(endTwips),
    firstLinePct: pct(firstLineTwips),
    beforePct: pct(source.beforeTwips),
    afterPct: pct(source.afterTwips),
    leavesRoom: textWidthTwips - startTwips - endTwips >= MIN_TEXT_WIDTH_TWIPS,
  };
}

/**
 * `line-height` חסר יחידה, לפי שלושת הכללים של OOXML.
 *
 * * **`auto`** — `w:line` הוא כפולה×240 („שורה בודדת”), ולכן החלוקה ב-240 היא
 *   ההמרה המדויקת ולא קירוב. זה גם למה הכלל הזה אינו נוגע בגודל הגופן בכלל.
 * * **`exact`** — `w:line` הוא מרחק, וגובה השורה הוא היחס שלו לגופן.
 * * **`atLeast`** — אותו חישוב, עם רצפה של תיבת ה-em. ראו „מה מקורב” בהערת
 *   הראש: הגובה הטבעי של הגופן אינו ידוע כאן, ולכן הרצפה היא המבנה (‏`1`)
 *   ולא קבוע טיפוגרפי מומצא.
 *
 * ערך פסול או אפס נופל ל-`1`: `line-height: 0` היה מקפל את כל השורות אחת על
 * השנייה, כלומר פס שאינו מראה דבר.
 */
function previewLineHeight(
  lineTwips: number,
  rule: LineSpacingRule,
  fontTwips: number,
): number {
  const line = Number.isFinite(lineTwips) && lineTwips > 0 ? lineTwips : 0;
  if (line === 0) return 1;
  if (rule === 'auto') return round(line / 240);
  const ratio = round(line / fontTwips);
  if (ratio <= 0) return 1;
  return rule === 'atLeast' ? Math.max(1, ratio) : ratio;
}
