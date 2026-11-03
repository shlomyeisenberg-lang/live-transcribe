/**
 * קריאת עיצוב התו של הבחירה הנוכחית — מה שהמנוע מדווח, בצורה שהמתג של
 * הקיצורים האישיים צורך (ui/shortcuts/preset-toggle.ts).
 *
 * ## למה מודול, ולא שלוש שורות באתר הקריאה
 *
 * כאן יושבת הכרעה שאי אפשר להסיק מהקוד: **מה פירושו של „המנוע לא דיווח
 * צבע”.** הוא מקפל „בחירה מעורבת” ו„אין צבע ישיר” לאותה תשובה, ולכן התשובה
 * הזאת נקראת אחרת לפי צורת הבחירה — ראו `readColor`. הכרעה כזאת צריכה בדיקה,
 * ובדיקה צריכה פונקציה שאפשר להזריק לה מתאם מזויף.
 *
 * המודול קורא **רק** דרך `CommandAdapter`, כלומר דרך אותו מסלול שכל פקד
 * ברצועה קורא בו. אין כאן שאילתה על ה-DOM של המנוע ואין גישה ל-Document API.
 */
import type { CommandAdapter } from './command-adapter';
import type { CommandId } from './capabilities';
import { parseColor, parseFontFamily, parseFontSizePt } from './payloads';
import type { FormatReading } from '../ui/shortcuts/format-preset';
import { TOGGLE_COMMANDS, TOGGLE_FIELDS } from '../ui/shortcuts/format-preset';

/** מה שנצרך מהמתאם. `Pick` כדי שבדיקה תכפול שני שדות ולא את כולו. */
export type FormatReadingSource = Pick<CommandAdapter, 'has' | 'getState'>;

/** צורת הבחירה, כפי ש-`readout-hold` מדווח אותה. */
export interface ReadingSelection {
  /** סמן מכווץ (או שאין בחירה בכלל). */
  empty: boolean;
  /** הקריאה של המנוע התיישבה. */
  settled: boolean;
}

function stateValue(source: FormatReadingSource, id: CommandId): unknown {
  return source.has(id) ? source.getState(id).value : undefined;
}

/**
 * צבע (או הדגשה).
 *
 * שלוש תשובות ולא שתיים, וזה כל טעמו של המודול:
 *
 * - המנוע דיווח צבע תקין → זה הצבע.
 * - המנוע לא דיווח, והבחירה היא **סמן מיושב** → „אין צבע ישיר”, כלומר `null`.
 *   לסמן אין „מעורב”: הוא עומד בהקשר עיצוב אחד ויחיד, ולכן שתיקה שם אינה
 *   דו-משמעית. בלי ההסקה הזאת המקרה הנפוץ ביותר היה נשבר — טקסט בצבע ברירת
 *   המחדל, קיצור שצובע אותו, ולחיצה שנייה שאינה יודעת לאן לחזור.
 * - המנוע לא דיווח, והבחירה היא טווח או שטרם התיישבה → **לא נודע**
 *   (`undefined`). על טווח מעורב אין צבע אחד לשחזר, וניחוש „אין צבע” היה
 *   מנקה את הצבע של החלק הצבוע.
 */
function readColor(
  source: FormatReadingSource,
  id: CommandId,
  selection: ReadingSelection,
): string | null | undefined {
  const reported = parseColor(stateValue(source, id));
  if (reported !== null) return reported;
  return selection.empty && selection.settled ? null : undefined;
}

/**
 * הקריאה. `settled: false` אינו חוסם את הקריאה כולה — גופן וגודל תקפים גם
 * לפני שהבחירה התיישבה — אלא רק את ההסקה על צבע שלא דווח.
 */
export function readFormat(
  source: FormatReadingSource,
  selection: ReadingSelection,
): FormatReading {
  const family = parseFontFamily(stateValue(source, 'font-family'));
  const size = parseFontSizePt(stateValue(source, 'font-size'));

  const reading: FormatReading = {
    // המתגים תמיד נודעים: המנוע מדווח `active: boolean` ואין לו ערך שלישי.
    // מתג שהמנוע אינו מכיר נקרא כ„כבוי” — פקודה שאינה קיימת אינה יכולה
    // להיות דלוקה.
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
  };

  if (family !== null) reading.fontFamily = family;
  if (size !== null) reading.fontSizePt = size;

  const color = readColor(source, 'text-color', selection);
  if (color !== undefined) reading.color = color;
  const highlight = readColor(source, 'highlight-color', selection);
  if (highlight !== undefined) reading.highlight = highlight;

  for (const field of TOGGLE_FIELDS) {
    const id = TOGGLE_COMMANDS[field];
    reading[field] = source.has(id) && source.getState(id).active === true;
  }

  return reading;
}
