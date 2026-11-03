/**
 * „ערכת עיצוב” — אוסף של תכונות תו שקיצור אישי מחיל בלחיצה אחת.
 *
 * זה הצד ההצהרתי בלבד: מה מותר להיות בערכה, איך ערכה שנקראה מהאחסון מנוקה,
 * ואיזו רשימת פקודות היא מתורגמת אליה. **אין כאן שום קריאה למנוע** — ההרצה
 * היא של המעטפת, דרך אותו `CommandAdapter` שכל כפתור ברצועה עובר בו. הפרדה
 * זאת היא מה שמאפשר לבדוק את כל ההכרעות כאן כקלט-פלט, בלי להרים מסמך.
 *
 * ## למה שדה חסר ושדה ריק אינם אותו דבר
 *
 * שדה שאינו בערכה = „אל תיגע”. שדה שערכו `null` (צבע והדגשה) = „נקה”. שתי
 * המשמעויות היו נופלות לאותו `undefined` אם הערכה הייתה מחזיקה ערכי ברירת
 * מחדל, ואז „ערכה שרק מגדילה את הגופן” הייתה מוחקת את הצבע של הטקסט. זה גם
 * החוזה של המנוע עצמו: `colorPayload(null)` הוא מסלול הניקוי, ומחרוזת ריקה
 * נדחית שם (ראו engine/payloads.ts).
 */
import type { CommandId } from '../../engine/capabilities';
import {
  colorPayload,
  fontFamilyPayload,
  fontSizePayload,
  parseColor,
  parseFontFamily,
  parseFontSizePt,
} from '../../engine/payloads';

/**
 * ערכת עיצוב. כל השדות אופציונליים, וערכה שכולה ריקה אינה חוקית — קיצור
 * שאינו מחיל דבר הוא קיצור שנראה שבור.
 */
export interface FormatPreset {
  fontFamily?: string;
  fontSizePt?: number;
  /** `null` = „ללא צבע” (אוטומטי). */
  color?: string | null;
  /** `null` = „ללא הדגשה”. */
  highlight?: string | null;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
}

/** המתגים — השדות שהמנוע מדווח עליהם ב-`active` ולא ב-`value`. */
export const TOGGLE_FIELDS = ['bold', 'italic', 'underline', 'strikethrough'] as const;
export type ToggleField = (typeof TOGGLE_FIELDS)[number];

/** הפקודה של כל מתג. אותם מזהים בדיוק שהרצועה משתמשת בהם. */
export const TOGGLE_COMMANDS: Record<ToggleField, CommandId> = {
  bold: 'bold',
  italic: 'italic',
  underline: 'underline',
  strikethrough: 'strikethrough',
};

/** התוויות העבריות, במקום אחד: הדיאלוג, הסיכום ושורת המצב מדברים אותו קול. */
export const FIELD_TITLES = {
  fontFamily: 'גופן',
  fontSizePt: 'גודל',
  color: 'צבע',
  highlight: 'הדגשה',
  bold: 'מודגש',
  italic: 'נטוי',
  underline: 'קו תחתון',
  strikethrough: 'קו חוצה',
} as const;

export type PresetField = keyof typeof FIELD_TITLES;

/** הסדר שבו שדות מוצגים ומוחלים. */
export const PRESET_FIELDS: readonly PresetField[] = [
  'fontFamily',
  'fontSizePt',
  'color',
  'highlight',
  ...TOGGLE_FIELDS,
];

/* ------------------------------------------------------------------ */
/* נורמליזציה                                                          */
/* ------------------------------------------------------------------ */

/**
 * ערכה מערך גולמי שנקרא מהאחסון או מטופס.
 *
 * כל שדה עובר את **אותו** פרסר שהמנוע נשען עליו (engine/payloads.ts), ולא
 * בדיקת טיפוס משלנו: ערכה שנשמרה עם `fontSizePt: 0` או `color: 'אדום'` הייתה
 * מגיעה לוולידטור של המנוע, נדחית שם סגור, והמשתמש היה רואה קיצור ששותק.
 * שדה שאינו עובר פשוט אינו נכנס לערכה — וזה גלוי בדיאלוג, כי הוא מציג את מה
 * שנשמר.
 */
export function normalizePreset(raw: unknown): FormatPreset {
  if (!raw || typeof raw !== 'object') return {};
  const source = raw as Record<string, unknown>;
  const preset: FormatPreset = {};

  const family = parseFontFamily(source.fontFamily);
  if (family !== null) preset.fontFamily = family;

  const size = parseFontSizePt(source.fontSizePt);
  if (size !== null) preset.fontSizePt = size;

  // `null` מפורש שורד, כל שאר מה שאינו צבע תקין נושר. `'color' in source`
  // הוא מה שמפריד „נשמר ניקוי” מ„לא נשמר כלום”.
  if ('color' in source) {
    if (source.color === null) preset.color = null;
    else {
      const color = parseColor(source.color);
      if (color !== null) preset.color = color;
    }
  }
  if ('highlight' in source) {
    if (source.highlight === null) preset.highlight = null;
    else {
      const highlight = parseColor(source.highlight);
      if (highlight !== null) preset.highlight = highlight;
    }
  }

  for (const field of TOGGLE_FIELDS) {
    const value = source[field];
    if (typeof value === 'boolean') preset[field] = value;
  }

  return preset;
}

/** האם הערכה אינה מחילה דבר. */
export function isEmptyPreset(preset: FormatPreset): boolean {
  return PRESET_FIELDS.every((field) => preset[field] === undefined);
}

/** השדות שהערכה קובעת, בסדר התצוגה. */
export function presetFields(preset: FormatPreset): PresetField[] {
  return PRESET_FIELDS.filter((field) => preset[field] !== undefined);
}

/* ------------------------------------------------------------------ */
/* תצוגה                                                              */
/* ------------------------------------------------------------------ */

/** הערך של שדה אחד כטקסט עברי. */
export function fieldText(preset: FormatPreset, field: PresetField): string {
  switch (field) {
    case 'fontFamily':
      return preset.fontFamily ?? '';
    case 'fontSizePt':
      return preset.fontSizePt === undefined ? '' : `${preset.fontSizePt} נק'`;
    case 'color':
      return preset.color === null ? 'צבע אוטומטי' : `צבע ${preset.color}`;
    case 'highlight':
      return preset.highlight === null ? 'ללא הדגשה' : `הדגשה ${preset.highlight}`;
    default:
      // מתג: „מודגש” כשהוא נדלק, „ללא מודגש” כשהוא נכבה. ערכה יכולה לבקש את
      // שני הכיוונים, ולכן „מודגש” לבדו היה מתאר גם כיבוי.
      return preset[field] === true ? FIELD_TITLES[field] : `ללא ${FIELD_TITLES[field]}`;
  }
}

/** סיכום הערכה בשורה אחת, לרשימה שבדיאלוג ולשורת המצב. */
export function presetSummary(preset: FormatPreset): string {
  return presetFields(preset)
    .map((field) => fieldText(preset, field))
    .join(', ');
}

/* ------------------------------------------------------------------ */
/* מה שהמנוע מדווח                                                     */
/* ------------------------------------------------------------------ */

/**
 * הקריאה מהמנוע על הבחירה הנוכחית. `undefined` בשדה פירושו **לא נודע** —
 * המנוע מקפל „בחירה מעורבת” ו„עוד לא נפתר” לאותו `undefined` ב-`value`
 * (ראו engine/readout-hold.ts), ואין לנו דרך להבחין ביניהם.
 *
 * המתגים אינם יכולים להיות „לא נודע”: המנוע מדווח עליהם `active: boolean`
 * ואין לו ערך שלישי. על בחירה מעורבת — חלק מודגש וחלק לא — הוא מדווח `false`,
 * ולכן „החזרה לעיצוב הקודם” על בחירה כזאת מיישרת את המתג ולא משחזרת את
 * הפסיפס. זו מגבלה של מה שהמנוע חושף, והיא מתועדת גם בהודעה למשתמש.
 */
export interface FormatReading {
  fontFamily?: string;
  fontSizePt?: number;
  color?: string | null;
  highlight?: string | null;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
}

/**
 * הערכה ש„מחזירה” את הקריאה — כלומר מה שצריך להחיל כדי לחזור למצב שנקרא,
 * מוגבל לשדות שהערכה הנכנסת נוגעת בהם.
 *
 * ההגבלה לשדות של הערכה היא העיקר: קיצור שמשנה רק גופן וגודל לא ישחזר גם צבע
 * שהמשתמש שינה בינתיים בעצמו. שדה שהערכה נוגעת בו ושהקריאה אינה יודעת עליו
 * דבר פשוט אינו נכנס — עדיף להשאיר אותו כפי שהוא מאשר לנחש ערך.
 */
export function reverseOf(reading: FormatReading, preset: FormatPreset): FormatPreset {
  const previous: FormatPreset = {};
  for (const field of presetFields(preset)) {
    switch (field) {
      case 'fontFamily':
        if (reading.fontFamily !== undefined) previous.fontFamily = reading.fontFamily;
        break;
      case 'fontSizePt':
        if (reading.fontSizePt !== undefined) previous.fontSizePt = reading.fontSizePt;
        break;
      case 'color':
        if (reading.color !== undefined) previous.color = reading.color;
        break;
      case 'highlight':
        if (reading.highlight !== undefined) previous.highlight = reading.highlight;
        break;
      default:
        previous[field] = reading[field];
    }
  }
  return previous;
}

/**
 * השדות שהערכה קובעת ושהקריאה לא ידעה לדווח — כלומר מה ש„חזרה לעיצוב הקודם”
 * לא תוכל לשחזר. מוצג למשתמש, כדי שהוא לא יחשב שהקיצור שכח.
 */
export function unreadableFields(reading: FormatReading, preset: FormatPreset): PresetField[] {
  const reverse = reverseOf(reading, preset);
  return presetFields(preset).filter((field) => reverse[field] === undefined);
}

/**
 * האם הטקסט שהסמן עומד בו כבר נראה כמו הערכה.
 *
 * זה מה שמכריע אם לחיצה שנייה מחזירה או מחילה שוב, ולכן זו ההכרעה החשובה
 * ביותר במודול: אחרי שהמשתמש הפעיל את הקיצור והמשיך לעבוד — עבר לפסקה אחרת,
 * שינה גופן ביד, פתח מסמך אחר — „העיצוב הקודם” ששמרנו אינו קודם לשום דבר,
 * והחזרתו הייתה מחילה על הטקסט הנוכחי עיצוב שלא היה לו מעולם.
 *
 * שדה שאינו נודע נחשב **כלא מתאים**: אם אי אפשר לאמת שהערכה בתוקף, אין
 * להתייחס אליה כאל פעילה.
 */
export function readingMatchesPreset(reading: FormatReading, preset: FormatPreset): boolean {
  return presetFields(preset).every((field) => {
    switch (field) {
      case 'fontFamily':
        // המנוע מדווח שם גופן כפי שהוא במסמך, וההשוואה חסינה לרישיות
        // ולרווחים — „David” ו-„david ” הם אותו גופן.
        return (
          reading.fontFamily !== undefined &&
          reading.fontFamily.trim().toLowerCase() === preset.fontFamily!.trim().toLowerCase()
        );
      case 'fontSizePt':
        return reading.fontSizePt === preset.fontSizePt;
      case 'color':
        return reading.color !== undefined && reading.color === preset.color;
      case 'highlight':
        return reading.highlight !== undefined && reading.highlight === preset.highlight;
      default:
        return reading[field] === preset[field];
    }
  });
}

/* ------------------------------------------------------------------ */
/* מה שמורץ                                                           */
/* ------------------------------------------------------------------ */

/**
 * צעד אחד בהחלת ערכה.
 *
 * שני סוגים, מפני שהמנוע חושף שני סוגי פקודות: `value` היא פקודה שמקבלת
 * payload ומציבה ערך, ו-`toggle` היא פקודה **שמחליפה מצב** — `bold` אינו
 * „הדגש” אלא „הפוך הדגשה”. ערכה שמבקשת „מודגש” על טקסט שכבר מודגש חייבת
 * לכן **לא** להריץ אותה, ומי שיודע את המצב הנוכחי הוא הקורא. הפרדת הסוגים
 * כאן היא מה שמונע את הבאג הזה בלי שהמודול יגע במנוע.
 */
export type PresetStep =
  | { kind: 'value'; command: CommandId; payload: unknown }
  | { kind: 'toggle'; command: CommandId; to: boolean };

/**
 * הצעדים להחלת הערכה, בסדר. `payload` נבנה בבוני ה-payload של המנוע ולא ביד
 * — הם החוזה מול הוולידטורים, ובדיקת החוזה מריצה אותם מולם.
 */
export function presetSteps(preset: FormatPreset): PresetStep[] {
  const steps: PresetStep[] = [];

  if (preset.fontFamily !== undefined) {
    const payload = fontFamilyPayload(preset.fontFamily);
    if (payload !== null) steps.push({ kind: 'value', command: 'font-family', payload });
  }
  if (preset.fontSizePt !== undefined) {
    const payload = fontSizePayload(preset.fontSizePt);
    if (payload !== null) steps.push({ kind: 'value', command: 'font-size', payload });
  }
  if (preset.color !== undefined) {
    steps.push({ kind: 'value', command: 'text-color', payload: colorPayload(preset.color) });
  }
  if (preset.highlight !== undefined) {
    steps.push({
      kind: 'value',
      command: 'highlight-color',
      payload: colorPayload(preset.highlight),
    });
  }
  for (const field of TOGGLE_FIELDS) {
    const wanted = preset[field];
    if (wanted !== undefined) {
      steps.push({ kind: 'toggle', command: TOGGLE_COMMANDS[field], to: wanted });
    }
  }

  return steps;
}
