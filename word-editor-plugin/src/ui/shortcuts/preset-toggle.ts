/**
 * המתג של הקיצור האישי: לחיצה מחילה ערכת עיצוב, לחיצה נוספת מחזירה את מה
 * שהיה לפניה.
 *
 * ## ההכרעה שכל הפיצ'ר תלוי בה
 *
 * „לחיצה נוספת מחזירה” נשמע כמו דגל בוליאני, וזה בדיוק מה שהוא **לא**. עיצוב
 * הוא תכונה של הטקסט, לא של הקיצור: בין הלחיצה הראשונה לשנייה המשתמש זז
 * לפסקה אחרת, שינה גופן ביד, ביטל פעולות או פתח מסמך אחר. דגל היה מחזיר במצב
 * כזה „עיצוב קודם” שאינו קודם לשום דבר — כלומר מחיל על הטקסט שלפניו עיצוב
 * שלא היה לו מעולם, וזה נראה למשתמש כמו שהקיצור הרס לו את המסמך.
 *
 * לכן המתג נשען על **מה שהמנוע מדווח עכשיו**: הקיצור נחשב „דלוק” רק כשמה
 * שהסמן עומד בו כבר נראה כמו הערכה (`readingMatchesPreset`). אם לא — הלחיצה
 * מחילה מחדש ולוכדת „קודם” חדש. אין מצב שבו החזרה מחילה עיצוב שלא נמדד.
 *
 * ## מה שאי אפשר לשחזר
 *
 * המנוע מקפל „בחירה מעורבת” ו„עוד לא נפתר” לאותו `undefined`, ולכן על בחירה
 * מעורבת אין „עיצוב קודם” אחד לשחזר. שדה כזה אינו נשמר ואינו משוחזר — והמשתמש
 * מקבל על כך הודעה בלחיצה עצמה, ולא גילוי מאוחר שהקיצור „שכח”.
 */
import {
  isEmptyPreset,
  presetSummary,
  readingMatchesPreset,
  reverseOf,
  unreadableFields,
  FIELD_TITLES,
  type FormatPreset,
  type FormatReading,
} from './format-preset';

/** מה שהמעטפת צריכה לעשות בעקבות הלחיצה. */
export interface ToggleDecision {
  /** הערכה להחלה. ריקה = אין מה להריץ. */
  apply: FormatPreset;
  /** מה לזכור כ„קודם”. `null` = לשכוח את מה שנזכר. */
  remember: FormatPreset | null;
  /** ההודעה לשורת המצב. תמיד יש אחת — לחיצה שלא אמרה דבר נראית כמו כלום. */
  status: string;
  /** האם הכיוון הוא „חזרה”. משמש את הבדיקות ואת נוסח ההודעה. */
  restoring: boolean;
}

export interface ToggleInput {
  /** השם שהמשתמש נתן לקיצור. */
  name: string;
  /** תווית הצירוף, לשם ההודעה „לחצו שוב כדי לחזור”. */
  label: string;
  preset: FormatPreset;
  /** מה שהמנוע מדווח על הבחירה **לפני** ההחלה. */
  reading: FormatReading;
  /** מה שנזכר בלחיצה קודמת, או `null` כשאין. */
  previous: FormatPreset | null;
}

/** „לא ניתן לשחזר את X ו-Y” — הרשימה בעברית. */
function unreadableNote(fields: readonly string[]): string {
  if (fields.length === 0) return '';
  const list =
    fields.length === 1
      ? fields[0]!
      : `${fields.slice(0, -1).join(', ')} ו${fields[fields.length - 1]}`;
  return ` (${list} לא נקראו מהבחירה ולא יוחזרו)`;
}

export function decideToggle(input: ToggleInput): ToggleDecision {
  const { name, label, preset, reading, previous } = input;

  // „דלוק” = יש מה להחזיר **וגם** מה שמצויר עכשיו הוא הערכה. ראו את ראש
  // הקובץ: זה מה שמפריד בין החזרה לבין החלה מחדש.
  const isOn = previous !== null && readingMatchesPreset(reading, preset);

  if (isOn) {
    if (isEmptyPreset(previous)) {
      // הלחיצה הראשונה נעשתה על בחירה שלא נקרא ממנה דבר. אין מה להחזיר,
      // והזיכרון נמחק כדי שהלחיצה הבאה תחיל מחדש במקום להיתקע כאן.
      return {
        apply: {},
        remember: null,
        status: `„${name}” — לא נשמר עיצוב קודם שאפשר להחזיר`,
        restoring: true,
      };
    }
    return {
      apply: previous,
      remember: null,
      status: `„${name}” — העיצוב הקודם הוחזר: ${presetSummary(previous)}`,
      restoring: true,
    };
  }

  const missing = unreadableFields(reading, preset).map((field) => FIELD_TITLES[field]);
  return {
    apply: preset,
    remember: reverseOf(reading, preset),
    status: `„${name}” הוחל — ${label} להחזרת העיצוב הקודם${unreadableNote(missing)}`,
    restoring: false,
  };
}

/**
 * הזיכרון של „מה היה לפני”, לפי מזהה קיצור.
 *
 * בזיכרון ולא באחסון, בכוונה: „העיצוב שהיה לפני הלחיצה” תקף לרצף עבודה אחד.
 * שמירתו בין הפעלות הייתה מבטיחה החזרה למצב של מסמך שנסגר לפני שבוע — ובדיוק
 * מהטעם הזה `forgetAll` נקרא בכל החלפת מסמך.
 */
export interface PresetToggles {
  /**
   * מכריע את הלחיצה **ומעדכן** את הזיכרון בהתאם.
   *
   * `previous` אינו נמסר — הוא **בדיוק** מה שהזיכרון הזה מחזיק, ומסירתו
   * מבחוץ הייתה מאפשרת לקורא לסתור אותו.
   */
  decide(input: Omit<ToggleInput, 'previous'> & { id: string }): ToggleDecision;
  /** שוכח קיצור אחד — למשל אחרי שנערך או נמחק. */
  forget(id: string): void;
  /** שוכח הכול. נקרא בהחלפת מסמך. */
  forgetAll(): void;
  /** האם קיצור מסוים מוחזק כדלוק. לבדיקות ולתצוגה. */
  isOn(id: string): boolean;
}

export function createPresetToggles(): PresetToggles {
  const memory = new Map<string, FormatPreset>();

  return {
    decide({ id, ...input }) {
      const decision = decideToggle({ ...input, previous: memory.get(id) ?? null });
      if (decision.remember === null) memory.delete(id);
      else memory.set(id, decision.remember);
      return decision;
    },
    forget(id) {
      memory.delete(id);
    },
    forgetAll() {
      memory.clear();
    },
    isOn(id) {
      return memory.has(id);
    },
  };
}
