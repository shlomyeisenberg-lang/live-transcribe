/**
 * כיווניות פסקה: `Ctrl` + `Shift` הימני = מימין לשמאל, `Ctrl` + `Shift`
 * השמאלי = משמאל לימין. זה הצירוף של Word עצמו, והוא היחיד בתוסף שאינו
 * „מקש עם מודיפיירים” אלא **מודיפייר שמשוחרר**.
 *
 * ## למה מודול נפרד
 *
 * המנתב הרגיל עובד על `keydown` ועל מקש שיש לו `code` או `key`. כאן המקש הוא
 * ה-`Shift` עצמו, וההכרעה נופלת על `keyup`: כל עוד ה-`Shift` לחוץ אי אפשר
 * לדעת אם המשתמש מתכוון לכיווניות או שהוא באמצע `Ctrl+Shift+X`. לכן זו מכונת
 * מצבים קטנה ומפורשת, ולא עוד רשומה במנתב.
 *
 * ההבחנה בין ימין לשמאל היא על `event.code` (`ShiftRight` / `ShiftLeft`),
 * ובנפילה על `event.location` — שניהם בלתי תלויים בפריסת המקלדת, כמו כל שאר
 * הקיצורים כאן.
 */
import type { CommandId } from '../../engine/capabilities';

export type ParagraphDirection = 'rtl' | 'ltr';

/** המינימום שהזיהוי קורא. לא `KeyboardEvent` מלא — כדי שהבדיקות לא ידרשו DOM. */
export interface ModifierEventLike {
  code?: string;
  key?: string;
  location?: number;
  ctrlKey: boolean;
  metaKey: boolean;
  altKey?: boolean;
}

/**
 * `Ctrl` ו-`Meta` — המודיפיירים שאינם מבטלים חימוש.
 *
 * לחיצת `Ctrl` היא אירוע `keydown` ככל אחר, ולכן „כל מקש אחר מבטל” ניקה
 * בגללה את החימוש — ומי שלחץ `Shift` ואז `Ctrl` לא קיבל כלום. הסדר ההפוך
 * עבד, ולכן זה נראה תקין. `Alt` **אינו** ברשימה בכוונה: AltGr מדווח
 * `ctrl+alt`, והוא שכבת תווים ולא צירוף שלנו.
 */
function isNeutralModifier(event: ModifierEventLike): boolean {
  if (event.key === 'Control' || event.key === 'Meta') return true;
  const code = event.code;
  return (
    code === 'ControlLeft' ||
    code === 'ControlRight' ||
    code === 'MetaLeft' ||
    code === 'MetaRight'
  );
}

/** `DOM_KEY_LOCATION_LEFT` / `_RIGHT`, כשאין `code`. */
const LOCATION_LEFT = 1;
const LOCATION_RIGHT = 2;

/** איזה `Shift` זה — או `null` אם זה אינו `Shift` בכלל. */
function shiftSide(event: ModifierEventLike): 'ShiftLeft' | 'ShiftRight' | null {
  if (event.code === 'ShiftLeft' || event.code === 'ShiftRight') return event.code;

  // דפדפן ישן, או אירוע סינתטי בלי `code`.
  if (event.key !== 'Shift') return null;
  if (event.location === LOCATION_RIGHT) return 'ShiftRight';
  if (event.location === LOCATION_LEFT) return 'ShiftLeft';

  // `Shift` בלי צד ידוע אינו כיווניות: ניחוש כאן היה משנה כיוון פסקה בטעות.
  return null;
}

export interface DirectionDetector {
  keydown: (event: ModifierEventLike) => void;
  /** הכיוון שיש להחיל, או `null`. */
  keyup: (event: ModifierEventLike) => ParagraphDirection | null;
  /**
   * מבטל חימוש שנתקע. נדרש כשהחלון מאבד פוקוס: ה-`keyup` מגיע לחלון האחר,
   * ובלי האיפוס הכיוון היה מתהפך כשהמשתמש חוזר ומשחרר.
   */
  reset: () => void;
}

export function createDirectionDetector(): DirectionDetector {
  /** ה-`Shift` שנלחץ בזמן ש-`Ctrl` היה לחוץ, ועדיין לא „התלכלך”. */
  let armed: 'ShiftLeft' | 'ShiftRight' | null = null;

  return {
    keydown(event) {
      const side = shiftSide(event);
      if (side) {
        // החימוש אינו דורש ש-`Ctrl` כבר יהיה לחוץ: סדר הלחיצה אינו משהו
        // שמשתמש חושב עליו, ומי שלחץ Shift ואז Ctrl התכוון לאותו דבר בדיוק.
        // מה ש**כן** נבדק הוא ה-`keyup`, שם `Ctrl` חייב עדיין להיות לחוץ.
        //
        // `Alt` פוסל: AltGr מדווח `ctrl+alt`, והוא שכבת תווים ולא צירוף שלנו.
        armed = event.altKey === true ? null : side;
        return;
      }

      // `Ctrl` או `Meta` שנלחצים **אחרי** ה-Shift אינם „מקש אחר”: הם בדיוק
      // מה שהצירוף מחכה לו. בלי החריגה הזאת ההבטחה שבהערה למעלה — שסדר
      // הלחיצה אינו משנה — לא התקיימה.
      if (isNeutralModifier(event)) return;

      // כל מקש אחר מבטל: `Ctrl+Shift+X` הוא קו חוצה, ושחרור ה-Shift אחריו
      // אינו אמור להפוך גם את כיוון הפסקה.
      armed = null;
    },

    keyup(event) {
      const side = armed;
      armed = null;

      if (!side || shiftSide(event) !== side) return null;
      // `Ctrl` שוחרר לפני ה-`Shift` — המשתמש התחרט, או שזה היה צירוף אחר.
      if (!event.ctrlKey && !event.metaKey) return null;
      if (event.altKey === true) return null;

      return side === 'ShiftRight' ? 'rtl' : 'ltr';
    },

    reset() {
      armed = null;
    },
  };
}

export const DIRECTION_COMMANDS: Record<ParagraphDirection, CommandId> = {
  rtl: 'direction-rtl',
  ltr: 'direction-ltr',
};

export interface DirectionShortcutDeps {
  runCommand: (id: CommandId) => void;
  /** האם להתעלם — פוקוס בשדה טקסט של הממשק, או דיאלוג מודאלי פתוח. */
  isBlocked?: (target: EventTarget | null) => boolean;
  target?: Pick<EventTarget, 'addEventListener' | 'removeEventListener'>;
}

/** מחבר את הזיהוי לאירועים האמיתיים. מחזיר ביטול. */
export function createDirectionShortcut(deps: DirectionShortcutDeps): { dispose: () => void } {
  const detector = createDirectionDetector();
  const isBlocked = deps.isBlocked ?? (() => false);
  const target = deps.target ?? (typeof window === 'undefined' ? undefined : window);

  const onKeyDown = (event: Event) => {
    detector.keydown(event as KeyboardEvent);
  };

  const onKeyUp = (event: Event) => {
    const keyEvent = event as KeyboardEvent;
    const direction = detector.keyup(keyEvent);
    if (!direction) return;
    if (isBlocked(keyEvent.target)) return;
    deps.runCommand(DIRECTION_COMMANDS[direction]);
  };

  // אובדן פוקוס באמצע: ה-`keyup` יגיע לחלון האחר, והחימוש היה נשאר תלוי
  // באוויר עד שהמשתמש חוזר ומשחרר Shift — ואז הכיוון מתהפך בלי שביקש.
  const onBlur = () => detector.reset();

  // `keydown` ב-capture: רשומה שרצה לפני המנוע (`runsBeforeEngine`, וכל קיצור
  // אישי) בולעת את ההקשה ב-`stopPropagation` על `window` בשלב ה-capture, ומאזין
  // bubble כאן לא היה רואה את ה-E של `Ctrl+Shift+E` — החימוש נשאר, ושחרור
  // ה-Shift הפך את כיוון הפסקה. על אותו צומת `stopPropagation` אינו חוסם.
  target?.addEventListener('keydown', onKeyDown, true);
  target?.addEventListener('keyup', onKeyUp);
  target?.addEventListener('blur', onBlur);

  let disposed = false;
  return {
    dispose() {
      if (disposed) return;
      disposed = true;
      target?.removeEventListener('keydown', onKeyDown, true);
      target?.removeEventListener('keyup', onKeyUp);
      target?.removeEventListener('blur', onBlur);
    },
  };
}
