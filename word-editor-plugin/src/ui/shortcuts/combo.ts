/**
 * צירוף מקשים כ**נתון** — מה שהמשתמש הצמיד בעצמו, בשונה מרשומה ברג'יסטרי.
 *
 * ## למה מודול נפרד, ולמה `code` ולא מחרוזת
 *
 * `superdoc-macros` כבר מקבלת קיצור אישי כמחרוזת („Ctrl+Alt+M”) ומפרקת אותה
 * ב-`parseShortcut`. הדרך הזאת לא נבחרה כאן, ולא מטעמי טעם: המחרוזת נוקבת
 * ב**תו**, והכלל של המאגר הזה — שנקבע אחרי באג שהרג את כל הקיצורים בפריסה
 * עברית — הוא ש**המקש הפיזי** (`event.code`) הוא מקור האמת. כדי לגשר בין
 * השניים צריך מפה הופכית מתו למקש, ובפריסה עברית היא אינה חד-חד-ערכית:
 * המקש `BracketLeft` מפיק „]”, כלומר בדיוק ההפך מהשם שלו.
 *
 * לכן הצירוף כאן **נלכד מהקשה אמיתית** ונשמר כמקש פיזי. אין מה לפרסר, אין
 * מה לשקף, והמשתמש רואה את מה שלחץ. זה גם מה שמאפשר להשתמש ב-`matchShortcut`
 * הקיים — הצירוף מומר לרשומה בצורת `Shortcut`, ועובר דרך אותו קוד התאמה
 * שנבדק כבר מול IME, מול מקש מוחזק ומול Meta של macOS.
 *
 * ## מה התווית מבטיחה, ומה לא
 *
 * `comboLabel` מציגה את **שם המקש הפיזי** בפריסה האמריקאית — זה השם שכתוב על
 * המקש ברוב המקלדות, וזה מה ש-Word מציג. באות ובספרה זה גם מה שהמשתמש רואה
 * על המקש בכל פריסה; בסימן פיסוק זה עשוי להיבדל מהתו שהמקש מפיק בעברית, וזה
 * מחיר מדוד: התווית מתארת את המקש שנלחץ, ולא את התו שיוצא ממנו.
 */
import { codesForKey, parseShortcut } from 'superdoc-macros';
import type { Shortcut } from './registry';
import { SHORTCUTS } from './registry';

/**
 * צירוף מקשים אחד.
 *
 * `ctrl` פירושו „Ctrl או Meta”, בדיוק כמו בשדה בעל אותו שם ברג'יסטרי: מקלדת
 * Mac שולחת Cmd, ואותו צירוף צריך לתפוס את שניהם.
 */
export interface KeyCombo {
  /** המקש הפיזי, בערכי `KeyboardEvent.code`. */
  code: string;
  ctrl: boolean;
  shift: boolean;
  alt: boolean;
}

/**
 * המקשים שאינם יכולים להיות המקש של הצירוף — הם המודיפייר עצמו. הקשה על
 * Ctrl לבדו אינה צירוף אלא חצי צירוף, ובלי הסינון הזה הלוכד היה „נסגר” על
 * המודיפייר ברגע שהמשתמש מתחיל להקיש.
 */
const MODIFIER_CODES = new Set([
  'ControlLeft',
  'ControlRight',
  'ShiftLeft',
  'ShiftRight',
  'AltLeft',
  'AltRight',
  'MetaLeft',
  'MetaRight',
  'CapsLock',
]);

/**
 * מקשי הפונקציה — היחידים שמותר לקשור **בלי** מודיפייר.
 *
 * למה הכלל הזה בכלל: קיצור אישי בלי Ctrl/Alt היה נורה על הקלדה רגילה. „ק”
 * הוא אות במסמך לפני שהוא קיצור, וקישור כזה היה הופך את העורך לבלתי שמיש
 * בלי שהמשתמש יבין למה. `superdoc-macros` אוכפת בדיוק את אותו כלל
 * (`hasBindingModifier`), ומאותו טעם. מקש פונקציה אינו מפיק תו, ולכן הוא
 * היוצא מן הכלל — וכך גם F3/F11/F12 ברג'יסטרי עצמו.
 */
function isFunctionKey(code: string): boolean {
  return /^F([1-9]|1[0-2])$/.test(code);
}

/** האם הצירוף ניתן לקשירה: מודיפייר אמיתי, או מקש פונקציה. */
export function isBindableCombo(combo: KeyCombo): boolean {
  if (combo.code === '' || MODIFIER_CODES.has(combo.code)) return false;
  if (combo.ctrl || combo.alt) return true;
  // Shift לבדו אינו מודיפייר לצורך קשירה: „Shift+K” הוא „K” גדולה.
  return isFunctionKey(combo.code);
}

/** המינימום שהלכידה קוראת מהאירוע. מאפשר בדיקות בלי DOM. */
export interface ComboEventLike {
  code: string;
  ctrlKey: boolean;
  metaKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
}

/**
 * הצירוף שהאירוע מתאר, או `null` כשאין בו צירוף שאפשר לקשור — הקשה על
 * מודיפייר לבדו, או אות בלי מודיפייר.
 */
export function comboFromEvent(event: ComboEventLike): KeyCombo | null {
  const combo: KeyCombo = {
    code: event.code,
    ctrl: event.ctrlKey || event.metaKey,
    shift: event.shiftKey,
    alt: event.altKey,
  };
  return isBindableCombo(combo) ? combo : null;
}

/**
 * שם המקש כפי שהוא מוצג. מכסה את מה שמקלדת רגילה מייצרת; קוד שאינו במפה
 * מוצג כמו שהוא, וזה עדיף על „מקש לא ידוע” — הוא לפחות מזהה את המקש.
 */
const KEY_NAMES: Record<string, string> = {
  Space: 'Space',
  Enter: 'Enter',
  NumpadEnter: 'Enter',
  Tab: 'Tab',
  Backspace: 'Backspace',
  Delete: 'Delete',
  Insert: 'Insert',
  Home: 'Home',
  End: 'End',
  PageUp: 'PageUp',
  PageDown: 'PageDown',
  ArrowUp: '↑',
  ArrowDown: '↓',
  ArrowLeft: '←',
  ArrowRight: '→',
  Minus: '-',
  Equal: '=',
  BracketLeft: '[',
  BracketRight: ']',
  Backslash: '\\',
  Semicolon: ';',
  Quote: "'",
  Backquote: '`',
  Comma: ',',
  Period: '.',
  Slash: '/',
  NumpadAdd: 'Numpad +',
  NumpadSubtract: 'Numpad -',
  NumpadMultiply: 'Numpad *',
  NumpadDivide: 'Numpad /',
  NumpadDecimal: 'Numpad .',
};

/** מקשי החצים — שם ה-code שלהם הוא גם ה-`key` שלהם. ראו `reservationText`. */
const ARROW_KEYS = new Set(['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight']);

/** שם המקש הבודד, בלי המודיפיירים. */
export function keyName(code: string): string {
  const letter = /^Key([A-Z])$/.exec(code);
  if (letter) return letter[1]!;
  const digit = /^Digit([0-9])$/.exec(code);
  if (digit) return digit[1]!;
  const numpad = /^Numpad([0-9])$/.exec(code);
  if (numpad) return `Numpad ${numpad[1]}`;
  if (isFunctionKey(code)) return code;
  return KEY_NAMES[code] ?? code;
}

/**
 * התווית שהמשתמש רואה — „Ctrl+Shift+K”.
 *
 * הסדר Ctrl→Shift→Alt הוא הסדר שכל תוויות הרג'יסטרי כתובות בו, ושתי רשימות
 * שמסדרות אחרת נראות כמו שני קיצורים שונים.
 */
export function comboLabel(combo: KeyCombo): string {
  const parts: string[] = [];
  if (combo.ctrl) parts.push('Ctrl');
  if (combo.shift) parts.push('Shift');
  if (combo.alt) parts.push('Alt');
  parts.push(keyName(combo.code));
  return parts.join('+');
}

/**
 * הצירוף כמחרוזת ש**הפרסר של מערכת המאקרו** מבין — לצורך הצהרה הפוכה:
 * הקיצורים האישיים שלנו נמסרים לה כ-`reservedShortcuts`, כדי שהיא לא תאפשר
 * להצמיד מאקרו לצירוף שכבר תפוס.
 *
 * נבדלת מ-`comboLabel` במקשים שאין להם תו: התווית מציגה „↑” מפני שזה מה
 * שהמשתמש רואה, והפרסר שלהם מתאים לפי `event.key` — ששם הוא „ArrowUp”.
 * שימוש בתווית לצורך ההצהרה היה משאיר את מקשי החצים בלי הגנה, בשקט.
 */
export function reservationText(combo: KeyCombo): string {
  const parts: string[] = [];
  if (combo.ctrl) parts.push('Ctrl');
  if (combo.shift) parts.push('Shift');
  if (combo.alt) parts.push('Alt');
  parts.push(ARROW_KEYS.has(combo.code) ? combo.code : keyName(combo.code));
  return parts.join('+');
}

/**
 * מפתח השוואה. שני צירופים מתנגשים אם ורק אם החתימה שלהם זהה — וזה גם מה
 * שהופך „Ctrl” ל„Ctrl או Meta” בהשוואה, בדיוק כמו בהתאמה בזמן ריצה.
 */
export function comboSignature(combo: KeyCombo): string {
  return [
    combo.ctrl ? 'Ctrl' : '',
    combo.shift ? 'Shift' : '',
    combo.alt ? 'Alt' : '',
    combo.code,
  ].join('|');
}

/**
 * הצירוף כרשומת `Shortcut`, כדי ש-`matchShortcut` הקיים יתאים אותו.
 *
 * `label`/`description`/`group` נדרשים בטיפוס ואינם משמשים בהתאמה עצמה;
 * הם ממולאים באמת ולא ב-placeholder, כדי שרשומה כזאת תוכל להיכנס גם לרשימת
 * העזרה בלי שיהיה בה שקר.
 */
export function comboAsShortcut(
  combo: KeyCombo,
  meta: { id: string; description: string },
): Shortcut {
  return {
    id: meta.id,
    label: comboLabel(combo),
    description: meta.description,
    group: 'custom',
    code: combo.code,
    ...(combo.ctrl ? { ctrl: true } : {}),
    ...(combo.shift ? { shift: true } : {}),
    ...(combo.alt ? { alt: true } : {}),
  };
}

/**
 * החתימות שהרג'יסטרי מחזיק, מהשדות המובנים שלו — ולא מהתווית.
 *
 * מהשדות ולא מהתווית בכוונה: `superdoc-macros` מקבלת את התוויות כמחרוזות
 * (`reservedShortcuts`) ומתעלמת ממה שאינה מצליחה לפרסר, ולכן תווית כמו
 * „Ctrl + Shift ימני” פשוט אינה מוגנת שם. כאן המקור הוא `code`/`ctrl`/
 * `shift`/`alt` עצמם, כלומר בדיוק מה שההתאמה בזמן ריצה משתמשת בו — ואין פער
 * בין מה שנחסם לבין מה שיירה.
 *
 * רשומה בלי `code` (כיווניות הפסקה, שמזוהה בשחרור מודיפייר) אינה נכנסת: אין
 * לה מקש, ולכן שום צירוף אישי אינו יכול להיות שווה לה.
 */
export function registryOwners(): Map<string, Shortcut> {
  const owners = new Map<string, Shortcut>();
  for (const shortcut of SHORTCUTS as readonly Shortcut[]) {
    if (shortcut.code === undefined) continue;
    const codes = typeof shortcut.code === 'string' ? [shortcut.code] : shortcut.code;
    for (const code of codes) {
      const signature = comboSignature({
        code,
        ctrl: shortcut.ctrl === true,
        shift: shortcut.shift === true,
        alt: shortcut.alt === true,
      });
      // הראשון זוכה: הרג'יסטרי אינו מכיל צירוף כפול (בדיקת החוזה אוכפת), ולכן
      // זה אינו מקרה שקורה — והכתיבה המפורשת חוסכת שאלה בקריאה.
      if (!owners.has(signature)) owners.set(signature, shortcut);
    }
  }
  return owners;
}

/**
 * החתימות של קיצור שנכתב כמחרוזת — הצורה שמערכת המאקרו שומרת בה קיצורים.
 *
 * ## למה זה נדרש בכלל
 *
 * הרג'יסטרי אינו המקור היחיד לקיצורים בעורך: מערכת המאקרו מצמידה קיצורים
 * **לנתונים** — הקלטות, קטעי טקסט, סקריפטים וכלים — ואלה נקשרים על מכל
 * המסמך (`kit.attachShortcuts`). כלומר שני מקורות של קיצורים אישיים, ומי
 * שאינו מכיר את השני מאפשר להצמיד צירוף שכבר תפוס. מה שהמשתמש רואה אז אינו
 * שגיאה אלא **אחד מהשניים ששותק** — הקשירה של המאקרו יושבת על שלב הלכידה של
 * המכל, ולכן היא זו שיורה, וההצמדה החדשה פשוט אינה עושה דבר.
 *
 * ## למה הפרסר שלהם, ולא שלנו
 *
 * `parseShortcut` ו-`codesForKey` הם **בדיוק** מה שהקשירה שלהם משתמשת בו
 * בזמן ריצה. פרסר משלנו על אותן מחרוזות היה מסכים איתם ברוב המקרים ונבדל
 * בקצוות — וכל פער כזה הוא צירוף שנראה פנוי ואינו פנוי. מחרוזת שהם אינם
 * מצליחים לפרסר אינה נכנסת: היא אינה יכולה לירות אצלם, ולכן אין מה להתנגש.
 */
export function signaturesOfShortcutText(text: string): string[] {
  const parsed = parseShortcut(text);
  if (!parsed) return [];

  // `codesForKey` ממפה אותיות וספרות בלבד — בכל השאר ההתאמה שלהם נופלת
  // ל-`event.key`, ולנו נדרש מקש פיזי כדי לבנות חתימה. `codesFromKeyName`
  // הוא ההיפוך של `keyName`, כלומר בדיוק אוצר המילים שהצירופים שלנו יכולים
  // לייצר; מחרוזת שאינה בו נשארת בלי חתימה, ואז הצירוף ההוא אינו מוגן מצידנו
  // (הצד שלהם עדיין שומר עליו — ראו `reservationText`).
  const codes = codesForKey(parsed.key);
  const resolved = codes.length > 0 ? codes : codesFromKeyName(parsed.key);

  // `Mod` שלהם הוא „Ctrl או Meta”, וזה בדיוק מה ש-`ctrl` שלנו מציין.
  const ctrl = parsed.ctrl || parsed.meta || parsed.mod;
  return resolved.map((code) =>
    comboSignature({ code, ctrl, shift: parsed.shift, alt: parsed.alt }),
  );
}

/**
 * ההיפוך של `keyName` — מהשם שמוצג אל המקש הפיזי.
 *
 * ההשוואה חסינה לרישיות מפני ש-`parseShortcut` מחזיר את המקש באותיות קטנות.
 * בסימני פיסוק ההיפוך נכון לפריסה האמריקאית בלבד, בדיוק כמו `keyName` עצמה —
 * וזה מדוד: השם שכתוב על המקש הוא מה שמשתמש מקליד בשדה של דיאלוג המאקרו.
 */
function codesFromKeyName(name: string): string[] {
  const wanted = name.trim().toLowerCase();
  if (wanted === '') return [];

  for (const code of ARROW_KEYS) {
    if (code.toLowerCase() === wanted) return [code];
  }
  if (/^f([1-9]|1[0-2])$/.test(wanted)) return [wanted.toUpperCase()];

  const matches = Object.entries(KEY_NAMES)
    .filter(([, label]) => label.toLowerCase() === wanted)
    .map(([code]) => code);
  return matches;
}

/**
 * חתימה → שם הפריט שמחזיק אותה, מרשימת פריטי מאקרו.
 *
 * הראשון זוכה, כמו בכל שאר מקומות ההתנגשות כאן: מערכת המאקרו אוכפת ייחודיות
 * בתוך עצמה, ולכן זה אינו מקרה שקורה — והכתיבה המפורשת חוסכת שאלה בקריאה.
 */
export function shortcutTextOwners(
  items: readonly { name: string; shortcut?: string }[],
): Map<string, string> {
  const owners = new Map<string, string>();
  for (const item of items) {
    if (!item.shortcut) continue;
    for (const signature of signaturesOfShortcutText(item.shortcut)) {
      if (!owners.has(signature)) owners.set(signature, item.name);
    }
  }
  return owners;
}
