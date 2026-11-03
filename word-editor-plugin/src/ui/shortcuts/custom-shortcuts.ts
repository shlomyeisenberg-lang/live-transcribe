/**
 * הקיצורים שהמשתמש הגדיר בעצמו — האוסף, האימות והנורמליזציה.
 *
 * ## מה זה כן, ומה זה לא
 *
 * זו מערכת **נוספת**, ולא שכבה מעל הרג'יסטרי. אף קיצור קיים אינו נגרע, אינו
 * מוסתר ואינו ניתן לדריסה: הרג'יסטרי הוא רשימה שמורה, וצירוף שיש לו רשומה
 * שם פשוט אינו זמין להצמדה. הסדר הזה גם עומד בזמן: קיצור אישי שנשמר על צירוף
 * פנוי, ובגרסה עתידית הצירוף ההוא נכנס לרג'יסטרי, **נושר בקריאה** — מה
 * שמובנה בעורך זוכה, והמשתמש מקבל הודעה ומצמיד צירוף אחר. בלי הכלל הזה
 * שדרוג היה משתיק בשקט פקודה מובנית.
 *
 * ## למה עיצוב, ולמה מתג
 *
 * הפעולה היחידה שקיצור אישי יכול להריץ כרגע היא „החל ערכת עיצוב”, והיא
 * **דו-מצבית**: לחיצה מחילה, לחיצה נוספת מחזירה את מה שהיה. זה מה שהמערכת
 * נועדה לו — לעבור לעיצוב עבודה ולחזור ממנו בלי לעבור בשלושה בוררים
 * ברצועה. ההכרעה מתי לחיצה שנייה „מחזירה” ומתי היא „מחילה מחדש” אינה כאן
 * אלא ב-preset-toggle.ts, כי היא מצב ריצה ולא נתון שמור.
 *
 * `kind` נשמר על כל רשומה בדיוק בשביל היום שבו תתווסף פעולה שנייה, והיא
 * נוספה: `apply-style` — „החל סגנון”, כמו קיצור המקשים שב-Word מוגדר בדיאלוג
 * „שנה סגנון”. הוא **אינו** מתג: החלת סגנון על פסקה שכבר בו אינה משנה דבר,
 * ו„חזרה למה שהיה” של סגנון היא החלת הסגנון הקודם — קיצור אחר.
 *
 * הסגנון נשמר לפי `styleId` ולא לפי שם: הקיצורים אינם של מסמך אחד, ו-`Heading1`
 * הוא אותו מזהה בכל מסמך Word. סגנון מותאם שאינו במסמך הפתוח נכשל בהודעה.
 */
import { comboAsShortcut, comboSignature, isBindableCombo, registryOwners, type KeyCombo } from './combo';
import { isEmptyPreset, normalizePreset, presetSummary, type FormatPreset } from './format-preset';
import type { Shortcut } from './registry';

/** הסגנון שקיצור מחיל: המזהה להחלה, והשם לתצוגה. */
export interface StyleRef {
  id: string;
  name: string;
}

interface ShortcutBase {
  /** מזהה יציב, נוצר בהוספה ואינו מוצג. */
  id: string;
  /** השם שהמשתמש נתן — זה מה שמופיע ברשימת הקיצורים ובשורת המצב. */
  name: string;
  combo: KeyCombo;
}

/** קיצור אישי אחד. */
export type CustomShortcut =
  | (ShortcutBase & { kind: 'format-preset'; preset: FormatPreset })
  | (ShortcutBase & { kind: 'apply-style'; style: StyleRef });

export type PresetShortcut = Extract<CustomShortcut, { kind: 'format-preset' }>;

/** מה שהקיצור עושה, במילים — לרשימות ולשורת המצב. */
export function shortcutSummary(entry: CustomShortcut): string {
  return entry.kind === 'apply-style' ? `החלת הסגנון „${entry.style.name}”` : presetSummary(entry.preset);
}

/** הקיצור שמחיל את הסגנון הזה, אם יש. */
export function styleShortcutFor(list: readonly CustomShortcut[], styleId: string): CustomShortcut | undefined {
  return list.find((entry) => entry.kind === 'apply-style' && entry.style.id === styleId);
}

function normalizeStyleRef(raw: unknown): StyleRef | null {
  if (!raw || typeof raw !== 'object') return null;
  const source = raw as Record<string, unknown>;
  const id = typeof source.id === 'string' ? source.id.trim() : '';
  if (id === '') return null;
  const name = typeof source.name === 'string' && source.name.trim() !== '' ? source.name.trim() : id;
  return { id, name: name.slice(0, MAX_NAME_LENGTH) };
}

/**
 * התקרה. אין בה קסם — היא קיימת מפני שהאחסון של אוצריא הוא KV של JSON, וגם
 * מפני שרשימה בת מאות רשומות אינה רשימה שאפשר לנהל בדיאלוג. מי שמגיע אליה
 * מקבל הודעה, ולא כתיבה שנכשלת בשקט.
 */
export const MAX_CUSTOM_SHORTCUTS = 40;

/** תקרת אורך לשם. שם ארוך יותר נחתך בקריאה ולא נדחה — הוא עדיין שם. */
export const MAX_NAME_LENGTH = 60;

/* ------------------------------------------------------------------ */
/* נורמליזציה                                                          */
/* ------------------------------------------------------------------ */

function normalizeCombo(raw: unknown): KeyCombo | null {
  if (!raw || typeof raw !== 'object') return null;
  const source = raw as Record<string, unknown>;
  if (typeof source.code !== 'string' || source.code === '') return null;
  const combo: KeyCombo = {
    code: source.code,
    ctrl: source.ctrl === true,
    shift: source.shift === true,
    alt: source.alt === true,
  };
  return isBindableCombo(combo) ? combo : null;
}

function normalizeName(raw: unknown): string {
  if (typeof raw !== 'string') return '';
  return raw.trim().slice(0, MAX_NAME_LENGTH);
}

/**
 * רשומה בודדת. `null` = לא ניתנת לשחזור, ואז היא נושרת: קיצור פגום שנשאר
 * ברשימה הוא שורה בדיאלוג שאי אפשר לתקן ואי אפשר להבין.
 */
function normalizeEntry(raw: unknown): CustomShortcut | null {
  if (!raw || typeof raw !== 'object') return null;
  const source = raw as Record<string, unknown>;

  if (source.kind !== 'format-preset' && source.kind !== 'apply-style') return null;

  const combo = normalizeCombo(source.combo);
  if (!combo) return null;

  if (source.kind === 'apply-style') {
    const style = normalizeStyleRef(source.style);
    const styleEntryId = typeof source.id === 'string' && source.id !== '' ? source.id : null;
    if (!style || styleEntryId === null) return null;
    return { id: styleEntryId, name: normalizeName(source.name) || style.name, kind: 'apply-style', combo, style };
  }

  const preset = normalizePreset(source.preset);
  // ערכה שאין בה שדה אחד תקין אינה מחילה כלום, וקיצור כזה נראה שבור. עדיף
  // שייעלם ויוגדר מחדש מאשר שיישאר ולא יעשה דבר.
  if (isEmptyPreset(preset)) return null;

  const id = typeof source.id === 'string' && source.id !== '' ? source.id : null;
  if (id === null) return null;

  // שם ריק אינו מפיל את הרשומה: הצירוף והערכה — מה שהקיצור **עושה** — שלמים,
  // וסיכום הערכה הוא שם אמיתי ולא placeholder. רשומה שנושרת בגלל שם היא
  // עיצוב שאבד בלי סיבה.
  const name = normalizeName(source.name) || presetSummary(preset);
  return { id, name, kind: 'format-preset', combo, preset };
}

/**
 * מה שנשר בקריאה, **לפי סיבה**.
 *
 * מספר אחד לא הספיק, וזה נמדד: המעטפת דיווחה כל נשירה כהתנגשות עם פעולה
 * מובנית, ומשתמש שרשומה אחת שלו הייתה פגומה נשלח לחפש התנגשות שאינה קיימת.
 * הסיבה היא מה שקובע מה אומרים לו — `reserved` הוא „הצירוף נלקח, בחרו אחר”,
 * וכל השאר הוא „הרשומה לא נקראה”.
 */
export interface DropCounts {
  /** סך כל מה שלא נטען. `0` = הכול נקרא. */
  total: number;
  /** הצירוף תפוס בידי פעולה מובנית — המסלול של שדרוג. ראו את ראש הקובץ. */
  reserved: number;
  /** רשומה שאינה ניתנת לשחזור: צורה פגומה, צירוף שאינו ניתן לקשירה, ערכה ריקה. */
  invalid: number;
  /** צירוף שרשומה קודמת באותה רשימה כבר לקחה. */
  duplicate: number;
  /** חריגה מהתקרה. */
  overflow: number;
}

/**
 * הרשימה מהערך הגולמי שנקרא מהאחסון.
 *
 * ארבעה סינונים: חריגה מהתקרה, רשומה פגומה, צירוף שתפוס ברג'יסטרי (ראו את
 * ראש הקובץ), וצירוף שרשומה קודמת כבר לקחה — הראשונה זוכה. הספירה מוחזרת
 * לפי סיבה כדי שהמעטפת תוכל לומר למשתמש את הדבר הנכון, במקום שהוא יגלה
 * בלחיצה שקיצור נעלם.
 */
export function normalizeCustomShortcuts(raw: unknown): {
  list: CustomShortcut[];
  dropped: DropCounts;
} {
  const dropped: DropCounts = { total: 0, reserved: 0, invalid: 0, duplicate: 0, overflow: 0 };
  if (!Array.isArray(raw)) return { list: [], dropped };

  const reserved = registryOwners();
  const taken = new Set<string>();
  const list: CustomShortcut[] = [];

  const drop = (reason: Exclude<keyof DropCounts, 'total'>): void => {
    dropped[reason] += 1;
    dropped.total += 1;
  };

  for (const item of raw) {
    if (list.length >= MAX_CUSTOM_SHORTCUTS) {
      drop('overflow');
      continue;
    }
    const entry = normalizeEntry(item);
    if (!entry) {
      drop('invalid');
      continue;
    }
    const signature = comboSignature(entry.combo);
    if (reserved.has(signature)) {
      drop('reserved');
      continue;
    }
    if (taken.has(signature)) {
      drop('duplicate');
      continue;
    }
    taken.add(signature);
    list.push(entry);
  }

  return { list, dropped };
}

/**
 * מה שנאמר למשתמש על מה שלא נטען, או `''` כשאין מה לומר.
 *
 * הנוסח כאן ולא במעטפת מפני שהוא **נגזר מהספירה**, וזה בדיוק מה שהיה שגוי:
 * הודעה שנוקבת בהתנגשות על רשומה פגומה שולחת אותו לחפש משהו שאינו קיים.
 * שני מסלולים, ולא אחד — ורק ה-`reserved` הוא זה שיש לו מה לעשות בתגובה.
 */
export function dropMessage(dropped: DropCounts): string {
  if (dropped.total === 0) return '';

  const other = dropped.total - dropped.reserved;

  if (dropped.reserved > 0 && other === 0) {
    return dropped.reserved === 1
      ? 'קיצור אישי אחד לא נטען — הצירוף שלו משמש כבר פעולה מובנית'
      : `${dropped.reserved} קיצורים אישיים לא נטענו — הצירופים שלהם משמשים כבר פעולות מובנות`;
  }

  if (dropped.reserved === 0) {
    return other === 1
      ? 'רשומת קיצור אחת באחסון לא נקראה'
      : `${other} רשומות קיצור באחסון לא נקראו`;
  }

  // שתי הסיבות יחד: כל אחת נאמרת בשמה, כי התגובה לכל אחת שונה.
  return `${dropped.reserved} קיצורים אישיים לא נטענו — הצירופים שלהם משמשים כבר פעולות מובנות; ועוד ${other} רשומות באחסון לא נקראו`;
}

/* ------------------------------------------------------------------ */
/* התנגשויות ואימות                                                    */
/* ------------------------------------------------------------------ */

/** מי מחזיק את הצירוף. `null` = פנוי. */
export type ComboConflict =
  | { kind: 'builtin'; description: string }
  | { kind: 'custom'; id: string; name: string }
  | { kind: 'macro'; name: string };

/**
 * חתימה → שם הפריט, לצירופים שאינם מהרג'יסטרי ואינם מהרשימה האישית.
 *
 * בפועל אלה קיצורי מערכת המאקרו (`shortcutTextOwners`). הם נמסרים מבחוץ ולא
 * נקראים כאן, מפני שהם תלויים ב**מסמך הפתוח**: `MacroKit` הוא של ה-session,
 * ומודול טהור אינו יכול לשאול אותו.
 */
export type TakenCombos = ReadonlyMap<string, string>;

/**
 * האם הצירוף תפוס, ובידי מי.
 *
 * שלושה מחזיקים אפשריים, ולא אחד: הרג'יסטרי, הרשימה האישית, ומערכת המאקרו.
 * השלישי הוא זה שקל לשכוח — ההנמקה המלאה ב-`signaturesOfShortcutText`.
 *
 * `exceptId` מחריג רשומה שנערכת כרגע — אחרת כל שמירה חוזרת של אותה רשומה
 * הייתה מתנגשת בעצמה.
 */
export function comboConflict(
  combo: KeyCombo,
  list: readonly CustomShortcut[],
  exceptId?: string,
  taken?: TakenCombos,
): ComboConflict | null {
  const signature = comboSignature(combo);

  const builtin = registryOwners().get(signature);
  if (builtin) return { kind: 'builtin', description: builtin.description };

  for (const entry of list) {
    if (entry.id === exceptId) continue;
    if (comboSignature(entry.combo) === signature) {
      return { kind: 'custom', id: entry.id, name: entry.name };
    }
  }

  const macro = taken?.get(signature);
  if (macro !== undefined) return { kind: 'macro', name: macro };

  return null;
}

/** ההודעה על התנגשות, בעברית. */
export function conflictMessage(conflict: ComboConflict): string {
  switch (conflict.kind) {
    case 'builtin':
      return `הצירוף הזה משמש כבר לפעולה מובנית: ${conflict.description}`;
    case 'macro':
      return `הצירוף הזה משמש כבר למאקרו „${conflict.name}”`;
    default:
      return `הצירוף הזה משמש כבר לקיצור „${conflict.name}”`;
  }
}

/** טיוטה מהדיאלוג — רשומה לפני שאושרה. */
export interface CustomShortcutDraft {
  /** קיים בעריכה, חסר בהוספה. */
  id?: string;
  name: string;
  combo: KeyCombo | null;
  preset: FormatPreset;
  /** קיצור „החל סגנון”. כשהוא קיים `preset` אינו נקרא. */
  style?: StyleRef;
}

/**
 * מה שמונע שמירה — ועם **סיווג**, ולא רק נוסח.
 *
 * ## למה הסיווג נדרש
 *
 * שני סוגי „מונע” שונים לגמרי נופלים כאן לאותה פונקציה: „עוד לא מילאת שם”
 * הוא **טופס שלא הושלם**, ו„הצירוף תפוס” הוא **תשובה שלילית** על משהו
 * שהמשתמש כן עשה. הצגתם באותו קול הפרה כלל מתועד של הבית: דיאלוג אינו נפתח
 * עם הודעת שגיאה אדומה לפני שאיש עשה כלום (ראו LinkDialog.vue, וכך גם
 * DocDefaultsDialog, FontAdvancedDialog, IndexDialog ו-MacrosDialog).
 *
 * ## למה ההתנגשות קודמת לחוסר
 *
 * מדוד: מי ש**לוכד צירוף לפני שהוא נותן שם** ראה „יש לתת שם לקיצור” בעוד
 * הלוכד מציג „Ctrl+S” — כלומר צירוף תפוס בלי שום סימן שהוא תפוס, והשם
 * מסתיר את התשובה החשובה. לכן ברגע שיש צירוף, ההתנגשות היא מה שנאמר.
 */
export type DraftIssue =
  /** הטופס טרם הושלם. הנחיה, לא שגיאה. */
  | { kind: 'incomplete'; message: string }
  /** בקשה שנדחתה — צירוף תפוס, או אין מקום. */
  | { kind: 'conflict'; message: string };

export function draftIssue(
  draft: CustomShortcutDraft,
  list: readonly CustomShortcut[],
  taken?: TakenCombos,
): DraftIssue | null {
  if (draft.combo && isBindableCombo(draft.combo)) {
    const conflict = comboConflict(draft.combo, list, draft.id, taken);
    if (conflict) return { kind: 'conflict', message: conflictMessage(conflict) };
  }

  if (draft.name.trim() === '') return { kind: 'incomplete', message: 'יש לתת שם לקיצור' };
  if (!draft.combo) return { kind: 'incomplete', message: 'יש ללחוץ על צירוף המקשים' };
  if (!isBindableCombo(draft.combo)) {
    return {
      kind: 'incomplete',
      message: 'צירוף חייב לכלול Ctrl או Alt, או להיות מקש פונקציה',
    };
  }
  if (!draft.style && isEmptyPreset(draft.preset)) {
    return { kind: 'incomplete', message: 'יש לבחור לפחות תכונת עיצוב אחת' };
  }

  if (draft.id === undefined && list.length >= MAX_CUSTOM_SHORTCUTS) {
    return {
      kind: 'conflict',
      message: `אין מקום לקיצור נוסף (${MAX_CUSTOM_SHORTCUTS} הוא המרב)`,
    };
  }

  return null;
}

/** הנוסח בלבד — שער השמירה, ומה שהבדיקות של האימות נשענות עליו. */
export function draftProblem(
  draft: CustomShortcutDraft,
  list: readonly CustomShortcut[],
  taken?: TakenCombos,
): string | null {
  return draftIssue(draft, list, taken)?.message ?? null;
}

/* ------------------------------------------------------------------ */
/* שינוי הרשימה                                                        */
/* ------------------------------------------------------------------ */

/**
 * מזהה חדש. `crypto.randomUUID` אינו בשימוש בכוונה: הוא אינו קיים בהקשר לא
 * מאובטח, ומזהה של שורה בדיאלוג אינו זקוק לאקראיות קריפטוגרפית — הוא זקוק
 * לייחודיות מול הרשימה, וזה מה שנבדק כאן.
 */
export function newShortcutId(list: readonly CustomShortcut[]): string {
  const used = new Set(list.map((entry) => entry.id));
  for (let index = 1; ; index += 1) {
    const id = `cs-${index}`;
    if (!used.has(id)) return id;
  }
}

/**
 * מוסיפה או מעדכנת. מחזירה רשימה חדשה — הקורא מחזיק אותה ב-ref, והחלפה
 * במקום הייתה מסתירה את השינוי מ-Vue.
 */
export function upsertShortcut(
  list: readonly CustomShortcut[],
  draft: CustomShortcutDraft,
  taken?: TakenCombos,
): { ok: true; list: CustomShortcut[]; id: string } | { ok: false; message: string } {
  const problem = draftProblem(draft, list, taken);
  if (problem !== null) return { ok: false, message: problem };

  const id = draft.id ?? newShortcutId(list);
  const base = { id, name: draft.name.trim().slice(0, MAX_NAME_LENGTH), combo: draft.combo! };
  const entry: CustomShortcut = draft.style
    ? { ...base, kind: 'apply-style', style: draft.style }
    : { ...base, kind: 'format-preset', preset: draft.preset };

  const index = list.findIndex((item) => item.id === id);
  const next = [...list];
  if (index === -1) next.push(entry);
  else next[index] = entry;

  return { ok: true, list: next, id };
}

/** מסירה רשומה. רשימה בלי הרשומה, גם כשלא הייתה שם. */
export function removeShortcut(
  list: readonly CustomShortcut[],
  id: string,
): CustomShortcut[] {
  return list.filter((entry) => entry.id !== id);
}

/* ------------------------------------------------------------------ */
/* הגשר להתאמה ולרשימת העזרה                                           */
/* ------------------------------------------------------------------ */

/**
 * הרשומות בצורת `Shortcut`, כדי שההתאמה תעבור דרך `matchShortcut` הקיים.
 *
 * זה העיקר בהחלטה הזאת: אין כאן מסלול התאמה שני. כל מה שנבדק על הקיצורים
 * המובנים — פריסת מקלדת, IME, מקש מוחזק, Meta של macOS — תופס גם כאן, בלי
 * שורת קוד נוספת ובלי שורת בדיקה נוספת.
 *
 * התיאור הוא השם שהמשתמש נתן ועוד סיכום הערכה, כדי שאותן רשומות יוכלו
 * להופיע בדיאלוג „קיצורי מקלדת” ולהיראות שם כמו כל השאר.
 */
export function customMatchers(list: readonly CustomShortcut[]): Shortcut[] {
  return list.map((entry) =>
    comboAsShortcut(entry.combo, { id: entry.id, description: entry.name }),
  );
}
