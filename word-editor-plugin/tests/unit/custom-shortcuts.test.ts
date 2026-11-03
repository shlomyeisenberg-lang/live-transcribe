/**
 * האוסף של הקיצורים האישיים: מה נקרא מהאחסון, מה נדחה, ומי מחזיק צירוף.
 *
 * הכלל שכל הקובץ הזה שומר עליו הוא היחיד שהמשתמש ביקש במפורש: **המערכת
 * נוספת ואינה משנה קיצור קיים.** הוא נאכף פעמיים — בשמירה (`draftProblem`)
 * ובקריאה (`normalizeCustomShortcuts`) — והכפילות מכוונת: הראשונה מונעת
 * התנגשות שנוצרת עכשיו, והשנייה התנגשות שנוצרה בשדרוג, אחרי שצירוף פנוי נכנס
 * לרג'יסטרי.
 */
import { describe, expect, it } from 'vitest';
import {
  MAX_CUSTOM_SHORTCUTS,
  MAX_NAME_LENGTH,
  draftIssue,
  dropMessage,
  comboConflict,
  conflictMessage,
  customMatchers,
  draftProblem,
  newShortcutId,
  normalizeCustomShortcuts,
  removeShortcut,
  shortcutSummary,
  styleShortcutFor,
  upsertShortcut,
  type CustomShortcut,
} from '../../src/ui/shortcuts/custom-shortcuts';
import { matchShortcut } from '../../src/ui/shortcuts/match';
import {
  comboLabel,
  comboSignature,
  registryOwners,
  reservationText,
  shortcutTextOwners,
  type KeyCombo,
} from '../../src/ui/shortcuts/combo';

const COMBO = { code: 'KeyK', ctrl: true, shift: false, alt: true };

/**
 * `count` צירופים **שפנויים באמת**.
 *
 * לא רשימה קשיחה: `Ctrl+Alt+R` ו-`Ctrl+Alt+P` תפוסים בידי המאקרו, ורשימה
 * שנכתבה ביד הייתה נשברת בכל קיצור מובנה שיתווסף — כלומר בדיוק במקום שבו
 * הבדיקה אמורה להמשיך לעבוד. המקור הוא `registryOwners`, אותו מקור שהקוד
 * הנבדק משתמש בו.
 */
function freeCombos(count: number): KeyCombo[] {
  const reserved = registryOwners();
  const letters = [...'ABCDEFGHIJKLMNOPQRSTUVWXYZ'];
  const combos: KeyCombo[] = [];

  for (const shift of [false, true]) {
    for (const letter of letters) {
      if (combos.length === count) return combos;
      const combo: KeyCombo = { code: `Key${letter}`, ctrl: true, shift, alt: true };
      if (!reserved.has(comboSignature(combo))) combos.push(combo);
    }
  }

  throw new Error(`אין ${count} צירופים פנויים`);
}

function record(over: Record<string, unknown> = {}): Record<string, unknown> {
  return {
    id: 'cs-1',
    name: 'כותרת קטע',
    kind: 'format-preset',
    combo: COMBO,
    preset: { fontFamily: 'David', fontSizePt: 14 },
    ...over,
  };
}

const ENTRY: CustomShortcut = normalizeCustomShortcuts([record()]).list[0]!;

describe('קריאה מהאחסון', () => {
  it('רשומה שלמה נקראת', () => {
    const { list, dropped } = normalizeCustomShortcuts([record()]);
    expect(dropped.total).toBe(0);
    expect(list).toEqual([
      {
        id: 'cs-1',
        name: 'כותרת קטע',
        kind: 'format-preset',
        combo: COMBO,
        preset: { fontFamily: 'David', fontSizePt: 14 },
      },
    ]);
  });

  it('ערך שאינו מערך הוא רשימה ריקה', () => {
    for (const raw of [null, undefined, {}, 'x', 7]) {
      expect(normalizeCustomShortcuts(raw)).toEqual({
        list: [],
        dropped: { total: 0, reserved: 0, invalid: 0, duplicate: 0, overflow: 0 },
      });
    }
  });

  it('רשומה בלי צירוף שאפשר לקשור נושרת', () => {
    // אות בלי מודיפייר הייתה נורית על הקלדה רגילה.
    const { list, dropped } = normalizeCustomShortcuts([
      record({ combo: { code: 'KeyK', ctrl: false, shift: false, alt: false } }),
    ]);
    expect(list).toEqual([]);
    expect(dropped).toMatchObject({ total: 1, invalid: 1 });
  });

  it('רשומה שערכתה ריקה נושרת — קיצור שאינו מחיל דבר נראה שבור', () => {
    expect(
      normalizeCustomShortcuts([record({ preset: { fontSizePt: -3 } })]).dropped,
    ).toMatchObject({ total: 1, invalid: 1 });
  });

  it('סוג שאינו מוכר נושר', () => {
    // אחסון בלי הבחנת סוג היה מחייב מיגרציה ביום שתתווסף פעולה שנייה.
    expect(normalizeCustomShortcuts([record({ kind: 'run-script' })]).dropped).toMatchObject({
      total: 1,
      invalid: 1,
    });
  });

  it('שם ריק אינו מפיל את הרשומה — סיכום הערכה הוא שם אמיתי', () => {
    const { list } = normalizeCustomShortcuts([record({ name: '   ' })]);
    expect(list[0]?.name).toBe("David, 14 נק'");
  });

  it('שם ארוך נחתך ולא נדחה', () => {
    const { list } = normalizeCustomShortcuts([record({ name: 'א'.repeat(200) })]);
    expect(list[0]?.name).toHaveLength(MAX_NAME_LENGTH);
  });

  it('**צירוף שתפוס ברג׳יסטרי נושר** — הפעולה המובנית זוכה', () => {
    // זה המסלול של שדרוג: הקיצור נשמר כשהצירוף היה פנוי, ובגרסה הזאת הוא
    // כבר של העורך. בלי הסינון הזה השדרוג היה משתיק בשקט פקודה מובנית.
    const { list, dropped } = normalizeCustomShortcuts([
      record({ combo: { code: 'KeyS', ctrl: true, shift: false, alt: false } }),
    ]);
    expect(list).toEqual([]);
    // **הסיבה** נשמרת, ולא רק המספר: היא מה שקובע מה נאמר למשתמש.
    expect(dropped).toMatchObject({ total: 1, reserved: 1, invalid: 0 });
  });

  it('שני קיצורים על אותו צירוף — הראשון זוכה', () => {
    const { list, dropped } = normalizeCustomShortcuts([record(), record({ id: 'cs-2' })]);
    expect(list.map((entry) => entry.id)).toEqual(['cs-1']);
    expect(dropped).toMatchObject({ total: 1, duplicate: 1 });
  });

  it('מעל התקרה נושר, ואינו מפיל את מה שנקרא', () => {
    const many = freeCombos(MAX_CUSTOM_SHORTCUTS + 3).map((combo, index) =>
      record({ id: `cs-${index}`, combo }),
    );
    const { list, dropped } = normalizeCustomShortcuts(many);
    expect(list).toHaveLength(MAX_CUSTOM_SHORTCUTS);
    expect(dropped).toMatchObject({ total: 3, overflow: 3 });
  });
});

describe('התנגשות', () => {
  it('צירוף פנוי אינו מתנגש', () => {
    expect(comboConflict({ code: 'KeyK', ctrl: true, shift: false, alt: true }, [])).toBeNull();
  });

  it('צירוף מובנה מדווח מי מחזיק אותו, בעברית', () => {
    const conflict = comboConflict({ code: 'KeyS', ctrl: true, shift: false, alt: false }, []);
    expect(conflict?.kind).toBe('builtin');
    expect(conflictMessage(conflict!)).toContain('פעולה מובנית');
  });

  it('צירוף של קיצור אישי אחר מדווח בשמו', () => {
    const conflict = comboConflict(COMBO, [ENTRY]);
    expect(conflict).toEqual({ kind: 'custom', id: 'cs-1', name: 'כותרת קטע' });
    expect(conflictMessage(conflict!)).toContain('כותרת קטע');
  });

  it('רשומה שנערכת אינה מתנגשת בעצמה', () => {
    expect(comboConflict(COMBO, [ENTRY], 'cs-1')).toBeNull();
  });
});

describe('אימות טיוטה', () => {
  const draft = { name: 'כותרת', combo: COMBO, preset: { fontSizePt: 14 } };

  it('טיוטה שלמה עוברת', () => {
    expect(draftProblem(draft, [])).toBeNull();
  });

  it('כל חסר מדווח בנפרד, בסדר מילוי הטופס', () => {
    expect(draftProblem({ ...draft, name: '  ' }, [])).toBe('יש לתת שם לקיצור');
    expect(draftProblem({ ...draft, combo: null }, [])).toBe('יש ללחוץ על צירוף המקשים');
    expect(draftProblem({ ...draft, preset: {} }, [])).toBe('יש לבחור לפחות תכונת עיצוב אחת');
  });

  it('צירוף בלי מודיפייר נדחה עם הסבר', () => {
    const problem = draftProblem(
      { ...draft, combo: { code: 'KeyK', ctrl: false, shift: true, alt: false } },
      [],
    );
    expect(problem).toContain('Ctrl או Alt');
  });

  it('התקרה חוסמת הוספה אך לא עריכה', () => {
    const full = normalizeCustomShortcuts(
      freeCombos(MAX_CUSTOM_SHORTCUTS).map((combo, index) => record({ id: `cs-${index}`, combo })),
    ).list;
    expect(full).toHaveLength(MAX_CUSTOM_SHORTCUTS);

    expect(draftProblem({ ...draft, combo: { code: 'F9', ctrl: false, shift: false, alt: false } }, full)).toContain(
      'אין מקום',
    );
    // עריכה של רשומה קיימת אינה מוסיפה שורה, ולכן אינה נחסמת.
    expect(
      draftProblem({ ...draft, id: full[0]!.id, combo: full[0]!.combo }, full),
    ).toBeNull();
  });
});

describe('שינוי הרשימה', () => {
  it('הוספה מקבלת מזהה חדש ואינה נוגעת ברשימה שנמסרה', () => {
    const before: CustomShortcut[] = [ENTRY];
    const result = upsertShortcut(before, {
      name: 'גוף',
      combo: { code: 'KeyG', ctrl: true, shift: false, alt: true },
      preset: { fontSizePt: 11 },
    });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.list).toHaveLength(2);
    expect(result.id).not.toBe('cs-1');
    expect(before).toHaveLength(1);
  });

  it('עריכה מחליפה במקום, ואינה מוסיפה שורה', () => {
    const result = upsertShortcut([ENTRY], {
      id: 'cs-1',
      name: 'כותרת גדולה',
      combo: COMBO,
      preset: { fontSizePt: 20 },
    });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.list).toHaveLength(1);
    expect(result.list[0]).toMatchObject({ name: 'כותרת גדולה', preset: { fontSizePt: 20 } });
  });

  it('טיוטה פסולה מסורבת עם אותה הודעה של האימות', () => {
    const result = upsertShortcut([], { name: '', combo: COMBO, preset: { fontSizePt: 14 } });
    expect(result).toEqual({ ok: false, message: 'יש לתת שם לקיצור' });
  });

  it('מזהה חדש אינו מתנגש בקיימים, גם אחרי מחיקות', () => {
    const list = [ENTRY, { ...ENTRY, id: 'cs-2' }, { ...ENTRY, id: 'cs-3' }];
    expect(newShortcutId(list)).toBe('cs-4');
    expect(newShortcutId(removeShortcut(list, 'cs-2'))).toBe('cs-2');
  });

  it('מחיקה של מזהה שאינו קיים אינה משנה דבר', () => {
    expect(removeShortcut([ENTRY], 'cs-9')).toEqual([ENTRY]);
  });
});

describe('הגשר להתאמה', () => {
  it('הרשומה מתאימה להקשה, גם בפריסה עברית', () => {
    const [matcher] = customMatchers([ENTRY]);
    expect(matcher?.id).toBe('cs-1');
    expect(matcher?.label).toBe('Ctrl+Alt+K');
    expect(
      matchShortcut(
        { key: 'ל', code: 'KeyK', ctrlKey: true, metaKey: false, shiftKey: false, altKey: true },
        matcher!,
      ),
    ).toBe(true);
  });
});

describe('התנגשות עם קיצורי המאקרו', () => {
  /**
   * המקור השני לקיצורים מבוססי-נתונים בעורך. הוא זה שקל לשכוח, וההשלכה שלו
   * אינה שגיאה אלא **שתיקה**: קשירת המאקרו יושבת על שלב הלכידה של מכל
   * המסמך, כלומר היא מקדימה את המנתב שלנו.
   */
  it('מחרוזת קיצור של מאקרו הופכת לחתימות שאפשר להשוות', () => {
    const owners = shortcutTextOwners([{ name: 'בסד', shortcut: 'Ctrl+Alt+1' }]);
    expect(owners.get(comboSignature({ code: 'Digit1', ctrl: true, shift: false, alt: true }))).toBe(
      'בסד',
    );
  });

  it('`Mod` של החבילה מכוסה — הוא Ctrl וגם Meta', () => {
    const owners = shortcutTextOwners([{ name: 'קטע', shortcut: 'Mod+Alt+J' }]);
    expect(owners.has(comboSignature({ code: 'KeyJ', ctrl: true, shift: false, alt: true }))).toBe(
      true,
    );
  });

  it('פריט בלי קיצור, ומחרוזת שאינה ניתנת לפירוק, אינם תופסים דבר', () => {
    // מחרוזת שהפרסר שלהם דוחה אינה יכולה לירות אצלם, ולכן אין מה להתנגש —
    // וחסימה עליה הייתה גוזלת מהמשתמש צירוף בלי סיבה.
    const owners = shortcutTextOwners([
      { name: 'בלי' },
      { name: 'זבל', shortcut: '???' },
    ]);
    expect(owners.size).toBe(0);
  });

  it('צירוף שמאקרו מחזיק נדחה, עם שמו', () => {
    const taken = shortcutTextOwners([{ name: 'בסד', shortcut: 'Ctrl+Alt+K' }]);

    const conflict = comboConflict(COMBO, [], undefined, taken);
    expect(conflict).toEqual({ kind: 'macro', name: 'בסד' });
    expect(conflictMessage(conflict!)).toContain('מאקרו „בסד”');

    expect(draftProblem({ name: 'כותרת', combo: COMBO, preset: { fontSizePt: 14 } }, [], taken)).toContain(
      'מאקרו',
    );
  });

  it('השמירה אוכפת את אותה בדיקה — ולא רק הטופס', () => {
    const taken = shortcutTextOwners([{ name: 'בסד', shortcut: 'Ctrl+Alt+K' }]);
    const result = upsertShortcut([], { name: 'כותרת', combo: COMBO, preset: { fontSizePt: 14 } }, taken);
    expect(result.ok).toBe(false);
  });

  it('בלי רשימת מאקרו אין חסימה — מסמך סגור אינו מחזיק צירופים', () => {
    expect(draftProblem({ name: 'כותרת', combo: COMBO, preset: { fontSizePt: 14 } }, [])).toBeNull();
  });

  it('ההצהרה ההפוכה נוקבת בשם שהפרסר שלהם מבין', () => {
    // `comboLabel` מציג „↑” מפני שזה מה שהמשתמש רואה; ההצהרה כלפי המאקרו
    // חייבת להיות „ArrowUp”, שהוא ה-`key` שהפרסר שלהם מתאים לפיו. שימוש
    // בתווית כאן היה משאיר את מקשי החצים בלי הגנה, בשקט.
    expect(comboLabel({ code: 'ArrowUp', ctrl: true, shift: false, alt: true })).toBe('Ctrl+Alt+↑');
    expect(reservationText({ code: 'ArrowUp', ctrl: true, shift: false, alt: true })).toBe(
      'Ctrl+Alt+ArrowUp',
    );

    // וההצהרה הזאת אכן חוזרת לאותה חתימה דרך הפרסר שלהם — כלומר היא מגנה
    // בפועל, ולא רק נראית נכונה.
    const combo = { code: 'ArrowUp', ctrl: true, shift: false, alt: true };
    const owners = shortcutTextOwners([{ name: 'קיצור אישי', shortcut: reservationText(combo) }]);
    expect(owners.has(comboSignature(combo))).toBe(true);
  });

  it('אות וספרה עוברות את המסלול ההפוך בלי אובדן', () => {
    for (const code of ['KeyK', 'Digit3']) {
      const combo = { code, ctrl: true, shift: true, alt: false };
      const owners = shortcutTextOwners([{ name: 'x', shortcut: reservationText(combo) }]);
      expect(owners.has(comboSignature(combo)), code).toBe(true);
    }
  });
});

describe('ההודעה על מה שלא נטען', () => {
  /**
   * זו הייתה תקלה אמיתית ש-QA מדד: כל נשירה דווחה כהתנגשות עם פעולה מובנית,
   * ומשתמש שרשומה אחת שלו הייתה פגומה נשלח לחפש התנגשות שאינה קיימת.
   */
  it('אין נשירה — אין הודעה', () => {
    expect(dropMessage({ total: 0, reserved: 0, invalid: 0, duplicate: 0, overflow: 0 })).toBe('');
  });

  it('התנגשות בלבד — ההודעה נוקבת בפעולה המובנית', () => {
    const message = dropMessage({ total: 1, reserved: 1, invalid: 0, duplicate: 0, overflow: 0 });
    expect(message).toContain('פעולה מובנית');
  });

  it('**רשומה פגומה בלבד — ההודעה אינה מזכירה התנגשות**', () => {
    const message = dropMessage({ total: 1, reserved: 0, invalid: 1, duplicate: 0, overflow: 0 });
    expect(message).not.toContain('מובנית');
    expect(message).toContain('לא נקראה');
  });

  it('שתי הסיבות יחד — כל אחת נאמרת בשמה', () => {
    const message = dropMessage({ total: 2, reserved: 1, invalid: 1, duplicate: 0, overflow: 0 });
    expect(message).toContain('מובנות');
    expect(message).toContain('לא נקראו');
  });
});

describe('סיווג מה שמונע שמירה', () => {
  /**
   * ההפרדה הזאת היא מה שמונע דיאלוג שנפתח עם הודעת שגיאה אדומה לפני שאיש
   * עשה כלום — כלל מתועד של הבית (LinkDialog.vue). „חסר” הוא הנחיה,
   * „תפוס” הוא תשובה שלילית.
   */
  it('טופס ריק הוא „חסר”, ולא שגיאה', () => {
    const issue = draftIssue({ name: '', combo: null, preset: {} }, []);
    expect(issue?.kind).toBe('incomplete');
  });

  it('צירוף תפוס הוא „התנגשות” — גם כשהשם עוד ריק', () => {
    // מדוד: מי שלכד צירוף לפני שנתן שם ראה „יש לתת שם” בעוד הלוכד מציג
    // „Ctrl+S” — כלומר צירוף תפוס בלי שום סימן שהוא תפוס.
    const issue = draftIssue(
      { name: '', combo: { code: 'KeyS', ctrl: true, shift: false, alt: false }, preset: {} },
      [],
    );
    expect(issue?.kind).toBe('conflict');
    expect(issue?.message).toContain('פעולה מובנית');
  });

  it('התקרה היא „התנגשות” ולא „חסר” — היא תשובה לבקשה', () => {
    const full = normalizeCustomShortcuts(
      freeCombos(MAX_CUSTOM_SHORTCUTS).map((combo, index) => record({ id: `cs-${index}`, combo })),
    ).list;
    const issue = draftIssue(
      { name: 'עוד', combo: { code: 'F9', ctrl: false, shift: false, alt: false }, preset: { fontSizePt: 12 } },
      full,
    );
    expect(issue?.kind).toBe('conflict');
  });

  it('טופס שלם אינו מחזיר דבר', () => {
    expect(draftIssue({ name: 'כותרת', combo: COMBO, preset: { fontSizePt: 14 } }, [])).toBeNull();
  });

  it('`draftProblem` נשאר הנוסח של אותו סיווג — שער אחד לשמירה', () => {
    const draft = { name: '', combo: null, preset: {} };
    expect(draftProblem(draft, [])).toBe(draftIssue(draft, [])?.message);
  });
});

describe('קיצור „החל סגנון”', () => {
  const styleRecord = (over: Record<string, unknown> = {}) => ({
    id: 'cs-2',
    name: '',
    kind: 'apply-style',
    combo: freeCombos(1)[0],
    style: { id: 'Heading1', name: 'כותרת 1' },
    ...over,
  });

  it('נקרא מהאחסון, ושם ריק נופל לשם הסגנון', () => {
    const { list, dropped } = normalizeCustomShortcuts([styleRecord()]);
    expect(dropped.total).toBe(0);
    expect(list[0]).toMatchObject({ kind: 'apply-style', name: 'כותרת 1', style: { id: 'Heading1', name: 'כותרת 1' } });
    expect(shortcutSummary(list[0]!)).toBe('החלת הסגנון „כותרת 1”');
  });

  it('רשומה בלי מזהה סגנון נושרת', () => {
    expect(normalizeCustomShortcuts([styleRecord({ style: { id: '  ' } })]).list).toEqual([]);
  });

  it('נשמר בלי ערכת עיצוב, ונמצא לפי הסגנון', () => {
    const combo = freeCombos(2)[1]!;
    const result = upsertShortcut([ENTRY], {
      name: 'כותרת 1',
      combo,
      preset: {},
      style: { id: 'Heading1', name: 'כותרת 1' },
    });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(styleShortcutFor(result.list, 'Heading1')).toMatchObject({ kind: 'apply-style', combo });
    expect(styleShortcutFor(result.list, 'Heading2')).toBeUndefined();
  });

  it('צירוף תפוס בידי קיצור אחר נדחה', () => {
    const result = upsertShortcut([ENTRY], {
      name: 'כותרת 1',
      combo: COMBO,
      preset: {},
      style: { id: 'Heading1', name: 'כותרת 1' },
    });
    expect(result).toMatchObject({ ok: false });
  });
});
