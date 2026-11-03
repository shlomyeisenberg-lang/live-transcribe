/**
 * שדות: מספר עמוד, מספר העמודים, תאריך ועדכון.
 *
 * שלוש טענות שהמודול עומד או נופל עליהן, ולכן הן עיקר מה שנמדד כאן:
 *
 * 1. **ה-`instruction` שנשלח הוא קוד השדה של Word, מילה במילה.** מחרוזת
 *    שגויה אינה שגיאה במנוע — היא שדה שגוי במסמך של המשתמש, ולכן היא נבדקת
 *    כאן על התו.
 * 2. **השדה נכנס בסמן, ומיד מחושב.** `fields.insert` מכניס תוצאה ריקה
 *    ו-`fields.rebuild` הוא שמחשב אותה; בלי הקריאה השנייה המשתמש רואה מקום
 *    ריק.
 * 3. **כשל של ה-rebuild אינו כשל של ההכנסה.** השדה כבר במסמך, והודעת כשל על
 *    פעולה שהצליחה היא רעש.
 */
import { describe, expect, it } from 'vitest';
import {
  emptyFieldsState,
  insertDate,
  insertDateOnly,
  insertTimeOnly,
  insertPageCount,
  insertPageNumber,
  readFieldsState,
  rebuildAllFields,
  FIELD_INSTRUCTIONS,
  type FieldsHost,
} from '../../src/engine/fields';

interface Call {
  op: string;
  input?: unknown;
}

interface RawField {
  address?: unknown;
  instruction?: string;
  resolvedText?: string;
}

interface FakeOptions {
  /** מה `fields.list` מחזיר כ-`items` — עמוד, לא בהכרח כל המסמך. */
  fields?: readonly RawField[];
  /**
   * כמו `fields`, אבל **מתחלף** בין קריאה מלאה אחת (עמוד/עמודים, `offset:0`
   * עד סיום) לשנייה — כדי לבדוק ש`rebuildAllFields` שואב רשימה טרייה אחרי
   * כשל, ולא מסתמך על התצלום הראשון. הקריאה המלאה ה-N-ית (`offset:0`)
   * מקדמת לאיבר ה-N ברשימה; אחרי הסוף היא נשארת על האחרון. גובר על `fields`.
   */
  fieldsSequence?: readonly (readonly RawField[])[];
  /** `total` של `DiscoveryOutput`. `undefined` = גרסה שאינה חושפת אותו. */
  total?: number;
  /** קבלה שנכשלת, לפי שם הפעולה — לכל קריאה לאותה פעולה. */
  failures?: Record<string, { code: string; message?: string }>;
  /**
   * `fields.rebuild` נכשל (`TARGET_NOT_FOUND`) כשה-`fieldId` של ה-`target`
   * ברשימה הזאת — לבדיקת „ממשיך לשאר ומדווח נכון” בלי להפיל את כל הפעולות.
   */
  rebuildFailFieldIds?: readonly string[];
  /** פעולות שזורקות. */
  throws?: readonly string[];
  /** מסלולים שאינם קיימים בפאסדה — גרסת מנוע שאין לה את היכולת. */
  missing?: readonly string[];
  /** מה `selection.current` מדווח. `null` = אין קטע שאפשר לפעול עליו. */
  blockId?: string | null;
  /** הקבלה של `fields.insert` תחזיר את הכתובת הזאת. `null` = בלי כתובת כלל. */
  insertedField?: unknown;
}

function fakeEngine(options: FakeOptions = {}) {
  const calls: Call[] = [];
  const missing = new Set(options.missing ?? []);
  const throwing = new Set(options.throws ?? []);
  const failures = options.failures ?? {};
  const rebuildFailFieldIds = new Set(options.rebuildFailFieldIds ?? []);

  function route<T>(op: string, impl: (input: unknown) => T): ((input: unknown) => T) | undefined {
    if (missing.has(op)) return undefined;
    return (input: unknown) => {
      calls.push({ op, input });
      if (throwing.has(op)) throw new Error('boom');
      return impl(input);
    };
  }

  const receipt = (op: string): { success: boolean; failure?: { code: string; message?: string } } =>
    failures[op] ? { success: false, failure: failures[op] } : { success: true };

  const blockId = options.blockId === undefined ? 'block-1' : options.blockId;
  const selectionTarget = {
    kind: 'text',
    segments: blockId ? [{ blockId, range: { start: 3, end: 3 } }] : [],
  };

  // אינדקס ה„דור” הנוכחי של `fieldsSequence`: מתקדם בכל קריאה עם `offset:0`,
  // ונשאר על האחרון אחרי הסוף. `fields` הרגיל הוא דור יחיד קבוע.
  let generation = -1;

  const doc = {
    selection: {
      current: route('selection.current', () => ({ empty: true, target: selectionTarget })),
    },
    fields: {
      // הכפיל מכבד `limit`/`offset` ולא מחזיר את אותו עמוד לנצח: `list` הוא
      // `DiscoveryOutput`, ו„עדכן שדות” שואב ממנו עמוד אחר עמוד. כפיל שמתעלם
      // מהעמוד היה מאשר בירוק גם מימוש שרץ על העמוד הראשון בלבד.
      list: route('fields.list', (input) => {
        const query = (input ?? {}) as { limit?: number; offset?: number };
        const offset = query.offset ?? 0;

        let all: readonly RawField[];
        if (options.fieldsSequence) {
          if (offset === 0) generation = Math.min(generation + 1, options.fieldsSequence.length - 1);
          all = options.fieldsSequence[Math.max(generation, 0)] ?? [];
        } else {
          all = options.fields ?? [];
        }

        const end = query.limit === undefined ? undefined : offset + query.limit;
        return {
          items: all.slice(offset, end),
          ...(options.total === undefined ? {} : { total: options.total }),
        };
      }),
      insert: route('fields.insert', () => {
        const result = receipt('fields.insert');
        if (!result.success) return result;
        const field =
          options.insertedField === undefined
            ? { kind: 'field', blockId: 'block-1', occurrenceIndex: 0, nestingDepth: 0 }
            : options.insertedField;
        return field === null ? result : { ...result, field };
      }),
      rebuild: route('fields.rebuild', (input) => {
        const target = (input as { target?: { fieldId?: string } } | undefined)?.target;
        if (target && typeof target.fieldId === 'string' && rebuildFailFieldIds.has(target.fieldId)) {
          return { success: false, failure: { code: 'TARGET_NOT_FOUND' } };
        }
        return receipt('fields.rebuild');
      }),
      remove: route('fields.remove', () => receipt('fields.remove')),
    },
  };

  const host = { activeEditor: { doc } } as unknown as FieldsHost;
  const ops = (): string[] => calls.map((call) => call.op);
  const inputs = (op: string): unknown[] =>
    calls.filter((call) => call.op === op).map((call) => call.input);

  return { host, calls, ops, inputs };
}

describe('הכנסת שדה', () => {
  it('„מספר עמוד” שולח `PAGE` גולמי בסמן, ומחשב את התוצאה מיד', async () => {
    const engine = fakeEngine();

    expect(await insertPageNumber(engine.host)).toEqual({ ok: true });

    expect(engine.inputs('fields.insert')).toEqual([
      {
        at: { kind: 'text', segments: [{ blockId: 'block-1', range: { start: 3, end: 3 } }] },
        instruction: 'PAGE',
        mode: 'raw',
      },
    ]);
    // ה-rebuild מקבל את הכתובת שהקבלה החזירה, ולא את `at` — שתי כתובות בשני
    // מודלים שונים, והחלפה ביניהן הייתה חריגת `INVALID_INPUT`.
    expect(engine.inputs('fields.rebuild')).toEqual([
      { target: { kind: 'field', blockId: 'block-1', occurrenceIndex: 0, nestingDepth: 0 } },
    ]);
  });

  it('„מספר העמודים” שולח `NUMPAGES`, ו„תאריך” שולח `DATE` עם מתג הפורמט', async () => {
    // המתג הוא העיקר כאן. `DATE` עירום נמדד במנוע האמיתי ומחזיר `2026-08-24` —
    // ISO לועזי, ואפילו יום אחורה (UTC). `DATE \\@ "dd/MM/yyyy HH:mm"` נמדד באותה
    // הרצה ומחזיר `25/08/2026`. הפירוט המלא בהערת הפתיחה של engine/fields.ts.
    const count = fakeEngine();
    expect(await insertPageCount(count.host)).toEqual({ ok: true });
    expect((count.inputs('fields.insert')[0] as { instruction: string }).instruction).toBe(
      'NUMPAGES',
    );

    const date = fakeEngine();
    expect(await insertDate(date.host)).toEqual({ ok: true });
    expect((date.inputs('fields.insert')[0] as { instruction: string }).instruction).toBe(
      'DATE \\@ "dd/MM/yyyy HH:mm"',
    );
  });

  it('שלושת הקודים הם בדיוק קודי השדות של Word', () => {
    // הקבוע הזה הוא מה שנכתב לתוך המסמך. שינוי בו הוא שינוי בקובץ שהמשתמש
    // ישמור, ולכן הוא ננעל כאן ולא רק בעקיפין דרך אתרי הקריאה.
    expect(FIELD_INSTRUCTIONS).toEqual({
      pageNumber: 'PAGE',
      pageCount: 'NUMPAGES',
      dateOnly: 'DATE \\@ "dd/MM/yyyy"',
      timeOnly: 'DATE \\@ "HH:mm"',
      date: 'DATE \\@ "dd/MM/yyyy HH:mm"',
    });
  });

  it('אין סמן במסמך: אין קריאה למנוע, ויש הודעה שאומרת מה לעשות', async () => {
    const engine = fakeEngine({ blockId: null });

    expect(await insertPageNumber(engine.host)).toEqual({
      ok: false,
      message:
        'הוספת מספר העמוד נכשלה: יש ללחוץ בגוף המסמך, על שורת טקסט שיש בה תו אחד לפחות, ואז להוסיף את השדה',
      reason: 'no-selection',
    });
    expect(engine.ops()).not.toContain('fields.insert');
  });

  it('קבלה שנכשלה מתורגמת לעברית, ואין rebuild על שדה שלא נכנס', async () => {
    const engine = fakeEngine({
      failures: { 'fields.insert': { code: 'PRECONDITION_FAILED' } },
    });

    expect(await insertDate(engine.host)).toEqual({
      ok: false,
      message: 'הוספת התאריך נכשלה: המסמך אינו במצב שמאפשר את הפעולה',
      reason: 'PRECONDITION_FAILED',
    });
    expect(engine.ops()).not.toContain('fields.rebuild');
  });

  it('`fields.insert` חסר בפאסדה: „אינו זמין בגרסה זו”', async () => {
    const engine = fakeEngine({ missing: ['fields.insert'] });

    expect(await insertPageCount(engine.host)).toEqual({
      ok: false,
      message: 'הוספת מספר העמודים נכשלה: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });

  it('חריגה מהמנוע הופכת להודעה ולא נזרקת החוצה', async () => {
    const engine = fakeEngine({ throws: ['fields.insert'] });

    expect(await insertPageNumber(engine.host)).toEqual({
      ok: false,
      message: 'הוספת מספר העמוד נכשלה: boom',
      reason: 'threw',
    });
  });

  it('rebuild שנכשל אחרי הכנסה מוצלחת אינו הופך אותה לכשל', async () => {
    // השדה **במסמך**. הוא ייראה בעדכון הבא, וכשל כאן היה מלמד את המשתמש
    // להתעלם מהודעות.
    const failing = fakeEngine({ failures: { 'fields.rebuild': { code: 'INTERNAL_ERROR' } } });
    expect(await insertPageNumber(failing.host)).toEqual({ ok: true });

    const throwing = fakeEngine({ throws: ['fields.rebuild'] });
    expect(await insertPageNumber(throwing.host)).toEqual({ ok: true });

    const missing = fakeEngine({ missing: ['fields.rebuild'] });
    expect(await insertPageNumber(missing.host)).toEqual({ ok: true });
  });

  it('קבלה בלי כתובת שדה: ההכנסה הצליחה, ואין rebuild על `undefined`', async () => {
    const engine = fakeEngine({ insertedField: null });

    expect(await insertPageNumber(engine.host)).toEqual({ ok: true });
    expect(engine.ops()).not.toContain('fields.rebuild');
  });

  it('אין Document API כלל: „המסמך עדיין נטען”', async () => {
    expect(await insertDate(null)).toEqual({
      ok: false,
      message: 'הוספת התאריך נכשלה: המסמך עדיין נטען',
      reason: 'document-api-unavailable',
    });
  });
});

describe('מניעת קינון — הכנסה על סמן שכבר בתוך שדה', () => {
  /**
   * זה בדיוק תרחיש הבאג שנמדד: `PAGE` שהיה אמור לפתור ל-`"1"` פתר
   * ל-`"28/08/202611"` אחרי שהוכנס עליו שדה נוסף באותה נקודה. הכפיל מדמה
   * זאת ב-`fieldsSequence`: התצלום „לפני” מראה שדה יחיד תקין, והתצלום
   * „אחרי” (אחרי ה-`insert`) מראה את אותו `fieldId` עם `resolvedText` שונה —
   * ההוכחה לקינון.
   */
  it('שדה קיים שה-`resolvedText` שלו השתנה בעקבות ההכנסה: ההכנסה מבוטלת ומדווחת כסירוב', async () => {
    const existing = { fieldId: 'f1', blockId: 'block-1', occurrenceIndex: 0, nestingDepth: 0 };
    const inserted = { fieldId: 'f2', blockId: 'block-1', occurrenceIndex: 1, nestingDepth: 0 };
    const engine = fakeEngine({
      fieldsSequence: [
        [{ address: existing, instruction: 'PAGE', resolvedText: '1' }],
        [
          // אחרי ההכנסה: השדה הקיים נשתל בתוכו והתוצאה שלו התלכלכה.
          { address: existing, instruction: 'PAGE', resolvedText: '11' },
          { address: inserted, instruction: 'NUMPAGES', resolvedText: '' },
        ],
      ],
      insertedField: inserted,
    });

    expect(await insertPageCount(engine.host)).toEqual({
      ok: false,
      message:
        'הוספת מספר העמודים נכשלה: הסמן היה בתוך שדה קיים, וההכנסה הייתה נבלעת בתוכו. ' +
        'ההכנסה בוטלה — יש להזיז את הסמן אל מחוץ לשדה הקיים ולנסות שוב',
      reason: 'field-in-field',
    });

    // השדה שקינן מוסר, וה-rebuild של „מספר העמודים” החדש לא נקרא — הוא בוטל.
    expect(engine.inputs('fields.remove')).toEqual([{ target: inserted, mode: 'raw' }]);
    expect(engine.ops()).not.toContain('fields.rebuild');
  });

  it('שדה קיים באותו בלוק שלא נגעו בו: לא מדווח קינון (ושינוי ב**בלוק אחר** מתעלם ממנו)', async () => {
    const sibling = { fieldId: 'f1', blockId: 'block-1', occurrenceIndex: 0, nestingDepth: 0 };
    const inserted = { fieldId: 'f2', blockId: 'block-1', occurrenceIndex: 1, nestingDepth: 0 };
    const elsewhere = { fieldId: 'fX', blockId: 'block-9', occurrenceIndex: 0, nestingDepth: 0 };
    const engine = fakeEngine({
      fieldsSequence: [
        [
          { address: sibling, instruction: 'DATE \\@ "dd/MM/yyyy HH:mm"', resolvedText: '28/08/2026 14:03' },
          { address: elsewhere, instruction: 'PAGE', resolvedText: '3' },
        ],
        [
          // ה-sibling נשאר בדיוק כפי שהיה — ההכנסה לא נגעה בו. `elsewhere`
          // דווקא כן השתנה, אבל הוא בבלוק אחר, ולכן אינו אמור להיספר.
          { address: sibling, instruction: 'DATE \\@ "dd/MM/yyyy HH:mm"', resolvedText: '28/08/2026 14:03' },
          { address: inserted, instruction: 'NUMPAGES', resolvedText: '' },
          { address: elsewhere, instruction: 'PAGE', resolvedText: '4' },
        ],
      ],
      insertedField: inserted,
    });

    expect(await insertPageCount(engine.host)).toEqual({ ok: true });
    expect(engine.ops()).not.toContain('fields.remove');
  });

  it('אין שדות קודמים בבלוק: אין קריאה שנייה ל-`fields.list` (מסלול מהיר)', async () => {
    const engine = fakeEngine({ fields: [] });

    expect(await insertDate(engine.host)).toEqual({ ok: true });
    // תצלום „לפני” יצא (רשימה ריקה), אבל בלי שדות קיימים אין מה להשוות אליו
    // אחרי ההכנסה — קריאה שנייה הייתה בזבוז על הפקד הכי נפוץ.
    expect(engine.inputs('fields.list')).toHaveLength(1);
  });

  it('שדה קיים בלי `fieldId` יציב: ההגנה מדלגת ואינה חוסמת הכנסה תקינה', async () => {
    // גרסת מנוע בלי handle יציב — אין דרך אמינה להשוות, ועדיף לוותר על ההגנה
    // מאשר לחסום הכנסה על בסיס ניחוש. ראו הערת הפתיחה של fields.ts.
    const engine = fakeEngine({
      fields: [{ address: { blockId: 'block-1', occurrenceIndex: 0 }, instruction: 'DATE', resolvedText: '28/08/2026' }],
    });

    expect(await insertPageNumber(engine.host)).toEqual({ ok: true });
    expect(engine.ops()).not.toContain('fields.remove');
  });
});

describe('rebuildAllFields', () => {
  it('מחשבת מחדש כל שדה שיש לו כתובת, ומדלגת על שדה בלעדיה', async () => {
    const engine = fakeEngine({
      fields: [
        { address: { kind: 'field', blockId: 'b1', occurrenceIndex: 0, nestingDepth: 0 } },
        { instruction: 'PAGE' },
        { address: { kind: 'field', blockId: 'b2', occurrenceIndex: 0, nestingDepth: 0 } },
      ],
    });

    expect(await rebuildAllFields(engine.host)).toEqual({ ok: true });
    expect(engine.inputs('fields.rebuild')).toEqual([
      { target: { kind: 'field', blockId: 'b1', occurrenceIndex: 0, nestingDepth: 0 } },
      { target: { kind: 'field', blockId: 'b2', occurrenceIndex: 0, nestingDepth: 0 } },
    ]);
  });

  it('מסמך בלי שדות הוא הצלחה שקטה', async () => {
    const engine = fakeEngine({ fields: [] });

    expect(await rebuildAllFields(engine.host)).toEqual({ ok: true });
    expect(engine.ops()).not.toContain('fields.rebuild');
  });

  it('`NO_OP` על שדה שכבר מעודכן אינו כשל', async () => {
    const engine = fakeEngine({
      fields: [{ address: { kind: 'field' } }],
      failures: { 'fields.rebuild': { code: 'NO_OP' } },
    });

    expect(await rebuildAllFields(engine.host)).toEqual({ ok: true });
  });

  it('כשל **אינו** עוצר את השאר: כל שדה ננסה, והתוצאה מדווחת כמה הצליחו וכמה נכשלו', async () => {
    // שני השדות נכשלים כאן (הכפיל מכשיל כל קריאה ל-`fields.rebuild`), אבל
    // שניהם **נוסו** — לא רק הראשון. זה ההפך המדויק מהבאג שנמדד: „עדכן שדות”
    // שעצר על הכשל הראשון והשאיר את כל מה שאחריו לא מעודכן, בלי שום סימן.
    const engine = fakeEngine({
      fields: [
        { address: { fieldId: 'a', blockId: 'b1', occurrenceIndex: 0, nestingDepth: 0 } },
        { address: { fieldId: 'b', blockId: 'b2', occurrenceIndex: 0, nestingDepth: 0 } },
      ],
      failures: { 'fields.rebuild': { code: 'PRECONDITION_FAILED' } },
    });

    expect(await rebuildAllFields(engine.host)).toEqual({
      ok: false,
      message: 'עדכון השדות לא הושלם: אף שדה לא עודכן, ו2 שדות נכשלו ולא עודכנו',
      reason: 'partial-rebuild',
    });
    expect(engine.inputs('fields.rebuild')).toHaveLength(2);
  });

  /**
   * זה בדיוק מה שנמדד בדפדפן: `rebuild` על שדה שקינן משקם/מסיר את המבנה
   * השגוי, וכתובת של שדה **אחר** שנשאבה **לפני** אותה מוטציה עלולה להתיישן.
   * הכפיל מדמה זאת: `fields.rebuild` נכשל עבור `fieldId: 'b'` (כאילו הכתובת
   * שלו התיישנה), וברענון שאחרי הכשל השדה השלישי (`c`) מופיע לראשונה —
   * ובכל זאת מתעדכן, כי הלולאה לא עצרה ולא הסתמכה על התצלום הישן.
   */
  it('אחרי כשל: הרשימה מתרעננת, ושדה שמתגלה רק ברענון עדיין מתעדכן', async () => {
    const a = { fieldId: 'a', blockId: 'b1', occurrenceIndex: 0, nestingDepth: 0 };
    const b = { fieldId: 'b', blockId: 'b1', occurrenceIndex: 1, nestingDepth: 0 };
    const c = { fieldId: 'c', blockId: 'b2', occurrenceIndex: 0, nestingDepth: 0 };
    const engine = fakeEngine({
      fieldsSequence: [
        [{ address: a }, { address: b }],
        // אחרי הכשל על `b`: `a` כבר נוסה (מדולג), `b` עדיין שם (ייכשל שוב —
        // לא נוסה בשנית), ו-`c` מתגלה לראשונה.
        [{ address: a }, { address: b }, { address: c }],
      ],
      rebuildFailFieldIds: ['b'],
    });

    expect(await rebuildAllFields(engine.host)).toEqual({
      ok: false,
      message: 'עדכון השדות לא הושלם: 2 שדות עודכנו, ושדה אחד נכשל ולא עודכן',
      reason: 'partial-rebuild',
    });
    expect(engine.inputs('fields.rebuild')).toEqual([{ target: a }, { target: b }, { target: c }]);
    // רשימה טרייה נשאבה רק אחרי הכשל על `b`, לא אחרי כל שדה.
    expect(engine.inputs('fields.list')).toHaveLength(2);
  });

  it('`fields.rebuild` חסר בפאסדה: „אינו זמין בגרסה זו”', async () => {
    const engine = fakeEngine({ missing: ['fields.rebuild'] });

    expect(await rebuildAllFields(engine.host)).toEqual({
      ok: false,
      message: 'עדכון השדות נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });

  /**
   * הבדיקה שמגנה על מסמך גדול: `list` מחזיר עמוד, ו„עדכן שדות” שרץ על העמוד
   * הראשון בלבד היה משאיר שדות לא מעודכנים **ומדווח הצלחה** — כשל שקט שאיש
   * לא היה מבחין בו עד שמישהו פותח את הקובץ ב-Word.
   */
  it('שואב את כל העמודים, ולא רק את הראשון', async () => {
    const fields = Array.from({ length: 250 }, (_, index) => ({ address: { fieldId: `f${index}` } }));
    const engine = fakeEngine({ fields, total: fields.length });

    expect(await rebuildAllFields(engine.host)).toEqual({ ok: true });

    expect(engine.inputs('fields.rebuild')).toHaveLength(250);
    // שני עמודים: 200 ואז 50. עמוד אחד היה מעדכן 200 שדות ומדווח „הצלחה”.
    expect(engine.inputs('fields.list')).toEqual([
      { limit: 200, offset: 0 },
      { limit: 200, offset: 200 },
    ]);
  });

  it('חריגה מ-`fields.list` הופכת להודעה', async () => {
    const engine = fakeEngine({ throws: ['fields.list'] });

    expect(await rebuildAllFields(engine.host)).toEqual({
      ok: false,
      message: 'עדכון השדות נכשל: boom',
      reason: 'threw',
    });
  });
});

describe('readFieldsState', () => {
  it('מונה את שדות המסמך', async () => {
    const engine = fakeEngine({ fields: [{ address: {} }, { address: {} }, { address: {} }] });

    expect(await readFieldsState(engine.host)).toEqual({ count: 3 });
  });

  it('`total` גובר על אורך העמוד', async () => {
    // `fields.list` הוא `DiscoveryOutput`: `items` הוא עמוד תחת `limit`/`offset`,
    // ו-`total` הוא המספר במסמך. ספירת העמוד הייתה אומרת „2 שדות” במסמך של 40.
    const engine = fakeEngine({ fields: [{ address: {} }, { address: {} }], total: 40 });

    expect(await readFieldsState(engine.host)).toEqual({ count: 40 });
  });

  it('גרסה בלי `total` נופלת לאורך העמוד ולא לאפס', async () => {
    // מספר חלקי הוא עדיין מידע; אפס היה מכבה את „עדכן שדות” על מסמך שיש בו שדות.
    const engine = fakeEngine({ fields: [{ address: {} }, { address: {} }] });

    expect(await readFieldsState(engine.host)).toEqual({ count: 2 });
  });

  it('כשל של קריאה מחזיר „אין שדות” ואינו זורק', async () => {
    // ה-tooltip יאמר „אין מה לעדכן”, וזה עדיף על מספר שהומצא.
    expect(await readFieldsState(fakeEngine({ throws: ['fields.list'] }).host)).toEqual(
      emptyFieldsState(),
    );
    expect(await readFieldsState(fakeEngine({ missing: ['fields.list'] }).host)).toEqual({
      count: 0,
    });
    expect(await readFieldsState(null)).toEqual({ count: 0 });
  });
});

describe('insertDateOnly / insertTimeOnly — מחרוזת ההוראה המדויקת', () => {
  it('„תאריך בלבד” שולח `DATE \\@ "dd/MM/yyyy"` בסמן', async () => {
    const engine = fakeEngine();

    expect(await insertDateOnly(engine.host)).toEqual({ ok: true });

    expect(engine.inputs('fields.insert')).toEqual([
      {
        at: { kind: 'text', segments: [{ blockId: 'block-1', range: { start: 3, end: 3 } }] },
        instruction: 'DATE \\@ "dd/MM/yyyy"',
        mode: 'raw',
      },
    ]);
  });

  it('„שעה בלבד” שולחת `DATE \\@ "HH:mm"` בסמן', async () => {
    const engine = fakeEngine();

    expect(await insertTimeOnly(engine.host)).toEqual({ ok: true });

    expect(engine.inputs('fields.insert')).toEqual([
      {
        at: { kind: 'text', segments: [{ blockId: 'block-1', range: { start: 3, end: 3 } }] },
        instruction: 'DATE \\@ "HH:mm"',
        mode: 'raw',
      },
    ]);
  });
});
