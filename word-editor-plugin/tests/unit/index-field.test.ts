/**
 * מפתח ערכים — סימון ערך, הוספה, עדכון, הסרה והתאמה אישית.
 *
 * חמש הטענות שנמדדות כאן, וכולן מקבעות ממצא שנמדד בדפדפן ולא הנחה (ההנמקה
 * המלאה ב-engine/index-field.ts):
 *
 * 1. **תת-ערך נשלח כנקודתיים בתוך הטקסט ולא כ-`subEntry`.** `IndexEntryData`
 *    מציע `subEntry`, המנוע מקבל אותו בהצלחה — וכותב `\s`, מתג שאינו קיים
 *    בשדה `XE` של Word. הצורה הקנונית `XE "ראשי:משני"` נמדדה כמתפרקת בחזרה
 *    ל-`text`+`subEntry` במנוע עצמו. הבדיקה מקבעת שזה מה שנשלח.
 * 2. **הוולידציה יושבת אצלנו, כי המנוע בולע.** טקסט של רווחים בלבד חוזר
 *    ממנו `success: true` וכותב ערך בלתי נראה, ו-`columns` שאינו מספר שלם
 *    חיובי נבלע בשקט. שניהם נדחים כאן לפני שנוגעים במסמך.
 * 3. **ההסרה היא צעד אחד.** המפתח הוא בלוק יחיד, ומימוש שהיה מנקה „שיירים”
 *    בעקבות תוכן העניינים היה מוחק פסקאות של המשתמש. הבדיקה דורשת שלא
 *    נשלחת שום קריאה ל-`blocks.*`.
 * 4. **דו-משמעות מוחזרת כהודעה ולא כניחוש** בהסרה ובהתאמה אישית, ואילו
 *    עדכון רץ על כולם — כי בנייה מחדש אינה הרסנית, ונמדד ששני מפתחות מקבלים
 *    כתובות שונות.
 * 5. **המודול לעולם אינו זורק.** חריגה, קבלה שנכשלה, ופעולה שאינה קיימת
 *    בגרסת המנוע — שלושתן `CommandOutcome`.
 */
import { describe, expect, it } from 'vitest';
import {
  DEFAULT_INDEX_COLUMNS,
  INDEX_COLUMNS_HINT,
  buildIndexEntryText,
  configureIndex,
  emptyIndexState,
  insertIndex,
  isValidIndexColumns,
  listIndexEntries,
  markIndexEntry,
  normalizeIndexEntryText,
  readIndexState,
  rebuildIndex,
  removeIndex,
  removeIndexEntry,
  type IndexHost,
} from '../../src/engine/index-field';

interface Call {
  op: string;
  input?: unknown;
}

interface FakeEntry {
  blockId: string;
  offset: number;
  text?: string;
  subEntry?: string;
  /** ערך בלי כתובת — מה שהמנוע עלול להחזיר, ומה שאסור להציג. */
  addressless?: boolean;
}

interface FakeOptions {
  /** המפתחות שבמסמך, לפי מזהה הבלוק שלהם. */
  indexes?: readonly string[];
  /** מה שכל מפתח מצהיר עליו. נקרא רק מהראשון. */
  config?: unknown;
  /** `total` שהמנוע מדווח, כשהוא שונה מאורך הרשימה — כלומר שאיבת עמודים. */
  total?: number;
  entries?: readonly FakeEntry[];
  entriesTotal?: number;
  /** מה `selection.current` מדווח. */
  blockId?: string | null;
  selectionText?: string;
  failures?: Record<string, { code: string; message?: string }>;
  throws?: readonly string[];
  /**
   * זריקה רק מהקריאה ה-N ואילך. נדרש כדי למדוד כשל **חלקי** של שאיבת
   * עמודים: עמוד ראשון שנקרא ועמוד שני שנפל הוא המצב שבו ספירה חלקית
   * הייתה מגיעה לממשק כאילו היא המספר האמיתי.
   */
  throwsFrom?: Record<string, number>;
  missing?: readonly string[];
}

function fakeEngine(options: FakeOptions = {}) {
  const calls: Call[] = [];
  const missing = new Set(options.missing ?? []);
  const throwing = new Set(options.throws ?? []);
  const throwsFrom = options.throwsFrom ?? {};
  const failures = options.failures ?? {};
  const seen: Record<string, number> = {};

  function route<T>(op: string, impl: (input: unknown) => T): ((input: unknown) => T) | undefined {
    if (missing.has(op)) return undefined;
    return (input: unknown) => {
      calls.push({ op, input });
      seen[op] = (seen[op] ?? 0) + 1;
      if (throwing.has(op)) throw new Error('boom');
      if (throwsFrom[op] !== undefined && seen[op] >= throwsFrom[op]) throw new Error('boom');
      return impl(input);
    };
  }

  const receipt = (op: string): { success: boolean; failure?: { code: string; message?: string } } =>
    failures[op] ? { success: false, failure: failures[op] } : { success: true };

  /** עמוד מתוך רשימה, בדיוק כמו `DiscoveryOutput` של המנוע. */
  const page = <T>(all: readonly T[], input: unknown, total: number | undefined) => {
    const query = (input ?? {}) as { limit?: number; offset?: number };
    const offset = query.offset ?? 0;
    const end = query.limit === undefined ? undefined : offset + query.limit;
    return { items: all.slice(offset, end), ...(total === undefined ? {} : { total }) };
  };

  const blockId = options.blockId === undefined ? 'block-1' : options.blockId;
  const selectionText = options.selectionText ?? 'טקסט נבחר';
  /** ה-`TextTarget` כפי שהמנוע מחזיר אותו, ומה ש-`at` אמור לקבל כמו שהוא. */
  const selectionTarget = {
    kind: 'text',
    story: { kind: 'body' },
    segments: blockId ? [{ blockId, range: { start: 0, end: selectionText.length } }] : [],
  };

  const doc = {
    index: {
      list: route('index.list', (input) =>
        page(
          (options.indexes ?? []).map((nodeId) => ({
            id: nodeId,
            address: { kind: 'block', nodeType: 'index', nodeId },
            config: options.config ?? {},
          })),
          input,
          options.total ?? (options.indexes ?? []).length,
        ),
      ),
      insert: route('index.insert', () => receipt('index.insert')),
      configure: route('index.configure', () => receipt('index.configure')),
      rebuild: route('index.rebuild', () => receipt('index.rebuild')),
      remove: route('index.remove', () => receipt('index.remove')),
      entries: {
        list: route('index.entries.list', (input) =>
          page(
            (options.entries ?? []).map((entry) => ({
              id: `${entry.blockId}#${entry.offset}`,
              ...(entry.addressless
                ? {}
                : {
                    address: {
                      kind: 'inline',
                      nodeType: 'indexEntry',
                      anchor: {
                        start: { blockId: entry.blockId, offset: entry.offset },
                        end: { blockId: entry.blockId, offset: entry.offset + 1 },
                      },
                    },
                  }),
              text: entry.text,
              subEntry: entry.subEntry,
            })),
            input,
            options.entriesTotal ?? (options.entries ?? []).length,
          ),
        ),
        insert: route('index.entries.insert', () => receipt('index.entries.insert')),
        remove: route('index.entries.remove', () => receipt('index.entries.remove')),
      },
    },
    selection: {
      current: route('selection.current', () => ({
        empty: !blockId,
        target: selectionTarget,
        text: selectionText,
      })),
    },
  };

  return { host: { activeEditor: { doc } } as unknown as IndexHost, calls, selectionTarget };
}

const ops = (calls: readonly Call[]): string[] => calls.map((call) => call.op);
const inputs = (calls: readonly Call[], op: string): unknown[] =>
  calls.filter((call) => call.op === op).map((call) => call.input);

/* ------------------------------------------------------------------ */
/* ולידציה                                                             */
/* ------------------------------------------------------------------ */

describe('ולידציה של קלט', () => {
  it('טקסט של רווחים בלבד אינו ערך', () => {
    // נמדד: המנוע **זורק** על `''` אבל מקבל בהצלחה את `'   '` וכותב
    // `XE "   "` — ערך בלתי נראה שאי אפשר למצוא.
    expect(normalizeIndexEntryText('   ')).toBeNull();
    expect(normalizeIndexEntryText('')).toBeNull();
    expect(normalizeIndexEntryText('  רש״י  ')).toBe('רש״י');
  });

  it('תת-ערך מקודד בנקודתיים, כמו ב-Word', () => {
    expect(buildIndexEntryText('אבות', 'אברהם')).toBe('אבות:אברהם');
    expect(buildIndexEntryText('אבות', '  ')).toBe('אבות');
    expect(buildIndexEntryText('אבות', '')).toBe('אבות');
  });

  it('מספר טורים חוקי הוא שלם שבין 1 ל-4', () => {
    // נמדד: `0`, `-5` ו-`2.5` חזרו מהמנוע `success: true` ופשוט לא הופיעו
    // ב-`instruction`. כלומר המשתמש היה לוחץ „אישור” ולא היה קורה כלום.
    expect(isValidIndexColumns(0)).toBe(false);
    expect(isValidIndexColumns(-5)).toBe(false);
    expect(isValidIndexColumns(2.5)).toBe(false);
    expect(isValidIndexColumns(99)).toBe(false);
    expect(isValidIndexColumns(1)).toBe(true);
    expect(isValidIndexColumns(DEFAULT_INDEX_COLUMNS)).toBe(true);
    expect(isValidIndexColumns(4)).toBe(true);
  });
});

/* ------------------------------------------------------------------ */
/* סימון ערך                                                           */
/* ------------------------------------------------------------------ */

describe('markIndexEntry', () => {
  it('שולח את הבחירה כמו שהיא, ואת תת-הערך בתוך הטקסט', async () => {
    const engine = fakeEngine({ selectionText: 'שבת' });

    expect(await markIndexEntry(engine.host, { text: 'שבת', subEntry: 'הדלקת נרות' })).toEqual({
      ok: true,
    });
    expect(inputs(engine.calls, 'index.entries.insert')).toEqual([
      { at: engine.selectionTarget, entry: { text: 'שבת:הדלקת נרות' } },
    ]);
  });

  it('ערך עברי מנוקד ועם גרשיים עובר שלם', async () => {
    // נמדד בדפדפן: `XE "רש״י"`, `XE "בְּרֵאשִׁית"` ו-`XE "ר' עקיבא"` נכתבו
    // אחד לאחד. שום דבר במסלול הזה אינו אמור לגעת בתווים.
    const engine = fakeEngine();
    await markIndexEntry(engine.host, { text: 'בְּרֵאשִׁית', subEntry: 'רש״י' });

    expect(inputs(engine.calls, 'index.entries.insert')).toEqual([
      { at: engine.selectionTarget, entry: { text: 'בְּרֵאשִׁית:רש״י' } },
    ]);
  });

  it('טקסט ריק נדחה לפני שנוגעים במנוע', async () => {
    const engine = fakeEngine();

    expect(await markIndexEntry(engine.host, { text: '   ', subEntry: '' })).toEqual({
      ok: false,
      message: 'סימון הערך נכשל: יש להקליד את טקסט הערך',
      reason: 'invalid-text',
    });
    expect(ops(engine.calls)).toEqual([]);
  });

  it('בלי טקסט מסומן במסמך — מסביר מה לעשות ואינו מנחש פסקה', async () => {
    // `index.entries.insert` מקבל `TextTarget`, כלומר טווח. זו גם ההתנהגות
    // של Word: „סמן ערך” מסמן את הטקסט שנבחר.
    const engine = fakeEngine({ blockId: null });

    expect(await markIndexEntry(engine.host, { text: 'שבת', subEntry: '' })).toEqual({
      ok: false,
      message: 'סימון הערך נכשל: יש לסמן במסמך את הטקסט שהערך יצביע אליו',
      reason: 'no-selection',
    });
    expect(ops(engine.calls)).not.toContain('index.entries.insert');
  });

  it('קבלה שנכשלה מתורגמת להודעה בעברית', async () => {
    const engine = fakeEngine({
      failures: { 'index.entries.insert': { code: 'TARGET_NOT_FOUND', message: 'no block' } },
    });

    const outcome = await markIndexEntry(engine.host, { text: 'שבת', subEntry: '' });
    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('TARGET_NOT_FOUND');
  });

  it('חריגה אינה יוצאת החוצה', async () => {
    const engine = fakeEngine({ throws: ['index.entries.insert'] });

    const outcome = await markIndexEntry(engine.host, { text: 'שבת', subEntry: '' });
    expect(outcome.ok === false && outcome.reason).toBe('threw');
  });

  it('פעולה שאינה בגרסת המנוע מדווחת כך', async () => {
    const engine = fakeEngine({ missing: ['index.entries.insert'] });

    expect(await markIndexEntry(engine.host, { text: 'שבת', subEntry: '' })).toEqual({
      ok: false,
      message: 'סימון הערך נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });

  it('מסמך שעדיין נטען אינו כשל סתום', async () => {
    expect(await markIndexEntry(null, { text: 'שבת', subEntry: '' })).toEqual({
      ok: false,
      message: 'סימון הערך נכשל: המסמך עדיין נטען',
      reason: 'document-api-unavailable',
    });
  });
});

describe('removeIndexEntry', () => {
  it('שולח את הכתובת שהתקבלה ברשימה, כמו שהיא', async () => {
    const engine = fakeEngine({ entries: [{ blockId: 'b1', offset: 0, text: 'שבת' }] });
    const [entry] = await listIndexEntries(engine.host);

    expect(await removeIndexEntry(engine.host, entry.address)).toEqual({ ok: true });
    expect(inputs(engine.calls, 'index.entries.remove')).toEqual([{ target: entry.address }]);
  });

  it('חריגה אינה יוצאת החוצה', async () => {
    const engine = fakeEngine({ throws: ['index.entries.remove'] });

    const outcome = await removeIndexEntry(engine.host, { kind: 'inline' });
    expect(outcome.ok === false && outcome.reason).toBe('threw');
  });

  it('פעולה שאינה בגרסת המנוע מדווחת כך', async () => {
    const engine = fakeEngine({ missing: ['index.entries.remove'] });

    expect(await removeIndexEntry(engine.host, {})).toEqual({
      ok: false,
      message: 'ביטול סימון הערך נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });
});

/* ------------------------------------------------------------------ */
/* קריאה                                                               */
/* ------------------------------------------------------------------ */

describe('readIndexState / listIndexEntries', () => {
  it('קוראת את המפתח, ההגדרות שלו והערכים שסומנו', async () => {
    const engine = fakeEngine({
      indexes: ['idx-1'],
      config: { columns: 3, runIn: true },
      entries: [{ blockId: 'b1', offset: 0, text: 'אבות', subEntry: 'אברהם' }],
    });

    expect(await readIndexState(engine.host)).toEqual({
      count: 1,
      columns: 3,
      runIn: true,
      entries: [
        {
          id: 'b1#0',
          text: 'אבות',
          subEntry: 'אברהם',
          address: {
            kind: 'inline',
            nodeType: 'indexEntry',
            anchor: { start: { blockId: 'b1', offset: 0 }, end: { blockId: 'b1', offset: 1 } },
          },
        },
      ],
    });
  });

  it('שואבת עמודים עד `total` ואינה נעצרת בעמוד הראשון', async () => {
    // ספר תורני עם מפתח ערכים הוא בדיוק המסמך שבו העמוד הראשון אינו הכול.
    const entries = Array.from({ length: 250 }, (_, index) => ({
      blockId: `b${index}`,
      offset: 0,
      text: `ערך ${index}`,
    }));
    const engine = fakeEngine({ entries, entriesTotal: 250 });

    expect((await listIndexEntries(engine.host)).length).toBe(250);
    expect(inputs(engine.calls, 'index.entries.list')).toEqual([
      { limit: 200, offset: 0 },
      { limit: 200, offset: 200 },
    ]);
  });

  it('ערך בלי כתובת אינו מוצג', async () => {
    // הצגתו הייתה מייצרת שורה שלחיצה עליה שולחת `undefined` למנוע.
    const engine = fakeEngine({
      entries: [
        { blockId: 'b1', offset: 0, text: 'תקין' },
        { blockId: 'b2', offset: 0, text: 'בלי כתובת', addressless: true },
      ],
    });

    expect((await listIndexEntries(engine.host)).map((entry) => entry.text)).toEqual(['תקין']);
  });

  it('כשל קריאה מוחזר כ„אין מפתח” ולא כמספר מומצא', async () => {
    const engine = fakeEngine({ indexes: ['idx-1'], throws: ['index.list'] });

    expect(await readIndexState(engine.host)).toEqual(emptyIndexState());
  });

  /**
   * כשל **חלקי** הוא המסוכן: `collectAll` מחזיר את העמודים שהספיק לקרוא גם
   * כשהוא נכשל, ובלי הבדיקה על `ok` הספירה החלקית הייתה מגיעה ל-tooltip
   * כאילו היא המספר האמיתי במסמך.
   */
  it('עמוד שני שנפל מחזיר „אין מפתח” ולא ספירה חלקית', async () => {
    const engine = fakeEngine({
      indexes: Array.from({ length: 250 }, (_, index) => `idx-${index}`),
      throwsFrom: { 'index.list': 2 },
    });

    expect(await readIndexState(engine.host)).toEqual(emptyIndexState());
  });

  it('מנוע בלי `index` מחזיר מצב ריק ואינו זורק', async () => {
    expect(await readIndexState(null)).toEqual(emptyIndexState());
    expect(await listIndexEntries(null)).toEqual([]);
  });
});

/* ------------------------------------------------------------------ */
/* הוספה                                                               */
/* ------------------------------------------------------------------ */

describe('insertIndex', () => {
  it('מכניס בסוף המסמך, עם ההגדרות כבר ביצירה', async () => {
    // נמדד: `insert` מקבל `config` וכותב אותו מיד ל-`instruction`. קריאה
    // אחת עדיפה על שתיים שהשנייה בהן עלולה להיכשל.
    const engine = fakeEngine();

    expect(await insertIndex(engine.host, { columns: 2, runIn: false })).toEqual({ ok: true });
    expect(inputs(engine.calls, 'index.insert')).toEqual([
      { at: { kind: 'documentEnd' }, config: { columns: 2, runIn: false } },
    ]);
  });

  /**
   * `NO_OP` היא קבלה של „לא היה מה לשנות”, לא של כשל. בלי הענף הזה משתמש
   * שמאשר את הדיאלוג בלי לשנות דבר היה מקבל „שינוי הגדרות המפתח נכשל”.
   */
  it('`NO_OP` היא הצלחה ולא כשל', async () => {
    const engine = fakeEngine({
      indexes: ['idx-1'],
      failures: { 'index.configure': { code: 'NO_OP' } },
    });

    expect(await configureIndex(engine.host, { columns: 2, runIn: false })).toEqual({ ok: true });
  });

  it('מספר טורים פסול נדחה לפני שנוגעים במנוע', async () => {
    const engine = fakeEngine();

    expect(await insertIndex(engine.host, { columns: 0, runIn: false })).toEqual({
      ok: false,
      message: `הוספת המפתח נכשלה: ${INDEX_COLUMNS_HINT}`,
      reason: 'invalid-columns',
    });
    expect(ops(engine.calls)).toEqual([]);
  });

  it('חריגה אינה יוצאת החוצה', async () => {
    const engine = fakeEngine({ throws: ['index.insert'] });

    const outcome = await insertIndex(engine.host, { columns: 2, runIn: false });
    expect(outcome.ok === false && outcome.reason).toBe('threw');
  });

  it('פעולה שאינה בגרסת המנוע מדווחת כך', async () => {
    const engine = fakeEngine({ missing: ['index.insert'] });

    expect(await insertIndex(engine.host, { columns: 2, runIn: false })).toEqual({
      ok: false,
      message: 'הוספת המפתח נכשלה: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });
});

/* ------------------------------------------------------------------ */
/* עדכון                                                               */
/* ------------------------------------------------------------------ */

describe('rebuildIndex', () => {
  it('בונה מחדש את כל המפתחות — נמדד שכתובותיהם שונות', async () => {
    const engine = fakeEngine({ indexes: ['idx-1', 'idx-2'] });

    expect(await rebuildIndex(engine.host)).toEqual({ ok: true });
    expect(inputs(engine.calls, 'index.rebuild')).toEqual([
      { target: { kind: 'block', nodeType: 'index', nodeId: 'idx-1' } },
      { target: { kind: 'block', nodeType: 'index', nodeId: 'idx-2' } },
    ]);
  });

  it('כתובת חוזרת נשלחת פעם אחת, והחוסר מדווח למשתמש', async () => {
    // הבלם עולה שורה, וגל 4 מדד מנוע שנותן לשני עצמים את אותו `nodeId`.
    const engine = fakeEngine({ indexes: ['idx-1', 'idx-1'] });

    expect(await rebuildIndex(engine.host)).toEqual({
      ok: false,
      message:
        'עדכון המפתח לא הושלם: יש במסמך כמה מפתחות שאינם ניתנים להבחנה זה מזה, ולכן אחד מהם לא עודכן',
      reason: 'ambiguous-index',
    });
    expect(inputs(engine.calls, 'index.rebuild').length).toBe(1);
  });

  it('מסמך בלי מפתח — כשל מנומק, לא הצלחה שקטה', async () => {
    const engine = fakeEngine();

    expect(await rebuildIndex(engine.host)).toEqual({
      ok: false,
      message: 'עדכון המפתח נכשל: אין במסמך מפתח',
      reason: 'no-index',
    });
  });

  it('נעצר בכשל הראשון ואינו ממשיך למפתח הבא', async () => {
    const engine = fakeEngine({
      indexes: ['idx-1', 'idx-2'],
      failures: { 'index.rebuild': { code: 'TARGET_NOT_FOUND' } },
    });

    const outcome = await rebuildIndex(engine.host);
    expect(outcome.ok === false && outcome.reason).toBe('TARGET_NOT_FOUND');
    expect(inputs(engine.calls, 'index.rebuild').length).toBe(1);
  });

  it('פעולה שאינה בגרסת המנוע מדווחת כך', async () => {
    const engine = fakeEngine({ indexes: ['idx-1'], missing: ['index.rebuild'] });

    expect(await rebuildIndex(engine.host)).toEqual({
      ok: false,
      message: 'עדכון המפתח נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });
});

/* ------------------------------------------------------------------ */
/* הסרה                                                                */
/* ------------------------------------------------------------------ */

describe('removeIndex', () => {
  it('צעד אחד בלבד — אין ניקוי שיירים, כי אין שיירים', async () => {
    // נמדד: המפתח הוא בלוק יחיד, ו-`index.remove` מפיל אותו כולו. ניקוי
    // שאין לו מה לנקות היה רק דרך למחוק פסקה של המשתמש.
    const engine = fakeEngine({ indexes: ['idx-1'] });

    expect(await removeIndex(engine.host)).toEqual({ ok: true });
    expect(inputs(engine.calls, 'index.remove')).toEqual([
      { target: { kind: 'block', nodeType: 'index', nodeId: 'idx-1' } },
    ]);
    expect(ops(engine.calls).some((op) => op.startsWith('blocks.'))).toBe(false);
  });

  it('שני מפתחות — מסרב ואינו מנחש איזה למחוק', async () => {
    const engine = fakeEngine({ indexes: ['idx-1', 'idx-2'] });

    expect(await removeIndex(engine.host)).toEqual({
      ok: false,
      message:
        'הסרת המפתח נכשלה: יש במסמך יותר ממפתח אחד, ואין דרך לדעת על איזה מהם הפעולה חלה',
      reason: 'ambiguous-index',
    });
    expect(ops(engine.calls)).not.toContain('index.remove');
  });

  it('מסמך בלי מפתח — כשל מנומק', async () => {
    const engine = fakeEngine();

    expect(await removeIndex(engine.host)).toEqual({
      ok: false,
      message: 'הסרת המפתח נכשלה: אין במסמך מפתח',
      reason: 'no-index',
    });
  });

  it('חריגה אינה יוצאת החוצה', async () => {
    const engine = fakeEngine({ indexes: ['idx-1'], throws: ['index.remove'] });

    const outcome = await removeIndex(engine.host);
    expect(outcome.ok === false && outcome.reason).toBe('threw');
  });

  it('פעולה שאינה בגרסת המנוע מדווחת כך', async () => {
    const engine = fakeEngine({ indexes: ['idx-1'], missing: ['index.remove'] });

    expect(await removeIndex(engine.host)).toEqual({
      ok: false,
      message: 'הסרת המפתח נכשלה: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });
});

/* ------------------------------------------------------------------ */
/* התאמה אישית                                                         */
/* ------------------------------------------------------------------ */

describe('configureIndex', () => {
  it('שולח את שני המתגים שנמדדו כהפיכים, ולא יותר', async () => {
    const engine = fakeEngine({ indexes: ['idx-1'] });

    expect(await configureIndex(engine.host, { columns: 3, runIn: true })).toEqual({ ok: true });
    expect(inputs(engine.calls, 'index.configure')).toEqual([
      {
        target: { kind: 'block', nodeType: 'index', nodeId: 'idx-1' },
        patch: { columns: 3, runIn: true },
      },
    ]);
  });

  /**
   * `NO_OP` היא קבלה של „לא היה מה לשנות”, לא של כשל. בלי הענף הזה משתמש
   * שמאשר את הדיאלוג בלי לשנות דבר היה מקבל „שינוי הגדרות המפתח נכשל”.
   */
  it('`NO_OP` היא הצלחה ולא כשל', async () => {
    const engine = fakeEngine({
      indexes: ['idx-1'],
      failures: { 'index.configure': { code: 'NO_OP' } },
    });

    expect(await configureIndex(engine.host, { columns: 2, runIn: false })).toEqual({ ok: true });
  });

  it('מספר טורים פסול נדחה לפני שנוגעים במנוע', async () => {
    const engine = fakeEngine({ indexes: ['idx-1'] });

    expect(await configureIndex(engine.host, { columns: 2.5, runIn: false })).toEqual({
      ok: false,
      message: `שינוי הגדרות המפתח נכשל: ${INDEX_COLUMNS_HINT}`,
      reason: 'invalid-columns',
    });
    expect(ops(engine.calls)).toEqual([]);
  });

  it('שני מפתחות — מסרב', async () => {
    const engine = fakeEngine({ indexes: ['idx-1', 'idx-2'] });

    const outcome = await configureIndex(engine.host, { columns: 2, runIn: false });
    expect(outcome.ok === false && outcome.reason).toBe('ambiguous-index');
    expect(ops(engine.calls)).not.toContain('index.configure');
  });

  it('קבלה שנכשלה מתורגמת להודעה בעברית', async () => {
    const engine = fakeEngine({
      indexes: ['idx-1'],
      failures: { 'index.configure': { code: 'INVALID_INPUT', message: 'nope' } },
    });

    const outcome = await configureIndex(engine.host, { columns: 2, runIn: false });
    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('INVALID_INPUT');
  });

  it('פעולה שאינה בגרסת המנוע מדווחת כך', async () => {
    const engine = fakeEngine({ indexes: ['idx-1'], missing: ['index.configure'] });

    expect(await configureIndex(engine.host, { columns: 2, runIn: false })).toEqual({
      ok: false,
      message: 'שינוי הגדרות המפתח נכשל: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });
});
