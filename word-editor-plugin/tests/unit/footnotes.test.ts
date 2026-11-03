/**
 * הוספת הערת שוליים והערת סיום.
 *
 * מה שנבדק כאן הוא בעיקר מה **לא** נשלח: `at` נשאר בחוץ כדי שההוספה תיפול על
 * הסמן (המסלול שהחוזה מגדיר לסרגל), ו-`content` הוא מחרוזת ריקה ולא טקסט
 * מקום. ומעל הכול — ש-namespace חסר מדווח „אינו זמין בגרסה זו” ולא מנסה מסלול
 * חלופי.
 */
import { describe, expect, it } from 'vitest';
import {
  insertNote,
  listNotes,
  readNotesState,
  removeNote,
  updateNote,
  type FootnotesHost,
  type NoteType,
} from '../../src/engine/footnotes';

function fakeEngine(receipt: () => unknown = () => ({ success: true })) {
  const calls: unknown[] = [];
  const host: FootnotesHost = {
    activeEditor: {
      doc: {
        footnotes: {
          insert: (input) => {
            calls.push(input);
            // הוולידציה של המנוע: type מבין השניים, ו-content מחרוזת.
            if (input.type !== 'footnote' && input.type !== 'endnote') {
              throw new Error("footnotes.insert requires a type of 'footnote' or 'endnote'.");
            }
            if (typeof input.content !== 'string') {
              throw new Error('footnotes.insert requires a content string.');
            }
            return receipt() as never;
          },
        },
      },
    },
  };
  return { host, calls };
}

describe('insertNote', () => {
  it('שולחת type ו-content בלבד — בלי `at`, כדי שההוספה תיפול על הסמן', async () => {
    const { host, calls } = fakeEngine();

    expect(await insertNote(host, 'footnote')).toEqual({ ok: true });
    expect(calls).toEqual([{ type: 'footnote', content: '' }]);
  });

  it('הערת סיום היא אותה פעולה עם type אחר', async () => {
    const { host, calls } = fakeEngine();

    expect(await insertNote(host, 'endnote')).toEqual({ ok: true });
    expect(calls).toEqual([{ type: 'endnote', content: '' }]);
  });

  it('namespace חסר מדווח „אינו זמין בגרסה זו” ואינו מנסה מסלול אחר', async () => {
    for (const host of [
      null,
      undefined,
      {},
      { activeEditor: null },
      { activeEditor: { doc: null } },
      { activeEditor: { doc: {} } },
      { activeEditor: { doc: { footnotes: {} } } },
    ]) {
      const outcome = await insertNote(host as FootnotesHost, 'footnote');

      expect(outcome).toEqual({
        ok: false,
        message: 'הוספת הערת שוליים נכשלה: אינו זמין בגרסה זו',
        reason: 'command-unsupported',
      });
    }
  });

  it('קבלה שנכשלה מתורגמת לעברית', async () => {
    const { host } = fakeEngine(() => ({ success: false, failure: { code: 'NO_SELECTION' } }));

    expect(await insertNote(host, 'footnote')).toEqual({
      ok: false,
      message: 'הוספת הערת שוליים נכשלה: יש למקם את הסמן במסמך',
      reason: 'NO_SELECTION',
    });
  });

  it('פעולה שזורקת מדווחת ואינה מפילה את התוסף', async () => {
    const { host } = fakeEngine(() => {
      throw new Error('boom');
    });

    expect(await insertNote(host, 'endnote')).toEqual({
      ok: false,
      message: 'הוספת הערת סיום נכשלה: boom',
      reason: 'threw',
    });
  });

  it('סובלת קבלה סינכרונית וקבלה כהבטחה', async () => {
    const sync = fakeEngine(() => ({ success: true }));
    expect(await insertNote(sync.host, 'footnote')).toEqual({ ok: true });

    const async = fakeEngine(() => Promise.resolve({ success: true }));
    expect(await insertNote(async.host, 'footnote')).toEqual({ ok: true });
  });
});

/* ------------------------------------------------------------------ */
/* ניהול ההערות שכבר במסמך — גל 9                                      */
/* ------------------------------------------------------------------ */

/**
 * כפיל שמדגם את מה שנמדד במנוע, ולא את מה שנוח לבדוק:
 *
 * 1. **שני רצפי מזהים.** הערות שוליים והערות סיום ממוספרות בנפרד, ולכן
 *    „1” הוא כתובת של שתי הערות שונות. זה המסלול שכל ההגנה כאן נבנתה בשבילו.
 * 2. **הכתובת נפתרת להערת השוליים תחילה**, ובהיעדרה להערת הסיום.
 * 3. **`get` זורק** על כתובת שאינה קיימת ואינו מחזיר קבלה.
 * 4. **`update` מחליף** את התוכן ואינו מוסיף עליו.
 * 5. **`list` מכבד `limit`/`offset`** ומחזיר `total` מלא.
 */
function fakeNotes(
  seed: readonly { type: NoteType; content: string }[],
  options: {
    receipt?: () => unknown;
    missing?: readonly ('list' | 'get' | 'update' | 'remove')[];
    pageSize?: number;
    /**
     * מאיזו שאיבה `list` **זורק**. כשל אמצע-שאיבה הוא המצב שבו העמוד הראשון
     * חזר והשני לא, והוא אינו תיאורטי: `total` מדווח את כל ההערות, וספר עם
     * מאות מהן נשאב בכמה סבבים. בלי האפשרות הזאת אין דרך לבדוק מה מוצג
     * כשחלק מהרשימה חסר.
     */
    failFromPage?: number;
  } = {},
) {
  let footnoteSeq = 0;
  let endnoteSeq = 0;
  const notes = seed.map((item) => ({
    type: item.type,
    noteId: String(item.type === 'endnote' ? ++endnoteSeq : ++footnoteSeq),
    content: item.content,
  }));
  const calls: { op: string; input?: unknown }[] = [];
  const missing = new Set(options.missing ?? []);
  const receipt = options.receipt ?? (() => ({ success: true }));

  const resolve = (noteId: string) =>
    notes.find((note) => note.type === 'footnote' && note.noteId === noteId) ??
    notes.find((note) => note.noteId === noteId);

  const shape = (note: (typeof notes)[number]) => ({
    address: { kind: 'entity', entityType: 'footnote', noteId: note.noteId },
    type: note.type,
    noteId: note.noteId,
    displayNumber: note.noteId,
    content: note.content,
  });

  const api: Record<string, unknown> = {
    insert: () => ({ success: true }),
  };
  if (!missing.has('list')) {
    api.list = (query?: { limit?: number; offset?: number }) => {
      calls.push({ op: 'list', input: query });
      const offset = query?.offset ?? 0;
      const limit = Math.min(query?.limit ?? notes.length, options.pageSize ?? notes.length);
      // לפי ה-`offset` ולא לפי מונה קריאות: הכפיל משמש כמה קריאות באותה
      // בדיקה, ומונה היה מפיל גם את השאיבה הראשונה של השנייה שבהן.
      if (options.failFromPage !== undefined && offset >= options.failFromPage * limit) {
        throw new Error('footnotes.list failed.');
      }
      return { items: notes.slice(offset, offset + limit).map(shape), total: notes.length };
    };
  }
  if (!missing.has('get')) {
    api.get = (input: { target: { noteId: string } }) => {
      calls.push({ op: 'get', input });
      const found = resolve(input.target.noteId);
      if (!found) throw new Error('footnote/endnote was not found.');
      return shape(found);
    };
  }
  if (!missing.has('update')) {
    api.update = (input: { target: { noteId: string }; patch: { content: string } }) => {
      calls.push({ op: 'update', input });
      const found = resolve(input.target.noteId);
      if (found) found.content = input.patch.content;
      return receipt();
    };
  }
  if (!missing.has('remove')) {
    api.remove = (input: { target: { noteId: string } }) => {
      calls.push({ op: 'remove', input });
      const found = resolve(input.target.noteId);
      if (found) notes.splice(notes.indexOf(found), 1);
      return receipt();
    };
  }

  const host = { activeEditor: { doc: { footnotes: api } } } as unknown as FootnotesHost;
  return { host, calls, notes, ops: () => calls.map((call) => call.op) };
}

describe('listNotes', () => {
  it('מחזירה את שני הסוגים, ואת התצוגה כפי שהיא במסמך', async () => {
    const { host } = fakeNotes([
      { type: 'footnote', content: 'רַשִׁ״י שם' },
      { type: 'endnote', content: 'מקורות' },
    ]);

    expect(await listNotes(host)).toEqual([
      {
        id: '1',
        type: 'footnote',
        number: '1',
        content: 'רַשִׁ״י שם',
        display: 'הערת שוליים 1: רַשִׁ״י שם',
      },
      { id: '1', type: 'endnote', number: '1', content: 'מקורות', display: 'הערת סיום 1: מקורות' },
    ]);
  });

  it('שואבת עמודים עד `total` — ספר עם מאות הערות אינו עמוד אחד', async () => {
    const seed = Array.from({ length: 7 }, (_, index) => ({
      type: 'footnote' as NoteType,
      content: `הערה ${index + 1}`,
    }));
    const { host, calls } = fakeNotes(seed, { pageSize: 3 });

    const notes = await listNotes(host);

    expect(notes).toHaveLength(7);
    expect(notes.map((note) => note.content)).toEqual(seed.map((item) => item.content));
    expect(calls.filter((call) => call.op === 'list')).toHaveLength(3);
  });

  it('כשל בעמוד השני מחזיר ריק — רשימה חלקית אינה מוצגת כמלאה', async () => {
    // הערת שוליים 1 בעמוד הראשון, הערת סיום 1 בשני. אילו החלקית הייתה
    // מוצגת, הדיאלוג היה מציע לערוך הערת סיום בלי לדעת שיש לה תאומה —
    // והמשתמש היה מקבל סירוב „אינו יודע להבדיל ביניהן” בלי שום סימן מקדים.
    const seed = [
      { type: 'footnote' as NoteType, content: 'רַשִׁ״י שם' },
      { type: 'endnote' as NoteType, content: 'מקורות' },
    ];
    const { host, calls } = fakeNotes(seed, { pageSize: 1, failFromPage: 1 });

    expect(await listNotes(host)).toEqual([]);
    expect(await readNotesState(host)).toEqual({ notes: [] });
    // ושתי השאיבות אכן קרו בכל קריאה — אחרת הבדיקה מודדת אוויר.
    expect(calls.filter((call) => call.op === 'list')).toHaveLength(4);
  });

  it('הערה בלי מזהה או בלי סוג מוכר מדולגת ואינה „הערת שוליים”', async () => {
    const host = {
      activeEditor: {
        doc: {
          footnotes: {
            list: () => ({
              items: [
                { noteId: '', type: 'footnote', content: 'בלי מזהה' },
                { noteId: '2', type: 'zigzag', content: 'סוג שאינו בחוזה' },
                { noteId: '3', type: 'footnote', displayNumber: '3', content: 'תקינה' },
              ],
              total: 3,
            }),
          },
        },
      },
    } as unknown as FootnotesHost;

    expect((await listNotes(host)).map((note) => note.id)).toEqual(['3']);
  });

  it('פעולה חסרה או קריאה שזרקה מחזירות רשימה ריקה ואינן מפילות', async () => {
    const { host } = fakeNotes([{ type: 'footnote', content: 'א' }], { missing: ['list'] });
    expect(await listNotes(host)).toEqual([]);
    expect(await readNotesState(null)).toEqual({ notes: [] });

    const throws = {
      activeEditor: {
        doc: {
          footnotes: {
            list: () => {
              throw new Error('boom');
            },
          },
        },
      },
    } as unknown as FootnotesHost;
    expect(await listNotes(throws)).toEqual([]);
  });
});

describe('updateNote', () => {
  it('מחליפה את התוכן בקריאה אחת — לא הסרה והוספה מחדש', async () => {
    const { host, calls, ops } = fakeNotes([{ type: 'footnote', content: 'הישן' }]);

    expect(await updateNote(host, { id: '1', type: 'footnote' }, 'החדש')).toEqual({ ok: true });
    expect(ops()).toEqual(['get', 'update']);
    expect(calls[1].input).toEqual({
      target: { kind: 'entity', entityType: 'footnote', noteId: '1' },
      patch: { content: 'החדש' },
    });
  });

  it('ירידת שורה מכווצת לרווח, ורווחי קצה יורדים', async () => {
    const { host, calls } = fakeNotes([{ type: 'footnote', content: 'א' }]);

    await updateNote(host, { id: '1', type: 'footnote' }, '  שורה\nשנייה  ');

    expect((calls[1].input as { patch: { content: string } }).patch.content).toBe('שורה שנייה');
  });

  it('תוכן ריק נדחה **לפני** הקריאה — המנוע היה מרוקן את ההערה', async () => {
    const { host, ops } = fakeNotes([{ type: 'footnote', content: 'הישן' }]);

    for (const empty of ['', '   ', '\n']) {
      expect(await updateNote(host, { id: '1', type: 'footnote' }, empty)).toEqual({
        ok: false,
        message: 'עריכת הערת השוליים נכשלה: יש להקליד את תוכן ההערה',
        reason: 'invalid-content',
      });
    }
    expect(ops()).toEqual([]);
  });

  it('הערת סיום שיש הערת שוליים באותו מספר — מסרבת ואינה נוגעת במסמך', async () => {
    const { host, ops, notes } = fakeNotes([
      { type: 'footnote', content: 'שוליים' },
      { type: 'endnote', content: 'סיום' },
    ]);

    const outcome = await updateNote(host, { id: '1', type: 'endnote' }, 'חדש');

    expect(outcome).toEqual({
      ok: false,
      message:
        'עריכת הערת הסיום נכשלה: הערת הסיום נושאת את אותו מספר כמו הערת שוליים שבמסמך, והמנוע אינו יודע להבדיל ביניהן',
      reason: 'note-ambiguous',
    });
    expect(ops()).toEqual(['get']);
    // וזה העיקר: הערת השוליים לא נגעה בה — היא זו שהייתה נערכת בלי הבדיקה.
    expect(notes.map((note) => note.content)).toEqual(['שוליים', 'סיום']);
  });

  it('הערת סיום שאין לה תאומה נערכת כרגיל', async () => {
    const { host, notes } = fakeNotes([{ type: 'endnote', content: 'סיום' }]);

    expect(await updateNote(host, { id: '1', type: 'endnote' }, 'סיום מתוקן')).toEqual({ ok: true });
    expect(notes[0].content).toBe('סיום מתוקן');
  });

  it('הערה שכבר אינה במסמך מדווחת בעברית, ולא בהודעת המנוע', async () => {
    const { host } = fakeNotes([{ type: 'footnote', content: 'א' }]);

    expect(await updateNote(host, { id: '9', type: 'footnote' }, 'חדש')).toEqual({
      ok: false,
      message: 'עריכת הערת השוליים נכשלה: ההערה אינה נמצאת במסמך',
      reason: 'note-not-found',
    });
  });

  it('פעולה חסרה, קבלה שנכשלה וחריגה — כולן מדווחות ולא מפילות', async () => {
    const missing = fakeNotes([{ type: 'footnote', content: 'א' }], { missing: ['update'] });
    expect(await updateNote(missing.host, { id: '1', type: 'footnote' }, 'חדש')).toEqual({
      ok: false,
      message: 'עריכת הערת השוליים נכשלה: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });

    const failed = fakeNotes([{ type: 'footnote', content: 'א' }], {
      receipt: () => ({ success: false, failure: { code: 'TARGET_NOT_FOUND' } }),
    });
    expect(await updateNote(failed.host, { id: '1', type: 'footnote' }, 'חדש')).toEqual({
      ok: false,
      message: 'עריכת הערת השוליים נכשלה: היעד של הפעולה לא נמצא במסמך',
      reason: 'TARGET_NOT_FOUND',
    });

    const threw = fakeNotes([{ type: 'footnote', content: 'א' }], {
      receipt: () => {
        throw new Error('boom');
      },
    });
    expect(await updateNote(threw.host, { id: '1', type: 'footnote' }, 'חדש')).toEqual({
      ok: false,
      message: 'עריכת הערת השוליים נכשלה: boom',
      reason: 'threw',
    });

    expect(await updateNote(null, { id: '1', type: 'footnote' }, 'חדש')).toEqual({
      ok: false,
      message: 'עריכת הערת השוליים נכשלה: המסמך עדיין נטען',
      reason: 'document-api-unavailable',
    });
  });
});

describe('removeNote', () => {
  it('מסירה את ההערה בצעד אחד, אחרי שאימתה לאיזו הערה הכתובת נפתרת', async () => {
    const { host, calls, ops, notes } = fakeNotes([
      { type: 'footnote', content: 'ראשונה' },
      { type: 'footnote', content: 'שנייה' },
    ]);

    expect(await removeNote(host, { id: '2', type: 'footnote' })).toEqual({ ok: true });
    expect(ops()).toEqual(['get', 'remove']);
    expect(calls[1].input).toEqual({
      target: { kind: 'entity', entityType: 'footnote', noteId: '2' },
    });
    expect(notes.map((note) => note.content)).toEqual(['ראשונה']);
  });

  it('„הסר” על הערת סיום שיש לה תאומה אינו מוחק את הערת השוליים', async () => {
    const { host, ops, notes } = fakeNotes([
      { type: 'footnote', content: 'שוליים' },
      { type: 'endnote', content: 'סיום' },
    ]);

    const outcome = await removeNote(host, { id: '1', type: 'endnote' });

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('note-ambiguous');
    expect(ops()).toEqual(['get']);
    expect(notes).toHaveLength(2);
  });

  it('הערת שוליים באותו מספר כן מוסרת — היא זו שהכתובת נפתרת אליה', async () => {
    const { host, notes } = fakeNotes([
      { type: 'footnote', content: 'שוליים' },
      { type: 'endnote', content: 'סיום' },
    ]);

    expect(await removeNote(host, { id: '1', type: 'footnote' })).toEqual({ ok: true });
    expect(notes.map((note) => note.type)).toEqual(['endnote']);
  });

  it('בלי מזהה, בלי פעולה, ובלי מסמך — כולן מסרבות בעברית', async () => {
    const { host, ops } = fakeNotes([{ type: 'footnote', content: 'א' }]);
    expect(await removeNote(host, { id: '', type: 'footnote' })).toEqual({
      ok: false,
      message: 'הסרת הערת השוליים נכשלה: יש לבחור הערה',
      reason: 'no-note',
    });
    expect(ops()).toEqual([]);

    const missing = fakeNotes([{ type: 'endnote', content: 'א' }], { missing: ['remove'] });
    expect(await removeNote(missing.host, { id: '1', type: 'endnote' })).toEqual({
      ok: false,
      message: 'הסרת הערת הסיום נכשלה: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });

    expect(await removeNote(null, { id: '1', type: 'endnote' })).toEqual({
      ok: false,
      message: 'הסרת הערת הסיום נכשלה: המסמך עדיין נטען',
      reason: 'document-api-unavailable',
    });
  });
});
