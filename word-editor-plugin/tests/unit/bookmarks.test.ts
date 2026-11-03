/**
 * סימניות.
 *
 * ארבע טענות שהמודול עומד או נופל עליהן, ולכן הן עיקר מה שנמדד כאן:
 *
 * 1. **הוולידציה של השם היא של Word, לא של המנוע.** המנוע אוכף „מחרוזת לא
 *    ריקה” ותו לא (נמדד בדפדפן, ראו engine/bookmarks.ts), ולכן שם עם רווח
 *    היה נכתב למסמך ונשבר רק כשהוא נפתח ב-Word. הוולידציה כאן היא מה שמונע
 *    את זה, ולכן היא נבדקת על התו — כולל, ובעיקר, ש**שם עברי מתקבל**.
 * 2. **שם פסול נעצר לפני המנוע.** לא „נשלח ונדחה”: אם הבדיקה תראה קריאה
 *    ל-`bookmarks.insert` על שם פסול, פירושו שהמסמך כבר נגוע.
 * 3. **הרשימה נשאבת עמוד אחר עמוד עד `total`.** `BookmarksListResult` הוא
 *    `DiscoveryOutput`; דיאלוג שמציג עמוד אחד היה מסתיר סימניות, והמשתמש היה
 *    מקבל „כבר קיים” על שם שאינו רואה.
 * 4. **המודול לעולם אינו זורק.** חריגה, קבלה שנכשלה, ופעולה שאינה קיימת
 *    בגרסת המנוע — שלושתן `CommandOutcome`.
 */
import { describe, expect, it } from 'vitest';
import {
  BOOKMARK_NAME_HINT,
  emptyBookmarksState,
  insertBookmark,
  normalizeBookmarkName,
  readBookmarks,
  removeBookmark,
  renameBookmark,
  type BookmarksHost,
} from '../../src/engine/bookmarks';

interface Call {
  op: string;
  input?: unknown;
}

interface FakeOptions {
  /** מה `bookmarks.list` מחזיק — כל המסמך, לא עמוד. */
  bookmarks?: readonly { name?: string }[];
  /** `total` של `DiscoveryOutput`. `undefined` = גרסה שאינה חושפת אותו. */
  total?: number;
  failures?: Record<string, { code: string; message?: string }>;
  throws?: readonly string[];
  missing?: readonly string[];
  /** מה `selection.current` מדווח. `null` = אין קטע שאפשר לפעול עליו. */
  blockId?: string | null;
}

function fakeEngine(options: FakeOptions = {}) {
  const calls: Call[] = [];
  const missing = new Set(options.missing ?? []);
  const throwing = new Set(options.throws ?? []);
  const failures = options.failures ?? {};

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

  const doc = {
    selection: {
      current: route('selection.current', () => ({ empty: true, target: selectionTarget })),
    },
    bookmarks: {
      // הכפיל מכבד `limit`/`offset`: כפיל שמתעלם מהעמוד היה מאשר בירוק גם
      // מימוש שקורא את העמוד הראשון בלבד. אותה החלטה כמו ב-fields.test.ts.
      list: route('bookmarks.list', (input) => {
        const all = options.bookmarks ?? [];
        const query = (input ?? {}) as { limit?: number; offset?: number };
        const offset = query.offset ?? 0;
        const end = query.limit === undefined ? undefined : offset + query.limit;
        return {
          items: all.slice(offset, end),
          ...(options.total === undefined ? {} : { total: options.total }),
        };
      }),
      insert: route('bookmarks.insert', () => receipt('bookmarks.insert')),
      rename: route('bookmarks.rename', () => receipt('bookmarks.rename')),
      remove: route('bookmarks.remove', () => receipt('bookmarks.remove')),
    },
  };

  const host = { activeEditor: { doc } } as unknown as BookmarksHost;
  const ops = (): string[] => calls.map((call) => call.op);
  const inputs = (op: string): unknown[] =>
    calls.filter((call) => call.op === op).map((call) => call.input);

  return { host, calls, ops, inputs };
}

describe('ולידציה של שם סימנייה', () => {
  it('שם עברי מתקבל — זו השאלה הקריטית לתוסף הזה', () => {
    expect(normalizeBookmarkName('פרק_ראשון')).toBe('פרק_ראשון');
    expect(normalizeBookmarkName('הקדמה')).toBe('הקדמה');
    expect(normalizeBookmarkName('סימן3')).toBe('סימן3');
    // מעורב עברית-לועזית, כמו שם של מקור לועזי בתוך חיבור עברי.
    expect(normalizeBookmarkName('פרקA')).toBe('פרקA');
  });

  it('שם מנוקד מתקבל — במאגר תורני זו הקלדה סבירה', () => {
    // סימני הניקוד הם תווים משולבים נפרדים (`\p{M}`). בלעדיהם בסט „שָׁלוֹם”
    // היה נדחה בעוד „שלום” עובר — הבדל שאינו נראה על המסך.
    expect(normalizeBookmarkName('שָׁלוֹם')).toBe('שָׁלוֹם');
    expect(normalizeBookmarkName('בְּרֵאשִׁית_א')).toBe('בְּרֵאשִׁית_א');
  });

  it('ספרה שאינה ASCII נדחית — הכלל של Word הוא ספרות בפועל', () => {
    // `\p{N}` היה מכניס אותה, והשם היה נשבר ב-Word.
    expect(normalizeBookmarkName('א١')).toBeNull();
    expect(normalizeBookmarkName('א1')).toBe('א1');
  });

  it('רווחים נדחים — Word אינו מקבל אותם, גם כשהמנוע כן', () => {
    expect(normalizeBookmarkName('שם עם רווח')).toBeNull();
    expect(normalizeBookmarkName('two words')).toBeNull();
  });

  it('שם שאינו מתחיל באות נדחה', () => {
    expect(normalizeBookmarkName('1מספר')).toBeNull();
    // קו תחתון פותח הוא ההחמרה **שלנו** על Word, לא כלל של Word: שם כזה הוא
    // סימנייה מוסתרת שם, ו-Word מייצר בעצמו שמות כאלה. ההנמקה ב-bookmarks.ts.
    expect(normalizeBookmarkName('_קו_תחתון')).toBeNull();
  });

  it('תווי פיסוק נדחים, וקו תחתון בגוף השם מתקבל', () => {
    expect(normalizeBookmarkName('סימנייה!')).toBeNull();
    expect(normalizeBookmarkName('א-ב')).toBeNull();
    expect(normalizeBookmarkName('פרק_ב_חלק_ג')).toBe('פרק_ב_חלק_ג');
  });

  it('הגבול הוא 40 תווים, כמו ב-Word', () => {
    expect(normalizeBookmarkName('א'.repeat(40))).toBe('א'.repeat(40));
    expect(normalizeBookmarkName('א'.repeat(41))).toBeNull();
  });

  it('רווחים בשוליים נגזרים ואינם פוסלים', () => {
    expect(normalizeBookmarkName('  הקדמה  ')).toBe('הקדמה');
    expect(normalizeBookmarkName('   ')).toBeNull();
    expect(normalizeBookmarkName('')).toBeNull();
  });
});

describe('הוספת סימנייה', () => {
  it('שולחת את השם ואת הבחירה כמו שהיא', async () => {
    const engine = fakeEngine();

    expect(await insertBookmark(engine.host, 'פרק_ראשון')).toEqual({ ok: true });

    expect(engine.inputs('bookmarks.insert')).toEqual([
      {
        name: 'פרק_ראשון',
        at: { kind: 'text', segments: [{ blockId: 'block-1', range: { start: 3, end: 3 } }] },
      },
    ]);
  });

  it('שם עברי מגיע למנוע בלי עיוות', async () => {
    const engine = fakeEngine();
    await insertBookmark(engine.host, 'הקדמת_המחבר');
    expect((engine.inputs('bookmarks.insert')[0] as { name: string }).name).toBe('הקדמת_המחבר');
  });

  it('שם פסול נעצר **לפני** המנוע, וההודעה מסבירה מה מותר', async () => {
    const engine = fakeEngine();

    const outcome = await insertBookmark(engine.host, 'שם עם רווח');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('invalid-name');
    expect(outcome.ok === false && outcome.message).toContain(BOOKMARK_NAME_HINT);
    // זה העיקר: המסמך לא נגע.
    expect(engine.ops()).not.toContain('bookmarks.insert');
  });

  it('בלי בחירה שאפשר לפעול עליה — כשל מוסבר, בלי מוטציה', async () => {
    const engine = fakeEngine({ blockId: null });

    const outcome = await insertBookmark(engine.host, 'הקדמה');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('no-selection');
    expect(engine.ops()).not.toContain('bookmarks.insert');
  });

  it('קבלה שנכשלה מתורגמת לעברית — ונוסח המנוע עצמו אינו מגיע למשתמש', async () => {
    const engine = fakeEngine({
      failures: {
        'bookmarks.insert': {
          code: 'INVALID_INPUT',
          message: 'bookmarks.insert: bookmark "הקדמה" already exists.',
        },
      },
    });

    const outcome = await insertBookmark(engine.host, 'הקדמה');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('INVALID_INPUT');
    expect(outcome.ok === false && outcome.message).toContain('הוספת הסימנייה נכשלה');
    // זו ההתנהגות בפועל, וכאן היא מתועדת ולא מתוארת יפה: `receiptFailureText`
    // מעדיף את התרגום הגנרי של הקוד על פני `failure.message`, ולכן המשתמש
    // שומע „ערך שאינו חוקי” ולא „כבר קיים”. הוא מודול משותף לכל הפקדים,
    // ושינוי שלו הוא גל בפני עצמו — ולכן „שם כפול” נחסם בדיאלוג לפני
    // הקריאה למנוע (ראו tests/component/bookmark-dialog.test.ts), והמסלול
    // הזה נשאר רק לכשל שלא צפינו.
    expect(outcome.ok === false && outcome.message).not.toContain('already exists');
  });

  it('`NO_OP` היא הצלחה — קבלה בלי שינוי אינה כשל', async () => {
    // המנוע מחזיר `NO_OP` על פעולה שלא היה בה מה לעשות. הצגתה כשגיאה אדומה
    // הייתה מאשימה את המשתמש במשהו שהצליח. אותה הכרעה כמו ב-cross-refs.ts.
    const engine = fakeEngine({ failures: { 'bookmarks.insert': { code: 'NO_OP' } } });

    expect(await insertBookmark(engine.host, 'הקדמה')).toEqual({ ok: true });
  });

  it('חריגה שנזרקה אינה יוצאת החוצה', async () => {
    const engine = fakeEngine({ throws: ['bookmarks.insert'] });

    const outcome = await insertBookmark(engine.host, 'הקדמה');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('threw');
  });

  it('פעולה שאינה קיימת בגרסה — הנוסח שהתכנית קובעת', async () => {
    const engine = fakeEngine({ missing: ['bookmarks.insert'] });

    expect(await insertBookmark(engine.host, 'הקדמה')).toEqual({
      ok: false,
      message: 'הוספת הסימנייה נכשלה: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });

  it('בלי Document API כלל — כשל מוסבר ולא זריקה', async () => {
    expect(await insertBookmark(null, 'הקדמה')).toEqual({
      ok: false,
      message: 'הוספת הסימנייה נכשלה: המסמך עדיין נטען',
      reason: 'document-api-unavailable',
    });
  });
});

describe('שינוי שם ומחיקה', () => {
  it('שינוי שם שולח כתובת ישות ושם חדש', async () => {
    const engine = fakeEngine();

    expect(await renameBookmark(engine.host, 'הקדמה', 'הקדמת_המחבר')).toEqual({ ok: true });

    expect(engine.inputs('bookmarks.rename')).toEqual([
      {
        target: { kind: 'entity', entityType: 'bookmark', name: 'הקדמה' },
        newName: 'הקדמת_המחבר',
      },
    ]);
  });

  it('השם החדש עובר את אותה ולידציה — זה המסלול שמכניס שם פסול למסמך תקין', async () => {
    const engine = fakeEngine();

    const outcome = await renameBookmark(engine.host, 'הקדמה', 'שם עם רווח');

    expect(outcome.ok === false && outcome.reason).toBe('invalid-name');
    expect(engine.ops()).not.toContain('bookmarks.rename');
  });

  it('מחיקה שולחת כתובת ישות', async () => {
    const engine = fakeEngine();

    expect(await removeBookmark(engine.host, 'הקדמה')).toEqual({ ok: true });

    expect(engine.inputs('bookmarks.remove')).toEqual([
      { target: { kind: 'entity', entityType: 'bookmark', name: 'הקדמה' } },
    ]);
  });

  it('סימנייה שאינה קיימת — הכשל של המנוע מגיע למשתמש', async () => {
    const engine = fakeEngine({
      failures: { 'bookmarks.remove': { code: 'TARGET_NOT_FOUND' } },
    });

    const outcome = await removeBookmark(engine.host, 'אין-כזו');

    expect(outcome.ok === false && outcome.reason).toBe('TARGET_NOT_FOUND');
    expect(outcome.ok === false && outcome.message).toContain('מחיקת הסימנייה נכשלה');
  });

  it('חריגה ופעולה חסרה — בשתיהן `CommandOutcome`', async () => {
    expect((await removeBookmark(fakeEngine({ throws: ['bookmarks.remove'] }).host, 'א')).ok).toBe(
      false,
    );
    expect(await renameBookmark(fakeEngine({ missing: ['bookmarks.rename'] }).host, 'א', 'ב')).toEqual(
      {
        ok: false,
        message: 'שינוי שם הסימנייה נכשל: אינו זמין בגרסה זו',
        reason: 'command-unsupported',
      },
    );
  });
});

describe('קריאת הסימניות', () => {
  it('שואבת עמודים עד `total` ולא עוצרת בעמוד הראשון', async () => {
    const bookmarks = Array.from({ length: 450 }, (_, index) => ({ name: `סימן_${index}` }));
    const engine = fakeEngine({ bookmarks, total: 450 });

    const state = await readBookmarks(engine.host);

    expect(state.names).toHaveLength(450);
    expect(state.names[0]).toBe('סימן_0');
    expect(state.names[449]).toBe('סימן_449');
    // שלושה עמודים של 200: 200, 200, 50.
    expect(engine.inputs('bookmarks.list')).toEqual([
      { limit: 200, offset: 0 },
      { limit: 200, offset: 200 },
      { limit: 200, offset: 400 },
    ]);
  });

  it('בלי `total` נעצרת אחרי עמוד אחד ואינה נכנסת ללולאה', async () => {
    // גרסה שאינה חושפת `total` אינה נותנת שום סימן מתי הרשימה נגמרה, וכל
    // המשך שאיבה הוא ניחוש. עמוד אחד — כמו `rebuildAllFields`, ומאותו טעם.
    const engine = fakeEngine({ bookmarks: [{ name: 'א' }, { name: 'ב' }] });

    expect((await readBookmarks(engine.host)).names).toEqual(['א', 'ב']);
    expect(engine.inputs('bookmarks.list')).toHaveLength(1);
  });

  it('רשומה בלי שם מדולגת ואינה הופכת ל-undefined ברשימה', async () => {
    const engine = fakeEngine({ bookmarks: [{ name: 'א' }, {}, { name: '' }], total: 3 });

    expect((await readBookmarks(engine.host)).names).toEqual(['א']);
  });

  it('חריגה מחזירה את מה שנאסף, ולא זורקת', async () => {
    const engine = fakeEngine({ throws: ['bookmarks.list'] });

    expect(await readBookmarks(engine.host)).toEqual(emptyBookmarksState());
  });

  it('גרסה בלי `bookmarks.list` — רשימה ריקה, בלי כשל', async () => {
    expect(await readBookmarks(fakeEngine({ missing: ['bookmarks.list'] }).host)).toEqual({
      names: [],
    });
    expect(await readBookmarks(null)).toEqual({ names: [] });
  });
});
