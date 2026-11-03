/**
 * ציטוטים וביבליוגרפיה — ניהול מקורות, הוספת ציטוט, ובניית הרשימה.
 *
 * שש הטענות שנמדדות כאן, וכולן מקבעות ממצא שנמדד בדפדפן ולא הנחה (ההנמקה
 * המלאה, כולל הפירוק של ה-docx המיוצא, ב-engine/citations.ts):
 *
 * 1. **ציטוט נשלח עם מקור אחד בדיוק.** `citations.insert` מקבל מערך, ובשניים
 *    הוא כותב `CITATION src-a;src-b` — ותחביר ריבוי המקורות של Word הוא
 *    המתג `\m`. הבדיקה מקבעת שהמערך שנשלח הוא תמיד באורך אחד.
 * 2. **היעד מכווץ לסוף הבחירה.** המנוע דוחה טווח ב-`INVALID_TARGET`, ו-Word
 *    מכניס את הציטוט אחרי הטקסט המסומן.
 * 3. **הוולידציה יושבת אצלנו, כי המנוע בולע.** כותרת של רווחים בלבד, סוג
 *    שאינו בחוזה, ושדות ריקים — כולם חוזרים מהמנוע `success: true` וכותבים
 *    לקובץ מקור פגום. שלושתם נדחים או מנוכים כאן, לפני שנוגעים במסמך.
 * 4. **מקור מצוטט אינו נמחק.** נמדד שהמנוע מוחק בהצלחה ומשאיר שדה
 *    `CITATION` מצביע לתג שאינו קיים — כלומר המסמך השבור שגל 5 נדחה בגללו.
 * 5. **הביבליוגרפיה נמצאת דרך `fields.list`**, מפני ש-`citations.bibliography`
 *    חסר `list` ו-`blocks.list` מציג אותה כפסקה רגילה. הבדיקה דורשת שלא
 *    נשלחת שום קריאה ל-`blocks.*`.
 * 6. **המודול לעולם אינו זורק.** חריגה, קבלה שנכשלה, ופעולה שאינה קיימת
 *    בגרסת המנוע — שלושתן `CommandOutcome`.
 */
import { describe, expect, it } from 'vitest';
import {
  CITATION_TITLE_HINT,
  CITATION_TYPE_HINT,
  addCitationSource,
  emptyCitationSourceDraft,
  emptyCitationsState,
  formatCitationAuthors,
  insertBibliography,
  insertCitation,
  listCitationSources,
  normalizeCitationTitle,
  parseCitationAuthors,
  readCitationsState,
  rebuildBibliography,
  removeBibliography,
  removeCitationSource,
  updateCitationSource,
  usesJournalFields,
  type CitationSourceDraft,
  type CitationsHost,
} from '../../src/engine/citations';

interface Call {
  op: string;
  input?: unknown;
}

interface FakeSource {
  sourceId: string;
  type?: string;
  title?: string;
  authors?: { last: string; first?: string }[];
  year?: string;
  /** מקור בלי `sourceId` — מה שהמנוע עלול להחזיר, ומה שאסור להציג. */
  idless?: boolean;
  /** מקור בלי `sourceId` וגם בלי `id`. אין לו כתובת, ולכן אין מה להציג. */
  anonymous?: boolean;
}

interface FakeOptions {
  sources?: readonly FakeSource[];
  /** `total` שהמנוע מדווח, כשהוא שונה מאורך הרשימה — כלומר שאיבת עמודים. */
  sourcesTotal?: number;
  /** לכל ציטוט במסמך, המקור שהוא מפנה אליו. */
  cited?: readonly string[];
  /** מזהי הבלוקים של הביבליוגרפיות, כפי ש-`fields.list` מדווח אותם. */
  bibliographies?: readonly string[];
  /** מה `selection.current` מדווח. `null` = אין בחירה. */
  caret?: { blockId: string; start: number; end: number } | null;
  /** בחירה רב-מקטעית, כמו שהמנוע מדווח על בחירה שחוצה פסקאות. גובר על `caret`. */
  segments?: readonly { blockId: string; range: { start: number; end: number } }[];
  failures?: Record<string, { code: string; message?: string }>;
  throws?: readonly string[];
  missing?: readonly string[];
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
      if (throwing.has(op)) throw new Error(`${op} התפוצץ`);
      return impl(input);
    };
  }

  function receipt(op: string): unknown {
    const failure = failures[op];
    return failure ? { success: false, failure } : { success: true };
  }

  const sources = options.sources ?? [];
  const cited = options.cited ?? [];
  const bibliographies = options.bibliographies ?? [];
  const caret = options.caret === undefined ? { blockId: 'block-1', start: 4, end: 9 } : options.caret;

  /** עמוד תחת `limit`/`offset`, כמו `DiscoveryOutput` האמיתי. */
  function page<T>(items: readonly T[], total: number, input: unknown) {
    const query = (input ?? {}) as { limit?: number; offset?: number };
    const offset = query.offset ?? 0;
    const end = query.limit === undefined ? undefined : offset + query.limit;
    return { items: items.slice(offset, end), total };
  }

  const doc = {
    fields: {
      list: route('fields.list', (input) =>
        page(
          bibliographies.map((nodeId) => ({
            address: { kind: 'field', blockId: nodeId },
            instruction: 'BIBLIOGRAPHY',
            fieldType: 'BIBLIOGRAPHY',
          })),
          bibliographies.length,
          input,
        ),
      ),
    },
    blocks: {
      list: route('blocks.list', () => ({ blocks: [], total: 0 })),
      deleteRange: route('blocks.deleteRange', () => receipt('blocks.deleteRange')),
    },
    selection: {
      current: route('selection.current', () => {
        const segments =
          options.segments ??
          (caret
            ? [{ blockId: caret.blockId, range: { start: caret.start, end: caret.end } }]
            : undefined);
        return {
          empty: segments === undefined,
          target: segments === undefined ? null : { kind: 'text', segments },
        };
      }),
    },
    citations: {
      list: route('citations.list', (input) =>
        page(
          cited.map((sourceId) => ({ sourceIds: [sourceId] })),
          cited.length,
          input,
        ),
      ),
      insert: route('citations.insert', () => receipt('citations.insert')),
      sources: {
        list: route('citations.sources.list', (input) =>
          page(
            sources.map((source) => ({
              id: source.anonymous ? undefined : source.sourceId,
              sourceId: source.idless || source.anonymous ? undefined : source.sourceId,
              tag: source.title,
              type: source.type ?? 'book',
              fields: {
                title: source.title,
                authors: source.authors,
                year: source.year,
              },
            })),
            options.sourcesTotal ?? sources.length,
            input,
          ),
        ),
        insert: route('citations.sources.insert', () => receipt('citations.sources.insert')),
        update: route('citations.sources.update', () => receipt('citations.sources.update')),
        remove: route('citations.sources.remove', () => receipt('citations.sources.remove')),
      },
      bibliography: {
        insert: route('citations.bibliography.insert', () =>
          receipt('citations.bibliography.insert'),
        ),
        rebuild: route('citations.bibliography.rebuild', () =>
          receipt('citations.bibliography.rebuild'),
        ),
        remove: route('citations.bibliography.remove', () =>
          receipt('citations.bibliography.remove'),
        ),
      },
    },
  };

  return {
    host: { activeEditor: { doc } } as unknown as CitationsHost,
    calls,
    ops: () => calls.map((call) => call.op),
    inputs: (op: string) => calls.filter((call) => call.op === op).map((call) => call.input),
  };
}

/** טופס מלא, כדי שכל בדיקה תשנה רק את מה שהיא מודדת. */
function draftOf(patch: Partial<CitationSourceDraft> = {}): CitationSourceDraft {
  return {
    ...emptyCitationSourceDraft(),
    title: 'שולחן ערוך',
    authors: 'קארו, יוסף',
    year: 'שכ״ה',
    city: 'ונציה',
    publisher: 'ג׳יוסטיניאן',
    ...patch,
  };
}

/* ------------------------------------------------------------------ */

describe('ולידציה של המקור — כאן ולא במנוע', () => {
  it('כותרת של רווחים בלבד נדחית, ולא מגיעה למסמך', async () => {
    // נמדד: `fields: { title: '   ' }` חוזר `success: true` וכותב
    // `<b:Title>   </b:Title>` — מקור שאי אפשר לזהות ואי אפשר למחוק.
    expect(normalizeCitationTitle('   ')).toBeNull();
    expect(normalizeCitationTitle('')).toBeNull();
    expect(normalizeCitationTitle('  רמב״ם  ')).toBe('רמב״ם');

    const engine = fakeEngine();
    const outcome = await addCitationSource(engine.host, draftOf({ title: '   ' }));

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.message).toContain(CITATION_TITLE_HINT);
    expect(engine.ops()).not.toContain('citations.sources.insert');
  });

  it('סוג שאינו ברשימה נדחה — המנוע בולע אותו וכותב אותו לקובץ', async () => {
    const engine = fakeEngine();
    const outcome = await addCitationSource(
      engine.host,
      draftOf({ type: 'zigzag' as never }),
    );

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.message).toContain(CITATION_TYPE_HINT);
    expect(engine.ops()).not.toContain('citations.sources.insert');
  });

  it('שדה ריק אינו נשלח בכלל, ולא כמחרוזת ריקה', async () => {
    const engine = fakeEngine();
    await addCitationSource(
      engine.host,
      draftOf({ year: '   ', city: '', publisher: '' }),
    );

    const [input] = engine.inputs('citations.sources.insert') as {
      fields: Record<string, unknown>;
    }[];
    expect(Object.keys(input.fields).sort()).toEqual(['authors', 'title']);
  });

  it('שדות כתב עת נשלחים רק לסוג שיש להם בו מובן', async () => {
    expect(usesJournalFields('book')).toBe(false);
    expect(usesJournalFields('journalArticle')).toBe(true);

    const asBook = fakeEngine();
    await addCitationSource(asBook.host, draftOf({ journalName: 'סיני', volume: 'ה' }));
    const [bookInput] = asBook.inputs('citations.sources.insert') as {
      fields: Record<string, unknown>;
    }[];
    expect(bookInput.fields.journalName).toBeUndefined();
    expect(bookInput.fields.volume).toBeUndefined();

    const asArticle = fakeEngine();
    await addCitationSource(
      asArticle.host,
      draftOf({ type: 'journalArticle', journalName: 'סיני', volume: 'ה' }),
    );
    const [articleInput] = asArticle.inputs('citations.sources.insert') as {
      fields: Record<string, unknown>;
    }[];
    expect(articleInput.fields.journalName).toBe('סיני');
  });
});

describe('שמות המחברים', () => {
  it('שם בלי פסיק נשמר כשם אחד — זה המצב הרגיל בספר תורני', () => {
    expect(parseCitationAuthors('רמב״ם')).toEqual([{ last: 'רמב״ם' }]);
    expect(parseCitationAuthors('חזון איש\nשולחן ערוך')).toEqual([
      { last: 'חזון איש' },
      { last: 'שולחן ערוך' },
    ]);
  });

  it('שם עם פסיק מתפצל ל„משפחה, פרטי”', () => {
    expect(parseCitationAuthors('כהן, יוסף')).toEqual([{ last: 'כהן', first: 'יוסף' }]);
    // פסיק בלי שם פרטי אינו מייצר `first` ריק.
    expect(parseCitationAuthors('כהן,')).toEqual([{ last: 'כהן' }]);
  });

  it('שורה ריקה ושורה בלי `last` מדולגות — נמדד שהמנוע מתפוצץ עליהן', () => {
    // מחבר בלי `last` הפיל את המנוע ב-`TypeError` גולמי ולא בקבלה.
    expect(parseCitationAuthors('\n  \n, יוסף\nלוי')).toEqual([{ last: 'לוי' }]);
  });

  it('העיצוב הוא ההפך המדויק של הפירוק', () => {
    expect(formatCitationAuthors([{ last: 'כהן', first: 'יוסף' }, { last: 'רמב״ם' }])).toBe(
      'כהן, יוסף\nרמב״ם',
    );
    expect(formatCitationAuthors(undefined)).toBe('');
  });
});

describe('הוספת ציטוט', () => {
  it('שולחת מקור אחד בדיוק, ביעד מכווץ לסוף הבחירה', async () => {
    // שתי ההכרעות שנמדדו: `CITATION a;b` אינו תחביר של Word, וטווח נדחה
    // ב-`INVALID_TARGET`.
    const engine = fakeEngine({
      sources: [{ sourceId: 'src-1', title: 'ספר' }],
      caret: { blockId: 'block-7', start: 4, end: 9 },
    });

    const outcome = await insertCitation(engine.host, 'src-1');

    expect(outcome).toEqual({ ok: true });
    const [input] = engine.inputs('citations.insert') as {
      at: { segments: { blockId: string; range: { start: number; end: number } }[] };
      sourceIds: string[];
    }[];
    expect(input.sourceIds).toEqual(['src-1']);
    expect(input.at.segments).toEqual([{ blockId: 'block-7', range: { start: 9, end: 9 } }]);
  });

  it('בחירה שחוצה פסקאות מתכווצת לסוף המקטע האחרון', async () => {
    // „סוף הבחירה” הוא סוף המקטע האחרון. עצירה על הראשון הייתה שותלת את
    // הציטוט באמצע הטקסט המסומן.
    const engine = fakeEngine({
      segments: [
        { blockId: '', range: { start: 0, end: 3 } },
        { blockId: 'block-1', range: { start: 4, end: 12 } },
        { blockId: 'block-2', range: { start: 0, end: 5 } },
        { blockId: 'block-3', range: {} as { start: number; end: number } },
      ],
    });

    expect(await insertCitation(engine.host, 'src-1')).toEqual({ ok: true });

    const [input] = engine.inputs('citations.insert') as {
      at: { segments: unknown[] };
    }[];
    expect(input.at.segments).toEqual([{ blockId: 'block-2', range: { start: 5, end: 5 } }]);
  });

  it('בלי סמן במסמך — כשל מנומק, ובלי לגעת במנוע', async () => {
    const engine = fakeEngine({ caret: null });

    const outcome = await insertCitation(engine.host, 'src-1');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('no-caret');
    expect(engine.ops()).not.toContain('citations.insert');
  });

  it('בלי מקור נבחר — כשל מנומק', async () => {
    const engine = fakeEngine();
    const outcome = await insertCitation(engine.host, '   ');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('no-source');
    expect(engine.ops()).not.toContain('citations.insert');
  });

  it('קבלה שנכשלה מתורגמת להודעה בעברית', async () => {
    const engine = fakeEngine({
      failures: { 'citations.insert': { code: 'TARGET_NOT_FOUND', message: 'no such source' } },
    });

    const outcome = await insertCitation(engine.host, 'src-1');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('TARGET_NOT_FOUND');
    expect(outcome.ok === false && outcome.message).toContain('הוספת הציטוט נכשלה');
  });

  it('חריגה נתפסת ואינה יוצאת החוצה', async () => {
    const engine = fakeEngine({ throws: ['citations.insert'] });

    const outcome = await insertCitation(engine.host, 'src-1');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('threw');
  });

  it('פעולה שאינה בגרסת המנוע מחזירה את נוסח §12', async () => {
    const engine = fakeEngine({ missing: ['citations.insert'] });

    const outcome = await insertCitation(engine.host, 'src-1');

    expect(outcome).toEqual({
      ok: false,
      message: 'הוספת הציטוט נכשלה: אינו זמין בגרסה זו',
      reason: 'command-unsupported',
    });
  });
});

describe('מחיקת מקור', () => {
  it('מסרבת כשיש למקור ציטוט במסמך, ואינה נוגעת במנוע', async () => {
    // נמדד: המנוע מוחק בהצלחה ומשאיר שדה `CITATION` מצביע לתג שאינו קיים.
    const engine = fakeEngine({
      sources: [{ sourceId: 'src-1', title: 'ספר' }],
      cited: ['src-1', 'src-1'],
    });

    const outcome = await removeCitationSource(engine.host, 'src-1');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('source-in-use');
    expect(outcome.ok === false && outcome.message).toContain('2 ציטוטים');
    expect(engine.ops()).not.toContain('citations.sources.remove');
  });

  it('מוחקת מקור שאין לו ציטוט', async () => {
    const engine = fakeEngine({
      sources: [{ sourceId: 'src-1', title: 'ספר' }],
      cited: ['src-2'],
    });

    const outcome = await removeCitationSource(engine.host, 'src-1');

    expect(outcome).toEqual({ ok: true });
    expect(engine.inputs('citations.sources.remove')).toEqual([
      { target: { kind: 'entity', entityType: 'citationSource', sourceId: 'src-1' } },
    ]);
  });

  it('כשל של ספירת הציטוטים מונע את המחיקה — „לא בדקנו” אינו „אין”', async () => {
    const engine = fakeEngine({ missing: ['citations.list'] });

    const outcome = await removeCitationSource(engine.host, 'src-1');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('citations-unreadable');
    expect(engine.ops()).not.toContain('citations.sources.remove');
  });
});

describe('קריאת המקורות', () => {
  it('נופלת חזרה ל-`id`, ומדלגת על מקור בלי מזהה כלל', async () => {
    const engine = fakeEngine({
      sources: [
        { sourceId: 'src-1', title: 'ספר א', year: 'תש״ף', authors: [{ last: 'כהן' }] },
        { sourceId: 'src-2', title: 'ספר ב', idless: true },
        { sourceId: 'src-3', title: 'ספר ג', anonymous: true },
        { sourceId: '', title: 'ספר ד' },
      ],
      cited: ['src-1'],
    });

    const sources = await listCitationSources(engine.host);

    // המקור השני חסר `sourceId` ובכל זאת מוצג: `DiscoveryItem.id` הוא חלק
    // מהחוזה הציבורי, ובמנוע האמיתי הוא זהה ל-`sourceId`. נפילה חזרה אליו
    // עדיפה על הסתרת מקור שהמשתמש רואה במסמך. השלישי חסר את שניהם, ולכן
    // נשמטים — שלישי בלי שניהם, רביעי עם מזהה ריק: שורה שלחיצה עליה
    // שולחת `undefined` או `''` ל-`sources.remove` גרועה משורה שאינה שם.
    expect(sources.map((source) => source.id)).toEqual(['src-1', 'src-2']);
    expect(sources[0].label).toBe('ספר א (כהן, תש״ף)');
    expect(sources[0].citedCount).toBe(1);
    expect(sources[1].citedCount).toBe(0);
  });

  it('שואבת את כל העמודים, ולא רק את הראשון', async () => {
    const sources = Array.from({ length: 250 }, (_, index) => ({
      sourceId: `src-${index}`,
      title: `ספר ${index}`,
    }));
    const engine = fakeEngine({ sources });

    expect(await listCitationSources(engine.host)).toHaveLength(250);
    // שני עמודים: 200 ואז 50. עמוד אחד היה מציג 200 מקורות ומסתיר 50,
    // ורשימת מקורות ארוכה היא בדיוק המסמך התורני שהתוסף נבנה בשבילו.
    expect(engine.inputs('citations.sources.list')).toEqual([
      { limit: 200, offset: 0 },
      { limit: 200, offset: 200 },
    ]);
  });

  it('`total` מנופח אינו מייצר לולאה אינסופית — עמוד ריק עוצר', async () => {
    const sources = Array.from({ length: 250 }, (_, index) => ({
      sourceId: `src-${index}`,
      title: `ספר ${index}`,
    }));
    const engine = fakeEngine({ sources, sourcesTotal: 500 });

    expect(await listCitationSources(engine.host)).toHaveLength(250);
    expect(engine.inputs('citations.sources.list')).toEqual([
      { limit: 200, offset: 0 },
      { limit: 200, offset: 200 },
      { limit: 200, offset: 250 },
    ]);
  });

  it('סוג שאינו מוכר יורד ל„אחר” ואינו נעלם', async () => {
    const engine = fakeEngine({ sources: [{ sourceId: 'src-1', title: 'ס', type: 'zigzag' }] });

    const [source] = await listCitationSources(engine.host);

    expect(source.draft.type).toBe('misc');
  });

  it('מסמך בלי כלום מדווח אפס בשלושת המונים', async () => {
    const engine = fakeEngine();

    expect(await readCitationsState(engine.host)).toEqual(emptyCitationsState());
  });

  it('סופרת ביבליוגרפיות דרך `fields.list` ולא דרך `blocks.*`', async () => {
    // אין ל-`citations.bibliography` פעולת `list`, ו-`blocks.list` מציג את
    // הביבליוגרפיה כפסקה רגילה (נמדד).
    const engine = fakeEngine({ bibliographies: ['BIB-1'], cited: ['src-1'] });

    const state = await readCitationsState(engine.host);

    expect(state.bibliographyCount).toBe(1);
    expect(state.citationCount).toBe(1);
    expect(engine.ops()).not.toContain('blocks.list');
  });

  it('כשל של הקריאה מחזיר „אין” ולא ספירה חלקית', async () => {
    const engine = fakeEngine({ throws: ['fields.list', 'citations.list'] });

    const state = await readCitationsState(engine.host);

    expect(state.bibliographyCount).toBe(0);
    expect(state.citationCount).toBe(0);
  });
});

describe('עריכת מקור', () => {
  it('שדה שרוקן נשלח כמחרוזת ריקה, כדי שהמחיקה תגיע למסמך', async () => {
    // נמדד: `patch` הוא `Partial` אמיתי — patch בלי `year` השאיר את הערך
    // הישן במסמך, ורק `year: ''` מחק אותו. השמטה כאן הייתה „הצלחה מדומה”:
    // `{ok:true}` בלי הודעה, והשנה חוזרת ברענון.
    const engine = fakeEngine();

    const outcome = await updateCitationSource(
      engine.host,
      'src-1',
      draftOf({ year: '', city: '   ', authors: '' }),
    );

    expect(outcome).toEqual({ ok: true });
    const [input] = engine.inputs('citations.sources.update') as {
      target: { sourceId: string };
      patch: Record<string, unknown>;
    }[];
    expect(input.target.sourceId).toBe('src-1');
    expect(input.patch.year).toBe('');
    expect(input.patch.city).toBe('');
    expect(input.patch.authors).toEqual([]);
    expect(input.patch.title).toBe('שולחן ערוך');
  });

  it('שדה שאינו רלוונטי לסוג מושמט ואינו נמחק', async () => {
    // ההשמטה נשארת רק כאן: לספר אין „כרך” להתחיל איתו, ומחרוזת ריקה הייתה
    // כותבת `<b:Volume></b:Volume>` לקובץ.
    const engine = fakeEngine();

    await updateCitationSource(engine.host, 'src-1', draftOf({ journalName: '', volume: '' }));

    const [input] = engine.inputs('citations.sources.update') as {
      patch: Record<string, unknown>;
    }[];
    expect(input.patch.journalName).toBeUndefined();
    expect(input.patch.volume).toBeUndefined();
    expect(input.patch.pages).toBeUndefined();
  });

  it('בהוספה — ההפך: שדה שרוקן מושמט ואינו נשלח ריק', async () => {
    // מקור חדש אין בו מה למחוק, ו-`<b:Year></b:Year>` ריק הוא רעש בקובץ.
    const engine = fakeEngine();

    await addCitationSource(engine.host, draftOf({ year: '', city: '' }));

    const [input] = engine.inputs('citations.sources.insert') as {
      fields: Record<string, unknown>;
    }[];
    expect(input.fields.year).toBeUndefined();
    expect(input.fields.city).toBeUndefined();
  });

  it('אינה שולחת `type` — אין בחוזה מסלול שמשנה אותו', async () => {
    const engine = fakeEngine();
    await updateCitationSource(engine.host, 'src-1', draftOf());

    const [input] = engine.inputs('citations.sources.update') as {
      patch: Record<string, unknown>;
    }[];
    expect(input.patch.type).toBeUndefined();
  });
});

describe('ביבליוגרפיה', () => {
  it('נוספת בסוף המסמך, ובלי סגנון — `\\sdStyle` אינו מתג של Word', async () => {
    const engine = fakeEngine();

    const outcome = await insertBibliography(engine.host);

    expect(outcome).toEqual({ ok: true });
    expect(engine.inputs('citations.bibliography.insert')).toEqual([
      { at: { kind: 'documentEnd' } },
    ]);
  });

  it('העדכון רץ על כל הביבליוגרפיות שבמסמך', async () => {
    const engine = fakeEngine({ bibliographies: ['BIB-1', 'BIB-2'] });

    const outcome = await rebuildBibliography(engine.host);

    expect(outcome).toEqual({ ok: true });
    expect(engine.inputs('citations.bibliography.rebuild')).toEqual([
      { target: { kind: 'block', nodeType: 'bibliography', nodeId: 'BIB-1' } },
      { target: { kind: 'block', nodeType: 'bibliography', nodeId: 'BIB-2' } },
    ]);
  });

  it('כתובת חוזרת מדווחת כעדכון שלא הושלם, ולא כ„בוצע”', async () => {
    // גל 4 מדד מנוע שנותן לשני עצמים את אותו `nodeId`. לולאה תמימה הייתה
    // בונה את הראשונה פעמיים ומדווחת „בוצע” על שנייה מיושנת.
    const engine = fakeEngine({ bibliographies: ['BIB-1', 'BIB-1'] });

    const outcome = await rebuildBibliography(engine.host);

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('ambiguous-bibliography');
    expect(engine.inputs('citations.bibliography.rebuild')).toHaveLength(1);
  });

  it('עדכון במסמך בלי ביבליוגרפיה מדווח למה, ואינו שותק', async () => {
    const engine = fakeEngine();

    const outcome = await rebuildBibliography(engine.host);

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('no-bibliography');
  });

  it('ההסרה היא צעד אחד, בלי ניקוי שיירים', async () => {
    // הביבליוגרפיה היא בלוק יחיד, ונמדד שאחרי `remove` לא נשארת ממנה פסקה.
    // ניקוי שאין לו מה לנקות היה רק דרך למחוק פסקה של המשתמש.
    const engine = fakeEngine({ bibliographies: ['BIB-1'] });

    const outcome = await removeBibliography(engine.host);

    expect(outcome).toEqual({ ok: true });
    expect(engine.ops()).not.toContain('blocks.deleteRange');
  });

  it('הסרה במסמך בלי ביבליוגרפיה מדווחת למה, ואינה שותקת', async () => {
    const engine = fakeEngine();

    const outcome = await removeBibliography(engine.host);

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('no-bibliography');
    expect(engine.ops()).not.toContain('citations.bibliography.remove');
  });

  it('ההסרה מסרבת כשיש יותר מאחת', async () => {
    const engine = fakeEngine({ bibliographies: ['BIB-1', 'BIB-2'] });

    const outcome = await removeBibliography(engine.host);

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.reason).toBe('ambiguous-bibliography');
    expect(engine.ops()).not.toContain('citations.bibliography.remove');
  });

  it('קבלה שנכשלה, חריגה ופעולה חסרה — שלושתן `CommandOutcome`', async () => {
    const failing = fakeEngine({
      bibliographies: ['BIB-1'],
      failures: { 'citations.bibliography.remove': { code: 'INVALID_TARGET' } },
    });
    const failed = await removeBibliography(failing.host);
    expect(failed.ok === false && failed.reason).toBe('INVALID_TARGET');

    const throwing = fakeEngine({
      bibliographies: ['BIB-1'],
      throws: ['citations.bibliography.remove'],
    });
    const threw = await removeBibliography(throwing.host);
    expect(threw.ok === false && threw.reason).toBe('threw');

    const absent = fakeEngine({ missing: ['citations.bibliography.remove'] });
    const missing = await removeBibliography(absent.host);
    expect(missing.ok === false && missing.reason).toBe('command-unsupported');
  });
});

describe('מסמך שעדיין נטען', () => {
  it('כל פעולה מחזירה כשל מנומק ואינה זורקת', async () => {
    const outcomes = await Promise.all([
      addCitationSource(null, draftOf()),
      updateCitationSource(null, 'src-1', draftOf()),
      removeCitationSource(null, 'src-1'),
      insertCitation(null, 'src-1'),
      insertBibliography(null),
      rebuildBibliography(null),
      removeBibliography(null),
    ]);

    for (const outcome of outcomes) {
      expect(outcome.ok).toBe(false);
      expect(outcome.ok === false && outcome.reason).toBe('document-api-unavailable');
    }

    expect(await listCitationSources(null)).toEqual([]);
    expect(await readCitationsState(null)).toEqual(emptyCitationsState());
  });
});
