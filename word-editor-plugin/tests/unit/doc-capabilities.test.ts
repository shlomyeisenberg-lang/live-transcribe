/**
 * בדיקת היסוד שכל הפקדים החדשים נשענים עליו.
 *
 * מה שנבדק כאן הוא ההתנהגות שהמודול נכתב בשבילה: **נכשל סגור**. כל תשובה
 * חלקית, חסרה, זרוקה או שאינה אובייקט חייבת להסתיים ב-`false` עם סיבה — כי
 * `true` בטעות אחת כאן מחזיר בדיוק את התקלה שהתוסף סבל ממנה: כפתור שנראה
 * עובד ואינו עובד.
 */
import { describe, expect, it, vi } from 'vitest';
import {
  DOC_CAPABILITY_QUESTIONS,
  readDocCapabilities,
  type CapabilitiesHost,
  type DocCapabilityQuestion,
} from '../../src/engine/doc-capabilities';

/** מפה מלאה שכל מה שהממשק שואל עליו זמין בה. */
function fullCapabilities() {
  return {
    global: {
      trackChanges: { enabled: true },
      comments: { enabled: true },
      lists: { enabled: true },
      dryRun: { enabled: true },
      history: { enabled: true },
    },
    operations: {
      'sections.setPageMargins': { available: true },
      'sections.setPageSetup': { available: true },
      'sections.setColumns': { available: true },
      'sections.setSectionDirection': { available: true },
      'sections.setBreakType': { available: true },
      'sections.setPageBorders': { available: true },
      'sections.clearPageBorders': { available: true },
      'sections.setLineNumbering': { available: true },
      'sections.setVerticalAlign': { available: true },
      'sections.setPageNumbering': { available: true },
      'sections.setHeaderFooterMargins': { available: true },
      'format.vertAlign': { available: true },
      'footnotes.insert': { available: true },
      'footnotes.list': { available: true },
      'footnotes.get': { available: true },
      'footnotes.update': { available: true },
      'footnotes.remove': { available: true },
      'format.paragraph.setFlowOptions': { available: true },
      'format.apply': { available: true },
      'styles.apply': { available: true },
      'lists.setLevelNumberStyle': { available: true },
      'lists.restartAt': { available: true },
      'lists.continuePrevious': { available: true },
      'lists.convertToText': { available: true },
      'format.paragraph.setIndentation': { available: true },
      'format.paragraph.clearIndentation': { available: true },
      'format.paragraph.setSpacing': { available: true },
      'format.paragraph.clearSpacing': { available: true },
      'format.paragraph.setKeepOptions': { available: true },
      'format.paragraph.setTabStop': { available: true },
      'format.paragraph.clearTabStop': { available: true },
      'format.paragraph.clearAllTabStops': { available: true },
      insert: { available: true },
      'create.image': { available: true },
      'hyperlinks.insert': { available: true },
      'create.sectionBreak': { available: true },
      'create.tableOfContents': { available: true },
      'create.table': { available: true },
      'headerFooters.parts.create': { available: true },
      'sections.setTitlePage': { available: true },
      'sections.setOddEvenHeadersFooters': { available: true },
      'headerFooters.refs.setLinkedToPrevious': { available: true },
      'fields.insert': { available: true },
      'fields.rebuild': { available: true },
      'bookmarks.list': { available: true },
      'bookmarks.insert': { available: true },
      'bookmarks.rename': { available: true },
      'bookmarks.remove': { available: true },
      'crossRefs.list': { available: true },
      'crossRefs.rebuild': { available: true },
      'toc.list': { available: true },
      'toc.update': { available: true },
      'toc.remove': { available: true },
      'toc.configure': { available: true },
      'toc.markEntry': { available: true },
      'toc.unmarkEntry': { available: true },
      'toc.listEntries': { available: true },
      'index.insert': { available: true },
      'index.list': { available: true },
      'index.rebuild': { available: true },
      'index.remove': { available: true },
      'index.configure': { available: true },
      'index.entries.list': { available: true },
      'index.entries.insert': { available: true },
      'index.entries.remove': { available: true },
      'blocks.list': { available: true },
      'blocks.deleteRange': { available: true },
      'fields.list': { available: true },
      'citations.list': { available: true },
      'citations.insert': { available: true },
      'captions.list': { available: true },
      'captions.insert': { available: true },
      'captions.remove': { available: true },
      'captions.update': { available: true },
      'captions.configure': { available: true },
      'citations.sources.list': { available: true },
      'citations.sources.insert': { available: true },
      'citations.sources.update': { available: true },
      'citations.sources.remove': { available: true },
      'citations.bibliography.insert': { available: true },
      'citations.bibliography.rebuild': { available: true },
      'citations.bibliography.remove': { available: true },
      'comments.create': { available: true },
      'clipboard.serializeSelection': { available: true },
      'clipboard.insert': { available: true },
      delete: { available: true },
      'ranges.resolve': { available: true },
    },
  };
}

function hostWith(get: () => unknown): CapabilitiesHost {
  return { activeEditor: { doc: { capabilities: { get: get as never } } } };
}

describe('readDocCapabilities', () => {
  it('עונה כן לכל השאלות כשהמנוע מדווח שהכול זמין', async () => {
    const report = await readDocCapabilities(hostWith(() => fullCapabilities()));

    expect(report.available).toBe(true);
    const denied = DOC_CAPABILITY_QUESTIONS.filter((q) => !report.can(q));
    expect(denied).toEqual([]);
  });

  it('שאלה זמינה אינה נושאת סיבה ואינה נושאת הסבר', async () => {
    const report = await readDocCapabilities(hostWith(() => fullCapabilities()));

    expect(report.reasons('canSetPageMargins')).toEqual([]);
    expect(report.explain('canSetPageMargins')).toBe('');
  });

  it('סובלת גם קריאה שמחזירה הבטחה — הפאסדה בדפדפן א-סינכרונית', async () => {
    const report = await readDocCapabilities(hostWith(() => Promise.resolve(fullCapabilities())));

    expect(report.available).toBe(true);
    expect(report.can('canInsertFootnote')).toBe(true);
  });

  it('קוראת ליכולות פעם אחת בלבד', async () => {
    // הדוח הוא תצלום. קריאה לכל שאלה הייתה מייצרת עשרות קריאות למנוע בכל
    // רינדור של הרצועה.
    const get = vi.fn(() => fullCapabilities());
    const report = await readDocCapabilities(hostWith(get));
    for (const question of DOC_CAPABILITY_QUESTIONS) report.can(question);

    expect(get).toHaveBeenCalledTimes(1);
  });

  it('אין Document API — הכול false עם DOCUMENT_API_UNAVAILABLE', async () => {
    for (const host of [
      null,
      undefined,
      {} as CapabilitiesHost,
      { activeEditor: null } as CapabilitiesHost,
      { activeEditor: { doc: null } } as CapabilitiesHost,
      { activeEditor: { doc: {} } } as CapabilitiesHost,
    ]) {
      const report = await readDocCapabilities(host);

      expect(report.available).toBe(false);
      for (const question of DOC_CAPABILITY_QUESTIONS) {
        expect(report.can(question), question).toBe(false);
        expect(report.reasons(question)).toEqual(['DOCUMENT_API_UNAVAILABLE']);
      }
      expect(report.explain('canSetPageMargins')).toBe('המסמך עדיין נטען');
    }
  });

  it('גרסה שאין בה `get` אינה מפילה ואינה מאשרת', async () => {
    const host = { activeEditor: { doc: { capabilities: {} } } } as CapabilitiesHost;

    const report = await readDocCapabilities(host);

    expect(report.available).toBe(false);
    expect(report.can('canSetColumns')).toBe(false);
  });

  it('קריאה שזורקת נכשלת סגור ואינה מפילה את הרצועה', async () => {
    const report = await readDocCapabilities(
      hostWith(() => {
        throw new Error('boom');
      }),
    );

    expect(report.available).toBe(false);
    expect(report.can('canSetPageMargins')).toBe(false);
  });

  it('הבטחה שנדחית נכשלת סגור', async () => {
    const report = await readDocCapabilities(hostWith(() => Promise.reject(new Error('boom'))));

    expect(report.available).toBe(false);
    expect(report.can('canSetPageMargins')).toBe(false);
  });

  it('תשובה שאינה אובייקט אינה תשובה', async () => {
    for (const value of [null, undefined, 'yes', 42, true]) {
      const report = await readDocCapabilities(hostWith(() => value));

      expect(report.available, String(value)).toBe(false);
      expect(report.can('canSetPageSetup')).toBe(false);
    }
  });

  it('מפה חלקית: פעולה שאינה בטבלה מקבלת OPERATION_UNAVAILABLE', async () => {
    // גרסה עתידית שתסיר פעולה תיראה בממשק ככפתור מנוטרל, לא ככפתור מת.
    const partial = fullCapabilities();
    delete (partial.operations as Record<string, unknown>)['sections.setColumns'];

    const report = await readDocCapabilities(hostWith(() => partial));

    expect(report.can('canSetColumns')).toBe(false);
    expect(report.reasons('canSetColumns')).toEqual(['OPERATION_UNAVAILABLE']);
    expect(report.explain('canSetColumns')).toBe('הפעולה אינה זמינה בגרסה הזאת של המנוע');
    // שאר השאלות לא נפגעו.
    expect(report.can('canSetPageMargins')).toBe(true);
  });

  it('namespace חסר: הערת שוליים מנוטרלת עם „אינו זמין בגרסה זו”', async () => {
    // `footnotes` הוא adapter אופציונלי בחוזה; כשהוא חסר המנוע מסמן את כל
    // הפעולות שלו NAMESPACE_UNAVAILABLE, וזה בדיוק הנוסח שהתכנית דורשת ב-§12.
    const raw = fullCapabilities();
    raw.operations['footnotes.insert'] = {
      available: false,
      reasons: ['NAMESPACE_UNAVAILABLE'],
    } as never;

    const report = await readDocCapabilities(hostWith(() => raw));

    expect(report.can('canInsertFootnote')).toBe(false);
    expect(report.reasons('canInsertFootnote')).toEqual(['NAMESPACE_UNAVAILABLE']);
    expect(report.explain('canInsertFootnote')).toBe('אינו זמין בגרסה זו');
  });

  it('namespace הלוח חסר: „העתק” מנוטרל, ו„מחק” אינו נפגע', async () => {
    // `clipboard` הוא adapter אופציונלי בחוזה, בדיוק כמו `footnotes`. שתי
    // השאלות מופרדות כדי שמנוע כזה יוכל להשאיר „גזור” מנוטרל בלי לנטרל את
    // המחיקה עצמה — ולהיפך.
    const raw = fullCapabilities();
    raw.operations['clipboard.serializeSelection'] = {
      available: false,
      reasons: ['NAMESPACE_UNAVAILABLE'],
    } as never;

    const report = await readDocCapabilities(hostWith(() => raw));

    expect(report.can('canCopySelection')).toBe(false);
    expect(report.explain('canCopySelection')).toBe('אינו זמין בגרסה זו');
    expect(report.can('canDeleteSelection')).toBe(true);
  });

  it('namespace הכותרות חסר: „כותרת עליונה” מנוטרל, והמתגים של `sections` לא', async () => {
    // `headerFooters` הוא namespace נפרד מ-`sections`, ומנוע שאין בו אותו
    // עדיין יודע „שונה בעמוד ראשון” — זו פעולה של `sections`. שתי השאלות
    // מופרדות בדיוק כדי שהפקדים לא ייכבו יחד בלי סיבה.
    const raw = fullCapabilities();
    raw.operations['headerFooters.parts.create'] = {
      available: false,
      reasons: ['NAMESPACE_UNAVAILABLE'],
    } as never;
    raw.operations['headerFooters.refs.setLinkedToPrevious'] = {
      available: false,
      reasons: ['NAMESPACE_UNAVAILABLE'],
    } as never;

    const report = await readDocCapabilities(hostWith(() => raw));

    expect(report.can('canEditHeaderFooter')).toBe(false);
    expect(report.explain('canEditHeaderFooter')).toBe('אינו זמין בגרסה זו');
    expect(report.can('canLinkToPrevious')).toBe(false);
    expect(report.can('canSetTitlePage')).toBe(true);
    expect(report.can('canSetOddEvenHeaders')).toBe(true);
  });

  it('namespace השדות חסר: „מספר עמוד” ו„עדכן שדות” מנוטרלים יחד', async () => {
    // `fields` הוא namespace שלם, ומנוע שאין בו אותו אינו יודע לא להכניס שדה
    // ולא לחשב אותו מחדש. שתי השאלות עדיין נפרדות (ראו CAPABILITY_SPECS),
    // ולכן הבדיקה מוודאת ששתיהן נכבות — ולא שאחת גוררת את השנייה.
    const raw = fullCapabilities();
    raw.operations['fields.insert'] = {
      available: false,
      reasons: ['NAMESPACE_UNAVAILABLE'],
    } as never;
    raw.operations['fields.rebuild'] = {
      available: false,
      reasons: ['NAMESPACE_UNAVAILABLE'],
    } as never;

    const report = await readDocCapabilities(hostWith(() => raw));

    expect(report.can('canInsertField')).toBe(false);
    expect(report.explain('canInsertField')).toBe('אינו זמין בגרסה זו');
    expect(report.can('canRebuildFields')).toBe(false);
  });

  it('מנוע שיודע להכניס שדה ואינו יודע לחשב מחדש: גם ההכנסה מנוטרלת', async () => {
    // הבדיקה הזאת קיבעה קודם את ההפך — „ההכנסה נשארת פעילה” — מתוך הנחה שאין
    // לה תלות ב-`rebuild`. יש: `fields.insert` מכניס שדה עם תוצאה **ריקה**,
    // וה-rebuild הוא שמחשב אותה. בלעדיו המשתמש מקבל דיווח „בוצע” ורואה מקום
    // ריק במסמך. שדה שאי אפשר לראות אינו פיצ'ר, ולכן הכפתור מנוטרל עם הסבר.
    const raw = fullCapabilities();
    raw.operations['fields.rebuild'] = {
      available: false,
      reasons: ['OPERATION_UNAVAILABLE'],
    } as never;

    const report = await readDocCapabilities(hostWith(() => raw));

    expect(report.can('canInsertField')).toBe(false);
    expect(report.explain('canInsertField')).toBe('הפעולה אינה זמינה בגרסה הזאת של המנוע');
    expect(report.can('canRebuildFields')).toBe(false);
  });

  it('מנוע שיודע לחשב מחדש ואינו יודע להכניס: „עדכן שדות” נשאר פעיל', async () => {
    // זו הסיבה ששתי השאלות נשארו נפרדות אחרי שההכנסה נעשתה תלויה גם ב-rebuild:
    // הן אינן זהות. מסמך שכבר יש בו שדות עדיין ניתן לעדכון גם כשאי אפשר
    // להוסיף שדה חדש.
    const raw = fullCapabilities();
    raw.operations['fields.insert'] = {
      available: false,
      reasons: ['OPERATION_UNAVAILABLE'],
    } as never;

    const report = await readDocCapabilities(hostWith(() => raw));

    expect(report.can('canInsertField')).toBe(false);
    expect(report.can('canRebuildFields')).toBe(true);
  });

  it('פעולת סימניות אחת שחסרה מנטרלת את כל הפקד', async () => {
    // „סימנייה” הוא כפתור אחד שפותח דיאלוג שמוסיף, מוחק ומשנה שם. פקד מנוטרל
    // למחצה אינו מצב שאפשר להציג, ולכן השאלה דורשת את כל ארבע הפעולות —
    // וכאן נמדד שגם היעדר `rename` לבדו מספיק כדי לכבות אותו.
    const raw = fullCapabilities();
    raw.operations['bookmarks.rename'] = {
      available: false,
      reasons: ['OPERATION_UNAVAILABLE'],
    } as never;

    const report = await readDocCapabilities(hostWith(() => raw));

    expect(report.can('canManageBookmarks')).toBe(false);
    expect(report.explain('canManageBookmarks')).toBe('הפעולה אינה זמינה בגרסה הזאת של המנוע');
  });

  it('כל פעולות הסימניות קיימות — הפקד פעיל', async () => {
    const report = await readDocCapabilities(hostWith(() => fullCapabilities()));

    expect(report.can('canManageBookmarks')).toBe(true);
    expect(report.explain('canManageBookmarks')).toBe('');
  });

  it('„עדכן הפניות” דורש גם מנייה וגם חישוב מחדש', async () => {
    // `crossRefs.rebuild` מקבל כתובת של הפניה מסוימת, ואין דרך אחרת להשיג
    // אותה מלבד `crossRefs.list`. מנוע שיודע לחשב ואינו יודע למנות היה
    // מדווח „בוצע” בלי לעדכן דבר. אותה תלות בדיוק כמו ב-`canInsertField`.
    const withoutList = fullCapabilities();
    withoutList.operations['crossRefs.list'] = {
      available: false,
      reasons: ['NAMESPACE_UNAVAILABLE'],
    } as never;
    expect(
      (await readDocCapabilities(hostWith(() => withoutList))).can('canRebuildCrossRefs'),
    ).toBe(false);

    const withoutRebuild = fullCapabilities();
    withoutRebuild.operations['crossRefs.rebuild'] = {
      available: false,
      reasons: ['NAMESPACE_UNAVAILABLE'],
    } as never;
    expect(
      (await readDocCapabilities(hostWith(() => withoutRebuild))).can('canRebuildCrossRefs'),
    ).toBe(false);

    expect(
      (await readDocCapabilities(hostWith(() => fullCapabilities()))).can('canRebuildCrossRefs'),
    ).toBe(true);
  });

  it('„הסר תוכן עניינים” דורש גם את מחיקת השורות שנשארות', async () => {
    // `toc.remove` מוחק את הבלוק הראשון של הטבלה בלבד ומשאיר את שאר השורות
    // כפסקאות `TOC1`…`TOC9` (נמדד בדפדפן). מנוע בלי `blocks.deleteRange`
    // היה מציג „הסר” שמחזיר „בוצע” ומשאיר את גוף הטבלה על המסך.
    const withoutSweep = fullCapabilities();
    withoutSweep.operations['blocks.deleteRange'] = {
      available: false,
      reasons: ['NAMESPACE_UNAVAILABLE'],
    } as never;
    expect(
      (await readDocCapabilities(hostWith(() => withoutSweep))).can('canRemoveTableOfContents'),
    ).toBe(false);
    // ואילו „עדכן טבלה” אינו נשען עליה, ולכן הוא נשאר פעיל.
    expect(
      (await readDocCapabilities(hostWith(() => withoutSweep))).can('canUpdateTableOfContents'),
    ).toBe(true);
  });

  it('„סמן ערך” דורש את שלוש הפעולות שהדיאלוג מריץ', async () => {
    // הדיאלוג מסמן, מציג את הערכים הקיימים ומבטל סימון. פקד מנוטרל למחצה
    // אינו מצב שאפשר להציג, ולכן היעדר כל אחת מהשלוש מנטרל אותו.
    for (const operation of ['toc.markEntry', 'toc.unmarkEntry', 'toc.listEntries']) {
      const raw = fullCapabilities();
      raw.operations[operation as 'toc.markEntry'] = {
        available: false,
        reasons: ['NAMESPACE_UNAVAILABLE'],
      } as never;
      expect(
        (await readDocCapabilities(hostWith(() => raw))).can('canMarkTocEntry'),
        operation,
      ).toBe(false);
    }

    expect(
      (await readDocCapabilities(hostWith(() => fullCapabilities()))).can('canMarkTocEntry'),
    ).toBe(true);
  });

  it('„הוסף מפתח” אינו נשען על `index.rebuild`, ו„הסר מפתח” אינו נשען על `blocks.*`', async () => {
    // שתי ההבחנות האלה נמדדו במנוע, והן מה שמבדיל את המפתח מתוכן העניינים:
    // `index.insert` מרנדר את המפתח מלא כבר ביצירה, ו-`index.remove` מוחק את
    // הבלוק כולו בלי להשאיר פסקאות יתומות. שאלה שהייתה מונה גם אותן הייתה
    // מנטרלת פקד עובד. ההנמקה המלאה ב-engine/index-field.ts.
    const raw = fullCapabilities();
    raw.operations['index.rebuild'] = { available: false } as never;
    raw.operations['blocks.deleteRange'] = { available: false } as never;

    const report = await readDocCapabilities(hostWith(() => raw));
    expect(report.can('canInsertIndex')).toBe(true);
    expect(report.can('canRemoveIndex')).toBe(true);
    // ואילו „עדכן מפתח” כן נשען עליה.
    expect(report.can('canRebuildIndex')).toBe(false);
  });

  it('„סמן ערך למפתח” דורש את שלוש הפעולות שהדיאלוג מריץ', async () => {
    for (const operation of [
      'index.entries.list',
      'index.entries.insert',
      'index.entries.remove',
    ]) {
      const raw = fullCapabilities();
      raw.operations[operation as 'index.entries.list'] = {
        available: false,
        reasons: ['NAMESPACE_UNAVAILABLE'],
      } as never;
      expect(
        (await readDocCapabilities(hostWith(() => raw))).can('canMarkIndexEntry'),
        operation,
      ).toBe(false);
    }

    expect(
      (await readDocCapabilities(hostWith(() => fullCapabilities()))).can('canMarkIndexEntry'),
    ).toBe(true);
  });

  it('„מחק מקור” נשען על `citations.list`, ו„עדכן ביבליוגרפיה” על `fields.list`', async () => {
    // שתי התלויות האלה נראות עקיפות ושתיהן נמדדו כהכרחיות: בלי `citations.list`
    // אי אפשר לדעת אילו ציטוטים מפנים אל מקור, והמחיקה הייתה משאירה שדה
    // `CITATION` מצביע לתג שאינו קיים; ובלי `fields.list` אין בכלל דרך למצוא
    // ביבליוגרפיה במסמך — `citations.bibliography` חסר `list`, ו-`blocks.list`
    // מציג אותה כפסקה רגילה. ההנמקה המלאה ב-engine/citations.ts.
    const withoutCitations = fullCapabilities();
    withoutCitations.operations['citations.list'] = { available: false } as never;
    const first = await readDocCapabilities(hostWith(() => withoutCitations));
    expect(first.can('canManageCitationSources')).toBe(false);
    // ואילו „הוסף ציטוט” אינו נשען עליה — הוא כותב ואינו סופר.
    expect(first.can('canInsertCitation')).toBe(true);

    const withoutFields = fullCapabilities();
    withoutFields.operations['fields.list'] = { available: false } as never;
    const second = await readDocCapabilities(hostWith(() => withoutFields));
    expect(second.can('canRebuildBibliography')).toBe(false);
    expect(second.can('canRemoveBibliography')).toBe(false);
    // ואילו ההוספה אינה צריכה למצוא כלום.
    expect(second.can('canInsertBibliography')).toBe(true);
  });

  it('„הוסף כיתוב” דורש את ארבע הפעולות שהדיאלוג מריץ — כולל `blocks.list`', async () => {
    // `blocks.list` נראה עקיף וגם הוא נמדד כהכרחי: העריכה היא `remove`+`insert`
    // (`captions.update` מוסיף את הטקסט החדש על הישן), והמקום שאליו הכיתוב
    // חוזר נגזר מהבלוק שלפניו. ההנמקה המלאה ב-engine/captions.ts.
    for (const operation of ['captions.list', 'captions.insert', 'captions.remove', 'blocks.list']) {
      const raw = fullCapabilities();
      raw.operations[operation as 'captions.list'] = {
        available: false,
        reasons: ['NAMESPACE_UNAVAILABLE'],
      } as never;
      expect(
        (await readDocCapabilities(hostWith(() => raw))).can('canManageCaptions'),
        operation,
      ).toBe(false);
    }

    expect(
      (await readDocCapabilities(hostWith(() => fullCapabilities()))).can('canManageCaptions'),
    ).toBe(true);
  });

  it('„נהל הערות” דורש את ארבע הפעולות שהדיאלוג מריץ — כולל `footnotes.get`', async () => {
    // `footnotes.get` נראה עקיף וגם הוא נמדד כהכרחי: כתובת ההערה אינה נושאת
    // את סוגה, ו-`get` הוא הדרך היחידה לשאול את המנוע לאיזו הערה היא נפתרת
    // לפני שנוגעים במסמך. ההנמקה המלאה ב-engine/footnotes.ts.
    for (const operation of [
      'footnotes.list',
      'footnotes.get',
      'footnotes.update',
      'footnotes.remove',
    ]) {
      const raw = fullCapabilities();
      raw.operations[operation as 'footnotes.list'] = {
        available: false,
        reasons: ['NAMESPACE_UNAVAILABLE'],
      } as never;
      expect(
        (await readDocCapabilities(hostWith(() => raw))).can('canManageNotes'),
        operation,
      ).toBe(false);
    }

    expect(
      (await readDocCapabilities(hostWith(() => fullCapabilities()))).can('canManageNotes'),
    ).toBe(true);
    // וההוספה אינה נשענת על אף אחת מהן — היא כותבת ואינה מחפשת.
    const onlyInsert = fullCapabilities();
    onlyInsert.operations['footnotes.get'] = { available: false } as never;
    expect((await readDocCapabilities(hostWith(() => onlyInsert))).can('canInsertFootnote')).toBe(
      true,
    );
  });

  it('„מספור ההערות” אינו שאלה בכלל', async () => {
    // `footnotes.configure` כותב OOXML קנוני, ובכל זאת אין לו פקד: אין דרך
    // לקרוא את ההגדרות שבמסמך, וכל קריאה מחליפה את `w:footnotePr` כולו.
    // שאלה בלי פקד היא הצהרת יכולת שאיש אינו קורא.
    const raw = fullCapabilities();
    (raw.operations as Record<string, { available: boolean }>)['footnotes.configure'] = {
      available: false,
    };
    expect((await readDocCapabilities(hostWith(() => raw))).can('canManageNotes')).toBe(true);
  });

  it('„מספור הכיתובים” אינו שאלה בכלל', async () => {
    // `captions.configure` מוצהר זמין, מחזיר `success: true`, ואינו משנה דבר
    // בקוד השדה (נמדד — `format: 'upperRoman'` נבלע, וכך גם `'zigzag'`). אין
    // לו פקד, ושאלה בלי פקד היא הצהרת יכולת שאיש אינו קורא. וכך גם
    // `captions.update`, שהמודול אינו קורא לו כלל.
    const raw = fullCapabilities();
    raw.operations['captions.configure'] = { available: false } as never;
    raw.operations['captions.update'] = { available: false } as never;
    expect((await readDocCapabilities(hostWith(() => raw))).can('canManageCaptions')).toBe(true);
  });

  it('„סגנון הביבליוגרפיה” אינו שאלה בכלל', async () => {
    // `citations.bibliography.configure` מוצהר זמין ועובד למחצה: הוא כותב את
    // הסגנון למקום הנכון ובאותה קריאה גם מתג `\sdStyle` שאינו של Word. אין
    // לו פקד, ושאלה בלי פקד היא הצהרת יכולת שאיש אינו קורא.
    expect(
      DOC_CAPABILITY_QUESTIONS.some((question) => question.toLowerCase().includes('style')),
    ).toBe(false);
  });

  it('`available` שאינו בדיוק true אינו „כן”', async () => {
    // המנוע מחזיר boolean, אבל תשובה מ-worker שעברה סריאליזציה עלולה להגיע
    // כמחרוזת. „truthy” אינו „זמין”.
    const raw = fullCapabilities();
    raw.operations['sections.setPageMargins'] = { available: 'true' } as never;

    const report = await readDocCapabilities(hostWith(() => raw));

    expect(report.can('canSetPageMargins')).toBe(false);
  });

  it('דגל global כבוי חוסם גם כשהפעולה זמינה', async () => {
    // „תגובה חדשה” דורשת גם את הפעולה וגם את הדגל: הראשון אומר שהיא קיימת,
    // השני שהיא מאופשרת במסמך הזה.
    const raw = fullCapabilities();
    raw.global.comments = { enabled: false, reasons: ['COLLABORATION_ACTIVE'] } as never;

    const report = await readDocCapabilities(hostWith(() => raw));

    expect(report.can('canAddComment')).toBe(false);
    expect(report.reasons('canAddComment')).toEqual(['COLLABORATION_ACTIVE']);
    expect(report.can('canInsertFootnote')).toBe(true);
  });

  it('מעקב אחר שינויים נשען על הדגל הגלובלי בלבד', async () => {
    const raw = fullCapabilities();
    raw.global.trackChanges = { enabled: false } as never;

    const report = await readDocCapabilities(hostWith(() => raw));

    expect(report.can('canTrackChanges')).toBe(false);
    expect(report.reasons('canTrackChanges')).toEqual(['OPERATION_UNAVAILABLE']);
  });

  it('אין `global` ואין `operations` — הכול סגור, אבל ה-API עצמו זמין', async () => {
    const report = await readDocCapabilities(hostWith(() => ({})));

    expect(report.available).toBe(true);
    const allowed = DOC_CAPABILITY_QUESTIONS.filter((q) => report.can(q));
    expect(allowed).toEqual([]);
  });

  it('קוד סיבה שאיננו מכירים אינו מפיל ואינו מוצג כזבל', async () => {
    const raw = fullCapabilities();
    raw.operations['sections.setColumns'] = {
      available: false,
      reasons: ['SOME_FUTURE_CODE'],
    } as never;

    const report = await readDocCapabilities(hostWith(() => raw));

    expect(report.can('canSetColumns')).toBe(false);
    // הקוד הלא-מוכר נופל, ובמקומו נשארת סיבה גנרית עם הסבר בעברית.
    expect(report.reasons('canSetColumns')).toEqual(['OPERATION_UNAVAILABLE']);
    expect(report.explain('canSetColumns')).not.toBe('');
  });

  it('שאלה שאינה מוכרת מקבלת תשובה סגורה ולא undefined', async () => {
    const report = await readDocCapabilities(hostWith(() => fullCapabilities()));
    const bogus = 'canDoSomethingElse' as DocCapabilityQuestion;

    expect(report.can(bogus)).toBe(false);
    expect(report.explain(bogus)).toBe('המסמך עדיין נטען');
  });
});
