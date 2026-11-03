/**
 * המקורות של „פתח מסמך” — מה מופיע בעץ, ומה נעלם ממנו בשקט.
 *
 * כל מה שנבדק כאן הוא החלטה שכשל בה אינו זורק: ספר שנגזם בטעות פשוט אינו
 * מופיע, תיקייה שנשמטה מ-storage פשוט אינה בעץ, וחיפוש שמדרג הפוך פשוט מראה
 * את הספר הלא-נכון ראשון.
 */
import { describe, expect, it } from 'vitest';
import {
  addFolder,
  findShelf,
  normalizeFolders,
  normalizeLibraryCache,
  normalizeListing,
  parsePlace,
  personalBooks,
  placeKey,
  pruneLibraryTree,
  removeFolder,
  searchLibrary,
  shelfTrail,
  withoutPersonalShelves,
  LIBRARY_CACHE_VERSION,
  type DocFolder,
} from '../../src/sessions/open-sources';

/** עץ בצורה של `library.getTree`, כולל ספרים שאינם Word שחייבים להיגזם. */
const RAW_TREE = {
  title: 'ספריית אוצריא',
  path: '/',
  books: [],
  categories: [
    {
      title: 'הלכה',
      path: '/הלכה',
      books: [
        { bookId: 'קונטרס', title: 'קונטרס', type: 'docx', bookUid: 'uid:1', source: 'library', author: 'פלוני' },
        { bookId: 'משנה ברורה', title: 'משנה ברורה', type: 'text' },
      ],
      categories: [
        {
          title: 'שו"ת',
          path: '/הלכה/שו"ת',
          books: [{ bookId: 'שו"ת דברי שלום', title: 'שו"ת דברי שלום', type: 'DOCM', id: 7 }],
          categories: [],
        },
        // ריקה מ-Word — חייבת להיעלם, גם כשיש בה ספרים אחרים.
        { title: 'פוסקים', path: '/הלכה/פוסקים', books: [{ title: 'טור', type: 'text' }], categories: [] },
      ],
    },
    {
      title: 'אישי',
      path: '/אישי',
      books: [{ bookId: 'חידושים', title: 'חידושים', type: 'docx', source: 'user' }],
      categories: [],
    },
    { title: 'תנך', path: '/תנך', books: [{ title: 'בראשית', type: 'text' }], categories: [] },
  ],
};

describe('pruneLibraryTree', () => {
  it('משאיר ספרי Word בלבד, וגוזם מדפים שאין מתחתם כאלה', () => {
    const tree = pruneLibraryTree(RAW_TREE)!;
    expect(tree.path).toBe('/');
    expect(tree.total).toBe(3);
    expect(tree.shelves.map((s) => s.title)).toEqual(['הלכה', 'אישי']);
    const halacha = tree.shelves[0]!;
    expect(halacha.books.map((b) => b.title)).toEqual(['קונטרס']);
    expect(halacha.shelves.map((s) => s.title)).toEqual(['שו"ת']);
    expect(halacha.total).toBe(2);
  });

  it('הסוג מושווה בלי תלות באותיות, ו-docm נחשב Word', () => {
    const shut = findShelf(pruneLibraryTree(RAW_TREE), '/הלכה/שו"ת')!;
    expect(shut.books[0]).toMatchObject({ type: 'docm', id: 7, bookUid: null });
  });

  it('ספרייה בלי אף ספר Word היא null, לא שורש ריק', () => {
    expect(pruneLibraryTree({ title: 'x', categories: [{ title: 'y', books: [{ title: 'z', type: 'pdf' }] }] })).toBeNull();
  });

  it('קלט פגום אינו זורק', () => {
    expect(pruneLibraryTree(null)).toBeNull();
    expect(pruneLibraryTree('x')).toBeNull();
    expect(pruneLibraryTree({ categories: 'x', books: [null, 3, { type: 'docx' }] })).toBeNull();
  });
});

describe('ניווט בעץ', () => {
  const tree = pruneLibraryTree(RAW_TREE);

  it('findShelf מוצא מדף מקונן, ו-"/" הוא השורש', () => {
    expect(findShelf(tree, '/')).toBe(tree);
    expect(findShelf(tree, '/הלכה/שו"ת')?.title).toBe('שו"ת');
    expect(findShelf(tree, '/אין')).toBeNull();
  });

  it('shelfTrail נותן את השרשרת מהשורש', () => {
    expect(shelfTrail(tree, '/הלכה/שו"ת').map((s) => s.title)).toEqual(['ספריית אוצריא', 'הלכה', 'שו"ת']);
    expect(shelfTrail(tree, '/').map((s) => s.path)).toEqual(['/']);
    expect(shelfTrail(tree, '/נעלם')).toEqual([]);
  });

  it('personalBooks אוסף את הספרים האישיים מכל העץ', () => {
    expect(personalBooks(tree).map((b) => b.title)).toEqual(['חידושים']);
  });
});

describe('withoutPersonalShelves', () => {
  it('מדף שכולו אישי יורד, והספירה מתעדכנת', () => {
    const view = withoutPersonalShelves(pruneLibraryTree(RAW_TREE))!;
    expect(view.shelves.map((s) => s.title)).toEqual(['הלכה']);
    expect(view.total).toBe(2);
  });

  it('מדף מעורב נשאר כמות שהוא, ועץ בלי אישיים חוזר כאותו אובייקט', () => {
    const mixed = pruneLibraryTree({
      categories: [{ title: 'א', books: [{ title: 'x', type: 'docx', source: 'user' }, { title: 'y', type: 'docx' }] }],
    });
    expect(withoutPersonalShelves(mixed)).toBe(mixed);
  });

  it('ספרייה שכולה אישית היא null — אין ענף ספרייה ריק', () => {
    expect(withoutPersonalShelves(pruneLibraryTree({ categories: [{ title: 'א', books: [{ title: 'x', type: 'docx', source: 'user' }] }] }))).toBeNull();
  });
});

describe('searchLibrary', () => {
  const tree = pruneLibraryTree(RAW_TREE);

  it('כותרת שמתחילה בשאילתה קודמת לכותרת שמכילה אותה', () => {
    const extra = pruneLibraryTree({
      categories: [
        {
          title: 'א',
          books: [
            { title: 'ספר שלום', type: 'docx' },
            { title: 'שלום עליכם', type: 'docx' },
          ],
        },
      ],
    });
    expect(searchLibrary(extra, 'שלום').map((h) => h.book.title)).toEqual(['שלום עליכם', 'ספר שלום']);
  });

  it('גרשיים עבריים מוצאים גרשיים ישרים', () => {
    expect(searchLibrary(tree, 'שו״ת').map((h) => h.book.title)).toEqual(['שו"ת דברי שלום']);
  });

  it('מחבר נמצא, והמיקום בעץ חוזר כשורה שנייה', () => {
    const [hit] = searchLibrary(tree, 'פלוני');
    expect(hit?.book.title).toBe('קונטרס');
    expect(hit?.where).toBe('הלכה');
  });

  it('שאילתה ריקה אינה מחזירה דבר, והתקרה נאכפת', () => {
    expect(searchLibrary(tree, '   ')).toEqual([]);
    expect(searchLibrary(tree, 'ש', 1)).toHaveLength(1);
  });
});

describe('normalizeLibraryCache', () => {
  const tree = pruneLibraryTree(RAW_TREE);

  it('מקבל את הצורה הנוכחית, כולל עץ null', () => {
    expect(normalizeLibraryCache({ version: LIBRARY_CACHE_VERSION, savedAt: 1, tree })?.tree).toEqual(tree);
    expect(normalizeLibraryCache({ version: LIBRARY_CACHE_VERSION, savedAt: 1, tree: null })?.tree).toBeNull();
  });

  it('גרסה אחרת או צורה שבורה — null', () => {
    expect(normalizeLibraryCache({ version: 0, savedAt: 1, tree })).toBeNull();
    expect(normalizeLibraryCache({ version: LIBRARY_CACHE_VERSION, savedAt: 1, tree: { title: 'x' } })).toBeNull();
    expect(normalizeLibraryCache({ version: LIBRARY_CACHE_VERSION, tree })).toBeNull();
    expect(normalizeLibraryCache(undefined)).toBeNull();
  });
});

describe('התיקיות', () => {
  const a: DocFolder = { token: 'a', name: 'חידושים', path: 'C:\\חידושים', addedAt: 1 };

  it('normalizeFolders משמיט שורות פגומות ו-token כפול, ומשלים שם מהנתיב', () => {
    expect(
      normalizeFolders([a, { ...a, name: 'כפול' }, { token: '' }, null, { token: 'b', path: 'D:\\x\\מכתבים' }]),
    ).toEqual([a, { token: 'b', name: 'מכתבים', path: 'D:\\x\\מכתבים', addedAt: 0 }]);
    expect(normalizeFolders('x')).toEqual([]);
  });

  it('addFolder על token קיים מעדכן את השורה ואינו מוסיף שנייה', () => {
    const next = addFolder([a], { ...a, name: 'שם חדש', addedAt: 99 });
    expect(next).toEqual([{ ...a, name: 'שם חדש' }]);
  });

  it('removeFolder מסיר לפי token', () => {
    expect(removeFolder([a], 'a')).toEqual([]);
  });
});

describe('normalizeListing', () => {
  it('תיקיות לפני קבצים, כל קבוצה לפי שם, ושורות פגומות נשמטות', () => {
    const listing = normalizeListing({
      entries: [
        { name: 'ב.docx', path: 'ב.docx', type: 'file', size: 10, modified: '2026-09-01T00:00:00Z' },
        { name: 'תיקייה', path: 'תיקייה', type: 'dir', size: 0, modified: null },
        { name: 'א.docx', path: 'א.docx', type: 'file', size: -1 },
        { name: '', path: 'x', type: 'file' },
        { name: 'y', path: 'y', type: 'link' },
      ],
      truncated: true,
    });
    expect(listing.entries.map((e) => e.name)).toEqual(['תיקייה', 'א.docx', 'ב.docx']);
    expect(listing.entries[1]).toMatchObject({ size: 0, modified: 0 });
    expect(listing.entries[2]!.modified).toBe(Date.parse('2026-09-01T00:00:00Z'));
    expect(listing.truncated).toBe(true);
  });

  it('תשובה ריקה או פגומה היא רשימה ריקה', () => {
    expect(normalizeListing(null)).toEqual({ entries: [], truncated: false });
  });
});

describe('המקום בעץ', () => {
  const folders: DocFolder[] = [{ token: 'tok', name: 'x', path: '', addedAt: 0 }];

  it('placeKey ו-parsePlace הם הלוך-חזור', () => {
    for (const place of [
      { kind: 'recent' },
      { kind: 'personal' },
      { kind: 'library', path: '/הלכה/שו"ת' },
      { kind: 'folder', token: 'tok', path: 'שיעורים/א' },
    ] as const) {
      expect(parsePlace(placeKey(place), folders)).toEqual(place);
    }
  });

  it('תיקייה שהוסרה, או ערך לא מוכר — חוזרים ל„אחרונים”', () => {
    expect(parsePlace(placeKey({ kind: 'folder', token: 'gone', path: '' }), folders)).toEqual({ kind: 'recent' });
    expect(parsePlace('library\u0000no-slash', folders)).toEqual({ kind: 'recent' });
    expect(parsePlace(42, folders)).toEqual({ kind: 'recent' });
  });
});
