/**
 * המצב של העץ — מה שמבטיח שהדיאלוג נפתח מהר.
 *
 * שלוש טענות מהירות, וכל אחת נשברת בשקט: העץ מגיע מהמטמון לפני תשובת
 * אוצריא; רענון שנכשל אינו מוחק עץ שעובד; וחזרה לתיקייה מציגה מיד את מה
 * שנראה קודם, בזמן שהרענון רץ ברקע.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';

const host = vi.hoisted(() => ({
  fetchLibraryShelf: vi.fn(),
  listDocFolder: vi.fn(),
  pickDocFolder: vi.fn(),
  revokeDocFolder: vi.fn(async () => undefined),
}));
const storage = vi.hoisted(() => ({
  values: {} as Record<string, unknown>,
  saved: [] as Array<[string, unknown]>,
}));

vi.mock('../../src/host/open-sources', () => host);
vi.mock('../../src/host/settings', () => ({
  loadDocFolders: async () => storage.values.folders,
  loadLibraryCache: async () => storage.values.cache,
  loadOpenDialogPlace: async () => storage.values.place,
  saveDocFolders: async (v: unknown) => void storage.saved.push(['folders', v]),
  saveLibraryCache: async (v: unknown) => void storage.saved.push(['cache', v]),
  saveOpenDialogPlace: async (v: unknown) => void storage.saved.push(['place', v]),
}));

import { useOpenSources } from '../../src/composables/use-open-sources';
import { LIBRARY_CACHE_VERSION, listingKey, pruneLibraryTree, placeKey } from '../../src/sessions/open-sources';

const CACHED = pruneLibraryTree({ categories: [{ title: 'ישן', books: [{ title: 'x', type: 'docx' }] }] });
const FRESH = pruneLibraryTree({ categories: [{ title: 'חדש', books: [{ title: 'y', type: 'docx' }] }] });
const listing = (name: string) => ({ entries: [{ name, path: name, type: 'file', size: 1, modified: 0 }], truncated: false });

beforeEach(() => {
  for (const fn of Object.values(host)) fn.mockReset();
  storage.values = {};
  storage.saved = [];
});

describe('useOpenSources — הספרייה', () => {
  it('init מציג את העץ מהמטמון, בלי לקרוא לאוצריא', async () => {
    storage.values.cache = { version: LIBRARY_CACHE_VERSION, savedAt: 1, tree: CACHED };
    const sources = useOpenSources();
    await sources.init();
    expect(sources.library.value).toEqual(CACHED);
    expect(sources.libraryState.value).toBe('ready');
    expect(host.fetchLibraryShelf).not.toHaveBeenCalled();
  });

  it('הרענון מחליף את העץ ושומר אותו למטמון — פעם אחת בהפעלה', async () => {
    host.fetchLibraryShelf.mockResolvedValue({ ok: true, value: FRESH });
    const sources = useOpenSources();
    await sources.refreshLibrary();
    await sources.refreshLibrary();
    expect(host.fetchLibraryShelf).toHaveBeenCalledTimes(1);
    expect(sources.library.value).toEqual(FRESH);
    expect(storage.saved.find(([k]) => k === 'cache')?.[1]).toMatchObject({ tree: FRESH });
    await sources.refreshLibrary(true);
    expect(host.fetchLibraryShelf).toHaveBeenCalledTimes(2);
  });

  it('רענון שנכשל משאיר את העץ מהמטמון על המסך', async () => {
    storage.values.cache = { version: LIBRARY_CACHE_VERSION, savedAt: 1, tree: CACHED };
    host.fetchLibraryShelf.mockResolvedValue({ ok: false, reason: 'failed', message: 'x' });
    const sources = useOpenSources();
    await sources.refreshLibrary();
    expect(sources.library.value).toEqual(CACHED);
    expect(sources.libraryState.value).toBe('ready');
  });

  it('„לא נתמך” מסתיר את הענף גם כשיש מטמון', async () => {
    storage.values.cache = { version: LIBRARY_CACHE_VERSION, savedAt: 1, tree: CACHED };
    host.fetchLibraryShelf.mockResolvedValue({ ok: false, reason: 'unsupported', message: 'x' });
    const sources = useOpenSources();
    await sources.refreshLibrary();
    expect(sources.libraryState.value).toBe('unsupported');
  });
});

describe('useOpenSources — תיקיות', () => {
  it('חזרה לתיקייה מציגה מיד את הישן, ומסמנת רענון עד שהחדש מגיע', async () => {
    host.listDocFolder.mockResolvedValueOnce({ ok: true, value: listing('א.docx') });
    const sources = useOpenSources();
    await sources.listFolder('ft', '');

    let resolve!: (v: unknown) => void;
    host.listDocFolder.mockReturnValueOnce(new Promise((r) => (resolve = r)));
    const pending = sources.listFolder('ft', '');
    const during = sources.listingOf('ft', '');
    expect(during).toMatchObject({ state: 'ready', refreshing: true });
    expect(during?.state === 'ready' && during.listing.entries[0]?.name).toBe('א.docx');

    resolve({ ok: true, value: listing('ב.docx') });
    await pending;
    const after = sources.listingOf('ft', '');
    expect(after?.state === 'ready' && after.listing.entries[0]?.name).toBe('ב.docx');
  });

  it('שתי בקשות בו-זמנית לאותה תיקייה הן בקשה אחת', async () => {
    host.listDocFolder.mockResolvedValue({ ok: true, value: listing('א.docx') });
    const sources = useOpenSources();
    await Promise.all([sources.listFolder('ft', ''), sources.listFolder('ft', '')]);
    expect(host.listDocFolder).toHaveBeenCalledTimes(1);
  });

  it('ensureListing אינו טוען שוב מה שכבר בזיכרון', async () => {
    host.listDocFolder.mockResolvedValue({ ok: true, value: listing('א.docx') });
    const sources = useOpenSources();
    await sources.listFolder('ft', 'sub');
    sources.ensureListing('ft', 'sub');
    expect(host.listDocFolder).toHaveBeenCalledTimes(1);
  });

  it('הוספת תיקייה שומרת אותה ובוחרת אותה מיד', async () => {
    host.pickDocFolder.mockResolvedValue({ ok: true, value: { token: 'ft', name: 'חדשה', path: 'C:\\חדשה', addedAt: 5 } });
    host.listDocFolder.mockResolvedValue({ ok: true, value: listing('א.docx') });
    const sources = useOpenSources();
    await sources.addDocFolder();
    expect(sources.folders.value.map((f) => f.token)).toEqual(['ft']);
    expect(sources.place.value).toEqual({ kind: 'folder', token: 'ft', path: '' });
    expect(storage.saved.some(([k]) => k === 'folders')).toBe(true);
    expect(storage.saved.find(([k]) => k === 'place')?.[1]).toBe(placeKey({ kind: 'folder', token: 'ft', path: '' }));
  });

  it('הסרת התיקייה הנבחרת: חוזרים ל„אחרונים”, הפירוטים שלה נמחקים, וההרשאה מבוטלת', async () => {
    storage.values.folders = [{ token: 'ft', name: 'x', path: '', addedAt: 1 }];
    storage.values.place = placeKey({ kind: 'folder', token: 'ft', path: '' });
    host.listDocFolder.mockResolvedValue({ ok: true, value: listing('א.docx') });
    const sources = useOpenSources();
    await sources.init();
    await sources.listFolder('ft', '');
    sources.removeDocFolder('ft');
    expect(sources.place.value).toEqual({ kind: 'recent' });
    expect(sources.listings.has(listingKey('ft', ''))).toBe(false);
    expect(host.revokeDocFolder).toHaveBeenCalledWith('ft');
  });
});
