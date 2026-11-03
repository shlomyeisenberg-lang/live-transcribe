/**
 * הקריאות שמאחורי העץ — מה נשלח לאוצריא, ומה חוזר לדיאלוג כשהיא מסרבת.
 *
 * שני המסלולים שכשל בהם שקט: ספר אישי שנדחה לכתיבה חייב ליפול לקריאה (ולא
 * להיכשל), וגרסת אוצריא שאינה מכירה את הקריאה חייבת להחזיר `unsupported` —
 * שעליו הדיאלוג מסתיר את הענף — ולא „נכשל” שמופיע בכל פתיחה.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';

const callMock = vi.hoisted(() => vi.fn());

vi.mock('../../src/host/otzaria-client', async (importOriginal) => {
  const real = await importOriginal<typeof import('../../src/host/otzaria-client')>();
  return { ...real, call: callMock, tryCall: vi.fn(async () => true) };
});

import { HostCallError } from '../../src/host/otzaria-client';
import {
  OPEN_SOURCE_PERMISSIONS,
  fetchLibraryShelf,
  listDocFolder,
  openDocFolderFile,
  openLibraryBook,
  pickDocFolder,
} from '../../src/host/open-sources';
import type { LibraryBook } from '../../src/sessions/open-sources';

const denied = (method: string) => new HostCallError('denied', 'error.permission_denied', method);
const unknown = (method: string) => new HostCallError('nope', 'error.unknown_method', method);

function book(patch: Partial<LibraryBook> = {}): LibraryBook {
  return {
    key: 'k',
    title: 'חידושים',
    type: 'docx',
    bookId: 'חידושים',
    bookUid: 'uid:3',
    id: null,
    source: 'user',
    author: '',
    ...patch,
  };
}

/** דחייה שנוצרת בזמן הקריאה ולא מראש — `mockRejectedValue` הפיל כאן את הבדיקה. */
const rejects = (error: Error) => async (): Promise<never> => {
  throw error;
};

const FILE = { token: 't', url: 'http://127.0.0.1/f/t', name: 'חידושים.docx', size: 10 };

beforeEach(() => callMock.mockReset());

describe('fetchLibraryShelf', () => {
  it('מבקש מאוצריא לגזום לסוגי Word, וגוזם גם בעצמו', async () => {
    callMock.mockResolvedValue({ categories: [{ title: 'א', books: [{ title: 'x', type: 'docx' }, { title: 'y', type: 'text' }] }] });
    const result = await fetchLibraryShelf();
    expect(callMock).toHaveBeenCalledWith('library.getTree', { includeBooks: true, types: ['docx', 'docm'] });
    expect(result.ok && result.value?.total).toBe(1);
  });

  it('אין הרשאה — reason ייעודי, לא כשל כללי', async () => {
    callMock.mockImplementationOnce(rejects(denied('library.getTree')));
    expect(await fetchLibraryShelf()).toMatchObject({ ok: false, reason: 'permission-denied' });
  });
});

describe('openLibraryBook', () => {
  it('ספר אישי מתבקש לכתיבה, לפי bookUid', async () => {
    callMock.mockResolvedValue({ ...FILE, access: 'readwrite' });
    const result = await openLibraryBook(book());
    expect(callMock).toHaveBeenCalledWith('library.openBookFile', { bookUid: 'uid:3', access: 'readwrite' });
    expect(result).toMatchObject({ ok: true, value: { token: 't', access: 'readwrite' } });
  });

  it('סירוב לכתיבה נופל לקריאה, והספר נפתח', async () => {
    callMock.mockImplementationOnce(rejects(denied('library.openBookFile'))).mockResolvedValueOnce({ ...FILE, access: 'read' });
    const result = await openLibraryBook(book());
    expect(callMock).toHaveBeenLastCalledWith('library.openBookFile', { bookUid: 'uid:3', access: 'read' });
    expect(result).toMatchObject({ ok: true, value: { access: 'read' } });
  });

  it('ספר של הספרייה מתבקש לקריאה בלבד, ובלי bookUid נופל ל-id ואז לכותרת+סוג', async () => {
    callMock.mockResolvedValue(FILE);
    await openLibraryBook(book({ source: 'library', bookUid: null, id: 7 }));
    expect(callMock).toHaveBeenLastCalledWith('library.openBookFile', { id: 7, access: 'read' });
    await openLibraryBook(book({ source: null, bookUid: null }));
    expect(callMock).toHaveBeenLastCalledWith('library.openBookFile', { bookId: 'חידושים', type: 'docx', access: 'read' });
  });

  it('אוצריא ישנה — unsupported', async () => {
    callMock.mockImplementationOnce(rejects(unknown('library.openBookFile')));
    expect(await openLibraryBook(book({ source: 'library' }))).toMatchObject({ ok: false, reason: 'unsupported' });
  });

  it('גישה שלא דווחה נגזרת ממה שהתבקש', async () => {
    callMock.mockResolvedValue(FILE);
    expect(await openLibraryBook(book({ source: 'library' }))).toMatchObject({ ok: true, value: { access: 'read' } });
  });
});

describe('התיקיות', () => {
  it('pickDocFolder: ביטול הוא value null, לא כשל', async () => {
    callMock.mockResolvedValue({ cancelled: true });
    expect(await pickDocFolder()).toEqual({ ok: true, value: null });
  });

  it('pickDocFolder: בחירה מחזירה token, שם ונתיב', async () => {
    callMock.mockResolvedValue({ cancelled: false, folderToken: 'ft', name: 'חידושים', path: 'C:\\חידושים' });
    expect(await pickDocFolder()).toMatchObject({ ok: true, value: { token: 'ft', name: 'חידושים', path: 'C:\\חידושים' } });
  });

  it('listDocFolder מסנן לסיומות Word, ותיקייה שנעלמה היא not-found', async () => {
    callMock.mockResolvedValueOnce({ entries: [], truncated: false });
    await listDocFolder('ft', 'שיעורים');
    expect(callMock).toHaveBeenCalledWith('fs.listUserFolder', { folderToken: 'ft', path: 'שיעורים', extensions: ['docx', 'docm'] });
    callMock.mockImplementationOnce(rejects(new HostCallError('gone', 'error.not_found', 'fs.listUserFolder')));
    expect(await listDocFolder('ft', '')).toMatchObject({ ok: false, reason: 'not-found' });
  });

  it('openDocFolderFile מבקש כתיבה, ונופל לקריאה בסירוב', async () => {
    callMock.mockImplementationOnce(rejects(denied('fs.openFolderFile'))).mockResolvedValueOnce(FILE);
    const result = await openDocFolderFile('ft', 'א.docx', 'א.docx');
    expect(callMock).toHaveBeenNthCalledWith(1, 'fs.openFolderFile', { folderToken: 'ft', path: 'א.docx', access: 'readwrite' });
    expect(result).toMatchObject({ ok: true, value: { access: 'read' } });
  });
});

describe('OPEN_SOURCE_PERMISSIONS', () => {
  it('מכסה כל מתודה שהמודול קורא לה', async () => {
    const { readFileSync } = await import('node:fs');
    const source = readFileSync('src/host/open-sources.ts', 'utf8');
    const called = [...source.matchAll(/(?:call|tryCall)<[^>]*>\(\s*'([a-zA-Z.]+)'|(?:call|tryCall)\('([a-zA-Z.]+)'/g)].map(
      (m) => m[1] ?? m[2],
    );
    expect(called.length).toBeGreaterThan(4);
    for (const method of called) expect(OPEN_SOURCE_PERMISSIONS, method).toHaveProperty([method!]);
  });
});
