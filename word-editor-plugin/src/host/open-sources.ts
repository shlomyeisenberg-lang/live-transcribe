/**
 * הקריאות לאוצריא שמאחורי העץ של „פתח מסמך”: ספרי Word מהספרייה, ותיקיות
 * שהמשתמש הוסיף לתוסף.
 *
 * `fs.pickUserFolder`, `fs.listUserFolder`, `fs.openFolderFile`,
 * `fs.revokeFolder` ו-`library.openBookFile` נוספו לאוצריא בשביל העץ הזה.
 * גרסה שאינה מכירה אותן עונה `unknown_method`, והתשובה מתורגמת כאן ל-
 * `reason: 'unsupported'` — כדי שהדיאלוג יסתיר את הענף במקום להציג שגיאה על
 * כל לחיצה. הפירוש של מה שחוזר (גיזום, מיון, מטמון) יושב ב-
 * sessions/open-sources.ts.
 *
 * התוצאות מטופסות ולא זורקות, מאותו טעם כמו `ReaderResult` ב-otzaria-reader.ts:
 * מי שקורא צריך הודעה אחת בעברית ולא מסלול טיפול לכל קוד.
 */
import { call, hostErrorCode, isPermissionDenied, tryCall } from './otzaria-client';
import type { UserFile } from './files';
import {
  EDITABLE_BOOK_TYPES,
  EDITABLE_EXTENSIONS,
  normalizeListing,
  pruneLibraryTree,
  type DocFolder,
  type FolderListing,
  type LibraryBook,
  type LibraryShelf,
} from '../sessions/open-sources';

/**
 * ההרשאה שכל קריאה כאן דורשת, לפי ה-API_REFERENCE של אוצריא. המקור לבדיקת
 * המניפסט (tests/unit/manifest.test.ts) — אותו תפקיד כמו `READER_PERMISSIONS`.
 */
export const OPEN_SOURCE_PERMISSIONS: Record<string, string> = {
  'library.getTree': 'library.books.read',
  'library.openBookFile': 'library.content.read',
  'fs.pickUserFolder': 'fs.user_files.read',
  'fs.listUserFolder': 'fs.user_files.read',
  'fs.openFolderFile': 'fs.user_files.read',
  'fs.revokeFolder': 'fs.user_files.read',
};

export type SourceResult<T> =
  | { ok: true; value: T }
  | { ok: false; reason: 'unsupported' | 'permission-denied' | 'not-found' | 'forbidden' | 'failed'; message: string };

/** האם הכשל הוא „אוצריא הזאת אינה מכירה את הקריאה”. */
export function isUnknownMethod(error: unknown): boolean {
  const code = hostErrorCode(error);
  if (code) return code.endsWith('unknown_method') || code.endsWith('unsupported_method');
  return error instanceof Error && /unknown_method|unknown method/i.test(error.message);
}

function failure<T>(action: string, error: unknown): SourceResult<T> {
  if (isUnknownMethod(error)) {
    return { ok: false, reason: 'unsupported', message: `${action}: גרסת אוצריא הזאת אינה תומכת בזה עדיין — נדרש עדכון` };
  }
  if (isPermissionDenied(error)) {
    return { ok: false, reason: 'permission-denied', message: `${action}: לתוסף חסרה הרשאה. יש לאשר אותה בהגדרות אוצריא` };
  }
  const code = hostErrorCode(error) ?? '';
  const detail = error instanceof Error ? error.message : String(error);
  if (code.endsWith('not_found') || /not_found/.test(detail)) {
    return { ok: false, reason: 'not-found', message: `${action}: לא נמצא — ייתכן שהוזז או נמחק` };
  }
  if (code.endsWith('forbidden') || /forbidden/.test(detail)) {
    return { ok: false, reason: 'forbidden', message: `${action}: אוצריא אינה מתירה את המיקום הזה` };
  }
  return { ok: false, reason: 'failed', message: `${action}: ${detail}` };
}

/* ------------------------------------------------------------------ */
/* הספרייה                                                              */
/* ------------------------------------------------------------------ */

/**
 * ספרי ה-Word שבספרייה, כעץ גזום. `value: null` = אין בספרייה אף ספר Word,
 * וזו תשובה ולא כשל.
 *
 * `types` מבקש מאוצריא לגזום בצד שלה, כדי שרק העץ הקטן יעבור בגשר. גרסה ישנה
 * מתעלמת ממנו ומחזירה את הכול — `pruneLibraryTree` גוזם בכל מקרה.
 */
export async function fetchLibraryShelf(): Promise<SourceResult<LibraryShelf | null>> {
  try {
    const raw = await call<unknown>('library.getTree', {
      includeBooks: true,
      types: [...EDITABLE_BOOK_TYPES],
    });
    return { ok: true, value: pruneLibraryTree(raw) };
  } catch (error) {
    return failure('קריאת הספרייה נכשלה', error);
  }
}

interface OpenedFile {
  token?: string;
  url?: string;
  name?: string;
  size?: number;
  access?: 'read' | 'readwrite';
}

function toUserFile(res: OpenedFile | null, fallbackName: string, mode: 'read' | 'readwrite'): UserFile | null {
  if (!res?.token || !res.url) return null;
  return {
    token: res.token,
    url: res.url,
    name: res.name ?? fallbackName,
    size: res.size ?? 0,
    access: res.access ?? mode,
  };
}

/**
 * פותח ספר מהספרייה. ספר אישי מתבקש לכתיבה — הוא קובץ של המשתמש, ו„שמור”
 * עליו אמור לכתוב אליו. ספר של הספרייה נפתח לקריאה, ו„שמור” הראשון עובר ב-
 * „שמור בשם”: קובץ שהספרייה מעדכנת אינו יעד כתיבה.
 *
 * סירוב הרשאה לכתיבה נופל לקריאה, בדיוק כמו `pickDocxFile`: עדיף ספר שנפתח
 * ואינו נשמר במקומו מאשר ספר שלא נפתח.
 */
export async function openLibraryBook(book: LibraryBook): Promise<SourceResult<UserFile>> {
  const identity = book.bookUid
    ? { bookUid: book.bookUid }
    : book.id !== null
      ? { id: book.id }
      : { bookId: book.bookId, type: book.type };
  const fileName = `${book.title}.${book.type}`;

  const request = async (mode: 'read' | 'readwrite'): Promise<UserFile | null> =>
    toUserFile(await call<OpenedFile>('library.openBookFile', { ...identity, access: mode }), fileName, mode);

  const wantsWrite = book.source === 'user';
  try {
    const file = await request(wantsWrite ? 'readwrite' : 'read');
    return file ? { ok: true, value: file } : { ok: false, reason: 'failed', message: 'אוצריא לא החזירה את קובץ הספר' };
  } catch (error) {
    if (wantsWrite && isPermissionDenied(error)) {
      try {
        const file = await request('read');
        if (file) return { ok: true, value: file };
      } catch (retry) {
        return failure(`פתיחת „${book.title}” נכשלה`, retry);
      }
    }
    return failure(`פתיחת „${book.title}” נכשלה`, error);
  }
}

/* ------------------------------------------------------------------ */
/* התיקיות                                                              */
/* ------------------------------------------------------------------ */

/** בורר התיקיות של אוצריא. `value: null` = המשתמש ביטל. */
export async function pickDocFolder(): Promise<SourceResult<DocFolder | null>> {
  try {
    const res = await call<{ cancelled?: boolean; folderToken?: string; name?: string; path?: string }>(
      'fs.pickUserFolder',
      { title: 'הוספת תיקייה ל„פתח מסמך”' },
    );
    if (!res || res.cancelled || !res.folderToken) return { ok: true, value: null };
    return {
      ok: true,
      value: { token: res.folderToken, name: res.name || 'תיקייה', path: res.path ?? '', addedAt: Date.now() },
    };
  } catch (error) {
    return failure('הוספת התיקייה נכשלה', error);
  }
}

/** תוכן תיקייה (לא רקורסיבי): תת-תיקיות, וקובצי Word בלבד. */
export async function listDocFolder(token: string, path: string): Promise<SourceResult<FolderListing>> {
  try {
    const raw = await call<unknown>('fs.listUserFolder', {
      folderToken: token,
      path,
      extensions: [...EDITABLE_EXTENSIONS],
    });
    return { ok: true, value: normalizeListing(raw) };
  } catch (error) {
    return failure('קריאת התיקייה נכשלה', error);
  }
}

/**
 * פותח קובץ מתוך תיקייה. `readwrite` תמיד: התיקייה נבחרה בידי המשתמש כמקום
 * שהוא עובד בו, ו„שמור” אמור לכתוב לקובץ עצמו. סירוב הרשאה נופל לקריאה.
 */
export async function openDocFolderFile(token: string, path: string, name: string): Promise<SourceResult<UserFile>> {
  const request = async (mode: 'read' | 'readwrite'): Promise<UserFile | null> =>
    toUserFile(await call<OpenedFile>('fs.openFolderFile', { folderToken: token, path, access: mode }), name, mode);
  try {
    const file = await request('readwrite');
    return file ? { ok: true, value: file } : { ok: false, reason: 'failed', message: 'אוצריא לא החזירה את הקובץ' };
  } catch (error) {
    if (isPermissionDenied(error)) {
      try {
        const file = await request('read');
        if (file) return { ok: true, value: file };
      } catch (retry) {
        return failure(`פתיחת „${name}” נכשלה`, retry);
      }
    }
    return failure(`פתיחת „${name}” נכשלה`, error);
  }
}

/** מסיר את ההרשאה על התיקייה. ניקוי — אינו זורק ואינו מדווח. */
export async function revokeDocFolder(token: string): Promise<void> {
  await tryCall('fs.revokeFolder', { folderToken: token });
}
