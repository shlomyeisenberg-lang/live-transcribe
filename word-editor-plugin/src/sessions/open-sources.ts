/**
 * המקורות של „פתח מסמך” מעבר לאחרונים: ספרי ה-Word שבספריית אוצריא, והתיקיות
 * שהמשתמש הוסיף לתוסף.
 *
 * ## למה מודול טהור
 *
 * אותו נימוק כמו `recent-documents.ts`: כאן מוחלט מה מופיע בעץ ומה נעלם ממנו —
 * איזה ספר נחשב „Word”, איזו קטגוריה נגזמת, מה נמצא בחיפוש, ומה נשאר מרשימת
 * התיקיות אחרי הפעלה מחדש. החלטות כאלה נשברות בשקט, ולכן הן נמדדות ישירות בלי
 * לזייף את הגשר. הקריאות עצמן יושבות ב-`host/open-sources.ts`.
 *
 * ## מהירות
 *
 * העץ שחוזר מאוצריא נגזם כאן פעם אחת לצורה קטנה (ספרי Word בלבד, וקטגוריות
 * שיש מתחתן כאלה), ורק הצורה הזאת נשמרת במטמון ומוצגת. הדיאלוג נפתח מהמטמון
 * מיד, והרענון רץ ברקע — ראו `composables/use-open-sources.ts`.
 */
import { foldForSearch } from './recent-documents';

/**
 * סוגי הספרים שהעורך באמת פותח. `doc`/`rtf`/`odt` מופיעים בספרייה, אבל המנוע
 * קורא OOXML בלבד — ספר כזה בעץ היה שורה שנכשלת בלחיצה. זו אותה רשימה שבורר
 * הקבצים מסנן לפיה (`pickDocxFile` ב-host/files.ts).
 */
export const EDITABLE_BOOK_TYPES = ['docx', 'docm'] as const;

/** הסיומות שנשלחות לפירוט תיקייה — אותה רשימה, בצורת סיומת. */
export const EDITABLE_EXTENSIONS: readonly string[] = EDITABLE_BOOK_TYPES;

export interface LibraryBook {
  /** מפתח יציב לרינדור: `bookUid` כשיש, אחרת הכותרת והסוג. */
  key: string;
  title: string;
  type: string;
  bookId: string;
  bookUid: string | null;
  id: number | null;
  /** `'user'` = ספר אישי. `null` כשהגרסה של אוצריא אינה מדווחת. */
  source: string | null;
  author: string;
}

export interface LibraryShelf {
  title: string;
  path: string;
  shelves: LibraryShelf[];
  books: LibraryBook[];
  /** כמה ספרים יש מתחת למדף, כולל כל תתי-המדפים. */
  total: number;
}

function text(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

function normalizeBook(raw: unknown, types: ReadonlySet<string>): LibraryBook | null {
  if (!raw || typeof raw !== 'object') return null;
  const value = raw as Record<string, unknown>;
  const type = text(value.type).toLowerCase();
  if (!types.has(type)) return null;
  const title = text(value.title) || text(value.bookId);
  if (title === '') return null;
  const bookUid = text(value.bookUid) || null;
  const id = typeof value.id === 'number' && Number.isFinite(value.id) ? value.id : null;
  return {
    key: bookUid ?? `${title}\u0000${type}`,
    title,
    type,
    bookId: text(value.bookId) || title,
    bookUid,
    id,
    source: text(value.source) || null,
    author: text(value.author),
  };
}

/**
 * גוזם את העץ של `library.getTree` לספרי Word בלבד.
 *
 * מדף בלי ספר Word בשום מקום מתחתיו נעלם — עץ של 2,000 קטגוריות ריקות אינו
 * „ספרי Word בספרייה”. גרסת אוצריא שמכירה את `types` כבר מחזירה עץ גזום, וזה
 * no-op עליו; גרסה ישנה מתעלמת מהארגומנט, והגיזום כאן הוא מה שמציל את התצוגה.
 *
 * `null` = אין עץ (קלט פגום, או ספרייה בלי אף ספר Word).
 */
export function pruneLibraryTree(
  raw: unknown,
  types: readonly string[] = EDITABLE_BOOK_TYPES,
): LibraryShelf | null {
  const allowed = new Set(types.map((type) => type.toLowerCase()));

  const walk = (node: unknown, fallbackTitle: string): LibraryShelf | null => {
    if (!node || typeof node !== 'object') return null;
    const value = node as Record<string, unknown>;
    const books = (Array.isArray(value.books) ? value.books : [])
      .map((book) => normalizeBook(book, allowed))
      .filter((book): book is LibraryBook => book !== null)
      .sort((a, b) => a.title.localeCompare(b.title, 'he'));
    const shelves = (Array.isArray(value.categories) ? value.categories : [])
      .map((child) => walk(child, ''))
      .filter((shelf): shelf is LibraryShelf => shelf !== null);
    const total = books.length + shelves.reduce((sum, shelf) => sum + shelf.total, 0);
    if (total === 0) return null;
    const title = text(value.title) || fallbackTitle;
    return { title, path: text(value.path) || `/${title}`, shelves, books, total };
  };

  const root = walk(raw, 'ספריית אוצריא');
  if (!root) return null;
  return { ...root, path: '/' };
}

/** המדף בנתיב הנתון, או `null`. `'/'` הוא השורש. */
export function findShelf(root: LibraryShelf | null, path: string): LibraryShelf | null {
  if (!root) return null;
  if (path === '/' || path === root.path) return root;
  const stack = [...root.shelves];
  while (stack.length > 0) {
    const shelf = stack.pop()!;
    if (shelf.path === path) return shelf;
    stack.push(...shelf.shelves);
  }
  return null;
}

/**
 * שרשרת המדפים מהשורש ועד המדף בנתיב — שביל הלחם שמעל הרשימה. ריקה כשהנתיב
 * אינו בעץ (למשל אחרי רענון שבו הקטגוריה נעלמה).
 */
export function shelfTrail(root: LibraryShelf | null, path: string): LibraryShelf[] {
  if (!root) return [];
  const visit = (shelf: LibraryShelf, trail: LibraryShelf[]): LibraryShelf[] | null => {
    const next = [...trail, shelf];
    if (shelf.path === path || (path === '/' && shelf === root)) return next;
    for (const child of shelf.shelves) {
      const found = visit(child, next);
      if (found) return found;
    }
    return null;
  };
  return visit(root, []) ?? [];
}

/** כל הספרים האישיים בעץ, ממוינים לפי שם — הצומת „ספרים אישיים”. */
export function personalBooks(root: LibraryShelf | null): LibraryBook[] {
  if (!root) return [];
  const out: LibraryBook[] = [];
  const stack = [root];
  while (stack.length > 0) {
    const shelf = stack.pop()!;
    for (const book of shelf.books) if (book.source === 'user') out.push(book);
    stack.push(...shelf.shelves);
  }
  return out.sort((a, b) => a.title.localeCompare(b.title, 'he'));
}

/**
 * העץ בלי מדפים שכל מה שבהם אישי.
 *
 * אוצריא מציבה את הספרים האישיים במדף משלה בעץ הספרייה, והעמודה מציגה אותם
 * גם בצומת „ספרים אישיים” — כלומר אותו מדף היה מופיע פעמיים, פעם כתיקייה
 * ופעם כצומת, בשני שמות זהים זה מתחת לזה (נראה בצילום השער). הצומת הוא
 * המקום הנכון להם: הם הקבצים של המשתמש, לא חלק מהספרייה.
 *
 * מדף מעורב (אישיים ושל הספרייה יחד) נשאר כמות שהוא — הוצאת הספרים האישיים
 * ממנו הייתה מזיזה ספר ממקום שהמשתמש מכיר.
 */
export function withoutPersonalShelves(root: LibraryShelf | null): LibraryShelf | null {
  if (!root) return null;
  const onlyPersonal = (shelf: LibraryShelf): boolean =>
    shelf.books.every((book) => book.source === 'user') && shelf.shelves.every(onlyPersonal);
  const shelves = root.shelves.filter((shelf) => !onlyPersonal(shelf));
  if (shelves.length === root.shelves.length) return root;
  const total = root.books.length + shelves.reduce((sum, shelf) => sum + shelf.total, 0);
  return total === 0 ? null : { ...root, shelves, total };
}

export interface LibraryHit {
  book: LibraryBook;
  /** שמות המדפים מתחת לשורש, מופרדים ב-„ / ” — השורה השנייה של התוצאה. */
  where: string;
}

/**
 * חיפוש ספרים לפי כותרת ומחבר, בכל העץ. `limit` חוסם את מה שמרונדר: שאילתה של
 * אות אחת בספרייה גדולה אינה אמורה לצייר אלף שורות.
 *
 * תוצאה שהכותרת שלה **מתחילה** בשאילתה קודמת לתוצאה שמכילה אותה באמצע — מי
 * שהקליד „שו"ת” מחפש ספר ששמו מתחיל כך.
 */
export function searchLibrary(root: LibraryShelf | null, query: string, limit = 50): LibraryHit[] {
  const needle = foldForSearch(query);
  const cap = Math.max(0, Math.floor(limit));
  if (!root || needle === '' || cap === 0) return [];
  const starts: LibraryHit[] = [];
  const contains: LibraryHit[] = [];
  /**
   * מחזיקים רק את מה שיכול להיכנס לתוצאה הסופית. כל תוצאות ה-`starts`
   * מדורגות לפני `contains`, לכן כשהגענו לתקרה שלהן אפשר לעצור את הסריקה.
   */
  const visit = (shelf: LibraryShelf, trail: string[]): boolean => {
    for (const book of shelf.books) {
      const title = foldForSearch(book.title);
      const where = trail.join(' / ');
      if (title.startsWith(needle)) {
        starts.push({ book, where });
        if (starts.length === cap) return true;
      } else if (title.includes(needle) || foldForSearch(book.author).includes(needle)) {
        if (contains.length < cap) contains.push({ book, where });
      }
    }
    for (const child of shelf.shelves) {
      if (visit(child, [...trail, child.title])) return true;
    }
    return false;
  };
  visit(root, []);
  return starts.length === cap ? starts : [...starts, ...contains].slice(0, cap);
}

/* ------------------------------------------------------------------ */
/* מטמון העץ                                                            */
/* ------------------------------------------------------------------ */

/** גרסת צורת המטמון. שינוי בצורה = מספר חדש, והמטמון הישן נזרק ולא מפוענח. */
export const LIBRARY_CACHE_VERSION = 1;

export interface LibraryCache {
  version: number;
  savedAt: number;
  tree: LibraryShelf | null;
}

/**
 * מפענח את המטמון מ-storage. כל דבר שאינו בדיוק הצורה הנוכחית הוא `null` —
 * מטמון פגום אינו סיבה לצייר עץ שבור; הרענון ממילא בדרך.
 */
export function normalizeLibraryCache(raw: unknown): LibraryCache | null {
  if (!raw || typeof raw !== 'object') return null;
  const value = raw as Partial<LibraryCache>;
  if (value.version !== LIBRARY_CACHE_VERSION) return null;
  if (typeof value.savedAt !== 'number' || !Number.isFinite(value.savedAt)) return null;
  const tree = value.tree === null ? null : isShelf(value.tree) ? value.tree : undefined;
  if (tree === undefined) return null;
  return { version: LIBRARY_CACHE_VERSION, savedAt: value.savedAt, tree };
}

function isShelf(value: unknown): value is LibraryShelf {
  if (!value || typeof value !== 'object') return false;
  const shelf = value as Partial<LibraryShelf>;
  return (
    typeof shelf.title === 'string' &&
    typeof shelf.path === 'string' &&
    typeof shelf.total === 'number' &&
    Array.isArray(shelf.books) &&
    Array.isArray(shelf.shelves) &&
    shelf.shelves.every(isShelf)
  );
}

/* ------------------------------------------------------------------ */
/* התיקיות של המשתמש                                                    */
/* ------------------------------------------------------------------ */

export interface DocFolder {
  /** ה-token האטום של אוצריא. הנתיב עצמו אינו מפתח — ה-token הוא ההרשאה. */
  token: string;
  name: string;
  /** הנתיב המלא, לטולטיפ בלבד. */
  path: string;
  addedAt: number;
}

/** מה שנשמר ב-storage, אחרי השמטת שורות פגומות ו-token כפול. */
export function normalizeFolders(raw: unknown): DocFolder[] {
  if (!Array.isArray(raw)) return [];
  const seen = new Set<string>();
  const out: DocFolder[] = [];
  for (const item of raw) {
    if (!item || typeof item !== 'object') continue;
    const value = item as Partial<DocFolder>;
    const token = text(value.token);
    if (token === '' || seen.has(token)) continue;
    seen.add(token);
    const path = text(value.path);
    out.push({
      token,
      name: text(value.name) || baseName(path) || 'תיקייה',
      path,
      addedAt: typeof value.addedAt === 'number' && Number.isFinite(value.addedAt) ? value.addedAt : 0,
    });
  }
  return out;
}

/**
 * מוסיף תיקייה. אוצריא מחזירה את אותו token לתיקייה שכבר נבחרה, ולכן בחירה
 * חוזרת אינה יוצרת שורה שנייה — היא רק מרעננת את השם והנתיב.
 */
export function addFolder(list: readonly DocFolder[], folder: DocFolder): DocFolder[] {
  const existing = list.findIndex((item) => item.token === folder.token);
  if (existing >= 0) {
    const next = [...list];
    next[existing] = { ...list[existing]!, name: folder.name, path: folder.path };
    return next;
  }
  return [...list, folder];
}

export function removeFolder(list: readonly DocFolder[], token: string): DocFolder[] {
  return list.filter((item) => item.token !== token);
}

function baseName(path: string): string {
  const parts = path.split(/[\\/]/).filter(Boolean);
  return parts[parts.length - 1] ?? '';
}

/* ------------------------------------------------------------------ */
/* פירוט תיקייה                                                         */
/* ------------------------------------------------------------------ */

export interface FolderEntry {
  name: string;
  /** יחסי לשורש התיקייה, מופרד ב-`/`. */
  path: string;
  type: 'file' | 'dir';
  size: number;
  /** מילישניות, או `0` כשלא דווח. */
  modified: number;
}

export interface FolderListing {
  entries: FolderEntry[];
  truncated: boolean;
}

/**
 * מפענח את התשובה של `fs.listUserFolder`. תיקיות לפני קבצים, כל קבוצה לפי שם
 * — אוצריא כבר ממיינת כך, והמיון כאן הוא רשת שמחזיקה את התצוגה גם מול גרסה
 * שממיינת אחרת.
 */
export function normalizeListing(raw: unknown): FolderListing {
  const value = raw && typeof raw === 'object' ? (raw as Record<string, unknown>) : {};
  const entries: FolderEntry[] = [];
  for (const item of Array.isArray(value.entries) ? value.entries : []) {
    if (!item || typeof item !== 'object') continue;
    const entry = item as Record<string, unknown>;
    const name = text(entry.name);
    const path = text(entry.path);
    if (name === '' || path === '') continue;
    const type = entry.type === 'dir' ? 'dir' : entry.type === 'file' ? 'file' : null;
    if (!type) continue;
    const modified = typeof entry.modified === 'string' ? Date.parse(entry.modified) : Number.NaN;
    entries.push({
      name,
      path,
      type,
      size: typeof entry.size === 'number' && entry.size > 0 ? entry.size : 0,
      modified: Number.isFinite(modified) ? modified : 0,
    });
  }
  entries.sort((a, b) =>
    a.type === b.type ? a.name.localeCompare(b.name, 'he') : a.type === 'dir' ? -1 : 1,
  );
  return { entries, truncated: value.truncated === true };
}

/** המפתח של פירוט במטמון: התיקייה ותת-הנתיב. */
export function listingKey(token: string, path: string): string {
  return `${token}\u0000${path}`;
}

/** שם תת-התיקייה האחרונה בנתיב יחסי, לכותרת ולשביל הלחם. */
export function lastSegment(path: string): string {
  return baseName(path);
}

/** האב של נתיב יחסי. `''` הוא שורש התיקייה. */
export function parentPath(path: string): string {
  const parts = path.split('/').filter(Boolean);
  parts.pop();
  return parts.join('/');
}

/* ------------------------------------------------------------------ */
/* המקום בעץ                                                            */
/* ------------------------------------------------------------------ */

/** הצומת שנבחר בעץ — מה שהרשימה משמאל מציגה. */
export type OpenPlace =
  | { kind: 'recent' }
  | { kind: 'library'; path: string }
  | { kind: 'personal' }
  | { kind: 'folder'; token: string; path: string };

export const RECENT_PLACE: OpenPlace = { kind: 'recent' };

/** מפתח יציב — לזהות הצומת בעץ, להשוואה, ולשמירה בין הפעלות. */
export function placeKey(place: OpenPlace): string {
  switch (place.kind) {
    case 'recent':
    case 'personal':
      return place.kind;
    case 'library':
      return `library\u0000${place.path}`;
    case 'folder':
      return `folder\u0000${place.token}\u0000${place.path}`;
  }
}

/**
 * מפענח מקום שנשמר. מה שאינו מוכר — או תיקייה שכבר הוסרה — חוזר
 * ל„אחרונים”: מקום שאינו קיים יותר אינו אמור לפתוח את הדיאלוג על רשימה ריקה.
 */
export function parsePlace(raw: unknown, folders: readonly DocFolder[]): OpenPlace {
  if (typeof raw !== 'string') return RECENT_PLACE;
  const [kind, a, b] = raw.split('\u0000');
  if (kind === 'personal') return { kind: 'personal' };
  if (kind === 'library' && typeof a === 'string' && a.startsWith('/')) return { kind: 'library', path: a };
  if (kind === 'folder' && a && folders.some((folder) => folder.token === a)) {
    return { kind: 'folder', token: a, path: b ?? '' };
  }
  return RECENT_PLACE;
}
