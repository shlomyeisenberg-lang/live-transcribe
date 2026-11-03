/**
 * המצב של העץ ב„פתח מסמך”: ספרי ה-Word מהספרייה, התיקיות, פירוטי התיקיות,
 * והמקום שנבחר.
 *
 * ## מהירות — מה המשתמש מחכה לו, ומה לא
 *
 * שום דבר כאן אינו חוסם את פתיחת הדיאלוג:
 *
 * - **הספרייה** נטענת מהמטמון ב-`init()`, עוד לפני שהדיאלוג נפתח לראשונה.
 *   הרענון מאוצריא רץ ברקע פעם אחת בהפעלה (`refreshLibrary`), ומחליף את העץ
 *   כשהוא מגיע. „טוען…” מוצג רק כשאין מטמון בכלל — כלומר בפעם הראשונה.
 * - **פירוט תיקייה** נשמר בזיכרון לכל ההפעלה. חזרה לתיקייה מציגה מיד את מה
 *   שנראה בפעם הקודמת, ורענון שקט מעדכן אותו (stale-while-revalidate). שתי
 *   בקשות לאותה תיקייה בו-זמנית הן בקשה אחת.
 * - **המקום** שנבחר נזכר בין הפעלות, כדי שמי שעובד מתוך תיקייה אחת לא יצטרך
 *   לנווט אליה בכל פתיחה.
 */
import { ref, shallowReactive, shallowRef } from 'vue';
import {
  fetchLibraryShelf,
  listDocFolder,
  pickDocFolder,
  revokeDocFolder,
  type SourceResult,
} from '../host/open-sources';
import {
  loadDocFolders,
  loadLibraryCache,
  loadOpenDialogPlace,
  saveDocFolders,
  saveLibraryCache,
  saveOpenDialogPlace,
} from '../host/settings';
import {
  LIBRARY_CACHE_VERSION,
  RECENT_PLACE,
  addFolder,
  listingKey,
  normalizeFolders,
  normalizeLibraryCache,
  parsePlace,
  placeKey,
  removeFolder,
  type DocFolder,
  type FolderListing,
  type LibraryShelf,
  type OpenPlace,
} from '../sessions/open-sources';

/**
 * מצב הספרייה כפי שהדיאלוג מציג אותו. `unsupported` ו-`denied` מסתירים את
 * הענף או מסבירים אותו — הם אינם שגיאה שחוזרת בכל פתיחה.
 */
export type LibraryState = 'idle' | 'loading' | 'ready' | 'unsupported' | 'denied' | 'error';

export type ListingState =
  | { state: 'loading' }
  | {
      state: 'ready';
      listing: FolderListing;
      refreshing: boolean;
      /** כשל ברענון שקט; הנתונים הקודמים נשארים זמינים. */
      refreshError?: { reason: string; message: string };
    }
  | { state: 'error'; reason: string; message: string };

export function useOpenSources() {
  const library = shallowRef<LibraryShelf | null>(null);
  const libraryState = ref<LibraryState>('idle');
  const libraryMessage = ref('');
  const folders = ref<DocFolder[]>([]);
  /** האם אוצריא תומכת בתיקיות. `null` = טרם נבדק. */
  const foldersSupported = ref<boolean | null>(null);
  const listings = shallowReactive(new Map<string, ListingState>());
  const place = shallowRef<OpenPlace>(RECENT_PLACE);

  let initialized: Promise<void> | null = null;
  let libraryRefresh: Promise<void> | null = null;
  const inflight = new Map<string, Promise<void>>();

  /** טעינת מה שנשמר. אידמפוטנטי — כל קורא מקבל את אותה הבטחה. */
  function init(): Promise<void> {
    initialized ??= (async () => {
      const [rawFolders, rawCache, rawPlace] = await Promise.all([
        loadDocFolders(),
        loadLibraryCache(),
        loadOpenDialogPlace(),
      ]);
      folders.value = normalizeFolders(rawFolders);
      const cache = normalizeLibraryCache(rawCache);
      if (cache && libraryState.value === 'idle') {
        library.value = cache.tree;
        libraryState.value = 'ready';
      }
      place.value = parsePlace(rawPlace, folders.value);
    })();
    return initialized;
  }

  /**
   * מרענן את הספרייה מאוצריא. פעם אחת בהפעלה, אלא אם `force` — הספרייה
   * אינה משתנה בין שתי פתיחות של הדיאלוג, ובקשה בכל פתיחה הייתה עבודה לשווא.
   */
  function refreshLibrary(force = false): Promise<void> {
    if (libraryRefresh && !force) return libraryRefresh;
    libraryRefresh = (async () => {
      await init();
      if (libraryState.value !== 'ready') libraryState.value = 'loading';
      const result = await fetchLibraryShelf();
      if (result.ok) {
        library.value = result.value;
        libraryState.value = 'ready';
        libraryMessage.value = '';
        void saveLibraryCache({ version: LIBRARY_CACHE_VERSION, savedAt: Date.now(), tree: result.value });
        return;
      }
      libraryMessage.value = result.message;
      // עץ מהמטמון נשאר על המסך כשהרענון נכשל: עדיף רשימה מאתמול מאשר
      // הודעת שגיאה במקום רשימה שעובדת. רק „לא נתמך” ו„אין הרשאה” מחליפים
      // אותו, כי הם אומרים שגם העץ הישן אינו ניתן לפתיחה.
      if (result.reason === 'unsupported') libraryState.value = 'unsupported';
      else if (result.reason === 'permission-denied') libraryState.value = 'denied';
      else if (libraryState.value !== 'ready') libraryState.value = 'error';
    })();
    return libraryRefresh;
  }

  function listingOf(token: string, path: string): ListingState | undefined {
    return listings.get(listingKey(token, path));
  }

  /**
   * פירוט תיקייה. מה שכבר בזיכרון מוצג מיד ומתרענן בשקט; מה שלא — „טוען”.
   */
  function listFolder(token: string, path: string): Promise<void> {
    const key = listingKey(token, path);
    const running = inflight.get(key);
    if (running) return running;

    const previous = listings.get(key);
    if (previous?.state === 'ready') listings.set(key, { ...previous, refreshing: true, refreshError: undefined });
    else listings.set(key, { state: 'loading' });

    const task = (async () => {
      const result = await listDocFolder(token, path);
      if (result.ok) {
        foldersSupported.value = true;
        listings.set(key, { state: 'ready', listing: result.value, refreshing: false });
        return;
      }
      if (result.reason === 'unsupported') foldersSupported.value = false;
      if (previous?.state === 'ready') {
        // רענון נכשל אינו מוחק תוכן שכבר הוצג. משאירים אותו זמין, ומסמנים
        // שהמידע עלול להיות ישן כדי שהרשימה תוכל להציג הודעה ורענון חוזר.
        listings.set(key, {
          ...previous,
          refreshing: false,
          refreshError: { reason: result.reason, message: result.message },
        });
      } else {
        listings.set(key, { state: 'error', reason: result.reason, message: result.message });
      }
    })().finally(() => inflight.delete(key));
    inflight.set(key, task);
    return task;
  }

  /**
   * פירוט שנדרש לעץ (פתיחת צומת): נטען רק אם אינו בזיכרון. פתיחה וסגירה
   * חוזרות של צומת אינן בקשה חדשה בכל פעם — הרענון הוא של התיקייה שנבחרה.
   */
  function ensureListing(token: string, path: string): void {
    if (!listings.has(listingKey(token, path))) void listFolder(token, path);
  }

  function setPlace(next: OpenPlace): void {
    if (placeKey(next) === placeKey(place.value)) return;
    place.value = next;
    void saveOpenDialogPlace(placeKey(next));
    if (next.kind === 'folder') void listFolder(next.token, next.path);
  }

  /**
   * „הוסף תיקייה”. התיקייה החדשה נבחרת מיד — מי שהוסיף תיקייה רוצה לראות
   * מה יש בה, לא לחפש אותה בעץ.
   */
  async function addDocFolder(): Promise<SourceResult<DocFolder | null>> {
    await init();
    const result = await pickDocFolder();
    if (!result.ok) {
      if (result.reason === 'unsupported') foldersSupported.value = false;
      return result;
    }
    if (!result.value) return result;
    foldersSupported.value = true;
    folders.value = addFolder(folders.value, result.value);
    void saveDocFolders(folders.value);
    setPlace({ kind: 'folder', token: result.value.token, path: '' });
    return result;
  }

  function removeDocFolder(token: string): void {
    folders.value = removeFolder(folders.value, token);
    void saveDocFolders(folders.value);
    for (const key of [...listings.keys()]) if (key.startsWith(`${token}\u0000`)) listings.delete(key);
    if (place.value.kind === 'folder' && place.value.token === token) setPlace(RECENT_PLACE);
    void revokeDocFolder(token);
  }

  return {
    library,
    libraryState,
    libraryMessage,
    folders,
    foldersSupported,
    listings,
    place,
    init,
    refreshLibrary,
    listingOf,
    listFolder,
    ensureListing,
    setPlace,
    addDocFolder,
    removeDocFolder,
  };
}

export type OpenSources = ReturnType<typeof useOpenSources>;
