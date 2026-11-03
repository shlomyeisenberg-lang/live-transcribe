/**
 * טעינת רשימות ההשלמה הסטטיות כנכס נפרד, בהזרקת `<script>` — אותה תבנית
 * בדיוק כמו `acronym-dictionary.ts` (ר' שם ההסבר: IIFE יחיד עם
 * `inlineDynamicImports`, ו-`file://` שחוסם `fetch`).
 *
 * הנכס נטען רק אחרי שההשלמה מהספר לא מצאה התאמה — כלומר רק אצל מי שהדליק את
 * „השלמה מהספר” (כבוי כברירת מחדל) ובאמת הקליד. בשונה מראשי-התיבות אין כאן
 * שער צורה: כל מילה עברית יכולה לפתוח ביטוי או שם, ולכן אין מבחן זול שיפסול
 * מראש. מה שמצדיק את זה הוא הגודל — 32KB, מול 0.7MB של ראשי-התיבות.
 */
import { STATIC_COMPLETION_FILE, STATIC_COMPLETION_GLOBAL } from './static-completion-constants';
import { buildStaticIndex, type StaticIndex } from './static-completion';

const DICTIONARY_SRC = `./${STATIC_COMPLETION_FILE}`;
const LOAD_TIMEOUT_MS = 20_000;

/** ההמתנה אחרי כשל, ותקרתה — כמו בראשי-התיבות, ומאותה סיבה. */
const FIRST_RETRY_DELAY_MS = 30_000;
const MAX_RETRY_DELAY_MS = 10 * 60_000;

/** מה שהנכס מציב ב-`window`. הסדר הוא סדר הניסיון. */
export interface PackedStaticCompletion {
  readonly phrases: readonly string[];
  readonly authors: readonly string[];
}

type PackedLoader = () => Promise<PackedStaticCompletion | null>;

function packedFromWindow(): PackedStaticCompletion | null {
  const value = (globalThis as Record<string, unknown>)[STATIC_COMPLETION_GLOBAL];
  if (typeof value !== 'object' || value === null) return null;
  const candidate = value as Partial<PackedStaticCompletion>;
  return Array.isArray(candidate.phrases) && Array.isArray(candidate.authors)
    ? (value as PackedStaticCompletion)
    : null;
}

function injectStaticCompletionScript(): Promise<PackedStaticCompletion | null> {
  const ready = packedFromWindow();
  if (ready !== null) return Promise.resolve(ready);

  if (typeof document === 'undefined' || typeof document.createElement !== 'function') {
    return Promise.resolve(null);
  }

  return new Promise((resolve) => {
    let settled = false;
    const finish = (value: PackedStaticCompletion | null): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      script.remove();
      resolve(value);
    };

    const timer = setTimeout(() => finish(null), LOAD_TIMEOUT_MS);

    const script = document.createElement('script');
    script.async = false;
    script.src = DICTIONARY_SRC;
    script.addEventListener('load', () => finish(packedFromWindow()));
    script.addEventListener('error', () => finish(null));
    document.head.appendChild(script);
  });
}

let pending: Promise<StaticIndex[] | null> | null = null;
let loaded: StaticIndex[] | null = null;
let retryAfter = 0;
let retryDelay = FIRST_RETRY_DELAY_MS;

/** המקורות, בטעינה עצלה וחד-פעמית. כשל משתיק ניסיונות עד תום ההמתנה. */
export function loadStaticSources(
  loader: PackedLoader = injectStaticCompletionScript,
): Promise<StaticIndex[] | null> {
  if (loaded) return Promise.resolve(loaded);
  if (pending) return pending;
  if (Date.now() < retryAfter) return Promise.resolve(null);

  pending = (async () => {
    /*
     * ה-`try` עוטף גם את בניית האינדקס, ולא רק את הטעינה: `buildStaticIndex`
     * זורק על רשומה מעל 31 מילים, וזריקה משם הייתה דוחה את ההבטחה בלי לקבוע
     * `retryAfter` — כלומר unhandled rejection בכל הקשה (‏`void evaluate()`),
     * ובלי שראשי-התיבות שאחריה ייבדקו בכלל.
     */
    let indexes: StaticIndex[] | null = null;
    try {
      const packed = await loader();
      if (packed !== null) {
        indexes = [
          buildStaticIndex('talmudic-phrases', packed.phrases),
          buildStaticIndex('authors', packed.authors),
        ];
      }
    } catch {
      indexes = null;
    }

    if (indexes === null) {
      retryAfter = Date.now() + retryDelay;
      retryDelay = Math.min(retryDelay * 2, MAX_RETRY_DELAY_MS);
      return null;
    }

    retryAfter = 0;
    retryDelay = FIRST_RETRY_DELAY_MS;
    loaded = indexes;
    return loaded;
  })().finally(() => {
    pending = null;
  });

  return pending;
}

/** לבדיקות בלבד. */
export function resetStaticSources(): void {
  loaded = null;
  pending = null;
  retryAfter = 0;
  retryDelay = FIRST_RETRY_DELAY_MS;
}
