/**
 * טעינת מילון ראשי-התיבות כנכס נפרד, בהזרקת `<script>` — אותה תבנית בדיוק
 * כמו `spellcheck-dictionary.ts` (ר' שם ההסבר המלא: IIFE יחיד עם
 * `inlineDynamicImports`, ו-`file://` שחוסם `fetch`). הנכס נטען רק כשיש
 * בפועל ראשי-תיבות שלמים להציע להם פירוש, ולא בעליית התוסף.
 *
 * ## שני הבדלים מהמילון התורני, ושניהם מכוונים
 *
 * **הנכס הוא אובייקט ולא מחרוזת.** `acronyms-asset` מציב ליטרל אובייקט, ולכן
 * `packedFromWindow` מחפש אובייקט. הגרסה הקודמת בדקה `typeof === 'string'`,
 * לא מצאה, והחזירה `null` בכל טעינה — ההשלמה לא עבדה מעולם.
 *
 * **כשל כן נזכר, לזמן מוגבל.** במילון התורני הכשל אינו נזכר כי הטעינה נובעת
 * מהדלקת מתג בידי המשתמש. כאן היא נובעת מהקלדה, ולכן כשל שאינו נזכר היה
 * מזריק את הנכס שוב ושוב בכל השהיה. במקומו — המתנה שמכפילה את עצמה.
 */
import { ACRONYMS_FILE, ACRONYMS_GLOBAL } from './acronyms-constants';
import { createAcronymDictionary, type AcronymDictionary, type PackedAcronyms } from './acronyms';

const DICTIONARY_SRC = `./${ACRONYMS_FILE}`;
const LOAD_TIMEOUT_MS = 20_000;

/** ההמתנה אחרי כשל, ותקרתה. נכס שאינו קיים אינו נבדק כל חצי דקה עד סוף ההפעלה. */
const FIRST_RETRY_DELAY_MS = 30_000;
const MAX_RETRY_DELAY_MS = 10 * 60_000;

type PackedLoader = () => Promise<PackedAcronyms | null>;

function packedFromWindow(): PackedAcronyms | null {
  const value = (globalThis as Record<string, unknown>)[ACRONYMS_GLOBAL];
  return typeof value === 'object' && value !== null && !Array.isArray(value)
    ? (value as PackedAcronyms)
    : null;
}

function injectAcronymsScript(): Promise<PackedAcronyms | null> {
  const ready = packedFromWindow();
  if (ready !== null) return Promise.resolve(ready);

  if (typeof document === 'undefined' || typeof document.createElement !== 'function') {
    return Promise.resolve(null);
  }

  return new Promise((resolve) => {
    let settled = false;
    const finish = (value: PackedAcronyms | null): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      // התגית סיימה את תפקידה — האובייקט כבר ב-`window`. בלי ההסרה כל ניסיון
      // חוזר היה משאיר אחריו תגית נוספת ב-`<head>`.
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

let pending: Promise<AcronymDictionary | null> | null = null;
let loaded: AcronymDictionary | null = null;
let retryAfter = 0;
let retryDelay = FIRST_RETRY_DELAY_MS;

/** המילון, בטעינה עצלה וחד-פעמית. כשל מחזיר `null` ומשתיק ניסיונות להמתנה הבאה. */
export function loadAcronymDictionary(
  loader: PackedLoader = injectAcronymsScript,
): Promise<AcronymDictionary | null> {
  if (loaded) return Promise.resolve(loaded);
  if (pending) return pending;
  if (Date.now() < retryAfter) return Promise.resolve(null);

  pending = (async () => {
    let packed: PackedAcronyms | null = null;
    try {
      packed = await loader();
    } catch {
      packed = null;
    }

    if (packed === null) {
      retryAfter = Date.now() + retryDelay;
      retryDelay = Math.min(retryDelay * 2, MAX_RETRY_DELAY_MS);
      return null;
    }

    retryAfter = 0;
    retryDelay = FIRST_RETRY_DELAY_MS;
    loaded = createAcronymDictionary(packed);
    return loaded;
  })().finally(() => {
    pending = null;
  });

  return pending;
}

/** לבדיקות בלבד. */
export function resetAcronymDictionary(): void {
  loaded = null;
  pending = null;
  retryAfter = 0;
  retryDelay = FIRST_RETRY_DELAY_MS;
}
