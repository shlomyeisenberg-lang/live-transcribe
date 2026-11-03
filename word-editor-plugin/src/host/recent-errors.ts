/**
 * השגיאות האחרונות שהדף ראה, בשביל „דיווח על באג”.
 *
 * ## למה בכלל
 *
 * מי שמדווח „זה נתקע” אינו יודע מה קרה, ואין לו קונסול: התוסף רץ ב-WebView
 * בתוך אוצריא, בלי DevTools. מה שהדף הדפיס ל-`console.error` ומה שזרק בלי
 * שנתפס הוא בדרך כלל כל מה שיש למפתח ביד — ולכן הוא נאסף כאן, בזיכרון בלבד,
 * ומוצג למשתמש בתוך הדיווח **לפני** שהוא נשלח.
 *
 * ## מה נשמר, ומה לא
 *
 * הודעה בלבד, מקוצרת, ובלי ערכים שאינם טקסט: אובייקט שנמסר ל-`console.error`
 * נרשם כשם הסוג שלו ולא כתוכנו, כי תוכן כזה עשוי להיות צומת של המסמך. הודעות
 * עוברות `scrubDiagnostic` (נתיבי משתמש, כתובות מייל) כבר ברישום, כך שגם
 * דפוסים מוכרים מוסתרים; עדיין נדרשת בדיקת המשתמש לפני צירוף הטקסט.
 *
 * אין כאן שמירה לדיסק ואין שליחה: הרשימה מתה עם הדף.
 */

/** כמה שגיאות נשמרות. דיווח עם מאה שורות אינו נקרא, ומתארך מעבר לגבול ה-URL. */
export const RECENT_ERRORS_LIMIT = 8;

/** אורך מרבי להודעה אחת. */
const MESSAGE_LIMIT = 300;

export interface RecentError {
  /** `HH:MM:SS` בשעון המקומי — מספיק כדי לראות מה קדם למה. */
  time: string;
  message: string;
  /** כמה פעמים ברצף נרשמה בדיוק אותה הודעה. */
  count: number;
}

const entries: RecentError[] = [];
let installed = false;

/** Maximum input inspected per diagnostic, before any pattern processing. */
export const DIAGNOSTIC_INPUT_LIMIT = 4096;

/**
 * Best-effort redaction of recognizable paths, document filenames and email.
 * Arbitrary log strings can still contain private text: inclusion is opt-in and
 * the user must review the preview. Never read document content for diagnostics.
 */
export function scrubDiagnostic(text: string): string {
  const bounded = text.slice(0, DIAGNOSTIC_INPUT_LIMIT);
  const scrubbed = bounded
    .replace(/[^\s<>"'()]+/g, (token) => token.includes('@') ? '<מייל>' : token)
    // Absolute paths can contain spaces, including the home directory itself.
    // Consume conservatively through a known filename or the next delimiter;
    // removing extra prose is preferable to exporting a private directory name.
    .replace(/(?:file:\/\/\/|[A-Za-z]:[\\/]|\\\\|\/)(?:[^\r\n"'<>|?*()„“”«»]*?\.(?:docx|dotx|docm|doc|pdf|txt|[cm]?[jt]s)(?=[:\s"'<>|?*()„“”«»]|$)(?::\d+(?::\d+)?)?|[^\r\n"'<>|?*()„“”«»]*)/gi, (path) => {
      const document = path.match(/\.(docx|dotx|docm|doc|pdf|txt)(?=[:\s]|$)/i);
      if (document) return `<מסמך>.${document[1]}`;
      // Only keep technical script basenames, never user-chosen directories.
      const script = path.match(/(?:^|[\\/])([^\\/\s]+\.[cm]?[jt]s(?::\d+(?::\d+)?)?)$/i);
      return script ? `file:///…/${path.includes('/assets/') ? 'assets/' : ''}${script[1]}` : '<נתיב>';
    })
    .replace(/(["'„“«])[^"'„“”«»\n]*?\.(docx|dotx|docm|doc|pdf|txt)(["'”»])/gi, '$1<מסמך>.$2$3')
    // Match each token once, rather than restarting a greedy basename search
    // at every character of an extension-free diagnostic (quadratic work).
    .replace(/[^\\/\s'"«»„“”:<>]+/g, (token) =>
      token.startsWith('.') ? token : token.replace(/^.*\.(docx|dotx|docm|doc|pdf|txt)\b/i, (_name, ext: string) => `<מסמך>.${ext}`));
  return scrubbed + (text.length > DIAGNOSTIC_INPUT_LIMIT ? '…' : '');
}

function clip(text: string): string {
  const flat = text.replace(/\s+/g, ' ').trim();
  return flat.length > MESSAGE_LIMIT ? `${flat.slice(0, MESSAGE_LIMIT)}…` : flat;
}

/** ערך אחד מתוך הארגומנטים של `console.error` — טקסט, או שם הסוג בלבד. */
function describe(value: unknown): string {
  if (typeof value === 'string') return value;
  if (value instanceof Error) return `${value.name.slice(0, 100)}: ${value.message.slice(0, DIAGNOSTIC_INPUT_LIMIT)}`;
  if (typeof value === 'number' || typeof value === 'boolean' || value == null) return String(value);
  const name = (value as { constructor?: { name?: string } }).constructor?.name;
  return `[${name || typeof value}]`;
}

function timestamp(): string {
  const now = new Date();
  return [now.getHours(), now.getMinutes(), now.getSeconds()]
    .map((part) => String(part).padStart(2, '0'))
    .join(':');
}

/** רושמת שגיאה. מיוצאת לבדיקות ולמי שתופס שגיאה ורוצה שתופיע בדיווח. */
export function recordError(...parts: unknown[]): void {
  // Stop before mapping/joining arbitrary argument lists; descriptions and
  // aggregate text are bounded even for a multi-megabyte console message.
  let raw = '';
  for (let i = 0; i < Math.min(parts.length, 32) && raw.length < DIAGNOSTIC_INPUT_LIMIT; i += 1) {
    const part = describe(parts[i]).slice(0, DIAGNOSTIC_INPUT_LIMIT - raw.length);
    raw += (raw ? ' ' : '') + part;
  }
  const message = clip(scrubDiagnostic(raw));
  if (!message) return;

  const last = entries[entries.length - 1];
  if (last && last.message === message) {
    last.count += 1;
    last.time = timestamp();
    return;
  }
  entries.push({ time: timestamp(), message, count: 1 });
  if (entries.length > RECENT_ERRORS_LIMIT) entries.splice(0, entries.length - RECENT_ERRORS_LIMIT);
}

/** עותק של הרשימה, מהישנה לחדשה. */
export function recentErrors(): RecentError[] {
  return entries.map((entry) => ({ ...entry }));
}

/** לבדיקות. */
export function clearRecentErrors(): void {
  entries.length = 0;
}

/**
 * מתחילה להאזין: `console.error`, `error` ו-`unhandledrejection`.
 *
 * `console.error` נעטף ולא מוחלף — המקור נקרא תמיד, ראשון, וכשל ברישום אינו
 * מגיע אליו. נקראת פעם אחת מ-main.ts; קריאה חוזרת אינה עוטפת שוב.
 */
export function installRecentErrors(): void {
  if (installed) return;
  installed = true;

  const original = console.error.bind(console);
  console.error = (...args: unknown[]) => {
    original(...args);
    try {
      recordError(...args);
    } catch {
      // הרישום הוא תוספת; הוא אינו רשאי להפיל את מי שדיווח.
    }
  };

  window.addEventListener('error', (event: ErrorEvent) => {
    // שגיאת טעינה של משאב (תמונה, סקריפט) מגיעה בלי message — אין מה לרשום.
    if (!event.message) return;
    const where = event.filename ? ` (${event.filename.slice(0, DIAGNOSTIC_INPUT_LIMIT)}:${event.lineno}:${event.colno})` : '';
    recordError(event.message, where);
  });

  window.addEventListener('unhandledrejection', (event: PromiseRejectionEvent) => {
    recordError('Unhandled rejection:', event.reason);
  });
}
