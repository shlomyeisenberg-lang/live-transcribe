/**
 * stub לפיתוח בדפדפן רגיל, כשאין window.Otzaria.
 * נטען רק ב-import.meta.env.DEV ולכן אינו נכנס ל-build.
 *
 * הוא משגר plugin.boot כ-CustomEvent על window — בדיוק כמו אוצריא — ולא
 * קורא ל-callbacks ישירות, כדי שה-latch שב-otzaria-client ייבדק בפיתוח באותו
 * מסלול שבו הוא עובד בייצור.
 *
 * מה שהוא **כן** מדמה: בורר קבצים אמיתי, מסלול הכתיבה המלא (begin → PUT →
 * commit) עם הורדה אמיתית ב„שמור בשם”, storage, ודיאלוגי אישור. כך אפשר לעבור
 * את כל הזרימה בדפדפן — פתיחה, עריכה, שמירה, שמירה בשם ופתיחה חוזרת — בלי
 * אוצריא. מה שהוא **אינו** מדמה: כתיבה אטומית לדיסק, הרשאות, ואת ה-WebView.
 * לאלה יש שער Windows.
 */
import type { BootPayload, OtzariaGlobal } from '../types/otzaria_plugin';
import { blankDocumentSource } from '../engine/blank-document-source';

const BOOT: BootPayload = {
  plugin: { id: 'dev', version: '0.0.0' },
  app: {
    version: '0.9.97',
    platform: 'dev',
    locale: 'he-IL',
    language: 'he',
    textDirection: 'rtl',
    devMode: true,
    runMode: 'foreground',
  },
  theme: {
    mode: 'light',
    colorScheme: {
      primary: '#1565C0',
      onPrimary: '#ffffff',
      secondary: '#6750A4',
      onSecondary: '#ffffff',
      surface: '#f8f9fa',
      onSurface: '#1a1a2e',
      onSurfaceVariant: '#49454f',
      surfaceContainerHigh: '#ece6f0',
      surfaceContainerHighest: '#e0e0e0',
      outline: '#cbd5e1',
      error: '#b00020',
      onError: '#ffffff',
    },
    typography: {
      fontFamily: 'FrankRuhlCLM',
      fontSize: 18,
      lineHeight: 1.5,
      commentatorsFontFamily: 'Shofar',
      commentatorsFontSize: 14,
    },
  },
  connectivity: { isOfflineMode: false, hasNetwork: false, isOnline: false },
  // `app.info.read` היא הרשאת בסיס שאוצריא מעניקה לכל תוסף. בלעדיה
  // explainHostGap מאבחן בטעות שהמארח חסם את fonts.listInstalled.
  permissions: ['app.info.read', 'fs.user_files.read', 'fs.user_files.write'],
};

/** קבצים שה„משתמש” בחר בריצה הזאת. נעלמים ברענון — כמו grant אמיתי שאבד. */
const files = new Map<string, { name: string; url: string; writable: boolean }>();
/** העלאות שממתינות ל-commit. */
const uploads = new Map<string, Blob>();
const UPLOAD_PREFIX = 'https://dev-stub.invalid/w/';
let counter = 0;

const storagePrefix = 'otzaria-word-dev:';
const workspacePrefix = 'otzaria-word-dev-fs:';

function pickFile(): Promise<{ file: File | null }> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.docx';
    input.addEventListener('change', () => resolve({ file: input.files?.[0] ?? null }), {
      once: true,
    });
    input.addEventListener('cancel', () => resolve({ file: null }), { once: true });
    input.click();
  });
}

function download(blob: Blob, name: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = name;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
}

/** ספרייה מדומה: שני מדפים מקוננים, ספר אישי אחד, וספר שאינו Word (נגזם). */
const DEV_LIBRARY = {
  title: 'ספריית אוצריא',
  path: '/',
  books: [],
  categories: [
    {
      title: 'הלכה',
      path: '/הלכה',
      books: [{ bookId: 'קונטרס בהלכות שבת', title: 'קונטרס בהלכות שבת', type: 'docx', bookUid: 'uid:1', source: 'library', author: 'ר׳ פלוני' }],
      categories: [
        {
          title: 'שו״ת',
          path: '/הלכה/שו״ת',
          books: [
            { bookId: 'שו״ת דברי שלום', title: 'שו״ת דברי שלום', type: 'docx', bookUid: 'uid:2', source: 'library' },
            { bookId: 'שו״ת טקסט', title: 'שו״ת טקסט', type: 'text' },
          ],
          categories: [],
        },
      ],
    },
    {
      title: 'ספרים אישיים',
      path: '/ספרים אישיים',
      books: [{ bookId: 'חידושי תורה שלי', title: 'חידושי תורה שלי', type: 'docx', bookUid: 'uid:3', source: 'user' }],
      categories: [],
    },
    { title: 'תנך', path: '/תנך', books: [{ bookId: 'בראשית', title: 'בראשית', type: 'text' }], categories: [] },
  ],
};

function devListing(path: string): unknown {
  const at = (days: number): string => new Date(Date.now() - days * 86_400_000).toISOString();
  if (path === '') {
    return {
      entries: [
        { name: 'שיעורים', path: 'שיעורים', type: 'dir', size: 0, modified: at(3) },
        { name: 'מכתבים', path: 'מכתבים', type: 'dir', size: 0, modified: at(9) },
        { name: 'חידושים על הדף.docx', path: 'חידושים על הדף.docx', type: 'file', size: 48_200, modified: at(0.1) },
        { name: 'Letter to the printer.docx', path: 'Letter to the printer.docx', type: 'file', size: 18_004, modified: at(12) },
      ],
      truncated: false,
    };
  }
  if (path === 'שיעורים') {
    return {
      entries: [{ name: 'שיעור א.docx', path: 'שיעורים/שיעור א.docx', type: 'file', size: 22_000, modified: at(1) }],
      truncated: false,
    };
  }
  return { entries: [], truncated: false };
}

function devOpen(name: string, writable: boolean): unknown {
  const blob = blankDocumentSource();
  if (!blob) throw new Error('error.internal: no blank document in dev');
  const token = `dev-file-${++counter}`;
  const url = URL.createObjectURL(blob);
  files.set(token, { name, url, writable });
  return { cancelled: false, token, url, name, size: blob.size, access: writable ? 'readwrite' : 'read' };
}

async function handle(method: string, payload: Record<string, unknown> = {}): Promise<unknown> {
  switch (method) {
    case 'fs.pickUserFile': {
      const { file } = await pickFile();
      if (!file) return { cancelled: true };
      const token = `dev-file-${++counter}`;
      const url = URL.createObjectURL(file);
      const writable = payload.access === 'readwrite';
      files.set(token, { name: file.name, url, writable });
      return {
        cancelled: false,
        token,
        url,
        name: file.name,
        size: file.size,
        access: writable ? 'readwrite' : 'read',
      };
    }

    case 'fs.resolveFileUrl': {
      const entry = files.get(String(payload.token));
      // ברענון הדף ה-blob אבד. זה בדיוק מה שקורה כשקובץ הוזז — והמסלול הזה
      // צריך להיבדק גם הוא.
      if (!entry) throw new Error('error.not_found: dev stub lost the file on reload');
      return { token: payload.token, url: entry.url, name: entry.name, size: 0 };
    }

    case 'fs.beginBinaryWrite': {
      const writeToken = `dev-write-${++counter}`;
      return {
        writeToken,
        uploadUrl: `${UPLOAD_PREFIX}${writeToken}`,
        expiresAt: new Date(Date.now() + 120_000).toISOString(),
        maxBytes: 100 * 1024 * 1024,
      };
    }

    case 'fs.abortBinaryWrite': {
      uploads.delete(String(payload.writeToken));
      return true;
    }

    case 'fs.commitUserFileWrite': {
      const blob = uploads.get(String(payload.writeToken));
      if (!blob) throw new Error('error.not_found: unknown or incomplete upload');
      uploads.delete(String(payload.writeToken));

      const target = payload.targetToken ? files.get(String(payload.targetToken)) : undefined;
      if (target) {
        // „שמור” לקובץ קיים. אין דיסק כאן, ולכן רק מדווחים — הבייטים אמיתיים.
        console.info(`[stub] נשמר ל-${target.name} (${blob.size} בייטים)`);
        return { cancelled: false, token: payload.targetToken, name: target.name, size: blob.size };
      }

      // הסיומת מהקריאה, כמו באוצריא — קבוע `.docx` היה מציע `ספר.txt.docx`
      // בייצוא לפורמט אוצריא.
      const suggestedBase = String(payload.suggestedName ?? 'מסמך');
      const extension = String(payload.extension ?? 'docx');
      const suggested = suggestedBase.toLowerCase().endsWith(`.${extension}`)
        ? suggestedBase
        : `${suggestedBase}.${extension}`;
      const name = window.prompt('שמור בשם (dev):', suggested);
      if (name === null) return { cancelled: true };
      // הורדה אמיתית, כדי שאפשר יהיה לפתוח את התוצר ב-Word.
      download(blob, name);
      const token = `dev-file-${++counter}`;
      files.set(token, { name, url: URL.createObjectURL(blob), writable: true });
      return { cancelled: false, token, name, size: blob.size };
    }

    // מה שמאפשר ל-resolveBoot לשחזר את מצב האתחול כשהאירוע אבד. אוצריא מעניקה
    // את app.info.read כהרשאת בסיס, ולכן זה עובד גם שם בלי הצהרה במניפסט.
    case 'app.getInfo':
      return BOOT.app;
    // „דיווח על באג”. כמו במאחז: http/https בלבד, ו-`true` בהצלחה.
    case 'app.openUrl': {
      const url = String(payload.url ?? '');
      if (!/^https?:\/\//i.test(url)) throw new Error('error.forbidden: only http/https URLs are allowed');
      window.open(url, '_blank', 'noopener');
      return true;
    }
    case 'app.getTheme':
      return BOOT.theme;
    /*
     * `PermissionSnapshot` — כלומר `{ permissions }` ולא המערך עצמו
     * (types/otzaria_plugin.d.ts). הדמה החזיר כאן מערך חשוף, והצרכן היחיד
     * (`explainHostGap` ב-engine/system-fonts.ts) קורא `info.permissions`,
     * ולכן בפיתוח הוא דיווח „לא ניתן לקרוא את ההרשאות שאושרו” על דמה שכן
     * ענה. `scripts/qa/host-stub.js` כבר החזיר את הצורה הנכונה.
     */
    case 'app.getGrantedPermissions':
      return { permissions: BOOT.permissions };

    case 'storage.get': {
      const raw = window.localStorage.getItem(storagePrefix + String(payload.key));
      return raw === null ? null : JSON.parse(raw);
    }
    case 'storage.set':
      window.localStorage.setItem(
        storagePrefix + String(payload.key),
        JSON.stringify(payload.value ?? null),
      );
      return true;
    case 'storage.remove':
      window.localStorage.removeItem(storagePrefix + String(payload.key));
      return true;

    /**
     * המרחב הפרטי של התוסף (`fs.*`), על localStorage.
     *
     * למה זה חייב להיות ב-stub: בלעדיו אין דרך לעבור בפיתוח את המסלול שבו
     * מסמך שלא נשמר חוזר אחרי סגירה — והמסלול הזה הוא בדיוק זה שאסור לו
     * להישבר. הנתונים שורדים רענון, כמו שהם שורדים סגירה באוצריא.
     */
    case 'fs.writeFile': {
      const content = String(payload.content ?? '');
      window.localStorage.setItem(workspacePrefix + String(payload.path), content);
      return { path: payload.path, size: content.length, usedBytes: content.length, quotaBytes: 0 };
    }
    case 'fs.readFile': {
      const content = window.localStorage.getItem(workspacePrefix + String(payload.path));
      if (content === null) throw new Error('error.not_found: no such file');
      return { path: payload.path, encoding: payload.encoding ?? 'utf8', size: content.length, content };
    }
    case 'fs.deleteEntry':
      window.localStorage.removeItem(workspacePrefix + String(payload.path));
      return true;
    case 'fs.stat': {
      const content = window.localStorage.getItem(workspacePrefix + String(payload.path));
      if (content === null) return { exists: false };
      return {
        exists: true,
        path: payload.path,
        name: String(payload.path),
        type: 'file',
        size: content.length,
        modified: new Date().toISOString(),
      };
    }

    case 'ui.showConfirm':
      return {
        confirmed: window.confirm(`${String(payload.title)}\n\n${String(payload.content)}`),
      };
    case 'ui.showError':
      console.error('[stub] שגיאה:', payload.message);
      return true;
    case 'ui.showMessage':
    case 'ui.showSuccess':
      console.info('[stub]', payload.message);
      return true;

    /**
     * `ui.exportPdf` — מחוץ לאוצריא אין דיאלוג „שמור בשם” של מערכת ההפעלה
     * ואין מנוע שמייצר PDF. הכפיל מדפיס את מה שהיה נשלח ומדווח ביטול: זהו
     * המצב שאינו יוצר קובץ ואינו מתחזה להצלחה, ובכל זאת מריץ את כל מסלול
     * ההכנה (`@page` והגלון) שאותו כן אפשר לבדוק בדפדפן.
     *
     * ה-`version` של הכפיל הוא 0.9.97, ולכן הפקד **כן** פעיל בפיתוח — אחרת
     * מסלול ההכנה לא היה נבדק כלל מחוץ לאוצריא.
     */
    case 'ui.exportPdf':
      console.info('[stub] ui.exportPdf', payload);
      return { saved: false, name: null };

    /**
     * `library.refreshUserBooks` — אין ספרייה בפיתוח. „אפס ספרים נקלטו” הוא
     * המצב האמיתי מחוץ לאוצריא, והוא גם מפעיל את נוסח ההודעה שמסביר על
     * תיקייה אישית — הנוסח שדווקא צריך בדיקה בדפדפן.
     */
    case 'library.refreshUserBooks':
      console.info('[stub] library.refreshUserBooks');
      return { addedBooks: 0, updatedBooks: 0, errors: [] };

    /*
     * העץ של „פתח מסמך”. ספרייה ותיקייה מדומות, כדי שכל הענפים — כולל
     * הקינון, החיפוש והמצבים הריקים — ייבדקו בדפדפן. „פתיחה” מחזירה את
     * המסמך הריק של התוסף: הבייטים אינם העניין כאן, המסלול כן.
     */
    case 'library.getTree':
      return DEV_LIBRARY;

    case 'library.openBookFile':
      return devOpen(`${String(payload.bookUid ?? payload.bookId ?? 'ספר')}.docx`, payload.access === 'readwrite');

    case 'fs.pickUserFolder': {
      const name = window.prompt('שם התיקייה (dev):', 'חידושים');
      if (!name) return { cancelled: true };
      return { cancelled: false, folderToken: `dev-folder-${name}`, name, path: `C:\\Users\\dev\\Documents\\${name}` };
    }

    case 'fs.listUserFolder':
      return devListing(String(payload.path ?? ''));

    case 'fs.openFolderFile':
      return devOpen(String(payload.path ?? 'מסמך.docx').split('/').pop()!, payload.access === 'readwrite');

    case 'fs.revokeFolder':
      return true;

    default:
      console.info('[stub] call לא ממומש:', method, payload);
      return null;
  }
}

export function installDevStub(): void {
  if ((window as Partial<Window>).Otzaria) return;

  // ה-PUT של ההעלאה אינו יוצא לרשת — נתפס כאן ונשמר בזיכרון.
  const realFetch = window.fetch.bind(window);
  window.fetch = async (input, init) => {
    const url = typeof input === 'string' ? input : input instanceof URL ? input.href : input.url;
    if (init?.method === 'PUT' && url.startsWith(UPLOAD_PREFIX)) {
      const writeToken = url.slice(UPLOAD_PREFIX.length);
      const body = init.body;
      uploads.set(writeToken, body instanceof Blob ? body : new Blob([]));
      return new Response(null, { status: 204 });
    }
    return realFetch(input, init);
  };

  const stub = {
    async call(method: string, payload?: Record<string, unknown>) {
      try {
        return { success: true, data: await handle(method, payload), error: null };
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        const [code] = message.split(':');
        return { success: false, data: null, error: { code, message } };
      }
    },
    on(event: string, cb: (payload: unknown) => void) {
      window.addEventListener(event, (e) => cb((e as CustomEvent).detail));
    },
    off() {
      // ה-stub אינו שומר הפניות; אין צורך בביטול בפיתוח.
    },
  };

  // ה-stub מממש רק את מה שהתוסף באמת קורא לו, ולא את כל העומסים של
  // OtzariaGlobal — לכן ההמרה המפורשת.
  window.Otzaria = stub as unknown as OtzariaGlobal;

  setTimeout(() => {
    window.dispatchEvent(new CustomEvent('plugin.boot', { detail: BOOT }));
  }, 0);
}
