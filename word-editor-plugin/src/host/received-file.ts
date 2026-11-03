/**
 * קובץ Word שתוסף אחר שולח לעורך — למשל „סייר קבצים”, בלחיצה על קובץ docx.
 *
 * ## המסלול
 *
 * התוסף השולח קורא ל-`plugin.openOther` עם `pluginId` שלנו ו-`param`, ואוצריא
 * פותחת את הלשונית ויורה אצלנו `plugin.page_opened` עם `{ param, openedBy }`.
 * אם הדף עוד עולה, אוצריא מחזיקה את האירוע עד סוף ה-boot ומוסרת אותו אז —
 * כלומר **לפני** שהבאנדל שלנו נטען והמאזין נרשם. לכן יש כאן latch, בדיוק כמו
 * ב„שלח למסמך” (`CONTEXT_MENU_LATCH_KEY` ב-otzaria-reader.ts).
 *
 * ## למה רק `dataUrl`
 *
 * הצורה שנתמכת היא `{ name, dataUrl }` — תוכן הקובץ עצמו, ב-base64. שתי
 * הצורות האחרות שנראות טבעיות אינן יכולות לעבוד, וזה בכוונה של אוצריא ולא פער:
 *
 * - **`url`** מ-`fs.resolveFileUrl` הוא `/f/<pluginId>/<token>` של התוסף
 *   השולח. שער ה-WebView מאפשר לכל תוסף רק את הנתיבים שנושאים את המזהה שלו
 *   (`PluginFileServer.isUriForPlugin` באוצריא), ולכן `fetch` שלנו נחסם.
 * - **`token`** של קובץ נשמר ב-KV של התוסף שקיבל אותו (`user_file_grants`),
 *   ו-`fs.resolveFileUrl` שלנו עונה עליו „unknown file token”.
 *
 * התוסף השולח יכול לקרוא את ה-`url` שלו עצמו ב-`fetch` ולשלוח את התוכן. מי
 * ששולח רק `url`, `token` או `path` מקבל הודעה שאומרת בדיוק את זה, ולא שקט:
 * „הלשונית נפתחה והמסמך לא” הוא בדיוק הבאג שהמודול הזה סוגר.
 */
import { base64ToBytes } from './base64';
import { DOCX_MIME } from '../engine/export';
import type { OtzariaEventMap } from '../types/otzaria_plugin';

/** ה-payload של `plugin.page_opened`, כפי שאוצריא יורה אותו. */
export type PageOpenedEvent = OtzariaEventMap['plugin.page_opened'];

/**
 * מה לעשות עם האירוע.
 *
 * `ignore` — אין כאן קובץ בכלל: כפתור „לשונית חדשה” של אוצריא יורה את אותו
 * אירוע בלי `param`, וכך גם `plugin.openSelf`. `rejected` — הייתה כוונה לשלוח
 * קובץ, והוא לא הגיע; ההודעה מיועדת לשורת המצב.
 */
export type ReceivedFile =
  | { kind: 'ignore' }
  | { kind: 'rejected'; message: string }
  | { kind: 'file'; name: string; blob: Blob };

/** השם כשהשולח לא נתן שם. */
const FALLBACK_NAME = 'מסמך';

/**
 * `data:[<mime>][;base64],<data>`. רק base64 נתמך: זה מה ש-`FileReader` ו-
 * `bytesToBase64` מייצרים, וקידוד אחוזים של קובץ בינארי אינו דבר ששולחים.
 */
const DATA_URL = /^data:[^,]*;base64,/i;

function nameOf(param: Record<string, unknown>): string {
  const name = typeof param.name === 'string' ? param.name.trim() : '';
  return name || FALLBACK_NAME;
}

/**
 * מפענחת את האירוע לקובץ שאפשר לפתוח, להודעה, או ל„אין כאן כלום”.
 *
 * אין כאן בדיקה שהבייטים הם docx: הסייר שולח רק `.docx`, ומי ששולח משהו אחר
 * נופל בפתיחה עצמה — ושם כבר יש הודעה. בדיקה כפולה הייתה ענף בלי קלט שמגיע
 * אליו.
 */
export function readReceivedFile(event: PageOpenedEvent | null | undefined): ReceivedFile {
  const param = event?.param;
  if (!param || typeof param !== 'object') return { kind: 'ignore' };

  const fields = param as Record<string, unknown>;
  const name = nameOf(fields);

  if (typeof fields.dataUrl === 'string') {
    const header = DATA_URL.exec(fields.dataUrl);
    const bytes = header ? base64ToBytes(fields.dataUrl.slice(header[0].length)) : null;
    if (!bytes || bytes.byteLength === 0) {
      return { kind: 'rejected', message: `„${name}” לא נפתח: התוכן שהתקבל מהתוסף השולח פגום` };
    }
    return { kind: 'file', name, blob: new Blob([bytes], { type: DOCX_MIME }) };
  }

  if (fields.url !== undefined || fields.token !== undefined || fields.path !== undefined) {
    return {
      kind: 'rejected',
      message: `„${name}” לא נפתח: התוסף ששלח אותו לא צירף את תוכן הקובץ. אפשר לפתוח אותו כאן דרך „פתח”`,
    };
  }

  return { kind: 'ignore' };
}

/**
 * השם של ה-latch ב-`index.html`. מוגדר כאן ושם, ו-
 * tests/unit/received-file.test.ts מקבע את הזהות — כמו `CONTEXT_MENU_LATCH_KEY`.
 */
export const PAGE_OPENED_LATCH_KEY = '__otzariaPageOpened';

interface PageOpenedLatch {
  queue: PageOpenedEvent[];
  live: boolean;
}

/**
 * מרוקנת את ה-latch ומעבירה אותו למצב „חי”.
 *
 * החוזה זהה ל-`takePendingContextMenuClicks`: קודם נרשם המאזין החי, ורק אחריו
 * הריקון, ושניהם באותה משימה סינכרונית — אין חלון שבו אירוע נופל בין השניים,
 * ומרגע ש-`live` דלוק ה-latch אינו צובר ואין טיפול כפול.
 */
export function takePendingPageOpens(): PageOpenedEvent[] {
  const latch = (window as unknown as Record<string, PageOpenedLatch | undefined>)[
    PAGE_OPENED_LATCH_KEY
  ];
  if (!latch) return [];
  latch.live = true;
  return latch.queue.splice(0, latch.queue.length);
}
