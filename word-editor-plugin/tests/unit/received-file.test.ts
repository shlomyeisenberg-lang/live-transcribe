/**
 * קובץ Word שתוסף אחר שולח (`plugin.page_opened`) — „סייר קבצים” פתח את
 * הלשונית והמסמך לא נטען, כי לא היה מי שמאזין.
 *
 * שני חלקים נבדקים כאן: הפענוח של ה-`param` (מה נפתח, מה מוסבר למשתמש ומה
 * נבלע בשקט), וה-latch שב-`index.html` — האירוע מגיע בסוף ה-boot, שניות לפני
 * שהבאנדל נטען, ובלעדיו התיקון כולו אינו עושה דבר בדיוק במקרה הנפוץ: הסייר
 * פותח את התוסף כשהוא עוד לא פתוח.
 */
import { describe, it, expect, afterEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { bytesToBase64 } from '../../src/host/base64';
import { DOCX_MIME } from '../../src/engine/export';
import {
  PAGE_OPENED_LATCH_KEY,
  readReceivedFile,
  takePendingPageOpens,
  type PageOpenedEvent,
} from '../../src/host/received-file';

/** jsdom אינו מממש `Blob.arrayBuffer`; `FileReader` כן. */
function blobBytes(blob: Blob): Promise<Uint8Array> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(new Uint8Array(reader.result as ArrayBuffer));
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(blob);
  });
}

/** תחילת חבילת zip, כלומר תחילת כל docx. */
const DOCX_BYTES = new Uint8Array([0x50, 0x4b, 0x03, 0x04, 0x14, 0x00, 0x06, 0x00, 0xff]);

/** בדיוק מה שהסייר שולח במצב גשר. */
function fromExplorer(param: unknown, openedBy = 'com.example.explorer'): PageOpenedEvent {
  return { param, openedBy };
}

const dataUrl = `data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,${bytesToBase64(DOCX_BYTES)}`;

describe('readReceivedFile', () => {
  it('dataUrl נפתח כבייטים של הקובץ, עם השם שנשלח', async () => {
    const received = readReceivedFile(fromExplorer({ name: 'חידושים.docx', dataUrl }));

    expect(received.kind).toBe('file');
    if (received.kind !== 'file') return;
    expect(received.name).toBe('חידושים.docx');
    expect(received.blob.type).toBe(DOCX_MIME);
    expect(Array.from(await blobBytes(received.blob))).toEqual(Array.from(DOCX_BYTES));
  });

  it('ה-MIME שב-dataUrl אינו קובע: הסייר יכול לשלוח application/octet-stream', () => {
    const generic = `data:application/octet-stream;base64,${bytesToBase64(DOCX_BYTES)}`;
    expect(readReceivedFile(fromExplorer({ name: 'א.docx', dataUrl: generic })).kind).toBe('file');
  });

  it('שם חסר או ריק — „מסמך”, ושם עם רווחים מסביב נחתך', () => {
    const noName = readReceivedFile(fromExplorer({ dataUrl }));
    const blank = readReceivedFile(fromExplorer({ name: '   ', dataUrl }));
    const padded = readReceivedFile(fromExplorer({ name: '  ב.docx ', dataUrl }));

    expect(noName.kind === 'file' && noName.name).toBe('מסמך');
    expect(blank.kind === 'file' && blank.name).toBe('מסמך');
    expect(padded.kind === 'file' && padded.name).toBe('ב.docx');
  });

  /**
   * „הלשונית נפתחה והמסמך לא” הוא הבאג. param שמנסה לשלוח קובץ ונכשל חייב
   * להגיע להודעה — לא לשקט שממנו התחלנו.
   */
  it.each([
    ['base64 פגום', 'data:application/octet-stream;base64,@@@'],
    ['dataUrl ריק', 'data:application/octet-stream;base64,'],
    ['לא base64', 'data:text/plain,hello'],
    ['לא data: בכלל', 'https://example.com/a.docx'],
  ])('%s — נדחה עם הודעה שמזכירה את הקובץ', (_label, value) => {
    const received = readReceivedFile(fromExplorer({ name: 'ג.docx', dataUrl: value }));

    expect(received.kind).toBe('rejected');
    expect(received.kind === 'rejected' && received.message).toContain('ג.docx');
  });

  /**
   * שלוש הצורות שאוצריא חוסמת בין תוספים (ראו ראש host/received-file.ts).
   * `{ name, path }` הוא בדיוק מה שהסייר שולח היום.
   */
  it.each([
    ['url של שרת הקבצים', { url: 'http://127.0.0.1:1/f/com.example.explorer/tok' }],
    ['token', { token: 'tok' }],
    ['path בלבד, כמו הסייר היום', { path: 'C:\\docs\\ד.docx' }],
    ['path null, כמו הסייר בלי גשר', { path: null }],
  ])('%s — נדחה עם הסבר ועם הדרך החלופית', (_label, fields) => {
    const received = readReceivedFile(fromExplorer({ name: 'ד.docx', ...fields }));

    expect(received.kind).toBe('rejected');
    if (received.kind !== 'rejected') return;
    expect(received.message).toContain('ד.docx');
    expect(received.message).toContain('לא צירף את תוכן הקובץ');
    expect(received.message).toContain('„פתח”');
  });

  it('dataUrl גובר על url שנשלח לצדו', () => {
    const received = readReceivedFile(fromExplorer({ name: 'ה.docx', url: 'http://x/f/y/z', dataUrl }));
    expect(received.kind).toBe('file');
  });

  /**
   * אותו אירוע נורה גם בלי קובץ: כפתור „לשונית חדשה” של אוצריא מוסר
   * `{ source: 'newTabButton' }` בלי `param`, ו-`plugin.openSelf` עשוי למסור
   * `param: null`. אלה אינם כשל ואסור שיציגו הודעה.
   */
  it.each([
    ['כפתור „לשונית חדשה”', { source: 'newTabButton' }],
    ['param null', { param: null }],
    ['param מחרוזת', { param: 'x' }],
    ['param מערך', { param: [dataUrl] }],
    ['param בלי שדות של קובץ', { param: { name: 'ו.docx' } }],
    ['אירוע בלי detail', null],
  ])('%s — מתעלמים', (_label, event) => {
    expect(readReceivedFile(event as PageOpenedEvent | null)).toEqual({ kind: 'ignore' });
  });
});

/* ===========================================================================
 *  ה-latch
 * ========================================================================= */

describe('ה-latch של plugin.page_opened', () => {
  // index.html ב-LF במאגר; הנרמול רק מונע שבירה אצל מי ש-git המיר ל-CRLF.
  const html = readFileSync(join(process.cwd(), 'index.html'), 'utf8').replace(/\r\n/g, '\n');

  afterEach(() => {
    delete (window as unknown as Record<string, unknown>)[PAGE_OPENED_LATCH_KEY];
  });

  it('ה-latch ב-index.html נושא את השם שהקוד קורא לו, ומאזין לאירוע של אוצריא', () => {
    expect(html).toContain(`window.${PAGE_OPENED_LATCH_KEY} =`);
    expect(html).toContain("window.addEventListener('plugin.page_opened'");
  });

  it('הוא רץ ב-head, לפני הבאנדל — אחרת אין לו טעם', () => {
    expect(html.indexOf(PAGE_OPENED_LATCH_KEY)).toBeLessThan(html.indexOf('src/main.ts'));
  });

  /** הסקריפט האמיתי מ-index.html, ולא latch מזויף שמניח איך הוא מתנהג. */
  function runLatchScript(): void {
    const at = html.indexOf(`window.${PAGE_OPENED_LATCH_KEY} =`);
    const start = html.lastIndexOf('(function () {', at);
    const end = html.indexOf('})();', at) + '})();'.length;
    new Function(html.slice(start, end))();
  }

  function pageOpened(detail: unknown): void {
    window.dispatchEvent(new CustomEvent('plugin.page_opened', { detail }));
  }

  it('צובר את ה-detail, ומוסר אותו בריקון — כך שהוא עובר את הפענוח', () => {
    runLatchScript();
    pageOpened(fromExplorer({ name: 'ז.docx', dataUrl }));

    const pending = takePendingPageOpens();
    expect(pending).toEqual([fromExplorer({ name: 'ז.docx', dataUrl })]);
    expect(readReceivedFile(pending[0]).kind).toBe('file');
  });

  it('אחרי הריקון הוא מפסיק לצבור — המאזין החי הוא היחיד שרואה אירועים', () => {
    runLatchScript();
    takePendingPageOpens();
    pageOpened(fromExplorer({ name: 'ח.docx', dataUrl }));

    expect(takePendingPageOpens()).toEqual([]);
  });

  it('שומר את הסדר, ומחזיק לכל היותר שמונה — הישנים יורדים', () => {
    runLatchScript();
    for (let i = 0; i < 10; i += 1) pageOpened({ param: { name: `${i}.docx`, dataUrl } });

    const names = takePendingPageOpens().map((event) => (event.param as { name: string }).name);
    expect(names).toEqual(['2.docx', '3.docx', '4.docx', '5.docx', '6.docx', '7.docx', '8.docx', '9.docx']);
  });

  it('בלי latch (בדיקות, דפדפן בלי index.html) — אין ממתינים, ואין זריקה', () => {
    expect(takePendingPageOpens()).toEqual([]);
  });
});
