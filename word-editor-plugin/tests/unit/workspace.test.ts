/**
 * המרחב הפרטי של התוסף — המקום היחיד שיכול להחזיק מסמך שטרם נשמר.
 *
 * שני דברים נבדקים כאן, וכל אחד מהם הוא נזק שקט אם הוא שבור:
 *
 * 1. **המגבלה נמדדת על המחרוזת ולא על הבייטים.** מה שעובר בגשר הוא base64,
 *    שתופס שליש יותר. בדיקה על הבייטים מאשרת קובץ שהמכסה תדחה — אחרי שכבר
 *    שילמנו על ההמרה.
 * 2. **שום מסלול אינו זורק.** זו רשת ביטחון: כשל בכתיבת טיוטה אינו סיבה
 *    להפיל עריכה, וכשל בקריאתה אינו סיבה שהתוסף לא יעלה.
 */
import { describe, it, expect, vi, afterEach } from 'vitest';
import { base64Length, base64ToBytes, bytesToBase64 } from '../../src/host/base64';
import {
  MAX_CONTENT_BYTES,
  MAX_PAYLOAD_BYTES,
  deleteWorkspaceEntry,
  readWorkspaceBytes,
  writeWorkspaceBytes,
} from '../../src/host/workspace';

function hostReturns(data: unknown): ReturnType<typeof vi.fn> {
  const call = vi.fn(async () => ({ success: true, data, error: null }));
  window.Otzaria = { call } as never;
  return call;
}

function hostFails(): ReturnType<typeof vi.fn> {
  const call = vi.fn(async () => ({
    success: false,
    data: null,
    error: { code: 'error.forbidden', message: 'לא' },
  }));
  window.Otzaria = { call } as never;
  return call;
}

afterEach(() => {
  delete (window as Partial<Window>).Otzaria;
  vi.restoreAllMocks();
});

describe('base64', () => {
  it('הלוך ושוב על בייטים שאינם טקסט', () => {
    const bytes = new Uint8Array([0, 1, 80, 75, 3, 4, 255, 254, 127]);
    const round = base64ToBytes(bytesToBase64(bytes));
    expect(round && Array.from(round)).toEqual(Array.from(bytes));
  });

  it('עובד על קלט גדול, ולא זורק RangeError', () => {
    // המלכודת שהמודול נכתב סביבה: `fromCharCode(...bytes)` על מערך גדול חורג
    // ממגבלת הארגומנטים — כלומר דווקא הקבצים הגדולים היו נכשלים.
    const bytes = new Uint8Array(300_000).fill(65);
    expect(() => bytesToBase64(bytes)).not.toThrow();
    expect(base64ToBytes(bytesToBase64(bytes))?.byteLength).toBe(300_000);
  });

  it('מחרוזת פגומה מוחזרת כ-null ולא כזריקה', () => {
    expect(base64ToBytes('לא base64 בכלל!!')).toBeNull();
  });

  it('אורך ה-base64 הוא שליש יותר מהבייטים', () => {
    expect(base64Length(3)).toBe(4);
    expect(base64Length(3_000)).toBe(4_000);
    expect(MAX_CONTENT_BYTES).toBeLessThan(MAX_PAYLOAD_BYTES);
    expect(base64Length(MAX_CONTENT_BYTES)).toBeLessThanOrEqual(MAX_PAYLOAD_BYTES);
  });
});

describe('כתיבה', () => {
  it('שולחת base64 בנתיב המבוקש', async () => {
    const call = hostReturns({ path: 'session-draft.docx', size: 3 });

    await expect(
      writeWorkspaceBytes('session-draft.docx', new Uint8Array([1, 2, 3])),
    ).resolves.toBe('written');

    expect(call).toHaveBeenCalledWith('fs.writeFile', {
      path: 'session-draft.docx',
      content: bytesToBase64(new Uint8Array([1, 2, 3])),
      encoding: 'base64',
    });
  });

  it('קובץ שה-base64 שלו חורג מהמכסה נדחה', async () => {
    // הבייטים עצמם מתחת ל-10MB; ה-base64 שלהם מעל. בדיקה על הבייטים הייתה
    // מאשרת אותו ורק אז נדחית על ידי אוצריא — אחרי ההמרה.
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    const call = hostReturns({});
    const oversize = new Uint8Array(MAX_CONTENT_BYTES + 1);
    expect(oversize.byteLength).toBeLessThan(MAX_PAYLOAD_BYTES);

    // `too-large` ולא `failed`: זהו כשל קבוע — המסמך לא יקטן מעצמו — והוא
    // היחיד שראוי להיאמר למשתמש. ההבחנה נעשית כאן מפני שכאן נמדד הגודל.
    await expect(writeWorkspaceBytes('draft.docx', oversize)).resolves.toBe('too-large');

    expect(call, 'לא נשלחה קריאה בכלל').not.toHaveBeenCalled();
    expect(warn).toHaveBeenCalled();
  });

  it('סירוב של אוצריא מוחזר כ-failed ואינו זורק', async () => {
    // חולף, ולא תכונה של המסמך: הסבב הבא עשוי להצליח, ולכן אין מדווחים
    // עליו למשתמש. ההפרדה מ-`too-large` היא בדיוק זו.
    const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
    hostFails();

    await expect(writeWorkspaceBytes('draft.docx', new Uint8Array([1]))).resolves.toBe('failed');
    expect(warn).toHaveBeenCalled();
  });

  it('היעדר SDK אינו מפיל עריכה', async () => {
    vi.spyOn(console, 'warn').mockImplementation(() => {});
    await expect(writeWorkspaceBytes('draft.docx', new Uint8Array([1]))).resolves.toBe('failed');
  });
});

describe('קריאה', () => {
  it('מחזירה את הבייטים שנכתבו', async () => {
    const bytes = new Uint8Array([80, 75, 3, 4]);
    hostReturns({ content: bytesToBase64(bytes), encoding: 'base64' });

    const read = await readWorkspaceBytes('draft.docx');
    expect(read && Array.from(read)).toEqual(Array.from(bytes));
  });

  it('קובץ חסר, תשובה זרה, או base64 פגום — כולם „אין מה לשחזר”', async () => {
    vi.spyOn(console, 'warn').mockImplementation(() => {});

    hostFails();
    await expect(readWorkspaceBytes('draft.docx')).resolves.toBeNull();

    hostReturns({ content: 42 });
    await expect(readWorkspaceBytes('draft.docx')).resolves.toBeNull();

    hostReturns({ content: '@@@ לא base64' });
    await expect(readWorkspaceBytes('draft.docx')).resolves.toBeNull();
  });
});

describe('מחיקה', () => {
  it('מוחקת בנתיב המבוקש', async () => {
    const call = hostReturns(true);
    await deleteWorkspaceEntry('draft.docx');
    expect(call).toHaveBeenCalledWith('fs.deleteEntry', { path: 'draft.docx' });
  });

  it('מחיקה שנכשלה אינה זורקת', async () => {
    hostFails();
    await expect(deleteWorkspaceEntry('draft.docx')).resolves.toBeUndefined();
  });
});
