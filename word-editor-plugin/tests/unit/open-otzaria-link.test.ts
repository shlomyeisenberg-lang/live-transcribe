/**
 * פתיחת קישור `otzaria://` מול המאחז.
 *
 * זה הצד השני של הלחיצה (‏#69), והוא נבדק כאן מפני שהכשל שדווח לא היה
 * „נזרקה שגיאה” אלא **שקט**: `reader.openBook` מחזיר `false` כשהספר אינו
 * נמצא — וגם כשה-`id` עמום, כי המאחז מסנן לפי `id`+`type` ושתי התאמות
 * מחזירות `null` אצלו. התוצאה נזרקה לפח, ולכן „לחצתי ולא קרה כלום, וגם לא
 * נאמר לי למה”.
 */
import { beforeEach, describe, expect, it, vi } from 'vitest';

const { callMock } = vi.hoisted(() => ({ callMock: vi.fn() }));

vi.mock('../../src/host/otzaria-client', () => ({
  call: callMock,
  isPermissionDenied: (error: unknown) =>
    error instanceof Error && error.message.includes('permission_denied'),
  hostErrorCode: (error: unknown) =>
    error instanceof Error && error.message.includes('not_found') ? 'error.not_found' : null,
}));

const { openOtzariaLink } = await import('../../src/host/otzaria-reader');

beforeEach(() => {
  callMock.mockReset();
  callMock.mockResolvedValue(true);
});

describe('openOtzariaLink', () => {
  it('ספר: שולח id, index ו-type', async () => {
    const result = await openOtzariaLink({ kind: 'book', id: 42, index: 1234 });

    expect(result).toEqual({ ok: true, value: undefined });
    expect(callMock).toHaveBeenCalledWith('reader.openBook', {
      id: 42,
      index: 1234,
      type: 'text',
      navigateToPositionIfReused: true,
    });
  });

  /** המזהה היציב מכריע אצל המאחז לפני כל ניחוש לפי `id`. */
  it('שולח bookUid כשהקישור נושא אותו', async () => {
    await openOtzariaLink({ kind: 'book', id: 42, index: 0, uid: 'id:42' });

    expect(callMock.mock.calls[0]![1]).toMatchObject({ bookUid: 'id:42', id: 42 });
  });

  it('קישור ישן בלי uid אינו שולח את השדה כלל', async () => {
    await openOtzariaLink({ kind: 'book', id: 42, index: 0 });

    expect(callMock.mock.calls[0]![1]).not.toHaveProperty('bookUid');
  });

  it('PDF נשלח עם type של pdf', async () => {
    await openOtzariaLink({ kind: 'pdf', id: 7, index: 3 });

    expect(callMock.mock.calls[0]![1]).toMatchObject({ type: 'pdf', id: 7, index: 3 });
  });

  /** זה הבאג של #69: `false` אינו זריקה, והוא נבלע. */
  it('false מהמאחז הוא כשל, עם הודעה שאומרת מה קרה', async () => {
    callMock.mockResolvedValue(false);

    const result = await openOtzariaLink({ kind: 'book', id: 42, index: 0 });

    expect(result).toEqual({
      ok: false,
      reason: 'book-not-found',
      message: 'פתיחת הקישור נכשלה: הספר אינו נמצא בספרייה של אוצריא',
    });
  });

  it('תשובה בצורה אחרת אינה כשל — רק false הוא סירוב', async () => {
    callMock.mockResolvedValue(null);

    expect(await openOtzariaLink({ kind: 'book', id: 42, index: 0 })).toEqual({
      ok: true,
      value: undefined,
    });
  });

  it('זריקה מתורגמת להודעה בעברית', async () => {
    callMock.mockRejectedValue(new Error('error.not_found: boom'));

    const result = await openOtzariaLink({ kind: 'book', id: 42, index: 0 });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.message).toContain('פתיחת הקישור נכשלה');
    expect(result.reason).toBe('error.not_found');
  });

  it('הרשאה חסרה אומרת איזו הרשאה חסרה', async () => {
    callMock.mockRejectedValue(new Error('permission_denied'));

    const result = await openOtzariaLink({ kind: 'book', id: 42, index: 0 });

    expect(result.ok).toBe(false);
    if (result.ok) return;
    expect(result.reason).toBe('permission-denied');
    expect(result.message).toContain('reader.open');
  });

  it('detection נפתח כלשונית חיפוש', async () => {
    await openOtzariaLink({ kind: 'detection', query: 'בסוגיא דדיורים בריבית' });

    expect(callMock).toHaveBeenCalledWith('reader.openSearchTab', {
      query: 'בסוגיא דדיורים בריבית',
    });
  });

  it('סירוב של detection מדווח אף הוא', async () => {
    callMock.mockResolvedValue(false);

    const result = await openOtzariaLink({ kind: 'detection', query: 'א' });

    expect(result.ok).toBe(false);
  });
});
