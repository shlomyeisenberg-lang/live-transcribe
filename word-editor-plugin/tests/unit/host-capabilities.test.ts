/**
 * זמינות קריאות ה-Host לפי גרסת אוצריא.
 *
 * הנקודה שנשמרת כאן היא שהברירה היא „לא”: פקד שנשען על קריאה שאינה קיימת
 * צריך להיראות מנוטרל, ולא להיכשל בלחיצה.
 */
import { beforeEach, describe, expect, it } from 'vitest';
import {
  EXPORT_PDF_MIN_APP_VERSION,
  compareAppVersions,
  hostAppVersion,
  setHostAppVersion,
  supportsPdfExport,
} from '../../src/host/host-capabilities';

beforeEach(() => setHostAppVersion(null));

describe('compareAppVersions', () => {
  it('משווה מקטע-מקטע ולא כמחרוזת', () => {
    // ההשוואה הלקסיקוגרפית הייתה אומרת ש-'0.9.100' < '0.9.97', וזה בדיוק
    // המספור שאוצריא נמצאת בו.
    expect(compareAppVersions('0.9.100', '0.9.97')).toBeGreaterThan(0);
    expect(compareAppVersions('0.9.97', '0.9.100')).toBeLessThan(0);
    expect(compareAppVersions('0.10.0', '0.9.99')).toBeGreaterThan(0);
  });

  it('שווה לעצמה, ואורך שונה משלים באפסים', () => {
    expect(compareAppVersions('0.9.97', '0.9.97')).toBe(0);
    expect(compareAppVersions('1', '1.0.0')).toBe(0);
    expect(compareAppVersions('1.0.1', '1')).toBeGreaterThan(0);
  });

  it('סיומת אחרי המספרים אינה משנה את הסדר', () => {
    expect(compareAppVersions('0.9.97-rc1', '0.9.97')).toBe(0);
    expect(compareAppVersions('0.9.97+build3', '0.9.96')).toBeGreaterThan(0);
  });
});

describe('supportsPdfExport', () => {
  it('בלי גרסה — נכשל סגור', () => {
    expect(hostAppVersion()).toBe(null);
    expect(supportsPdfExport.value).toBe(false);
  });

  it('גרסה ריקה או רווחים בלבד נחשבות כאין גרסה', () => {
    for (const value of ['', '   ', null, undefined]) {
      setHostAppVersion(value);
      expect(hostAppVersion(), String(value)).toBe(null);
      expect(supportsPdfExport.value, String(value)).toBe(false);
    }
  });

  it('גרסה ישנה — לא נתמך; הגרסה עצמה ומעלה — נתמך', () => {
    setHostAppVersion('0.9.96');
    expect(supportsPdfExport.value).toBe(false);

    setHostAppVersion(EXPORT_PDF_MIN_APP_VERSION);
    expect(supportsPdfExport.value).toBe(true);

    setHostAppVersion('0.9.98');
    expect(supportsPdfExport.value).toBe(true);

    setHostAppVersion('1.0.0');
    expect(supportsPdfExport.value).toBe(true);
  });

  it('הערך תגובתי — שינוי גרסה מעדכן את החישוב', () => {
    setHostAppVersion('0.9.90');
    expect(supportsPdfExport.value).toBe(false);
    setHostAppVersion('0.9.97');
    expect(supportsPdfExport.value).toBe(true);
  });
});
