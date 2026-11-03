/**
 * צורת התשובות של דמה הפיתוח — לא ההתנהגות שלו, אלא **המעטפת**.
 *
 * ## למה זה קובץ בדיקה בכלל
 *
 * לתוסף שלושה דמים של המאחז: `host/dev-stub.ts` (פיתוח בדפדפן),
 * `scripts/qa/host-stub.js` (שערי הדפדפן) ו-`scripts/session-probe.mjs`.
 * שלושתם עונים על אותן מתודות, ואף אחד מהם לא היה נבדק — ולכן שניים מהם
 * נסחפו מהחוזה בשקט, כל אחד במתודה אחרת:
 *
 * - `host-stub.js` החזיר `storage.get` עטוף ב-`{ value }`, ולכן כל העדפה
 *   שנזרעה בשער נדחתה בנרמול ונפלה לברירת המחדל. ~60 שערים רצו כך בירוק.
 * - `dev-stub.ts` החזיר את **מערך** ההרשאות במקום `PermissionSnapshot`
 *   (`{ permissions }`), ולכן `explainHostGap` דיווח בפיתוח „לא ניתן לקרוא
 *   את ההרשאות שאושרו” על דמה שכן ענה.
 *
 * שני הכשלים מאותה משפחה: **מעטפת שגויה אינה שגיאה אלא שתיקה** — הצרכן מקבל
 * אובייקט, לא מזהה את השדה שהוא מחפש, ונופל למסלול „אין נתונים”. ולכן מה
 * שנבדק כאן הוא הצורה, ולא מה שיש בתוכה.
 *
 * ההתנהגות מול המוצר נמדדת ב-`scripts/qa/settings-restore-qa.mjs`, שזורע
 * העדפה לפני שהתוסף קם ובודק מה המשתמש רואה. שם החוזה, כאן הצורה — ומי ששובר
 * אחד מהם רואה אדום לפני שהשני נדרש.
 */
import { beforeAll, describe, expect, it } from 'vitest';
import { installDevStub } from '../../src/host/dev-stub';

interface Envelope {
  success: boolean;
  data: unknown;
  error: unknown;
}

function otzaria(): { call: (method: string, payload?: Record<string, unknown>) => Promise<Envelope> } {
  const global = window as unknown as {
    Otzaria?: { call: (method: string, payload?: Record<string, unknown>) => Promise<Envelope> };
  };
  if (!global.Otzaria) throw new Error('הדמה לא הותקן');
  return global.Otzaria;
}

beforeAll(() => {
  installDevStub();
});

describe('דמה הפיתוח — צורת התשובות', () => {
  it('„מה אושר” הוא PermissionSnapshot ולא המערך עצמו', async () => {
    const res = await otzaria().call('app.getGrantedPermissions');

    expect(res.success).toBe(true);
    // `{ permissions }` — types/otzaria_plugin.d.ts, וזה גם מה שהצרכן היחיד
    // קורא (`explainHostGap` ב-engine/system-fonts.ts). מערך חשוף כאן נראה
    // למי שקורא `info.permissions` בדיוק כמו מארח שלא ענה.
    expect(Array.isArray(res.data)).toBe(false);
    expect(res.data).toEqual({
      permissions: expect.arrayContaining(['app.info.read', 'fs.user_files.read']),
    });
  });

  it('`storage.get` מחזיר את הערך עצמו, לא `{ value }`', async () => {
    await otzaria().call('storage.set', { key: 'shape-test', value: { mm: 7 } });
    const res = await otzaria().call('storage.get', { key: 'shape-test' });

    expect(res.success).toBe(true);
    // `call()` ב-host/otzaria-client.ts מוסר את `res.data` כמו שהוא, וכל
    // הצרכנים ב-host/settings.ts קוראים ישירות — `raw !== false`,
    // `raw === true`, `Array.isArray(raw)`.
    expect(res.data).toEqual({ mm: 7 });
  });

  it('מפתח שאינו קיים חוזר כ-null ולא כאובייקט', async () => {
    const res = await otzaria().call('storage.get', { key: 'shape-test-missing' });

    expect(res.success).toBe(true);
    // אובייקט ריק כאן היה עובר את השומר של „אין ערך” ב-`loadSetting`
    // (`raw === null || raw === undefined`) ומגיע ל-`parse` — ומי שה-parse
    // שלו מרשני (`raw => raw`, ShulchanTab.vue) היה מקבל אותו כערך אמיתי.
    expect(res.data).toBeNull();
  });
});
