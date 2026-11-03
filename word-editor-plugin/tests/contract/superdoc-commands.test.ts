/**
 * שלב 0.2 בתכנית: בדיקת חוזה מול superdoc עצמו, בגרסה שהתוסף נועל.
 *
 * הבדיקה מריצה את ה-registry שלנו מול הקטלוג האמיתי של המנוע. מזהה שהמנוע
 * לא מכיר, או שהוא מסמן unsupported, מפיל את הבדיקה — במקום להפוך לכפתור שלא
 * עושה כלום. שדרוג גרסת superdoc שמשנה מזהה ייתפס כאן ולא אצל המשתמש.
 *
 * כאן — ורק כאן — מותר לקרוא ל-createSuperDocUI: הבדיקה בונה controller על
 * host מבני ריק, מחזיקה בו בבעלות ומפרקת אותו. קוד התוסף משתמש ב-superdoc.ui
 * המושאל (ראו create-editor.ts). הקטלוג סטטי, ולכן הוא זמין גם בלי להרים את
 * מנוע ה-DOCX — מה שמאפשר להריץ את הבדיקה הזאת ב-CI בלי דפדפן אמיתי.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { createSuperDocUI, BUILT_IN_COMMAND_IDS } from 'superdoc/ui';
import type { SuperDocUI } from 'superdoc/ui';
import {
  COMMAND_GROUPS,
  COMMAND_IDS,
  KNOWN_UNSUPPORTED_COMMANDS,
  inspectCommandSupport,
} from '../../src/engine/capabilities';

let ui: SuperDocUI;

beforeAll(() => {
  ui = createSuperDocUI({ superdoc: {} });
});

afterAll(() => {
  ui.destroy();
});

describe('registry הפקודות מול המנוע', () => {
  it('כל מזהה ב-registry מוכר למנוע ומנותב', () => {
    const report = inspectCommandSupport(ui);

    expect(report.unknown).toEqual([]);
    expect(report.unsupported).toEqual([]);
    expect(report.routed).toHaveLength(COMMAND_IDS.length);
  });

  it('אין מזהה כפול, וכל תחום אינו ריק', () => {
    expect(new Set(COMMAND_IDS).size).toBe(COMMAND_IDS.length);

    for (const [group, ids] of Object.entries(COMMAND_GROUPS)) {
      expect(ids.length, `התחום ${group} ריק`).toBeGreaterThan(0);
    }
  });

  it('פקודה שהמנוע מסמן unsupported אינה ב-registry', () => {
    for (const id of Object.keys(KNOWN_UNSUPPORTED_COMMANDS)) {
      expect(COMMAND_IDS).not.toContain(id);

      // has() מחזיר true גם ל-unsupported — המזהה בקטלוג. ההבחנה היא ב-source,
      // וזה מה שמצדיק את ההוצאה מה-registry.
      expect(ui.commands.has(id)).toBe(true);
      expect(ui.commands.get(id).getState().source).toBe('unsupported');
    }
  });

  it('מזהה שאינו בקטלוג מדווח כלא מוכר', () => {
    expect(ui.commands.has('otzaria-no-such-command')).toBe(false);
    expect(ui.commands.get('otzaria-no-such-command').getState().source).toBe('unsupported');
  });

  it('כל המזהים הקנוניים של superdoc מוכרים למנוע', () => {
    for (const id of Object.values(BUILT_IN_COMMAND_IDS)) {
      expect(ui.commands.has(id), id).toBe(true);
    }
  });

  it('ה-registry מכסה את המזהים הקנוניים שאינם aliases', () => {
    // setFontFamily/setFontSize הם aliases ל-font-family/font-size ולכן אינם
    // ב-registry: פקד אחד לכל פעולה.
    const aliases = new Set(['setFontFamily', 'setFontSize']);
    const missing = Object.values(BUILT_IN_COMMAND_IDS)
      .filter((id) => !aliases.has(id))
      .filter((id) => !COMMAND_IDS.includes(id as (typeof COMMAND_IDS)[number]));

    expect(missing).toEqual([]);
  });
});
