/**
 * האדפטר הוא המקום היחיד שמתרגם כשל של פקודה להודעה בעברית. הבדיקות כאן
 * מכסות את ארבעת מסלולי הכשל שהמנוע מבחין ביניהם, כי הבלבול ביניהם הוא
 * שמייצר "הכפתור לא עושה כלום": false עם reason, receipt כושל, זריקה, ומזהה
 * שאינו מוכר.
 */
import { describe, it, expect, vi } from 'vitest';
import type { BorrowedSuperDocUI } from 'superdoc';
import type { CommandExecutionResult, CommandState } from 'superdoc/ui';
import { createCommandAdapter } from '../../src/engine/command-adapter';

interface FakeCommand {
  state?: Partial<CommandState>;
  result?: CommandExecutionResult | (() => Promise<CommandExecutionResult>);
}

function fakeUi(commands: Record<string, FakeCommand>): {
  ui: BorrowedSuperDocUI;
  executed: string[];
  observers: Array<(state: CommandState) => void>;
} {
  const executed: string[] = [];
  const observers: Array<(state: CommandState) => void> = [];

  const baseState: CommandState = {
    enabled: true,
    active: false,
    supported: true,
    source: 'builtin',
  };

  const ui = {
    commands: {
      has: (id: string) => id in commands,
      get: (id: string) => ({
        id,
        getState: (): CommandState => ({ ...baseState, ...commands[id]?.state }),
        observe: (listener: (state: CommandState) => void) => {
          observers.push(listener);
          return () => {
            observers.splice(observers.indexOf(listener), 1);
          };
        },
        execute: () => true,
        executeAsync: async () => true,
      }),
      async executeAsync(id: string): Promise<CommandExecutionResult> {
        executed.push(id);
        const result = commands[id]?.result;
        if (typeof result === 'function') return result();
        return result ?? true;
      },
    },
  } as unknown as BorrowedSuperDocUI;

  return { ui, executed, observers };
}

describe('createCommandAdapter', () => {
  it('פקודה שהצליחה מחזירה ok', async () => {
    const { ui, executed } = fakeUi({ bold: { result: true } });

    await expect(createCommandAdapter(ui).run('bold')).resolves.toEqual({ ok: true });
    expect(executed).toEqual(['bold']);
  });

  it('receipt מוצלח נחשב הצלחה', async () => {
    const { ui } = fakeUi({ 'table-insert': { result: { success: true } as never } });

    await expect(createCommandAdapter(ui).run('table-insert')).resolves.toEqual({ ok: true });
  });

  it('false מתורגם לפי ה-reason של מצב הפקד', async () => {
    const { ui } = fakeUi({
      'table-delete-row': { result: false, state: { reason: 'table-context-unavailable' } },
    });

    await expect(createCommandAdapter(ui).run('table-delete-row')).resolves.toEqual({
      ok: false,
      message: 'יש למקם את הסמן בתוך תא בטבלה',
      reason: 'table-context-unavailable',
    });
  });

  it('false בלי reason אינו נעלם בשקט', async () => {
    const { ui } = fakeUi({ bold: { result: false } });

    const outcome = await createCommandAdapter(ui).run('bold');

    expect(outcome).toEqual({ ok: false, message: 'הפעולה נכשלה', reason: undefined });
  });

  it('reason שאינו מוכר מוצג עם הקוד שלו', async () => {
    const { ui } = fakeUi({ bold: { result: false, state: { reason: 'future-reason' as never } } });

    const outcome = await createCommandAdapter(ui).run('bold');

    expect(outcome.ok).toBe(false);
    expect(outcome.ok === false && outcome.message).toContain('future-reason');
  });

  it('receipt עם NO_OP נחשב הצלחה שקטה, לא כשל', async () => {
    // באג #9 בסקר הפקדים: "הקטן הזחה" על הזחה אפס ו"RTL" על פסקה שכבר RTL
    // עוברים כאן בדיוק (ui.commands.executeAsync), והמנוע מחזיר NO_OP כשהערך
    // המבוקש כבר קיים. ב-Word זו אינה שגיאה, וכך גם מודולי ה-Document API
    // הישיר (page-setup.ts ואחרים) כבר מכריעים.
    const { ui } = fakeUi({
      'indent-decrease': {
        result: { success: false, failure: { code: 'NO_OP', message: 'no change' } } as never,
      },
    });

    await expect(createCommandAdapter(ui).run('indent-decrease')).resolves.toEqual({ ok: true });
  });

  it('receipt כושל מתורגם לפי קוד הכשל', async () => {
    const { ui } = fakeUi({
      link: {
        result: {
          success: false,
          failure: { code: 'DOCUMENT_READONLY', message: 'read only' },
        } as never,
      },
    });

    await expect(createCommandAdapter(ui).run('link')).resolves.toEqual({
      ok: false,
      message: 'המסמך פתוח לקריאה בלבד',
      reason: 'DOCUMENT_READONLY',
    });
  });

  it('קוד כשל שאין לו תרגום מוצג עם ההסבר של המנוע ועם הקוד', async () => {
    const { ui } = fakeUi({
      link: {
        result: {
          success: false,
          failure: { code: 'PLAN_CONFLICT_OVERLAP', message: 'overlapping plan steps' },
        } as never,
      },
    });

    const outcome = await createCommandAdapter(ui).run('link');

    expect(outcome.ok === false && outcome.message).toBe(
      'הפעולה נכשלה: overlapping plan steps (PLAN_CONFLICT_OVERLAP)',
    );
  });

  it('קוד כשל בלי הסבר מוצג עם הקוד לבדו', async () => {
    const { ui } = fakeUi({
      link: { result: { success: false, failure: { code: 'PLAN_CONFLICT_OVERLAP' } } as never },
    });

    const outcome = await createCommandAdapter(ui).run('link');

    expect(outcome.ok === false && outcome.message).toBe('הפעולה נכשלה (PLAN_CONFLICT_OVERLAP)');
  });

  it('זריקה מן המנוע נתפסת ומוחזרת כהודעה', async () => {
    const { ui } = fakeUi({
      bold: {
        result: () => Promise.reject(new Error('המנוע קרס')),
      },
    });

    await expect(createCommandAdapter(ui).run('bold')).resolves.toEqual({
      ok: false,
      message: 'המנוע קרס',
      reason: 'threw',
    });
  });

  it('מזהה שאינו מוכר למנוע אינו מורץ', async () => {
    const { ui, executed } = fakeUi({});

    const outcome = await createCommandAdapter(ui).run('no-such');

    expect(outcome).toEqual({
      ok: false,
      message: 'הפעולה no-such אינה מוכרת למנוע',
      reason: 'unknown-command',
    });
    expect(executed).toEqual([]);
  });

  it('observe מחזיר ביטול שמסיר את המאזין', () => {
    const { ui, observers } = fakeUi({ bold: {} });
    const listener = vi.fn();

    const stop = createCommandAdapter(ui).observe('bold', listener);
    expect(observers).toHaveLength(1);

    stop();
    expect(observers).toHaveLength(0);
  });
});
