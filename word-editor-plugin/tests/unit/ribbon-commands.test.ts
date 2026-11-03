/**
 * חיבור פקדי ה-Ribbon לפקודות המנוע.
 *
 * מה שהיה בקובץ הזה, ולמה הוא נמחק: שמונה בדיקות שהריצו payloads מול mock
 * שכל תפקידו היה `calls.push({ id, payload }); return true;`, ואחר כך השוו
 * את `calls` לאותם payloads. ה-docblock הבטיח שהן מאמתות „העברה במדויק...
 * למנוע SuperDoc”, אבל mock אינו מנוע: ארבעה מהם — `{ fontFamily }`,
 * `{ fontSize }`, `{ color }` ו-`{ zoom }` — נדחים בשקט על ידי superdoc,
 * והבדיקה אישרה אותם בירוק. היא לא בדקה חוזה אלא ש-JavaScript מעביר
 * ארגומנטים.
 *
 * במקומן: tests/contract/command-payloads.test.ts, שמריץ את ה-payloads מול
 * הוולידטורים האמיתיים של החבילה. כאן נשאר מה שאותה בדיקה **אינה** יכולה
 * לכסות — שני חיבורים בין הקוד שלנו לבין החוזה הזה.
 */
import { describe, it, expect } from 'vitest';
import { readFileSync, readdirSync } from 'node:fs';
import { join, relative } from 'node:path';
import type { BorrowedSuperDocUI } from 'superdoc';
import type { CommandExecutionResult, CommandState } from 'superdoc/ui';
import { createCommandAdapter } from '../../src/engine/command-adapter';
import { COMMAND_IDS } from '../../src/engine/capabilities';

// vitest רץ משורש המאגר, ולכן cwd הוא השורש.
const SRC = join(process.cwd(), 'src');

function sourceFiles(dir = SRC): string[] {
  const files: string[] = [];
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) files.push(...sourceFiles(full));
    else if (/\.(ts|vue)$/.test(entry.name) && !entry.name.endsWith('.d.ts')) files.push(full);
  }
  return files;
}

/** הערות מוסרות: הן מתעדות במפורש את ה-payloads השגויים, וזה אינו הפרה. */
function stripComments(text: string): string {
  return text
    .replace(/\/\*[\s\S]*?\*\//g, (block) => block.replace(/[^\n]/g, ' '))
    .replace(/\/\/.*$/gm, '')
    .replace(/<!--[\s\S]*?-->/g, (block) => block.replace(/[^\n]/g, ' '));
}

const sources = sourceFiles().map((path) => ({
  path: relative(SRC, path),
  text: stripComments(readFileSync(path, 'utf8')),
}));

/** התאמות בפורמט "נתיב:שורה", כדי שכשל יצביע למקום ולא רק לכלל. */
function hits(pattern: RegExp): string[] {
  const found: string[] = [];
  for (const { path, text } of sources) {
    text.split('\n').forEach((line, index) => {
      if (pattern.test(line)) found.push(`${path}:${index + 1}`);
    });
  }
  return found;
}

describe('מזהי הפקודות שפקדי ה-Ribbon מבקשים', () => {
  it('יש קבצי מקור לבדוק', () => {
    expect(sources.length).toBeGreaterThan(0);
  });

  it('כל `useCommand` מבקש מזהה שנמצא ב-registry', () => {
    // המזהים ב-registry מאומתים מול הקטלוג האמיתי של המנוע
    // (tests/contract/superdoc-commands.test.ts). מזהה שמוקלד שגוי באתר
    // הקריאה אינו מגיע לשם, ולכן היה מייצר פקד שנראה תקין ולא עושה כלום.
    const requested = new Set<string>();
    for (const { text } of sources) {
      for (const match of text.matchAll(/useCommand\(\s*'([^']+)'/g)) requested.add(match[1]);
    }

    expect(requested.size).toBeGreaterThan(0);
    const unknown = [...requested].filter((id) => !COMMAND_IDS.includes(id as never));
    expect(unknown).toEqual([]);
  });
});

describe('חוזה ה-payload באתרי הקריאה', () => {
  it('אין פקד שבונה payload לפי שם השדה של הפקודה', () => {
    // הרגרסיה שהקובץ הזה לא תפס: `fontFamily`, `fontSize`, `color` ו-`zoom`
    // אינם מפתחות שהמנוע מפרק (`unwrapScalar` מכיר `value`, `alignment`,
    // `lineHeight`, `style` — לא אותם), ולכן payload כזה נכשל סגור. בניית
    // ה-payload היא ב-engine/payloads.ts ונבדקת מול המנוע עצמו.
    expect(hits(/\.run\(\s*\{\s*(fontFamily|fontSize|color|zoom)\b/)).toEqual([]);
    expect(hits(/run\(\s*'[^']+'\s*,\s*\{\s*(fontFamily|fontSize|color|zoom)\b/)).toEqual([]);
  });
});

describe('createCommandAdapter', () => {
  it('מעביר את ה-payload למנוע בלי לשנות אותו', () => {
    // זה כל מה שהאדפטר מבטיח לגבי payload, וזו הסיבה שעיצוב ה-payload חייב
    // להיות במקום אחד למעלה: אין שכבה שמתקנת אותו בדרך.
    const calls: Array<{ id: string; payload?: unknown }> = [];
    const ui = {
      commands: {
        has: () => true,
        get: () => ({
          getState: (): CommandState => ({
            enabled: true,
            active: false,
            supported: true,
            source: 'builtin',
          }),
          observe: () => () => {},
        }),
        async executeAsync(id: string, payload?: unknown): Promise<CommandExecutionResult> {
          calls.push({ id, payload });
          return true;
        },
      },
    } as unknown as BorrowedSuperDocUI;

    const adapter = createCommandAdapter(ui);
    const payload = { value: '#0055FF' };
    void adapter.run('text-color', payload);
    void adapter.run('bold');

    expect(calls[0].payload).toBe(payload);
    expect(calls[1].payload).toBeUndefined();
  });
});
