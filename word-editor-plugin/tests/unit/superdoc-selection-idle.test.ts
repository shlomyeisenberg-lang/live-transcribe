import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import { describe, expect, it } from 'vitest';
import { patchSelectionIdle } from '../../scripts/superdoc-selection-idle';

const chunks = join(process.cwd(), 'node_modules/superdoc/dist/chunks');
const name = readdirSync(chunks).find((entry) => /^create-super-doc-ui-[\w-]+\.es\.js$/.test(entry));
if (!name) throw new Error('SuperDoc UI chunk not found');
const source = readFileSync(join(chunks, name), 'utf8');

describe('SuperDoc typed-caret idle workaround', () => {
  it('patches the shipped engine while retaining every other byte, including heavy-read and mutation guards', () => {
    const patched = patchSelectionIdle(source);
    expect(patched).toContain('const TYPED_CARET_IDLE_MS = 500;');
    expect(patched.match(/lastTypingMutationAtMs < TYPED_CARET_IDLE_MS/g)).toHaveLength(3);
    expect(patched).toContain('if (Date.now() - lastTypingMutationAtMs < HEAVY_READ_IDLE_MS)');
    const restored = patched
      .replace('\nconst TYPED_CARET_IDLE_MS = 500;', '')
      .split('lastTypingMutationAtMs < TYPED_CARET_IDLE_MS').join('lastTypingMutationAtMs < HEAVY_READ_IDLE_MS');
    expect(restored).toBe(source);
  });

  it('requires review when upstream changes or removes one of the selection guards', () => {
    expect(() => patchSelectionIdle(source.replace('typedCaretValidationHeld && Date.now()', 'Date.now()'))).toThrow(/no longer matches/);
    expect(() => patchSelectionIdle(source.replace('6500', '1000'))).toThrow(/no longer matches/);
  });
});
