import { afterEach, describe, expect, it, vi } from 'vitest';
const { load, save } = vi.hoisted(() => ({ load: vi.fn(), save: vi.fn(async () => {}) }));
vi.mock('../../src/host/settings', () => ({ loadPasteDefault: load, savePasteDefault: save }));
import { loadPasteDefault, pasteDefaultMode, setPasteDefault } from '../../src/composables/paste-default';
afterEach(() => { pasteDefaultMode.value = 'keep-source'; vi.clearAllMocks(); });
describe('paste preference boot race', () => {
  it('does not overwrite a user choice with a late stored preference', async () => {
    let resolve!: (value: string) => void;
    load.mockImplementationOnce(() => new Promise<string>((done) => { resolve = done; }));
    const pending = loadPasteDefault();
    setPasteDefault('merge');
    resolve('keep-source');
    await pending;
    expect(pasteDefaultMode.value).toBe('merge');
    expect(save).toHaveBeenCalledWith('merge');
  });
  it('persists an explicit current-default choice while loading another old mode', async () => {
    let resolve!: (value: string) => void;
    load.mockImplementationOnce(() => new Promise<string>((done) => { resolve = done; }));
    const pending = loadPasteDefault();
    setPasteDefault('keep-source');
    resolve('merge');
    await pending;
    expect(pasteDefaultMode.value).toBe('keep-source');
    expect(save).toHaveBeenCalledWith('keep-source');
  });
});
