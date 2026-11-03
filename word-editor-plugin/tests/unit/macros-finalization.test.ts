import { beforeEach, describe, expect, it, vi } from 'vitest';
import type { SuperDoc } from 'superdoc';
import { DEFAULT_STORAGE_KEY, type PersistedMacroState } from 'superdoc-macros';
import {
  MACRO_STATUS,
  SEEDED_FLAG_KEY,
  installMacros,
  type MacrosSession,
} from '../../src/engine/macros';

interface FakeEngine {
  ui: {
    commands: {
      has(id: string): boolean;
      get(id: string): { getState(): { reason?: string } };
      executeAsync(id: string, payload?: unknown): Promise<unknown>;
    };
  };
  activeEditor: null;
}

function createFakeSession(): { session: MacrosSession; engine: FakeEngine } {
  const engine: FakeEngine = {
    ui: {
      commands: {
        has: (id) => id === 'bold',
        get: () => ({ getState: () => ({}) }),
        executeAsync: async () => ({ success: true }),
      },
    },
    activeEditor: null,
  };
  return { session: { superdoc: engine as unknown as SuperDoc }, engine };
}

function typeText(container: HTMLElement, text: string): void {
  container.dispatchEvent(
    new InputEvent('beforeinput', { inputType: 'insertText', data: text, bubbles: true }),
  );
}

async function flushAsyncWork(): Promise<void> {
  await new Promise((resolve) => setTimeout(resolve, 0));
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function lastStatus(statuses: ReadonlyArray<{ message: string; isError?: boolean }>) {
  return statuses[statuses.length - 1];
}

describe('installMacros — recording finalization', () => {
  beforeEach(() => {
    vi.restoreAllMocks();
    localStorage.clear();
  });

  it('reports and discards a capture containing only an uncapturable image command', async () => {
    const { session, engine } = createFakeSession();
    const statuses: Array<{ message: string; isError?: boolean }> = [];
    let confirmations = 0;
    const handle = installMacros(
      session,
      document.createElement('div'),
      (message, isError) => statuses.push({ message, isError }),
      {
        confirmIncomplete: async () => {
          confirmations += 1;
          return true;
        },
      },
    );

    handle.toggleRecording();
    expect(handle.recording.value).toBe(true);
    await engine.ui.commands.executeAsync('bold', { src: 'x'.repeat(20_000) });
    await flushAsyncWork();
    handle.toggleRecording();
    await flushAsyncWork();

    expect(confirmations).toBe(0);
    expect(handle.kit.listRecordings()).toHaveLength(0);
    expect(handle.kit.hasPendingRecording).toBe(false);
    expect(lastStatus(statuses)).toMatchObject({ isError: true });
    expect(lastStatus(statuses)?.message).toContain('bold');
    handle.dispose();
  });

  it('retains an incomplete capture when storage fails after consent, then saves it on retry', async () => {
    const { session, engine } = createFakeSession();
    const container = document.createElement('div');
    const statuses: Array<{ message: string; isError?: boolean }> = [];
    let confirmations = 0;
    const handle = installMacros(
      session,
      container,
      (message, isError) => statuses.push({ message, isError }),
      {
        confirmIncomplete: async () => {
          confirmations += 1;
          return true;
        },
      },
    );

    handle.toggleRecording();
    typeText(container, 'א');
    await engine.ui.commands.executeAsync('bold', { src: 'x'.repeat(20_000) });
    await flushAsyncWork();

    const blocked = vi.spyOn(Storage.prototype, 'setItem').mockImplementation(() => {
      throw new DOMException('quota', 'QuotaExceededError');
    });
    handle.toggleRecording();
    await flushAsyncWork();

    expect(confirmations).toBe(1);
    expect(handle.kit.listRecordings()).toHaveLength(0);
    expect(handle.kit.hasPendingRecording).toBe(true);
    expect(lastStatus(statuses)?.message).toContain(MACRO_STATUS.recordingKept);

    blocked.mockRestore();
    handle.toggleRecording();
    await flushAsyncWork();

    expect(confirmations).toBe(1); // consent stays attached to the pending snapshot
    expect(handle.kit.hasPendingRecording).toBe(false);
    expect(handle.kit.listRecordings()).toHaveLength(1);
    expect(handle.kit.listRecordings()[0]?.steps).toEqual([
      { type: 'insert-text', text: 'א' },
    ]);
    handle.dispose();
  });

  it('keeps a stopped capture pending when the saved-recording list is full', async () => {
    const fullState: PersistedMacroState = {
      version: 1,
      scripts: [],
      snippets: [],
      recordings: Array.from({ length: 500 }, (_value, index) => ({
        version: 1 as const,
        id: `r-${index}`,
        name: `recording ${index}`,
        steps: [{ type: 'insert-text' as const, text: 'x' }],
      })),
    };
    localStorage.setItem(DEFAULT_STORAGE_KEY, JSON.stringify(fullState));
    localStorage.setItem(SEEDED_FLAG_KEY, 'yes');

    const { session } = createFakeSession();
    const container = document.createElement('div');
    const statuses: Array<{ message: string; isError?: boolean }> = [];
    const handle = installMacros(session, container, (message, isError) =>
      statuses.push({ message, isError }),
    );

    handle.toggleRecording();
    typeText(container, 'חדש');
    handle.toggleRecording();
    await flushAsyncWork();

    expect(handle.recording.value).toBe(false);
    expect(handle.kit.hasPendingRecording).toBe(true);
    expect(handle.kit.listRecordings()).toHaveLength(500);
    expect(lastStatus(statuses)?.message).toContain(MACRO_STATUS.recordingKept);

    handle.kit.removeRecording('r-0');
    handle.toggleRecording();
    await flushAsyncWork();

    const recordings = handle.kit.listRecordings();
    expect(handle.kit.hasPendingRecording).toBe(false);
    expect(recordings).toHaveLength(500);
    expect(recordings[recordings.length - 1]?.steps).toEqual([
      { type: 'insert-text', text: 'חדש' },
    ]);
    handle.dispose();
  });

  it('blocks a second finalization while the host confirmation is still open', async () => {
    const { session, engine } = createFakeSession();
    const container = document.createElement('div');
    const confirmation = { resolve: null as ((value: boolean) => void) | null };
    let confirmations = 0;
    const handle = installMacros(session, container, () => undefined, {
      confirmIncomplete: () => {
        confirmations += 1;
        return new Promise<boolean>((resolve) => {
          confirmation.resolve = resolve;
        });
      },
    });

    handle.toggleRecording();
    typeText(container, 'א');
    await engine.ui.commands.executeAsync('bold', { src: 'x'.repeat(20_000) });
    await flushAsyncWork();

    handle.toggleRecording();
    handle.toggleRecording();
    expect(confirmations).toBe(1);

    const completeConfirmation = confirmation.resolve;
    if (!completeConfirmation) throw new Error('confirmation was not opened');
    completeConfirmation(false);
    await flushAsyncWork();
    expect(handle.kit.hasPendingRecording).toBe(false);
    expect(handle.kit.listRecordings()).toHaveLength(0);
    handle.dispose();
  });

  it('retains the stopped capture when the host confirmation itself fails', async () => {
    const { session, engine } = createFakeSession();
    const container = document.createElement('div');
    const statuses: Array<{ message: string; isError?: boolean }> = [];
    const handle = installMacros(
      session,
      container,
      (message, isError) => statuses.push({ message, isError }),
      {
        confirmIncomplete: async () => {
          throw new Error('confirm unavailable');
        },
      },
    );

    handle.toggleRecording();
    typeText(container, 'א');
    await engine.ui.commands.executeAsync('bold', { src: 'x'.repeat(20_000) });
    await flushAsyncWork();
    handle.toggleRecording();
    await flushAsyncWork();

    expect(handle.kit.hasPendingRecording).toBe(true);
    expect(lastStatus(statuses)?.message).toContain('confirm unavailable');
    expect(lastStatus(statuses)?.message).toContain(MACRO_STATUS.recordingKept);
    handle.dispose();
  });

  it('preserves readable macros when storage becomes unwritable during default seeding', () => {
    const existing: PersistedMacroState = {
      version: 1,
      scripts: [],
      snippets: [],
      recordings: [
        {
          version: 1,
          id: 'important',
          name: 'חשוב',
          steps: [{ type: 'insert-text', text: 'נשמר' }],
        },
      ],
    };
    localStorage.setItem(DEFAULT_STORAGE_KEY, JSON.stringify(existing));

    const originalSetItem = Storage.prototype.setItem;
    vi.spyOn(Storage.prototype, 'setItem').mockImplementation(function (
      this: Storage,
      key: string,
      value: string,
    ) {
      // The large macro-state write fails, but the tiny seeded marker would
      // still succeed. A memory-only seed must deliberately avoid setting it.
      if (key === DEFAULT_STORAGE_KEY) {
        throw new DOMException('quota', 'QuotaExceededError');
      }
      return originalSetItem.call(this, key, value);
    });

    const { session } = createFakeSession();
    const statuses: Array<{ message: string; isError?: boolean }> = [];
    const handle = installMacros(
      session,
      document.createElement('div'),
      (message, isError) => statuses.push({ message, isError }),
    );

    expect(handle.kit.listRecordings().map((item) => item.id)).toEqual(['important']);
    expect(handle.kit.listSnippets()).toHaveLength(1);
    expect(localStorage.getItem(SEEDED_FLAG_KEY)).toBeNull();
    expect(statuses).toContainEqual({ message: MACRO_STATUS.storageUnavailable, isError: true });
    handle.dispose();
  });
});
