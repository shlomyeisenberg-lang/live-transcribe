import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createTextLocator, locateText } from '../../src/engine/painted-lines';

const originalRects = Object.getOwnPropertyDescriptor(Range.prototype, 'getClientRects');
const box = { left: 0, right: 10, top: 0, bottom: 10, width: 10, height: 10 };
let root: HTMLElement;

beforeEach(() => {
  root = document.createElement('div');
  document.body.append(root);
  Object.defineProperty(Range.prototype, 'getClientRects', { configurable: true, value: vi.fn(() => [box]) });
});
afterEach(() => {
  root.remove();
  if (originalRects) Object.defineProperty(Range.prototype, 'getClientRects', originalRects);
  else Reflect.deleteProperty(Range.prototype, 'getClientRects');
  vi.restoreAllMocks();
});

describe('autocorrect text location', () => {
  function splitWord(text: string, secondStart = 2) {
    root.innerHTML = `<div data-source-node-id="b" data-pm-start="1" data-pm-end="5"><span data-pm-start="1" data-pm-end="2">א</span><span data-pm-start="${secondStart}" data-pm-end="${secondStart + text.length}">${text}</span></div>`;
  }

  it('requires the whole replacement to match across formatting runs', () => {
    splitWord('בד ');
    expect(locateText('b', 0, 'אבג', root, { before: '', after: ' ' })).toBeNull();
  });

  it('locates all parts of a matching replacement across formatting runs', () => {
    splitWord('בג ');
    const found = locateText('b', 0, 'אבג', root, { before: '', after: ' ' });
    expect(found?.start).toBe(0);
    expect(found?.rects).toHaveLength(2);
  });

  it('does not join text across a non-text PM gap', () => {
    splitWord('בג ', 3);
    expect(locateText('b', 0, 'אבג', root)).toBeNull();
  });

  it('reads the block once for all indicators in a scan', () => {
    splitWord('בג ');
    const fragment = root.firstElementChild!;
    const roots = vi.spyOn(root, 'querySelectorAll');
    const runs = vi.spyOn(fragment, 'querySelectorAll');
    const locate = createTextLocator('b', root);
    for (let i = 0; i < 100; i += 1) expect(locate(0, 'אבג')).not.toBeNull();
    expect(roots).toHaveBeenCalledTimes(1);
    expect(runs).toHaveBeenCalledTimes(1);
  });

  it('a new scan sees the current DOM after an edit', () => {
    splitWord('בג ');
    expect(createTextLocator('b', root)(0, 'אבג')).not.toBeNull();
    splitWord('בד ');
    expect(createTextLocator('b', root)(0, 'אבג')).toBeNull();
  });

  it('does not identify another identical word at the recorded offset', () => {
    root.innerHTML = '<div data-source-node-id="b" data-pm-start="1" data-pm-end="11"><span data-pm-start="1" data-pm-end="11">zzX X X X </span></div>';
    // Before inserting zz, the last X was at6 and had this context.
    // Current offset6 is the third X, and offset8 is the real last X.
    expect(locateText('b', 6, 'X', root, { before: ' X ', after: ' ' })).toBeNull();
  });

  it('follows a moved word only when its context identifies one occurrence', () => {
    root.innerHTML = '<div data-source-node-id="b" data-pm-start="1" data-pm-end="11"><span data-pm-start="1" data-pm-end="11">zzA X B Y </span></div>';
    expect(locateText('b', 2, 'X', root, { before: 'A ', after: ' B ' })?.start).toBe(4);
  });
});
