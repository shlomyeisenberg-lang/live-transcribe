import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';
import type { Plugin } from 'vite';

/**
 * SuperDoc 2.23 holds typed-caret validation behind the 6.5s catalog idle gate.
 * That also holds activeChangeIds, disabling Accept/Reject after tracked input.
 * Only the three selection guards get a shorter idle window; catalog reads and
 * the active/pending foreground-mutation guards remain unchanged.
 *
 * Temporary upstream workaround: remove on an engine upgrade that fixes this.
 * Applied in memory for dev/build, without changing installed dependency files.
 */
export function patchSelectionIdle(source: string): string {
  const declaration = 'const HEAVY_READ_IDLE_MS = 6500;';
  const guard = 'typedCaretValidationHeld && Date.now() - lastTypingMutationAtMs < HEAVY_READ_IDLE_MS';
  if (source.split(declaration).length !== 2 || source.split(guard).length !== 4) {
    throw new Error('SuperDoc selection-idle patch no longer matches: review the upstream workaround before upgrading.');
  }
  return source
    .replace(declaration, `${declaration}\nconst TYPED_CARET_IDLE_MS = 500;`)
    .split(guard).join('typedCaretValidationHeld && Date.now() - lastTypingMutationAtMs < TYPED_CARET_IDLE_MS');
}

export function superdocSelectionIdle(): Plugin {
  const chunkName = /^create-super-doc-ui-[\w-]+\.es\.js$/;
  let root = process.cwd();
  return {
    name: 'superdoc-selection-idle',
    enforce: 'pre',
    configResolved(config) {
      root = config.root;
    },
    buildStart() {
      const dir = join(root, 'node_modules/superdoc');
      const { version } = JSON.parse(readFileSync(join(dir, 'package.json'), 'utf8'));
      if (version !== '2.23.0') {
        throw new Error(`Review the SuperDoc selection-idle workaround for version ${version}.`);
      }
      const chunks = join(dir, 'dist/chunks');
      const matches = readdirSync(chunks).filter((name) => chunkName.test(name));
      if (matches.length !== 1) throw new Error('Cannot locate the SuperDoc UI chunk for the selection-idle workaround.');
      // Fail early even if a future entry point stops importing this chunk.
      patchSelectionIdle(readFileSync(join(chunks, matches[0]), 'utf8'));
    },
    transform(source, id) {
      const path = id.split('\\').join('/').split('?')[0];
      if (!/\/superdoc\/dist\/chunks\/create-super-doc-ui-[\w-]+\.es\.js$/.test(path)) return;
      return { code: patchSelectionIdle(source), map: null };
    },
  };
}
