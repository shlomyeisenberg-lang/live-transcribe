/** Paste QA entrypoint. Covers supported collapsed/blank destinations, HTML
 * script emphasis, keep-source baseline, and pre-mutation refusals for ranges
 * and split fonts. Native clipboardData is produced through CDP's documented
 * editing command; keydown interception and macOS Ctrl+V fallback are also
 * covered by focused unit tests. Run with a freshly built SuperDoc 2.20 bundle.
 */
await import('./paste-destination-qa.mjs');
