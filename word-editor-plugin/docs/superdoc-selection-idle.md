# SuperDoc 2.23: typed selection validation

In 2.23.0, the UI uses `HEAVY_READ_IDLE_MS = 6500` both for full-document
catalog reads and for typed-caret validation. The latter supplies
`activeChangeIds`; Accept/Reject remains disabled for about 6.5 seconds after
tracked input unless the caret moves.

`scripts/superdoc-selection-idle.ts` applies a temporary Vite transform to the
installed UI chunk in development and production. It gives the three
`typedCaretValidationHeld` guards their own 500 ms idle threshold. Full-document
reads retain their 6500 ms gate. Active and pending foreground mutations still
prevent selection validation; the patch does not force reads during a mutation.
No dependency files are modified on disk and no install hook is needed.

The dependency is pinned to 2.23.0. The build rejects another version or changed
guard shapes so a future upgrade requires reviewing or removing the workaround.
`superdoc` is excluded from Vite dependency optimization so development uses the
same transform as production.

Validation:

- `tests/unit/superdoc-selection-idle.test.ts` checks the actual installed chunk,
  unchanged heavy-read guards, and rejection of upstream drift.
- `node scripts/qa/review-view-qa.mjs b` types tracked changes and accepts/rejects
  them through the ribbon without moving the caret, then inspects exported DOCX.
  It fails if the command still requires more than three seconds of additional
  waiting after the preceding export. This is not an end-to-end latency budget.
- `npm run check:typing-recalc` checks style-recalculation scope during typing
  in a document with several pages. It is not a general engine benchmark.
