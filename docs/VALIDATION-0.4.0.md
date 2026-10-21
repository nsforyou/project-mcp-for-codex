# 0.4.0 validation / 测试记录

Date: 2026-10-21. Windows, 64-bit Python 3.13.7, desktop Microsoft Project 16.0, PyMuPDF and local Java/MPXJ runtime.

## Automated coverage

- 25 unit/protocol tests: existing live/offline safeguards and notes, plus deterministic read-only layout preview, source/settings/output hash binding, refusal to overwrite, malformed profiles, lock/process refusal, optional PDF checks, UTF-8 stdio initialization and all 30 tool schemas.
- Native A/B fixture: A3 landscape/day scale, A4 landscape/day scale and A4 portrait/week scale.
- A separate disposable 33-task/2-resource synthetic plan: A3/A4 landscape/day scale, including completion values 0/25/50/75/100%, a long Chinese name, dependencies and a milestone.
- Native exports compare tasks, resources and assignments before layout and after reopening the saved copy; source SHA256 remains unchanged.
- Re-export after reopening checks that page structure and extracted text remain the same, including inside-bar progress labels.
- PDF percentage counts match table plus non-milestone bar percentages for the accepted test profiles. Final pages were rendered and inspected; low-resolution renderers can omit thin glyph details, so detailed review used higher-resolution rendering.
- Existing native integration smoke test passed: 21 create checks, 2 write checks, structured long-note round trip, preserved view streams and exact SHA256 rollback.

Commands:

```powershell
python -m unittest discover -s mcp/tests -p 'test_*.py'
python mcp/tests/print_integration.py --extended
python mcp/tests/print_integration.py --file PATH_TO_DISPOSABLE_SYNTHETIC_PLAN.mpp
python mcp/tests/integration_smoke.py
```

Outputs, rendered pages, COM logs and external synthetic fixtures stay local in ignored folders. No business plans or local configuration are distributed. An integration worktree used the existing server's shared lock and never restarted the production MCP process.

## Limits and observed edge cases

This validates the tested local setup and synthetic fixtures, not every Project document or language. The standard English/Simplified Chinese normal/manual bar style names are supported. Font display names are installation-dependent; the default keeps the existing family.

With the larger synthetic plan compressed into A4 portrait/week scale, some short bars cannot display full percentage text. The PDF exporter therefore reports `exported_needs_visual_review`, and emits a warning when an available full-range percentage-count heuristic disagrees with expected counts. The heuristic is not a completeness or clipping proof; date-window exports and custom styles still require manual review.

An A3 final page can have substantial whitespace. The exporter does not automatically rebalance pages or reduce text below the configured limit to force a single page. It does not support linked/inserted subprojects, physical printer submission or print-completion feedback. Drawing/OLE/custom view preservation is not certified by the task-data checks.

No new native live-edit integration was needed for this release: live write implementation is unchanged. Its existing unit safeguards passed; the existing offline native regression passed. Tests did not operate a user's active Project window.
