# Estimates — S02_externalise_hardcoded_paths

Generated 2026-06-12T12:24:17Z · estimator level factor 0.75 · canonical record: `estimates.json`

| AI | status | complexity (1-10) | est. tokens | executing model | est. seconds | estimator model |
|---|---|---|---|---|---|---|
| claude | ok | 7 | 320000 | claude-opus-4-8 | 1200 | claude-opus-4-8 |
| grok | ok | 5 | 105000 | grok-build | 540 | grok-build |
| codex | ok | 6 | 75000 | gpt-5.4 | 4200 | gpt-5.5 |

## claude — notes

Wide but mechanical refactor across a ~3,400-line const-heavy Main.cs: extract one user-specific path into App.config/typed options and consolidate dozens of c*Col/c*Row/cMostPA*/cIIandGOV* constants into a WorkbookLayout class, then rewire every reference. Token cost is driven by reading the large file plus multiple edit passes; complexity is raised by regression risk if any constant is missed or remapped. Key risk: the 'smoke run against a sample appraisal folder' Done-when cannot be verified by the agent (needs Windows+Office Interop), so build correctness must be argued statically. Estimate assumes specification.md §4 gives a clear template contract to document against.

## grok — notes

Assumptions: dozens means approximately 30-60 constants scattered across the ~3400-line Main.cs, primarily in one globals region, with references throughout UI handlers and logic; paths are few but critical; App.config exists and will be extended with <appSettings> or similar, and a simple POCO or static class suffices for WorkbookLayout. Risks: constants may have subtle dependencies or be used in calculations/Excel interop calls that could break if extraction isn't exact; full verification requires building the .NET 4.8 project and running against sample data which may expose environment-specific issues. Estimate driven by the need for large context (full or near-full file reads multiple times for accurate global search/replace and to preserve semantics), plus iterative build/smoke/test-fix cycles and adding contract documentation per spec.md §4.

## codex — notes

This is a behavior-preserving refactor in a large legacy WinForms Main.cs, with many constants and references that must be moved without changing workbook behavior. The main effort is inventorying all path/layout literals, introducing a typed WorkbookLayout/App.config pattern consistent with .NET Framework 4.8, and updating usages safely. Build and smoke-run validation may be slowed by Office Interop and sample appraisal folder availability.

