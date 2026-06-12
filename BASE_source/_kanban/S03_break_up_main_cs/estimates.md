# Estimates — S03_break_up_main_cs

Generated 2026-06-12T12:24:59Z · estimator level factor 0.75 · canonical record: `estimates.json`

| AI | status | complexity (1-10) | est. tokens | executing model | est. seconds | estimator model |
|---|---|---|---|---|---|---|
| claude | ok | 9 | 2600000 | claude-opus-4-8 | 5400 | claude-opus-4-8 |
| grok | ok | 8 | 220000 | grok-build | 5400 | grok-build |
| codex | ok | 9 | 850000 | gpt-5.5 | 28800 | gpt-5.5 |

## claude — notes

Driver is a 3,427-line WinForms form with heavy Office Interop that must be re-read in full to safely extract five distinct services behind interfaces without breaking event wiring or Interop lifetimes. Iterative work: identify region seams, define interfaces, move logic, rewire handlers, then stand up a net48 xUnit project and write smoke coverage — each pass re-reads large swaths of Main.cs, inflating tokens. Key risks: Interop COM objects are hard to mock so 'testable behind interfaces' is partial; merge-collision risk during extraction; can't actually build/run (Windows+Office required) so verification is read-only. Assumes incremental S03a-style commits; the test project + one service smoke test is the realistic floor, full delegation the ceiling.

## grok — notes

Key assumptions: regions in Main.cs are sufficiently separable per specification.md §2 (with existing Data/* and Excel/* helpers reusable), interop can be isolated behind small interfaces without full COM shims, only smoke tests for one service are needed rather than full coverage, and incremental commits (S03a etc.) can be estimated as one coherent body of work. Risks: hidden shared mutable state, event wiring, and RCW lifetime in the 3427-line WinForms class; .NET 4.8 + xUnit project + packages.config + .sln edits are fiddly; any behavioral drift in CAS XML, Excel setup, findings, or PPT generation is unacceptable and will force re-reads. Estimate driven by: monolithic file size (multiple targeted reads + context rebuilds), 5 service extractions each requiring interface design + Main.cs surgery + delegate changes, plus concurrent addition of test infrastructure in a legacy Office-Interop codebase.

## codex — notes

This is a broad legacy refactor across a 3,400-line WinForms form with Office Interop, so most effort is understanding behavior, extracting services safely, and preserving event-handler semantics. The xUnit project and smoke tests are straightforward, but making services unit-testable requires introducing seams around Interop and persistent XML access. Estimate assumes incremental extraction of the suggested five service areas with compile/debug cycles, and possible friction from .NET Framework 4.8 and Office dependencies.

