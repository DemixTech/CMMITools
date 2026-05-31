# S03 — Break up Main.cs into services

> Status: **backlog**. Depends conceptually on `S02` (constants externalised first).

## Problem

`Main.cs` is ~3,427 lines in a single form class. Every AI edit has high merge-collision
risk and forces re-reading the whole file to be safe.

## Scope

Extract behaviour by region into services behind small interfaces, leaving `Main.cs` as a
thin UI shell that delegates. Suggested seams (from `specification.md` §2):

- `AppraisalPlanService` — CAS Plan load/save (XML), schedule generation.
- `CasPopulateService` — main-workbook setup, interview insertion.
- `OEdbImportService` — OE/OEdb import, merge, OU maps.
- `FindingsService` — findings extraction, IIGOV characterization & rating, OoS row hiding.
- `PresentationBuilderService` — "full tool" + presentation generation.

Add an **xUnit test project** at the same time so subsequent tasks have somewhere to land
tests.

## Done when

- `Main.cs` event handlers delegate to the services; no business logic left inline beyond UI wiring.
- Services are unit-testable (Interop boundaries behind interfaces where practical).
- A test project exists and builds, with at least smoke coverage of one extracted service.

## Notes

- Do this incrementally — one service per commit is fine (`S03a`, `S03b`, …) if it gets large.
