# Estimates — O03_cas_ingestion_api

Generated 2026-06-12T12:28:17Z · estimator level factor 0.75 · canonical record: `estimates.json`

| AI | status | complexity (1-10) | est. tokens | executing model | est. seconds | estimator model |
|---|---|---|---|---|---|---|
| claude | ok | 6 | 240000 | claude-sonnet-4-6 | 1020 | claude-opus-4-8 |
| grok | ok | 5 | 19500 | grok-build | 265 | grok-build |
| codex | ok | 7 | 95000 | gpt-5.5 | 14400 | gpt-5.5 |

## claude — notes

Assumes O01 (scaffold) and O02 (EF Core/Azure) are done, so this is endpoint + DTO + persistence work on an existing project. Main driver is mirroring the full CAS field map into a shared DTO — the map size/fidelity is the biggest risk and could balloon the DTO and validation code. Auth is unspecified (TBD in O01), so a sane default (API key/bearer) is assumed; idempotent upsert semantics add care but are routine. Cross-repo contract sync with C01 and documenting it in specification.md is light but must be kept consistent. Lands mid-band, hence sonnet.

## grok — notes

Assumes O01/O02 done (MVC+EF+entities+Azure wiring+auth pinned) and that the CAS map (107 rows in xlsx, partial in md) yields a stable but non-trivial DTO with sections + collections. Primary work: DTO definition mirroring the map (for contract sharing), POST /api/cas/appraisal endpoint with validation + DTO-to-EF-entity mapping + idempotent upsert, auth application, and precise documentation of the contract in specification.md. Lower than O02 (7) and C01 (7) because no schema work or multi-file TS refactoring; risks are mapping impedance between CAS-form DTO and domain entities plus potential auth gaps for the machine client. estTokens/seconds reflect 6-9 focused turns on an already-scaffolded project (reads of post-O02 code + map docs, targeted model+controller+spec edits, build+basic verify).

## codex — notes

Assumes O01/O02 have already created the ASP.NET Core app, EF Core model, Azure SQL connection, and basic test setup. The main cost is reading the CAS field map, designing a shared DTO contract, mapping it into the existing EF model, adding validation/auth/idempotent upsert behavior, and documenting the contract for CAS_updater C01. Risk is driven by unknown CAS field-map size and whether O02’s persistence model already matches the appraisal shape cleanly.

