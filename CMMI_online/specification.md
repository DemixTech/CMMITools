# CMMI_online — specification

> **Status: not yet implemented.** This describes the *intended* architecture. Update it to
> describe what actually exists as `O01`+ land. Binds to root `../specification.md`
> (canonical CMMI domain model §1, Azure DB contract §2, roadmap §3).

## 1. Intended architecture

- **ASP.NET (Core) MVC** — controllers + views; services behind interfaces; DI throughout.
- **EF Core** model mapping the master spec §1 domain entities to the Azure SQL schema
  defined in master spec §2 (`R01`). This app **owns all writes** to that DB.
- A small **ingestion API** (e.g. `POST /api/cas/appraisal`) consumed by `CAS_updater`.

## 2. Binding to master spec

- Domain entities are **not** re-defined here — they come from master spec §1. This spec
  documents the **EF Core mapping** (entity → table/columns) once the schema (`R01`) exists.
- The Azure DB contract (who reads/writes which tables) is master spec §2; this component is
  the write owner.

## 3. Open decisions (resolve as tasks land)

- ASP.NET version, hosting target (Azure App Service?), auth model — pinned in `O01`.
- API surface and DTO shapes for CAS ingestion — defined in `O03`, shared with `CAS_updater`.
- Migration sequencing from `BASE_source` (I-4) and macro ports from the NASA xlsm (I-5).

## 4. Gotchas

_(none yet — add as the project grows: EF Core migration pitfalls, Azure connection-string
handling, etc.)_
