# CMMI_online — component README

The online port of BASE: an **ASP.NET MVC** web application and the long-term migration
target. It owns writes to the centralised **Azure** database.

> **Status: greenfield / empty.** Nothing is scaffolded yet. The first task (`O01`) creates
> the project. Until then, the architecture below is *intended*, not implemented.

Kanban prefix: **O**. See `specification.md` for the intended architecture and the root
`../specification.md` for the shared CMMI domain model and Azure roadmap.

## Intended stack (decide & pin in `O01`)

- **ASP.NET (Core) MVC**.
- **EF Core** against **Azure SQL** — this app is the single owner of writes to that DB.
- Auth + hosting target: TBD in `O01`.

## Role in the system

```
CAS_updater ──API──▶ CMMI_online ──EF Core──▶ Azure DB ◀──(future) BASE_source reads
```

- Exposes an ingestion API so `CAS_updater` can POST parsed `CASxxxx.xlsm` data instead of
  driving the live CAS site (Initiative I-3 → tasks `O03` + `CAS_updater/C01`).
- Over time, absorbs functionality migrated from `BASE_source` and macros ported from the
  NASA `…CAS_Plan.xlsm` (master spec §3 Initiatives I-4, I-5).

## Component workflow

- Tasks in `_kanban/` (prefix **O**), per the root README's workflow.
- One commit per task, e.g. `O02: Wire EF Core to Azure schema`.
