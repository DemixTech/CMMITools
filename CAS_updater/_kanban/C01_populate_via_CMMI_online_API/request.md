# C01 — Populate via the CMMI_online API (instead of the live CAS web UI)

> Status: **backlog** (future work — not yet assigned). Flip `status.json` to `"assigned"` when ready to start.

## Goal

Today `CAS_updater` reads `CASxxxx.xlsm` data and populates the **live CAS web system**
(`cas.cmmiinstitute.com`) by driving the browser / form fields directly
(`scraper.ts`, `populator.ts`, the `*.handler.ts` page handlers).

The target architecture is for `CAS_updater` to instead **hand its parsed data to the
`CMMI_online` MVC app via an HTTP API**, and let `CMMI_online` own the writes to the
centralised Azure database. `CAS_updater` becomes a *client* of `CMMI_online`, not a
direct writer to the upstream CAS site.

```
CASxxxx.xlsm ──read──▶ CAS_updater ──HTTP API──▶ CMMI_online ──EF Core──▶ Azure DB
```

## Why

- Single owner of writes to the Azure schema = `CMMI_online` (see root `specification.md` §Azure DB contract). Avoids two components writing the same tables.
- Decouples `CAS_updater` from the fragile live-site DOM / form-field scraping.
- Reuses `CMMI_online`'s validation, auth, and domain model rather than duplicating it.

## Prerequisites (sequence)

1. **CMMI_online exists** with an ingestion endpoint (e.g. `POST /api/cas/appraisal`) and the Azure DB contract is defined in the **root `specification.md` §Roadmap** as a cross-component Initiative.
2. The CAS field mapping (`cas-scraper/CAS_Field_Mapping.md`, `_xlsCasMap_MASTER.xlsx`) is expressed as a stable DTO shared with `CMMI_online`.

## Scope of this task (when picked up)

- Define the request DTO that mirrors the current CAS field map.
- Add an API-client mode to `CAS_updater` that POSTs parsed `CASxxxx.xlsm` data to `CMMI_online` instead of driving the browser.
- Keep the existing direct-CAS-site populator behind a config flag during transition.

## Notes

- This task is the `CAS_updater` half of a cross-component initiative; the `CMMI_online`
  half (the ingestion endpoint) lives in `CMMI_online/_kanban/` under prefix `O`.
- Reference the initiative number from the root master spec in both tasks once it exists.
