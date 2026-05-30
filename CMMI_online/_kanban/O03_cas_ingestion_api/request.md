# O03 — CAS ingestion API endpoint (Initiative I-3, CMMI_online half)

> Status: **backlog**. Depends on `O02`. Counterpart task: `CAS_updater/_kanban/C01`.

## Goal

Expose an HTTP API that `CAS_updater` POSTs parsed `CASxxxx.xlsm` data to, so CAS data
flows `CAS_updater → CMMI_online → Azure DB` instead of `CAS_updater` driving the live
`cas.cmmiinstitute.com` site directly.

## Scope

- Define the request **DTO** mirroring the CAS field map
  (`../../CAS_updater/cas-scraper/CAS_Field_Mapping.md`, `_xlsCasMap_MASTER.xlsx`). Share
  the DTO shape with `CAS_updater` (`C01`).
- Add an endpoint (e.g. `POST /api/cas/appraisal`) that validates and persists via EF Core
  (`O02`).
- Auth for the endpoint; idempotency/upsert semantics for re-submitting an appraisal.

## Done when

- The endpoint accepts the DTO and writes the appraisal to Azure.
- The DTO contract is documented in `specification.md` and referenced by `C01`.

## Notes

- This is the server half of Initiative I-3; keep it in sync with `CAS_updater/C01` (the client half).
