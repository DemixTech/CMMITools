# R01 — Define the v1 Azure DB schema (Initiative I-1)

> Status: **backlog**. This is a **root** task because the schema is a shared contract no
> single component owns yet. Blocks I-2 (`O02`) and I-3 (`O03`/`C01`).

## Goal

Translate the canonical CMMI domain model (root `specification.md` §1) into a concrete
**Azure SQL** schema: tables, primary/foreign keys, naming convention, and an initial
**EF Core migration** that `CMMI_online` will own.

## Scope

- Map each §1 entity (Organisation, OrganizationalUnit, OUProcess, PracticeArea, Practice,
  Staff, Schedule, OEdb, CasPlan, Question, …) to one or more tables.
- Decide naming convention (table/column casing, key naming) and write it into
  master spec §2.
- Define which component reads vs. writes which tables. Baseline: `CMMI_online` owns all
  writes; `BASE_source` and `CAS_updater` go through its API.
- Produce the EF Core migration in `CMMI_online` once that project is scaffolded (`O01`),
  or as a standalone schema script if `O01` isn't done yet.

## Done when

- Master spec §2 is filled in with the concrete schema + naming convention.
- An EF Core migration (or SQL DDL) exists and applies cleanly to a fresh Azure SQL DB.
- §2's "who reads/writes which tables" table is complete.

## Notes

- Don't over-design: v1 only needs to cover what `CMMI_online` (`O02`) and the CAS
  ingestion path (`O03`) actually touch first. Extend per Initiative afterwards.
