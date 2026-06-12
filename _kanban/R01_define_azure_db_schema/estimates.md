# Estimates — R01_define_azure_db_schema

Generated 2026-06-12T12:26:49Z · estimator level factor 0.75 · canonical record: `estimates.json`

| AI | status | complexity (1-10) | est. tokens | executing model | est. seconds | estimator model |
|---|---|---|---|---|---|---|
| claude | ok | 7 | 240000 | claude-opus-4-8 | 540 | claude-opus-4-8 |
| grok | ok | 6 | 15000 | grok-build | 360 | grok-build |
| codex | ok | 7 | 85000 | gpt-5.5 | 4800 | gpt-5.5 |

## claude — notes

Drives: translating ~10+ CMMI domain entities into a normalized Azure SQL schema, settling a naming convention, authoring the read/write ownership matrix, and producing a working EF Core migration (or standalone DDL). High because it's a shared contract that blocks two initiatives, so correctness and forward-compat matter, and it requires reading/absorbing the master spec §1 plus component bindings. Tempered by the explicit 'don't over-design, cover only what O02/O03 touch first' scope. Main risks: O01 not yet scaffolded (forces standalone DDL path) and verifying the migration 'applies cleanly to a fresh Azure SQL DB' without DB access — likely delivered untested, inflating later rework. Token estimate assumes a few design-iteration passes over the spec and migration.

## grok — notes

Assumes root specification.md §1 contains a clear but not exhaustive description of the ~10 listed entities and their relationships; v1 scope keeps table count to ~10-15 with basic FKs and no advanced features like partitioning. The bulk of work is mapping to relational model, codifying naming (e.g. PascalCase tables, Id PKs), populating the ownership matrix, and generating a valid EF Core migration or DDL script. Risks include entity ambiguities requiring reasonable defaults and ensuring the artifact is immediately applicable to Azure SQL without further O01 scaffolding.

## codex — notes

This is a shared contract/design task, not just mechanical coding: the agent must read the canonical domain model, decide a durable Azure SQL naming/key convention, document ownership boundaries, and produce either EF Core migration code or standalone DDL. The estimate assumes CMMI_online may not yet be scaffolded, so a SQL DDL fallback is likely. Main risks are incomplete domain detail in the current specs and needing to reconcile v1 scope against future entities without over-designing.

