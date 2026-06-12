# Estimates — O02_wire_ef_core_to_azure

Generated 2026-06-12T12:26:31Z · estimator level factor 0.75 · canonical record: `estimates.json`

| AI | status | complexity (1-10) | est. tokens | executing model | est. seconds | estimator model |
|---|---|---|---|---|---|---|
| claude | ok | 8 | 460000 | claude-opus-4-8 | 1500 | claude-opus-4-8 |
| grok | ok | 7 | 22500 | grok-build | 510 | grok-build |
| codex | ok | 8 | 85000 | gpt-5.5 | 10800 | gpt-5.5 |

## claude — notes

Drivers: a full EF Core model + DbContext mapping ~10+ domain entities with relationships, plus migration application to a live dev Azure SQL DB and connection-string/secrets wiring. Greenfield component (only README/spec exist) so all scaffolding is net-new, and it depends on O01/R01 being done first. Key risks: the actual Azure SQL connectivity and migration apply can't be fully verified by the agent without real credentials/provisioned DB, and entity↔table fidelity to spec §2 plus read/write endpoint guarding adds iteration. Complexity 8 lands in the opus band; token estimate assumes several edit/build/migration cycles.

## grok — notes

Key assumptions: O01 will have created a basic ASP.NET Core MVC project with EF Core packages and DI skeleton in place; R01 will have produced a defined Azure SQL schema plus an applyable migration matching the master spec §1 entities (Organisation, OU, OUProcess, PracticeArea, Practice, Staff, Schedule, OEdb, CasPlan, Question, …). Scope is focused on POCOs + DbContext + fluent mappings, secure config/secrets for the connection string, read endpoints (writes API-guarded), applying the migration to a dev Azure SQL instance, and writing the §2 as-built entity→table docs. Risks: potential column, constraint, or relationship mismatches between the textual schema and live DB that require iteration; Azure-specific setup (firewalls, auth method, retry policies) and the 'never committed' secret requirement may surface extra friction; greenfield nature means initial wiring, service registration, and basic test reads must be built rather than extended. Estimate driven by ~10+ entities plus relationships, the canonical write-owner role, partial API surface, migration execution, and mandatory spec update in a non-trivial domain model.

## codex — notes

This is a broad data-layer integration task touching domain modeling, EF Core mappings, configuration/secrets, Azure SQL migration/application, API read/write behavior, and specification updates. The main risks are schema ambiguity from R01, relationship/cardinality mismatches across the CMMI domain, and external Azure DB access or credential availability. Estimate assumes O01 has scaffolded a working ASP.NET Core project and R01 has produced a usable SQL schema/migration, but still requires substantial implementation and verification.

