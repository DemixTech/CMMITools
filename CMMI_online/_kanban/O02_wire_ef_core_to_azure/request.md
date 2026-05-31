# O02 — Wire EF Core to the Azure DB schema (Initiative I-2)

> Status: **backlog**. Depends on `O01` (project exists) and `R01` (schema defined).

## Goal

Implement the EF Core model + DbContext mapping master spec §1 entities to the §2 Azure
schema, and connect `CMMI_online` to an Azure SQL database as the **write owner**.

## Scope

- EF Core entities + DbContext for the §1 domain (Organisation, OU, OUProcess, PracticeArea,
  Practice, Staff, Schedule, OEdb, CasPlan, Question, …).
- Apply the `R01` migration to a dev Azure SQL DB.
- Connection string via configuration/secrets, never committed.
- Read endpoints first; writes guarded behind the API.

## Done when

- The app reads/writes the domain against Azure SQL through EF Core.
- `specification.md` §2 documents the entity→table mapping (as-built).
