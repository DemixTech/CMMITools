# O01 — Scaffold the ASP.NET MVC project

> Status: **backlog**. First task for this component; unblocks O02/O03.

## Goal

Create the `CMMI_online` ASP.NET MVC project in this directory and pin the stack decisions.

## Scope

- Scaffold an ASP.NET (Core) MVC app (csproj/sln, Program, a home controller + view).
- Add **EF Core** with the Azure SQL provider (no schema yet — that's `R01`/`O02`).
- Pin and document in `specification.md` §1/§3: ASP.NET version, hosting target, auth model,
  configuration/secrets approach (connection strings out of source control).
- Add a test project.

## Done when

- The project builds and runs locally (home page renders).
- `specification.md` is updated from "intended" to "as-built" for what now exists.
- `.gitignore` covers `bin/`, `obj/`, and any local secrets.
