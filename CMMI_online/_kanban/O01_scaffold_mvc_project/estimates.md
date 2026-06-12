# Estimates — O01_scaffold_mvc_project

Generated 2026-06-12T12:25:40Z · estimator level factor 0.75 · canonical record: `estimates.json`

| AI | status | complexity (1-10) | est. tokens | executing model | est. seconds | estimator model |
|---|---|---|---|---|---|---|
| claude | ok | 5 | 320000 | claude-sonnet-4-6 | 780 | claude-opus-4-8 |
| grok | ok | 4 | 36500 | grok-build | 390 | grok-build |
| codex | ok | 3 | 9000 | gpt-5.4-mini | 420 | gpt-5.5 |

## claude — notes

Mostly well-trodden scaffolding: `dotnet new` for MVC + xUnit test project, adding EF Core + Microsoft.Data.SqlClient packages (no schema/migrations yet), .gitignore, and a moderate specification.md rewrite from 'intended' to 'as-built' pinning ASP.NET Core version, hosting, auth, and a secrets approach (user-secrets/connection-string-out-of-source). Drivers are the doc/decision work and iterating until it builds and the home page runs. Main risks: whether the .NET SDK is installed/available in the run environment and a build-fix loop if package versions conflict; if no SDK, verification of 'builds and runs' can't be completed. No business logic, so complexity stays mid-range.

## grok — notes

Greenfield .NET scaffold (dotnet new mvc + sln, EF Core SqlServer package only, basic HomeController/view, xUnit test proj, standard .gitignore, local secrets via User Secrets). Agent must also read component + root specs, then rewrite specification.md §1/§3 from intended to as-built with pinned ASP.NET version, hosting target, auth model, and config approach. Complexity 4 (not 2-3) due to judgment calls on stack pins that must align with Azure DB owner role and unblock later O tasks, plus full verification that the app builds and home page renders. ~7-10 LLM turns expected (context reads, CLI outputs, targeted edits, build/run loops); estTokens and seconds reflect capable grok-build handling of boilerplate + precise docs with low iteration risk.

## codex — notes

Greenfield scaffold with standard dotnet templates, EF Core Azure SQL package, a basic test project, gitignore, and specification updates. Main time drivers are reading/updating the existing spec, package restore/build verification, and documenting concrete stack decisions. Risk is mostly local SDK/package availability or needing minor template/version adjustments.

