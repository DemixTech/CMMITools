# BASE (BASE_source) — component README

The legacy desktop appraisal tool. WinForms, .NET Framework **4.8**, heavy Microsoft
Office Interop. ClickOnce-published to `../BASE_install/`.

Kanban prefix: **S**. See `specification.md` for architecture & gotchas, and the root
`../specification.md` for the shared CMMI domain model and Azure roadmap this component
will migrate toward.

## Stack

- **WinForms**, `net48` (`<TargetFrameworkVersion>v4.8</TargetFrameworkVersion>`, `WinExe`, namespace `BASE`).
- Office Interop: **Excel 15.0**, **Word 15.0**, **PowerPoint 15.0**, **Outlook 14.0** — requires a machine with the matching Office installed.
- Domain persisted to local **XML** files (e.g. `BASE\CasPlanFileXML.xml`, `BASE\OEdbFile.xml`) via `Data/PersistentData.cs`.

## Build / run

- Open `BASE.sln` in Visual Studio (Windows + Office required for Interop) and run, or build the `BASE` project (`BASE.csproj`).
- Publish: ClickOnce → outputs to `../BASE_install/`. Signing instructions in `20220710_v01_SignTool_Instructions.docx`.
- `bin/`, `obj/`, `packages/` are build output (gitignored).

## What it does (UI surface)

`Main.cs` (one ~3,400-line form) drives the workflow via tab/button handlers:
load/save the CAS Plan (XML), generate schedules, set up the main appraisal workbook,
insert interviews, merge OE sources, hide out-of-scope rows, extract findings,
IIGOV characterization & rating, build OU maps, import the model/questions, and generate
the "full tool" + presentations.

## Component workflow

- Pick tasks from `_kanban/` (prefix **S**) per the root README's task workflow.
- Source of the backlog: `../04_BugAndEnhancementList.xlsx`.
- One commit per task, e.g. `S03: Externalise hard-coded paths`.

## Heads-up before editing

`Main.cs` is large and const-heavy with a hard-coded 2020 appraisal path. Read
`specification.md` §Gotchas first — several seed tasks (`S02`, `S03`) exist specifically to
make this codebase amenable to small, safe edits.
