# CMMI Tools

Tooling that helps CMMI lead appraisers **plan & perform appraisals, populate CAS, and build their presentations**. The original rationale is in `01_StartupEmail.pdf`.

This repo is a **single git repository with several components**, run with a translator3-style AI workflow (kanban queue + living docs per component). There is no live production system to break, so the workflow is a plain build/test/edit loop — not the headless/cache machinery used for the live WordPress projects.

## Components

| Dir | What it is | Kanban prefix |
|---|---|---|
| `BASE_source/` | Legacy **WinForms .NET Framework 4.8** desktop app "BASE" — Excel/Word/PowerPoint/Outlook Interop automation over appraisal workbooks. ClickOnce-published to `BASE_install/`. | **S** |
| `CMMI_online/` | The **ASP.NET MVC** online port (greenfield). Owns writes to the centralised **Azure** database. The long-term migration target. | **O** |
| `CAS_updater/` | TypeScript/Node tool that reads `CASxxxx.xlsm` and populates the CAS system (`cas.cmmiinstitute.com`). Moving toward feeding `CMMI_online`'s API instead of driving the live site. | **C** |
| `_kanban/` (root) | Tasks no single component owns (e.g. defining the shared Azure schema). | **R** |

Migration sources (read-only, **not** components): `2024-05-04to05-10 (A5) C384400 NASA/` and its `…CAS_Plan.xlsm` (macros to port), `04_BugAndEnhancementList.xlsx` (BASE backlog), `02_Overview.pptx` / `05_IntergalacticSPIN.pptx`, the `9999-PPTX Linker/` helper tool.

## Docs layout

- **Root `README.md`** (this file) — orientation, component map, routing rule.
- **Root `specification.md`** — the **master spec**: the canonical CMMI domain model, the Azure DB contract, and the cross-component migration roadmap. *Defined here once; component specs bind to it.*
- **Per component** — a `README.md` (stack + how to build/run + component workflow) and a `specification.md` (architecture today + gotchas + how it binds to the master spec).

## Task workflow (per component)

1. `~/find_assigned.sh /mnt/c/GitHub/CMMITools` lists every `_kanban/<ID>_Title/status.json` whose `status == "assigned"` (it already walks all components recursively, pruning `node_modules`/`.git`/`bin`/`obj`).
2. Pick one, set `status.json` → `"wip"` with `agentStartTime`, read `request.md`, do the work.
3. Write `outcome.md`, set `status.json` → `"completed"` with `agentCompletionTime`.
4. **One commit per task**, message prefixed with the task ID — e.g. `O14: Add CAS ingestion endpoint`, `S03: Externalise hard-coded paths`.

### Routing rule (which `_kanban/` does a task go in?)

- Touches **one** component → that component's `_kanban/` (prefix S/O/C).
- Defines a **shared contract** no component owns yet (e.g. the Azure schema) → root `_kanban/` (prefix R).
- **Cross-cutting** work → it is *not* one shared task. Record it as a numbered **Initiative** in `specification.md` §Roadmap, then split into one single-owner task per component, each referencing the Initiative.

## Drive mapping

`subst-for-Gdrive.bat` / `subst-for-Xdrive.bat` map a drive letter to this directory for easy access to appraisal files. Edit to taste.
