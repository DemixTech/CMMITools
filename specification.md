# CMMI Tools — Master Specification

This is the **master spec** for the cmmitools repo. It defines what is shared across all
components: the canonical CMMI domain model, the Azure database contract, and the
cross-component migration roadmap.

> **Rule: master spec *defines*, component specs *bind*.** The domain model and Azure
> schema live here exactly once. Each component's `specification.md` documents how *it*
> maps to these definitions (its EF mapping, its DTO shape, its Form/View layout) — it does
> not re-define them. When a shared concept changes, change it here first.

See the root `README.md` for the component map, kanban prefixes, and routing rule.

---

## 1. Canonical CMMI domain model

These are the appraisal-domain entities the tools manipulate. They exist today as the
C# types in `BASE_source/Data/` (XML-persisted via `PersistentData`); they are the
starting point for the Azure schema and the `CMMI_online` EF Core model.

| Entity | Meaning | Today in `BASE_source/Data/` |
|---|---|---|
| **Organisation** | The appraised organisation. | `Organisation.cs` |
| **OrganizationalUnit (OU)** | The unit within the org that is in scope for the appraisal. | `OrganizationalUnit.cs` |
| **OUProcess** | A process the OU runs, mapped to model practices. | `OUProcess.cs` |
| **PracticeArea (PA)** | A CMMI practice area (e.g. GOV, II, …). | `PracticeArea.cs` |
| **Practice** | An individual practice within a PA (e.g. `GOV 2.3`), with rating/strength/weakness. | `Practice.cs` |
| **Staff** | Appraisal team members and org participants. | `Staff.cs` |
| **Schedule (1 & 2)** | Appraisal schedule / interview plan. | `Schedule1.cs`, `Schedule2.cs` |
| **OEdb (Objective Evidence db)** | The objective-evidence database; rows of evidence mapped to practices. | `OEdbFile.cs`, `OEdbProcessor.cs` |
| **CAS Plan** | The CAS appraisal plan workbook contents. | `CasPlanFile.cs` |
| **Question / Questions** | Appraisal question model. | `Question.cs`, `QuestionsFile.cs` |
| **MDD Toolkit / Map records** | Method-definition toolkit & practice→evidence maps. | `MddToolkit.cs`, `MapRecord.cs` |
| **Example artifacts** | Example activities / work products per practice. | `ExampleActivity.cs`, `ExampleWorkProduct.cs`, `EPAcode.cs`, `ESampleType.cs` |
| **DataReference / WorkUnit** | Supporting reference data and units of work. | `DataReference.cs`, `WorkUnit.cs` |

Glossary: **CAS** = Conformance Appraisal System (`cas.cmmiinstitute.com`); **PA** = Practice
Area; **OU** = Organizational Unit; **OEdb** = Objective Evidence database; **OoS** =
out of scope; **IIGOV** = the II/GOV practice areas given special handling in BASE.

## 2. Azure database contract (TARGET — to be defined)

The long-term plan is a **centralised Azure database** as the single source of truth, with:

- `CMMI_online` as the **owner of all writes** (via EF Core), exposing an API.
- `BASE_source` migrating over time from its local XML persistence to *reading* (then
  optionally writing through the API) from Azure.
- `CAS_updater` POSTing parsed `CASxxxx.xlsm` data to `CMMI_online`'s API rather than
  writing CAS directly.

```
CAS_updater ──API──▶
                     CMMI_online ──EF Core──▶  Azure DB  (single source of truth)
BASE_source ──API/read──▶
```

The concrete schema (tables, keys, naming, migration strategy, who reads/writes which
tables) is **not yet defined** — that is Initiative I-1 below and root task
`_kanban/R01_define_azure_db_schema`. Until it lands, no component should assume a schema.

## 3. Migration roadmap (cross-component Initiatives)

Cross-cutting work lives here as numbered Initiatives, each split into one single-owner
task per component (see routing rule in README).

- **I-1 — Define the v1 Azure DB schema.** Owner: root (`R01`). Translate §1 into tables +
  keys + an EF Core migration. Blocks I-2 and I-3.
- **I-2 — `CMMI_online` reads/writes Azure.** Owner: `CMMI_online` (`O02`). Wire EF Core to
  the I-1 schema.
- **I-3 — `CAS_updater` populates via the `CMMI_online` API.** Owners: `CMMI_online`
  (`O03`, the ingestion endpoint) + `CAS_updater` (`C01`, the client). Replaces direct
  scraping of the live CAS site.
- **I-4 — `BASE_source` reads from Azure.** Owner: `BASE_source` (future `Sxx`). Behind a
  feature flag; defaults to local XML during transition.
- **I-5 — Port NASA `…CAS_Plan.xlsm` macros.** Owner: split per macro into `CMMI_online`
  (or `BASE_source`) tasks. The xlsm stays unmodified (archival source); mark each macro
  "ported" here once its target task is done.

Keep this list ordered by dependency. Add an Initiative *before* breaking it into tasks.

## 4. Conventions (shared)

- Kanban: `_kanban/<ID>_Title/{status.json, request.md, outcome.md}`. `status.json` =
  `{ "status", "agentStartTime", "agentCompletionTime" }`; statuses: `backlog` (untriaged),
  `assigned`, `wip`, `completed`, plus `on hold` / `deferred` as needed.
- One commit per task, ID-prefixed message.
- IDs: **R** root, **S** BASE_source, **O** CMMI_online, **C** CAS_updater. Prefixes prevent
  collisions across the three queues and make commit messages self-locating.
