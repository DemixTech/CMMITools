# cas_helper — system specifications

Audience: an AI coder (Claude or similar) who has not seen this codebase before. This document is dense and assumes you can read code; it gives you the map, the invariants, and the gotchas so you can make accurate edits without re-deriving everything.

## 1. Purpose

`cas_helper` automates the work of filling out a CMMI **CAS (Conformance Appraisal System)** appraisal in two directions:

- **Excel ← startup file.** Populate a project-specific CAS Plan workbook (`.xlsm`) from a "demix"/startup workbook supplied by the customer. Driven by a `_FieldMap` sheet *inside the target workbook*. Python, run via the two skills under `skills/`.
- **Web ← Excel.** Drive a Puppeteer browser session against `https://cas.cmmiinstitute.com` and populate dozens of CAS web pages from the CAS Plan workbook. Driven by an `_xlsCasMap` sheet (lives in the workbook, with a master copy in `cas-scraper/_xlsCasMap_MASTER.xlsx`). TypeScript, run via `npm run populate` from `cas-scraper/`.

These are **two separate "FieldMap" sheets in two separate workflows**. Do not conflate them:

| Sheet | Lives where | Used by | Drives |
|---|---|---|---|
| `_FieldMap` | Inside each project's CAS Plan `.xlsm` | `skills/skill-setup-base-cas-plan/setup_cas_plan.py`, `skills/skill-audit-cas-plan/audit_cas_plan.py` | Copying values from startup file → CAS Plan workbook |
| `_xlsCasMap` | `cas-scraper/_xlsCasMap_MASTER.xlsx` (fallback: inside CAS Plan `.xlsm`) | `cas-scraper/populator.ts` | Filling CAS web form fields from CAS Plan workbook |

A scraper (`cas-scraper/scraper.ts`) exists primarily to discover/refresh the web-side field inventory; it is **not** part of the day-to-day populate workflow.

## 2. Top-level layout

```
cas_helper/
├── cas-project-config.json     ← per-project paths + CAS settings (single source of truth)
├── .secrets/
│   └── keys.json               ← CAS credentials (gitignored)
├── README.md                   ← summary of recent config-paths refactor
├── specifications.md           ← (this file)
│
├── cas-scraper/                ← TypeScript: web populator + scraper
│   ├── package.json            ← npm run populate / start / build
│   ├── config.ts               ← loads ../cas-project-config.json + ../.secrets/keys.json
│   ├── types.ts                ← FieldMapping, ExcelData, PopulateLog interfaces
│   ├── populator.ts            ← MAIN entry for `npm run populate`. CASPopulator class + main loop.
│   ├── scraper.ts              ← `npm start` — re-discovers form fields across all CAS pages
│   ├── test-download.ts        ← one-off download debug script
│   ├── _xlsCasMap_MASTER.xlsx  ← canonical web field-map (preferred over per-project copy)
│   ├── helpers/
│   │   ├── fieldPopulators.ts  ← mixin: populateText / Select / Radio / Checkbox / Number /
│   │   │                          Multiselect / RadioLevel / Date / DateParts + populateField dispatcher
│   │   └── injectTemplate.py   ← Python helper invoked via execSync for "template workflow" pages
│   ├── handlers/               ← one mixin module per page family (see §5)
│   │   ├── orgScope.handlers.ts        (12 page handlers, ~2000 LOC — biggest file)
│   │   ├── objectiveEvidence.handlers.ts
│   │   ├── resourceEstimates.handler.ts
│   │   ├── logisticalRequirements.handler.ts
│   │   ├── appraisalConstraints.handler.ts
│   │   ├── riskIdentification.handler.ts
│   │   └── conflictsOfInterest.handler.ts
│   ├── .browser-data/          ← persistent Puppeteer profile (keeps login across runs)
│   ├── html_logs/              ← before/after HTML snapshots per page (timestamp-prefixed)
│   ├── populate_log.json       ← per-page success/fail/skipped record (PopulateLog[])
│   ├── CAS_Field_Mapping.md    ← human-readable cross-reference; informational only
│   ├── README.md
│   └── run_scraper.bat         ← Windows launcher for scraper.ts (interactive CAS_EMAIL prompt is now obsolete; keys.json supersedes it)
│
└── skills/                     ← Python skills, packaged for invocation by Claude
    ├── skill-audit-cas-plan/
    │   ├── SKILL.md            ← skill metadata + usage
    │   ├── audit_cas_plan.py   ← validates _FieldMap against actual sheet content; writes results to col G
    │   └── run_audit.bat       ← reads files.target from cas-project-config.json
    └── skill-setup-base-cas-plan/
        ├── SKILL.md
        ├── setup_cas_plan.py   ← copies values from source startup file → target CAS Plan workbook
        ├── field_mappings.json ← alias hints
        └── run_setup.bat       ← reads files.source + files.target from cas-project-config.json
```

**Files to ignore.** At the top of `cas-scraper/` there are duplicate handler files (`appraisalConstraints.handler.ts`, `conflictsOfInterest.handler.ts`, `logisticalRequirements.handler.ts`, `orgScope.handlers.ts`, `riskIdentification.handler.ts`, plus `handleOrgProjectsPage.ts`, `populator - Copy.ts`, `populator_3.ts`). The canonical versions live under `handlers/`. The root-level copies are dev backups and **are not imported** by `populator.ts`. Do not edit them; prefer to leave them or delete on request.

## 3. Configuration

After the May 2026 refactor (see `README.md`), `config.ts` resolves both files **relative to `__dirname` (`cas-scraper/`)**, so `npm run populate` works regardless of cwd:

- `../cas-project-config.json` — project paths, CAS settings, no secrets
- `../.secrets/keys.json` — credentials only

`cas-project-config.json` shape (load this before doing anything project-specific):

```jsonc
{
  "project":  { "name": "...", "casId": "83325", "description": "..." },
  "files":    { "source": "C:/.../Demix.xlsm", "target": "C:/.../CAS_Plan.xlsm" },
  "utilities": { "cas-scraper": "..." },
  "skills":   { "skill-audit-cas-plan": "...", "skill-setup-base-cas-plan": "..." },
  "cas": {
    "baseUrl":            "https://cas.cmmiinstitute.com",
    "loginUrl":           "https://cmmiinstitute.com/login",
    "continueFromPage":   "/appraisal-participants",   // "" = start from beginning
    "debugMode":          false,                       // true = prompt after each page
    "autoExitOnComplete": false                        // true = exit on reaching finalPage
  }
}
```

`keys.json` shape:

```json
{ "cas": { "email": "...", "password": "...", "staySignedIn": "yes" } }
```

Env vars `CAS_EMAIL` / `CAS_PASSWORD` are a fallback if `keys.json` is missing. `config.ts` warns rather than silently swallowing parse errors, and exports `CONFIG.projectConfigPath` / `CONFIG.keysPath` so callers can reference the actual file paths in error messages.

## 4. The populator (web ← Excel)

### 4.1 Main loop

`CASPopulator.run()` (`populator.ts:833`):

1. `init()` — launch Puppeteer with `userDataDir: cas-scraper/.browser-data/` so login persists.
2. `tryDirectNavigation()` — navigate to the first/continueFrom page; if it doesn't redirect to login, session is valid.
3. Otherwise `login()` — fills `#UserName` / `#Password`, handles the "Remember me" checkbox per `staySignedIn`.
4. Navigate to `CONFIG.continueFromPage` (or `/name-and-type`).
5. Per-page loop:
   - Extract `pagePath` from `URL.pathname` by stripping `/appraisals/{casId}`.
   - If `pagePath === CONFIG.finalPage` (`/sample-scope`) → `handlePhaseComplete()` and exit.
   - If `pagePath` is in `CONFIG.skipPages` (`/org-unit-project-appraisal-scope`, `/include-project`) → click Next; on failure prompt via `askWhenStuck` (see below).
   - Otherwise `processCurrentPage(pagePath)`, push log, prompt user via `askForFeedback` (only in `debugMode`), then click Next; if Next fails prompt via `askWhenStuck` regardless of `debugMode`.
6. On exit, flush `populate_log.json` and leave the browser open for manual review.

#### 4.1.1 Interactive prompts

Two prompts can pause the loop:

**`askForFeedback()`** — runs after every page when `debugMode: true`. Returns `{action: 'exit' | 'next' | 'reprocess', feedback: string}`:

| Input | Action |
|---|---|
| `a` / `exit` / `quit` / `q` / `s` / `stop` | `exit` — break the loop |
| `b` / `next` / `continue` / *(Enter)* | `next` — click the page's Next button (default) |
| `c` / `reprocess` / `current` / `here` | `reprocess` — skip Next; next iteration re-reads `page.url()` and processes whatever's there |
| *(any other text)* | save as `userFeedback`, then `next` |
| *(text)* `a` / *(text)* `c` | save feedback, then take that action |

Legacy `[s]` / `[q]` suffix patterns are still recognised and map to `exit`.

**`askWhenStuck(reason)`** — runs whenever `clickNextButton()` returns false, **regardless of `debugMode`** (dead-end fallback so the user can drive the browser manually and resume). Returns `{action: 'exit' | 'reprocess', feedback: string}`:

| Input | Action |
|---|---|
| `a` / `exit` / `quit` / etc. | `exit` |
| `c` / `current` / `here` / *(Enter)* | `reprocess` — re-read current URL (default) |
| *(any other text)* | save as `userFeedback`, then `reprocess` |
| *(text)* `a` | save feedback, then `exit` |

The `reprocess` action does **not** call `processCurrentPage` directly; it just `continue`s the while-loop, which then reads `page.url()` fresh at the top of the next iteration. This means the user can manually navigate anywhere in the open Puppeteer window — even to a page outside the original CAS workflow — and the next iteration will run that page through the normal dispatch (finalPage check, skipPages check, handler routing). If the new URL is a skipPages page, it'll be skipped; if it's `finalPage`, the populator finishes.

### 4.2 `processCurrentPage` dispatch

Lives at `populator.ts:622`. Three-tier routing:

**Tier 1 — pre-handlers (run before the field loop, then fall through).** `/organizations`, `/org-units`, `/org-unit-targets` use these to detect Edit-vs-Add mode by checking for an `.item-card`.

**Tier 2 — full handlers (replace the field loop entirely; return early).** All the OE pages, both `/org-projects` variants, sampling factors, subgroups, the timeline, readiness reviews, and the five "template workflow" pages (resource estimates, logistical requirements, appraisal constraints, risk identification, conflicts of interest).

**Tier 3 — generic field loop.** For any remaining page: filter `fieldMap` by `CAS_Page === pagePath` and `CAS_Type !== 'skip'` and `CAS_Selector ∉ {'SKIP', 'HANDLER'}`, then for each field look up `excelData[Sheet][Row]` and call `populateField(mapping, value)`. If anything changed, click the page's save button.

### 4.3 The `_xlsCasMap` field map

`loadFieldMap()` reads from `cas-scraper/_xlsCasMap_MASTER.xlsx` if present, otherwise from `_xlsCasMap` inside `CONFIG.excelFile`. Schema (8 columns, 107 rows in the MASTER at time of writing):

| Col | Field | Notes |
|---|---|---|
| A | `Row`           | Row number in the source `Sheet` to read the value from |
| B | `Sheet`         | Either `P1-OrgScope` or `P1PA-R` (the two Excel sheets that hold appraisal data) |
| C | `FieldLabel`    | Human label, e.g. "Appraisal Name" |
| D | `CAS_Page`      | URL path under `/appraisals/{casId}`, e.g. `/name-and-type` |
| E | `CAS_Selector`  | CSS selector. Sentinel values: `SKIP`, `HANDLER` (handled by a tier-2 method, not the generic loop) |
| F | `CAS_FieldName` | Form field `name` attribute (informational) |
| G | `CAS_Type`      | One of: `text`, `textarea`, `select`, `radio`, `radio-level`, `checkbox`, `number`, `multiselect`, `date`, `date-parts`, `skip` |
| H | `Notes`         | Free text; `radio` types parse `"Yes=#sel1; No=#sel2"` here; `date-parts` types use comma-separated selectors in col E |

Distinct `CAS_Page` values: `/follow-on-activities`, `/logistical-requirements`, `/name-and-type`, `/objective-evidence/{additional-info, collection-approach, collection-responsibilities, collection-techniques, data-collection-timing, performance-report-approaches}`, `/org-unit-sampling-factor-values`, `/org-unit-sampling-factors`, `/org-unit-subgroups`, `/org-unit-targets`, `/org-units`, `/organizations`, `/readiness-reviews`, `/required-outputs`, `/timeline`.

`loadExcelData()` reads values from the project workbook (`CONFIG.excelFile`, which is `files.target`). It reads column B at each row referenced by `fieldMap` for that sheet, **plus** hardcoded "extra rows" for `P1-OrgScope` (rows 69, 70, 73-75, 77-79, 85, 87-88, 91-92, 101-106, 112-117) and `P1PA-R` (rows 50-51, 54-57, 62-64, 67-69, 72, 75-77, 79-81, 86). These extras feed the tier-2 handlers (e.g. timeline phase dates, readiness review details) that need cells the `_xlsCasMap` doesn't list. **If you add a new tier-2 handler that needs extra rows, you must add them to the matching `extraRows` array in `loadExcelData()`.**

`resolveCellValue()` handles rich text, hyperlinks, formulas (with cached results), formula references like `'Sheet Name'!B12`, dates, and booleans → strings. It recurses with depth limit 10. Critical for surviving the Excel formula soup these workbooks accumulate.

### 4.4 Field populators (`helpers/fieldPopulators.ts`)

A single mixin function `applyFieldPopulators(cls)` attaches:

| Method | `CAS_Type` | Notes |
|---|---|---|
| `populateField`        | — (dispatcher) | Switches on `CAS_Type`, returns `{success, changed, error?}` |
| `populateTextInput`    | `text`, `textarea` | Skips if current value already matches |
| `populateSelect`       | `select` | Three-pass option matching: exact, contains, contained-by |
| `populateRadio`        | `radio` | Parses `Notes` as `value=selector;...`; falls back to splitting `CAS_Selector` on `|` and assuming Yes/No |
| `populateRadioLevel`   | `radio-level` | Like `radio` but `Notes` uses comma-separated `n=selector`, default selector `#level-{n}` |
| `populateCheckbox`     | `checkbox` | Truthy values: `yes`/`true`/`1`/`x`/`checked`. Supports `#123` numeric IDs via `getElementById` and the `[data-test="input-virtual-selection_{n}"]` pattern |
| `populateNumberInput`  | `number` | Sets value via direct DOM assignment + `input`/`change` events |
| `populateMultiselect`  | `multiselect` | React-style multiselect; tries many selector families for the option list |
| `populateDateInput`    | `date` | Converts `YYYY-MM-DD` → `MM/DD/YYYY` |
| `populateDateParts`    | `date-parts` | `CAS_Selector` is comma-separated `yearSel,monthSel,daySel`; supports both number inputs and select dropdowns |
| — | `skip` | Short-circuits to success without touching the page |

All populators are idempotent: they check current value and skip if already correct.

### 4.5 Page handlers (`handlers/*.ts`)

Pattern: each module exports `applyXxxHandlers(cls)` which adds methods to `CASPopulator.prototype`. They are wired up at the bottom of `populator.ts` (lines 940-947). Method signatures are duplicated into the `interface CASPopulator` declaration (lines 951-993) for TS type-checking — **when you add a handler, you must update the apply call (line 940-ish), the `interface CASPopulator` declaration (line 951+), AND the dispatch in `processCurrentPage()` (line 622+)**.

Handler inventory:

**`orgScope.handlers.ts` (`applyOrgScopeHandlers`)** — 12 methods covering Phase 1 Org Scope:
- `handleOrganizationsPage` (`/organizations`), `handleOrgUnitsPage` (`/org-units`), `handleOrgUnitTargetsPage` (`/org-unit-targets`) — Edit-vs-Add detection only (tier 1).
- `handleTimelinePage` (`/timeline`) — three-phase flow using `EditPhase` URL param; reads P1-OrgScope rows 69-70, 85, 87-88, 91-92.
- `handleReadinessReviewsPage` (`/readiness-reviews`) — first RR uses rows 73-75 + 101-106; second uses 77-79 + 112-117. Finds RR cards by name match, clicks Edit, sets a Yes/No radio for `characterizedEvidence`.
- `handleOrgProjectsPage` and `handleOUProjectsPage` (`/org-projects` with `IsOrganizational=True`/`False`) — template-download/upload pattern; sources are `C_SupportV2` (cols A-R) and `C_ProjectsV2` (cols A-U, dates in J/K formatted as `mm-dd-yyyy`).
- `handleSamplingFactorsPage`, `handleSamplingFactorValuesPage`, `handleSubgroupsPage`, `handleSubgroupAssignmentPage`, `handleOrgProjectAppraisalScopePage` — each handles its own page; the first three use existence checks to decide between Add and Edit.

**`objectiveEvidence.handlers.ts` (`applyOEHandlers`)** — 7 page handlers plus a `fillAndSaveSimpleForm` helper. Most are P1PA-R driven:
- `handleOECollectionApproachPage` (rows 50-51), `handleOECollectionTechniquesPage` (54-57), `handleOECollectionResponsibilitiesPage` (62-64), `handlePerformanceReportApproachesPage` (67-69), `handleOEInitialSummaryPage` / `handleInitialSummaryPage` (row 72), `handleDataCollectionTimingPage` (75-77 + 79-81; deletes all existing entries first), `handleOEAdditionalInfoPage` (86).

**Five "template workflow" handlers** — each in its own file:
- `handleResourceEstimatesPage` (`/resource-estimates`) — source `C_Resource_Estimates`, rows 24-N, 4 cols.
- `handleLogisticalRequirementsPage` (`/logistical-requirements`) — source P1PA-R, **5-row repeating blocks starting at row 117, separated by 1 blank row, terminated by `"COPY AND REPEAT"` sentinel**, hard cap 200 rows.
- `handleAppraisalConstraintsPage` (`/appraisal-constraints`) — source `C_AppraisalConstraints`, 3 cols; col A is shared-string dropdown.
- `handleRiskIdentificationPage` (`/risk-identification`) — source `C_RiskIdentification`, 7 cols; cols B, E, G are shared-string dropdowns.
- `handleConflictsOfInterestPage` (`/conflicts-of-interest`) — source `C_COI`, 6 cols; cols A, D, F are dropdowns.

### 4.6 Template workflow (recurring pattern)

For pages where CAS itself uses an Excel upload, the handler follows this shape:

1. Delete existing rows in CAS (loop on `a[href*="ConfirmDelete"]` or `?handler=ConfirmRemove`; safety cap 10-30).
2. Download CAS's template `.xlsx` (preserves styles + data validations).
3. Read the source rows from the project workbook (cell-by-cell, via `resolveCellValue`).
4. Spawn Python (`helpers/injectTemplate.py`) via `execSync`, passing the column config + row data as JSON. Python writes into the downloaded template, preserving shared-string dropdowns (`"ss"` columns) and inline-text columns.
5. Upload via file input + a submit button found by data-test/text search.

### 4.7 Deletion loop pattern

Used wherever existing CAS rows need to be cleared before upload:

```
loop up to N (10 / 20 / 30):
  find first a[href*="ConfirmDelete"] (or ?handler=ConfirmRemove)
  if none → break
  click it
  on the confirmation page, click the Confirm/Submit button
  wait for navigation
```

The safety cap varies per handler; if you change one, look for the magic number constant near the top of the method.

### 4.8 Navigation helpers

- `clickNextButton()` — the CAS "Next" button is an `<a class="button blue-button">` containing an SVG `<polygon points="80,60 ...">`. The handler walks all blue buttons and picks the one whose polygon points start with `"80,60"`.
- `clickSaveButton()` — tries a list of selectors (`button[data-test="button-update-appraisal"]`, `button[data-test="button-add-edit-org"]`, `button[data-test*="update"]`, etc.), then falls back to text search for "update"/"save"/"add target".
- `clickAddEditButton(selector, name)` and `clickAddButton(text)` — for the various Add/Edit affordances.
- `savePageHtml(path, suffix)` — dumps `page.content()` into `html_logs/{timestamp}_{path}_{suffix}.html`. Called before and after every page.

### 4.9 Skipping vs handlers vs generic loop

A field can be opted out of automation in three ways:

- **`CAS_Type === 'skip'`** in the map — generic loop short-circuits as success.
- **`CAS_Selector === 'SKIP'` or `'HANDLER'`** — generic loop filters it out. `'HANDLER'` indicates "a tier-2 handler covers this field"; `'SKIP'` indicates "do not touch this on the web side".
- **A tier-2 handler** for the page — generic loop is skipped entirely for that page.

## 5. The scraper (`scraper.ts`)

Not part of the populate workflow. Run via `npm start`. It walks an explicit list of `PAGES_TO_SCRAPE` (`scraper.ts:70-111`), saves each page's HTML + screenshot, extracts form field metadata (`name`, `label`, `type`, options, etc.), and writes:

- `cas_form_fields.json` — final structured output
- `cas_form_fields_intermediate.json` — saved after each page so a crash doesn't lose work
- `html/*.html`, `screenshots/*.png`

Use this when CAS changes its forms and `_xlsCasMap` needs to be regenerated/updated. The handcrafted `_xlsCasMap_MASTER.xlsx` is then maintained against the scraper's output.

`scraper.ts` has its own copy of CONFIG loading and was already cwd-relative (`'../cas-project-config.json'`) before the May 2026 refactor.

## 6. The skills (Excel ← startup file)

### 6.1 `skill-setup-base-cas-plan` — `setup_cas_plan.py`

Driven by the `_FieldMap` sheet inside the **target** workbook (not `_xlsCasMap`). Schema:

| Col | Field | Notes |
|---|---|---|
| A | Sheet     | Target sheet name |
| B | Heading   | Section heading in source (for disambiguation) |
| C | FieldName | Field label, or `[HEADING]` for section markers |
| D | Row       | Target row number |
| E | Type      | `heading` / `value` / `formula` |
| F | Aliases   | Comma-separated alternate names for source lookup |
| G | Notes     | Audit results live here |

Processing rules:
- `heading` rows: skipped (structural only).
- `formula` rows: **never overwritten** — formula cells in the target are protected.
- `value` rows: build a (heading, fieldname) → cell index from source sheets, then copy. Try field name first, then each alias in order, then field name without heading.
- Sheets handled specially: `Project&Support` does a full row copy by `WorkID` (cols A-AB); `Staff` copies only cols A/C/D and preserves formulas in B/E, only processes `p#` and `s#` WorkIDs.
- If a value is found in the source at a row that differs from the target row, the source cell is marked **yellow fill / red bold text** as an anomaly (but the value is still copied).
- `--backup` snapshots both files before editing. `--report` writes a human-readable text report.

### 6.2 `skill-audit-cas-plan` — `audit_cas_plan.py`

Validates `_FieldMap` against the actual sheet content. For each row: does the sheet exist? Does col A at the specified row match `FieldName`? Is the cell type (value vs formula) what `_FieldMap` claims? Results go into col G with conditional formatting (green ok / red mismatch / yellow warning). `--dry-run` previews without writing.

### 6.3 Launching

Both skills have a `.bat` wrapper that reads `files.source` / `files.target` from `cas-project-config.json` via a Python one-liner (no hardcoded paths after the May 2026 refactor):

```bat
set "PROJECT_CONFIG=%~dp0..\..\cas-project-config.json"
for /f "usebackq delims=" %%I in (`python -c "import json; print(json.load(open(r'%PROJECT_CONFIG%','r',encoding='utf-8'))['files']['target'])"`) do set "TARGET=%%I"
```

The report file lands next to the target file with a `YYYYMMDD` date stamp.

## 7. Data flow (end-to-end)

```
┌─────────────────────────┐  setup_cas_plan.py  ┌─────────────────────────────┐
│ Demix / startup .xlsm   │ ──────────────────► │  CAS Plan .xlsm (target)   │
│ (customer-supplied)     │   _FieldMap inside  │  – Agreement 3 / Planning   │
└─────────────────────────┘   target drives it  │  – StartupInfo / Staff      │
                                                │  – Project&Support          │
                                                │  – P1-OrgScope / P1PA-R     │
                                                │  – C_Resource_Estimates etc │
                                                │  – _xlsCasMap (fallback)    │
                                                └──────────────┬──────────────┘
                                                               │  populator.ts
                                                               │  driven by
                                                               │  _xlsCasMap_MASTER.xlsx
                                                               ▼
                                                 ┌─────────────────────────┐
                                                 │  CAS web portal         │
                                                 │  cas.cmmiinstitute.com  │
                                                 └─────────────────────────┘
```

## 8. Logging & artifacts

- `cas-scraper/populate_log.json` — `PopulateLog[]`, one entry per page. Includes per-field `status` (success / failed / skipped), `error`, and any free-text `userFeedback` typed during a `debugMode` pause.
- `cas-scraper/html_logs/` — `{ISO_timestamp}_{pagePath}_{before|after}.html`.
- `cas-scraper/screenshots/`, `cas-scraper/html/` — populated by `scraper.ts`, not the populator.
- `cas-scraper/.browser-data/` — persistent Chromium profile. Kept around so login survives across `npm run populate` invocations. Safe to delete if you need a clean session.

## 9. Key invariants / gotchas

- **Two FieldMap sheets exist** and they are different. `_FieldMap` is inside each project workbook and drives the Python skills. `_xlsCasMap` lives in a MASTER xlsx and drives the TypeScript populator. Conflating these is the most common source of confusion.
- **`continueFromPage` is sticky.** If the populator crashed at `/objective-evidence/initial-summary`, `continueFromPage` is probably still set. Clear it (set to `""`) before a fresh run, or it will skip everything earlier.
- **The page-routing block in `processCurrentPage` is order-sensitive.** Earlier `if` checks short-circuit later ones. When adding a new handler, place it next to its siblings and look for substring overlap (e.g. `/org-unit-sampling-factors` vs `/org-unit-sampling-factor-values`).
- **`loadExcelData` extra rows must be kept in sync.** A tier-2 handler that reads from a row not present in `_xlsCasMap` only works if that row is in the `extraRows` array for its sheet.
- **`_xlsCasMap_MASTER.xlsx` takes precedence over the per-project copy.** If you edit the project workbook's `_xlsCasMap`, the populator may still load MASTER. Update MASTER too, or move/delete it.
- **Mixin interface duplication.** When adding a handler method, both `applyXxxHandlers(CASPopulator)` (line 940-ish) and the `interface CASPopulator` declaration (lines 949-993) must list it; otherwise TS won't type-check `this.handleXxx()` from inside `processCurrentPage`.
- **Formula cells are sacred in `setup_cas_plan.py`.** Never overwrite them; the audit script will help flag drift.
- **CAS uses `__RequestVerificationToken` CSRF tokens** on form submits. Always navigate via Puppeteer; do not raw-POST.
- **Windows mount + JSON writes.** On the dev workspace, in-place edits that shorten a JSON file can leave trailing null bytes that break `JSON.parse`. Prefer delete-then-rewrite for `.json` files on this filesystem (see `README.md`).
- **Duplicate handler files at `cas-scraper/` root are dead code.** Only `cas-scraper/handlers/*.ts` is imported.
- **Field-population idempotency.** All populators check the current value first. Re-running the populator on an already-completed page should be a no-op.
- **The "Next" button is identified by SVG polygon points** (`points="80,60 ..."`), not by text. If CAS changes its button design, `clickNextButton()` is the first thing to update.
- **Target-file version drift.** `files.target` in `cas-project-config.json` hardcodes the full filename including version suffix (e.g. `10_102v08_*_CAS_Plan.xlsm`). Every time you save a new revision of the plan workbook, you must update the config. ExcelJS will throw `File not found` if the version digits don't match. If this keeps biting, consider switching `files.target` to a glob pattern and resolving it at startup.
- **Two prompts pause the loop, governed by different rules.** `askForFeedback()` only runs when `debugMode: true`. `askWhenStuck()` runs whenever no Next button is found, **regardless of `debugMode`** — it's the dead-end safety net. If you ever change `debugMode` semantics, be careful not to swallow the `askWhenStuck` path.
- **`reprocess` does not bypass dispatch.** Choosing `c` re-runs the URL through the full `processCurrentPage` pipeline including `skipPages` and `finalPage` checks. If you manually navigate to `/sample-scope` and pick `c`, the populator will exit via `handlePhaseComplete`.
- **Excel data is loaded once at startup.** Editing the workbook between pages does not update `excelData` until you restart `npm run populate`. There is currently no in-session reload.

## 10. Common edit tasks

- **Add a field to an existing page:** add a row to `_xlsCasMap_MASTER.xlsx` (8 columns), make sure the target sheet has a value at the indicated row, and re-run `npm run populate`. If the page has a tier-2 handler, set `CAS_Selector = 'HANDLER'` and extend the handler instead.
- **Add a handler for a new page:** create `cas-scraper/handlers/{name}.handler.ts` exporting `apply{Name}Handler(cls)`, add `apply{Name}Handler(CASPopulator)` to `populator.ts` line 940-ish, add the method signature to the `interface CASPopulator` block, add the dispatch case inside `processCurrentPage()` (mind the ordering), and update `loadExcelData` extras if you read non-mapped rows.
- **Start a new project:** copy a previous `cas-project-config.json`, update `project.casId`, `files.source`, `files.target`. Drop the project workbook into the target path. Clear `cas.continueFromPage`. Run `skill-setup-base-cas-plan` to populate from demix, then `skill-audit-cas-plan` to verify, then `npm run populate` for the web side.
- **Rotate credentials:** edit `cas_helper/.secrets/keys.json`. No code change needed.
- **Debug a single page:** set `cas.continueFromPage` to that page and `cas.debugMode: true`, then `npm run populate`. After processing, the `askForFeedback` prompt appears; type `c` to re-process the same page after tweaking the workbook in the browser, `a` to exit, or any text to save feedback to `populate_log.json`.
- **Work a page that's outside the auto-Next chain:** start the populator anywhere, let it dead-end (or set `continueFromPage` to a page with no Next button), and when the `askWhenStuck` prompt fires, manually navigate the open Puppeteer window to the page you want, then press Enter / type `c`. Repeat as needed.
