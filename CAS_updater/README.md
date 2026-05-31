# CAS_updater — component README

TypeScript/Node tooling that reads a project's **CAS Plan workbook** (`CASxxxx.xlsm`) and
**populates the CMMI CAS system** (`cas.cmmiinstitute.com`). Formerly the standalone
`cas_scraper` / `cas_helper` repo; now a component of cmmitools.

Kanban prefix: **C**. The detailed architecture lives in `specification.md`; day-to-day
usage of the populator/scraper lives in `cas-scraper/README.md`. Shared domain model and
Azure roadmap: root `../specification.md`.

## What it does (two workflows — see `specification.md` §1)

- **Web ← Excel** (`cas-scraper/`, TypeScript): drives a Puppeteer session against the CAS
  site and fills web form fields from the CAS Plan workbook. `npm run populate`.
- **Excel ← startup file** (`skills/`, Python): populates a project CAS Plan workbook from a
  customer "startup" workbook.

## Layout

```
CAS_updater/
├── README.md                ← this file
├── specification.md         ← dense architecture / invariants / gotchas (read before editing)
├── cas-project-config.json  ← per-project paths + CAS settings (single source of truth)
├── .secrets/keys.json       ← CAS credentials (GITIGNORED — never commit)
├── cas-scraper/             ← TS web populator + scraper (npm run populate / start / build)
└── skills/                  ← Python: Excel-side setup & audit skills
```

## Run

```
cd CAS_updater/cas-scraper
npm install          # deps are not committed
npm run populate     # web ← Excel populator
```

Credentials go in `.secrets/keys.json` (gitignored); per-project paths in
`cas-project-config.json`. See `cas-scraper/README.md` for the full setup.

## Future direction

Move from driving the live CAS site to **POSTing parsed data to the `CMMI_online` API**
(`CMMI_online` then owns the Azure writes). Tracked as Initiative I-3 in the master spec →
task `_kanban/C01_populate_via_CMMI_online_API` (this side) + `CMMI_online/_kanban/O03`
(the endpoint).

## Security note

`.secrets/keys.json` was **committed in the old `cas_helper` repo history** (remote
`PieterVZ-Demix/cas_helper.git`). It is gitignored here, but the old credentials still
exist upstream — rotate them.

---

> The section below is the original `cas_helper` working notes — a refactor changelog plus
> hard-won operational gotchas. Kept for reference; new architecture lives in
> `specification.md`.

## Appendix — config & paths refactor (history & gotchas)

Summary of the config-path cleanup that fixed the spurious "Credentials not set!" error from `npm run populate`.

### The original bug

`cas-scraper/config.ts` looked for the project config at `C:/WorkDir-Claude/cas-project-config.json`, but the file actually lives at `C:/WorkDir-Claude/cas_helper/cas-project-config.json`. Because the file was never found, `projectConfig` stayed `null`, the `keysFile` pointer was never followed, and `CONFIG.email` fell through to `''`. The populator then printed the misleading "configure credentials in `C:\WorkDir-Claude\keys.json`" message — even though `keys.json` existed and was fine.

Compounding the mess, the `try/catch` in `config.ts` swallowed all errors silently with `// Config optional, continue with defaults`, so nothing surfaced.

### New layout

All paths are now relative to `cas-scraper/` (i.e. wherever `npm run populate` runs from):

```
cas_helper/
├── cas-project-config.json     ← single source of truth for project paths
├── .secrets/
│   └── keys.json               ← credentials (gitignored)
└── cas-scraper/
    ├── config.ts               ← reads ../cas-project-config.json and ../.secrets/keys.json
    └── ...
```

`config.ts` resolves both files relative to `__dirname`, so it works regardless of where you invoke it from.

### Files changed

`cas-scraper/config.ts` — fully rewritten. Resolves `../cas-project-config.json` and `../.secrets/keys.json` relative to its own location. Warns (instead of silently swallowing) when files are missing or malformed. Exports `projectConfigPath` and `keysPath` so other modules can show accurate paths in error messages.

`cas-scraper/test-download.ts` — same hardcoded-path bug fixed.

`cas-scraper/populator.ts`, `cas-scraper/scraper.ts` — error messages now print `CONFIG.keysPath` (the actual file location) instead of a hardcoded string.

`cas-scraper/README.md` — Configuration and Credentials sections updated to reflect new paths.

`cas-scraper/cas-project-config.json` — deleted (it was a duplicate of the real one in `cas_helper/`).

`cas_helper/cas-project-config.json` — removed the now-redundant `keysFile` field; `config.ts` hardcodes the keys location.

`skills/skill-audit-cas-plan/run_audit.bat` — no longer hardcodes the NASA `BASE_DIR`. Reads `files.target` from `cas-project-config.json` via a Python one-liner; `REPORT` is written next to the target file with a `YYYYMMDD` date stamp.

`skills/skill-setup-base-cas-plan/run_setup.bat` — same treatment, reads `files.source` and `files.target` from the project config.

### Files deliberately left alone

`cas-scraper/run_scraper.bat` — nothing in it maps cleanly to `cas-project-config.json`. The interactive `CAS_EMAIL` / `CAS_PASSWORD` prompts (lines 17–22) are now obsolete (since `keys.json` covers credentials) but harmless as a fallback.

`cas-scraper/populator - Copy.ts`, `cas-scraper/populator_3.ts` — appear to be dev backups / older revisions. Not updated to avoid touching code paths that aren't in active use.

`cas_helper/filelist.txt` and `filelist - Copy.txt` — reference the old `C:\WorkDir-Claude\keys.json` path but appear to be one-off file lists, not source-of-truth config.

### Verification

A smoke test run from `cas-scraper/` confirmed:

```
projectConfigPath: .../cas_helper/cas-project-config.json
keysPath:          .../cas_helper/.secrets/keys.json
email loaded:      pwj...
password loaded:   (set)
staySignedIn:      true
appraisalId:       83325
continueFromPage:  /appraisal-participants
excelFile:         .../Beijing PDE Info Tech_CAS_Plan.xlsm
```

`npm run populate` from `cas-scraper/` now proceeds past the credentials check.

### Gotcha worth remembering

On this Windows-mounted folder, in-place edits that shorten a file left trailing null bytes (which broke `cas-project-config.json` as valid JSON the first time it was edited). When editing JSON or other strict-parse files here, prefer delete-then-rewrite over partial edits.

### Subsequent changes

#### Interactive prompts in the populator

The populator now has two prompts that pause the main loop:

**`askForFeedback`** runs after each page when `cas.debugMode: true`. Three explicit options plus free-text feedback:

```
Options:
  a               - Exit
  b  (or Enter)   - Continue to next page (default)
  c               - Re-read current browser URL and process it
  [any text]      - Save as feedback, then continue
  [any text] a    - Save feedback, then exit
  [any text] c    - Save feedback, then re-process current URL
```

Option `c` skips the auto-Next step. The next iteration reads `page.url()` fresh, so you can manually navigate the Puppeteer window to any page and have the populator process it on the next pass.

**`askWhenStuck`** runs whenever `clickNextButton()` returns false — **regardless of `debugMode`**. This is the dead-end fallback: when CAS doesn't expose a Next button (e.g. on `/appraisal-participants` with no fields to populate), the populator pauses and asks you to navigate manually instead of silently quitting.

```
Options:
  a               - Exit
  c  (or Enter)   - Process the page currently in the browser (default)
  [any text]      - Save as feedback, then process current page
  [any text] a    - Save feedback, then exit
```

Both prompts feed back into the same main loop, so a session can mix auto-advance pages with manually-navigated ones freely.

#### Target-file version drift

`files.target` in `cas-project-config.json` hardcodes the full plan-workbook filename including the version digits (e.g. `10_102v08_Beijing PDE Info Tech_CAS_Plan.xlsm`). Every time you save a new revision, you must update the config — ExcelJS throws `File not found` otherwise. Caught this one when the configured target was still pointing at `10_102v04_…` after the workbook had been saved through v08 and the older versions moved to `_delete/`.

If this keeps biting, options worth considering: switch `files.target` to a glob, store just the version digits and template the filename, or have `populator.ts` auto-pick the highest-numbered match in the project folder.

#### Excel data is a one-shot load

`loadExcelData()` runs once during `init()`. Editing the workbook between pages does not propagate until you restart `npm run populate`. No in-session reload yet.
