# BASE_source — specification

Architecture of the BASE WinForms desktop tool as it exists today, its gotchas, and how it
binds to the master spec. Audience: an AI coder making edits without re-deriving everything.

Binds to: root `../specification.md` (canonical CMMI domain model §1, Azure roadmap §3).

## 1. Project shape

- `net48` WinForms (`WinExe`, namespace `BASE`). Entry: `Program.cs` → `Main` form.
- UI: `Main.cs` (~3,427 lines, the single form class), `Main.Designer.cs`, `Main.resx`.
- `Data/` — the CMMI domain types (see master spec §1 for the mapping). `PersistentData.cs`
  serialises them to/from XML files under a `BASE\` folder.
- `Excel/` — Excel Interop helper(s).
- Office Interop references: Excel 15.0, Word 15.0, PowerPoint 15.0, Outlook 14.0.
- `App.config`, `packages.config`, `BASE.sln`, `BASEIcon.*`, ClickOnce → `../BASE_install/`.

## 2. `Main.cs` structure

One form class. Regions seen: `#region globals` (≈ lines 37–125) and
`#region Process Worksheet options` (≈ 2531–2643). Everything else is a flat list of
event handlers + `Helper*` methods. Notable handlers:

- **CAS Plan / XML**: `btnLoadXML_Click`, `btnSaveXML_Click`, `btnOpenBaseCASPlan_Click`,
  `btnSelectPlan_Click`.
- **Schedule**: `buttonGenerateSchedule_Click`, `btnLoadSchedule2_Click`.
- **Main workbook setup**: `btnSetupMain_Click`, `btnSetupMain2_Click`, `btnSelectMainTool_Click`, `btnSelectWorkingDir_Click`.
- **Interviews**: `btnInsertInterviews_Click`, `btnInsertInterviews2_Click`.
- **OE / OEdb**: `btnOEdb_Click`, `btnOEdbMain_Click`, `btnOEdbSource_Click`, `btnMergeSourceToMain_Click`, `btnMergeSources2_Click`, `btnImportOE_Click`, `btnDemixOEMerge_Click`, `btnBuildOUMaps_Click`.
- **Out-of-scope rows**: `btnHideOoSRows_Click`, `btnHideOoS2_Click`, `HelperHideRows1/2`.
- **Findings**: `btnExtractFindings_Click`, `HelperExtractFindings`, `HelperExtractFindingsDemixOE`.
- **IIGOV**: `btnIIGOVrating_Click`, `btnTestAndEngl2_Click`, `HelperIIGOVcharacterization`.
- **Model / questions / toolkit**: `btnImportModel_Click`, `btnSelectQuestionFile_Click`, `btnGenerateFullTool_Click`, `btnBuildTmpDictionary_Click`.
- **Row processing**: `ProcessRowsUsingExcel`, `ProcessRowsUsingObject` (the two strategies in the Process-Worksheet region).

## 3. Persistence

Domain objects (master spec §1) ↔ XML via `PersistentData`. File name constants live in
`Main.cs` globals, e.g. `CasPlanFileXML.xml`, `OEdbFile.xml`,
`TargetOEdbImportFileXML.xml`, `TargetQuestionModelFileXML.xml`,
`TargetDataReferenceFileXML.xml`, `TargetPresentationFileXML.xml`,
`TargetToolkitMasterFileXML.xml`. All under a `BASE\` subfolder.

## 4. Gotchas

- **Hard-coded path.** `cPath_start` points at a literal 2020 path
  (`C:\Users\PietervanZyl\…\2020-12-11 (A5) … Goshine Tech`). Any run against a different
  appraisal folder depends on the working-dir selectors, not this constant. (`S02`.)
- **Const-heavy layout assumptions.** Dozens of column/row constants in `#region globals`
  (`cProjectHeadingStartRow`, `cOEDatabaseHeadingStartRow`, `cPAtestColumn`,
  `cXXWeaknessCol`, `cIIandGOV*`, `cMostPAStartRow/EndRow`, …) bake the workbook layout
  into code. A worksheet whose columns moved will silently mis-process. Treat these as a
  fragile contract with the template workbooks.
- **Office Interop required.** No headless path — needs Excel/Word/PowerPoint/Outlook
  installed at the referenced versions. COM lifetime / orphan-process care applies.
- **Monolith edit risk.** ~3,400 lines in one form = high re-read cost and merge-collision
  risk for any AI edit. (`S03`.)

## 5. Binding to master spec / migration

Today BASE owns its data as local XML. Per master spec §3 Initiative **I-4**, BASE will
eventually read from the centralised Azure DB (behind a feature flag, defaulting to local
XML). Until the Azure schema (`R01`/I-1) exists, keep persistence as-is.

## 6. Common edit tasks

- Add a worksheet-processing step → new `btn*_Click` + `Helper*`; reuse `ProcessRowsUsing*`.
- Change a workbook layout assumption → find the relevant `c*Col`/`c*Row` const first.
- Touch persistence → go through `PersistentData`, keep the XML file-name constants in sync.
