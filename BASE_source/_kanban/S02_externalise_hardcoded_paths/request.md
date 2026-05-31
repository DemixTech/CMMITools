# S02 — Externalise hard-coded paths and layout constants

> Status: **backlog**. Pre-flight cleanup so later edits are safe.

## Problem

`Main.cs` hard-codes a 2020 appraisal path (`cPath_start = C:\Users\PietervanZyl\…\Goshine Tech`)
and dozens of workbook column/row constants in `#region globals`. Any AI edit has to step
around these or risk silent regression on a different appraisal folder/template.

## Scope

- Move `cPath_start` and other environment paths into `App.config` (or a typed options
  class), defaulting sensibly and overridable per appraisal.
- Group the workbook layout constants (`c*Col`, `c*Row`, `cMostPA*`, `cIIandGOV*`, …) into
  a typed `WorkbookLayout` options object, documented as the contract with the template
  workbooks (see `specification.md` §4).

## Done when

- No literal user-specific path remains in `Main.cs`.
- Layout constants are in one named place, referenced everywhere, with a comment block
  explaining the template-workbook contract.
- App still builds and a smoke run against a sample appraisal folder works.
