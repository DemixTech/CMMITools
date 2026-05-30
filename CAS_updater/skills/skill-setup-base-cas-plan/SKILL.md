# Skill: Setup Base CAS Plan

## Description
This skill automates the population of a CAS (CMMI Appraisal System) Plan workbook from a startup/demix source file. It uses the `_FieldMap` sheet embedded in the TARGET file to define field structure and find corresponding fields in the source.

## Version
2.0.0

## Key Features
- **_FieldMap Driven**: Reads field definitions from target file's `_FieldMap` sheet
- **Smart Field Matching**: Matches by heading + field name, supports aliases
- **Formula Protection**: Never overwrites formula cells in target
- **Anomaly Detection**: Marks source rows with different positions (yellow/red)

## How It Works

1. **Load _FieldMap** from target file
2. **Build source index** by scanning source sheets for heading + field combinations
3. **For each _FieldMap entry**:
   - Skip if type = "formula" or "heading"
   - Find matching field in source (by heading + name, or aliases)
   - Copy value to target row specified in _FieldMap
   - Mark anomalies if source row ≠ target row

## Required Files

| File | Requirements |
|------|--------------|
| Target (.xlsm) | Must contain `_FieldMap` sheet |
| Source (.xlsm) | Startup/demix file with data |

## _FieldMap Structure

The target file must have a `_FieldMap` sheet with:

| Column | Content |
|--------|---------|
| A: Sheet | Sheet name |
| B: Heading | Section heading |
| C: FieldName | Field label (or `[HEADING]`) |
| D: Row | Target row number |
| E: Type | `heading`, `value`, or `formula` |
| F: Aliases | Alternate field names (comma-separated) |
| G: Notes | Audit results / notes |

## Processing Rules

| Type | Action |
|------|--------|
| `heading` | Skipped (structure only) |
| `formula` | Skipped (never overwrite formulas) |
| `value` | Copy from source if found |

## Alias Support

If a field has a different name in source, add aliases in Column F:

```
FieldName: CMMI Appraiser Team Lead ID
Aliases: CMMI v2.0 Appraiser Team Lead ID, ATL ID
```

The script will try:
1. Exact field name match
2. Each alias in order
3. Field name without heading context

## Usage

### Command Line
```bash
python setup_cas_plan.py \
  --source "path/to/SYSTEC-April-CAS Plan-Demix.xlsm" \
  --target "path/to/10_98v01_Base Camp Artemis_CAS_Plan.xlsm" \
  --backup \
  --report "setup_report.txt"
```

### Arguments
| Argument | Required | Description |
|----------|----------|-------------|
| `--source`, `-s` | Yes | Source startup/demix file |
| `--target`, `-t` | Yes | Target CAS Plan file (with _FieldMap) |
| `--backup`, `-b` | No | Backup both files before modification |
| `--report`, `-r` | No | Save detailed report |
| `--dry-run`, `-n` | No | Preview without changes |
| `--verbose`, `-v` | No | Detailed logging |

## Visual Markers in Source File

| Condition | Marking | Meaning |
|-----------|---------|---------|
| Row mismatch | Yellow bg / Red text | Source row ≠ Target row (value still copied) |

## Workflow

1. **Run audit** first: `skill-audit-cas-plan` to verify _FieldMap
2. **Add aliases** in Column F for any field name differences
3. **Run setup**: `skill-setup-base-cas-plan`
4. **Review report** for warnings and anomalies

## Example Report Output

```
======================================================================
CAS PLAN SETUP REPORT (v2.0.0)
======================================================================
Source: ...\SYSTEC-April-CAS Plan-Demix.xlsm
Target: ...\10_98v01_Base Camp Artemis_CAS_Plan.xlsm

SUMMARY
--------------------------------------------------
Total Cells Copied: 45
Warnings: 12
Row Anomalies: 8

SHEET-BY-SHEET RESULTS
--------------------------------------------------
  Agreement 3: 25 cells copied
  Planning: 12 cells copied
  StartupInfo: 8 cells copied
```

## Notes
- Always run audit before setup to verify _FieldMap
- Add aliases for fields with different names in source
- Formula cells are always protected
- Anomalies are marked but values are still copied correctly
