# Skill: Audit CAS Plan

## Description
This skill audits the `_FieldMap` sheet in a CAS Plan workbook against the actual sheet content. It validates that each field definition matches the actual data in the corresponding sheet.

## Version
1.0.0

## What It Checks

For each entry in `_FieldMap`, the audit verifies:

| Check | Description |
|-------|-------------|
| Field Name | Does Column A at the specified row match the FieldName? |
| Cell Type | Is the cell a formula/value as expected? |
| Sheet Exists | Does the referenced sheet exist? |

## Results Written to Column G

| Result | Style | Meaning |
|--------|-------|---------|
| `ok` | 🟢 Green | Field matches _FieldMap definition |
| Mismatch details | 🔴 Red | Name or type doesn't match |
| Warning | 🟡 Yellow | Other issue (sheet not found, etc.) |

## Use Cases

1. **After generating _FieldMap** - Verify the auto-generated mapping is correct
2. **After editing sheets** - Check if row numbers have shifted
3. **Before running setup** - Ensure target structure matches expectations
4. **Template validation** - Verify a new template matches the expected structure

## Usage

### Command Line
```bash
python audit_cas_plan.py --file "path/to/CAS_Plan.xlsm"
```

### With Report
```bash
python audit_cas_plan.py --file "path/to/CAS_Plan.xlsm" --report "audit_report.txt"
```

### Dry Run (Preview Only)
```bash
python audit_cas_plan.py --file "path/to/CAS_Plan.xlsm" --dry-run
```

### Arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `--file`, `-f` | Yes | Path to CAS Plan file with _FieldMap |
| `--dry-run`, `-n` | No | Preview without writing results |
| `--report`, `-r` | No | Save detailed report to file |
| `--verbose`, `-v` | No | Enable detailed logging |

## Example Output

### Console Output
```
============================================================
CAS Plan Audit - Starting (v1.0.0)
============================================================
File: C:\...\10_97v01_Base Camp Artemis_CAS_Plan.xlsm
Dry run: False
Found _FieldMap with 153 entries
Row 115: StartupInfo!=Planning!A61 - Name mismatch: expected '=Planning!A61', found 'Appraisal objectives'
============================================================
CAS Plan Audit - Complete
Total fields audited: 153
  OK: 152
  Mismatches: 1
  Warnings: 0
============================================================
```

### Column G in _FieldMap
After running, Column G will show:
- `ok` (green background) for matching fields
- Error message (red background) for mismatches

## Fixing Mismatches

When mismatches are found:

1. **Name mismatch** - Either:
   - Update the FieldName in _FieldMap to match actual
   - Or update the actual sheet to match _FieldMap

2. **Type mismatch** - Either:
   - Update the Type in _FieldMap (value/formula)
   - Or check if the cell formula was accidentally deleted

3. **Row shifted** - Update the Row number in _FieldMap

## Integration with Other Skills

Run this audit:
- **Before** `setup_cas_plan.py` - Ensure target structure is valid
- **After** modifying template - Verify changes didn't break mapping

## File Structure
```
skill-audit-cas-plan/
├── SKILL.md              # This documentation
├── audit_cas_plan.py     # Main Python script
├── run_audit.bat         # Windows batch launcher
└── README.md             # User instructions
```
