# Audit CAS Plan

A Python tool to validate the `_FieldMap` sheet against actual sheet content in a CAS Plan workbook.

## Overview

This tool audits each entry in the `_FieldMap` sheet and verifies that:
- The field exists at the specified row
- The field name matches what's in Column A
- The cell type (value/formula) is correct

Results are written to Column G with color coding.

## Installation

### Prerequisites
- Python 3.8 or higher
- openpyxl library

### Install Dependencies
```bash
pip install openpyxl
```

## Usage

### Basic Usage
```bash
python audit_cas_plan.py --file "path/to/CAS_Plan.xlsm"
```

### With Report File
```bash
python audit_cas_plan.py --file "path/to/CAS_Plan.xlsm" --report "audit_report.txt"
```

### Dry Run (Preview)
```bash
python audit_cas_plan.py --file "path/to/CAS_Plan.xlsm" --dry-run
```

## Results

After running, check the `_FieldMap` sheet:

| Column G Value | Color | Meaning |
|----------------|-------|---------|
| `ok` | Green | Field matches definition |
| Error message | Red | Mismatch found |
| Warning | Yellow | Other issue |

## Example

```
$ python audit_cas_plan.py -f "CAS_Plan.xlsm"

============================================================
CAS Plan Audit - Complete
Total fields audited: 153
  OK: 152
  Mismatches: 1
  Warnings: 0
============================================================
```

## Troubleshooting

### "File not found"
- Check the file path is correct
- Ensure the file exists

### "openpyxl not found"
```bash
pip install openpyxl
```

### "_FieldMap sheet not found"
- The file must contain a `_FieldMap` sheet
- Generate it first using the setup skill

## Version History

- **1.0.0** (2026-02-14) - Initial release
