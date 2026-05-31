# Setup Base CAS Plan

A Python tool to populate a CAS (CMMI Appraisal System) Plan workbook from a startup/demix source file.

## Overview

This tool automates the process of transferring data from a source startup/demix Excel file to a target CAS Plan workbook. It processes specific tabs and ensures that formulas in the target file are never overwritten.

## Features

- ✅ Copies values from 5 specific tabs: Agreement 3, Planning, StartupInfo, Project&Support, Staff
- ✅ Preserves formulas in the target workbook
- ✅ Flags potentially overwritten formulas in source with red background/white text
- ✅ Supports dry-run mode for preview
- ✅ Automatic backup creation
- ✅ Detailed reporting

## Installation

### Prerequisites
- Python 3.8 or higher
- pip (Python package installer)

### Install Dependencies
```bash
pip install openpyxl
```

## Usage

### Basic Usage
```bash
python setup_cas_plan.py --source "path/to/startup-file.xlsm" --target "path/to/CAS_Plan.xlsm"
```

### With Backup and Report
```bash
python setup_cas_plan.py \
    --source "C:\WorkDir-Claude\2024-05-04to05-10 (A5) C384400 NASA\00_Startup\SYSTEC-April-CAS Plan-Demix.xlsm" \
    --target "C:\WorkDir-Claude\2024-05-04to05-10 (A5) C384400 NASA\10_97v01_Base Camp Artemis_CAS_Plan.xlsm" \
    --backup \
    --report "setup_report.txt"
```

### Dry Run (Preview Only)
```bash
python setup_cas_plan.py \
    --source "path/to/source.xlsm" \
    --target "path/to/target.xlsm" \
    --dry-run
```

### Command Line Options

| Option | Short | Description |
|--------|-------|-------------|
| `--source` | `-s` | Path to source file (required) |
| `--target` | `-t` | Path to target CAS Plan file (required) |
| `--dry-run` | `-n` | Preview changes without modifying files |
| `--backup` | `-b` | Create backup of target before modification |
| `--report` | `-r` | Path to save detailed report |
| `--json-output` | `-j` | Path to save JSON results |
| `--verbose` | `-v` | Enable verbose logging |

## Processed Tabs

The tool processes these sheets from the source file:

1. **Agreement 3** - Appraisal agreement details
2. **Planning** - Planning and scheduling information
3. **StartupInfo** - Initial startup configuration
4. **Project&Support** - Project and support function details
5. **Staff** - Staff/team member information

## Formula Protection

The tool implements several safeguards:

1. **Target formulas are never overwritten** - If a target cell contains a formula, it is skipped and a warning is logged.

2. **Source formula detection** - If the tool detects that a source cell may have had a formula that was overwritten with a plain value, it marks that cell with:
   - Red background (RGB: 255, 0, 0)
   - White text (RGB: 255, 255, 255)

3. **Change logging** - All changes are logged for review.

## Output

After running, the tool provides:

1. **Console output** - Summary of changes
2. **Modified target file** - With populated data
3. **Modified source file** - With flagged cells (if any)
4. **Report file** (optional) - Detailed text report
5. **JSON output** (optional) - Machine-readable results

## Example Workflow

### For Claude AI Assistant:

```
1. User: "Setup the CAS Plan from the startup file"

2. Claude reads SKILL.md

3. Claude executes:
   python setup_cas_plan.py \
       --source "...\00_Startup\SYSTEC-April-CAS Plan-Demix.xlsm" \
       --target "...\10_97v01_Base Camp Artemis_CAS_Plan.xlsm" \
       --backup \
       --report "setup_report.txt"

4. Claude reviews report and informs user of:
   - Number of cells copied
   - Any warnings (formulas not overwritten)
   - Any flagged cells requiring manual review
```

## Troubleshooting

### "openpyxl not found"
```bash
pip install openpyxl
```

### "Permission denied" error
- Close the Excel files before running
- Check file permissions

### "Sheet not found" warning
- The sheet may have a different name in your file
- Check sheet names match exactly (case-sensitive)

## Support

For issues or questions, please review the SKILL.md documentation or contact the maintainer.

## Version History

- **1.0.0** (2026-02-13) - Initial release
