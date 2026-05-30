# CAS Web Automation Project

Tools for automating CMMI CAS (CMMI Appraisal System) form population and data extraction.

## Project Configuration

The project uses a central configuration file at `cas_helper\cas-project-config.json` (one level above `cas-scraper\`):

```json
{
  "project": {
    "name": "NASA SYSTEC Appraisal",
    "casId": "81846"
  },
  "files": {
    "source": "path/to/source.xlsm",
    "target": "path/to/target.xlsm"
  },
  "utilities": {
    "cas-scraper": "C:\\WorkDir-Claude\\cas-scraper"
  },
  "skills": {
    "skill-audit-cas-plan": "C:\\WorkDir-Claude\\skills\\skill-audit-cas-plan",
    "skill-setup-base-cas-plan": "C:\\WorkDir-Claude\\skills\\skill-setup-base-cas-plan"
  }
}
```

## Tools

### 1. Form Scraper (`scraper.ts`)

Extracts form field information from CAS pages.

```bash
npm run scrape
```

### 2. Form Populator (`populator.ts`)

Populates CAS forms from Excel data using field mappings.

```bash
npm run populate
```

## Prerequisites

- Node.js 18+ installed
- npm installed

## Installation

```bash
cd C:\WorkDir-Claude\cas-scraper
npm install
```

## Credentials Setup

Credentials are stored in `cas_helper\.secrets\keys.json` (separate from config for security, and `.secrets\` should be gitignored):

```json
{
  "cas": {
    "email": "your@email.com",
    "password": "your_password",
    "staySignedIn": "yes"
  }
}
```

**Options:**
- `staySignedIn`: `"yes"` or `"no"` - controls the "Remember me" checkbox on login

**Note:** Environment variables (`CAS_EMAIL`, `CAS_PASSWORD`) are still supported as fallback.

## Field Mappings

The `fieldmap_cas.json` file maps Excel rows to CAS form fields:

```json
{
  "Row": 9,
  "Sheet": "P1-OrgScope",
  "FieldLabel": "Virtual Phase 1: Plan and Prepare",
  "CAS_Page": "/name-and-type",
  "CAS_Selector": "input[data-test='input-virtual-selection_5']",
  "CAS_FieldName": "virtual-phase-1",
  "CAS_Type": "checkbox",
  "Notes": "Check if Yes. HTML id=5, use data-test attribute"
}
```

### Selector Strategies

- **Standard selectors**: `#appraisal-name`, `select[name="TimeZone"]`
- **Data-test attributes**: `input[data-test='input-virtual-selection_5']` (preferred for CAS checkboxes)
- **Numeric ID fallback**: The populator handles `#5`, `#6`, `#7` by using `getElementById()`

## Excel Data

The `excel_data.json` file contains values to populate:

```json
{
  "P1-OrgScope": {
    "3": "Appraisal Name",
    "9": "Yes",  // Virtual Phase 1 checkbox
    "10": "Yes", // Virtual Phase 2 checkbox
    "11": "Yes"  // Virtual Phase 3 checkbox
  }
}
```

## Output Files

| File | Description |
|------|-------------|
| `cas_form_fields.json` | Complete field extraction by page |
| `cas_summary.json` | Page and field count summary |
| `fieldmap_cas.json` | Excel-to-CAS field mappings |
| `excel_data.json` | Values to populate |
| `populate_log.json` | Population results log |
| `screenshots/` | Screenshots for debugging |

## Troubleshooting

### Checkboxes not found
CAS uses numeric IDs (e.g., `id="5"`) which are invalid CSS selectors. Solutions:
1. Use `data-test` attribute: `input[data-test='input-virtual-selection_5']`
2. The populator auto-detects numeric IDs and uses `getElementById()`

### Login issues
1. Check credentials in environment variables
2. Review screenshots in `screenshots/` folder
3. The CAS portal may have session timeouts

## Notes

- Browser runs in visible mode for manual review
- Interactive mode pauses after each page for feedback
- Use `[s]` or `[q]` to stop, `[Enter]` to continue
