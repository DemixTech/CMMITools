# CAS Web Interface Field Mapping
## Generated from scraper results: 2026-02-15

## Overview
This document maps fields from the CAS (CMMI Appraisal System) web interface to the 
_FieldMap structure in the CAS Plan Excel workbook.

## CAS Pages Discovered

### 1. Organization Scope (/appraisals/{id}/name-and-type)
| CAS Field | Selector | Type | _FieldMap Field | Notes |
|-----------|----------|------|-----------------|-------|
| Name | #appraisal-name | text | Planning.Appraisal Name (actual) | Main appraisal name |
| TimeZone | select[name="TimeZone"] | select | Planning.Time Zone | Time zone dropdown |
| PartnerOrganizationId | select[name="PartnerOrganizationId"] | select | *NEW FIELD NEEDED* | Partner organization |
| Objectives | #appraisal-objectives | textarea | Planning.Business Objective Description | Combined BO + AO |
| TeamMemberSignaturesRequired | #signatures-required-yes/no | radio | *NEW FIELD NEEDED* | ATM signatures |
| delivered-virtually | #delivered-virtually-yes/no | radio | Planning.Virtual | Virtual delivery flag |

### Virtual Activity Checkboxes (Organization Scope page)
| Checkbox ID | Label | Purpose |
|-------------|-------|---------|
| 5 | 1: Plan and Prepare for Appraisal | Phase 1 virtual |
| 6 | 1: Plan and Prepare for Appraisal (sub) | Phase 1 sub-activity |
| 9 | 2.1.2 Collect and Examine Affirmations | Phase 2 activity |
| 10 | 2.2.1 Characterize Model Practices | Phase 2 activity |
| 11 | 2.2.2 Generate Preliminary Findings | Phase 2 activity |
| 12 | 2.2.3 Validate Preliminary Results | Phase 2 activity |
| 13 | 2.3.1 Derive Final Findings | Phase 2 activity |
| 14 | 2.3.2 Determine Practice Group Ratings | Phase 2 activity |
| 15 | 2.3.3 Determine Practice Area and Maturity Level | Phase 2 activity |
| 16 | 2.3.4 Record Appraisal Results | Phase 2 activity |
| 17 | Other Activities | Other virtual |
| 7 | 3: Report Results | Phase 3 virtual |

### 2. Appraisal Personnel (/appraisals/{id}/training-schedule)
| CAS Field | Selector | Type | Notes |
|-----------|----------|------|-------|
| ImportFile | #file-input-ImportFile | file | ATM training import |

### 3. OE Collection Approach (/appraisals/{id}/objective-evidence/collection-approach)
| CAS Field | Selector | Type | _FieldMap Field | Notes |
|-----------|----------|------|-----------------|-------|
| Type | select[name="Type"] | select | *NEW FIELD NEEDED* | Discovery/Managed Discovery/Verification |
| Comment | #Comment | textarea | *NEW FIELD NEEDED* | OE approach description |

## Current _FieldMap to CAS Mappings

| Sheet | FieldName | CAS Page | CAS Selector | CAS Field |
|-------|-----------|----------|--------------|-----------|
| Planning | Time Zone | /name-and-type | select[name="TimeZone"] | TimeZone |
| Planning | Virtual | /name-and-type | #delivered-virtually-yes | delivered-virtually |
| Planning | Appraisal Name (actual) | /name-and-type | #appraisal-name | Name |
| Planning | Business Objective Title | /name-and-type | #appraisal-objectives | Objectives |
| Planning | Business Objective Description | /name-and-type | #appraisal-objectives | Objectives |

## Pages Requiring Additional Scraping
The scraper discovered these pages need to be scraped for full mapping:

### OE Collection Plan
1. `/objective-evidence/collection-approach` - OE Collection Approach ✓ (partial)
2. `/objective-evidence/collection-techniques` - OE Collection Techniques
3. `/objective-evidence/collection-responsibilities` - Responsibility for Collection
4. `/objective-evidence/performance-report-approaches` - Performance Report Collection Approach
5. `/objective-evidence/initial-summary` - Summary of Initial OE
6. `/objective-evidence/data-collection-timing` - Data Collection Timing
7. `/objective-evidence/additional-info` - Additional Information

### Logistics and Constraints
8. `/resource-estimates` - Resource Effort Estimate
9. `/logistical-requirements` - Logistics Requirements
10. `/appraisal-constraints` - Appraisal Constraints
11. `/risk-identification` - Risk Identification and Management
12. `/conflicts-of-interest` - COI Identification and Management
13. `/follow-on-activities` - Optional Follow-on Activities

### Appraisal Documents
14. `/appraisal-plan-summary` - Appraisal Plan Summary and Signature
15. `/supporting-documents` - Supporting Document Upload

### Appraisal Outputs
16. `/required-outputs` - Required Outputs
17. `/performance-report` - Performance Report Output
18. `/optional-outputs` - Optional Outputs

### Organization Info
19. `/organizations` - Organizational Info (OU details, addresses)
20. `/timeline` - Appraisal Timeline (key dates)
21. `/readiness-reviews` - Readiness Reviews
22. `/confidentiality-agreement` - Confidentiality and Non-Attribution

## Recommended _FieldMap Schema Updates

Add these columns to _FieldMap for CAS automation:

| Column | Purpose | Example |
|--------|---------|---------|
| CAS_Page | URL path relative to /appraisals/{id} | /name-and-type |
| CAS_Selector | CSS selector for Puppeteer | #appraisal-name |
| CAS_FieldName | Field name attribute in HTML | Name |
| CAS_Type | Field type | text, select, radio, checkbox, textarea |

## Next Steps

1. **Extend Scraper**: Update scraper to visit all 22+ subpages listed above
2. **Update _FieldMap**: Add CAS columns to Excel workbook
3. **Create Sync Tool**: Build bidirectional CAS↔Excel sync automation
4. **Test Automation**: Validate field mappings with test appraisal

## Technical Notes

- Appraisal ID: 81334 (test/dummy appraisal)
- Base URL: https://cas.cmmiinstitute.com
- Form submission uses __RequestVerificationToken (CSRF protection)
- Radio buttons use true/false string values
- Checkboxes use "on" value when checked
