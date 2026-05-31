/**
 * CAS Form Populator - Interactive Mode
 *
 * Populates CAS forms from Excel data using _xlsCasMap mappings.
 * Processes one page at a time, waits for user feedback before continuing.
 *
 * Usage:
 *   1. Configure C:\WorkDir-Claude\cas-project-config.json and keys.json
 *   2. npm run populate
 *
 * Module structure:
 *   config.ts                              - CONFIG object (project + keys loading)
 *   types.ts                               - Shared interfaces
 *   helpers/fieldPopulators.ts             - populateXxx methods (mixin)
 *   handlers/orgScope.handlers.ts          - Phase 1 Org Scope page handlers (mixin)
 *   handlers/objectiveEvidence.handlers.ts - OE page handlers (mixin)
 *
 * To add a new page handler:
 *   1. Write applyMyHandlers(cls) in a new handlers/xxx.handlers.ts file
 *   2. Import and call it below the existing applyXxx calls
 *   3. Add the method signature to the CASPopulator interface at the bottom
 */

import puppeteer, { Browser, Page } from 'puppeteer';
import * as fs from 'fs';
import * as path from 'path';
import * as readline from 'readline';
import * as ExcelJS from 'exceljs';

import { CONFIG } from './config';
import { FieldMapping, ExcelData, PopulateLog } from './types';
import { applyFieldPopulators } from './helpers/fieldPopulators';
import { applyOrgScopeHandlers } from './handlers/orgScope.handlers';
import { applyOEHandlers } from './handlers/objectiveEvidence.handlers';

class CASPopulator {
  browser: Browser | null = null;
  page: Page | null = null;
  fieldMap: FieldMapping[] = [];
  excelData: ExcelData = {};
  private logs: PopulateLog[] = [];
  rl: readline.Interface;

  constructor() {
    this.rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
  }

  // ── Initialisation ───────────────────────────────────────────────────────

  async init(): Promise<void> {
    console.log('🚀 CAS Form Populator - Interactive Mode');
    console.log('='.repeat(60));

    if (!CONFIG.email || !CONFIG.password) {
      console.error('❌ Credentials not set!');
      console.log('Please configure credentials in C:\\WorkDir-Claude\\keys.json:');
      console.log('  { "cas": { "email": "your@email.com", "password": "yourpassword", "staySignedIn": "yes" } }');
      process.exit(1);
    }

    console.log(`🔑 Using credentials from keys.json (email: ${CONFIG.email.substring(0, 3)}...)`);

    if (CONFIG.continueFromPage) {
      console.log(`⏩ Will continue from page: ${CONFIG.continueFromPage}`);
    } else {
      console.log(`▶️  Starting from the beginning (continueFromPage not set)`);
    }

    await this.loadFieldMap();
    await this.loadExcelData();

    const userDataDir = path.join(process.cwd(), '.browser-data');

    this.browser = await puppeteer.launch({
      headless: false,
      defaultViewport: null,
      userDataDir: userDataDir,
      args: ['--start-maximized', '--window-size=1400,900']
    });

    console.log(`📁 Browser data directory: ${userDataDir}`);

    if (!fs.existsSync(CONFIG.htmlLogDir)) {
      fs.mkdirSync(CONFIG.htmlLogDir, { recursive: true });
    }
    console.log(`📄 HTML logs directory: ${CONFIG.htmlLogDir}`);

    this.page = await this.browser.newPage();
    this.page.setDefaultNavigationTimeout(CONFIG.navigationTimeout);

    console.log('✅ Browser launched (visible mode)');
  }

  async loadFieldMap(): Promise<void> {
    const scraperDir = path.dirname(path.resolve(__filename));
    const masterPath = path.join(scraperDir, '_xlsCasMap_MASTER.xlsx');

    let sourceFile: string;
    let sheetName: string;

    if (fs.existsSync(masterPath)) {
      sourceFile = masterPath;
      sheetName  = '_xlsCasMap';
      console.log(`📋 Loading _xlsCasMap from MASTER: ${path.basename(masterPath)}`);
    } else {
      if (!CONFIG.excelFile || !fs.existsSync(CONFIG.excelFile)) {
        console.error('❌ Neither _xlsCasMap_MASTER.xlsx nor target Excel file found!');
        console.log(`   MASTER expected at: ${masterPath}`);
        console.log(`   Target expected at: ${CONFIG.excelFile}`);
        process.exit(1);
      }
      sourceFile = CONFIG.excelFile;
      sheetName  = '_xlsCasMap';
      console.log(`📋 Loading _xlsCasMap from Excel: ${path.basename(CONFIG.excelFile)} (MASTER not found, using fallback)`);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(sourceFile);

    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      console.error(`❌ Sheet '${sheetName}' not found in ${path.basename(sourceFile)}!`);
      process.exit(1);
    }

    sheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
      if (rowNumber === 1) return;

      const rowVal = row.getCell(1).value;
      if (!rowVal) return;

      const mapping: FieldMapping = {
        Row:           typeof rowVal === 'number' ? rowVal : parseInt(String(rowVal), 10),
        Sheet:         String(row.getCell(2).value || ''),
        FieldLabel:    String(row.getCell(3).value || ''),
        CAS_Page:      String(row.getCell(4).value || ''),
        CAS_Selector:  String(row.getCell(5).value || ''),
        CAS_FieldName: String(row.getCell(6).value || ''),
        CAS_Type:      String(row.getCell(7).value || ''),
        Notes:         String(row.getCell(8).value || ''),
      };

      this.fieldMap.push(mapping);
    });

    console.log(`✅ Loaded ${this.fieldMap.length} field mappings from _xlsCasMap`);
  }

  // ── Formula / Excel value resolution ────────────────────────────────────

  resolveCellValue(workbook: ExcelJS.Workbook, cellValue: any, depth: number = 0): string {
    if (depth > 10) return '';
    if (cellValue === null || cellValue === undefined) return '';
    if (typeof cellValue === 'string')  return cellValue;
    if (typeof cellValue === 'number')  return String(cellValue);
    if (typeof cellValue === 'boolean') return cellValue ? 'Yes' : 'No';

    if (cellValue instanceof Date) {
      if (!isNaN(cellValue.getTime())) {
        return cellValue.toISOString().split('T')[0];
      }
      return '';
    }

    if (typeof cellValue === 'object') {
      const cellObj = cellValue as any;

      if ('richText' in cellObj) {
        return cellObj.richText.map((rt: any) => rt.text).join('');
      }
      if ('text' in cellObj && 'hyperlink' in cellObj) {
        return String(cellObj.text);
      }
      if ('formula' in cellObj) {
        const formula      = cellObj.formula as string;
        const cachedResult = 'result' in cellObj ? cellObj.result : undefined;

        const isSimpleRef = /^'?[^'!]+'?![A-Z]+\d+$/.test(formula);
        if (isSimpleRef) {
          const resolved = this.resolveFormulaReference(workbook, formula, depth + 1);
          if (resolved) return resolved;
        }

        const isInvalidDate = cachedResult instanceof Date && isNaN((cachedResult as Date).getTime());
        if (cachedResult !== null && cachedResult !== undefined && cachedResult !== '' && !isInvalidDate) {
          return this.resolveCellValue(workbook, cachedResult, depth + 1);
        }

        const resolved = this.resolveFormulaReference(workbook, formula, depth + 1);
        if (resolved) return resolved;

        return '';
      }
      if ('result' in cellObj) {
        return this.resolveCellValue(workbook, cellObj.result, depth + 1);
      }
    }

    return String(cellValue);
  }

  private resolveFormulaReference(workbook: ExcelJS.Workbook, formula: string, depth: number = 0): string {
    if (depth > 10) return '';

    const simpleRef = formula.match(/^(?:'([^']+)'|([A-Za-z0-9_]+))!([A-Z]+)(\d+)$/);
    if (simpleRef) {
      return this.getCellValueFromSheet(workbook, simpleRef[1] || simpleRef[2], simpleRef[3], parseInt(simpleRef[4], 10), depth);
    }

    const complexRef = formula.match(/(?:'([^']+)'|([A-Za-z0-9_]+))!([A-Z]+)(\d+)/);
    if (complexRef) {
      return this.getCellValueFromSheet(workbook, complexRef[1] || complexRef[2], complexRef[3], parseInt(complexRef[4], 10), depth);
    }

    return '';
  }

  private getCellValueFromSheet(workbook: ExcelJS.Workbook, sheetName: string, col: string, row: number, depth: number): string {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) return '';
    return this.resolveCellValue(workbook, sheet.getCell(`${col}${row}`).value, depth);
  }

  async loadExcelData(): Promise<void> {
    console.log(`📊 Loading data from Excel sheets...`);

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(CONFIG.excelFile);

    const sheetNames = [...new Set(this.fieldMap.map(f => f.Sheet))];

    for (const sheetName of sheetNames) {
      const sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        console.log(`   ⚠️  Sheet ${sheetName} not found`);
        continue;
      }

      this.excelData[sheetName] = {};

      let rowsNeeded = this.fieldMap
        .filter(f => f.Sheet === sheetName)
        .map(f => f.Row);

      if (sheetName === 'P1-OrgScope') {
        const extraRows = [
          69, 70,
          73, 74, 75,
          77, 78, 79,
          85, 87, 88,
          91, 92,
          101, 102, 103, 104, 105, 106,
          112, 113, 114, 115, 116, 117
        ];
        rowsNeeded = [...new Set([...rowsNeeded, ...extraRows])];
      }

      if (sheetName === 'P1PA-R') {
        const extraRows = [
          50, 51,
          54, 55, 56, 57,
          62, 63, 64,
          67, 68, 69,
          72,
          75, 76, 77,
          79, 80, 81,
          86
        ];
        rowsNeeded = [...new Set([...rowsNeeded, ...extraRows])];
      }

      for (const rowNum of rowsNeeded) {
        const value = this.resolveCellValue(workbook, sheet.getRow(rowNum).getCell(2).value);
        if (value && value.trim()) {
          this.excelData[sheetName][rowNum] = value;
        }
      }

      console.log(`   ✅ ${sheetName}: loaded ${Object.keys(this.excelData[sheetName]).length} values`);
    }

    console.log(`✅ Excel data loaded`);
  }

  // ── Session / Login ──────────────────────────────────────────────────────

  async tryDirectNavigation(): Promise<boolean> {
    if (!this.page) return false;

    console.log('\n🔄 Checking if session is still valid...');

    try {
      const firstPage = this.fieldMap[0]?.CAS_Page || '/name-and-type';
      const url = `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}${firstPage}`;

      console.log(`   Navigating to: ${url}`);
      await this.page.goto(url, { waitUntil: 'networkidle2' });
      await new Promise(resolve => setTimeout(resolve, 2000));

      const currentUrl = this.page.url();

      if (currentUrl.includes('login') || currentUrl.includes('Login')) {
        console.log('   ⚠️  Redirected to login page - session expired');
        return false;
      }

      if (currentUrl.includes('cas.cmmiinstitute.com') && currentUrl.includes(CONFIG.appraisalId)) {
        console.log('   ✅ Session is valid - already authenticated!');
        return true;
      }

      console.log(`   ⚠️  Unexpected URL: ${currentUrl}`);
      return false;

    } catch (error) {
      console.log('   ⚠️  Navigation failed, will try login');
      return false;
    }
  }

  async handleStaySignedIn(): Promise<void> {
    if (!this.page) return;

    try {
      const staySignedInSelectors = [
        '#RememberMe',
        'input[name="RememberMe"]',
        'input[type="checkbox"][name*="remember"]',
        'input[type="checkbox"][name*="Remember"]',
        'input[type="checkbox"][id*="remember"]',
        'input[type="checkbox"][id*="Remember"]',
        '.remember-me input[type="checkbox"]'
      ];

      for (const selector of staySignedInSelectors) {
        try {
          const checkbox = await this.page.$(selector);
          if (checkbox) {
            const isChecked = await this.page.evaluate(el => (el as HTMLInputElement).checked, checkbox);

            if (CONFIG.staySignedIn && !isChecked) {
              await checkbox.click();
              console.log('    ✅ Checked "Stay signed in"');
            } else if (!CONFIG.staySignedIn && isChecked) {
              await checkbox.click();
              console.log('    ✅ Unchecked "Stay signed in"');
            } else {
              console.log(`    ℹ️  "Stay signed in" already ${isChecked ? 'checked' : 'unchecked'}`);
            }
            return;
          }
        } catch (e) {
          // Try next selector
        }
      }

      console.log('    ℹ️  No "Stay signed in" checkbox found');
    } catch (error) {
      console.log('    ⚠️  Could not handle "Stay signed in" checkbox:', error);
    }
  }

  async login(): Promise<boolean> {
    if (!this.page) throw new Error('Page not initialized');

    console.log('\n🔐 Logging in to CMMI Institute...');

    try {
      await this.page.goto(CONFIG.loginUrl, { waitUntil: 'networkidle2' });
      await new Promise(resolve => setTimeout(resolve, 3000));

      await this.page.waitForSelector('#UserName', { timeout: 15000 });

      await this.page.click('#UserName');
      await this.page.type('#UserName', CONFIG.email, { delay: 30 });
      console.log('    ✅ Entered username');

      await this.page.click('#Password');
      await this.page.type('#Password', CONFIG.password, { delay: 30 });
      console.log('    ✅ Entered password');

      await this.handleStaySignedIn();

      await Promise.all([
        this.page.waitForNavigation({ waitUntil: 'networkidle2' }),
        this.page.click('input[type="submit"]')
      ]);

      const currentUrl = this.page.url();
      if (!currentUrl.includes('login')) {
        console.log('✅ Login successful');
        return true;
      }
      return false;

    } catch (error) {
      console.error('❌ Login failed:', error);
      return false;
    }
  }

  // ── Navigation helpers ───────────────────────────────────────────────────

  async navigateToPage(pagePath: string): Promise<boolean> {
    if (!this.page) return false;

    const url = `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}${pagePath}`;
    console.log(`\n📄 Navigating to: ${pagePath}`);
    console.log(`   URL: ${url}`);

    try {
      await this.page.goto(url, { waitUntil: 'networkidle2' });
      await new Promise(resolve => setTimeout(resolve, 2000));
      console.log('   ✅ Page loaded');
      return true;
    } catch (error) {
      console.error('   ❌ Navigation failed:', error);
      return false;
    }
  }

  async clickNextButton(): Promise<boolean> {
    if (!this.page) return false;

    console.log('\n   ➡️  Looking for Next button...');

    try {
      const result = await this.page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll('a.button.blue-button'));
        for (const btn of buttons) {
          const polygon = btn.querySelector('svg polygon');
          if (polygon) {
            const points = polygon.getAttribute('points') || '';
            if (points.startsWith('80,60')) {
              const href = btn.getAttribute('href');
              if (href) {
                (btn as HTMLElement).click();
                return { success: true, href };
              }
            }
          }
        }
        return { success: false, href: null };
      });

      if (result.success && result.href) {
        console.log(`      Found Next button: ${result.href}`);
        console.log('      ✅ Next button clicked');
        await new Promise(resolve => setTimeout(resolve, 2000));
        try {
          await this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 10000 });
          console.log('      ✅ Navigated to next page');
        } catch (e) { /* may already be done */ }
        return true;
      }
    } catch (e) {
      console.log(`      ⚠️  Error finding Next button: ${e}`);
    }

    console.log('      ℹ️  No Next button found');
    return false;
  }

  async clickSaveButton(): Promise<boolean> {
    if (!this.page) return false;

    console.log('\n   💾 Looking for save/update button...');

    const saveButtonSelectors = [
      'button[data-test="button-update-appraisal"]',
      'button[data-test="button-add-edit-org"]',
      'button[data-test*="update"]',
      'button[data-test*="save"]',
      'button[data-test*="add"]',
      '.actions button:not(.p2):not(.red-button)',
      'input[type="submit"][value*="Update"]',
      'input[type="submit"][value*="Add"]',
      'button[type="submit"]:not([disabled])',
      'input[type="submit"][value*="Save"]',
      '.btn-primary[type="submit"]',
      'input.btn[type="submit"]'
    ];

    for (const selector of saveButtonSelectors) {
      try {
        const button = await this.page.$(selector);
        if (button) {
          const buttonText = await this.page.evaluate(el => {
            return (el as HTMLInputElement).value || el.textContent || '';
          }, button);
          console.log(`      Found button: "${buttonText.trim()}" (${selector})`);
          await button.click();
          console.log('      ✅ Button clicked');
          await new Promise(resolve => setTimeout(resolve, 2000));
          try {
            await this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 5000 });
            console.log('      ✅ Page navigation completed');
          } catch (navError) {
            console.log('      ℹ️  No page navigation (AJAX update)');
          }
          await new Promise(resolve => setTimeout(resolve, 1000));
          return true;
        }
      } catch (e) { /* try next */ }
    }

    // Fallback: find by text content
    try {
      const clicked = await this.page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll('button, input[type="submit"], input[type="button"]'));
        for (const btn of buttons) {
          const text = (btn.textContent || (btn as HTMLInputElement).value || '').toLowerCase();
          if (text.includes('update') || text.includes('save') || text === 'add target') {
            (btn as HTMLElement).click();
            return (btn.textContent || (btn as HTMLInputElement).value || '').trim();
          }
        }
        return null;
      });
      if (clicked) {
        console.log(`      Found button by text: "${clicked}"`);
        console.log('      ✅ Button clicked');
        await new Promise(resolve => setTimeout(resolve, 3000));
        return true;
      }
    } catch (e) { /* fallback failed */ }

    console.log('      ⚠️  No save/update button found');
    return false;
  }

  async clickAddEditButton(selector: string, buttonName: string): Promise<boolean> {
    if (!this.page) return false;

    console.log(`\n   ➕ Looking for "${buttonName}" button...`);

    try {
      await new Promise(resolve => setTimeout(resolve, 1000));

      if (selector) {
        try {
          await this.page.waitForSelector(selector, { timeout: 5000 });
          const button = await this.page.$(selector);
          if (button) {
            await this.page.evaluate((sel: string) => {
              const el = document.querySelector(sel);
              if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }, selector);
            await new Promise(resolve => setTimeout(resolve, 500));
            await button.click();
            console.log(`      ✅ Clicked "${buttonName}" (${selector})`);
            await new Promise(resolve => setTimeout(resolve, 2000));
            return true;
          }
        } catch (e) {
          console.log(`      ℹ️  Selector ${selector} not found, trying text search...`);
        }
      }

      const clicked = await this.page.evaluate((text: string) => {
        const buttons = Array.from(document.querySelectorAll('a, button'));
        for (const btn of buttons) {
          const btnText = btn.textContent?.trim().toLowerCase() || '';
          if (btnText.includes(text.toLowerCase())) {
            (btn as HTMLElement).scrollIntoView({ behavior: 'smooth', block: 'center' });
            (btn as HTMLElement).click();
            return btn.textContent?.trim() || text;
          }
        }
        return null;
      }, buttonName);

      if (clicked) {
        console.log(`      ✅ Clicked "${clicked}" (by text)`);
        await new Promise(resolve => setTimeout(resolve, 2000));
        return true;
      }
    } catch (e) {
      console.log(`      ⚠️  Error clicking button: ${e}`);
    }

    console.log(`      ⚠️  "${buttonName}" button not found`);
    return false;
  }

  async clickAddButton(buttonText: string): Promise<boolean> {
    return this.clickAddEditButton('', buttonText);
  }

  // ── HTML snapshot ────────────────────────────────────────────────────────

  async savePageHtml(pagePath: string, suffix: string = ''): Promise<string> {
    if (!this.page) return '';

    try {
      const html      = await this.page.content();
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const safePath  = pagePath.replace(/\//g, '_').replace(/^_/, '');
      const filename  = `${timestamp}_${safePath}${suffix ? '_' + suffix : ''}.html`;
      const filepath  = path.join(CONFIG.htmlLogDir, filename);

      fs.writeFileSync(filepath, html, 'utf-8');
      console.log(`   💾 Saved HTML: ${filename}`);

      return filepath;
    } catch (error) {
      console.log(`   ⚠️  Could not save HTML: ${error}`);
      return '';
    }
  }

  // ── Page processing ──────────────────────────────────────────────────────

  async processCurrentPage(pagePath: string): Promise<PopulateLog> {
    const pageFields = this.fieldMap.filter(f =>
      f.CAS_Page === pagePath &&
      f.CAS_Type !== 'skip' &&
      f.CAS_Selector !== 'SKIP' &&
      f.CAS_Selector !== 'HANDLER'
    );

    const log: PopulateLog = {
      timestamp:        new Date().toISOString(),
      page:             pagePath,
      fieldsAttempted:  pageFields.length,
      fieldsSuccessful: 0,
      fieldsFailed:     0,
      details:          []
    };

    console.log('\n' + '-'.repeat(60));
    console.log(`📋 Processing page: ${pagePath}`);
    console.log(`   Fields to populate: ${pageFields.length}`);
    console.log('-'.repeat(60));

    await this.savePageHtml(pagePath, 'before');

    // Pages that need pre-processing before field loop (edit vs add mode)
    if (pagePath === '/organizations')    { await this.handleOrganizationsPage(); }
    if (pagePath === '/org-units')        { await this.handleOrgUnitsPage(); }
    if (pagePath === '/org-unit-targets') { await this.handleOrgUnitTargetsPage(); }

    // Pages that are fully self-contained (bypass field loop)
    if (pagePath === '/timeline')          { await this.handleTimelinePage();         return log; }
    if (pagePath === '/readiness-reviews') { await this.handleReadinessReviewsPage(); return log; }

    if (pagePath.startsWith('/org-projects')) {
      if (pagePath.includes('/true') || pagePath.includes('IsOrganizational=True')) {
        await this.handleOrgProjectsPage();
      } else {
        await this.handleOUProjectsPage();
      }
      return log;
    }

    if (pagePath.includes('/org-unit-sampling-factors') && !pagePath.includes('values')) {
      await this.handleSamplingFactorsPage();      return log;
    }
    if (pagePath.includes('/org-unit-sampling-factor-values')) {
      await this.handleSamplingFactorValuesPage(); return log;
    }
    if (pagePath.includes('/org-unit-subgroups') && !pagePath.includes('assignment')) {
      await this.handleSubgroupsPage();            return log;
    }
    if (pagePath.includes('/org-unit-project-subgroups')) {
      await this.handleSubgroupAssignmentPage();   return log;
    }
    if (pagePath.includes('/organizational-project-appraisal-scope')) {
      await this.handleOrgProjectAppraisalScopePage(); return log;
    }

    // OE pages
    if (pagePath === '/objective-evidence/collection-approach')        { await this.handleOECollectionApproachPage();        return log; }
    if (pagePath === '/objective-evidence/initial-summary')            { await this.handleOEInitialSummaryPage();            return log; }
    if (pagePath === '/objective-evidence/collection-techniques')      { await this.handleOECollectionTechniquesPage();      return log; }
    if (pagePath === '/objective-evidence/collection-responsibilities') { await this.handleOECollectionResponsibilitiesPage(); return log; }
    if (pagePath === '/objective-evidence/performance-report-approaches') { await this.handlePerformanceReportApproachesPage(); return log; }
    if (pagePath === '/objective-evidence/data-collection-timing')     { await this.handleDataCollectionTimingPage();        return log; }
    if (pagePath === '/objective-evidence/additional-info')            { await this.handleOEAdditionalInfoPage();            return log; }

    // ── Standard field-by-field population ──────────────────────────────
    let anyFieldsChanged = false;

    for (const field of pageFields) {
      const value = this.excelData[field.Sheet]?.[field.Row];

      if (!value) {
        console.log(`\n   ⏭️  Skipping ${field.FieldLabel} (no value in Excel data)`);
        log.details.push({ field: field.FieldLabel, selector: field.CAS_Selector, value: '', status: 'skipped' });
        continue;
      }

      const result = await this.populateField(field, value);

      log.details.push({
        field:    field.FieldLabel,
        selector: field.CAS_Selector,
        value:    value,
        status:   result.success ? 'success' : 'failed',
        error:    result.error
      });

      if (result.success) {
        log.fieldsSuccessful++;
        if (result.changed) anyFieldsChanged = true;
      } else {
        log.fieldsFailed++;
      }
    }

    if (anyFieldsChanged) {
      console.log('\n   💾 Changes detected, saving...');
      // /org-unit-targets may show "Add Target" instead of the generic save button
      if (pagePath === '/org-unit-targets') {
        const addTargetClicked = await this.page?.evaluate(() => {
          const buttons = Array.from(document.querySelectorAll('.actions button, button'));
          for (const btn of buttons) {
            const text = (btn.textContent || '').trim().toLowerCase();
            if (text === 'add target' || text.includes('add target')) {
              (btn as HTMLElement).click();
              return btn.textContent?.trim() || 'Add Target';
            }
          }
          return null;
        });
        if (addTargetClicked) {
          console.log(`   ✅ Clicked: "${addTargetClicked}"`);
          await new Promise(resolve => setTimeout(resolve, 3000));
        } else {
          await this.clickSaveButton();
        }
      } else {
        await this.clickSaveButton();
      }
    } else if (pageFields.length > 0) {
      console.log('\n   ℹ️  No changes needed, skipping save');
    }

    await this.savePageHtml(pagePath, 'after');

    return log;
  }

  // Legacy navigation-first method kept for backward compatibility
  async processPage(pagePath: string): Promise<PopulateLog> {
    const navSuccess = await this.navigateToPage(pagePath);
    const log = await this.processCurrentPage(pagePath);
    if (!navSuccess) {
      log.fieldsFailed = log.fieldsAttempted;
    }
    return log;
  }

  // ── User interaction ─────────────────────────────────────────────────────

  async askForFeedback(): Promise<{ continue: boolean; feedback: string }> {
    if (!CONFIG.debugMode) {
      console.log('\n   ▶️  Debug mode OFF - auto-continuing to next page...');
      return { continue: true, feedback: '' };
    }

    return new Promise((resolve) => {
      console.log('\n' + '='.repeat(60));
      console.log('📋 PAGE POPULATION COMPLETE');
      console.log('='.repeat(60));
      console.log('\nPlease review the populated fields in the browser.');
      console.log('\nOptions:');
      console.log('  [Enter]     - Continue to next page');
      console.log('  [c]         - Continue to next page');
      console.log('  [s] or [q]  - Stop processing');
      console.log('  [any text]  - Save as feedback and continue');
      console.log('  [any text][s] or [any text][q] - Save feedback AND stop');
      console.log('');

      this.rl.question('Your input: ', (answer: string) => {
        const input      = answer.trim();
        const inputLower = input.toLowerCase();

        if (inputLower === 's' || inputLower === 'q' || inputLower === 'stop' || inputLower === 'quit') {
          resolve({ continue: false, feedback: '' });
        } else if (input === '' || inputLower === 'c' || inputLower === 'continue') {
          resolve({ continue: true, feedback: '' });
        } else {
          const endsWithStop = /\[s\]\s*$/i.test(input) || /\[q\]\s*$/i.test(input);
          let feedback = input;
          if (endsWithStop) {
            feedback = input.replace(/\s*\[[sq]\]\s*$/i, '').trim();
          }
          console.log(`💬 Feedback saved: "${feedback}"`);
          if (endsWithStop) {
            console.log('🛑 Stop requested after feedback.');
            resolve({ continue: false, feedback });
          } else {
            resolve({ continue: true, feedback });
          }
        }
      });
    });
  }

  async handlePhaseComplete(): Promise<void> {
    console.log('\n' + '═'.repeat(60));
    console.log('🎉 PHASE 1 COMPLETE - Reached ' + CONFIG.finalPage);
    console.log('═'.repeat(60));
    console.log('\n✅ All Phase 1 pages have been processed successfully!');

    if (CONFIG.autoExitOnComplete) {
      console.log('\n👋 Auto-exit enabled. Closing browser and exiting...');
      await this.browser?.close();
      process.exit(0);
    } else {
      return new Promise((resolve) => {
        this.rl.question('\nPress any key to stop... ', () => resolve());
      });
    }
  }

  // ── Main run loop ────────────────────────────────────────────────────────

  async run(): Promise<void> {
    try {
      await this.init();

      const sessionValid = await this.tryDirectNavigation();

      if (!sessionValid) {
        console.log('\n🔐 Session expired or not found, logging in...');
        const loginSuccess = await this.login();
        if (!loginSuccess) {
          console.log('❌ Login failed. Exiting.');
          return;
        }
      }

      let startPage = '/name-and-type';

      if (CONFIG.continueFromPage) {
        startPage = CONFIG.continueFromPage;
        console.log(`\n⏩ Continuing from page: ${startPage}`);
      } else {
        console.log(`\n▶️  Starting from first page: ${startPage}`);
      }

      await this.navigateToPage(startPage);

      let pageCount = 0;
      let continueProcessing = true;

      while (continueProcessing) {
        pageCount++;

        const currentUrl = this.page?.url() || '';
        const urlPath    = new URL(currentUrl).pathname;
        const pagePath   = urlPath.replace(`/appraisals/${CONFIG.appraisalId}`, '');

        console.log(`\n${'═'.repeat(60)}`);
        console.log(`PAGE ${pageCount}: ${pagePath}`);
        console.log('═'.repeat(60));

        if (pagePath === CONFIG.finalPage || pagePath.includes(CONFIG.finalPage)) {
          await this.handlePhaseComplete();
          continueProcessing = false;
          break;
        }

        const shouldSkip = CONFIG.skipPages.some(skip => pagePath.includes(skip));
        if (shouldSkip) {
          console.log(`   ⏭️  Skipping page (in skipPages list)`);
          const hasNextPage = await this.clickNextButton();
          if (!hasNextPage) {
            console.log('\n✅ No more pages (reached end of workflow)');
            continueProcessing = false;
          } else {
            await new Promise(resolve => setTimeout(resolve, 2000));
          }
          continue;
        }

        const log = await this.processCurrentPage(pagePath);
        this.logs.push(log);

        console.log(`\n📊 Page Summary:`);
        console.log(`   ✅ Successful: ${log.fieldsSuccessful}`);
        console.log(`   ❌ Failed:     ${log.fieldsFailed}`);
        console.log(`   ⏭️  Skipped:   ${log.details.filter(d => d.status === 'skipped').length}`);

        const { continue: shouldContinue, feedback } = await this.askForFeedback();
        if (feedback) log.userFeedback = feedback;

        this.saveLogs();

        if (!shouldContinue) {
          console.log('\n🛑 Stopping as requested.');
          break;
        }

        const hasNextPage = await this.clickNextButton();

        if (!hasNextPage) {
          console.log('\n✅ No more pages (reached end of workflow)');
          continueProcessing = false;
        } else {
          await new Promise(resolve => setTimeout(resolve, 2000));
        }
      }

      console.log('\n' + '═'.repeat(60));
      console.log('✅ PROCESSING COMPLETE');
      console.log('═'.repeat(60));
      console.log(`📝 Log saved to: ${CONFIG.logFile}`);

    } catch (error) {
      console.error('❌ Error:', error);
    } finally {
      this.saveLogs();
      this.rl.close();
      console.log('\n🌐 Browser left open for review. Close manually when done.');
    }
  }

  saveLogs(): void {
    fs.writeFileSync(CONFIG.logFile, JSON.stringify(this.logs, null, 2));
  }
}

// ── Apply mixins ─────────────────────────────────────────────────────────────
applyFieldPopulators(CASPopulator);
applyOrgScopeHandlers(CASPopulator);
applyOEHandlers(CASPopulator);

// ── TypeScript interface merging (gives full type checking for mixin methods) ─
// When you add a new handler module, also add its method signatures here.
interface CASPopulator {
  // Field populators (from helpers/fieldPopulators.ts)
  populateField(mapping: FieldMapping, value: string): Promise<{ success: boolean; changed: boolean; error?: string }>;
  populateTextInput(selector: string, value: string): Promise<{ changed: boolean }>;
  populateSelect(selector: string, value: string): Promise<{ changed: boolean }>;
  populateRadio(selector: string, value: string, notes: string): Promise<{ changed: boolean }>;
  populateCheckbox(selector: string, value: string): Promise<{ changed: boolean }>;
  populateNumberInput(selector: string, value: string): Promise<{ changed: boolean }>;
  populateDateParts(selector: string, value: string): Promise<{ changed: boolean }>;
  populateDateInput(selector: string, value: string): Promise<{ changed: boolean }>;
  populateRadioLevel(selector: string, value: string, notes: string): Promise<{ changed: boolean }>;
  populateMultiselect(selector: string, value: string): Promise<{ changed: boolean }>;

  // Org Scope handlers (from handlers/orgScope.handlers.ts)
  handleOrganizationsPage(): Promise<void>;
  handleOrgUnitsPage(): Promise<void>;
  handleOrgUnitTargetsPage(): Promise<void>;
  handleTimelinePage(): Promise<void>;
  handleReadinessReviewsPage(): Promise<void>;
  handleOrgProjectsPage(): Promise<void>;
  handleOUProjectsPage(): Promise<void>;
  handleSamplingFactorsPage(): Promise<void>;
  handleSamplingFactorValuesPage(): Promise<void>;
  handleSubgroupsPage(): Promise<void>;
  handleSubgroupAssignmentPage(): Promise<void>;
  handleOrgProjectAppraisalScopePage(): Promise<void>;

  // OE handlers (from handlers/objectiveEvidence.handlers.ts)
  fillAndSaveSimpleForm(fields: Array<{ selector: string; type: 'select' | 'text' | 'textarea' | 'date'; value: string }>): Promise<void>;
  handleOECollectionApproachPage(): Promise<void>;
  handleOECollectionTechniquesPage(): Promise<void>;
  handleOECollectionResponsibilitiesPage(): Promise<void>;
  handlePerformanceReportApproachesPage(): Promise<void>;
  handleInitialSummaryPage(): Promise<void>;
  handleDataCollectionTimingPage(): Promise<void>;
  handleOEInitialSummaryPage(): Promise<void>;
  handleOEAdditionalInfoPage(): Promise<void>;
}

// ── Main entry point ──────────────────────────────────────────────────────────
const populator = new CASPopulator();
populator.run();
