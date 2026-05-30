/**
 * CAS Form Populator - Interactive Mode
 * 
 * Populates CAS forms from Excel data using _xlsCasMap mappings.
 * Processes one page at a time, waits for user feedback before continuing.
 * 
 * Usage:
 *   1. Set environment variables: CAS_EMAIL, CAS_PASSWORD
 *   2. npm run populate
 */

import puppeteer, { Browser, Page } from 'puppeteer';
import * as fs from 'fs';
import * as path from 'path';
import * as readline from 'readline';
import * as ExcelJS from 'exceljs';
import AdmZip from 'adm-zip';

// Configuration
// Load project config if available
// Use absolute path to ensure config is found regardless of working directory
const projectConfigPath = 'C:/WorkDir-Claude/cas-project-config.json';
let projectConfig: any = null;
let keysConfig: any = null;

try {
  if (fs.existsSync(projectConfigPath)) {
    projectConfig = JSON.parse(fs.readFileSync(projectConfigPath, 'utf-8'));
    
    // Load keys from separate file (best practice: keep secrets separate)
    if (projectConfig?.keysFile && fs.existsSync(projectConfig.keysFile)) {
      keysConfig = JSON.parse(fs.readFileSync(projectConfig.keysFile, 'utf-8'));
    }
  }
} catch (e) {
  // Config optional, continue with defaults
}

const CONFIG = {
  casBaseUrl: projectConfig?.cas?.baseUrl || 'https://cas.cmmiinstitute.com',
  loginUrl: projectConfig?.cas?.loginUrl || 'https://cmmiinstitute.com/login',
  appraisalId: projectConfig?.project?.casId || '81846',
  
  // Continue from specific page (empty string = start from beginning)
  continueFromPage: projectConfig?.cas?.continueFromPage || '',
  
  // Credentials from keys.json (preferred) or environment variables (fallback)
  email: keysConfig?.cas?.email || process.env.CAS_EMAIL || '',
  password: keysConfig?.cas?.password || process.env.CAS_PASSWORD || '',
  staySignedIn: keysConfig?.cas?.staySignedIn?.toLowerCase() === 'yes',
  
  // Excel source file (from cas-project-config.json files.target)
  excelFile: projectConfig?.files?.target || '',
  
  // Timing
  navigationTimeout: 90000,
  waitAfterAction: 1000,
  
  // Logging
  logFile: 'populate_log.json',
  htmlLogDir: 'html_logs',
  
  // Debug mode - if true, prompt for input after each page; if false, auto-continue
  debugMode: projectConfig?.cas?.debugMode ?? true,
  
  // Auto exit on complete - if true, exit when reaching /sample-scope; if false, prompt
  autoExitOnComplete: projectConfig?.cas?.autoExitOnComplete ?? false,
  
  // Pages to skip (no processing needed)
  skipPages: [
    '/org-unit-project-appraisal-scope',
    '/include-project'
  ],
  
  // Final page - reaching this means Phase 1 is complete
  finalPage: '/sample-scope'
};

interface FieldMapping {
  Row: number;
  Sheet: string;
  FieldLabel: string;
  CAS_Page: string;
  CAS_Selector: string;
  CAS_FieldName: string;
  CAS_Type: string;
  Notes: string;
}

interface ExcelData {
  [sheet: string]: {
    [row: number]: string;
  };
}

interface PopulateLog {
  timestamp: string;
  page: string;
  fieldsAttempted: number;
  fieldsSuccessful: number;
  fieldsFailed: number;
  details: {
    field: string;
    selector: string;
    value: string;
    status: 'success' | 'failed' | 'skipped';
    error?: string;
  }[];
  userFeedback?: string;
}

class CASPopulator {
  private browser: Browser | null = null;
  private page: Page | null = null;
  private fieldMap: FieldMapping[] = [];
  private excelData: ExcelData = {};
  private logs: PopulateLog[] = [];
  private rl: readline.Interface;

  constructor() {
    this.rl = readline.createInterface({
      input: process.stdin,
      output: process.stdout
    });
  }

  async init(): Promise<void> {
    console.log('🚀 CAS Form Populator - Interactive Mode');
    console.log('=' .repeat(60));
    
    if (!CONFIG.email || !CONFIG.password) {
      console.error('❌ Credentials not set!');
      console.log('Please configure credentials in C:\\WorkDir-Claude\\keys.json:');
      console.log('  {');
      console.log('    "cas": {');
      console.log('      "email": "your@email.com",');
      console.log('      "password": "yourpassword",');
      console.log('      "staySignedIn": "yes"');
      console.log('    }');
      console.log('  }');
      process.exit(1);
    }
    
    console.log(`🔑 Using credentials from keys.json (email: ${CONFIG.email.substring(0, 3)}...)`);
    
    // Show continueFromPage setting
    if (CONFIG.continueFromPage) {
      console.log(`⏩ Will continue from page: ${CONFIG.continueFromPage}`);
    } else {
      console.log(`▶️  Starting from the beginning (continueFromPage not set)`);
    }

    // Load field mappings
    await this.loadFieldMap();
    
    // Load Excel data
    await this.loadExcelData();

    // Launch browser (visible with scrollbars)
    // Use persistent user data directory to save cookies/session
    const userDataDir = path.join(process.cwd(), '.browser-data');
    
    this.browser = await puppeteer.launch({
      headless: false,
      defaultViewport: null,  // Use actual window size, enables scrollbars
      userDataDir: userDataDir,  // Persist cookies and session data
      args: [
        '--start-maximized',
        '--window-size=1400,900'
      ]
    });
    
    console.log(`📁 Browser data directory: ${userDataDir}`);
    
    // Create HTML logs directory
    if (!fs.existsSync(CONFIG.htmlLogDir)) {
      fs.mkdirSync(CONFIG.htmlLogDir, { recursive: true });
    }
    console.log(`📄 HTML logs directory: ${CONFIG.htmlLogDir}`);

    this.page = await this.browser.newPage();
    this.page.setDefaultNavigationTimeout(CONFIG.navigationTimeout);
    
    console.log('✅ Browser launched (visible mode)');
  }

  async loadFieldMap(): Promise<void> {
    // Prefer _xlsCasMap_MASTER.xlsx in the scraper folder - this is the authoritative
    // source for selectors and survives Excel version bumps.
    // Falls back to _xlsCasMap sheet inside files.target if MASTER not found.
    const scraperDir = path.dirname(path.resolve(__filename));
    const masterPath = path.join(scraperDir, '_xlsCasMap_MASTER.xlsx');
    
    let sourceFile: string;
    let sheetName: string;
    
    if (fs.existsSync(masterPath)) {
      sourceFile = masterPath;
      sheetName = '_xlsCasMap';
      console.log(`📋 Loading _xlsCasMap from MASTER: ${path.basename(masterPath)}`);
    } else {
      if (!CONFIG.excelFile || !fs.existsSync(CONFIG.excelFile)) {
        console.error('❌ Neither _xlsCasMap_MASTER.xlsx nor target Excel file found!');
        console.log(`   MASTER expected at: ${masterPath}`);
        console.log(`   Target expected at: ${CONFIG.excelFile}`);
        process.exit(1);
      }
      sourceFile = CONFIG.excelFile;
      sheetName = '_xlsCasMap';
      console.log(`📋 Loading _xlsCasMap from Excel: ${path.basename(CONFIG.excelFile)} (MASTER not found, using fallback)`);
    }
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(sourceFile);
    
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) {
      console.error(`❌ Sheet '${sheetName}' not found in ${path.basename(sourceFile)}!`);
      process.exit(1);
    }
    
    // Read field mappings (skip header row)
    sheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
      if (rowNumber === 1) return; // Skip header
      
      const rowVal = row.getCell(1).value;
      if (!rowVal) return;
      
      const mapping: FieldMapping = {
        Row: typeof rowVal === 'number' ? rowVal : parseInt(String(rowVal), 10),
        Sheet: String(row.getCell(2).value || ''),
        FieldLabel: String(row.getCell(3).value || ''),
        CAS_Page: String(row.getCell(4).value || ''),
        CAS_Selector: String(row.getCell(5).value || ''),
        CAS_FieldName: String(row.getCell(6).value || ''),
        CAS_Type: String(row.getCell(7).value || ''),
        Notes: String(row.getCell(8).value || ''),
      };
      
      this.fieldMap.push(mapping);
    });
    
    console.log(`✅ Loaded ${this.fieldMap.length} field mappings from _xlsCasMap`);
  }

  // Helper to resolve formula references and get the actual cell value
  // This follows formula references like "'Agreement 3'!B16" to get the source value
  private resolveCellValue(workbook: ExcelJS.Workbook, cellValue: any, depth: number = 0): string {
    if (depth > 10) return ''; // Prevent infinite recursion
    
    if (cellValue === null || cellValue === undefined) return '';
    
    // Direct string or number
    if (typeof cellValue === 'string') return cellValue;
    if (typeof cellValue === 'number') return String(cellValue);
    if (typeof cellValue === 'boolean') return cellValue ? 'Yes' : 'No';
    
    // Date object (including Invalid Date from ExcelJS formula cache misses)
    if (cellValue instanceof Date) {
      if (!isNaN(cellValue.getTime())) {
        return cellValue.toISOString().split('T')[0];
      }
      // Invalid Date - ExcelJS sometimes stores these as cached results
      // for formulas it couldn't evaluate. Treat as empty so the formula
      // reference resolver can follow the chain instead.
      return '';
    }
    
    // Object (formula, richText, etc.)
    if (typeof cellValue === 'object') {
      const cellObj = cellValue as any;
      
      // Rich text
      if ('richText' in cellObj) {
        return cellObj.richText.map((rt: any) => rt.text).join('');
      }
      
      // Hyperlink
      if ('text' in cellObj && 'hyperlink' in cellObj) {
        return String(cellObj.text);
      }
      
      // Formula cell
      if ('formula' in cellObj) {
        const formula = cellObj.formula as string;
        const cachedResult = 'result' in cellObj ? cellObj.result : undefined;
        
        // Simple direct cell reference (e.g. "Sheet!B4" or "'Sheet Name'!B4")
        // These are safe to follow directly - more reliable than cached result in .xlsm
        const isSimpleRef = /^'?[^'!]+'?![A-Z]+\d+$/.test(formula);
        
        if (isSimpleRef) {
          // Follow the reference directly for simple cross-sheet refs
          const resolvedValue = this.resolveFormulaReference(workbook, formula, depth + 1);
          if (resolvedValue) return resolvedValue;
        }
        
        // For complex formulas (IF, concatenation, etc.) use the cached result
        // Guard against Invalid Date objects: ExcelJS may cache these for
        // formulas it couldn't evaluate (JSON.stringify shows null but !== null is true)
        const isInvalidDate = cachedResult instanceof Date && isNaN((cachedResult as Date).getTime());
        if (cachedResult !== null && cachedResult !== undefined && cachedResult !== '' && !isInvalidDate) {
          return this.resolveCellValue(workbook, cachedResult, depth + 1);
        }
        
        // Last resort: try formula reference even for complex formulas
        // (extracts first cell reference found inside the formula)
        const resolvedValue = this.resolveFormulaReference(workbook, formula, depth + 1);
        if (resolvedValue) return resolvedValue;
        
        return '';
      }
      
      // Result only (shouldn't happen but handle it)
      if ('result' in cellObj) {
        return this.resolveCellValue(workbook, cellObj.result, depth + 1);
      }
    }
    
    // Fallback
    return String(cellValue);
  }
  
  // Resolve a formula reference like "'Agreement 3'!B16" or "StartupInfo!B4"
  private resolveFormulaReference(workbook: ExcelJS.Workbook, formula: string, depth: number = 0): string {
    if (depth > 10) return ''; // Prevent infinite recursion
    
    // Try to match simple cell reference patterns:
    // 'Sheet Name'!B16  or  SheetName!B4
    const simpleRefMatch = formula.match(/^(?:'([^']+)'|([A-Za-z0-9_]+))!([A-Z]+)(\d+)$/);
    if (simpleRefMatch) {
      const sheetName = simpleRefMatch[1] || simpleRefMatch[2];
      const col = simpleRefMatch[3];
      const row = parseInt(simpleRefMatch[4], 10);
      return this.getCellValueFromSheet(workbook, sheetName, col, row, depth);
    }
    
    // Try to extract a reference from more complex formulas
    // e.g., IF('Agreement 3'!B17<>"", 'Agreement 3'!B17, "")
    const complexRefMatch = formula.match(/(?:'([^']+)'|([A-Za-z0-9_]+))!([A-Z]+)(\d+)/);
    if (complexRefMatch) {
      const sheetName = complexRefMatch[1] || complexRefMatch[2];
      const col = complexRefMatch[3];
      const row = parseInt(complexRefMatch[4], 10);
      return this.getCellValueFromSheet(workbook, sheetName, col, row, depth);
    }
    
    return '';
  }
  
  // Get cell value from a specific sheet and cell address
  private getCellValueFromSheet(workbook: ExcelJS.Workbook, sheetName: string, col: string, row: number, depth: number): string {
    const sheet = workbook.getWorksheet(sheetName);
    if (!sheet) return '';
    
    const cell = sheet.getCell(`${col}${row}`);
    return this.resolveCellValue(workbook, cell.value, depth);
  }

  async loadExcelData(): Promise<void> {
    console.log(`📊 Loading data from Excel sheets...`);
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(CONFIG.excelFile);
    
    // Get unique sheet names from field mappings
    const sheetNames = [...new Set(this.fieldMap.map(f => f.Sheet))];
    
    for (const sheetName of sheetNames) {
      const sheet = workbook.getWorksheet(sheetName);
      if (!sheet) {
        console.log(`   ⚠️  Sheet ${sheetName} not found`);
        continue;
      }
      
      this.excelData[sheetName] = {};
      
      // Get the rows we need for this sheet
      let rowsNeeded = this.fieldMap
        .filter(f => f.Sheet === sheetName)
        .map(f => f.Row);
      
      // Add timeline and readiness review rows that are handled specially
      // These may have DYNAMIC selectors in _xlsCasMap but we still need the data
      if (sheetName === 'P1-OrgScope') {
        const extraRows = [
          69, 70,           // Phase 1 dates
          73, 74, 75,       // Readiness Review 1 (name, start, end)
          77, 78, 79,       // Readiness Review 2 (name, start, end)
          85, 87, 88,       // Phase 2 dates + days on site
          91, 92,           // Phase 3 dates
          101, 102, 103, 104, 105, 106,  // Readiness Review 1 additional fields
          112, 113, 114, 115, 116, 117   // Readiness Review 2 additional fields
        ];
        rowsNeeded = [...new Set([...rowsNeeded, ...extraRows])];
      }

      if (sheetName === 'P1PA-R') {
        // Ensure date rows are included even when they have HANDLER entries in the fieldmap
        const extraRows = [
          50, 51,           // OE Collection Approach + Description
          54, 55, 56, 57,   // Collection Techniques
          62, 63, 64,       // Collection Responsibilities
          67, 68, 69,       // Performance Report Approach
          72,               // Summary of Initial OE
          75, 76, 77,       // Data Collection Timing milestone 1
          79, 80, 81,       // Data Collection Timing milestone 2
          86                // Additional Information
        ];
        rowsNeeded = [...new Set([...rowsNeeded, ...extraRows])];
      }
      
      for (const rowNum of rowsNeeded) {
        const row = sheet.getRow(rowNum);
        const cellValue = row.getCell(2).value; // Column B contains the data
        
        // Use the new resolver that follows formula references
        const value = this.resolveCellValue(workbook, cellValue);
        
        if (value && value.trim()) {
          this.excelData[sheetName][rowNum] = value;
        }
      }
      
      const loadedCount = Object.keys(this.excelData[sheetName]).length;
      console.log(`   ✅ ${sheetName}: loaded ${loadedCount} values`);
    }
    
    console.log(`✅ Excel data loaded`);
  }

  // Try to navigate directly to CAS - returns true if session is still valid
  async tryDirectNavigation(): Promise<boolean> {
    if (!this.page) return false;
    
    console.log('\n🔄 Checking if session is still valid...');
    
    try {
      // Try to go directly to the first CAS page
      const firstPage = this.fieldMap[0]?.CAS_Page || '/name-and-type';
      const url = `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}${firstPage}`;
      
      console.log(`   Navigating to: ${url}`);
      await this.page.goto(url, { waitUntil: 'networkidle2' });
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      const currentUrl = this.page.url();
      
      // Check if we were redirected to login page
      if (currentUrl.includes('login') || currentUrl.includes('Login')) {
        console.log('   ⚠️  Redirected to login page - session expired');
        return false;
      }
      
      // Check if we're on the CAS page (not redirected)
      if (currentUrl.includes('cas.cmmiinstitute.com') && currentUrl.includes(CONFIG.appraisalId)) {
        console.log('   ✅ Session is valid - already authenticated!');
        return true;
      }
      
      // Unknown state, assume login needed
      console.log(`   ⚠️  Unexpected URL: ${currentUrl}`);
      return false;
      
    } catch (error) {
      console.log('   ⚠️  Navigation failed, will try login');
      return false;
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
      
      // Handle "Stay signed in" checkbox
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

  async handleStaySignedIn(): Promise<void> {
    if (!this.page) return;
    
    try {
      // Common selectors for "Stay signed in" / "Remember me" checkboxes
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

  // Handle organizations page - check if org exists (edit mode) or not (add mode)
  async handleOrganizationsPage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   🔍 Checking organizations page state...');
    
    try {
      // Check if an organization card exists (meaning org already added)
      const orgCardExists = await this.page.evaluate(() => {
        return document.querySelector('.item-card') !== null;
      });
      
      if (orgCardExists) {
        // Scenario 1: Organization exists - click Edit button
        console.log('   📋 Organization already exists - switching to Edit mode');
        
        const editClicked = await this.page.evaluate(() => {
          // Find the Edit button in the item-card__actions
          const editLink = document.querySelector('.item-card__actions a.button[href*="Edit"]') as HTMLAnchorElement;
          if (editLink) {
            editLink.click();
            return true;
          }
          return false;
        });
        
        if (editClicked) {
          console.log('   ✅ Clicked "Edit" button');
          // Wait for edit form to load
          await new Promise(resolve => setTimeout(resolve, 2000));
          
          // Wait for the Name field to be available
          try {
            await this.page.waitForSelector('#Name', { timeout: 10000 });
            console.log('   ✅ Edit form loaded');
          } catch (e) {
            console.log('   ⚠️  Edit form may not have loaded properly');
          }
        } else {
          console.log('   ⚠️  Could not find Edit button');
        }
      } else {
        // Scenario 2: No organization yet - form should already be visible
        console.log('   📝 No organization yet - Add mode (form already visible)');
        
        // Verify the form is visible
        try {
          await this.page.waitForSelector('#Name', { timeout: 5000 });
          console.log('   ✅ Add form is ready');
        } catch (e) {
          console.log('   ⚠️  Form fields may not be visible');
        }
      }
    } catch (error) {
      console.log('   ⚠️  Error checking organizations page:', error);
    }
  }

  // Handle org-units page - check if OU exists (edit mode) or not (add mode)
  async handleOrgUnitsPage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   🔍 Checking org-units page state...');
    
    try {
      // Check if an OU card exists (meaning OU already added)
      // The item-card-list will contain item-card elements if OUs exist
      const ouCardExists = await this.page.evaluate(() => {
        const cardList = document.querySelector('.item-card-list');
        if (!cardList) return false;
        // Check if there's an actual item-card inside (not just empty list)
        return cardList.querySelector('.item-card') !== null;
      });
      
      if (ouCardExists) {
        // Scenario 1: OU exists - click Edit button
        console.log('   📋 Organizational Unit already exists - switching to Edit mode');
        
        const editClicked = await this.page.evaluate(() => {
          // Find the Edit button/link in the item-card
          // It could be an <a> tag with "Edit" text or href containing "Edit"
          const editLinks = Array.from(document.querySelectorAll('.item-card a.button, .item-card-list a.button'));
          for (const link of editLinks) {
            const href = link.getAttribute('href') || '';
            const text = link.textContent || '';
            if (href.toLowerCase().includes('edit') || text.toLowerCase().includes('edit')) {
              (link as HTMLElement).click();
              return true;
            }
          }
          
          // Fallback: look for any Edit link on the page
          const allLinks = Array.from(document.querySelectorAll('a'));
          for (const link of allLinks) {
            const href = link.getAttribute('href') || '';
            if (href.includes('/org-units/') && href.includes('Edit')) {
              (link as HTMLElement).click();
              return true;
            }
          }
          
          return false;
        });
        
        if (editClicked) {
          console.log('   ✅ Clicked "Edit" button');
          // Wait for edit form to load
          await new Promise(resolve => setTimeout(resolve, 2000));
          
          // Wait for the Name field to be available
          try {
            await this.page.waitForSelector('#Name', { timeout: 10000 });
            console.log('   ✅ Edit form loaded');
          } catch (e) {
            console.log('   ⚠️  Edit form may not have loaded properly');
          }
        } else {
          console.log('   ⚠️  Could not find Edit button - will try to use existing form');
        }
      } else {
        // Scenario 2: No OU yet - form should already be visible for adding
        console.log('   📝 No Organizational Unit yet - Add mode (form already visible)');
        
        // Verify the form is visible
        try {
          await this.page.waitForSelector('#Name', { timeout: 5000 });
          console.log('   ✅ Add form is ready');
        } catch (e) {
          console.log('   ⚠️  Form fields may not be visible');
        }
      }
    } catch (error) {
      console.log('   ⚠️  Error checking org-units page:', error);
    }
  }

  // Handle org-unit-targets page - check if target exists (edit mode) or not (add mode)
  async handleOrgUnitTargetsPage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   🔍 Checking org-unit-targets page state...');
    
    try {
      // Check if a target card exists (meaning target already added)
      const targetCardExists = await this.page.evaluate(() => {
        const cardList = document.querySelector('.item-card-list');
        if (!cardList) return false;
        return cardList.querySelector('.item-card') !== null;
      });
      
      if (targetCardExists) {
        // Scenario 1: Target exists - click Edit button
        console.log('   📋 Target Level already exists - switching to Edit mode');
        
        const editClicked = await this.page.evaluate(() => {
          // Find the Edit button/link
          const editLinks = Array.from(document.querySelectorAll('.item-card a.button, .item-card-list a.button, a.button'));
          for (const link of editLinks) {
            const href = link.getAttribute('href') || '';
            const text = link.textContent || '';
            if (href.toLowerCase().includes('edit') || text.toLowerCase().includes('edit')) {
              (link as HTMLElement).click();
              return true;
            }
          }
          return false;
        });
        
        if (editClicked) {
          console.log('   ✅ Clicked "Edit" button');
          await new Promise(resolve => setTimeout(resolve, 2000));
          
          try {
            await this.page.waitForSelector('#maturity-level', { timeout: 10000 });
            console.log('   ✅ Edit form loaded');
          } catch (e) {
            console.log('   ⚠️  Edit form may not have loaded properly');
          }
        } else {
          console.log('   ⚠️  Could not find Edit button');
        }
      } else {
        // Scenario 2: No target yet - form should already be visible
        console.log('   📝 No Target Level yet - Add mode (form already visible)');
        
        try {
          await this.page.waitForSelector('#maturity-level', { timeout: 5000 });
          console.log('   ✅ Add form is ready');
        } catch (e) {
          console.log('   ⚠️  Form fields may not be visible');
        }
      }
    } catch (error) {
      console.log('   \u26a0\ufe0f  Error checking org-unit-targets page:', error);
    }
  }

  // Handle the timeline page which has multiple sub-forms
  // Note: Readiness Reviews are handled on /readiness-reviews page separately
  async handleTimelinePage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   \ud83d\udcc5 Processing Timeline page (multi-form)...');
    
    // Debug: Show what data we have loaded
    console.log('\n   DEBUG: Excel data for P1-OrgScope timeline rows:');
    console.log(`      Row 69: ${this.excelData['P1-OrgScope']?.[69] || 'NOT LOADED'}`);
    console.log(`      Row 70: ${this.excelData['P1-OrgScope']?.[70] || 'NOT LOADED'}`);
    console.log(`      Row 85: ${this.excelData['P1-OrgScope']?.[85] || 'NOT LOADED'}`);
    console.log(`      Row 87: ${this.excelData['P1-OrgScope']?.[87] || 'NOT LOADED'}`);
    console.log(`      Row 88: ${this.excelData['P1-OrgScope']?.[88] || 'NOT LOADED'}`);
    console.log(`      Row 91: ${this.excelData['P1-OrgScope']?.[91] || 'NOT LOADED'}`);
    console.log(`      Row 92: ${this.excelData['P1-OrgScope']?.[92] || 'NOT LOADED'}`);
    
    const appraisalId = CONFIG.appraisalId;
    const baseUrl = CONFIG.casBaseUrl;
    
    try {
      // --- Phase 1: Plan Appraisal ---
      console.log('\n   === Phase 1: Plan Appraisal ===');
      await this.page.goto(`${baseUrl}/appraisals/${appraisalId}/timeline?EditPhase=Phase1#Form`, { waitUntil: 'networkidle2' });
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // Save HTML for debugging
      await this.savePageHtml('/timeline-phase1', 'before');
      
      // Populate Phase 1 fields using date-parts (Year/Month/Day number inputs)
      const phase1StartDate = this.excelData['P1-OrgScope']?.[69];
      const phase1EndDate = this.excelData['P1-OrgScope']?.[70];
      
      console.log(`   DEBUG: phase1StartDate = "${phase1StartDate}"`);
      console.log(`   DEBUG: phase1EndDate = "${phase1EndDate}"`);
      
      if (phase1StartDate) {
        console.log(`   Setting Phase 1 Start Date: ${phase1StartDate}`);
        await this.populateDateParts('#StartDateYear,#StartDateMonth,#StartDateDay', phase1StartDate);
      } else {
        console.log(`   \u26a0\ufe0f  No Phase 1 Start Date found in Excel data!`);
      }
      if (phase1EndDate) {
        console.log(`   Setting Phase 1 End Date: ${phase1EndDate}`);
        await this.populateDateParts('#EndDateYear,#EndDateMonth,#EndDateDay', phase1EndDate);
      } else {
        console.log(`   \u26a0\ufe0f  No Phase 1 End Date found in Excel data!`);
      }
      await this.clickSaveButton();
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // --- Phase 2: Conduct Appraisal ---
      console.log('\n   === Phase 2: Conduct Appraisal ===');
      await this.page.goto(`${baseUrl}/appraisals/${appraisalId}/timeline?EditPhase=Phase2#Form`, { waitUntil: 'networkidle2' });
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // Save HTML for debugging
      await this.savePageHtml('/timeline-phase2', 'before');
      
      const phase2StartDate = this.excelData['P1-OrgScope']?.[85];
      const phase2EndDate = this.excelData['P1-OrgScope']?.[87];
      const daysOnSite = this.excelData['P1-OrgScope']?.[88];
      
      if (phase2StartDate) {
        console.log(`   Setting Phase 2 Start Date: ${phase2StartDate}`);
        await this.populateDateParts('#StartDateYear,#StartDateMonth,#StartDateDay', phase2StartDate);
      }
      if (phase2EndDate) {
        console.log(`   Setting Phase 2 End Date: ${phase2EndDate}`);
        await this.populateDateParts('#EndDateYear,#EndDateMonth,#EndDateDay', phase2EndDate);
      }
      if (daysOnSite) {
        console.log(`   Setting Days On Site: ${daysOnSite}`);
        await this.populateNumberInput('#DaysOnSite', daysOnSite);
      }
      await this.clickSaveButton();
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // --- Phase 3: Report Results ---
      console.log('\n   === Phase 3: Report Results ===');
      await this.page.goto(`${baseUrl}/appraisals/${appraisalId}/timeline?EditPhase=Phase3#Form`, { waitUntil: 'networkidle2' });
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // Save HTML for debugging
      await this.savePageHtml('/timeline-phase3', 'before');
      
      const phase3StartDate = this.excelData['P1-OrgScope']?.[91];
      const phase3EndDate = this.excelData['P1-OrgScope']?.[92];
      
      if (phase3StartDate) {
        console.log(`   Setting Phase 3 Start Date: ${phase3StartDate}`);
        await this.populateDateParts('#StartDateYear,#StartDateMonth,#StartDateDay', phase3StartDate);
      }
      if (phase3EndDate) {
        console.log(`   Setting Phase 3 End Date: ${phase3EndDate}`);
        await this.populateDateParts('#EndDateYear,#EndDateMonth,#EndDateDay', phase3EndDate);
      }
      await this.clickSaveButton();
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // --- Readiness Reviews (on timeline page) ---
      // Readiness Reviews are added via the timeline page, not a separate page
      const readinessReviews = [
        {
          name: this.excelData['P1-OrgScope']?.[73],
          startDate: this.excelData['P1-OrgScope']?.[74],
          endDate: this.excelData['P1-OrgScope']?.[75],
        },
        {
          name: this.excelData['P1-OrgScope']?.[77],
          startDate: this.excelData['P1-OrgScope']?.[78],
          endDate: this.excelData['P1-OrgScope']?.[79],
        }
      ].filter(rr => rr.name); // Only include RRs with names
      
      for (let i = 0; i < readinessReviews.length; i++) {
        const rr = readinessReviews[i];
        console.log(`\n   === Readiness Review ${i + 1}: ${rr.name} ===`);
        
        // Navigate to timeline page first to check if RR exists
        await this.page.goto(`${baseUrl}/appraisals/${appraisalId}/timeline`, { waitUntil: 'networkidle2' });
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        // Check if this RR already exists
        const rrExists = await this.page.evaluate((rrName) => {
          const cards = document.querySelectorAll('.appraisal-timeline-readiness-review, .item-card');
          for (const card of Array.from(cards)) {
            if (card.textContent?.includes(rrName)) {
              return true;
            }
          }
          return false;
        }, rr.name);
        
        if (rrExists) {
          console.log(`   \u2139\ufe0f  Readiness Review "${rr.name}" already exists, skipping`);
          continue;
        }
        
        // Navigate to add readiness review form
        await this.page.goto(`${baseUrl}/appraisals/${appraisalId}/timeline?NewReadinessReview=true#Form`, { waitUntil: 'networkidle2' });
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        // Save HTML for debugging
        await this.savePageHtml(`/timeline-readiness-review-${i + 1}`, 'before');
        
        // Fill in Name
        if (rr.name) {
          console.log(`   Setting Name: ${rr.name}`);
          try {
            await this.populateTextInput('#Name', rr.name);
          } catch (e) {
            console.log(`   \u26a0\ufe0f  Could not fill Name field: ${e}`);
          }
        }
        
        // Fill in Start Date
        if (rr.startDate) {
          console.log(`   Setting Start Date: ${rr.startDate}`);
          await this.populateDateParts('#StartDateYear,#StartDateMonth,#StartDateDay', rr.startDate);
        }
        
        // Fill in End Date
        if (rr.endDate) {
          console.log(`   Setting End Date: ${rr.endDate}`);
          await this.populateDateParts('#EndDateYear,#EndDateMonth,#EndDateDay', rr.endDate);
        }
        
        // Save
        await this.clickSaveButton();
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        console.log(`   \u2705 Readiness Review ${i + 1} saved`);
      }
      
      // Return to timeline main page
      await this.page.goto(`${baseUrl}/appraisals/${appraisalId}/timeline`, { waitUntil: 'networkidle2' });
      
      console.log('\n   \u2705 Timeline page processing complete (including Readiness Reviews)');
      
    } catch (error) {
      console.log(`   \u26a0\ufe0f  Error processing timeline: ${error}`);
    }
  }

  // Handle the readiness-reviews page which has repeating sections
  // This page is for EDITING existing readiness reviews with additional fields
  // The basic readiness reviews (name, dates) are created on the timeline page
  async handleReadinessReviewsPage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   \ud83d\udccb Processing Readiness Reviews page (editing additional fields)...');
    
    const appraisalId = CONFIG.appraisalId;
    const baseUrl = CONFIG.casBaseUrl;
    
    // Readiness Review data from P1-OrgScope
    // First RR: rows 73-75 (name, start, end) + additional fields from rows 101-106
    // Second RR: rows 77-79 (name, start, end)
    // Note: Additional fields (101-106) are shared - we'll apply them to the first RR for now
    
    const readinessReviews = [
      {
        // First RR: rows 73-75 for name/dates, rows 101-106 for additional fields
        name: this.excelData['P1-OrgScope']?.[73],
        startDate: this.excelData['P1-OrgScope']?.[74],
        endDate: this.excelData['P1-OrgScope']?.[75],
        objectives: this.excelData['P1-OrgScope']?.[101],
        successCriteria: this.excelData['P1-OrgScope']?.[102],
        requiredMembers: this.excelData['P1-OrgScope']?.[103],
        outcomes: this.excelData['P1-OrgScope']?.[104],
        furtherDetails: this.excelData['P1-OrgScope']?.[105],
        characterizedEvidence: this.excelData['P1-OrgScope']?.[106],
      },
      {
        // Second RR: rows 77-79 for name/dates, rows 112-117 for additional fields
        name: this.excelData['P1-OrgScope']?.[77],
        startDate: this.excelData['P1-OrgScope']?.[78],
        endDate: this.excelData['P1-OrgScope']?.[79],
        objectives: this.excelData['P1-OrgScope']?.[112],
        successCriteria: this.excelData['P1-OrgScope']?.[113],
        requiredMembers: this.excelData['P1-OrgScope']?.[114],
        outcomes: this.excelData['P1-OrgScope']?.[115],
        furtherDetails: this.excelData['P1-OrgScope']?.[116],
        characterizedEvidence: this.excelData['P1-OrgScope']?.[117],
      }
    ].filter(rr => rr.name); // Only include RRs with names
    
    try {
      for (let i = 0; i < readinessReviews.length; i++) {
        const rr = readinessReviews[i];
        console.log(`\n   === Readiness Review ${i + 1}: ${rr.name} ===`);
        
        // Navigate to readiness reviews page
        await this.page.goto(`${baseUrl}/appraisals/${appraisalId}/readiness-reviews`, { waitUntil: 'networkidle2' });
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        // Save HTML for debugging
        await this.savePageHtml(`/readiness-reviews-${i + 1}`, 'before');
        
        // Find and click Edit button for this readiness review
        const editClicked = await this.page.evaluate((rrName) => {
          const cards = document.querySelectorAll('.item-card');
          for (const card of Array.from(cards)) {
            if (card.textContent?.includes(rrName)) {
              // Found the card, now find the Edit button
              const editBtn = card.querySelector('a.button[href*="Edit"]') as HTMLAnchorElement;
              if (editBtn) {
                editBtn.click();
                return { found: true, clicked: true };
              }
              return { found: true, clicked: false };
            }
          }
          return { found: false, clicked: false };
        }, rr.name);
        
        if (!editClicked.found) {
          console.log(`   \u26a0\ufe0f  Readiness Review "${rr.name}" not found on page`);
          continue;
        }
        
        if (!editClicked.clicked) {
          console.log(`   \u26a0\ufe0f  Could not find Edit button for "${rr.name}"`);
          continue;
        }
        
        console.log(`   \u2705 Clicked Edit for "${rr.name}"`);
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        // Wait for form to load
        try {
          await this.page.waitForSelector('#Name', { timeout: 5000 });
        } catch (e) {
          console.log(`   \u26a0\ufe0f  Edit form did not load`);
          continue;
        }
        
        // Fill in the additional fields
        // Objectives
        if (rr.objectives) {
          console.log(`   Setting Objectives...`);
          try {
            await this.populateTextInput('#Objectives', rr.objectives);
          } catch (e) {
            console.log(`   \u26a0\ufe0f  Could not fill Objectives: ${e}`);
          }
        }
        
        // Success Criteria
        if (rr.successCriteria) {
          console.log(`   Setting Success Criteria...`);
          try {
            await this.populateTextInput('#SuccessCriteria', rr.successCriteria);
          } catch (e) {
            console.log(`   \u26a0\ufe0f  Could not fill Success Criteria: ${e}`);
          }
        }
        
        // Required Members
        if (rr.requiredMembers) {
          console.log(`   Setting Required Members...`);
          try {
            await this.populateTextInput('#RequiredMembers', rr.requiredMembers);
          } catch (e) {
            console.log(`   \u26a0\ufe0f  Could not fill Required Members: ${e}`);
          }
        }
        
        // Outcomes
        if (rr.outcomes) {
          console.log(`   Setting Outcomes...`);
          try {
            await this.populateTextInput('#Outcomes', rr.outcomes);
          } catch (e) {
            console.log(`   \u26a0\ufe0f  Could not fill Outcomes: ${e}`);
          }
        }
        
        // Further Details
        if (rr.furtherDetails) {
          console.log(`   Setting Further Details...`);
          try {
            await this.populateTextInput('#FurtherDetails', rr.furtherDetails);
          } catch (e) {
            console.log(`   \u26a0\ufe0f  Could not fill Further Details: ${e}`);
          }
        }
        
        // Characterized Objective Evidence (Yes/No radio)
        if (rr.characterizedEvidence) {
          console.log(`   Setting Characterized Evidence: ${rr.characterizedEvidence}`);
          const isYes = rr.characterizedEvidence.toLowerCase() === 'yes';
          const radioSelector = isYes ? '#org-unit-project-sensitive-yes' : '#org-unit-project-sensitive-no';
          try {
            await this.page.click(radioSelector);
          } catch (e) {
            console.log(`   \u26a0\ufe0f  Could not set Characterized Evidence: ${e}`);
          }
        }
        
        // Click save/update button
        await this.clickSaveButton();
        await new Promise(resolve => setTimeout(resolve, 1500));
        
        console.log(`   \u2705 Readiness Review ${i + 1} updated`);
      }
      
      // Return to readiness reviews main page
      await this.page.goto(`${baseUrl}/appraisals/${appraisalId}/readiness-reviews`, { waitUntil: 'networkidle2' });
      
      console.log('\n   \u2705 Readiness Reviews page processing complete');
      
    } catch (error) {
      console.log(`   \u26a0\ufe0f  Error processing readiness reviews: ${error}`);
    }
  }

  // Handle the org-projects page - downloads template, populates from C_SupportV2, uploads
  async handleOrgProjectsPage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   📋 Processing Org Projects page (template download/upload workflow)...');
    
    const downloadDir = path.join(process.cwd(), 'downloads');
    
    // Ensure download directory exists
    if (!fs.existsSync(downloadDir)) {
      fs.mkdirSync(downloadDir, { recursive: true });
    }
    
    try {
      // Step 1: Download the template using fetch within authenticated browser context
      console.log('\n   === Step 1: Download Template ===');
      
      // Clear old files from download directory
      const oldFiles = fs.readdirSync(downloadDir);
      for (const file of oldFiles) {
        if (file.endsWith('.xlsx') || file.endsWith('.xls')) {
          fs.unlinkSync(path.join(downloadDir, file));
          console.log(`   🗑️  Deleted old file: ${file}`);
        }
      }
      
      // Find the download link URL
      const downloadUrl = await this.page.evaluate(() => {
        const links = Array.from(document.querySelectorAll('a'));
        for (const link of links) {
          const href = link.getAttribute('href') || '';
          const text = link.textContent || '';
          if (href.includes('/template/download') || 
              (href.includes('download') && (text.toLowerCase().includes('template') || href.includes('Template')))) {
            // Return absolute URL
            return link.href;
          }
        }
        return null;
      });
      
      if (!downloadUrl) {
        console.log('   ⚠️  Could not find download template link');
        // Try to find any download-related link for debugging
        const allLinks = await this.page.evaluate(() => {
          return Array.from(document.querySelectorAll('a')).map(a => ({
            href: a.getAttribute('href'),
            text: a.textContent?.trim().substring(0, 50)
          }));
        });
        console.log('   Available links:');
        allLinks.filter(l => l.href).forEach(l => console.log(`      ${l.text} -> ${l.href}`));
        return;
      }
      
      console.log(`   📥 Download URL: ${downloadUrl}`);
      
      // Use fetch within the page context to download the file (preserves cookies/auth)
      const downloadResult = await this.page.evaluate(async (url) => {
        try {
          const response = await fetch(url, {
            method: 'GET',
            credentials: 'include' // Include cookies for authentication
          });
          
          if (!response.ok) {
            return { success: false, error: `HTTP ${response.status}: ${response.statusText}` };
          }
          
          // Get the filename from Content-Disposition header or URL
          const contentDisposition = response.headers.get('Content-Disposition');
          let filename = 'template.xlsx';
          if (contentDisposition) {
            const match = contentDisposition.match(/filename[^;=\n]*=(["']?)([^"';\n]*)/i);
            if (match && match[2]) {
              filename = match[2];
            }
          }
          
          // Get the file as base64
          const blob = await response.blob();
          const reader = new FileReader();
          
          return new Promise((resolve) => {
            reader.onloadend = () => {
              const base64 = (reader.result as string).split(',')[1];
              resolve({ success: true, base64, filename });
            };
            reader.onerror = () => {
              resolve({ success: false, error: 'FileReader error' });
            };
            reader.readAsDataURL(blob);
          });
        } catch (e) {
          return { success: false, error: String(e) };
        }
      }, downloadUrl) as { success: boolean; base64?: string; filename?: string; error?: string };
      
      if (!downloadResult.success || !downloadResult.base64) {
        console.log(`   ⚠️  Download failed: ${downloadResult.error}`);
        return;
      }
      
      // Save the file from base64
      const templatePath = path.join(downloadDir, downloadResult.filename || 'template.xlsx');
      const buffer = Buffer.from(downloadResult.base64, 'base64');
      fs.writeFileSync(templatePath, buffer);
      console.log(`   ✅ Template downloaded: ${downloadResult.filename} (${buffer.length} bytes)`);
      
      // Step 2: Load C_SupportV2 data from CAS Plan Excel
      console.log('\n   === Step 2: Load C_SupportV2 Data ===');
      
      const sourceWorkbook = new ExcelJS.Workbook();
      await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
      const sourceSheet = sourceWorkbook.getWorksheet('C_SupportV2');
      
      if (!sourceSheet) {
        console.log('   ⚠️  C_SupportV2 sheet not found in source Excel');
        return;
      }
      
      // Helper function to extract cell value (handles formulas with formula resolution)
      const getCellValue = (cell: ExcelJS.Cell): any => {
        // Use class method to properly resolve formula references
        return this.resolveCellValue(sourceWorkbook, cell.value);
      };
      
      // Read header row to understand structure
      const headerRow = sourceSheet.getRow(1);
      const headers: string[] = [];
      headerRow.eachCell((cell, colNum) => {
        headers[colNum] = String(getCellValue(cell) || '');
      });
      console.log(`   Headers: ${headers.filter(h => h).join(', ').substring(0, 100)}...`);
      
      // Find where data ends (look for empty Column B - Function Name)
      // Column B contains formulas, so we need to check the formula result
      let lastDataRow = 1;
      for (let row = 2; row <= sourceSheet.rowCount + 10; row++) {
        const cellB = sourceSheet.getCell(row, 2);
        const valueB = getCellValue(cellB);
        
        if (valueB && String(valueB).trim() && String(valueB).trim() !== 'undefined') {
          lastDataRow = row;
        } else {
          break;
        }
      }
      
      console.log(`   ✅ C_SupportV2 data rows: 2 to ${lastDataRow} (${lastDataRow - 1} support functions)`);
      
      // Read source data (all columns A-R = columns 1-18)
      interface RowData {
        [col: number]: any;
      }
      const sourceData: RowData[] = [];
      for (let row = 2; row <= lastDataRow; row++) {
        const rowData: RowData = {};
        for (let col = 1; col <= 18; col++) {
          const cell = sourceSheet.getCell(row, col);
          const value = getCellValue(cell);
          // Skip 'undefined' string results from empty formula references
          if (value !== null && value !== undefined && String(value) !== 'undefined') {
            rowData[col] = value;
          }
        }
        const funcName = rowData[2] || rowData[1] || '(unnamed)';
        console.log(`      Row ${row}: ${String(funcName).substring(0, 40)}`);
        sourceData.push(rowData);
      }
      
      console.log(`   ✅ Loaded ${sourceData.length} support functions from C_SupportV2`);
      
      // Step 3: Populate the template (direct copy from rows 2-N, columns A-R)
      // Create a NEW workbook with the same structure since the downloaded template
      // has Google Sheets extensions that ExcelJS cannot parse
      console.log('\n   === Step 3: Populate Template ===');
      
      // Extract headers from downloaded template using AdmZip
      let templateHeaders: string[] = [];
      try {
        const zip = new AdmZip(templatePath);
        const sharedStringsXml = zip.readAsText('xl/sharedStrings.xml');
        // Extract first 18 strings (headers A-R)
        const stringMatches = sharedStringsXml.match(/<x:t[^>]*>([^<]*)<\/x:t>/g) || [];
        templateHeaders = stringMatches.slice(0, 18).map(m => {
          const match = m.match(/<x:t[^>]*>([^<]*)<\/x:t>/);
          return match ? match[1].replace(/\r\n/g, ' ').trim() : '';
        });
        console.log(`   Extracted ${templateHeaders.length} headers from template`);
      } catch (e) {
        console.log(`   ⚠️  Could not extract headers: ${e}`);
        // Use default headers matching C_SupportV2
        templateHeaders = [
          'Database ID (blank for new records)',
          'Function Name (required)',
          'Project Type',
          'Size (FTEs) (required)',
          'Function Description',
          'Is this function sensitive? (required)',
          'Point of Contact (required if project marked sensitive)',
          'Point of Contact\'s email address (required if project marked sensitive)',
          'Function Uses Suppliers (required)',
          'Current Manager\'s Name(s)',
          'Same as organization\'s address',
          'Address Line 1 (required)',
          'Address Line 2',
          'City (required)',
          'State/Province/Region (required)',
          'ZIP/Postal Code (required)',
          'Country/Region (required)',
          'Additional Function Information'
        ];
      }
      
      // Create new workbook
      const templateWorkbook = new ExcelJS.Workbook();
      const templateSheet = templateWorkbook.addWorksheet('Support Functions');
      
      // Add headers
      const templateHeaderRow = templateSheet.getRow(1);
      templateHeaders.forEach((header, idx) => {
        templateHeaderRow.getCell(idx + 1).value = header;
      });
      templateHeaderRow.font = { bold: true };
      
      console.log(`   Created new workbook with ${templateHeaders.length} columns`);
      
      // Copy data rows
      for (let i = 0; i < sourceData.length; i++) {
        const srcRow = sourceData[i];
        const targetRowNum = i + 2; // Starting from row 2 (1-indexed)
        const templateRow = templateSheet.getRow(targetRowNum);
        
        // Copy columns A-R (1-18)
        for (let col = 1; col <= 18; col++) {
          if (srcRow[col] !== undefined && srcRow[col] !== null) {
            templateRow.getCell(col).value = srcRow[col];
          }
        }
        
        const funcName = srcRow[2] || '(unnamed)';
        console.log(`      Row ${targetRowNum}: ${String(funcName).substring(0, 40)}`);
      }
      
      console.log(`   ✅ Copied ${sourceData.length} rows to template`);
      
      // Save the populated template
      const populatedFileName = 'populated_' + (downloadResult.filename || 'template.xlsx');
      const populatedTemplatePath = path.join(downloadDir, populatedFileName);
      await templateWorkbook.xlsx.writeFile(populatedTemplatePath);
      console.log(`   ✅ Saved populated template: ${populatedFileName}`);
      
      // Step 4: Upload the populated template
      console.log('\n   === Step 4: Upload Populated Template ===');
      
      // Find the file input element
      const fileInput = await this.page.$('input[type="file"]');
      
      if (fileInput) {
        await fileInput.uploadFile(populatedTemplatePath);
        console.log('   ✅ File selected for upload');
        
        // Wait for upload processing
        await new Promise(resolve => setTimeout(resolve, 3000));
        
        // Click upload/import button if present
        const uploadClicked = await this.page.evaluate(() => {
          const buttons = Array.from(document.querySelectorAll('button, input[type="submit"]'));
          for (const btn of buttons) {
            const text = (btn.textContent || (btn as HTMLInputElement).value || '').toLowerCase();
            if (text.includes('upload') || text.includes('import') || text.includes('submit')) {
              (btn as HTMLElement).click();
              return btn.textContent?.trim() || 'button';
            }
          }
          return null;
        });
        
        if (uploadClicked) {
          console.log(`   ✅ Upload button clicked: "${uploadClicked}"`);
          await new Promise(resolve => setTimeout(resolve, 5000));
        } else {
          console.log('   ℹ️  No upload button found - file may auto-submit');
        }
      } else {
        console.log('   ⚠️  File input not found - template may need manual upload');
        console.log(`   📁 Populated template saved at: ${populatedTemplatePath}`);
      }
      
      console.log('\n   ✅ Org Projects page processing complete');
      
    } catch (error) {
      console.log(`   ⚠️  Error processing org-projects: ${error}`);
      if (error instanceof Error) {
        console.log(`   Stack: ${error.stack}`);
      }
    }
  }

  // Handle the OU projects page - downloads template, populates from C_ProjectsV2, uploads
  // This is for "OU Sample Eligible Projects" page at /org-unit-projects/true
  async handleOUProjectsPage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   📋 Processing OU Projects page (template download/upload workflow)...');
    
    const downloadDir = path.join(process.cwd(), 'downloads');
    
    // Ensure download directory exists
    if (!fs.existsSync(downloadDir)) {
      fs.mkdirSync(downloadDir, { recursive: true });
    }
    
    try {
      // Step 1: Download the template using fetch within authenticated browser context
      console.log('\n   === Step 1: Download Template ===');
      
      // Clear old project files from download directory
      const oldFiles = fs.readdirSync(downloadDir);
      for (const file of oldFiles) {
        if (file.includes('Project') && (file.endsWith('.xlsx') || file.endsWith('.xls'))) {
          fs.unlinkSync(path.join(downloadDir, file));
          console.log(`   🗑️  Deleted old file: ${file}`);
        }
      }
      
      // Find the download link URL
      const downloadUrl = await this.page.evaluate(() => {
        const links = Array.from(document.querySelectorAll('a'));
        for (const link of links) {
          const href = link.getAttribute('href') || '';
          const text = link.textContent || '';
          if (href.includes('/template/download') || 
              (href.includes('download') && (text.toLowerCase().includes('template') || href.includes('Template')))) {
            return link.href;
          }
        }
        return null;
      });
      
      if (!downloadUrl) {
        console.log('   ⚠️  Could not find download template link');
        const allLinks = await this.page.evaluate(() => {
          return Array.from(document.querySelectorAll('a')).map(a => ({
            href: a.getAttribute('href'),
            text: a.textContent?.trim().substring(0, 50)
          }));
        });
        console.log('   Available links:');
        allLinks.filter(l => l.href).forEach(l => console.log(`      ${l.text} -> ${l.href}`));
        return;
      }
      
      console.log(`   📥 Download URL: ${downloadUrl}`);
      
      // Use fetch within the page context to download the file (preserves cookies/auth)
      const downloadResult = await this.page.evaluate(async (url) => {
        try {
          const response = await fetch(url, {
            method: 'GET',
            credentials: 'include'
          });
          
          if (!response.ok) {
            return { success: false, error: `HTTP ${response.status}: ${response.statusText}` };
          }
          
          const contentDisposition = response.headers.get('Content-Disposition');
          let filename = 'projects_template.xlsx';
          if (contentDisposition) {
            const match = contentDisposition.match(/filename[^;=\n]*=(["']?)([^"';\n]*)/i);
            if (match && match[2]) {
              filename = match[2];
            }
          }
          
          const blob = await response.blob();
          const reader = new FileReader();
          
          return new Promise((resolve) => {
            reader.onloadend = () => {
              const base64 = (reader.result as string).split(',')[1];
              resolve({ success: true, base64, filename });
            };
            reader.onerror = () => {
              resolve({ success: false, error: 'FileReader error' });
            };
            reader.readAsDataURL(blob);
          });
        } catch (e) {
          return { success: false, error: String(e) };
        }
      }, downloadUrl) as { success: boolean; base64?: string; filename?: string; error?: string };
      
      if (!downloadResult.success || !downloadResult.base64) {
        console.log(`   ⚠️  Download failed: ${downloadResult.error}`);
        return;
      }
      
      // Save the file from base64
      const templatePath = path.join(downloadDir, downloadResult.filename || 'projects_template.xlsx');
      const buffer = Buffer.from(downloadResult.base64, 'base64');
      fs.writeFileSync(templatePath, buffer);
      console.log(`   ✅ Template downloaded: ${downloadResult.filename} (${buffer.length} bytes)`);
      
      // Step 2: Load C_ProjectsV2 data from CAS Plan Excel
      console.log('\n   === Step 2: Load C_ProjectsV2 Data ===');
      
      const sourceWorkbook = new ExcelJS.Workbook();
      await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
      const sourceSheet = sourceWorkbook.getWorksheet('C_ProjectsV2');
      
      if (!sourceSheet) {
        console.log('   ⚠️  C_ProjectsV2 sheet not found in source Excel');
        return;
      }
      
      // Helper function to extract cell value (handles formulas with formula resolution)
      const getCellValue = (cell: ExcelJS.Cell): any => {
        const val = cell.value;
        // Handle Date objects specially (they need to stay as Date for Excel formatting)
        if (val instanceof Date) return val;
        // Use class method to properly resolve formula references
        const resolved = this.resolveCellValue(sourceWorkbook, val);
        // Return null for empty strings
        if (resolved === '') return null;
        return resolved;
      };
      
      // Read header row
      const headerRow = sourceSheet.getRow(1);
      const headers: string[] = [];
      headerRow.eachCell((cell, colNum) => {
        headers[colNum] = String(getCellValue(cell) || '');
      });
      console.log(`   Headers: ${headers.filter(h => h).slice(0, 10).join(', ').substring(0, 100)}...`);
      
      // Find where data ends (look for empty Column B - Project Name)
      let lastDataRow = 1;
      for (let row = 2; row <= sourceSheet.rowCount + 10; row++) {
        const cellB = sourceSheet.getCell(row, 2);
        const valueB = getCellValue(cellB);
        
        if (valueB && String(valueB).trim() && String(valueB).trim() !== 'undefined') {
          lastDataRow = row;
        } else {
          break;
        }
      }
      
      console.log(`   ✅ C_ProjectsV2 data rows: 2 to ${lastDataRow} (${lastDataRow - 1} projects)`);
      
      // Read source data (columns A-U = 1-21 for projects)
      interface RowData {
        [col: number]: any;
      }
      const sourceData: RowData[] = [];
      for (let row = 2; row <= lastDataRow; row++) {
        const rowData: RowData = {};
        for (let col = 1; col <= 21; col++) {
          const cell = sourceSheet.getCell(row, col);
          const value = getCellValue(cell);
          if (value !== null && value !== undefined && String(value) !== 'undefined') {
            rowData[col] = value;
          }
        }
        const projectName = rowData[2] || rowData[1] || '(unnamed)';
        console.log(`      Row ${row}: ${String(projectName).substring(0, 40)}`);
        sourceData.push(rowData);
      }
      
      console.log(`   ✅ Loaded ${sourceData.length} projects from C_ProjectsV2`);
      
      // Step 3: Populate the template
      console.log('\n   === Step 3: Populate Template ===');
      
      // Extract headers from downloaded template using AdmZip
      let templateHeaders: string[] = [];
      try {
        const zip = new AdmZip(templatePath);
        const sharedStringsXml = zip.readAsText('xl/sharedStrings.xml');
        const stringMatches = sharedStringsXml.match(/<x:t[^>]*>([^<]*)<\/x:t>/g) || [];
        templateHeaders = stringMatches.slice(0, 21).map(m => {
          const match = m.match(/<x:t[^>]*>([^<]*)<\/x:t>/);
          return match ? match[1].replace(/\r\n/g, ' ').trim() : '';
        });
        console.log(`   Extracted ${templateHeaders.length} headers from template`);
      } catch (e) {
        console.log(`   ⚠️  Could not extract headers: ${e}`);
        templateHeaders = [
          'Database ID (blank for new records)',
          'Project Name (required)',
          'Project Type',
          'Size (FTEs) (required)',
          'Project Description (required)',
          'Is this project sensitive? (required)',
          'Point of Contact',
          'Point of Contact\'s email address',
          'Current Life Cycle Phase (required)',
          'Start Date (required)',
          'Projected End Date (required)',
          'Project Uses Suppliers (required)',
          'Current Manager\'s Name(s) (required)',
          'Same as organization\'s address',
          'Address Line 1 (required)',
          'Address Line 2',
          'City (required)',
          'State/Province/Region (required)',
          'ZIP/Postal Code (required)',
          'Country/Region (required)',
          'Additional Project Information'
        ];
      }
      
      // Create new workbook
      const templateWorkbook = new ExcelJS.Workbook();
      const templateSheet = templateWorkbook.addWorksheet('Projects');
      
      // Add headers
      const templateHeaderRow = templateSheet.getRow(1);
      templateHeaders.forEach((header, idx) => {
        templateHeaderRow.getCell(idx + 1).value = header;
      });
      templateHeaderRow.font = { bold: true };
      
      console.log(`   Created new workbook with ${templateHeaders.length} columns`);
      
      // Copy data rows
      for (let i = 0; i < sourceData.length; i++) {
        const srcRow = sourceData[i];
        const targetRowNum = i + 2;
        const templateRow = templateSheet.getRow(targetRowNum);
        
        // Copy columns A-U (1-21)
        for (let col = 1; col <= 21; col++) {
          if (srcRow[col] !== undefined && srcRow[col] !== null) {
            const cell = templateRow.getCell(col);
            
            // Special handling for date columns J (10) and K (11)
            if ((col === 10 || col === 11) && srcRow[col] instanceof Date) {
              cell.value = srcRow[col];
              cell.numFmt = 'mm-dd-yyyy';  // Set date format
            } else {
              cell.value = srcRow[col];
            }
          }
        }
        
        const projectName = srcRow[2] || '(unnamed)';
        console.log(`      Row ${targetRowNum}: ${String(projectName).substring(0, 40)}`);
      }
      
      console.log(`   ✅ Copied ${sourceData.length} rows to template`);
      
      // Save the populated template
      const populatedFileName = 'populated_' + (downloadResult.filename || 'projects_template.xlsx');
      const populatedTemplatePath = path.join(downloadDir, populatedFileName);
      await templateWorkbook.xlsx.writeFile(populatedTemplatePath);
      console.log(`   ✅ Saved populated template: ${populatedFileName}`);
      
      // Step 4: Upload the populated template
      console.log('\n   === Step 4: Upload Populated Template ===');
      
      const fileInput = await this.page.$('input[type="file"]');
      
      if (fileInput) {
        await fileInput.uploadFile(populatedTemplatePath);
        console.log('   ✅ File selected for upload');
        
        await new Promise(resolve => setTimeout(resolve, 3000));
        
        const uploadClicked = await this.page.evaluate(() => {
          const buttons = Array.from(document.querySelectorAll('button, input[type="submit"]'));
          for (const btn of buttons) {
            const text = (btn.textContent || (btn as HTMLInputElement).value || '').toLowerCase();
            if (text.includes('upload') || text.includes('import') || text.includes('submit')) {
              (btn as HTMLElement).click();
              return btn.textContent?.trim() || 'button';
            }
          }
          return null;
        });
        
        if (uploadClicked) {
          console.log(`   ✅ Upload button clicked: "${uploadClicked}"`);
          await new Promise(resolve => setTimeout(resolve, 5000));
        } else {
          console.log('   ℹ️  No upload button found - file may auto-submit');
        }
      } else {
        console.log('   ⚠️  File input not found - template may need manual upload');
        console.log(`   📁 Populated template saved at: ${populatedTemplatePath}`);
      }
      
      console.log('\n   ✅ OU Projects page processing complete');
      
    } catch (error) {
      console.log(`   ⚠️  Error processing OU projects: ${error}`);
      if (error instanceof Error) {
        console.log(`   Stack: ${error.stack}`);
      }
    }
  }

  // Handle Sampling Factors page - /org-unit-sampling-factors
  // Data from P1-OrgScope rows 127-129
  // Form fields:
  //   select[name="StandardSamplingFactorId"] - Sampling Factor Name (options: Customer=2, Domain=6, Location=1, Organizational Structure=4, Other=-1, Size=3, Type of Work=5)
  //   input[name="OtherName"] - Other Sampling Factor (text, required if "Other" selected)
  //   textarea[name="Definition"] - Definition (textarea, required)
  async handleSamplingFactorsPage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   📋 Processing Sampling Factors page...');
    
    try {
      // Load data from P1-OrgScope
      const sourceWorkbook = new ExcelJS.Workbook();
      await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
      const sourceSheet = sourceWorkbook.getWorksheet('P1-OrgScope');
      
      if (!sourceSheet) {
        console.log('   ⚠️  P1-OrgScope sheet not found');
        return;
      }
      
      // Helper to get cell value (uses class method for formula resolution)
      const getCellValue = (row: number, col: number): string => {
        const cell = sourceSheet.getCell(row, col);
        return this.resolveCellValue(sourceWorkbook, cell.value);
      };
      
      // Get sampling factor data
      const samplingFactorName = getCellValue(127, 2);  // Row 127, Col B - e.g., "Type of Work"
      const otherSamplingFactor = getCellValue(128, 2); // Row 128, Col B - only if "Other" selected
      const definition = getCellValue(129, 2);          // Row 129, Col B
      
      console.log(`   Sampling Factor Name: ${samplingFactorName}`);
      console.log(`   Other: ${otherSamplingFactor || '(none)'}`);
      console.log(`   Definition: ${definition}`);
      
      // Check if an entry already exists with the same name
      const existingEntry = await this.page.evaluate((searchName) => {
        const cards = document.querySelectorAll('.item-card');
        for (const card of Array.from(cards)) {
          const title = card.querySelector('.item-card__title h3')?.textContent?.trim();
          if (title && title.toLowerCase() === searchName.toLowerCase()) {
            return { exists: true, title };
          }
        }
        return { exists: false };
      }, samplingFactorName);
      
      if (existingEntry.exists) {
        console.log(`   ℹ️  Sampling factor "${existingEntry.title}" already exists - skipping`);
        return;
      }
      
      // Check if form is already visible or if we need to click Add
      const formVisible = await this.page.$('select[name="StandardSamplingFactorId"]');
      
      if (!formVisible) {
        // Click "Add Sampling Factor" link/button
        console.log('   Adding new sampling factor...');
        const addClicked = await this.page.evaluate(() => {
          const links = Array.from(document.querySelectorAll('a.button, button'));
          for (const link of links) {
            const text = link.textContent?.toLowerCase() || '';
            if (text.includes('add') && text.includes('sampling')) {
              (link as HTMLElement).click();
              return true;
            }
          }
          return false;
        });
        
        if (addClicked) {
          await new Promise(resolve => setTimeout(resolve, 2000));
        }
      }
      
      // Fill in the form
      // 1. Sampling Factor Name (select dropdown)
      if (samplingFactorName) {
        try {
          await this.page.waitForSelector('select[name="StandardSamplingFactorId"]', { timeout: 5000 });
          const selectResult = await this.page.evaluate((searchText) => {
            const select = document.querySelector('select[name="StandardSamplingFactorId"]') as HTMLSelectElement;
            if (!select) return { found: false, error: 'Select not found' };
            
            for (const option of Array.from(select.options)) {
              if (option.text.toLowerCase().includes(searchText.toLowerCase())) {
                select.value = option.value;
                select.dispatchEvent(new Event('change', { bubbles: true }));
                return { found: true, value: option.text, optionValue: option.value };
              }
            }
            return { found: false, error: `Option not found: ${searchText}` };
          }, samplingFactorName);
          
          if (selectResult.found) {
            console.log(`   ✅ Selected: ${selectResult.value} (value=${selectResult.optionValue})`);
          } else {
            console.log(`   ⚠️  ${selectResult.error}`);
          }
        } catch (e) {
          console.log(`   ⚠️  Could not set sampling factor name: ${e}`);
        }
      }
      
      // 2. Other Sampling Factor (text input - only if "Other" was selected)
      if (otherSamplingFactor) {
        try {
          await this.page.waitForSelector('input[name="OtherName"], #OtherName', { timeout: 2000 });
          await this.page.evaluate((val) => {
            const el = document.querySelector('input[name="OtherName"], #OtherName') as HTMLInputElement;
            if (el) {
              el.value = val;
              el.dispatchEvent(new Event('input', { bubbles: true }));
            }
          }, otherSamplingFactor);
          console.log(`   ✅ Set Other Sampling Factor: ${otherSamplingFactor}`);
        } catch (e) {
          // This field may not be visible if "Other" wasn't selected
          console.log(`   ℹ️  Other Sampling Factor field not needed`);
        }
      }
      
      // 3. Definition (textarea - required)
      if (definition) {
        try {
          await this.page.waitForSelector('textarea[name="Definition"], #Definition', { timeout: 3000 });
          await this.page.evaluate((val) => {
            const el = document.querySelector('textarea[name="Definition"], #Definition') as HTMLTextAreaElement;
            if (el) {
              el.value = val;
              el.dispatchEvent(new Event('input', { bubbles: true }));
            }
          }, definition);
          console.log(`   ✅ Set Definition: ${definition.substring(0, 50)}...`);
        } catch (e) {
          console.log(`   ⚠️  Could not set definition: ${e}`);
        }
      }
      
      // Click Save/Add button
      const saveClicked = await this.page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll('button, input[type="submit"]'));
        for (const btn of buttons) {
          const text = (btn.textContent || (btn as HTMLInputElement).value || '').toLowerCase();
          if (text.includes('add') || text.includes('save') || text.includes('submit')) {
            (btn as HTMLElement).click();
            return btn.textContent?.trim() || 'button';
          }
        }
        return null;
      });
      
      if (saveClicked) {
        console.log(`   ✅ Clicked: "${saveClicked}"`);
        await new Promise(resolve => setTimeout(resolve, 3000));
      }
      
      console.log('\n   ✅ Sampling Factors page complete');
      
    } catch (error) {
      console.log(`   ⚠️  Error processing sampling factors: ${error}`);
    }
  }

  // Handle Sampling Factor Values page - /org-unit-sampling-factor-values
  // Data from P1-OrgScope rows 132-134
  // Form fields:
  //   select[name="AppraisalOrgUnitSamplingFactorId"] - Sampling Factor (dropdown of defined factors)
  //   input[name="Value"] - Value (text)
  //   textarea[name="Description"] - Description (textarea)
  async handleSamplingFactorValuesPage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   📋 Processing Sampling Factor Values page...');
    
    try {
      // Load data from P1-OrgScope
      const sourceWorkbook = new ExcelJS.Workbook();
      await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
      const sourceSheet = sourceWorkbook.getWorksheet('P1-OrgScope');
      
      if (!sourceSheet) {
        console.log('   ⚠️  P1-OrgScope sheet not found');
        return;
      }
      
      // Helper to get cell value (uses class method for formula resolution)
      const getCellValue = (row: number, col: number): string => {
        const cell = sourceSheet.getCell(row, col);
        return this.resolveCellValue(sourceWorkbook, cell.value);
      };
      
      // Get sampling factor value data
      const samplingFactor = getCellValue(132, 2);  // Row 132, Col B - e.g., "Type of Work"
      const value = getCellValue(133, 2);           // Row 133, Col B - e.g., "All projects"
      const description = getCellValue(134, 2);     // Row 134, Col B - e.g., "All development projects"
      
      console.log(`   Sampling Factor: ${samplingFactor}`);
      console.log(`   Value: ${value}`);
      console.log(`   Description: ${description}`);
      
      // Check if an entry already exists with the same value
      const existingEntry = await this.page.evaluate((searchValue) => {
        const cards = document.querySelectorAll('.item-card');
        for (const card of Array.from(cards)) {
          // Look for the Value section
          const sections = card.querySelectorAll('.item-card__section');
          for (const section of Array.from(sections)) {
            const header = section.querySelector('h3')?.textContent?.trim();
            const valueText = section.querySelector('small')?.textContent?.trim();
            if (header === 'Value' && valueText && valueText.toLowerCase() === searchValue.toLowerCase()) {
              return { exists: true, value: valueText };
            }
          }
        }
        return { exists: false };
      }, value);
      
      if (existingEntry.exists) {
        console.log(`   ℹ️  Sampling factor value "${existingEntry.value}" already exists - skipping`);
        return;
      }
      
      // Check if form is already visible or if we need to click Add
      const formVisible = await this.page.$('select[name="AppraisalOrgUnitSamplingFactorId"]');
      
      if (!formVisible) {
        // Click "Add" button
        console.log('   Adding new sampling factor value...');
        const addClicked = await this.page.evaluate(() => {
          const buttons = Array.from(document.querySelectorAll('a.button, button'));
          for (const btn of buttons) {
            const text = btn.textContent?.toLowerCase() || '';
            if (text.includes('add')) {
              (btn as HTMLElement).click();
              return true;
            }
          }
          return false;
        });
        
        if (addClicked) {
          await new Promise(resolve => setTimeout(resolve, 2000));
        }
      }
      
      // Fill in the form
      // 1. Sampling Factor (select dropdown)
      if (samplingFactor) {
        try {
          await this.page.waitForSelector('select[name="AppraisalOrgUnitSamplingFactorId"]', { timeout: 5000 });
          const selectResult = await this.page.evaluate((searchText) => {
            const select = document.querySelector('select[name="AppraisalOrgUnitSamplingFactorId"]') as HTMLSelectElement;
            if (!select) return { found: false, error: 'Select not found' };
            
            for (const option of Array.from(select.options)) {
              if (option.text.toLowerCase().includes(searchText.toLowerCase())) {
                select.value = option.value;
                select.dispatchEvent(new Event('change', { bubbles: true }));
                return { found: true, value: option.text, optionValue: option.value };
              }
            }
            return { found: false, error: `Option not found: ${searchText}` };
          }, samplingFactor);
          
          if (selectResult.found) {
            console.log(`   ✅ Selected Sampling Factor: ${selectResult.value} (value=${selectResult.optionValue})`);
          } else {
            console.log(`   ⚠️  ${selectResult.error}`);
          }
        } catch (e) {
          console.log(`   ⚠️  Could not set sampling factor: ${e}`);
        }
      }
      
      // 2. Value (text input)
      if (value) {
        try {
          await this.page.waitForSelector('input[name="Value"], #org-unit-sample-value', { timeout: 5000 });
          await this.page.evaluate((val) => {
            const el = document.querySelector('input[name="Value"], #org-unit-sample-value') as HTMLInputElement;
            if (el) {
              el.value = val;
              el.dispatchEvent(new Event('input', { bubbles: true }));
            }
          }, value);
          console.log(`   ✅ Set Value: ${value}`);
        } catch (e) {
          console.log(`   ⚠️  Could not set value: ${e}`);
        }
      }
      
      // 3. Description (textarea)
      if (description) {
        try {
          await this.page.waitForSelector('textarea[name="Description"], #org-unit-sample-description', { timeout: 3000 });
          await this.page.evaluate((val) => {
            const el = document.querySelector('textarea[name="Description"], #org-unit-sample-description') as HTMLTextAreaElement;
            if (el) {
              el.value = val;
              el.dispatchEvent(new Event('input', { bubbles: true }));
            }
          }, description);
          console.log(`   ✅ Set Description: ${description.substring(0, 50)}...`);
        } catch (e) {
          console.log(`   ⚠️  Could not set description: ${e}`);
        }
      }
      
      // Click Save/Add button
      const saveClicked = await this.page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll('button, input[type="submit"]'));
        for (const btn of buttons) {
          const text = (btn.textContent || (btn as HTMLInputElement).value || '').toLowerCase();
          if (text.includes('add') || text.includes('save') || text.includes('submit')) {
            (btn as HTMLElement).click();
            return btn.textContent?.trim() || 'button';
          }
        }
        return null;
      });
      
      if (saveClicked) {
        console.log(`   ✅ Clicked: "${saveClicked}"`);
        await new Promise(resolve => setTimeout(resolve, 3000));
      }
      
      console.log('\n   ✅ Sampling Factor Values page complete');
      
    } catch (error) {
      console.log(`   ⚠️  Error processing sampling factor values: ${error}`);
    }
  }

  // Handle Project Subgroups page - /org-unit-subgroups
  // Data from P1-OrgScope rows 137-139
  // Form fields:
  //   input[name="Name"] - Subgroup Name (text)
  //   input[name="Abbreviation"] - Abbreviation (text)
  //   input[type="checkbox"][id*="Selected"] - Sampling factor value checkboxes
  async handleSubgroupsPage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   📋 Processing Project Subgroups page...');
    
    try {
      // Load data from P1-OrgScope
      const sourceWorkbook = new ExcelJS.Workbook();
      await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
      const sourceSheet = sourceWorkbook.getWorksheet('P1-OrgScope');
      
      if (!sourceSheet) {
        console.log('   ⚠️  P1-OrgScope sheet not found');
        return;
      }
      
      // Helper to get cell value (uses class method for formula resolution)
      const getCellValue = (row: number, col: number): string => {
        const cell = sourceSheet.getCell(row, col);
        return this.resolveCellValue(sourceWorkbook, cell.value);
      };
      
      // Get subgroup data
      const name = getCellValue(137, 2);         // Row 137, Col B - e.g., "All projects"
      const abbreviation = getCellValue(138, 2); // Row 138, Col B - e.g., "AP"
      const checkboxInfo = getCellValue(139, 2); // Row 139, Col B - e.g., "[x] All projects" indicates which to check
      
      console.log(`   Name: ${name}`);
      console.log(`   Abbreviation: ${abbreviation}`);
      console.log(`   Checkbox: ${checkboxInfo}`);
      
      // Check if a subgroup already exists with the same name
      // Note: The title displays as "Name (Abbreviation)" e.g., "All projects (AP)"
      const existingEntry = await this.page.evaluate((searchName, searchAbbr) => {
        const cards = document.querySelectorAll('.item-card');
        for (const card of Array.from(cards)) {
          const title = card.querySelector('.item-card__title h3')?.textContent?.trim();
          if (title) {
            // Check if title matches "Name (Abbr)" format or just the name
            const expectedTitle = searchAbbr ? `${searchName} (${searchAbbr})` : searchName;
            if (title.toLowerCase() === expectedTitle.toLowerCase() ||
                title.toLowerCase().startsWith(searchName.toLowerCase())) {
              return { exists: true, title };
            }
          }
        }
        return { exists: false };
      }, name, abbreviation);
      
      if (existingEntry.exists) {
        console.log(`   ℹ️  Subgroup "${existingEntry.title}" already exists - skipping`);
        return;
      }
      
      // Check if form is already visible or if we need to click Add
      const formVisible = await this.page.$('input[name="Name"]');
      
      if (!formVisible) {
        // Click "Add" button
        console.log('   Adding new subgroup...');
        const addClicked = await this.page.evaluate(() => {
          const buttons = Array.from(document.querySelectorAll('a.button, button'));
          for (const btn of buttons) {
            const text = btn.textContent?.toLowerCase() || '';
            if (text.includes('add')) {
              (btn as HTMLElement).click();
              return true;
            }
          }
          return false;
        });
        
        if (addClicked) {
          await new Promise(resolve => setTimeout(resolve, 2000));
        }
      }
      
      // Fill in the form
      // 1. Name (text input)
      if (name) {
        try {
          await this.page.waitForSelector('input[name="Name"], #Name', { timeout: 5000 });
          await this.page.evaluate((val) => {
            const el = document.querySelector('input[name="Name"], #Name') as HTMLInputElement;
            if (el) {
              el.value = val;
              el.dispatchEvent(new Event('input', { bubbles: true }));
            }
          }, name);
          console.log(`   ✅ Set Name: ${name}`);
        } catch (e) {
          console.log(`   ⚠️  Could not set name: ${e}`);
        }
      }
      
      // 2. Abbreviation (text input)
      if (abbreviation) {
        try {
          await this.page.waitForSelector('input[name="Abbreviation"], #Abbreviation', { timeout: 3000 });
          await this.page.evaluate((val) => {
            const el = document.querySelector('input[name="Abbreviation"], #Abbreviation') as HTMLInputElement;
            if (el) {
              el.value = val;
              el.dispatchEvent(new Event('input', { bubbles: true }));
            }
          }, abbreviation);
          console.log(`   ✅ Set Abbreviation: ${abbreviation}`);
        } catch (e) {
          console.log(`   ⚠️  Could not set abbreviation: ${e}`);
        }
      }
      
      // 3. Check the sampling factor value checkbox(es)
      // The checkboxes have IDs like: SamplingFactors_0__Values_0__Selected
      // For simple case, we check all available checkboxes (or match by label text)
      try {
        const checkboxResult = await this.page.evaluate((checkboxLabel) => {
          // Find all sampling factor value checkboxes
          const checkboxes = document.querySelectorAll('input[type="checkbox"][id*="Selected"]');
          let checkedCount = 0;
          
          for (const checkbox of Array.from(checkboxes)) {
            const cb = checkbox as HTMLInputElement;
            // Get the associated label text
            const label = document.querySelector(`label[for="${cb.id}"]`);
            const labelText = label?.textContent?.trim() || '';
            
            // If checkboxLabel contains "[x]" followed by text, check matching checkbox
            // Or if simple mode, check all checkboxes
            const shouldCheck = checkboxLabel.includes('[x]') 
              ? checkboxLabel.toLowerCase().includes(labelText.toLowerCase())
              : true; // Default: check all
            
            if (shouldCheck && !cb.checked) {
              cb.checked = true;
              cb.dispatchEvent(new Event('change', { bubbles: true }));
              checkedCount++;
            }
          }
          
          return { total: checkboxes.length, checked: checkedCount };
        }, checkboxInfo);
        
        console.log(`   ✅ Checked ${checkboxResult.checked} of ${checkboxResult.total} sampling factor value(s)`);
      } catch (e) {
        console.log(`   ⚠️  Could not check sampling factor values: ${e}`);
      }
      
      // Click Save/Add button
      const saveClicked = await this.page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll('button, input[type="submit"]'));
        for (const btn of buttons) {
          const text = (btn.textContent || (btn as HTMLInputElement).value || '').toLowerCase();
          if (text.includes('add') || text.includes('save') || text.includes('submit')) {
            (btn as HTMLElement).click();
            return btn.textContent?.trim() || 'button';
          }
        }
        return null;
      });
      
      if (saveClicked) {
        console.log(`   ✅ Clicked: "${saveClicked}"`);
        await new Promise(resolve => setTimeout(resolve, 3000));
      }
      
      console.log('\n   ✅ Project Subgroups page complete');
      
    } catch (error) {
      console.log(`   ⚠️  Error processing subgroups: ${error}`);
    }
  }

  // Handle Project Subgroup Assignment page - /org-unit-project-subgroups
  // This page assigns each project to a subgroup
  // For simple case: select the first (and only) subgroup option for all projects
  async handleSubgroupAssignmentPage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   📋 Processing Project Subgroup Assignment page...');
    
    try {
      // Check if form is visible
      const formVisible = await this.page.$('form.org-unit-project-subgroups__form');
      
      if (!formVisible) {
        console.log('   ⚠️  Subgroup assignment form not found');
        return;
      }
      
      // Select the first available subgroup option for all project dropdowns
      const result = await this.page.evaluate(() => {
        const selects = document.querySelectorAll('select[name*="SubgroupId"]');
        let assignedCount = 0;
        let alreadyAssigned = 0;
        
        for (const select of Array.from(selects)) {
          const sel = select as HTMLSelectElement;
          
          // Find the first non-empty option
          let firstValidOption: HTMLOptionElement | null = null;
          for (const option of Array.from(sel.options)) {
            if (option.value && option.value !== '') {
              firstValidOption = option;
              break;
            }
          }
          
          if (firstValidOption) {
            if (sel.value === firstValidOption.value) {
              // Already assigned
              alreadyAssigned++;
            } else {
              // Assign to first subgroup
              sel.value = firstValidOption.value;
              sel.dispatchEvent(new Event('change', { bubbles: true }));
              assignedCount++;
            }
          }
        }
        
        return { 
          total: selects.length, 
          assigned: assignedCount, 
          alreadyAssigned: alreadyAssigned 
        };
      });
      
      console.log(`   ✅ Assigned ${result.assigned} projects to subgroup (${result.alreadyAssigned} already assigned, ${result.total} total)`);
      
      // Click Save button
      const saveClicked = await this.page.evaluate(() => {
        const buttons = Array.from(document.querySelectorAll('button, input[type="submit"]'));
        for (const btn of buttons) {
          const text = (btn.textContent || (btn as HTMLInputElement).value || '').toLowerCase();
          if (text.includes('save')) {
            (btn as HTMLElement).click();
            return btn.textContent?.trim() || 'button';
          }
        }
        return null;
      });
      
      if (saveClicked) {
        console.log(`   ✅ Clicked: "${saveClicked}"`);
        await new Promise(resolve => setTimeout(resolve, 3000));
      }
      
      console.log('\n   ✅ Project Subgroup Assignment page complete');
      
    } catch (error) {
      console.log(`   ⚠️  Error processing subgroup assignment: ${error}`);
    }
  }

  // Handle Organizational Support Function PA Exceptions page - /organizational-project-appraisal-scope
  // Data from P1-OrgScope rows 147-154
  // This page requires selecting each support function, then setting PA exceptions
  async handleOrgProjectAppraisalScopePage(): Promise<void> {
    if (!this.page) return;
    
    console.log('\n   📋 Processing Organizational Support Function PA Exceptions page...');
    
    try {
      // Load data from P1-OrgScope
      const sourceWorkbook = new ExcelJS.Workbook();
      await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
      const sourceSheet = sourceWorkbook.getWorksheet('P1-OrgScope');
      
      if (!sourceSheet) {
        console.log('   ⚠️  P1-OrgScope sheet not found');
        return;
      }
      
      // Helper to get cell value (uses class method for formula resolution)
      const getCellValue = (row: number, col: number): string => {
        const cell = sourceSheet.getCell(row, col);
        return this.resolveCellValue(sourceWorkbook, cell.value);
      };
      
      // PA abbreviation to full name mapping
      const paNameMap: { [key: string]: string } = {
        'CM': 'Configuration Management',
        'MPM': 'Managing Performance and Measurement',
        'PAD': 'Process Asset Development',
        'PCM': 'Process Management',
        'CAR': 'Causal Analysis and Resolution',
        'DAR': 'Decision Analysis and Resolution',
        'OT': 'Organizational Training',
        'PQA': 'Process Quality Assurance',
        'PI': 'Product Integration',
        'TS': 'Technical Solution',
        'PR': 'Peer Reviews',
        'RDM': 'Requirements Development and Management',
        'VV': 'Verification and Validation',
        'RSK': 'Risk and Opportunity Management',
        'EST': 'Estimating',
        'MC': 'Monitor and Control',
        'PLAN': 'Planning',
      };
      
      // Scope value mapping
      const scopeMap: { [key: string]: string } = {
        'IS': 'InScope',
        'IS-OoS': 'InScopeDefaultOthersToOutOfScope',
        'OoS': 'OutOfScope',
      };
      
      // Parse support function exceptions from Excel
      // Format: "PA1 (IS), PA2 (IS-OoS), ..." 
      const parseExceptions = (exceptionStr: string): Array<{pa: string, scope: string}> => {
        const exceptions: Array<{pa: string, scope: string}> = [];
        if (!exceptionStr) return exceptions;
        
        // Split by comma and parse each
        const parts = exceptionStr.split(',').map(p => p.trim());
        for (const part of parts) {
          // Match pattern like "CM (IS)" or "PAD (IS-OoS)"
          const match = part.match(/^(\w+)\s*\(([^)]+)\)$/);
          if (match) {
            const paAbbr = match[1].trim();
            const scopeAbbr = match[2].trim();
            const paName = paNameMap[paAbbr];
            const scopeValue = scopeMap[scopeAbbr];
            
            if (paName && scopeValue) {
              exceptions.push({ pa: paName, scope: scopeValue });
            } else {
              console.log(`      ⚠️  Unknown PA or scope: ${paAbbr} (${scopeAbbr})`);
            }
          }
        }
        return exceptions;
      };
      
      // NEW data structure from P1-OrgScope (v06):
      // Each support function block:
      //   Row N: "Select support function" | "S1-CM" (or S2-EPG, S3-HR, S4-QA)
      //   Row N+1: "Practice Area" | "Configuration Management"
      //   Row N+2: "Select scope" | "In-Scope" (or "In-Scope (default other Projects to Out-of-Scope)")
      //   Row N+3: "Justification for in-Scope" | "The process is performed..."
      //   (repeat for additional PAs)
      //   Empty row separates support functions
      
      // Parse all support functions and their PA exceptions from the new format
      interface PAException {
        practiceArea: string;
        scope: string;
        justification: string;
      }
      interface SupportFunction {
        name: string;
        exceptions: PAException[];
      }
      
      const supportFunctions: SupportFunction[] = [];
      let currentSF: SupportFunction | null = null;
      let currentPA: Partial<PAException> = {};
      
      // Scan rows 147-178 to build support function data
      for (let row = 147; row <= 178; row++) {
        const colA = getCellValue(row, 1);
        const colB = getCellValue(row, 2);
        
        if (!colA && !colB) {
          // Empty row - save current PA if any, then continue
          if (currentPA.practiceArea && currentSF) {
            currentSF.exceptions.push(currentPA as PAException);
            currentPA = {};
          }
          continue;
        }
        
        const colALower = colA.toLowerCase();
        
        if (colALower.includes('select support function') || (colALower === 'select' && colB.startsWith('S'))) {
          // New support function - save previous one if exists
          if (currentPA.practiceArea && currentSF) {
            currentSF.exceptions.push(currentPA as PAException);
            currentPA = {};
          }
          if (currentSF && currentSF.exceptions.length > 0) {
            supportFunctions.push(currentSF);
          }
          currentSF = { name: colB, exceptions: [] };
        } else if (colALower === 'practice area') {
          // Save previous PA if exists
          if (currentPA.practiceArea && currentSF) {
            currentSF.exceptions.push(currentPA as PAException);
          }
          currentPA = { practiceArea: colB };
        } else if (colALower === 'select scope') {
          currentPA.scope = colB;
        } else if (colALower.includes('justification')) {
          currentPA.justification = colB;
        }
      }
      
      // Don't forget the last PA and SF
      if (currentPA.practiceArea && currentSF) {
        currentSF.exceptions.push(currentPA as PAException);
      }
      if (currentSF && currentSF.exceptions.length > 0) {
        supportFunctions.push(currentSF);
      }
      
      console.log(`   Found ${supportFunctions.length} support functions with exceptions`);
      
      // Map scope text to dropdown value
      const scopeTextToValue: { [key: string]: string } = {
        'in-scope': 'InScope',
        'in-scope (default other projects to out-of-scope)': 'InScopeDefaultOthersToOutOfScope',
        'out-of-scope': 'OutOfScope',
      };
      
      for (const sf of supportFunctions) {
        console.log(`\n   === Processing ${sf.name} ===`);
        console.log(`   PA Exceptions: ${sf.exceptions.length}`);
        for (const exc of sf.exceptions) {
          console.log(`      - ${exc.practiceArea}: ${exc.scope}`);
        }
        
        // Step 1: Select the support function from the OrgUnitProjectId dropdown
        const selectResult = await this.page.evaluate((sfName) => {
          // Specifically target the dropdown in the organizational-project-appraisal-scope__project-selector form
          const projectForm = document.querySelector('.organizational-project-appraisal-scope__project-selector form');
          if (!projectForm) return { found: false, error: 'Project selector form not found' };
          
          const select = projectForm.querySelector('select[name="OrgUnitProjectId"]') as HTMLSelectElement;
          if (!select) return { found: false, error: 'OrgUnitProjectId select not found' };
          
          for (const option of Array.from(select.options)) {
            if (option.text.trim() === sfName || option.text.includes(sfName)) {
              select.value = option.value;
              select.dispatchEvent(new Event('change', { bubbles: true }));
              return { found: true, value: option.text.trim(), optionValue: option.value };
            }
          }
          return { found: false, error: `Option not found for ${sfName}` };
        }, sf.name);
        
        if (!selectResult.found) {
          console.log(`   ⚠️  ${selectResult.error}`);
          continue;
        }
        
        console.log(`   ✅ Selected: ${selectResult.value}`);
        
        // Step 2: Click the Select button in the project selector form (NOT the org unit target form)
        console.log(`   Clicking Select button...`);
        await Promise.all([
          this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 }).catch(() => {}),
          this.page.evaluate(() => {
            // Specifically target the button in the project selector form
            const projectForm = document.querySelector('.organizational-project-appraisal-scope__project-selector form');
            if (projectForm) {
              const btn = projectForm.querySelector('button');
              if (btn) {
                btn.click();
                return true;
              }
            }
            return false;
          })
        ]);
        
        // Wait for the form to fully load
        await new Promise(resolve => setTimeout(resolve, 2000));
        
        // Step 3: Verify the PA form loaded
        const formHeader = await this.page.evaluate(() => {
          const h2 = document.querySelector('#Form h2');
          return h2?.textContent?.trim() || '';
        });
        console.log(`   Form header: "${formHeader}"`);
        
        // Wait for PA selectors to be visible
        try {
          await this.page.waitForSelector('.project-appraisal-scope-form__practice-area-selector', { timeout: 5000 });
          console.log(`   ✅ PA form loaded`);
        } catch (e) {
          console.log(`   ⚠️  PA form did not load`);
          continue;
        }
        
        // Step 4: Set each PA exception from the new data structure
        console.log(`   Processing ${sf.exceptions.length} PA exceptions...`);
        
        for (const exc of sf.exceptions) {
          // Convert scope text to dropdown value
          const scopeValue = scopeTextToValue[exc.scope.toLowerCase()] || 'InScope';
          console.log(`      Setting PA: "${exc.practiceArea}" to scope: ${scopeValue}`);
          
          const setResult = await this.page.evaluate((paName, scopeValue, justification) => {
            // Normalize string for comparison (lowercase, remove extra spaces)
            const normalize = (s: string) => s.toLowerCase().replace(/\s+/g, ' ').trim();
            const targetPA = normalize(paName);
            
            // Find all practice area selector containers
            const containers = document.querySelectorAll('.project-appraisal-scope-form__practice-area-selector');
            
            for (const container of Array.from(containers)) {
              // Find the label with the PA name
              // The first label in the container (not the "Justification" one) is the PA name
              const labels = container.querySelectorAll('label');
              let matchedLabel: string | null = null;
              
              for (const label of Array.from(labels)) {
                const labelText = label.textContent?.trim() || '';
                // Skip "Justification for In-Scope" labels
                if (labelText.toLowerCase().startsWith('justification')) continue;
                
                // Check if this label matches our target PA
                if (normalize(labelText) === targetPA) {
                  matchedLabel = labelText;
                  break;
                }
              }
              
              if (matchedLabel) {
                // Found the PA container, now find and set the scope dropdown
                const select = container.querySelector('select[name^="PracticeAreaInclusionStatus"]') as HTMLSelectElement;
                if (!select) {
                  return { success: false, pa: paName, error: 'Select element not found in container', matchedLabel };
                }
                
                // Set the scope value
                let scopeSet = false;
                for (const option of Array.from(select.options)) {
                  if (option.value === scopeValue) {
                    select.value = option.value;
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    scopeSet = true;
                    break;
                  }
                }
                
                if (!scopeSet) {
                  return { success: false, pa: paName, error: `Scope value "${scopeValue}" not found in options`, matchedLabel };
                }
                
                // Set the justification - only for In-Scope options
                if (justification && (scopeValue === 'InScope' || scopeValue === 'InScopeDefaultOthersToOutOfScope')) {
                  const textarea = container.querySelector('textarea[name^="PracticeAreaJustification"]') as HTMLTextAreaElement;
                  if (textarea) {
                    textarea.value = justification;
                    textarea.dispatchEvent(new Event('input', { bubbles: true }));
                  }
                }
                
                return { success: true, pa: paName, scope: select.options[select.selectedIndex]?.text?.trim() || scopeValue, matchedLabel };
              }
            }
            
            // Debug: List available PA names
            const availablePAs: string[] = [];
            for (const container of Array.from(containers)) {
              const labels = container.querySelectorAll('label');
              for (const label of Array.from(labels)) {
                const text = label.textContent?.trim();
                if (text && !text.toLowerCase().startsWith('justification')) {
                  availablePAs.push(text);
                  break; // Only take the first non-justification label from each container
                }
              }
            }
            
            return { success: false, pa: paName, error: 'PA not found', availablePAs: availablePAs, searchedFor: targetPA };
          }, exc.practiceArea, scopeValue, exc.justification);
          
          if (setResult.success) {
            console.log(`      ✅ ${setResult.pa}: ${setResult.scope}`);
          } else {
            console.log(`      ⚠️  Failed: ${exc.practiceArea}`);
            console.log(`         Error: ${(setResult as any).error}`);
            if ((setResult as any).searchedFor) {
              console.log(`         Searched for (normalized): "${(setResult as any).searchedFor}"`);
            }
            if ((setResult as any).availablePAs) {
              console.log(`         Available PAs on page: ${(setResult as any).availablePAs.slice(0, 8).join(', ')}...`);
            }
          }
        }
        
        // Step 5: Click Save Exceptions button
        console.log(`   Clicking Save Exceptions...`);
        await Promise.all([
          this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 }).catch(() => {}),
          this.page.evaluate(() => {
            // Find the Save Exceptions button in the #Form
            const form = document.querySelector('#Form form');
            if (form) {
              const buttons = form.querySelectorAll('button');
              for (const btn of Array.from(buttons)) {
                if (btn.textContent?.toLowerCase().includes('save')) {
                  btn.click();
                  return true;
                }
              }
            }
            return false;
          })
        ]);
        
        console.log(`   ✅ Saved exceptions for ${sf.name}`);
        await new Promise(resolve => setTimeout(resolve, 2000));
      }
      
      console.log('\n   ✅ Organizational Support Function PA Exceptions page complete');
      
    } catch (error) {
      console.log(`   ⚠️  Error processing org project appraisal scope: ${error}`);
    }
  }

  // ─────────────────────────────────────────────────────────────
  // OE COLLECTION PLAN HANDLERS
  // Pages: /objective-evidence/collection-approach
  //         /objective-evidence/collection-techniques
  //         /objective-evidence/collection-responsibilities
  //         /objective-evidence/performance-report-approaches
  //         /objective-evidence/data-collection-timing
  //         /objective-evidence/additional-info
  // ─────────────────────────────────────────────────────────────

  // Generic helper: fill a simple form (select + optional text + textarea) and save.
  // Used by collection-approach, collection-techniques, collection-responsibilities,
  // performance-report-approaches and additional-info.
  private async fillAndSaveSimpleForm(fields: Array<{ selector: string; type: 'select' | 'text' | 'textarea' | 'date'; value: string }>): Promise<void> {
    if (!this.page) return;
    for (const field of fields) {
      if (!field.value) continue;
      try {
        switch (field.type) {
          case 'select':
            await this.populateSelect(field.selector, field.value);
            break;
          case 'text':
            await this.populateTextInput(field.selector, field.value);
            break;
          case 'textarea':
            await this.populateTextInput(field.selector, field.value);
            break;
          case 'date':
            await this.populateDateInput(field.selector, field.value);
            break;
        }
      } catch (e) {
        console.log(`      ⚠️  Could not set ${field.selector}: ${e}`);
      }
    }
    await this.clickSaveButton();
    await new Promise(resolve => setTimeout(resolve, 1500));
  }

  // /objective-evidence/collection-approach
  // Rows 50-51 of P1PA-R
  // This page always shows the edit form + existing entry card above it.
  // Compare the current entry to the desired values; skip the update if they match.
  async handleOECollectionApproachPage(): Promise<void> {
    if (!this.page) return;
    console.log('\n   📋 Processing OE Collection Approach page...');

    const d = this.excelData['P1PA-R'] || {};
    const approach = d[50];  // e.g. "Managed Discovery"
    const comment  = d[51];  // description

    console.log(`   Collection Approach: ${approach}`);
    console.log(`   Description: ${comment?.substring(0, 60)}...`);

    // Read the existing entry from the .card > .item-card block
    const existing = await this.page.evaluate(() => {
      const title = document.querySelector('.card .item-card .item-card__title h3')?.textContent?.trim() || '';
      const desc  = document.querySelector('.card .item-card .item-card__section small')?.textContent?.trim() || '';
      return { title, desc };
    });

    console.log(`   Existing entry: "${existing.title}"`);

    // Normalize whitespace for comparison (trim + collapse internal spaces)
    const norm = (s: string) => (s || '').trim().replace(/\s+/g, ' ');

    // Compare: if title matches the approach and description is substantially the same, skip
    const titleMatches = norm(existing.title).toLowerCase() === norm(approach).toLowerCase();
    const descMatches  = norm(existing.desc) === norm(comment);

    if (titleMatches && descMatches) {
      console.log('   ℹ️  Entry already matches - skipping update');
      return;
    }

    if (existing.title && titleMatches && !descMatches) {
      console.log(`   ⚠️  Approach matches but description differs - updating`);
    } else if (existing.title && !titleMatches) {
      console.log(`   ⚠️  Approach differs ("${existing.title}" vs "${approach}") - updating`);
    } else if (!existing.title) {
      console.log(`   ℹ️  No existing entry found - populating`);
    }

    await this.fillAndSaveSimpleForm([
      { selector: 'select[name="Type"]', type: 'select',   value: approach },
      { selector: '#Comment',            type: 'textarea', value: comment  },
    ]);

    console.log('   ✅ OE Collection Approach saved');
  }

  // /objective-evidence/collection-techniques
  // Rows 54-57 of P1PA-R
  async handleOECollectionTechniquesPage(): Promise<void> {
    if (!this.page) return;
    console.log('\n   📋 Processing OE Collection Techniques page...');

    const d = this.excelData['P1PA-R'] || {};
    const technique   = d[54];  // e.g. "OE Database"
    const otherTech   = d[55];  // (usually blank)
    const oeType      = d[56];  // e.g. "Artifacts and Affirmations"
    const description = d[57];

    console.log(`   Technique: ${technique}`);
    console.log(`   OE Type:   ${oeType}`);

    // Check if an entry with this technique already exists
    const exists = await this.page.evaluate((techName) => {
      const cards = document.querySelectorAll('.item-card');
      for (const c of Array.from(cards)) {
        if (c.textContent?.includes(techName)) return true;
      }
      return false;
    }, technique || '');

    if (exists) {
      console.log(`   ℹ️  Technique "${technique}" already exists - skipping`);
      return;
    }

    // Click Add if the form is not already visible
    const formVisible = await this.page.$('select[name="Type"]');
    if (!formVisible) {
      await this.page.evaluate(() => {
        const btn = Array.from(document.querySelectorAll('a.button, button'))
          .find(b => b.textContent?.toLowerCase().includes('add'));
        (btn as HTMLElement | undefined)?.click();
      });
      await new Promise(r => setTimeout(r, 1500));
    }

    await this.fillAndSaveSimpleForm([
      { selector: 'select[name="Type"]',                 type: 'select',   value: technique  },
      { selector: '[name="Name"]',                       type: 'text',     value: otherTech  },
      { selector: 'select[name="ObjectiveEvidenceType"]', type: 'select',  value: oeType     },
      { selector: '[name="Description"]',                type: 'textarea', value: description },
    ]);

    console.log('   ✅ OE Collection Technique saved');
  }

  // /objective-evidence/collection-responsibilities
  // Rows 62-64 of P1PA-R
  async handleOECollectionResponsibilitiesPage(): Promise<void> {
    if (!this.page) return;
    console.log('\n   📋 Processing OE Collection Responsibilities page...');

    const d = this.excelData['P1PA-R'] || {};
    const entity      = d[62];  // e.g. "Other"
    const otherEntity = d[63];  // e.g. "EPG member"
    const description = d[64];

    console.log(`   Entity: ${entity}  Other: ${otherEntity}`);

    // Check if an entry already exists
    const exists = await this.page.evaluate(() => {
      return document.querySelector('.item-card') !== null;
    });

    if (exists) {
      console.log('   ℹ️  Responsibility entry already exists - skipping');
      return;
    }

    const formVisible = await this.page.$('select[name="Type"]');
    if (!formVisible) {
      await this.page.evaluate(() => {
        const btn = Array.from(document.querySelectorAll('a.button, button'))
          .find(b => b.textContent?.toLowerCase().includes('add'));
        (btn as HTMLElement | undefined)?.click();
      });
      await new Promise(r => setTimeout(r, 1500));
    }

    await this.fillAndSaveSimpleForm([
      { selector: 'select[name="Type"]',  type: 'select',   value: entity      },
      { selector: '[name="Name"]',         type: 'text',     value: otherEntity },
      { selector: '[name="Description"]',  type: 'textarea', value: description },
    ]);

    console.log('   ✅ OE Collection Responsibility saved');
  }

  // /objective-evidence/performance-report-approaches
  // Rows 67-69 of P1PA-R
  async handlePerformanceReportApproachesPage(): Promise<void> {
    if (!this.page) return;
    console.log('\n   📋 Processing Performance Report Approaches page...');

    const d = this.excelData['P1PA-R'] || {};
    const approach      = d[67];  // e.g. "ATL"
    const otherApproach = d[68];  // e.g. "na" (might be blank/skipped)
    const description   = d[69];

    console.log(`   Approach: ${approach}  Other: ${otherApproach}`);

    const exists = await this.page.evaluate(() => {
      return document.querySelector('.item-card') !== null;
    });

    if (exists) {
      console.log('   ℹ️  Approach entry already exists - skipping');
      return;
    }

    const formVisible = await this.page.$('select[name="Type"]');
    if (!formVisible) {
      await this.page.evaluate(() => {
        const btn = Array.from(document.querySelectorAll('a.button, button'))
          .find(b => b.textContent?.toLowerCase().includes('add'));
        (btn as HTMLElement | undefined)?.click();
      });
      await new Promise(r => setTimeout(r, 1500));
    }

    // Only include OtherApproach if it has a meaningful value
    const fields: Array<{ selector: string; type: 'select' | 'text' | 'textarea' | 'date'; value: string }> = [
      { selector: 'select[name="Type"]',     type: 'select',   value: approach    },
      { selector: '[name="Description"]',     type: 'textarea', value: description },
    ];
    if (otherApproach && otherApproach.toLowerCase() !== 'na' && otherApproach.trim() !== '') {
      fields.splice(1, 0, { selector: '[name="Name"]', type: 'text', value: otherApproach });
    }

    await this.fillAndSaveSimpleForm(fields);

    console.log('   ✅ Performance Report Approach saved');
  }

  // /objective-evidence/initial-summary
  // Row 72 of P1PA-R
  // Single textarea (#Summary). Page always shows existing value in a .card above the form.
  // Skip the update if the current content already matches.
  async handleInitialSummaryPage(): Promise<void> {
    if (!this.page) return;
    console.log('\n   📋 Processing Initial OE Summary page...');

    const d = this.excelData['P1PA-R'] || {};
    const summary = d[72];

    if (!summary) {
      console.log('   ⚠️  No summary value in Excel (row 72) - skipping');
      return;
    }

    console.log(`   Summary: ${summary.substring(0, 80)}...`);

    // Read existing value from the .card above the form
    const existingText = await this.page.evaluate(() => {
      // The existing summary is shown in a .card .item-card section
      const small = document.querySelector('.card .item-card__section small');
      return small?.textContent?.trim() || '';
    });

    const norm = (s: string) => (s || '').trim().replace(/\s+/g, ' ');

    if (norm(existingText) === norm(summary)) {
      console.log('   ℹ️  Summary already matches - skipping update');
      return;
    }

    if (existingText) {
      console.log(`   ⚠️  Summary differs - updating`);
    } else {
      console.log('   ℹ️  No existing summary - populating');
    }

    await this.fillAndSaveSimpleForm([
      { selector: '#Summary', type: 'textarea', value: summary },
    ]);

    console.log('   ✅ Initial OE Summary saved');
  }

  // /objective-evidence/data-collection-timing
  // Milestone 1: P1PA-R rows 75 (name), 76 (date), 77 (participants)
  // Milestone 2: P1PA-R rows 79 (name), 80 (date), 81 (participants)
  //
  // Page structure (from HTML snapshot):
  //   .item-card-list
  //     .card.data-collection-timing-card  (one per existing entry)
  //       .item-card
  //         .item-card__title h3           (milestone name)
  //         .item-card__actions
  //           a.red-button  href="?Id=NNN&handler=ConfirmDelete"  <- navigates to confirm page
  //   #Form  (always present at bottom)
  //     form.data-collection-timing-form
  //       #Name, #CompletedYear/#CompletedMonth/#CompletedDay, #ParticipantNamesListing
  //       button "Update Data Collection Timing"
  //
  // Strategy:
  //   1. Delete ALL existing entries one by one (click red Delete link → confirm → back)
  //   2. For each desired milestone, fill the always-present form and submit
  //   NO page.goto() calls — work entirely on the already-loaded page
  async handleDataCollectionTimingPage(): Promise<void> {
    if (!this.page) return;
    console.log('\n   \ud83d\udccb Processing Data Collection Timing page...');

    const appraisalId = CONFIG.appraisalId;
    const baseUrl     = CONFIG.casBaseUrl;
    const d           = this.excelData['P1PA-R'] || {};

    const milestones = [
      { name: d[75], date: d[76], participants: d[77] },
      { name: d[79], date: d[80], participants: d[81] },
    ].filter(m => m.name);

    console.log(`   Milestones to create: ${milestones.length}`);
    if (milestones.length === 0) {
      console.log('   \u26a0\ufe0f  No milestone data found in Excel - skipping');
      return;
    }

    // ── STEP 1: Delete all existing entries ──────────────────────────────
    console.log('\n   \ud83d\uddd1\ufe0f  Clearing existing entries...');
    let safetyLimit = 10;
    while (safetyLimit-- > 0) {
      // Find the first red Delete link in the item-card-list
      const deleteHref = await this.page.evaluate(() => {
        const link = document.querySelector(
          '.item-card-list a.red-button[href*="ConfirmDelete"], ' +
          '.item-card-list a[href*="handler=ConfirmDelete"]'
        ) as HTMLAnchorElement | null;
        return link ? link.href : null;
      });

      if (!deleteHref) {
        console.log('   \u2705 No more entries to delete');
        break;
      }

      console.log(`   \u2192 Clicking Delete: ${deleteHref}`);
      await this.page.goto(deleteHref, { waitUntil: 'networkidle2' });
      await new Promise(r => setTimeout(r, 1000));

      // The delete link goes to a ConfirmDelete page - look for a confirm/yes button
      const confirmed = await this.page.evaluate(() => {
        // Look for a confirm submit button ("Delete", "Yes", "Confirm", or any red/submit button)
        const btns = Array.from(document.querySelectorAll('button, input[type="submit"]'));
        for (const btn of btns) {
          const text = (btn.textContent || (btn as HTMLInputElement).value || '').toLowerCase();
          if (text.includes('delete') || text.includes('yes') || text.includes('confirm')) {
            (btn as HTMLElement).click();
            return true;
          }
        }
        // Fallback: submit any form on the confirm page
        const form = document.querySelector('form') as HTMLFormElement | null;
        if (form) { form.submit(); return true; }
        return false;
      });

      await this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }).catch(() => {});
      await new Promise(r => setTimeout(r, 1000));

      if (!confirmed) {
        console.log('   \u26a0\ufe0f  Could not confirm delete - navigating back to timing page');
        await this.page.goto(
          `${baseUrl}/appraisals/${appraisalId}/objective-evidence/data-collection-timing`,
          { waitUntil: 'networkidle2' }
        );
        await new Promise(r => setTimeout(r, 1000));
        break;
      }

      console.log('   \u2705 Entry deleted');

      // After confirm the server redirects back to the timing page automatically;
      // if not, navigate back explicitly
      const currentUrl = this.page.url();
      if (!currentUrl.includes('data-collection-timing')) {
        await this.page.goto(
          `${baseUrl}/appraisals/${appraisalId}/objective-evidence/data-collection-timing`,
          { waitUntil: 'networkidle2' }
        );
        await new Promise(r => setTimeout(r, 1000));
      }
    }

    // ── STEP 2: Create each milestone using the always-present form ───────
    for (let i = 0; i < milestones.length; i++) {
      const m = milestones[i];
      console.log(`\n   === Creating Milestone ${i + 1}: ${m.name} ===`);
      console.log(`      Date: ${m.date}  Participants: ${m.participants?.substring(0, 60)}`);

      // The form (#Form) is always at the bottom of the page.
      // If we navigated away during delete-loop we may need to get back.
      const currentUrl = this.page.url();
      if (!currentUrl.includes('data-collection-timing')) {
        await this.page.goto(
          `${baseUrl}/appraisals/${appraisalId}/objective-evidence/data-collection-timing`,
          { waitUntil: 'networkidle2' }
        );
        await new Promise(r => setTimeout(r, 1000));
      }

      // Scroll form into view
      await this.page.evaluate(() => {
        document.getElementById('Form')?.scrollIntoView({ block: 'center' });
      });
      await new Promise(r => setTimeout(r, 500));

      // Wait for the form to be ready
      try {
        await this.page.waitForSelector('form.data-collection-timing-form #Name', { timeout: 8000 });
      } catch (e) {
        console.log(`   \u26a0\ufe0f  Form not found on page`);
        continue;
      }

      // ── Fill: Name ─────────────────────────────────────────────────
      try {
        await this.page.evaluate(() => {
          const el = document.querySelector('form.data-collection-timing-form #Name') as HTMLInputElement | null;
          if (el) { el.value = ''; }
        });
        await this.populateTextInput('form.data-collection-timing-form #Name', m.name);
        console.log(`      \u2705 Name: ${m.name}`);
      } catch (e) {
        console.log(`      \u26a0\ufe0f  Name error: ${e}`);
      }

      // ── Fill: Date (CompletedYear/Month/Day) ────────────────────────
      if (m.date) {
        try {
          await this.populateDateParts(
            'form.data-collection-timing-form #CompletedYear,' +
            'form.data-collection-timing-form #CompletedMonth,' +
            'form.data-collection-timing-form #CompletedDay',
            m.date
          );
          console.log(`      \u2705 Date: ${m.date}`);
        } catch (e) {
          console.log(`      \u26a0\ufe0f  Date error: ${e}`);
        }
      }

      // ── Fill: Participants ──────────────────────────────────────────
      if (m.participants) {
        try {
          await this.populateTextInput(
            'form.data-collection-timing-form #ParticipantNamesListing',
            m.participants
          );
          console.log(`      \u2705 Participants: ${m.participants.substring(0, 60)}`);
        } catch (e) {
          console.log(`      \u26a0\ufe0f  Participants error: ${e}`);
        }
      }

      // ── Submit the form ─────────────────────────────────────────────
      // Click the submit button inside the data-collection-timing-form specifically
      const submitted = await this.page.evaluate(() => {
        const form = document.querySelector('form.data-collection-timing-form') as HTMLFormElement | null;
        if (!form) return false;
        const btn = form.querySelector('button') as HTMLButtonElement | null;
        if (btn) { btn.click(); return true; }
        form.submit();
        return true;
      });

      if (!submitted) {
        console.log(`      \u26a0\ufe0f  Could not submit form`);
        continue;
      }

      await this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }).catch(() => {});
      await new Promise(r => setTimeout(r, 1500));
      console.log(`   \u2705 Milestone ${i + 1} saved`);

      // After submit the server redirects back to the timing page;
      // verify we're still on it before next iteration
      const urlAfter = this.page.url();
      if (!urlAfter.includes('data-collection-timing')) {
        console.log(`   \u26a0\ufe0f  Unexpected redirect to: ${urlAfter} - navigating back`);
        await this.page.goto(
          `${baseUrl}/appraisals/${appraisalId}/objective-evidence/data-collection-timing`,
          { waitUntil: 'networkidle2' }
        );
        await new Promise(r => setTimeout(r, 1000));
      }
    }

    console.log('\n   \u2705 Data Collection Timing complete');
  }

  // /objective-evidence/additional-info
  // Row 86 of P1PA-R
  // /objective-evidence/initial-summary
  // Row 72 of P1PA-R
  // Single textarea (#Summary). Always-visible form with existing value shown above it.
  // Compare current text; skip if already matches.
  async handleOEInitialSummaryPage(): Promise<void> {
    if (!this.page) return;
    console.log('\n   📋 Processing OE Initial Summary page...');

    const d = this.excelData['P1PA-R'] || {};
    const summary = d[72];

    if (!summary) {
      console.log('   ⚠️  No summary value in Excel data (row 72) - skipping');
      return;
    }

    console.log(`   Summary: ${summary.substring(0, 80)}...`);

    // Read existing value from the textarea (the form is always visible)
    const existing = await this.page.evaluate(() => {
      const ta = document.querySelector('#Summary') as HTMLTextAreaElement | null;
      return ta ? ta.value.trim() : '';
    });

    const norm = (s: string) => (s || '').trim().replace(/\s+/g, ' ');

    if (norm(existing) === norm(summary)) {
      console.log('   ℹ️  Summary already matches - skipping update');
      return;
    }

    if (existing) {
      console.log(`   ⚠️  Existing summary differs - updating`);
    } else {
      console.log(`   ℹ️  No existing summary - populating`);
    }

    await this.fillAndSaveSimpleForm([
      { selector: '#Summary', type: 'textarea', value: summary },
    ]);

    console.log('   ✅ OE Initial Summary saved');
  }

  async handleOEAdditionalInfoPage(): Promise<void> {
    if (!this.page) return;
    console.log('\n   📋 Processing OE Additional Info page...');

    const d    = this.excelData['P1PA-R'] || {};
    const info = d[86];

    console.log(`   Additional Info: ${info?.substring(0, 60)}`);

    if (!info) {
      console.log('   ℹ️  No additional info to populate');
      return;
    }

    await this.fillAndSaveSimpleForm([
      { selector: '[name="AdditionalInfo"]', type: 'textarea', value: info },
    ]);

    console.log('   ✅ OE Additional Info saved');
  }

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

  async populateField(mapping: FieldMapping, value: string): Promise<{ success: boolean; changed: boolean; error?: string }> {
    if (!this.page || !value) {
      return { success: false, changed: false, error: 'No page or empty value' };
    }

    const { CAS_Selector, CAS_Type, FieldLabel } = mapping;
    
    console.log(`\n   📝 Field: ${FieldLabel}`);
    console.log(`      Selector: ${CAS_Selector}`);
    console.log(`      Type: ${CAS_Type}`);
    console.log(`      Value: ${value.substring(0, 50)}${value.length > 50 ? '...' : ''}`);

    try {
      let result: { changed: boolean } = { changed: false };
      
      switch (CAS_Type) {
        case 'text':
        case 'textarea':
          result = await this.populateTextInput(CAS_Selector, value);
          break;
          
        case 'select':
          result = await this.populateSelect(CAS_Selector, value);
          break;
          
        case 'radio':
          result = await this.populateRadio(CAS_Selector, value, mapping.Notes);
          break;
          
        case 'checkbox':
          result = await this.populateCheckbox(CAS_Selector, value);
          break;
        
        case 'number':
          result = await this.populateNumberInput(CAS_Selector, value);
          break;
          
        case 'multiselect':
          result = await this.populateMultiselect(CAS_Selector, value);
          break;
        
        case 'radio-level':
          result = await this.populateRadioLevel(CAS_Selector, value, mapping.Notes);
          break;
        
        case 'date':
          result = await this.populateDateInput(CAS_Selector, value);
          break;
        
        case 'date-parts':
          result = await this.populateDateParts(CAS_Selector, value);
          break;
          
        case 'skip':
          console.log(`      ⏭️  Skipped (field not present in CAS)`);
          return { success: true, changed: false };
          
        default:
          console.log(`      ⚠️  Unknown field type: ${CAS_Type}`);
          return { success: false, changed: false, error: `Unknown type: ${CAS_Type}` };
      }
      
      if (result.changed) {
        console.log(`      ✅ Field updated`);
      }
      return { success: true, changed: result.changed };
      
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      console.log(`      ❌ Failed: ${errorMsg}`);
      return { success: false, changed: false, error: errorMsg };
    }
  }

  async populateTextInput(selector: string, value: string): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };
    
    await this.page.waitForSelector(selector, { timeout: 5000 });
    
    // Check if value already matches
    const currentValue = await this.page.evaluate((sel) => {
      const el = document.querySelector(sel) as HTMLInputElement | HTMLTextAreaElement;
      return el ? el.value : '';
    }, selector);
    
    if (currentValue.trim() === value.trim()) {
      console.log(`      ℹ️  Value already set, skipping`);
      return { changed: false };
    }
    
    await this.page.click(selector);
    
    // Clear existing value
    await this.page.evaluate((sel) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      if (el) el.value = '';
    }, selector);
    
    // Type new value
    await this.page.type(selector, value, { delay: 10 });
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));
    
    return { changed: true };
  }

  async populateSelect(selector: string, value: string): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };
    
    await this.page.waitForSelector(selector, { timeout: 5000 });
    
    // Check current selection and find the target option
    const result = await this.page.evaluate((sel, searchText) => {
      const select = document.querySelector(sel) as HTMLSelectElement;
      if (!select) return { currentText: '', targetValue: null };
      
      const currentText = select.options[select.selectedIndex]?.text || '';
      const searchLower = searchText.toLowerCase();
      
      // Pass 1: exact match
      for (const option of Array.from(select.options)) {
        if (option.text === searchText || option.value === searchText) {
          return { currentText, targetValue: option.value, targetText: option.text };
        }
      }
      // Pass 2: option text contains search text
      for (const option of Array.from(select.options)) {
        if (option.text.toLowerCase().includes(searchLower)) {
          return { currentText, targetValue: option.value, targetText: option.text };
        }
      }
      // Pass 3: search text contains option text (e.g. "Demix (Pty) Ltd" contains "Demix")
      for (const option of Array.from(select.options)) {
        const optLower = option.text.toLowerCase();
        if (optLower.length > 3 && searchLower.includes(optLower)) {
          return { currentText, targetValue: option.value, targetText: option.text };
        }
      }
      return { currentText, targetValue: null };
    }, selector, value);
    
    if (!result.targetValue) {
      throw new Error(`Option not found for: ${value}`);
    }
    
    // Check if already selected
    if (result.currentText.includes(value) || result.currentText === result.targetText) {
      console.log(`      ℹ️  Value already selected, skipping`);
      return { changed: false };
    }
    
    await this.page.select(selector, result.targetValue);
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));
    
    return { changed: true };
  }

  async populateRadio(selector: string, value: string, notes: string): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };
    
    // Parse radio options from notes.
    // Notes format uses '; ' as separator: "Yes=#yes-id; No=#no-id"
    const options = notes.split(';').map(s => s.trim()).filter(s => s.includes('='));
    let targetSelector = '';
    
    for (const opt of options) {
      const eqIdx = opt.indexOf('=');
      const optValue = opt.substring(0, eqIdx).trim();
      const optSelector = opt.substring(eqIdx + 1).trim();
      if (value.toLowerCase() === optValue.toLowerCase()) {
        targetSelector = optSelector;
        break;
      }
    }
    
    if (!targetSelector) {
      // Fallback: try first or second selector based on Yes/No
      const selectors = selector.split('|');
      targetSelector = value.toLowerCase() === 'yes' ? selectors[0] : selectors[1];
    }
    
    await this.page.waitForSelector(targetSelector, { timeout: 5000 });
    
    // Check if already selected
    const isAlreadyChecked = await this.page.evaluate((sel) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      return el ? el.checked : false;
    }, targetSelector);
    
    if (isAlreadyChecked) {
      console.log(`      ℹ️  Radio already selected, skipping`);
      return { changed: false };
    }
    
    await this.page.click(targetSelector);
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));
    
    return { changed: true };
  }

  async populateCheckbox(selector: string, value: string): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };
    
    const shouldBeChecked = ['yes', 'true', '1', 'x', 'checked'].includes(value.toLowerCase());
    
    // Handle numeric ID selectors (e.g., #5, #6, #7)
    // These need special handling because CSS selectors can't start with a digit
    let elementId: string | null = null;
    if (selector.match(/^#\d+$/)) {
      elementId = selector.substring(1); // Extract the numeric ID
      console.log(`      Using getElementById for numeric ID: ${elementId}`);
    }
    
    // Use JavaScript to find and click the checkbox
    const result = await this.page.evaluate((sel, elemId, shouldCheck) => {
      let el: HTMLInputElement | null = null;
      
      if (elemId) {
        // Use getElementById for numeric IDs
        el = document.getElementById(elemId) as HTMLInputElement;
      } else {
        // Use querySelector for regular selectors
        el = document.querySelector(sel) as HTMLInputElement;
      }
      
      if (!el) {
        // Fallback: Try data-test attribute for CAS-specific selectors
        const dataTestSelector = `input[data-test="input-virtual-selection_${elemId || sel.replace('#', '')}"]`;
        el = document.querySelector(dataTestSelector) as HTMLInputElement;
      }
      
      if (!el) {
        return { success: false, error: `Element not found: ${sel} (id: ${elemId})`, wasChecked: false, nowChecked: false, changed: false };
      }
      
      const isChecked = el.checked;
      
      // Check if already in desired state
      if (shouldCheck === isChecked) {
        return { success: true, wasChecked: isChecked, nowChecked: isChecked, changed: false };
      }
      
      el.click();  // Use element's click method
      // Also dispatch change event
      el.dispatchEvent(new Event('change', { bubbles: true }));
      
      return { success: true, wasChecked: isChecked, nowChecked: el.checked, changed: true };
    }, selector, elementId, shouldBeChecked);
    
    if (!result.success) {
      throw new Error(result.error || 'Unknown error');
    }
    
    if (!result.changed) {
      console.log(`      ℹ️  Checkbox already ${result.wasChecked ? 'checked' : 'unchecked'}, skipping`);
    } else {
      console.log(`      Checkbox state: was=${result.wasChecked}, now=${result.nowChecked}`);
    }
    
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));
    
    return { changed: result.changed };
  }

  async populateNumberInput(selector: string, value: string): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };
    
    await this.page.waitForSelector(selector, { timeout: 5000 });
    
    // Check if value already matches
    const currentValue = await this.page.evaluate((sel) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      return el ? el.value : '';
    }, selector);
    
    if (currentValue.trim() === value.trim()) {
      console.log(`      ℹ️  Value already set, skipping`);
      return { changed: false };
    }
    
    // Clear and set the value
    await this.page.evaluate((sel, val) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      if (el) {
        el.value = '';
        el.value = val;
        // Trigger input and change events to ensure React picks it up
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }
    }, selector, value);
    
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));
    
    return { changed: true };
  }

  async populateDateParts(selector: string, value: string): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };
    
    console.log(`      populateDateParts called with selector="${selector}", value="${value}"`);
    
    // selector format: "#StartDateYear,#StartDateMonth,#StartDateDay"
    const [yearSelector, monthSelector, daySelector] = selector.split(',').map(s => s.trim());
    console.log(`      Parsed selectors: year="${yearSelector}", month="${monthSelector}", day="${daySelector}"`);
    
    // Parse date - expected format YYYY-MM-DD
    let year: string, month: string, day: string;
    if (value.match(/^\d{4}-\d{2}-\d{2}/)) {
      const parts = value.split('-');
      year = parts[0];
      month = parts[1];
      day = parts[2];
    } else {
      console.log(`      ERROR: Invalid date format: ${value}`);
      throw new Error(`Invalid date format: ${value}. Expected YYYY-MM-DD`);
    }
    
    // Remove leading zeros for month and day (CAS expects plain numbers)
    const monthValue = parseInt(month, 10).toString();
    const dayValue = parseInt(day, 10).toString();
    
    console.log(`      Date parts: Year=${year}, Month=${monthValue}, Day=${dayValue}`);
    
    let changed = false;
    
    // Check if element exists first
    const elementExists = await this.page.evaluate((sel) => {
      const el = document.querySelector(sel);
      return el !== null;
    }, yearSelector);
    
    if (!elementExists) {
      console.log(`      ERROR: Element not found: ${yearSelector}`);
      throw new Error(`Element not found: ${yearSelector}`);
    }
    
    // Determine if we're dealing with number inputs or select dropdowns
    // by checking the element type
    const elementInfo = await this.page.evaluate((sel) => {
      const el = document.querySelector(sel);
      if (!el) return { exists: false, tagName: '', type: '' };
      return { 
        exists: true, 
        tagName: el.tagName, 
        type: (el as HTMLInputElement).type || '',
        id: el.id,
        name: (el as HTMLInputElement).name || ''
      };
    }, yearSelector);
    
    console.log(`      Element info: ${JSON.stringify(elementInfo)}`);
    
    const isNumberInput = elementInfo.tagName === 'INPUT' && elementInfo.type === 'number';
    
    if (isNumberInput) {
      // Handle number inputs (Timeline phase dates)
      console.log(`      Using number input mode`);
      
      // Set Year
      if (yearSelector) {
        console.log(`      Setting year to: ${year}`);
        await this.page.waitForSelector(yearSelector, { timeout: 5000 });
        const yearResult = await this.page.evaluate((sel, val) => {
          const el = document.querySelector(sel) as HTMLInputElement;
          if (el) {
            el.value = val;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            return { success: true, newValue: el.value };
          }
          return { success: false, newValue: '' };
        }, yearSelector, year);
        console.log(`      Year result: ${JSON.stringify(yearResult)}`);
        changed = true;
      }
      
      // Set Month
      if (monthSelector) {
        console.log(`      Setting month to: ${monthValue}`);
        await this.page.waitForSelector(monthSelector, { timeout: 5000 });
        const monthResult = await this.page.evaluate((sel, val) => {
          const el = document.querySelector(sel) as HTMLInputElement;
          if (el) {
            el.value = val;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            return { success: true, newValue: el.value };
          }
          return { success: false, newValue: '' };
        }, monthSelector, monthValue);
        console.log(`      Month result: ${JSON.stringify(monthResult)}`);
        changed = true;
      }
      
      // Set Day
      if (daySelector) {
        console.log(`      Setting day to: ${dayValue}`);
        await this.page.waitForSelector(daySelector, { timeout: 5000 });
        const dayResult = await this.page.evaluate((sel, val) => {
          const el = document.querySelector(sel) as HTMLInputElement;
          if (el) {
            el.value = val;
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            return { success: true, newValue: el.value };
          }
          return { success: false, newValue: '' };
        }, daySelector, dayValue);
        console.log(`      Day result: ${JSON.stringify(dayResult)}`);
        changed = true;
      }
    } else {
      // Handle select dropdowns (Readiness reviews)
      console.log(`      Using select dropdown mode`);
      
      // Set Year (select dropdown)
      if (yearSelector) {
        await this.page.waitForSelector(yearSelector, { timeout: 5000 });
        await this.page.select(yearSelector, year);
        changed = true;
      }
      
      // Set Month (select dropdown)
      if (monthSelector) {
        await this.page.waitForSelector(monthSelector, { timeout: 5000 });
        await this.page.select(monthSelector, monthValue);
        changed = true;
      }
      
      // Set Day (select dropdown)
      if (daySelector) {
        await this.page.waitForSelector(daySelector, { timeout: 5000 });
        await this.page.select(daySelector, dayValue);
        changed = true;
      }
    }
    
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));
    
    console.log(`      populateDateParts complete, changed=${changed}`);
    return { changed };
  }

  async populateDateInput(selector: string, value: string): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };
    
    await this.page.waitForSelector(selector, { timeout: 5000 });
    
    // Format date to MM/DD/YYYY for CAS date inputs
    let formattedDate = value;
    if (value.match(/^\d{4}-\d{2}-\d{2}/)) {
      // Convert from YYYY-MM-DD to MM/DD/YYYY
      const parts = value.split('-');
      formattedDate = `${parts[1]}/${parts[2]}/${parts[0]}`;
    }
    
    // Check if value already matches
    const currentValue = await this.page.evaluate((sel) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      return el ? el.value : '';
    }, selector);
    
    if (currentValue === formattedDate) {
      console.log(`      \u2139\ufe0f  Date already set, skipping`);
      return { changed: false };
    }
    
    // Clear and set the date value
    await this.page.evaluate((sel, val) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      if (el) {
        el.value = '';
        el.value = val;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }
    }, selector, formattedDate);
    
    console.log(`      Set date to: ${formattedDate}`);
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));
    
    return { changed: true };
  }

  async populateRadioLevel(selector: string, value: string, notes: string): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };
    
    // Normalise the value: strip leading "Level " prefix (case-insensitive) so that
    // Excel values like "Level 5" match notes keys like "5".
    // e.g. "Level 5" -> "5",  "3" -> "3"
    const normValue = value.trim().replace(/^level\s+/i, '');
    
    // Parse the notes to find the selector for the given level value
    // Notes format: "1=#level-1, 2=#level-2, 3=#level-3, 4=#level-4, 5=#level-5"
    const options = notes.split(',').map(s => s.trim());
    let targetSelector = '';
    
    for (const opt of options) {
      const [optValue, optSelector] = opt.split('=').map(s => s.trim());
      if (normValue === optValue) {
        targetSelector = optSelector;
        break;
      }
    }
    
    if (!targetSelector) {
      // Fallback: construct selector from normValue (e.g., "5" -> "#level-5")
      targetSelector = `#level-${normValue}`;
    }
    
    console.log(`      Target level selector: ${targetSelector}`);
    
    await this.page.waitForSelector(targetSelector, { timeout: 5000 });
    
    // Check if already selected
    const isAlreadyChecked = await this.page.evaluate((sel) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      return el ? el.checked : false;
    }, targetSelector);
    
    if (isAlreadyChecked) {
      console.log(`      \u2139\ufe0f  Level already selected, skipping`);
      return { changed: false };
    }
    
    await this.page.click(targetSelector);
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));
    
    return { changed: true };
  }

  async populateMultiselect(selector: string, value: string): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };
    
    console.log(`      🔄 Handling React multiselect...`);
    
    try {
      // The multiselect is a React Select component
      // We need to: 1) click to open dropdown, 2) click the option
      
      // Find the multiselect container
      const containerSelector = selector || '.multiselect';
      await this.page.waitForSelector(containerSelector, { timeout: 5000 });
      
      // Check if any value is already selected (try multiple React Select class patterns)
      const alreadySelected = await this.page.evaluate((searchText) => {
        // React Select v2+/v3+ class patterns
        const selectors = [
          '.multiselect__multi-value__label',
          '[class*="multiValue"] [class*="label"]',
          '[class*="multi-value"] [class*="label"]',
          '.css-1rhbuit-multiValue .css-1v99tuv',  // React Select generated classes
          '[class*="ValueContainer"] [class*="MultiValue"]'
        ];
        for (const sel of selectors) {
          const vals = document.querySelectorAll(sel);
          for (const val of Array.from(vals)) {
            if (val.textContent?.includes(searchText)) return true;
          }
        }
        return false;
      }, value);
      
      if (alreadySelected) {
        console.log(`      ℹ️  Value already selected, skipping`);
        return { changed: false };
      }
      
      // Click to open the dropdown
      await this.page.click(containerSelector);
      console.log(`      Clicked to open dropdown`);
      await new Promise(resolve => setTimeout(resolve, 800));
      
      // Wait for dropdown menu to appear (try multiple class patterns)
      const menuSelectors = [
        '.multiselect__menu',
        '[class*="MenuList"]',
        '[class*="menu-list"]',
        '[class*="menuList"]',
        '[id*="listbox"]'
      ];
      
      let menuFound = false;
      for (const menuSel of menuSelectors) {
        try {
          await this.page.waitForSelector(menuSel, { timeout: 2000 });
          menuFound = true;
          console.log(`      Dropdown menu opened (${menuSel})`);
          break;
        } catch { /* try next */ }
      }
      
      if (!menuFound) {
        console.log(`      ⚠️  No dropdown menu detected, trying to find options anyway`);
      }
      
      // Dump available options for debugging, then select first item
      const clicked = await this.page.evaluate((searchText) => {
        // Try multiple option selector patterns
        const optionSelectors = [
          '.multiselect__option',
          '[class*="option"]',
          '[class*="Option"]',
          '[id*="option"]',
          '[role="option"]'
        ];
        
        let options: Element[] = [];
        let usedSelector = '';
        for (const sel of optionSelectors) {
          const found = document.querySelectorAll(sel);
          if (found.length > 0) {
            options = Array.from(found);
            usedSelector = sel;
            break;
          }
        }
        
        if (!options.length) {
          return { clicked: false, debug: 'No options found with any selector pattern' };
        }
        
        // Log available options
        const availableOptions = options.map(o => (o.textContent || '').trim()).slice(0, 10);
        
        // Pass 1: exact text match
        for (const option of options) {
          const text = (option.textContent || '').trim();
          if (text === searchText) {
            (option as HTMLElement).click();
            return { clicked: true, text, method: 'exact match', selector: usedSelector, availableOptions };
          }
        }
        
        // Pass 2: partial match
        for (const option of options) {
          const text = (option.textContent || '').trim();
          if (text.includes(searchText) || searchText.includes(text)) {
            (option as HTMLElement).click();
            return { clicked: true, text, method: 'partial match', selector: usedSelector, availableOptions };
          }
        }
        
        // Pass 3: select first option
        const first = options[0] as HTMLElement;
        const firstText = (first.textContent || '').trim();
        first.click();
        return { clicked: true, text: firstText, method: 'first option (no text match)', selector: usedSelector, availableOptions };
      }, value);
      
      if (clicked.clicked) {
        console.log(`      ✅ Selected: "${clicked.text?.substring(0, 60)}" (${clicked.method})`);
        if (clicked.method?.includes('first option')) {
          console.log(`      ℹ️  Available options were: ${JSON.stringify(clicked.availableOptions)}`);
        }
        await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));
        return { changed: true };
      } else {
        console.log(`      ⚠️  Debug: ${clicked.debug}`);
        // Don't throw - treat as non-fatal
        console.log(`      ⚠️  Could not select multiselect option, continuing`);
        return { changed: false };
      }
      
    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      console.log(`      ❌ Multiselect error: ${errorMsg}`);
      throw error;
    }
  }

  // Click the Next button to go to the next page in the CAS workflow
  async clickNextButton(): Promise<boolean> {
    if (!this.page) return false;
    
    console.log('\n   ➡️  Looking for Next button...');
    
    try {
      // Find the Next button - it's a blue-button with an SVG arrow inside
      // The arrow points right (polygon starts with "80,60")
      const result = await this.page.evaluate(() => {
        // Find all blue-button links
        const buttons = Array.from(document.querySelectorAll('a.button.blue-button'));
        
        for (const btn of buttons) {
          // Must contain an SVG with a polygon (the arrow)
          const svg = btn.querySelector('svg');
          const polygon = btn.querySelector('svg polygon');
          
          if (svg && polygon) {
            const points = polygon.getAttribute('points') || '';
            // Forward arrow has points starting with "80,60" (pointing right)
            // Back arrow would start differently
            if (points.startsWith('80,60')) {
              const href = btn.getAttribute('href');
              if (href) {
                (btn as HTMLElement).click();
                return { success: true, href: href };
              }
            }
          }
        }
        
        return { success: false, href: null };
      });
      
      if (result.success && result.href) {
        console.log(`      Found Next button: ${result.href}`);
        console.log('      ✅ Next button clicked');
        
        // Wait for navigation
        await new Promise(resolve => setTimeout(resolve, 2000));
        try {
          await this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 10000 });
          console.log('      ✅ Navigated to next page');
        } catch (e) {
          // Navigation might have already completed
        }
        
        return true;
      }
      
    } catch (e) {
      console.log(`      ⚠️  Error finding Next button: ${e}`);
    }
    
    console.log('      ℹ️  No Next button found');
    return false;
  }

  // Click save/update button to persist changes
  async clickSaveButton(): Promise<boolean> {
    if (!this.page) return false;
    
    console.log('\n   💾 Looking for save/update button...');
    
    // Try different button selectors - CAS-specific first, then generic
    const saveButtonSelectors = [
      'button[data-test="button-update-appraisal"]',  // CAS "Update Appraisal" button
      'button[data-test="button-add-edit-org"]',      // CAS "Add Organization" or "Update Organization" button
      'button[data-test*="update"]',                  // Any CAS update button
      'button[data-test*="save"]',                    // Any CAS save button
      'button[data-test*="add"]',                     // Any CAS add button
      '.actions button:not(.p2):not(.red-button)',    // Action buttons (not cancel/remove)
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
          
          // Click the button
          await button.click();
          console.log('      ✅ Button clicked');
          
          // Wait for any AJAX/navigation to complete
          await new Promise(resolve => setTimeout(resolve, 2000));
          
          // Try to wait for navigation, but don't fail if there's no navigation
          // (some buttons use AJAX instead of page reload)
          try {
            await this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 5000 });
            console.log('      ✅ Page navigation completed');
          } catch (navError) {
            // No navigation occurred - likely an AJAX update, which is fine
            console.log('      ℹ️  No page navigation (AJAX update)');
          }
          
          // Additional wait to ensure any animations/updates complete
          await new Promise(resolve => setTimeout(resolve, 1000));
          
          return true;
        }
      } catch (e) {
        // Try next selector
      }
    }
    
    // Fallback: try to find button by text content
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
    } catch (e) {
      // Fallback failed
    }
    
    console.log('      ⚠️  No save/update button found');
    return false;
  }

  // Click Add/Edit button for repeating sections (like Add Organization)
  async clickAddEditButton(selector: string, buttonName: string): Promise<boolean> {
    if (!this.page) return false;
    
    console.log(`\n   ➕ Looking for "${buttonName}" button...`);
    
    try {
      // Wait a moment for page to stabilize
      await new Promise(resolve => setTimeout(resolve, 1000));
      
      // First try the specific selector
      if (selector) {
        try {
          await this.page.waitForSelector(selector, { timeout: 5000 });
          const button = await this.page.$(selector);
          if (button) {
            // Scroll button into view
            await this.page.evaluate((sel) => {
              const el = document.querySelector(sel);
              if (el) el.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }, selector);
            await new Promise(resolve => setTimeout(resolve, 500));
            
            await button.click();
            console.log(`      ✅ Clicked "${buttonName}" (${selector})`);
            
            // Wait for form/modal to appear
            await new Promise(resolve => setTimeout(resolve, 2000));
            return true;
          }
        } catch (e) {
          console.log(`      ℹ️  Selector ${selector} not found, trying text search...`);
        }
      }
      
      // Fallback: Find button by text content
      const clicked = await this.page.evaluate((text) => {
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
        // Wait for form/modal to appear
        await new Promise(resolve => setTimeout(resolve, 2000));
        return true;
      }
    } catch (e) {
      console.log(`      ⚠️  Error clicking button: ${e}`);
    }
    
    console.log(`      ⚠️  "${buttonName}" button not found`);
    return false;
  }

  // Legacy method for backward compatibility
  async clickAddButton(buttonText: string): Promise<boolean> {
    return this.clickAddEditButton('', buttonText);
  }

  async askForFeedback(): Promise<{ continue: boolean; feedback: string }> {
    // If debug mode is disabled, auto-continue without prompting
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
      
      this.rl.question('Your input: ', (answer) => {
        const input = answer.trim();
        const inputLower = input.toLowerCase();
        
        if (inputLower === 's' || inputLower === 'q' || inputLower === 'stop' || inputLower === 'quit') {
          resolve({ continue: false, feedback: '' });
        } else if (input === '' || inputLower === 'c' || inputLower === 'continue') {
          resolve({ continue: true, feedback: '' });
        } else {
          // Check if feedback ends with [s] or [q] to stop after saving
          const endsWithStop = /\[s\]\s*$/i.test(input) || /\[q\]\s*$/i.test(input);
          
          // Remove the trailing [s] or [q] from the feedback
          let feedback = input;
          if (endsWithStop) {
            feedback = input.replace(/\s*\[[sq]\]\s*$/i, '').trim();
          }
          
          console.log(`💬 Feedback saved: "${feedback}"`);
          
          if (endsWithStop) {
            console.log('🛑 Stop requested after feedback.');
            resolve({ continue: false, feedback: feedback });
          } else {
            resolve({ continue: true, feedback: feedback });
          }
        }
      });
    });
  }
  
  // Handle Phase 1 completion
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
      // Prompt user to confirm exit
      return new Promise((resolve) => {
        this.rl.question('\nPress any key to stop... ', () => {
          resolve();
        });
      });
    }
  }

  // Process the current page (already navigated to it)
  async processCurrentPage(pagePath: string): Promise<PopulateLog> {
    const pageFields = this.fieldMap.filter(f => 
      f.CAS_Page === pagePath && 
      f.CAS_Type !== 'skip' && 
      f.CAS_Selector !== 'SKIP' && 
      f.CAS_Selector !== 'HANDLER'
    );
    
    const log: PopulateLog = {
      timestamp: new Date().toISOString(),
      page: pagePath,
      fieldsAttempted: pageFields.length,
      fieldsSuccessful: 0,
      fieldsFailed: 0,
      details: []
    };

    console.log('\n' + '─'.repeat(60));
    console.log(`📋 Processing page: ${pagePath}`);
    console.log(`   Fields to populate: ${pageFields.length}`);
    console.log('─'.repeat(60));
    
    // Save HTML before processing (for debugging)
    await this.savePageHtml(pagePath, 'before');

    // No special handling needed for /name-and-type
    // Virtual checkboxes (rows 9-11) are handled by the standard checkbox populator

    // Special handling for /organizations page - check if org exists and needs editing
    if (pagePath === '/organizations') {
      await this.handleOrganizationsPage();
    }
    
    // Special handling for /org-units page - check if OU exists and needs editing
    if (pagePath === '/org-units') {
      await this.handleOrgUnitsPage();
    }
    
    // Special handling for /org-unit-targets page - check if target exists and needs editing
    if (pagePath === '/org-unit-targets') {
      await this.handleOrgUnitTargetsPage();
    }
    
    // Special handling for /timeline page - has multiple sub-forms
    if (pagePath === '/timeline') {
      await this.handleTimelinePage();
      return log; // Timeline page has its own processing
    }
    
    // Special handling for /readiness-reviews page - repeating sections
    if (pagePath === '/readiness-reviews') {
      await this.handleReadinessReviewsPage();
      return log; // Readiness reviews page has its own processing
    }
    
    // Special handling for /org-projects pages - template download/upload workflow
    // /org-projects/true = Organizational Support Functions (uses C_SupportV2)
    // /org-projects/false = OU Sample Eligible Projects (uses C_ProjectsV2)
    if (pagePath.startsWith('/org-projects')) {
      if (pagePath.includes('/true') || pagePath.includes('IsOrganizational=True')) {
        // Organizational Support Functions
        await this.handleOrgProjectsPage();
      } else {
        // OU Projects (false or no suffix)
        await this.handleOUProjectsPage();
      }
      return log;
    }
    
    // Special handling for sampling-related pages
    // Data from P1-OrgScope rows 127-139
    if (pagePath.includes('/org-unit-sampling-factors') && !pagePath.includes('values')) {
      await this.handleSamplingFactorsPage();
      return log;
    }
    
    if (pagePath.includes('/org-unit-sampling-factor-values')) {
      await this.handleSamplingFactorValuesPage();
      return log;
    }
    
    if (pagePath.includes('/org-unit-subgroups') && !pagePath.includes('assignment')) {
      await this.handleSubgroupsPage();
      return log;
    }
    
    // Special handling for project subgroup assignment page
    // This assigns all projects to their subgroups
    if (pagePath.includes('/org-unit-project-subgroups')) {
      await this.handleSubgroupAssignmentPage();
      return log;
    }
    
    // Special handling for organizational support function PA exceptions page
    if (pagePath.includes('/organizational-project-appraisal-scope')) {
      await this.handleOrgProjectAppraisalScopePage();
      return log;
    }

    // ── OE Collection Plan pages ──────────────────────────────────
    if (pagePath === '/objective-evidence/collection-approach') {
      await this.handleOECollectionApproachPage();
      return log;
    }

    if (pagePath === '/objective-evidence/initial-summary') {
      await this.handleOEInitialSummaryPage();
      return log;
    }

    if (pagePath === '/objective-evidence/collection-techniques') {
      await this.handleOECollectionTechniquesPage();
      return log;
    }

    if (pagePath === '/objective-evidence/collection-responsibilities') {
      await this.handleOECollectionResponsibilitiesPage();
      return log;
    }

    if (pagePath === '/objective-evidence/performance-report-approaches') {
      await this.handlePerformanceReportApproachesPage();
      return log;
    }

    if (pagePath === '/objective-evidence/data-collection-timing') {
      await this.handleDataCollectionTimingPage();
      return log;
    }

    if (pagePath === '/objective-evidence/additional-info') {
      await this.handleOEAdditionalInfoPage();
      return log;
    }

    // Track if any fields were actually changed
    let anyFieldsChanged = false;

    // Process each field
    for (const field of pageFields) {
      const value = this.excelData[field.Sheet]?.[field.Row];
      
      if (!value) {
        console.log(`\n   ⏭️  Skipping ${field.FieldLabel} (no value in Excel data)`);
        log.details.push({
          field: field.FieldLabel,
          selector: field.CAS_Selector,
          value: '',
          status: 'skipped'
        });
        continue;
      }

      const result = await this.populateField(field, value);
      
      log.details.push({
        field: field.FieldLabel,
        selector: field.CAS_Selector,
        value: value,
        status: result.success ? 'success' : 'failed',
        error: result.error
      });

      if (result.success) {
        log.fieldsSuccessful++;
        if (result.changed) {
          anyFieldsChanged = true;
        }
      } else {
        log.fieldsFailed++;
      }
    }

    // Click save/update button only if we made changes
    if (anyFieldsChanged) {
      console.log('\n   💾 Changes detected, saving...');
      // For org-unit-targets in Add mode, the button is "Add Target" - try it first
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
    
    // Save HTML after processing (for debugging)
    await this.savePageHtml(pagePath, 'after');

    return log;
  }

  // Process page by navigating to it first (legacy method)
  async processPage(pagePath: string): Promise<PopulateLog> {
    const pageFields = this.fieldMap.filter(f => 
      f.CAS_Page === pagePath && 
      f.CAS_Type !== 'skip' && 
      f.CAS_Selector !== 'SKIP' && 
      f.CAS_Selector !== 'HANDLER'
    );
    
    const log: PopulateLog = {
      timestamp: new Date().toISOString(),
      page: pagePath,
      fieldsAttempted: pageFields.length,
      fieldsSuccessful: 0,
      fieldsFailed: 0,
      details: []
    };

    console.log('\n' + '─'.repeat(60));
    console.log(`📋 Processing page: ${pagePath}`);
    console.log(`   Fields to populate: ${pageFields.length}`);
    console.log('─'.repeat(60));

    // Navigate to page
    const navSuccess = await this.navigateToPage(pagePath);
    if (!navSuccess) {
      log.fieldsFailed = pageFields.length;
      return log;
    }

    // Special handling for /organizations page - check if org exists and needs editing
    if (pagePath === '/organizations') {
      await this.handleOrganizationsPage();
    }

    // Process each field
    for (const field of pageFields) {
      const value = this.excelData[field.Sheet]?.[field.Row];
      
      if (!value) {
        console.log(`\n   ⏭️  Skipping ${field.FieldLabel} (no value in Excel data)`);
        log.details.push({
          field: field.FieldLabel,
          selector: field.CAS_Selector,
          value: '',
          status: 'skipped'
        });
        continue;
      }

      const result = await this.populateField(field, value);
      
      log.details.push({
        field: field.FieldLabel,
        selector: field.CAS_Selector,
        value: value,
        status: result.success ? 'success' : 'failed',
        error: result.error
      });

      if (result.success) {
        log.fieldsSuccessful++;
      } else {
        log.fieldsFailed++;
      }
    }

    // Click save/update button to persist changes
    if (log.fieldsSuccessful > 0) {
      await this.clickSaveButton();
    }

    return log;
  }

  async run(): Promise<void> {
    try {
      await this.init();
      
      // Try to navigate directly first (using saved session/cookies)
      const sessionValid = await this.tryDirectNavigation();
      
      if (!sessionValid) {
        // Session not valid, need to login
        console.log('\n🔐 Session expired or not found, logging in...');
        const loginSuccess = await this.login();
        if (!loginSuccess) {
          console.log('❌ Login failed. Exiting.');
          return;
        }
      }

      // Get the starting page
      let startPage = '/name-and-type';  // Default first page
      
      if (CONFIG.continueFromPage) {
        startPage = CONFIG.continueFromPage;
        console.log(`\n⏩ Continuing from page: ${startPage}`);
      } else {
        console.log(`\n▶️  Starting from first page: ${startPage}`);
      }
      
      // Navigate to the starting page
      await this.navigateToPage(startPage);
      
      // Process pages by following the Next button workflow
      let pageCount = 0;
      let continueProcessing = true;
      
      while (continueProcessing) {
        pageCount++;
        
        // Get current page path from URL
        const currentUrl = this.page?.url() || '';
        const urlPath = new URL(currentUrl).pathname;
        const pagePath = urlPath.replace(`/appraisals/${CONFIG.appraisalId}`, '');
        
        console.log(`\n${'═'.repeat(60)}`);
        console.log(`PAGE ${pageCount}: ${pagePath}`);
        console.log('═'.repeat(60));
        
        // Check if we've reached the final page (Phase 1 complete)
        if (pagePath === CONFIG.finalPage || pagePath.includes(CONFIG.finalPage)) {
          await this.handlePhaseComplete();
          continueProcessing = false;
          break;
        }
        
        // Check if this page should be skipped
        const shouldSkip = CONFIG.skipPages.some(skip => pagePath.includes(skip));
        if (shouldSkip) {
          console.log(`   ⏭️  Skipping page (in skipPages list)`);
          
          // Click Next to move to the next page
          const hasNextPage = await this.clickNextButton();
          if (!hasNextPage) {
            console.log('\n✅ No more pages (reached end of workflow)');
            continueProcessing = false;
          } else {
            await new Promise(resolve => setTimeout(resolve, 2000));
          }
          continue;
        }
        
        // Process the current page
        const log = await this.processCurrentPage(pagePath);
        this.logs.push(log);
        
        // Show summary
        console.log(`\n📊 Page Summary:`);
        console.log(`   ✅ Successful: ${log.fieldsSuccessful}`);
        console.log(`   ❌ Failed: ${log.fieldsFailed}`);
        console.log(`   ⏭️  Skipped: ${log.details.filter(d => d.status === 'skipped').length}`);
        
        // Ask for feedback
        const { continue: shouldContinue, feedback } = await this.askForFeedback();
        
        if (feedback) {
          log.userFeedback = feedback;
        }
        
        // Save logs after each page
        this.saveLogs();
        
        if (!shouldContinue) {
          console.log('\n🛑 Stopping as requested.');
          break;
        }
        
        // Try to click the Next button to go to the next page
        const hasNextPage = await this.clickNextButton();
        
        if (!hasNextPage) {
          console.log('\n✅ No more pages (reached end of workflow)');
          continueProcessing = false;
        } else {
          // Wait for the new page to load
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
      
      // Keep browser open for review
      console.log('\n🌐 Browser left open for review. Close manually when done.');
    }
  }

  saveLogs(): void {
    fs.writeFileSync(CONFIG.logFile, JSON.stringify(this.logs, null, 2));
  }

  // Save HTML of current page for debugging/analysis
  async savePageHtml(pagePath: string, suffix: string = ''): Promise<string> {
    if (!this.page) return '';
    
    try {
      const html = await this.page.content();
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const safePath = pagePath.replace(/\//g, '_').replace(/^_/, '');
      const filename = `${timestamp}_${safePath}${suffix ? '_' + suffix : ''}.html`;
      const filepath = path.join(CONFIG.htmlLogDir, filename);
      
      fs.writeFileSync(filepath, html, 'utf-8');
      console.log(`   💾 Saved HTML: ${filename}`);
      
      return filepath;
    } catch (error) {
      console.log(`   ⚠️  Could not save HTML: ${error}`);
      return '';
    }
  }
}

// Main
const populator = new CASPopulator();
populator.run();