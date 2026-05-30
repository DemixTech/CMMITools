/**
 * CAS Online Form Field Scraper v2
 * 
 * This script logs into the CMMI CAS portal and extracts form fields
 * from ALL pages of a specific appraisal, including all subpages.
 * 
 * Usage:
 *   1. npm install
 *   2. Set environment variables:
 *      set CAS_EMAIL=your@email.com
 *      set CAS_PASSWORD=yourpassword
 *   3. npm start
 * 
 * Output: 
 *   - cas_form_fields.json (structured data)
 *   - screenshots/*.png (visual captures)
 *   - html/*.html (raw HTML for each page)
 */

import puppeteer, { Browser, Page } from 'puppeteer';
import * as fs from 'fs';
import * as path from 'path';

// Load project config and keys
const projectConfigPath = '../cas-project-config.json';
let projectConfig: any = null;
let keysConfig: any = null;

try {
  if (fs.existsSync(projectConfigPath)) {
    projectConfig = JSON.parse(fs.readFileSync(projectConfigPath, 'utf-8'));
    
    // Load keys from separate file
    if (projectConfig?.keysFile && fs.existsSync(projectConfig.keysFile)) {
      keysConfig = JSON.parse(fs.readFileSync(projectConfig.keysFile, 'utf-8'));
    }
  }
} catch (e) {
  // Config optional, continue with defaults
}

// Configuration
const CONFIG = {
  loginUrl: projectConfig?.cas?.loginUrl || 'https://cmmiinstitute.com/login',
  dashboardUrl: 'https://cmmiinstitute.com/dashboard',
  casBaseUrl: projectConfig?.cas?.baseUrl || 'https://cas.cmmiinstitute.com',
  appraisalId: projectConfig?.project?.casId || '81846',
  
  // Credentials from keys.json (preferred) or environment variables (fallback)
  email: keysConfig?.cas?.email || process.env.CAS_EMAIL || '',
  password: keysConfig?.cas?.password || process.env.CAS_PASSWORD || '',
  staySignedIn: keysConfig?.cas?.staySignedIn?.toLowerCase() === 'yes',
  
  // Timing - increased for slow server
  navigationTimeout: 90000,
  waitBetweenPages: 3000,
  
  // Retry settings
  maxRetries: 3,
  retryDelay: 5000,
  
  // Output directories
  outputFile: 'cas_form_fields.json',
  screenshotDir: 'screenshots',
  htmlDir: 'html'
};

// EXPLICIT LIST OF ALL CAS PAGES TO SCRAPE
// This ensures we visit every page, not just those in navigation
const PAGES_TO_SCRAPE = [
  // Phase 1: Organization Scope
  { path: '/name-and-type', name: 'Appraisal Name and Type' },
  { path: '/organizations', name: 'Organizational Info' },
  { path: '/timeline', name: 'Appraisal Timeline' },
  { path: '/readiness-reviews', name: 'Readiness Reviews' },
  
  // Phase 1: Appraisal Personnel
  { path: '/training-schedule', name: 'ATM Training Schedule' },
  { path: '/confidentiality-agreement', name: 'Confidentiality and Non-Attribution' },
  
  // Phase 1: Readiness - OE Collection Plan
  { path: '/objective-evidence/collection-approach', name: 'OE Collection Approach' },
  { path: '/objective-evidence/collection-techniques', name: 'OE Collection Techniques' },
  { path: '/objective-evidence/collection-responsibilities', name: 'Responsibility for Collection' },
  { path: '/objective-evidence/performance-report-approaches', name: 'Performance Report Collection Approach' },
  { path: '/objective-evidence/initial-summary', name: 'Summary of Initial OE' },
  { path: '/objective-evidence/data-collection-timing', name: 'Data Collection Timing' },
  { path: '/objective-evidence/additional-info', name: 'Additional Information' },
  
  // Phase 1: Readiness - Logistics and Constraints
  { path: '/resource-estimates', name: 'Resource Effort Estimate' },
  { path: '/logistical-requirements', name: 'Logistics Requirements' },
  { path: '/appraisal-constraints', name: 'Appraisal Constraints' },
  { path: '/risk-identification', name: 'Risk Identification and Management' },
  { path: '/conflicts-of-interest', name: 'COI Identification and Management' },
  { path: '/follow-on-activities', name: 'Optional Follow-on Activities' },
  
  // Appraisal Documents
  { path: '/appraisal-plan-summary', name: 'Appraisal Plan Summary and Signature' },
  { path: '/supporting-documents', name: 'Supporting Document Upload' },
  
  // Appraisal Outputs
  { path: '/required-outputs', name: 'Required Outputs' },
  { path: '/performance-report', name: 'Performance Report Output' },
  { path: '/optional-outputs', name: 'Optional Outputs' },
  
  // Additional pages (sample scope, etc.)
  { path: '/sample-scope', name: 'Sample Scope' },
  { path: '/random-sample', name: 'Random Sample' },
  { path: '/menu', name: 'Appraisal Menu' },
];

interface FormField {
  name: string;
  label: string;
  type: string;
  id?: string;
  required?: boolean;
  options?: string[];
  value?: string;
  placeholder?: string;
  cssSelector?: string;
}

interface PageData {
  url: string;
  title: string;
  pageName: string;
  pagePath: string;
  htmlFile: string;
  screenshotFile: string;
  sections: {
    sectionName: string;
    fields: FormField[];
  }[];
  allLabels: string[];
  allButtons: string[];
  allLinks: { text: string; href: string }[];
  timestamp: string;
  status: 'success' | 'error' | 'not_found';
  errorMessage?: string;
}

interface ScrapedData {
  appraisalId: string;
  scrapedAt: string;
  version: string;
  pages: PageData[];
  navigation: {
    mainMenu: string[];
    discoveredLinks: string[];
  };
  summary: {
    totalPages: number;
    successfulPages: number;
    totalFields: number;
  };
}

class CASScraper {
  private browser: Browser | null = null;
  private page: Page | null = null;
  private scrapedData: ScrapedData;
  private pageCounter: number = 0;
  private visitedUrls: Set<string> = new Set();

  private async reconnectBrowser(): Promise<void> {
    try {
      if (this.browser) {
        try { await this.browser.close(); } catch { }
      }
      
      this.browser = await puppeteer.launch({
        headless: false,
        defaultViewport: { width: 1920, height: 1080 },
        args: ['--start-maximized']
      });

      this.page = await this.browser.newPage();
      this.page.setDefaultNavigationTimeout(CONFIG.navigationTimeout);
      
      // Re-login
      console.log('    🔐 Re-logging in...');
      await this.login();
      
      console.log('    ✅ Browser reconnected');
    } catch (error) {
      console.error('    ❌ Failed to reconnect browser:', error);
      throw error;
    }
  }

  private saveIntermediateResults(): void {
    try {
      const intermediateFile = 'cas_form_fields_intermediate.json';
      fs.writeFileSync(intermediateFile, JSON.stringify(this.scrapedData, null, 2));
    } catch (error) {
      console.error('Failed to save intermediate results:', error);
    }
  }

  constructor() {
    this.scrapedData = {
      appraisalId: CONFIG.appraisalId,
      scrapedAt: new Date().toISOString(),
      version: '2.0.0',
      pages: [],
      navigation: {
        mainMenu: [],
        discoveredLinks: []
      },
      summary: {
        totalPages: 0,
        successfulPages: 0,
        totalFields: 0
      }
    };
  }

  async init(): Promise<void> {
    console.log('🚀 Starting CAS Scraper v2.0...');
    console.log(`📋 Will scrape ${PAGES_TO_SCRAPE.length} predefined pages`);
    
    // Validate credentials
    if (!CONFIG.email || !CONFIG.password) {
      console.error('❌ Credentials not set!');
      console.log(`Please configure credentials in ${CONFIG.keysPath}:`);
      console.log('  {');
      console.log('    "cas": {');
      console.log('      "email": "your@email.com",');
      console.log('      "password": "yourpassword",');
      console.log('      "staySignedIn": "yes"');
      console.log('    }');
      console.log('  }');
      process.exit(1);
    }
    
    console.log(`🔑 Using credentials from ${CONFIG.keysPath} (email: ${CONFIG.email.substring(0, 3)}...)`);
    
    // Create output directories
    for (const dir of [CONFIG.screenshotDir, CONFIG.htmlDir]) {
      if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
      }
    }

    this.browser = await puppeteer.launch({
      headless: false, // Set to true for headless mode
      defaultViewport: { width: 1920, height: 1080 },
      args: ['--start-maximized']
    });

    this.page = await this.browser.newPage();
    this.page.setDefaultNavigationTimeout(CONFIG.navigationTimeout);
    
    console.log('✅ Browser initialized');
  }

  async login(): Promise<boolean> {
    if (!this.page) throw new Error('Page not initialized');

    console.log('🔐 Logging in to CMMI Institute...');
    
    try {
      await this.page.goto(CONFIG.loginUrl, { waitUntil: 'networkidle2' });
      await this.capturePageState('01_login_page');

      // Wait longer for page to fully load (cookie consent dialogs, etc.)
      await new Promise(resolve => setTimeout(resolve, 3000));
      
      // CMMI login uses #UserName (type="text") not type="email"
      await this.page.waitForSelector('#UserName, input[name="UserName"]', { timeout: 15000 });

      const emailSelector = await this.findSelector([
        '#UserName',
        'input[name="UserName"]',
        'input[type="email"]',
        'input[name="email"]',
        'input#email'
      ]);
      
      if (emailSelector) {
        // Click first to focus, then clear any existing value, then type
        await this.page.click(emailSelector);
        await this.page.evaluate((sel) => {
          const el = document.querySelector(sel) as HTMLInputElement;
          if (el) el.value = '';
        }, emailSelector);
        await this.page.type(emailSelector, CONFIG.email, { delay: 30 });
        console.log(`    ✅ Entered email in ${emailSelector}`);
      } else {
        console.log('    ❌ Could not find email/username field!');
      }

      const passwordSelector = await this.findSelector([
        '#Password',
        'input[name="Password"]',
        'input[type="password"]'
      ]);
      
      if (passwordSelector) {
        await this.page.click(passwordSelector);
        await this.page.type(passwordSelector, CONFIG.password, { delay: 30 });
        console.log(`    ✅ Entered password in ${passwordSelector}`);
      } else {
        console.log('    ❌ Could not find password field!');
      }

      // Handle "Stay signed in" checkbox
      await this.handleStaySignedIn();

      await this.capturePageState('02_credentials_entered');

      const submitSelector = await this.findSelector([
        'button[type="submit"]',
        'input[type="submit"]',
        '.login-button',
        '#login-button'
      ]);

      if (submitSelector) {
        await Promise.all([
          this.page.waitForNavigation({ waitUntil: 'networkidle2' }),
          this.page.click(submitSelector)
        ]);
      }

      await this.capturePageState('03_after_login');

      const currentUrl = this.page.url();
      if (currentUrl.includes('dashboard') || !currentUrl.includes('login')) {
        console.log('✅ Login successful');
        return true;
      } else {
        console.log('❌ Login may have failed. Current URL:', currentUrl);
        return false;
      }

    } catch (error) {
      console.error('❌ Login failed:', error);
      await this.capturePageState('error_login');
      return false;
    }
  }

  private async handleStaySignedIn(): Promise<void> {
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

  async navigateToCAS(): Promise<boolean> {
    if (!this.page) throw new Error('Page not initialized');

    console.log('📂 Navigating to CAS Dashboard...');

    try {
      const appraisalUrl = `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}`;
      await this.page.goto(appraisalUrl, { waitUntil: 'networkidle2' });
      
      await new Promise(resolve => setTimeout(resolve, 3000));
      await this.capturePageState('04_cas_dashboard');

      console.log('✅ Navigated to CAS appraisal:', CONFIG.appraisalId);
      return true;

    } catch (error) {
      console.error('❌ Navigation to CAS failed:', error);
      await this.capturePageState('error_cas_navigation');
      return false;
    }
  }

  async scrapeAllPredefinedPages(): Promise<void> {
    if (!this.page) throw new Error('Page not initialized');

    console.log('\n📄 Scraping all predefined pages...\n');

    for (let i = 0; i < PAGES_TO_SCRAPE.length; i++) {
      const pageInfo = PAGES_TO_SCRAPE[i];
      const fullUrl = `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}${pageInfo.path}`;
      
      console.log(`[${i + 1}/${PAGES_TO_SCRAPE.length}] ${pageInfo.name}`);
      console.log(`    URL: ${fullUrl}`);

      let success = false;
      let lastError: any = null;
      
      for (let attempt = 1; attempt <= CONFIG.maxRetries && !success; attempt++) {
        try {
          if (attempt > 1) {
            console.log(`    🔄 Retry attempt ${attempt}/${CONFIG.maxRetries}...`);
            await new Promise(resolve => setTimeout(resolve, CONFIG.retryDelay));
          }
          
          // Navigate to the page with timeout handling
          const response = await this.page.goto(fullUrl, { 
            waitUntil: 'networkidle2',
            timeout: CONFIG.navigationTimeout 
          });
          await new Promise(resolve => setTimeout(resolve, CONFIG.waitBetweenPages));

          // Check if page exists (not 404)
          const status = response?.status() || 0;
          
          if (status === 404 || status >= 400) {
            console.log(`    ⚠️  Page returned status ${status}, skipping`);
            this.scrapedData.pages.push({
              url: fullUrl,
              title: '',
              pageName: pageInfo.name,
              pagePath: pageInfo.path,
              htmlFile: '',
              screenshotFile: '',
              sections: [],
              allLabels: [],
              allButtons: [],
              allLinks: [],
              timestamp: new Date().toISOString(),
              status: 'not_found',
              errorMessage: `HTTP ${status}`
            });
            success = true; // Don't retry 404s
            continue;
          }

          // Extract page data
          const pageData = await this.extractPageFormFields(pageInfo.name, pageInfo.path);
          this.scrapedData.pages.push(pageData);
          this.visitedUrls.add(fullUrl);
          
          const fieldCount = pageData.sections.reduce((acc, s) => acc + s.fields.length, 0);
          console.log(`    ✅ ${pageData.sections.length} sections, ${fieldCount} fields, ${pageData.allLabels.length} labels`);
          success = true;

        } catch (error) {
          lastError = error;
          console.error(`    ❌ Attempt ${attempt} failed:`, String(error).substring(0, 100));
          
          // Check if browser is still alive
          if (this.page?.isClosed() || !this.browser?.isConnected()) {
            console.log('    🔧 Browser disconnected, attempting to reconnect...');
            await this.reconnectBrowser();
          }
        }
      }
      
      if (!success) {
        console.error(`    ❌ All ${CONFIG.maxRetries} attempts failed for ${pageInfo.name}`);
        this.scrapedData.pages.push({
          url: fullUrl,
          title: '',
          pageName: pageInfo.name,
          pagePath: pageInfo.path,
          htmlFile: '',
          screenshotFile: '',
          sections: [],
          allLabels: [],
          allButtons: [],
          allLinks: [],
          timestamp: new Date().toISOString(),
          status: 'error',
          errorMessage: String(lastError)
        });
      }
      
      // Save intermediate results every 5 pages
      if (i > 0 && i % 5 === 0) {
        console.log('    💾 Saving intermediate results...');
        this.saveIntermediateResults();
      }
    }

    // Discover any additional links we might have missed
    await this.discoverAdditionalLinks();
  }

  async discoverAdditionalLinks(): Promise<void> {
    if (!this.page) return;

    console.log('\n🔍 Discovering additional links...');

    // Go back to the menu page to find any links we missed
    try {
      const menuUrl = `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}/menu`;
      await this.page.goto(menuUrl, { waitUntil: 'networkidle2' });
      
      const allLinks = await this.page.evaluate((appraisalId) => {
        const links: string[] = [];
        document.querySelectorAll('a[href]').forEach(a => {
          const href = a.getAttribute('href');
          if (href && href.includes(appraisalId) && !href.includes('#')) {
            links.push(href);
          }
        });
        return [...new Set(links)];
      }, CONFIG.appraisalId);

      // Find links we haven't visited
      const newLinks = allLinks.filter(link => {
        const fullUrl = link.startsWith('/') ? `${CONFIG.casBaseUrl}${link}` : link;
        return !this.visitedUrls.has(fullUrl);
      });

      if (newLinks.length > 0) {
        console.log(`  Found ${newLinks.length} additional unvisited links`);
        this.scrapedData.navigation.discoveredLinks = newLinks;
      } else {
        console.log('  No additional links found');
      }

    } catch (error) {
      console.error('Error discovering links:', error);
    }
  }

  async extractPageFormFields(pageName: string, pagePath: string): Promise<PageData> {
    if (!this.page) throw new Error('Page not initialized');

    this.pageCounter++;
    const safeName = pageName.replace(/[^a-z0-9]/gi, '_').substring(0, 40);
    const filePrefix = `${String(this.pageCounter).padStart(2, '0')}_${safeName}`;
    
    const { htmlFile, screenshotFile } = await this.capturePageState(filePrefix);

    const pageData: PageData = {
      url: this.page.url(),
      title: await this.page.title(),
      pageName: pageName,
      pagePath: pagePath,
      htmlFile: htmlFile,
      screenshotFile: screenshotFile,
      sections: [],
      allLabels: [],
      allButtons: [],
      allLinks: [],
      timestamp: new Date().toISOString(),
      status: 'success'
    };

    try {
      // Extract ALL labels on the page
      pageData.allLabels = await this.page.evaluate(() => {
        const labels: string[] = [];
        document.querySelectorAll('label, .label, .field-label, .form-label, th, dt, legend, .control-label').forEach(el => {
          const text = el.textContent?.trim();
          if (text && text.length < 200 && text.length > 0) {
            labels.push(text);
          }
        });
        return [...new Set(labels)];
      });

      // Extract ALL buttons
      pageData.allButtons = await this.page.evaluate(() => {
        const buttons: string[] = [];
        document.querySelectorAll('button, input[type="button"], input[type="submit"], .btn, [role="button"]').forEach(el => {
          const text = el.textContent?.trim() || (el as HTMLInputElement).value;
          if (text) {
            buttons.push(text);
          }
        });
        return [...new Set(buttons)];
      });

      // Extract ALL links
      pageData.allLinks = await this.page.evaluate(() => {
        const links: { text: string; href: string }[] = [];
        document.querySelectorAll('a[href]').forEach(el => {
          const text = el.textContent?.trim();
          const href = el.getAttribute('href');
          if (text && href) {
            links.push({ text: text.substring(0, 100), href });
          }
        });
        return links;
      });

      // Extract all form fields with detailed information
      const fields = await this.page.evaluate(() => {
        const extractedFields: any[] = [];
        const inputs = document.querySelectorAll('input, select, textarea');
        
        inputs.forEach((input, index) => {
          const el = input as HTMLInputElement | HTMLSelectElement | HTMLTextAreaElement;
          
          // Find associated label
          let label = '';
          const id = el.id;
          
          // Method 1: label[for="id"]
          if (id) {
            const labelEl = document.querySelector(`label[for="${id}"]`);
            label = labelEl?.textContent?.trim() || '';
          }
          
          // Method 2: Parent label
          if (!label) {
            const parentLabel = el.closest('label');
            if (parentLabel) {
              const clone = parentLabel.cloneNode(true) as HTMLElement;
              clone.querySelectorAll('input, select, textarea').forEach(i => i.remove());
              label = clone.textContent?.trim() || '';
            }
          }
          
          // Method 3: Previous sibling or nearby elements
          if (!label) {
            const parent = el.closest('.form-group, .field, .input-group, .form-field, .field-wrapper, tr, .row, .control-group');
            if (parent) {
              const labelEl = parent.querySelector('label, .label, .field-label, th, dt, .control-label');
              label = labelEl?.textContent?.trim() || '';
            }
          }

          // Method 4: aria-label
          if (!label) {
            label = el.getAttribute('aria-label') || '';
          }

          // Method 5: placeholder as fallback
          if (!label && (el as HTMLInputElement).placeholder) {
            label = (el as HTMLInputElement).placeholder;
          }

          // Get field type
          let fieldType = el.tagName.toLowerCase();
          if (fieldType === 'input') {
            fieldType = (el as HTMLInputElement).type || 'text';
          }

          // Get options for select elements
          const options: string[] = [];
          if (el.tagName === 'SELECT') {
            const selectEl = el as HTMLSelectElement;
            Array.from(selectEl.options).forEach(opt => {
              const optText = opt.text.trim();
              if (optText) {
                options.push(optText);
              }
            });
          }

          // Get options for radio/checkbox groups
          if (el.type === 'radio' || el.type === 'checkbox') {
            const name = el.name;
            if (name) {
              const siblings = document.querySelectorAll(`input[name="${name}"]`);
              siblings.forEach(sib => {
                const sibLabel = document.querySelector(`label[for="${sib.id}"]`);
                const optText = sibLabel?.textContent?.trim() || (sib as HTMLInputElement).value;
                if (optText && !options.includes(optText)) {
                  options.push(optText);
                }
              });
            }
          }

          // Build CSS selector
          let cssSelector = el.tagName.toLowerCase();
          if (el.id) {
            cssSelector = `#${el.id}`;
          } else if (el.name) {
            cssSelector = `${el.tagName.toLowerCase()}[name="${el.name}"]`;
          }

          extractedFields.push({
            name: el.name || el.id || `unnamed_${index}`,
            label: label,
            type: fieldType,
            id: el.id || undefined,
            required: el.required || el.hasAttribute('required') || el.hasAttribute('aria-required'),
            options: options.length > 0 ? options : undefined,
            value: el.value || undefined,
            placeholder: (el as HTMLInputElement).placeholder || undefined,
            cssSelector: cssSelector
          });
        });

        return extractedFields;
      });

      // Put all fields in one section for simplicity
      if (fields.length > 0) {
        pageData.sections.push({
          sectionName: 'Form Fields',
          fields: fields as FormField[]
        });
      }

    } catch (error) {
      console.error('Error extracting form fields:', error);
      pageData.status = 'error';
      pageData.errorMessage = String(error);
    }

    return pageData;
  }

  async saveResults(): Promise<void> {
    // Calculate summary
    this.scrapedData.summary = {
      totalPages: this.scrapedData.pages.length,
      successfulPages: this.scrapedData.pages.filter(p => p.status === 'success').length,
      totalFields: this.scrapedData.pages.reduce((acc, p) => 
        acc + p.sections.reduce((a, s) => a + s.fields.length, 0), 0)
    };

    fs.writeFileSync(CONFIG.outputFile, JSON.stringify(this.scrapedData, null, 2));
    console.log(`\n💾 Results saved to: ${CONFIG.outputFile}`);
    
    // Create summary
    const summary = {
      version: '2.0.0',
      totalPages: this.scrapedData.summary.totalPages,
      successfulPages: this.scrapedData.summary.successfulPages,
      totalFields: this.scrapedData.summary.totalFields,
      pages: this.scrapedData.pages.map(p => ({
        name: p.pageName,
        path: p.pagePath,
        url: p.url,
        status: p.status,
        htmlFile: p.htmlFile,
        screenshotFile: p.screenshotFile,
        fields: p.sections.reduce((acc, s) => acc + s.fields.length, 0),
        labels: p.allLabels.length,
        buttons: p.allButtons.length
      }))
    };

    fs.writeFileSync('cas_summary.json', JSON.stringify(summary, null, 2));
    console.log(`📊 Summary saved to: cas_summary.json`);

    // Generate HTML index
    const indexHtml = this.generateHtmlIndex();
    fs.writeFileSync('cas_index.html', indexHtml);
    console.log(`📑 HTML index saved to: cas_index.html`);

    // Generate _FieldMapCAS update suggestions
    this.generateFieldMapSuggestions();
  }

  private generateFieldMapSuggestions(): void {
    const suggestions: string[] = [];
    suggestions.push('Row,Sheet,FieldLabel,CAS_Page,CAS_Selector,CAS_FieldName,CAS_Type,Notes');

    for (const page of this.scrapedData.pages) {
      if (page.status !== 'success') continue;

      for (const section of page.sections) {
        for (const field of section.fields) {
          if (field.type === 'hidden') continue;
          
          suggestions.push([
            '', // Row - to be filled manually
            '', // Sheet - to be filled manually
            `"${field.label || field.name}"`,
            page.pagePath,
            `"${field.cssSelector}"`,
            field.name,
            field.type,
            field.options ? `Options: ${field.options.slice(0, 3).join(', ')}...` : ''
          ].join(','));
        }
      }
    }

    fs.writeFileSync('cas_fieldmap_suggestions.csv', suggestions.join('\n'));
    console.log(`📋 Field mapping suggestions saved to: cas_fieldmap_suggestions.csv`);
  }

  private generateHtmlIndex(): string {
    let html = `<!DOCTYPE html>
<html>
<head>
  <title>CAS Scrape Results v2 - Appraisal ${CONFIG.appraisalId}</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; background: #f5f5f5; }
    h1, h2, h3 { color: #333; }
    .container { max-width: 1400px; margin: 0 auto; }
    .summary-box { background: #fff; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .summary-stats { display: flex; gap: 20px; }
    .stat { background: #4472C4; color: white; padding: 15px 25px; border-radius: 8px; text-align: center; }
    .stat-value { font-size: 2em; font-weight: bold; }
    .stat-label { font-size: 0.9em; opacity: 0.9; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 20px; background: white; }
    th, td { border: 1px solid #ddd; padding: 10px; text-align: left; }
    th { background-color: #4472C4; color: white; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    tr:hover { background-color: #f0f0f0; }
    .page-section { background: white; margin-bottom: 20px; padding: 20px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .status-success { color: green; }
    .status-error { color: red; }
    .status-not_found { color: orange; }
    .field-required { color: red; font-weight: bold; }
    a { color: #0066cc; }
    code { background: #eee; padding: 2px 6px; border-radius: 3px; font-size: 0.9em; }
  </style>
</head>
<body>
  <div class="container">
    <h1>🔍 CAS Scrape Results - Appraisal ${CONFIG.appraisalId}</h1>
    
    <div class="summary-box">
      <h2>Summary</h2>
      <p>Scraped at: ${this.scrapedData.scrapedAt}</p>
      <div class="summary-stats">
        <div class="stat">
          <div class="stat-value">${this.scrapedData.summary.totalPages}</div>
          <div class="stat-label">Total Pages</div>
        </div>
        <div class="stat">
          <div class="stat-value">${this.scrapedData.summary.successfulPages}</div>
          <div class="stat-label">Successful</div>
        </div>
        <div class="stat">
          <div class="stat-value">${this.scrapedData.summary.totalFields}</div>
          <div class="stat-label">Total Fields</div>
        </div>
      </div>
    </div>
    
    <div class="summary-box">
      <h2>Pages Overview</h2>
      <table>
        <tr>
          <th>#</th>
          <th>Page Name</th>
          <th>Path</th>
          <th>Status</th>
          <th>Fields</th>
          <th>Files</th>
        </tr>`;

    this.scrapedData.pages.forEach((page, index) => {
      const fieldCount = page.sections.reduce((acc, s) => acc + s.fields.length, 0);
      html += `
        <tr>
          <td>${index + 1}</td>
          <td><a href="#page-${index}">${page.pageName}</a></td>
          <td><code>${page.pagePath}</code></td>
          <td class="status-${page.status}">${page.status}</td>
          <td>${fieldCount}</td>
          <td>
            ${page.htmlFile ? `<a href="${page.htmlFile}" target="_blank">HTML</a>` : '-'} | 
            ${page.screenshotFile ? `<a href="${page.screenshotFile}" target="_blank">Screenshot</a>` : '-'}
          </td>
        </tr>`;
    });

    html += `
      </table>
    </div>
    
    <h2>Detailed Field Information</h2>`;

    this.scrapedData.pages.forEach((page, pageIndex) => {
      if (page.status !== 'success') return;
      
      html += `
    <div class="page-section" id="page-${pageIndex}">
      <h3>${pageIndex + 1}. ${page.pageName}</h3>
      <p><strong>Path:</strong> <code>${page.pagePath}</code></p>
      <p><strong>URL:</strong> <a href="${page.url}" target="_blank">${page.url}</a></p>`;

      page.sections.forEach(section => {
        if (section.fields.length === 0) return;
        
        html += `
      <table>
        <tr>
          <th>Field Name</th>
          <th>Label</th>
          <th>Type</th>
          <th>CSS Selector</th>
          <th>Options/Value</th>
        </tr>`;

        section.fields.forEach(field => {
          if (field.type === 'hidden') return;
          const options = field.options ? field.options.slice(0, 3).join(', ') + (field.options.length > 3 ? '...' : '') : '';
          html += `
        <tr>
          <td><code>${field.name}</code></td>
          <td>${field.label || '-'}</td>
          <td>${field.type}</td>
          <td><code>${field.cssSelector}</code></td>
          <td>${options || field.value || '-'}</td>
        </tr>`;
        });

        html += `
      </table>`;
      });

      html += `
    </div>`;
    });

    html += `
  </div>
</body>
</html>`;

    return html;
  }

  async close(): Promise<void> {
    if (this.browser) {
      await this.browser.close();
      console.log('🔒 Browser closed');
    }
  }

  private async findSelector(selectors: string[]): Promise<string | null> {
    if (!this.page) return null;
    
    for (const selector of selectors) {
      try {
        const element = await this.page.$(selector);
        if (element) return selector;
      } catch { }
    }
    return null;
  }

  private async capturePageState(name: string): Promise<{ htmlFile: string; screenshotFile: string }> {
    const htmlFile = path.join(CONFIG.htmlDir, `${name}.html`);
    const screenshotFile = path.join(CONFIG.screenshotDir, `${name}.png`);

    if (!this.page) return { htmlFile, screenshotFile };

    try {
      const html = await this.page.content();
      fs.writeFileSync(htmlFile, html, 'utf-8');

      await this.page.screenshot({ path: screenshotFile, fullPage: true });

    } catch (error) {
      console.error(`  Error capturing page state for ${name}:`, error);
    }

    return { htmlFile, screenshotFile };
  }
}

// Main execution
async function main() {
  const scraper = new CASScraper();

  try {
    await scraper.init();
    
    const loginSuccess = await scraper.login();
    if (!loginSuccess) {
      console.log('\n⚠️  Login failed. Please check credentials.');
      return;
    }

    const navSuccess = await scraper.navigateToCAS();
    if (!navSuccess) {
      console.log('\n⚠️  Could not navigate to CAS.');
      return;
    }

    await scraper.scrapeAllPredefinedPages();
    await scraper.saveResults();

    console.log('\n✅ Scraping completed successfully!');
    console.log('\n📁 Output files:');
    console.log('   - cas_form_fields.json (structured data)');
    console.log('   - cas_summary.json (overview)');
    console.log('   - cas_index.html (browsable report)');
    console.log('   - cas_fieldmap_suggestions.csv (for _FieldMapCAS)');
    console.log('   - html/*.html (raw HTML for each page)');
    console.log('   - screenshots/*.png (screenshots)');

  } catch (error) {
    console.error('\n❌ Error during scraping:', error);
  } finally {
    await scraper.close();
  }
}

main();
