/**
 * Org Scope Page Handlers - Mixin applied to CASPopulator
 *
 * Covers all Phase 1 Org Scope pages:
 *   /organizations, /org-units, /org-unit-targets, /timeline,
 *   /readiness-reviews, /org-projects (true/false),
 *   /org-unit-sampling-factors, /org-unit-sampling-factor-values,
 *   /org-unit-subgroups, /org-unit-project-subgroups,
 *   /organizational-project-appraisal-scope
 *
 * Applied via: applyOrgScopeHandlers(CASPopulator) in populator.ts
 */

import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import AdmZip from 'adm-zip';
import { CONFIG } from '../config';

export function applyOrgScopeHandlers(cls: any): void {
// Handle organizations page - check if org exists (edit mode) or not (add mode)
cls.prototype.handleOrganizationsPage = async function(): Promise<void> {
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
cls.prototype.handleOrgUnitsPage = async function(): Promise<void> {
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
cls.prototype.handleOrgUnitTargetsPage = async function(): Promise<void> {
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
    console.log('   ⚠️  Error checking org-unit-targets page:', error);
  }
}

// Handle the timeline page which has multiple sub-forms
// Note: Readiness Reviews are handled on /readiness-reviews page separately
cls.prototype.handleTimelinePage = async function(): Promise<void> {
  if (!this.page) return;
  
  console.log('\n   📅 Processing Timeline page (multi-form)...');
  
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
      console.log(`   ⚠️  No Phase 1 Start Date found in Excel data!`);
    }
    if (phase1EndDate) {
      console.log(`   Setting Phase 1 End Date: ${phase1EndDate}`);
      await this.populateDateParts('#EndDateYear,#EndDateMonth,#EndDateDay', phase1EndDate);
    } else {
      console.log(`   ⚠️  No Phase 1 End Date found in Excel data!`);
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
    ].filter((rr: any) => rr.name); // Only include RRs with names
    
    for (let i = 0; i < readinessReviews.length; i++) {
      const rr = readinessReviews[i];
      console.log(`\n   === Readiness Review ${i + 1}: ${rr.name} ===`);
      
      // Navigate to timeline page first to check if RR exists
      await this.page.goto(`${baseUrl}/appraisals/${appraisalId}/timeline`, { waitUntil: 'networkidle2' });
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      // Check if this RR already exists
      const rrExists = await this.page.evaluate((rrName: any) => {
        const cards = document.querySelectorAll('.appraisal-timeline-readiness-review, .item-card');
        for (const card of Array.from(cards)) {
          if (card.textContent?.includes(rrName)) {
            return true;
          }
        }
        return false;
      }, rr.name);
      
      if (rrExists) {
        console.log(`   ℹ️  Readiness Review "${rr.name}" already exists, skipping`);
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
          console.log(`   ⚠️  Could not fill Name field: ${e}`);
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
      
      console.log(`   ✅ Readiness Review ${i + 1} saved`);
    }
    
    // Return to timeline main page
    await this.page.goto(`${baseUrl}/appraisals/${appraisalId}/timeline`, { waitUntil: 'networkidle2' });
    
    console.log('\n   ✅ Timeline page processing complete (including Readiness Reviews)');
    
  } catch (error) {
    console.log(`   ⚠️  Error processing timeline: ${error}`);
  }
}

// Handle the readiness-reviews page which has repeating sections
// This page is for EDITING existing readiness reviews with additional fields
// The basic readiness reviews (name, dates) are created on the timeline page
cls.prototype.handleReadinessReviewsPage = async function(): Promise<void> {
  if (!this.page) return;
  
  console.log('\n   📋 Processing Readiness Reviews page (editing additional fields)...');
  
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
  ].filter((rr: any) => rr.name); // Only include RRs with names
  
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
      const editClicked = await this.page.evaluate((rrName: any) => {
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
        console.log(`   ⚠️  Readiness Review "${rr.name}" not found on page`);
        continue;
      }
      
      if (!editClicked.clicked) {
        console.log(`   ⚠️  Could not find Edit button for "${rr.name}"`);
        continue;
      }
      
      console.log(`   ✅ Clicked Edit for "${rr.name}"`);
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      // Wait for form to load
      try {
        await this.page.waitForSelector('#Name', { timeout: 5000 });
      } catch (e) {
        console.log(`   ⚠️  Edit form did not load`);
        continue;
      }
      
      // Fill in the additional fields
      // Objectives
      if (rr.objectives) {
        console.log(`   Setting Objectives...`);
        try {
          await this.populateTextInput('#Objectives', rr.objectives);
        } catch (e) {
          console.log(`   ⚠️  Could not fill Objectives: ${e}`);
        }
      }
      
      // Success Criteria
      if (rr.successCriteria) {
        console.log(`   Setting Success Criteria...`);
        try {
          await this.populateTextInput('#SuccessCriteria', rr.successCriteria);
        } catch (e) {
          console.log(`   ⚠️  Could not fill Success Criteria: ${e}`);
        }
      }
      
      // Required Members
      if (rr.requiredMembers) {
        console.log(`   Setting Required Members...`);
        try {
          await this.populateTextInput('#RequiredMembers', rr.requiredMembers);
        } catch (e) {
          console.log(`   ⚠️  Could not fill Required Members: ${e}`);
        }
      }
      
      // Outcomes
      if (rr.outcomes) {
        console.log(`   Setting Outcomes...`);
        try {
          await this.populateTextInput('#Outcomes', rr.outcomes);
        } catch (e) {
          console.log(`   ⚠️  Could not fill Outcomes: ${e}`);
        }
      }
      
      // Further Details
      if (rr.furtherDetails) {
        console.log(`   Setting Further Details...`);
        try {
          await this.populateTextInput('#FurtherDetails', rr.furtherDetails);
        } catch (e) {
          console.log(`   ⚠️  Could not fill Further Details: ${e}`);
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
          console.log(`   ⚠️  Could not set Characterized Evidence: ${e}`);
        }
      }
      
      // Click save/update button
      await this.clickSaveButton();
      await new Promise(resolve => setTimeout(resolve, 1500));
      
      console.log(`   ✅ Readiness Review ${i + 1} updated`);
    }
    
    // Return to readiness reviews main page
    await this.page.goto(`${baseUrl}/appraisals/${appraisalId}/readiness-reviews`, { waitUntil: 'networkidle2' });
    
    console.log('\n   ✅ Readiness Reviews page processing complete');
    
  } catch (error) {
    console.log(`   ⚠️  Error processing readiness reviews: ${error}`);
  }
}

// Handle the org-projects page - downloads template, populates from C_SupportV2, uploads
cls.prototype.handleOrgProjectsPage = async function(): Promise<void> {
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
        return Array.from(document.querySelectorAll('a')).map((a: any) => ({
          href: a.getAttribute('href'),
          text: a.textContent?.trim().substring(0, 50)
        }));
      });
      console.log('   Available links:');
      allLinks.filter((l: any) => l.href).forEach((l: any) => console.log(`      ${l.text} -> ${l.href}`));
      return;
    }
    
    console.log(`   📥 Download URL: ${downloadUrl}`);
    
    // Use fetch within the page context to download the file (preserves cookies/auth)
    const downloadResult = await this.page.evaluate(async (url: any) => {
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
    console.log(`   Headers: ${headers.filter((h: any) => h).join(', ').substring(0, 100)}...`);
    
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
      templateHeaders = stringMatches.slice(0, 18).map((m: any) => {
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
cls.prototype.handleOUProjectsPage = async function(): Promise<void> {
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
        return Array.from(document.querySelectorAll('a')).map((a: any) => ({
          href: a.getAttribute('href'),
          text: a.textContent?.trim().substring(0, 50)
        }));
      });
      console.log('   Available links:');
      allLinks.filter((l: any) => l.href).forEach((l: any) => console.log(`      ${l.text} -> ${l.href}`));
      return;
    }
    
    console.log(`   📥 Download URL: ${downloadUrl}`);
    
    // Use fetch within the page context to download the file (preserves cookies/auth)
    const downloadResult = await this.page.evaluate(async (url: any) => {
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
    console.log(`   Headers: ${headers.filter((h: any) => h).slice(0, 10).join(', ').substring(0, 100)}...`);
    
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
      templateHeaders = stringMatches.slice(0, 21).map((m: any) => {
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
cls.prototype.handleSamplingFactorsPage = async function(): Promise<void> {
  if (!this.page) return;
  
  console.log('\n   📋 Processing Sampling Factors page...');
  
  try {
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
    const sourceSheet = sourceWorkbook.getWorksheet('P1-OrgScope');
    
    if (!sourceSheet) {
      console.log('   ⚠️  P1-OrgScope sheet not found');
      return;
    }
    
    const getCellValue = (row: number, col: number): string => {
      const cell = sourceSheet.getCell(row, col);
      return this.resolveCellValue(sourceWorkbook, cell.value);
    };
    
    const samplingFactorName = getCellValue(127, 2);
    const otherSamplingFactor = getCellValue(128, 2);
    const definition = getCellValue(129, 2);
    
    console.log(`   Sampling Factor Name: ${samplingFactorName}`);
    console.log(`   Other: ${otherSamplingFactor || '(none)'}`);
    console.log(`   Definition: ${definition}`);
    
    const existingEntry = await this.page.evaluate((searchName: any) => {
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
    
    const formVisible = await this.page.$('select[name="StandardSamplingFactorId"]');
    
    if (!formVisible) {
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
    
    if (samplingFactorName) {
      try {
        await this.page.waitForSelector('select[name="StandardSamplingFactorId"]', { timeout: 5000 });
        const selectResult = await this.page.evaluate((searchText: any) => {
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
    
    if (otherSamplingFactor) {
      try {
        await this.page.waitForSelector('input[name="OtherName"], #OtherName', { timeout: 2000 });
        await this.page.evaluate((val: any) => {
          const el = document.querySelector('input[name="OtherName"], #OtherName') as HTMLInputElement;
          if (el) {
            el.value = val;
            el.dispatchEvent(new Event('input', { bubbles: true }));
          }
        }, otherSamplingFactor);
        console.log(`   ✅ Set Other Sampling Factor: ${otherSamplingFactor}`);
      } catch (e) {
        console.log(`   ℹ️  Other Sampling Factor field not needed`);
      }
    }
    
    if (definition) {
      try {
        await this.page.waitForSelector('textarea[name="Definition"], #Definition', { timeout: 3000 });
        await this.page.evaluate((val: any) => {
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
cls.prototype.handleSamplingFactorValuesPage = async function(): Promise<void> {
  if (!this.page) return;
  
  console.log('\n   📋 Processing Sampling Factor Values page...');
  
  try {
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
    const sourceSheet = sourceWorkbook.getWorksheet('P1-OrgScope');
    
    if (!sourceSheet) {
      console.log('   ⚠️  P1-OrgScope sheet not found');
      return;
    }
    
    const getCellValue = (row: number, col: number): string => {
      const cell = sourceSheet.getCell(row, col);
      return this.resolveCellValue(sourceWorkbook, cell.value);
    };
    
    const samplingFactor = getCellValue(132, 2);
    const value = getCellValue(133, 2);
    const description = getCellValue(134, 2);
    
    console.log(`   Sampling Factor: ${samplingFactor}`);
    console.log(`   Value: ${value}`);
    console.log(`   Description: ${description}`);
    
    const existingEntry = await this.page.evaluate((searchValue: any) => {
      const cards = document.querySelectorAll('.item-card');
      for (const card of Array.from(cards)) {
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
    
    const formVisible = await this.page.$('select[name="AppraisalOrgUnitSamplingFactorId"]');
    
    if (!formVisible) {
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
    
    if (samplingFactor) {
      try {
        await this.page.waitForSelector('select[name="AppraisalOrgUnitSamplingFactorId"]', { timeout: 5000 });
        const selectResult = await this.page.evaluate((searchText: any) => {
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
    
    if (value) {
      try {
        await this.page.waitForSelector('input[name="Value"], #org-unit-sample-value', { timeout: 5000 });
        await this.page.evaluate((val: any) => {
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
    
    if (description) {
      try {
        await this.page.waitForSelector('textarea[name="Description"], #org-unit-sample-description', { timeout: 3000 });
        await this.page.evaluate((val: any) => {
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
cls.prototype.handleSubgroupsPage = async function(): Promise<void> {
  if (!this.page) return;
  
  console.log('\n   📋 Processing Project Subgroups page...');
  
  try {
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
    const sourceSheet = sourceWorkbook.getWorksheet('P1-OrgScope');
    
    if (!sourceSheet) {
      console.log('   ⚠️  P1-OrgScope sheet not found');
      return;
    }
    
    const getCellValue = (row: number, col: number): string => {
      const cell = sourceSheet.getCell(row, col);
      return this.resolveCellValue(sourceWorkbook, cell.value);
    };
    
    const name = getCellValue(137, 2);
    const abbreviation = getCellValue(138, 2);
    const checkboxInfo = getCellValue(139, 2);
    
    console.log(`   Name: ${name}`);
    console.log(`   Abbreviation: ${abbreviation}`);
    console.log(`   Checkbox: ${checkboxInfo}`);
    
    const existingEntry = await this.page.evaluate((searchName: any, searchAbbr: any) => {
      const cards = document.querySelectorAll('.item-card');
      for (const card of Array.from(cards)) {
        const title = card.querySelector('.item-card__title h3')?.textContent?.trim();
        if (title) {
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
    
    const formVisible = await this.page.$('input[name="Name"]');
    
    if (!formVisible) {
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
    
    if (name) {
      try {
        await this.page.waitForSelector('input[name="Name"], #Name', { timeout: 5000 });
        await this.page.evaluate((val: any) => {
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
    
    if (abbreviation) {
      try {
        await this.page.waitForSelector('input[name="Abbreviation"], #Abbreviation', { timeout: 3000 });
        await this.page.evaluate((val: any) => {
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
    
    try {
      const checkboxResult = await this.page.evaluate((checkboxLabel: any) => {
        const checkboxes = document.querySelectorAll('input[type="checkbox"][id*="Selected"]');
        let checkedCount = 0;
        
        for (const checkbox of Array.from(checkboxes)) {
          const cb = checkbox as HTMLInputElement;
          const label = document.querySelector(`label[for="${cb.id}"]`);
          const labelText = label?.textContent?.trim() || '';
          
          const shouldCheck = checkboxLabel.includes('[x]') 
            ? checkboxLabel.toLowerCase().includes(labelText.toLowerCase())
            : true;
          
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
cls.prototype.handleSubgroupAssignmentPage = async function(): Promise<void> {
  if (!this.page) return;
  
  console.log('\n   📋 Processing Project Subgroup Assignment page...');
  
  try {
    const formVisible = await this.page.$('form.org-unit-project-subgroups__form');
    
    if (!formVisible) {
      console.log('   ⚠️  Subgroup assignment form not found');
      return;
    }
    
    const result = await this.page.evaluate(() => {
      const selects = document.querySelectorAll('select[name*="SubgroupId"]');
      let assignedCount = 0;
      let alreadyAssigned = 0;
      
      for (const select of Array.from(selects)) {
        const sel = select as HTMLSelectElement;
        
        let firstValidOption: HTMLOptionElement | null = null;
        for (const option of Array.from(sel.options)) {
          if (option.value && option.value !== '') {
            firstValidOption = option;
            break;
          }
        }
        
        if (firstValidOption) {
          if (sel.value === firstValidOption.value) {
            alreadyAssigned++;
          } else {
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
cls.prototype.handleOrgProjectAppraisalScopePage = async function(): Promise<void> {
  if (!this.page) return;
  
  console.log('\n   📋 Processing Organizational Support Function PA Exceptions page...');
  
  try {
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
    const sourceSheet = sourceWorkbook.getWorksheet('P1-OrgScope');
    
    if (!sourceSheet) {
      console.log('   ⚠️  P1-OrgScope sheet not found');
      return;
    }
    
    const getCellValue = (row: number, col: number): string => {
      const cell = sourceSheet.getCell(row, col);
      return this.resolveCellValue(sourceWorkbook, cell.value);
    };
    
    const scopeTextToValue: { [key: string]: string } = {
      'in-scope': 'InScope',
      'in-scope (default other projects to out-of-scope)': 'InScopeDefaultOthersToOutOfScope',
      'out-of-scope': 'OutOfScope',
    };
    
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
    
    for (let row = 147; row <= 178; row++) {
      const colA = getCellValue(row, 1);
      const colB = getCellValue(row, 2);
      
      if (!colA && !colB) {
        if (currentPA.practiceArea && currentSF) {
          currentSF.exceptions.push(currentPA as PAException);
          currentPA = {};
        }
        continue;
      }
      
      const colALower = colA.toLowerCase();
      
      if (colALower.includes('select support function') || (colALower === 'select' && colB.startsWith('S'))) {
        if (currentPA.practiceArea && currentSF) {
          currentSF.exceptions.push(currentPA as PAException);
          currentPA = {};
        }
        if (currentSF && currentSF.exceptions.length > 0) {
          supportFunctions.push(currentSF);
        }
        currentSF = { name: colB, exceptions: [] };
      } else if (colALower === 'practice area') {
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
    
    if (currentPA.practiceArea && currentSF) {
      currentSF.exceptions.push(currentPA as PAException);
    }
    if (currentSF && currentSF.exceptions.length > 0) {
      supportFunctions.push(currentSF);
    }
    
    console.log(`   Found ${supportFunctions.length} support functions with exceptions`);
    
    for (const sf of supportFunctions) {
      console.log(`\n   === Processing ${sf.name} ===`);
      console.log(`   PA Exceptions: ${sf.exceptions.length}`);
      for (const exc of sf.exceptions) {
        console.log(`      - ${exc.practiceArea}: ${exc.scope}`);
      }
      
      const selectResult = await this.page.evaluate((sfName: any) => {
        const projectForm = document.querySelector('.organizational-project-appraisal-scope__project-selector form');
        if (!projectForm) return { found: false, error: 'Project selector form not found' };
        
        const select = projectForm.querySelector('select[name="OrgUnitProjectId"]') as HTMLSelectElement;
        if (!select) return { found: false, error: 'OrgUnitProjectId select not found' };
        
        const allOptions = Array.from(select.options).map(o => ({ text: o.text.trim(), value: o.value }));

        // Extract the SF code from the name (e.g. 'S4-QA' -> 'S4', 'S1-CM' -> 'S1')
        const sfCode = sfName.match(/^(S\d+)/i)?.[1]?.toUpperCase() || sfName;

        // Match strategy: exact > contains full name > starts with SF code > contains SF code
        let matchedOption = allOptions.find(o => o.text === sfName)
          || allOptions.find(o => o.text.includes(sfName))
          || allOptions.find(o => o.text.toUpperCase().startsWith(sfCode + '-') || o.text.toUpperCase().startsWith(sfCode + ' '))
          || allOptions.find(o => o.text.toUpperCase().includes(sfCode));

        if (matchedOption) {
          select.value = matchedOption.value;
          select.dispatchEvent(new Event('change', { bubbles: true }));
          return { found: true, value: matchedOption.text, optionValue: matchedOption.value, allOptions };
        }
        return { found: false, error: `Option not found for ${sfName} (code: ${sfCode})`, allOptions };
      }, sf.name);
      
      if (!selectResult.found) {
        console.log(`   ⚠️  ${selectResult.error}`);
        if ((selectResult as any).allOptions?.length) {
          console.log(`   Available dropdown options:`);
          (selectResult as any).allOptions.forEach((o: any) => console.log(`      "${o.text}" (value: ${o.value})`));
        }
        continue;
      }
      
      console.log(`   ✅ Selected: ${selectResult.value} (matched from "${sf.name}")`);
      
      console.log(`   Clicking Select button...`);
      await Promise.all([
        this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 }).catch(() => {}),
        this.page.evaluate(() => {
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
      
      await new Promise(resolve => setTimeout(resolve, 2000));
      
      const formHeader = await this.page.evaluate(() => {
        const h2 = document.querySelector('#Form h2');
        return h2?.textContent?.trim() || '';
      });
      console.log(`   Form header: "${formHeader}"`);
      
      try {
        await this.page.waitForSelector('.project-appraisal-scope-form__practice-area-selector', { timeout: 5000 });
        console.log(`   ✅ PA form loaded`);
      } catch (e) {
        console.log(`   ⚠️  PA form did not load`);
        continue;
      }
      
      console.log(`   Processing ${sf.exceptions.length} PA exceptions...`);
      
      for (const exc of sf.exceptions) {
        const scopeValue = scopeTextToValue[exc.scope.toLowerCase()] || 'InScope';
        console.log(`      Setting PA: "${exc.practiceArea}" to scope: ${scopeValue}`);
        
        const setResult = await this.page.evaluate((paName: any, scopeValue: any, justification: any) => {
          const normalize = (s: string) => s.toLowerCase().replace(/\s+/g, ' ').trim();
          const targetPA = normalize(paName);
          
          const containers = document.querySelectorAll('.project-appraisal-scope-form__practice-area-selector');
          
          for (const container of Array.from(containers)) {
            const labels = container.querySelectorAll('label');
            let matchedLabel: string | null = null;
            
            for (const label of Array.from(labels)) {
              const labelText = label.textContent?.trim() || '';
              if (labelText.toLowerCase().startsWith('justification')) continue;
              
              if (normalize(labelText) === targetPA) {
                matchedLabel = labelText;
                break;
              }
            }
            
            if (matchedLabel) {
              const select = container.querySelector('select[name^="PracticeAreaInclusionStatus"]') as HTMLSelectElement;
              if (!select) {
                return { success: false, pa: paName, error: 'Select element not found in container', matchedLabel };
              }
              
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
          
          const availablePAs: string[] = [];
          for (const container of Array.from(containers)) {
            const labels = container.querySelectorAll('label');
            for (const label of Array.from(labels)) {
              const text = label.textContent?.trim();
              if (text && !text.toLowerCase().startsWith('justification')) {
                availablePAs.push(text);
                break;
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
      
      console.log(`   Clicking Save Exceptions...`);
      await Promise.all([
        this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 30000 }).catch(() => {}),
        this.page.evaluate(() => {
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
}
