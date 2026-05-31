/**
 * Appraisal Constraints Page Handler - Mixin applied to CASPopulator
 *
 * Page: /appraisal-constraints
 * Workflow:
 *   1. Delete all existing entries (ConfirmRemove pattern)
 *   2. Download the template
 *   3. Populate from C_AppraisalConstraints sheet (headers match template)
 *   4. Upload the populated template
 *
 * Source: C_AppraisalConstraints sheet in target xlsm
 *   Row 1: headers (* Constraint | * Other Constraint | * Constraint Description)
 *   Row 2+: data rows copied directly to template (same column order)
 *
 * Applied via: applyAppraisalConstraintsHandler(CASPopulator) in populator.ts
 */

import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import AdmZip from 'adm-zip';
import { CONFIG } from '../config';

export function applyAppraisalConstraintsHandler(cls: any): void {

cls.prototype.handleAppraisalConstraintsPage = async function(): Promise<void> {
  if (!this.page) return;

  console.log('\n   📋 Processing Appraisal Constraints page...');

  const downloadDir = path.join(process.cwd(), 'downloads');
  if (!fs.existsSync(downloadDir)) fs.mkdirSync(downloadDir, { recursive: true });

  try {
    // ── STEP 1: Delete all existing entries ──────────────────────────────
    // Note: this page uses "ConfirmRemove" not "ConfirmDelete"
    console.log('\n   === Step 1: Delete Existing Entries ===');

    let safetyLimit = 20;
    while (safetyLimit-- > 0) {
      const deleteHref = await this.page.evaluate(() => {
        const link = document.querySelector(
          '.item-card-list a[href*="handler=ConfirmRemove"], ' +
          '.item-card-list a.red-button[href*="ConfirmRemove"], ' +
          '.item-card-list a[href*="handler=ConfirmDelete"], ' +
          '.item-card-list a.red-button[href*="ConfirmDelete"]'
        ) as HTMLAnchorElement | null;
        return link ? link.href : null;
      });

      if (!deleteHref) {
        console.log('   ✅ No more entries to delete');
        break;
      }

      console.log(`   → Removing: ${deleteHref}`);
      await this.page.goto(deleteHref, { waitUntil: 'networkidle2' });
      await new Promise(r => setTimeout(r, 1000));

      // Confirm (Remove/Delete/Yes/Confirm button, or submit form)
      const confirmed = await this.page.evaluate(() => {
        const btns = Array.from(document.querySelectorAll('button, input[type="submit"]'));
        for (const btn of btns) {
          const text = (btn.textContent || (btn as HTMLInputElement).value || '').toLowerCase();
          if (text.includes('remove') || text.includes('delete') || text.includes('yes') || text.includes('confirm')) {
            (btn as HTMLElement).click();
            return true;
          }
        }
        const form = document.querySelector('form') as HTMLFormElement | null;
        if (form) { form.submit(); return true; }
        return false;
      });

      await this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }).catch(() => {});
      await new Promise(r => setTimeout(r, 1000));

      if (!confirmed) {
        console.log('   ⚠️  Could not confirm removal - navigating back');
        await this.page.goto(
          `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}/appraisal-constraints`,
          { waitUntil: 'networkidle2' }
        );
        await new Promise(r => setTimeout(r, 1000));
        break;
      }

      console.log('   ✅ Entry removed');

      const currentUrl = this.page.url();
      if (!currentUrl.includes('appraisal-constraints')) {
        await this.page.goto(
          `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}/appraisal-constraints`,
          { waitUntil: 'networkidle2' }
        );
        await new Promise(r => setTimeout(r, 1000));
      }
    }

    // ── STEP 2: Download the template ────────────────────────────────────
    console.log('\n   === Step 2: Download Template ===');

    const oldFiles = fs.readdirSync(downloadDir);
    for (const file of oldFiles) {
      if (file.endsWith('.xlsx') || file.endsWith('.xls')) {
        fs.unlinkSync(path.join(downloadDir, file));
        console.log(`   🗑️  Deleted old file: ${file}`);
      }
    }

    const downloadUrl = await this.page.evaluate(() => {
      const links = Array.from(document.querySelectorAll('a'));
      const allLinks = links.map(a => ({ href: a.getAttribute('href') || '', text: (a.textContent || '').trim().substring(0, 60) }));
      for (const link of links) {
        const href = link.getAttribute('href') || '';
        const text = link.textContent || '';
        if (href.includes('/template/download') ||
            (href.includes('download') && (text.toLowerCase().includes('template') || href.includes('Template')))) {
          return { url: (link as HTMLAnchorElement).href, allLinks };
        }
      }
      return { url: null, allLinks };
    });

    if (!downloadUrl.url) {
      console.log('   ⚠️  Could not find template download link');
      console.log(`   🔍 All links on page (${downloadUrl.allLinks.length} total):`);
      downloadUrl.allLinks.filter((l: any) => l.href).forEach((l: any) => console.log(`      "${l.text}" -> ${l.href}`));
      return;
    }

    console.log(`   📥 Download URL: ${downloadUrl.url}`);

    const downloadResult = await this.page.evaluate(async (url: any) => {
      try {
        const response = await fetch(url, { method: 'GET', credentials: 'include' });
        if (!response.ok) {
          return { success: false, error: `HTTP ${response.status}: ${response.statusText}` };
        }

        const contentDisposition = response.headers.get('Content-Disposition');
        let filename = 'template.xlsx';
        if (contentDisposition) {
          const match = contentDisposition.match(/filename[^;=\n]*=(["']?)([^"';\n]*)/i);
          if (match && match[2]) filename = match[2];
        }

        const blob = await response.blob();
        const reader = new FileReader();
        return new Promise((resolve) => {
          reader.onloadend = () => {
            const base64 = (reader.result as string).split(',')[1];
            resolve({ success: true, base64, filename });
          };
          reader.onerror = () => resolve({ success: false, error: 'FileReader error' });
          reader.readAsDataURL(blob);
        });
      } catch (e) {
        return { success: false, error: String(e) };
      }
    }, downloadUrl.url) as { success: boolean; base64?: string; filename?: string; error?: string };

    if (!downloadResult.success || !downloadResult.base64) {
      console.log(`   ⚠️  Download failed: ${downloadResult.error}`);
      return;
    }

    const templatePath = path.join(downloadDir, downloadResult.filename || 'template.xlsx');
    const buffer = Buffer.from(downloadResult.base64, 'base64');
    fs.writeFileSync(templatePath, buffer);
    console.log(`   ✅ Template downloaded: ${downloadResult.filename} (${buffer.length} bytes)`);

    // ── STEP 3: Load data from C_AppraisalConstraints ────────────────────
    console.log('\n   === Step 3: Load C_AppraisalConstraints Data ===');

    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
    const sourceSheet = sourceWorkbook.getWorksheet('C_AppraisalConstraints');

    if (!sourceSheet) {
      console.log('   ⚠️  C_AppraisalConstraints sheet not found in source Excel');
      return;
    }

    const getCellValue = (cell: ExcelJS.Cell): any => {
      return this.resolveCellValue(sourceWorkbook, cell.value);
    };

    // Row 1 = headers; data starts at row 2. Read until first fully empty row.
    // Cols: A=Constraint, B=Other Constraint, C=Constraint Description
    interface ConstraintRow {
      constraint: string;
      otherConstraint: string;
      description: string;
    }

    const dataRows: ConstraintRow[] = [];

    for (let rowNum = 2; rowNum <= sourceSheet.rowCount + 5; rowNum++) {
      const a = getCellValue(sourceSheet.getCell(rowNum, 1));
      const b = getCellValue(sourceSheet.getCell(rowNum, 2));
      const c = getCellValue(sourceSheet.getCell(rowNum, 3));

      const valA = (a != null && String(a) !== 'undefined') ? String(a).trim() : '';
      const valB = (b != null && String(b) !== 'undefined') ? String(b).trim() : '';
      const valC = (c != null && String(c) !== 'undefined') ? String(c).trim() : '';

      // Stop at first fully empty row
      if (!valA && !valB && !valC) break;

      dataRows.push({ constraint: valA, otherConstraint: valB, description: valC });
      console.log(`      Row ${rowNum}: "${valA.substring(0, 40)}" | "${valC.substring(0, 40)}"`);
    }

    console.log(`   ✅ Loaded ${dataRows.length} constraint entries`);

    if (dataRows.length === 0) {
      console.log('   ⚠️  No data in C_AppraisalConstraints - skipping upload');
      return;
    }

    // ── STEP 4: Populate the template ────────────────────────────────────
    console.log('\n   === Step 4: Populate Template ===');

    // Extract headers from downloaded template using AdmZip
    // Inject data into downloaded template using Python XML manipulation
    // Preserves portal's styles, dropdown validation, and hidden lookup sheet
    const populatedFileName = 'populated_' + (downloadResult.filename || 'template.xlsx');
    const populatedPath = path.join(downloadDir, populatedFileName);
    const injectScriptPath = path.join(__dirname, '..', 'helpers', 'injectTemplate.py');
    const { execSync } = require('child_process');

    const colConfig = JSON.stringify([["A","constraint","ss"],["B","otherConstraint","inline"],["C","description","inline"]]);
    const rowsData  = JSON.stringify(dataRows);
    const pyOut = execSync(
      `python "${injectScriptPath}" "${templatePath}" "${populatedPath}" ${JSON.stringify(colConfig)} ${JSON.stringify(rowsData)}`,
      { encoding: 'utf8', timeout: 30000 }
    ).trim();
    if (!pyOut.startsWith('ok:')) throw new Error(`Template inject failed: ${pyOut}`);
    console.log(`   ✅ Populated template: ${populatedFileName} (${pyOut.split(':')[1]} rows)`);

    // ── STEP 5: Upload the populated template ─────────────────────────────
    console.log('\n   === Step 5: Upload Populated Template ===');

    // Log all inputs and buttons on page to understand upload mechanism
    const pageUploadInfo = await this.page.evaluate(() => {
      const inputs = Array.from(document.querySelectorAll('input')).map(i => ({
        type: i.type, name: i.name, id: i.id, accept: i.getAttribute('accept')
      }));
      const buttons = Array.from(document.querySelectorAll('button, input[type="submit"]')).map(b => ({
        text: (b.textContent || (b as HTMLInputElement).value || '').trim().substring(0, 50),
        type: (b as HTMLInputElement).type || 'button'
      }));
      return { inputs, buttons };
    });
    console.log(`   🔍 Inputs on page: ${JSON.stringify(pageUploadInfo.inputs)}`);
    console.log(`   🔍 Buttons on page: ${JSON.stringify(pageUploadInfo.buttons)}`);

    const fileInput = await this.page.$('input[type="file"]');

    if (fileInput) {
      await fileInput.uploadFile(populatedPath);
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
      console.log(`   📁 Populated template saved at: ${populatedPath}`);
    }

    console.log('\n   ✅ Appraisal Constraints page processing complete');

  } catch (error) {
    console.log(`   ⚠️  Error processing appraisal constraints: ${error}`);
    if (error instanceof Error) console.log(`   Stack: ${error.stack}`);
  }
}

} // end applyAppraisalConstraintsHandler
