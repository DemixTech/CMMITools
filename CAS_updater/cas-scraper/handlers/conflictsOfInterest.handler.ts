/**
 * Conflicts of Interest Page Handler - Mixin applied to CASPopulator
 *
 * Page: /conflicts-of-interest
 * Workflow:
 *   1. Delete all existing entries (ConfirmDelete pattern)
 *   2. Download the template
 *   3. Populate from C_COI sheet (headers match template)
 *   4. Upload the populated template
 *
 * Source: C_COI sheet in target xlsm
 *   Row 1: headers (6 columns, identical to template)
 *   Row 2+: data rows copied directly (same column order)
 *
 * Columns (A-F):
 *   A: * Conflict of Interests
 *   B: * Other Conflict of Interest
 *   C: * Description
 *   D: * Impact
 *   E: * Management Strategy
 *   F: * Status
 *
 * Applied via: applyCOIHandler(CASPopulator) in populator.ts
 */

import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import AdmZip from 'adm-zip';
import { CONFIG } from '../config';

export function applyCOIHandler(cls: any): void {

cls.prototype.handleConflictsOfInterestPage = async function(): Promise<void> {
  if (!this.page) return;

  console.log('\n   📋 Processing Conflicts of Interest page...');

  const downloadDir = path.join(process.cwd(), 'downloads');
  if (!fs.existsSync(downloadDir)) fs.mkdirSync(downloadDir, { recursive: true });

  try {
    // ── STEP 1: Delete all existing entries ──────────────────────────────
    console.log('\n   === Step 1: Delete Existing Entries ===');

    let safetyLimit = 30;
    while (safetyLimit-- > 0) {
      const deleteHref = await this.page.evaluate(() => {
        const link = document.querySelector(
          '.item-card-list a[href*="handler=ConfirmDelete"], ' +
          '.item-card-list a.red-button[href*="ConfirmDelete"]'
        ) as HTMLAnchorElement | null;
        return link ? link.href : null;
      });

      if (!deleteHref) {
        console.log('   ✅ No more entries to delete');
        break;
      }

      console.log(`   → Deleting: ${deleteHref}`);
      await this.page.goto(deleteHref, { waitUntil: 'networkidle2' });
      await new Promise(r => setTimeout(r, 1000));

      const confirmed = await this.page.evaluate(() => {
        const btns = Array.from(document.querySelectorAll('button, input[type="submit"]'));
        for (const btn of btns) {
          const text = (btn.textContent || (btn as HTMLInputElement).value || '').toLowerCase();
          if (text.includes('delete') || text.includes('yes') || text.includes('confirm')) {
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
        console.log('   ⚠️  Could not confirm delete - navigating back');
        await this.page.goto(
          `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}/conflicts-of-interest`,
          { waitUntil: 'networkidle2' }
        );
        await new Promise(r => setTimeout(r, 1000));
        break;
      }

      console.log('   ✅ Entry deleted');

      const currentUrl = this.page.url();
      if (!currentUrl.includes('conflicts-of-interest')) {
        await this.page.goto(
          `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}/conflicts-of-interest`,
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

    // ── STEP 3: Load data from C_COI ─────────────────────────────────────
    console.log('\n   === Step 3: Load C_COI Data ===');

    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
    const sourceSheet = sourceWorkbook.getWorksheet('C_COI');

    if (!sourceSheet) {
      console.log('   ⚠️  C_COI sheet not found in source Excel');
      return;
    }

    const getCellValue = (cell: ExcelJS.Cell): any => {
      return this.resolveCellValue(sourceWorkbook, cell.value);
    };

    // Row 1 = headers; data starts at row 2. 6 columns (A-F). Stop at first fully empty row.
    interface COIRow {
      conflict: string;
      otherConflict: string;
      description: string;
      impact: string;
      managementStrategy: string;
      status: string;
    }

    const dataRows: COIRow[] = [];

    for (let rowNum = 2; rowNum <= sourceSheet.rowCount + 5; rowNum++) {
      const vals: string[] = [];
      let allEmpty = true;
      for (let col = 1; col <= 6; col++) {
        const v = getCellValue(sourceSheet.getCell(rowNum, col));
        const s = (v != null && String(v) !== 'undefined' && String(v) !== 'nan') ? String(v).trim() : '';
        vals.push(s);
        if (s) allEmpty = false;
      }

      if (allEmpty) break;

      dataRows.push({
        conflict:           vals[0],
        otherConflict:      vals[1],
        description:        vals[2],
        impact:             vals[3],
        managementStrategy: vals[4],
        status:             vals[5],
      });
      console.log(`      Row ${rowNum}: "${vals[0].substring(0, 45)}" [${vals[3]}] [${vals[5]}]`);
    }

    console.log(`   ✅ Loaded ${dataRows.length} COI entries`);

    if (dataRows.length === 0) {
      console.log('   ⚠️  No data in C_COI - skipping upload');
      return;
    }

    // ── STEP 4: Populate the template ────────────────────────────────────
    console.log('\n   === Step 4: Populate Template ===');

    // Inject data into downloaded template using Python XML manipulation
    // Preserves portal's styles, dropdown validation, and hidden lookup sheet
    const populatedFileName = 'populated_' + (downloadResult.filename || 'template.xlsx');
    const populatedPath = path.join(downloadDir, populatedFileName);
    const injectScriptPath = path.join(__dirname, '..', 'helpers', 'injectTemplate.py');
    const { execSync } = require('child_process');

    const colConfig = JSON.stringify([["A","conflict","ss"],["B","otherConflict","inline"],["C","description","inline"],["D","impact","ss"],["E","managementStrategy","inline"],["F","status","ss"]]);
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

    console.log('\n   ✅ Conflicts of Interest page processing complete');

  } catch (error) {
    console.log(`   ⚠️  Error processing conflicts of interest: ${error}`);
    if (error instanceof Error) console.log(`   Stack: ${error.stack}`);
  }
}

} // end applyCOIHandler
