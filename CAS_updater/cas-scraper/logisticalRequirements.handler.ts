/**
 * Logistical Requirements Page Handler - Mixin applied to CASPopulator
 *
 * Page: /logistical-requirements
 * Workflow:
 *   1. Delete all existing entries (click Delete → confirm → repeat)
 *   2. Download the template
 *   3. Populate from P1PA-R logistical requirement blocks (rows 117+)
 *   4. Upload the populated template
 *
 * Source: P1PA-R sheet, repeating 5-row blocks starting at row 117:
 *   Row +0: Logistical requirement        → Template col A
 *   Row +1: Logistical requirement Equipment → Template col B
 *   Row +2: Other logistical requirement  → Template col C
 *   Row +3: Description of the requirement → Template col D
 *   Row +4: Role responsible for requirement → Template col E
 *
 * Applied via: applyLogisticalRequirementsHandler(CASPopulator) in populator.ts
 */

import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import AdmZip from 'adm-zip';
import { CONFIG } from '../config';

export function applyLogisticalRequirementsHandler(cls: any): void {

cls.prototype.handleLogisticalRequirementsPage = async function(): Promise<void> {
  if (!this.page) return;

  console.log('\n   📋 Processing Logistical Requirements page...');

  const downloadDir = path.join(process.cwd(), 'downloads');
  if (!fs.existsSync(downloadDir)) fs.mkdirSync(downloadDir, { recursive: true });

  try {
    // ── STEP 1: Delete all existing entries ──────────────────────────────
    console.log('\n   === Step 1: Delete Existing Entries ===');

    let safetyLimit = 20;
    while (safetyLimit-- > 0) {
      // Find the first ConfirmDelete link in the item-card-list
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

      // Confirm the deletion (click Delete/Yes/Confirm button or submit form)
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
          `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}/logistical-requirements`,
          { waitUntil: 'networkidle2' }
        );
        await new Promise(r => setTimeout(r, 1000));
        break;
      }

      console.log('   ✅ Entry deleted');

      // After confirm, server redirects back to the page; ensure we're there
      const currentUrl = this.page.url();
      if (!currentUrl.includes('logistical-requirements')) {
        await this.page.goto(
          `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}/logistical-requirements`,
          { waitUntil: 'networkidle2' }
        );
        await new Promise(r => setTimeout(r, 1000));
      }
    }

    // ── STEP 2: Download the template ────────────────────────────────────
    console.log('\n   === Step 2: Download Template ===');

    // Clear old xlsx files from download directory
    const oldFiles = fs.readdirSync(downloadDir);
    for (const file of oldFiles) {
      if (file.endsWith('.xlsx') || file.endsWith('.xls')) {
        fs.unlinkSync(path.join(downloadDir, file));
        console.log(`   🗑️  Deleted old file: ${file}`);
      }
    }

    // Find the download link
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

    // Fetch within authenticated browser context (preserves cookies/auth)
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
    }, downloadUrl) as { success: boolean; base64?: string; filename?: string; error?: string };

    if (!downloadResult.success || !downloadResult.base64) {
      console.log(`   ⚠️  Download failed: ${downloadResult.error}`);
      return;
    }

    const templatePath = path.join(downloadDir, downloadResult.filename || 'template.xlsx');
    const buffer = Buffer.from(downloadResult.base64, 'base64');
    fs.writeFileSync(templatePath, buffer);
    console.log(`   ✅ Template downloaded: ${downloadResult.filename} (${buffer.length} bytes)`);

    // ── STEP 3: Load logistical requirement blocks from P1PA-R ───────────
    console.log('\n   === Step 3: Load Logistical Requirements Data ===');

    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
    const sourceSheet = sourceWorkbook.getWorksheet('P1PA-R');

    if (!sourceSheet) {
      console.log('   ⚠️  P1PA-R sheet not found in source Excel');
      return;
    }

    const getCellValue = (cell: ExcelJS.Cell): any => {
      return this.resolveCellValue(sourceWorkbook, cell.value);
    };

    // Parse repeating 5-row blocks starting at row 117
    // Each block: [0]=Logistical requirement, [1]=Equipment, [2]=Other, [3]=Description, [4]=Role
    // Blocks are separated by one empty row
    interface LogReqBlock {
      requirement: string;
      equipment: string;
      other: string;
      description: string;
      role: string;
    }

    const blocks: LogReqBlock[] = [];
    let rowNum = 117;
    const MAX_ROWS = 200; // safety cap

    while (rowNum <= Math.min(sourceSheet.rowCount + 10, 117 + MAX_ROWS)) {
      const cellA = getCellValue(sourceSheet.getCell(rowNum, 1));
      const valA = cellA != null ? String(cellA).trim() : '';

      // Block starts when col A = 'Logistical requirement'
      if (valA === 'Logistical requirement') {
        const getVal = (r: number) => {
          const v = getCellValue(sourceSheet.getCell(r, 2));
          return (v != null && String(v) !== 'undefined' && String(v) !== 'nan') ? String(v).trim() : '';
        };

        const block: LogReqBlock = {
          requirement: getVal(rowNum),
          equipment:   getVal(rowNum + 1),
          other:       getVal(rowNum + 2),
          description: getVal(rowNum + 3),
          role:        getVal(rowNum + 4),
        };

        // Skip if all fields empty (template placeholder block)
        if (!block.requirement && !block.description && !block.role) {
          rowNum += 6;
          continue;
        }

        blocks.push(block);
        console.log(`      Block ${blocks.length}: "${block.requirement.substring(0, 50)}"`);
        rowNum += 6; // 5 data rows + 1 empty separator
      } else if (valA === 'COPY AND REPEAT ABOVE FOR ADDITIONAL' || valA === 'Appraisal Constraints') {
        // End of logistical requirements section
        break;
      } else {
        rowNum++;
      }
    }

    console.log(`   ✅ Loaded ${blocks.length} logistical requirement entries`);

    if (blocks.length === 0) {
      console.log('   ⚠️  No logistical requirement data found - skipping upload');
      return;
    }

    // ── STEP 4: Populate the template ────────────────────────────────────
    console.log('\n   === Step 4: Populate Template ===');

    // Extract headers from downloaded template using AdmZip
    let templateHeaders: string[] = [];
    try {
      const zip = new AdmZip(templatePath);
      const sharedStringsXml = zip.readAsText('xl/sharedStrings.xml');
      const stringMatches = sharedStringsXml.match(/<(?:x:)?t[^>]*>([^<]*)<\/(?:x:)?t>/g) || [];
      templateHeaders = stringMatches.slice(0, 5).map((m: any) => {
        const match = m.match(/<(?:x:)?t[^>]*>([^<]*)<\/(?:x:)?t>/);
        return match ? match[1].replace(/\r\n/g, ' ').trim() : '';
      });
      console.log(`   Extracted ${templateHeaders.length} headers from template: ${templateHeaders.join(' | ')}`);
    } catch (e) {
      console.log(`   ⚠️  Could not extract template headers: ${e}`);
      templateHeaders = [
        '* Logistical requirement',
        '* Logistical requirement Equipment',
        '* Other logistical requirement',
        '* Description of the requirement',
        '* Role responsible for requirement',
      ];
    }

    // Create new workbook matching the template structure
    const outWorkbook = new ExcelJS.Workbook();
    const outSheet = outWorkbook.addWorksheet('Logistical Requirements');

    // Write header row
    const headerRow = outSheet.getRow(1);
    templateHeaders.forEach((h, i) => { headerRow.getCell(i + 1).value = h; });
    headerRow.font = { bold: true };

    // Write data rows - headings are the same in source and destination
    // Col A: Logistical requirement
    // Col B: Logistical requirement Equipment
    // Col C: Other logistical requirement
    // Col D: Description of the requirement
    // Col E: Role responsible for requirement
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      const row = outSheet.getRow(i + 2);
      row.getCell(1).value = b.requirement;
      row.getCell(2).value = b.equipment;
      row.getCell(3).value = b.other;
      row.getCell(4).value = b.description;
      row.getCell(5).value = b.role;
      console.log(`      Row ${i + 2}: "${b.requirement.substring(0, 45)}"`);
    }

    console.log(`   ✅ Populated ${blocks.length} rows`);

    // Save populated template
    const populatedFileName = 'populated_' + (downloadResult.filename || 'template.xlsx');
    const populatedPath = path.join(downloadDir, populatedFileName);
    await outWorkbook.xlsx.writeFile(populatedPath);
    console.log(`   ✅ Saved: ${populatedFileName}`);

    // ── STEP 5: Upload the populated template ─────────────────────────────
    console.log('\n   === Step 5: Upload Populated Template ===');

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

    console.log('\n   ✅ Logistical Requirements page processing complete');

  } catch (error) {
    console.log(`   ⚠️  Error processing logistical requirements: ${error}`);
    if (error instanceof Error) console.log(`   Stack: ${error.stack}`);
  }
}

} // end applyLogisticalRequirementsHandler
