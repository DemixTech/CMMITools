/**
 * Resource Estimates Page Handler - Mixin applied to CASPopulator
 *
 * Page: /resource-estimates
 * Workflow: download template → populate from C_Resource_Estimates rows 24-N → upload
 *
 * Source columns A,B,C,D → Template columns B,C,D,E (col A = Database ID, left blank)
 *
 * Applied via: applyResourceEstimatesHandler(CASPopulator) in populator.ts
 */

import * as fs from 'fs';
import * as path from 'path';
import * as ExcelJS from 'exceljs';
import AdmZip from 'adm-zip';
import { CONFIG } from '../config';

export function applyResourceEstimatesHandler(cls: any): void {

cls.prototype.handleResourceEstimatesPage = async function(): Promise<void> {
  if (!this.page) return;

  console.log('\n   📋 Processing Resource Estimates page (template download/upload workflow)...');

  const downloadDir = path.join(process.cwd(), 'downloads');
  if (!fs.existsSync(downloadDir)) fs.mkdirSync(downloadDir, { recursive: true });

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
      console.log('   ⚠️  Could not find download template link');
      console.log(`   🔍 All links on page (${downloadUrl.allLinks.length} total):`);
      downloadUrl.allLinks.filter((l: any) => l.href).forEach((l: any) => console.log(`      "${l.text}" -> ${l.href}`));
      return;
    }

    console.log(`   📥 Download URL: ${downloadUrl.url}`);

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
        let filename = 'template.xlsx';
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
    }, downloadUrl.url) as { success: boolean; base64?: string; filename?: string; error?: string };

    if (!downloadResult.success || !downloadResult.base64) {
      console.log(`   ⚠️  Download failed: ${downloadResult.error}`);
      return;
    }

    // Save the file from base64
    const templatePath = path.join(downloadDir, downloadResult.filename || 'template.xlsx');
    const buffer = Buffer.from(downloadResult.base64, 'base64');
    fs.writeFileSync(templatePath, buffer);
    console.log(`   ✅ Template downloaded: ${downloadResult.filename} (${buffer.length} bytes)`);

    // Step 2: Load C_Resource_Estimates data from CAS Plan Excel
    console.log('\n   === Step 2: Load C_Resource_Estimates Data ===');

    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
    const sourceSheet = sourceWorkbook.getWorksheet('C_Resource_Estimates');

    if (!sourceSheet) {
      console.log('   ⚠️  C_Resource_Estimates sheet not found in source Excel');
      return;
    }

    // Helper function to extract cell value (handles formulas with formula resolution)
    const getCellValue = (cell: ExcelJS.Cell): any => {
      return this.resolveCellValue(sourceWorkbook, cell.value);
    };

    // Data starts at row 24, columns A-D
    // Find where data ends (look for empty Column A - Participant Group)
    let lastDataRow = 23;
    for (let row = 24; row <= sourceSheet.rowCount + 10; row++) {
      const cellA = sourceSheet.getCell(row, 1);
      const valueA = getCellValue(cellA);

      if (valueA && String(valueA).trim() && String(valueA).trim() !== 'undefined') {
        lastDataRow = row;
      } else {
        break;
      }
    }

    console.log(`   ✅ C_Resource_Estimates data rows: 24 to ${lastDataRow} (${lastDataRow - 23} entries)`);

    // Read source data (columns A-D = columns 1-4)
    interface RowData {
      [col: number]: any;
    }
    const sourceData: RowData[] = [];
    for (let row = 24; row <= lastDataRow; row++) {
      const rowData: RowData = {};
      for (let col = 1; col <= 4; col++) {
        const cell = sourceSheet.getCell(row, col);
        const value = getCellValue(cell);
        if (value !== null && value !== undefined && String(value) !== 'undefined') {
          rowData[col] = value;
        }
      }
      const groupName = rowData[1] || '(unnamed)';
      console.log(`      Row ${row}: ${String(groupName).substring(0, 40)}`);
      sourceData.push(rowData);
    }

    console.log(`   ✅ Loaded ${sourceData.length} resource estimate entries`);

    // Step 3: Populate template by injecting data into downloaded template XML
    console.log('\n   === Step 3: Populate Template ===');

    // Inject data into downloaded template using Python XML manipulation
    // Preserves portal's styles, dropdown validation, and hidden lookup sheet
    const populatedFileName = 'populated_' + (downloadResult.filename || 'template.xlsx');
    const populatedTemplatePath = path.join(downloadDir, populatedFileName);
    const injectScriptPath = path.join(__dirname, '..', 'helpers', 'injectTemplate.py');
    const { execSync } = require('child_process');

    // Col A=dbId(blank), B=group, C=individuals, D=task, E=hours
    const colConfig = JSON.stringify([["A","dbId","inline"],["B","group","inline"],["C","individuals","inline"],["D","task","inline"],["E","hours","inline"]]);
    const rowsData  = JSON.stringify(sourceData.map((r: any) => ({
      dbId: '',
      group:       String(r[1] || ''),
      individuals: String(r[2] || ''),
      task:        String(r[3] || ''),
      hours:       String(r[4] || ''),
    })));

    const pyOut = execSync(
      `python "${injectScriptPath}" "${templatePath}" "${populatedTemplatePath}" ${JSON.stringify(colConfig)} ${JSON.stringify(rowsData)}`,
      { encoding: 'utf8', timeout: 30000 }
    ).trim();
    if (!pyOut.startsWith('ok:')) throw new Error(`Template inject failed: ${pyOut}`);
    console.log(`   ✅ Populated template: ${populatedFileName} (${pyOut.split(':')[1]} rows)`);

    // Step 4: Upload the populated template
    console.log('\n   === Step 4: Upload Populated Template ===');

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

    console.log('\n   ✅ Resource Estimates page processing complete');

  } catch (error) {
    console.log(`   ⚠️  Error processing resource estimates: ${error}`);
    if (error instanceof Error) {
      console.log(`   Stack: ${error.stack}`);
    }
  }
}

} // end applyResourceEstimatesHandler
