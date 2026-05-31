// Handle the org-projects page - downloads template, populates from C_SupportV2, uploads
// ADD THIS FUNCTION after handleReadinessReviewsPage() in populator.ts
// AND ADD the call in processCurrentPage():
//   if (pagePath === '/org-projects') {
//     await this.handleOrgProjectsPage();
//     return log;
//   }

async handleOrgProjectsPage(): Promise<void> {
  if (!this.page) return;
  
  console.log('\n   📋 Processing Org Projects page (template download/upload workflow)...');
  
  const downloadDir = path.join(process.cwd(), 'downloads');
  
  // Ensure download directory exists
  if (!fs.existsSync(downloadDir)) {
    fs.mkdirSync(downloadDir, { recursive: true });
  }
  
  try {
    // Step 1: Download the template
    console.log('\n   === Step 1: Download Template ===');
    
    // Configure download behavior
    const client = await this.page.target().createCDPSession();
    await client.send('Page.setDownloadBehavior', {
      behavior: 'allow',
      downloadPath: downloadDir
    });
    
    // Clear old files from download directory
    const oldFiles = fs.readdirSync(downloadDir);
    for (const file of oldFiles) {
      if (file.endsWith('.xlsx') || file.endsWith('.xls')) {
        fs.unlinkSync(path.join(downloadDir, file));
      }
    }
    
    // Find and click the download template link
    const downloadClicked = await this.page.evaluate(() => {
      const links = Array.from(document.querySelectorAll('a'));
      for (const link of links) {
        const href = link.getAttribute('href') || '';
        const text = link.textContent || '';
        if (href.includes('download') || href.includes('template') || href.includes('Template') ||
            text.toLowerCase().includes('download') || text.toLowerCase().includes('template')) {
          link.click();
          return { clicked: true, href: href, text: text.trim() };
        }
      }
      return { clicked: false, href: '', text: '' };
    });
    
    if (!downloadClicked.clicked) {
      console.log('   ⚠️  Could not find download template link');
      return;
    }
    
    console.log(`   ✅ Clicked download link: ${downloadClicked.text}`);
    
    // Wait for download to complete
    console.log('   ⏳ Waiting for download...');
    await new Promise(resolve => setTimeout(resolve, 5000));
    
    // Find the downloaded file
    const files = fs.readdirSync(downloadDir);
    const templateFile = files.find(f => f.endsWith('.xlsx') || f.endsWith('.xls'));
    
    if (!templateFile) {
      console.log('   ⚠️  Downloaded template file not found in: ' + downloadDir);
      console.log('   Files in directory: ' + files.join(', '));
      return;
    }
    
    const templatePath = path.join(downloadDir, templateFile);
    console.log(`   ✅ Template downloaded: ${templateFile}`);
    
    // Step 2: Load C_SupportV2 data from CAS Plan Excel
    console.log('\n   === Step 2: Load C_SupportV2 Data ===');
    
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.readFile(CONFIG.excelFile);
    const sourceSheet = sourceWorkbook.getWorksheet('C_SupportV2');
    
    if (!sourceSheet) {
      console.log('   ⚠️  C_SupportV2 sheet not found in source Excel');
      return;
    }
    
    // Find where Column B data ends
    let lastDataRow = 1;
    for (let row = 2; row <= sourceSheet.rowCount; row++) {
      const cellB = sourceSheet.getCell(row, 2).value; // Column B
      if (cellB && String(cellB).trim()) {
        lastDataRow = row;
      } else {
        break;
      }
    }
    
    console.log(`   ✅ C_SupportV2 data rows: 2 to ${lastDataRow} (${lastDataRow - 1} functions)`);
    
    // Read source data
    interface RowData {
      [col: number]: any;
    }
    const sourceData: RowData[] = [];
    for (let row = 2; row <= lastDataRow; row++) {
      const rowData: RowData = {};
      for (let col = 1; col <= sourceSheet.columnCount; col++) {
        const cellValue = sourceSheet.getCell(row, col).value;
        rowData[col] = cellValue;
      }
      sourceData.push(rowData);
      console.log(`      Row ${row}: ${rowData[2]} (Function Name)`);
    }
    
    console.log(`   ✅ Loaded ${sourceData.length} support functions from C_SupportV2`);
    
    // Step 3: Populate the template
    console.log('\n   === Step 3: Populate Template ===');
    
    const templateWorkbook = new ExcelJS.Workbook();
    await templateWorkbook.xlsx.readFile(templatePath);
    const templateSheet = templateWorkbook.worksheets[0]; // First sheet
    
    if (!templateSheet) {
      console.log('   ⚠️  No worksheet found in template');
      return;
    }
    
    console.log(`   Template sheet: ${templateSheet.name}`);
    
    // Build a map of Column B values to row numbers in template
    const templateMap: { [key: string]: number } = {};
    for (let row = 2; row <= templateSheet.rowCount; row++) {
      const cellB = templateSheet.getCell(row, 2).value; // Column B
      if (cellB && String(cellB).trim()) {
        templateMap[String(cellB).trim()] = row;
      }
    }
    
    console.log(`   Template has ${Object.keys(templateMap).length} existing entries in Column B`);
    
    // Copy data from source to template matching on Column B
    let updatedCount = 0;
    let addedCount = 0;
    let nextEmptyRow = 2;
    
    // Find first empty row
    while (templateSheet.getCell(nextEmptyRow, 2).value) {
      nextEmptyRow++;
    }
    
    for (const srcRow of sourceData) {
      const functionName = String(srcRow[2] || '').trim(); // Column B = Function Name
      
      if (!functionName) continue;
      
      let targetRow: number;
      
      if (templateMap[functionName]) {
        // Update existing row
        targetRow = templateMap[functionName];
        updatedCount++;
        console.log(`      Updating row ${targetRow}: ${functionName}`);
      } else {
        // Add to next empty row
        targetRow = nextEmptyRow;
        nextEmptyRow++;
        addedCount++;
        console.log(`      Adding new row ${targetRow}: ${functionName}`);
      }
      
      // Copy all columns from source to template
      for (let col = 1; col <= Object.keys(srcRow).length; col++) {
        if (srcRow[col] !== undefined && srcRow[col] !== null) {
          templateSheet.getCell(targetRow, col).value = srcRow[col];
        }
      }
      
      // Set Column K to "Yes" for data rows
      templateSheet.getCell(targetRow, 11).value = 'Yes'; // Column K
    }
    
    // Clear Column K "Yes" values beyond the last data row
    // The last data row is row 2 + sourceData.length - 1 = sourceData.length + 1
    const lastTemplateDataRow = 1 + sourceData.length;
    for (let row = lastTemplateDataRow + 1; row <= templateSheet.rowCount; row++) {
      const cellK = templateSheet.getCell(row, 11);
      if (cellK.value === 'Yes' || cellK.value === 'yes') {
        cellK.value = '';
        console.log(`      Cleared Column K 'Yes' from row ${row}`);
      }
    }
    
    console.log(`   ✅ Updated: ${updatedCount}, Added: ${addedCount}`);
    
    // Save the populated template
    const populatedTemplatePath = path.join(downloadDir, 'populated_' + templateFile);
    await templateWorkbook.xlsx.writeFile(populatedTemplatePath);
    console.log(`   ✅ Saved populated template: populated_${templateFile}`);
    
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
            return true;
          }
        }
        return false;
      });
      
      if (uploadClicked) {
        console.log('   ✅ Upload button clicked');
        await new Promise(resolve => setTimeout(resolve, 5000));
      }
    } else {
      console.log('   ⚠️  File input not found - template may need manual upload');
      console.log(`   📁 Populated template saved at: ${populatedTemplatePath}`);
    }
    
    console.log('\n   ✅ Org Projects page processing complete');
    
  } catch (error) {
    console.log(`   ⚠️  Error processing org-projects: ${error}`);
  }
}
