/**
 * Test script to debug template download on /org-projects page
 * First finds the correct URL by navigating through CAS
 */

import puppeteer from 'puppeteer';
import * as fs from 'fs';
import * as path from 'path';

const projectConfigPath = path.resolve(__dirname, '..', 'cas-project-config.json');
const projectConfig = JSON.parse(fs.readFileSync(projectConfigPath, 'utf-8'));

const CONFIG = {
  casBaseUrl: projectConfig?.cas?.baseUrl || 'https://cas.cmmiinstitute.com',
  appraisalId: projectConfig?.project?.casId || '81846',
};

async function testDownload() {
  const downloadDir = path.join(process.cwd(), 'downloads');
  
  // Ensure download directory exists
  if (!fs.existsSync(downloadDir)) {
    fs.mkdirSync(downloadDir, { recursive: true });
  }
  
  console.log(`Download directory: ${downloadDir}`);

  // Use persistent user data directory to preserve authentication
  const browser = await puppeteer.launch({
    headless: false,
    defaultViewport: null,
    userDataDir: path.join(process.cwd(), '.browser-data'),
    args: ['--start-maximized']
  });

  const page = await browser.newPage();

  // Start from the main appraisal page
  const baseUrl = `${CONFIG.casBaseUrl}/appraisals/${CONFIG.appraisalId}`;
  console.log(`\nNavigating to base appraisal URL: ${baseUrl}`);
  await page.goto(baseUrl, { waitUntil: 'networkidle2' });
  await new Promise(resolve => setTimeout(resolve, 2000));

  // Check if we're authenticated
  let currentUrl = page.url();
  console.log(`Current URL: ${currentUrl}`);
  
  if (currentUrl.includes('login')) {
    console.log('\n❌ Not authenticated! Please run the main populator.ts first to login.');
    console.log('Browser left open for manual login...');
    return;
  }
  console.log('✅ Authenticated session detected');

  // Find all navigation links to understand the CAS structure
  console.log('\n=== Finding navigation links ===');
  const navLinks = await page.evaluate(() => {
    return Array.from(document.querySelectorAll('a')).map(a => ({
      href: a.getAttribute('href') || '',
      text: a.textContent?.trim().substring(0, 60) || '',
    })).filter(l => l.href && l.href.includes('/appraisals/'));
  });

  console.log(`Found ${navLinks.length} appraisal-related links:`);
  navLinks.forEach((l, i) => {
    console.log(`  [${i}] ${l.text} -> ${l.href}`);
  });

  // Look for org-unit-projects or similar links
  const orgProjectLinks = navLinks.filter(l => 
    l.href.includes('org-unit-project') || 
    l.href.includes('org-project') ||
    l.href.includes('projects') ||
    l.text.toLowerCase().includes('project')
  );

  console.log(`\n=== Project-related links (${orgProjectLinks.length}) ===`);
  orgProjectLinks.forEach((l, i) => {
    console.log(`  [${i}] ${l.text} -> ${l.href}`);
  });

  // Try to find download template links anywhere on this page
  const downloadLinks = navLinks.filter(l =>
    l.href.includes('download') || l.href.includes('template')
  );

  console.log(`\n=== Download/Template links (${downloadLinks.length}) ===`);
  downloadLinks.forEach((l, i) => {
    console.log(`  [${i}] ${l.text} -> ${l.href}`);
  });

  // If we found a project-related link, navigate to it
  if (orgProjectLinks.length > 0) {
    const projectLink = orgProjectLinks[0];
    const projectUrl = projectLink.href.startsWith('/') 
      ? CONFIG.casBaseUrl + projectLink.href 
      : projectLink.href;
    
    console.log(`\n📍 Navigating to: ${projectUrl}`);
    await page.goto(projectUrl, { waitUntil: 'networkidle2' });
    await new Promise(resolve => setTimeout(resolve, 2000));
    
    currentUrl = page.url();
    console.log(`Now at: ${currentUrl}`);

    // Save page HTML
    const html = await page.content();
    fs.writeFileSync(path.join(downloadDir, 'project_page.html'), html);
    console.log('📄 Saved page HTML to downloads/project_page.html');

    // Find download links on this page
    const pageDownloadLinks = await page.evaluate(() => {
      return Array.from(document.querySelectorAll('a')).map(a => ({
        href: a.getAttribute('href') || '',
        fullHref: a.href,
        text: a.textContent?.trim() || '',
        className: a.className || ''
      })).filter(l => 
        l.href.includes('download') || 
        l.href.includes('template') ||
        l.text.toLowerCase().includes('download') ||
        l.text.toLowerCase().includes('template')
      );
    });

    console.log(`\n=== Download links on project page (${pageDownloadLinks.length}) ===`);
    pageDownloadLinks.forEach((l, i) => {
      console.log(`  [${i}] "${l.text}"`);
      console.log(`       href: ${l.href}`);
      console.log(`       fullHref: ${l.fullHref}`);
    });

    if (pageDownloadLinks.length > 0) {
      // Attempt download
      const downloadUrl = pageDownloadLinks[0].fullHref;
      console.log(`\n📥 Attempting download from: ${downloadUrl}`);
      
      const downloadResult = await page.evaluate(async (url) => {
        try {
          const response = await fetch(url, {
            method: 'GET',
            credentials: 'include'
          });
          
          if (!response.ok) {
            return { 
              success: false, 
              error: `HTTP ${response.status}: ${response.statusText}`
            };
          }
          
          const contentDisposition = response.headers.get('Content-Disposition');
          let filename = 'template.xlsx';
          if (contentDisposition) {
            const match = contentDisposition.match(/filename[^;=\n]*=(["']?)([^"';\n]*)/i);
            if (match && match[2]) {
              filename = match[2];
            }
          }
          
          const contentType = response.headers.get('Content-Type');
          const blob = await response.blob();
          const reader = new FileReader();
          
          return new Promise((resolve) => {
            reader.onloadend = () => {
              const base64 = (reader.result as string).split(',')[1];
              resolve({ 
                success: true, 
                base64, 
                filename, 
                contentType,
                size: blob.size 
              });
            };
            reader.onerror = () => {
              resolve({ success: false, error: 'FileReader error' });
            };
            reader.readAsDataURL(blob);
          });
        } catch (e) {
          return { success: false, error: String(e) };
        }
      }, downloadUrl) as any;

      if (downloadResult.success && downloadResult.base64) {
        const filePath = path.join(downloadDir, downloadResult.filename || 'template.xlsx');
        const buffer = Buffer.from(downloadResult.base64, 'base64');
        fs.writeFileSync(filePath, buffer);
        console.log(`\n✅ Downloaded: ${filePath} (${buffer.length} bytes)`);
      } else {
        console.log(`\n❌ Download failed: ${downloadResult.error}`);
      }
    }
  }

  console.log('\n✅ Test complete. Browser left open for inspection.');
}

testDownload().catch(err => {
  console.error('Test failed with error:', err);
});
