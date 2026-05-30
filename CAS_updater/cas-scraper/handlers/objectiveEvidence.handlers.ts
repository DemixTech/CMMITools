/**
 * Objective Evidence Page Handlers - Mixin applied to CASPopulator
 *
 * Covers all /objective-evidence/* pages:
 *   collection-approach, collection-techniques, collection-responsibilities,
 *   performance-report-approaches, initial-summary,
 *   data-collection-timing, additional-info
 *
 * Also contains fillAndSaveSimpleForm() helper used by OE handlers.
 *
 * Applied via: applyOEHandlers(CASPopulator) in populator.ts
 */

import { CONFIG } from '../config';

export function applyOEHandlers(cls: any): void {
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
cls.prototype.fillAndSaveSimpleForm = async function(fields: Array<{ selector: string; type: 'select' | 'text' | 'textarea' | 'date'; value: string }>): Promise<void> {
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
cls.prototype.handleOECollectionApproachPage = async function(): Promise<void> {
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
cls.prototype.handleOECollectionTechniquesPage = async function(): Promise<void> {
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
  const exists = await this.page.evaluate((techName: any) => {
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
cls.prototype.handleOECollectionResponsibilitiesPage = async function(): Promise<void> {
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
cls.prototype.handlePerformanceReportApproachesPage = async function(): Promise<void> {
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
cls.prototype.handleInitialSummaryPage = async function(): Promise<void> {
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
cls.prototype.handleDataCollectionTimingPage = async function(): Promise<void> {
  if (!this.page) return;
  console.log('\n   📋 Processing Data Collection Timing page...');

  const appraisalId = CONFIG.appraisalId;
  const baseUrl     = CONFIG.casBaseUrl;
  const d           = this.excelData['P1PA-R'] || {};

  const milestones = [
    { name: d[75], date: d[76], participants: d[77] },
    { name: d[79], date: d[80], participants: d[81] },
  ].filter(m => m.name);

  console.log(`   Milestones to create: ${milestones.length}`);
  if (milestones.length === 0) {
    console.log('   ⚠️  No milestone data found in Excel - skipping');
    return;
  }

  // ── STEP 1: Delete all existing entries ──────────────────────────────
  console.log('\n   🗑️  Clearing existing entries...');
  let safetyLimit = 10;
  while (safetyLimit-- > 0) {
    const deleteHref = await this.page.evaluate(() => {
      const link = document.querySelector(
        '.item-card-list a.red-button[href*="ConfirmDelete"], ' +
        '.item-card-list a[href*="handler=ConfirmDelete"]'
      ) as HTMLAnchorElement | null;
      return link ? link.href : null;
    });

    if (!deleteHref) {
      console.log('   ✅ No more entries to delete');
      break;
    }

    console.log(`   → Clicking Delete: ${deleteHref}`);
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
      console.log('   ⚠️  Could not confirm delete - navigating back to timing page');
      await this.page.goto(
        `${baseUrl}/appraisals/${appraisalId}/objective-evidence/data-collection-timing`,
        { waitUntil: 'networkidle2' }
      );
      await new Promise(r => setTimeout(r, 1000));
      break;
    }

    console.log('   ✅ Entry deleted');

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

    const currentUrl = this.page.url();
    if (!currentUrl.includes('data-collection-timing')) {
      await this.page.goto(
        `${baseUrl}/appraisals/${appraisalId}/objective-evidence/data-collection-timing`,
        { waitUntil: 'networkidle2' }
      );
      await new Promise(r => setTimeout(r, 1000));
    }

    await this.page.evaluate(() => {
      document.getElementById('Form')?.scrollIntoView({ block: 'center' });
    });
    await new Promise(r => setTimeout(r, 500));

    try {
      await this.page.waitForSelector('form.data-collection-timing-form #Name', { timeout: 8000 });
    } catch (e) {
      console.log(`   ⚠️  Form not found on page`);
      continue;
    }

    try {
      await this.page.evaluate(() => {
        const el = document.querySelector('form.data-collection-timing-form #Name') as HTMLInputElement | null;
        if (el) { el.value = ''; }
      });
      await this.populateTextInput('form.data-collection-timing-form #Name', m.name);
      console.log(`      ✅ Name: ${m.name}`);
    } catch (e) {
      console.log(`      ⚠️  Name error: ${e}`);
    }

    if (m.date) {
      try {
        await this.populateDateParts(
          'form.data-collection-timing-form #CompletedYear,' +
          'form.data-collection-timing-form #CompletedMonth,' +
          'form.data-collection-timing-form #CompletedDay',
          m.date
        );
        console.log(`      ✅ Date: ${m.date}`);
      } catch (e) {
        console.log(`      ⚠️  Date error: ${e}`);
      }
    }

    if (m.participants) {
      try {
        await this.populateTextInput(
          'form.data-collection-timing-form #ParticipantNamesListing',
          m.participants
        );
        console.log(`      ✅ Participants: ${m.participants.substring(0, 60)}`);
      } catch (e) {
        console.log(`      ⚠️  Participants error: ${e}`);
      }
    }

    const submitted = await this.page.evaluate(() => {
      const form = document.querySelector('form.data-collection-timing-form') as HTMLFormElement | null;
      if (!form) return false;
      const btn = form.querySelector('button') as HTMLButtonElement | null;
      if (btn) { btn.click(); return true; }
      form.submit();
      return true;
    });

    if (!submitted) {
      console.log(`      ⚠️  Could not submit form`);
      continue;
    }

    await this.page.waitForNavigation({ waitUntil: 'networkidle2', timeout: 15000 }).catch(() => {});
    await new Promise(r => setTimeout(r, 1500));
    console.log(`   ✅ Milestone ${i + 1} saved`);

    const urlAfter = this.page.url();
    if (!urlAfter.includes('data-collection-timing')) {
      console.log(`   ⚠️  Unexpected redirect to: ${urlAfter} - navigating back`);
      await this.page.goto(
        `${baseUrl}/appraisals/${appraisalId}/objective-evidence/data-collection-timing`,
        { waitUntil: 'networkidle2' }
      );
      await new Promise(r => setTimeout(r, 1000));
    }
  }

  console.log('\n   ✅ Data Collection Timing complete');
}

// /objective-evidence/initial-summary (duplicate handler for alternate page mapping)
// Row 72 of P1PA-R
cls.prototype.handleOEInitialSummaryPage = async function(): Promise<void> {
  if (!this.page) return;
  console.log('\n   📋 Processing OE Initial Summary page...');

  const d = this.excelData['P1PA-R'] || {};
  const summary = d[72];

  if (!summary) {
    console.log('   ⚠️  No summary value in Excel data (row 72) - skipping');
    return;
  }

  console.log(`   Summary: ${summary.substring(0, 80)}...`);

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

// /objective-evidence/additional-info
// Row 86 of P1PA-R
cls.prototype.handleOEAdditionalInfoPage = async function(): Promise<void> {
  if (!this.page) return;
  console.log('\n   📋 Processing OE Additional Info page...');

  const d    = this.excelData['P1PA-R'] || {};
  const info = d[86];

  console.log(`   Additional Info: ${info?.substring(0, 60)}`);

  if (!info) {
    console.log('   ℹ️  No additional info value in Excel (row 86) - skipping');
    return;
  }

  const existing = await this.page.evaluate(() => {
    const ta = document.querySelector('#AdditionalInformation') as HTMLTextAreaElement | null;
    return ta ? ta.value.trim() : '';
  });

  const norm = (s: string) => (s || '').trim().replace(/\s+/g, ' ');

  if (norm(existing) === norm(info)) {
    console.log('   ℹ️  Already matches - skipping update');
    return;
  }

  console.log(existing ? '   ⚠️  Existing value differs - updating' : '   ℹ️  No existing value - populating');

  await this.fillAndSaveSimpleForm([
    { selector: '#AdditionalInformation', type: 'textarea', value: info },
  ]);

  console.log('   ✅ OE Additional Info saved');
}

}
