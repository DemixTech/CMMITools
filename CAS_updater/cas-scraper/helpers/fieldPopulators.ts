/**
 * Field Populator Helpers - Mixin applied to CASPopulator
 *
 * Contains all populateXxx methods and populateField dispatcher.
 * Applied via: applyFieldPopulators(CASPopulator) in populator.ts
 */

import { CONFIG } from '../config';
import { FieldMapping } from '../types';

export function applyFieldPopulators(cls: any): void {

  cls.prototype.populateField = async function(
    mapping: FieldMapping,
    value: string
  ): Promise<{ success: boolean; changed: boolean; error?: string }> {
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
  };

  cls.prototype.populateTextInput = async function(
    selector: string,
    value: string
  ): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };

    await this.page.waitForSelector(selector, { timeout: 5000 });

    const currentValue = await this.page.evaluate((sel: string) => {
      const el = document.querySelector(sel) as HTMLInputElement | HTMLTextAreaElement;
      return el ? el.value : '';
    }, selector);

    if (currentValue.trim() === value.trim()) {
      console.log(`      ℹ️  Value already set, skipping`);
      return { changed: false };
    }

    await this.page.click(selector);

    await this.page.evaluate((sel: string) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      if (el) el.value = '';
    }, selector);

    await this.page.type(selector, value, { delay: 10 });
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));

    return { changed: true };
  };

  cls.prototype.populateSelect = async function(
    selector: string,
    value: string
  ): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };

    await this.page.waitForSelector(selector, { timeout: 5000 });

    const result = await this.page.evaluate((sel: string, searchText: string) => {
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
      // Pass 3: search text contains option text
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

    if (result.currentText.includes(value) || result.currentText === result.targetText) {
      console.log(`      ℹ️  Value already selected, skipping`);
      return { changed: false };
    }

    await this.page.select(selector, result.targetValue);
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));

    return { changed: true };
  };

  cls.prototype.populateRadio = async function(
    selector: string,
    value: string,
    notes: string
  ): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };

    const options = notes.split(';').map((s: string) => s.trim()).filter((s: string) => s.includes('='));
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
      const selectors = selector.split('|');
      targetSelector = value.toLowerCase() === 'yes' ? selectors[0] : selectors[1];
    }

    await this.page.waitForSelector(targetSelector, { timeout: 5000 });

    const isAlreadyChecked = await this.page.evaluate((sel: string) => {
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
  };

  cls.prototype.populateCheckbox = async function(
    selector: string,
    value: string
  ): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };

    const shouldBeChecked = ['yes', 'true', '1', 'x', 'checked'].includes(value.toLowerCase());

    let elementId: string | null = null;
    if (selector.match(/^#\d+$/)) {
      elementId = selector.substring(1);
      console.log(`      Using getElementById for numeric ID: ${elementId}`);
    }

    const result = await this.page.evaluate(
      (sel: string, elemId: string | null, shouldCheck: boolean) => {
        let el: HTMLInputElement | null = null;

        if (elemId) {
          el = document.getElementById(elemId) as HTMLInputElement;
        } else {
          el = document.querySelector(sel) as HTMLInputElement;
        }

        if (!el) {
          const dataTestSelector = `input[data-test="input-virtual-selection_${elemId || sel.replace('#', '')}"]`;
          el = document.querySelector(dataTestSelector) as HTMLInputElement;
        }

        if (!el) {
          return { success: false, error: `Element not found: ${sel} (id: ${elemId})`, wasChecked: false, nowChecked: false, changed: false };
        }

        const isChecked = el.checked;

        if (shouldCheck === isChecked) {
          return { success: true, wasChecked: isChecked, nowChecked: isChecked, changed: false };
        }

        el.click();
        el.dispatchEvent(new Event('change', { bubbles: true }));

        return { success: true, wasChecked: isChecked, nowChecked: el.checked, changed: true };
      },
      selector, elementId, shouldBeChecked
    );

    if (!result.success) {
      throw new Error((result as any).error || 'Unknown error');
    }

    if (!result.changed) {
      console.log(`      ℹ️  Checkbox already ${result.wasChecked ? 'checked' : 'unchecked'}, skipping`);
    } else {
      console.log(`      Checkbox state: was=${result.wasChecked}, now=${result.nowChecked}`);
    }

    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));

    return { changed: result.changed };
  };

  cls.prototype.populateNumberInput = async function(
    selector: string,
    value: string
  ): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };

    await this.page.waitForSelector(selector, { timeout: 5000 });

    const currentValue = await this.page.evaluate((sel: string) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      return el ? el.value : '';
    }, selector);

    if (currentValue.trim() === value.trim()) {
      console.log(`      ℹ️  Value already set, skipping`);
      return { changed: false };
    }

    await this.page.evaluate((sel: string, val: string) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      if (el) {
        el.value = '';
        el.value = val;
        el.dispatchEvent(new Event('input', { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }
    }, selector, value);

    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));

    return { changed: true };
  };

  cls.prototype.populateDateParts = async function(
    selector: string,
    value: string
  ): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };

    console.log(`      populateDateParts called with selector="${selector}", value="${value}"`);

    const [yearSelector, monthSelector, daySelector] = selector.split(',').map((s: string) => s.trim());
    console.log(`      Parsed selectors: year="${yearSelector}", month="${monthSelector}", day="${daySelector}"`);

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

    const monthValue = parseInt(month, 10).toString();
    const dayValue   = parseInt(day, 10).toString();

    console.log(`      Date parts: Year=${year}, Month=${monthValue}, Day=${dayValue}`);

    let changed = false;

    const elementExists = await this.page.evaluate((sel: string) => {
      return document.querySelector(sel) !== null;
    }, yearSelector);

    if (!elementExists) {
      console.log(`      ERROR: Element not found: ${yearSelector}`);
      throw new Error(`Element not found: ${yearSelector}`);
    }

    const elementInfo = await this.page.evaluate((sel: string) => {
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
      console.log(`      Using number input mode`);

      const setPart = async (sel: string, val: string, label: string) => {
        if (!sel) return;
        console.log(`      Setting ${label} to: ${val}`);
        await this.page.waitForSelector(sel, { timeout: 5000 });
        const r = await this.page.evaluate((s: string, v: string) => {
          const el = document.querySelector(s) as HTMLInputElement;
          if (el) {
            el.value = v;
            el.dispatchEvent(new Event('input',  { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            return { success: true, newValue: el.value };
          }
          return { success: false, newValue: '' };
        }, sel, val);
        console.log(`      ${label} result: ${JSON.stringify(r)}`);
        changed = true;
      };

      await setPart(yearSelector,  year,       'Year');
      await setPart(monthSelector, monthValue, 'Month');
      await setPart(daySelector,   dayValue,   'Day');

    } else {
      console.log(`      Using select dropdown mode`);

      if (yearSelector)  { await this.page.waitForSelector(yearSelector,  { timeout: 5000 }); await this.page.select(yearSelector,  year);       changed = true; }
      if (monthSelector) { await this.page.waitForSelector(monthSelector, { timeout: 5000 }); await this.page.select(monthSelector, monthValue); changed = true; }
      if (daySelector)   { await this.page.waitForSelector(daySelector,   { timeout: 5000 }); await this.page.select(daySelector,   dayValue);   changed = true; }
    }

    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));

    console.log(`      populateDateParts complete, changed=${changed}`);
    return { changed };
  };

  cls.prototype.populateDateInput = async function(
    selector: string,
    value: string
  ): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };

    await this.page.waitForSelector(selector, { timeout: 5000 });

    let formattedDate = value;
    if (value.match(/^\d{4}-\d{2}-\d{2}/)) {
      const parts = value.split('-');
      formattedDate = `${parts[1]}/${parts[2]}/${parts[0]}`;
    }

    const currentValue = await this.page.evaluate((sel: string) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      return el ? el.value : '';
    }, selector);

    if (currentValue === formattedDate) {
      console.log(`      ℹ️  Date already set, skipping`);
      return { changed: false };
    }

    await this.page.evaluate((sel: string, val: string) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      if (el) {
        el.value = '';
        el.value = val;
        el.dispatchEvent(new Event('input',  { bubbles: true }));
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }
    }, selector, formattedDate);

    console.log(`      Set date to: ${formattedDate}`);
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));

    return { changed: true };
  };

  cls.prototype.populateRadioLevel = async function(
    selector: string,
    value: string,
    notes: string
  ): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };

    const normValue = value.trim().replace(/^level\s+/i, '');

    const options = notes.split(',').map((s: string) => s.trim());
    let targetSelector = '';

    for (const opt of options) {
      const [optValue, optSelector] = opt.split('=').map((s: string) => s.trim());
      if (normValue === optValue) {
        targetSelector = optSelector;
        break;
      }
    }

    if (!targetSelector) {
      targetSelector = `#level-${normValue}`;
    }

    console.log(`      Target level selector: ${targetSelector}`);

    await this.page.waitForSelector(targetSelector, { timeout: 5000 });

    const isAlreadyChecked = await this.page.evaluate((sel: string) => {
      const el = document.querySelector(sel) as HTMLInputElement;
      return el ? el.checked : false;
    }, targetSelector);

    if (isAlreadyChecked) {
      console.log(`      ℹ️  Level already selected, skipping`);
      return { changed: false };
    }

    await this.page.click(targetSelector);
    await new Promise(resolve => setTimeout(resolve, CONFIG.waitAfterAction));

    return { changed: true };
  };

  cls.prototype.populateMultiselect = async function(
    selector: string,
    value: string
  ): Promise<{ changed: boolean }> {
    if (!this.page) return { changed: false };

    console.log(`      🔄 Handling React multiselect...`);

    try {
      const containerSelector = selector || '.multiselect';
      await this.page.waitForSelector(containerSelector, { timeout: 5000 });

      const alreadySelected = await this.page.evaluate((searchText: string) => {
        const selectors = [
          '.multiselect__multi-value__label',
          '[class*="multiValue"] [class*="label"]',
          '[class*="multi-value"] [class*="label"]',
          '.css-1rhbuit-multiValue .css-1v99tuv',
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

      await this.page.click(containerSelector);
      console.log(`      Clicked to open dropdown`);
      await new Promise(resolve => setTimeout(resolve, 800));

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

      const clicked = await this.page.evaluate((searchText: string) => {
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

        const availableOptions = options.map(o => (o.textContent || '').trim()).slice(0, 10);

        for (const option of options) {
          const text = (option.textContent || '').trim();
          if (text === searchText) {
            (option as HTMLElement).click();
            return { clicked: true, text, method: 'exact match', selector: usedSelector, availableOptions };
          }
        }

        for (const option of options) {
          const text = (option.textContent || '').trim();
          if (text.includes(searchText) || searchText.includes(text)) {
            (option as HTMLElement).click();
            return { clicked: true, text, method: 'partial match', selector: usedSelector, availableOptions };
          }
        }

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
        console.log(`      ⚠️  Debug: ${(clicked as any).debug}`);
        console.log(`      ⚠️  Could not select multiselect option, continuing`);
        return { changed: false };
      }

    } catch (error) {
      const errorMsg = error instanceof Error ? error.message : String(error);
      console.log(`      ❌ Multiselect error: ${errorMsg}`);
      throw error;
    }
  };
}
