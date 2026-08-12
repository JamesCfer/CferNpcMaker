/**
 * PF2e Store SystemAdapter — generates a single shop from a short description.
 *
 * The store lands as a JournalEntry carrying the payload under
 * flags.Pf2eStoreGenerator.store, rendered by StoreSheet.
 *
 * @typedef {object} StoreFormData
 * @property {string} name         Shop name.
 * @property {string} storeType    One of STORE_TYPES.
 * @property {number} level        Party/settlement level the stock is pitched at.
 * @property {string} priceTier    low | standard | high | luxury.
 * @property {number} stockCount   How many inventory lines to ask for.
 * @property {string} settlement   Free-text home settlement.
 * @property {string} description  Free-text description for the AI.
 */

import { SystemAdapter, postToN8n }        from './core/adapter.js';
import { N8N_BASE, devUrl }                from './core/n8n.js';
import { detectModuleFolder }              from './core/utils.js';
import { MODULE_ID, FLAG_SCOPE, FLAG_KEY } from './constants.js';
import { sanitizeStore, storeTypeLabel,
         STORE_TYPES, PRICE_TIERS }        from './sanitizer.js';
import { StoreSheet }                      from './store-sheet.js';

const MODULE_FOLDER  = detectModuleFolder(MODULE_ID);
const STORE_ENDPOINT = `${N8N_BASE}/webhook/pf2e-store-builder`;

/** Monthly uses charged for one store. */
export const STORE_COST = 7;

export class Pf2eStoreAdapter extends SystemAdapter {
  get moduleFolder() { return MODULE_FOLDER; }

  get module() {
    return {
      id:           MODULE_ID,
      label:        'PF2e Store',
      icon:         'fa-solid fa-shop',
      githubUrl:    'https://github.com/JamesCfer/Pf2eStoreGenerator',
      historyLabel: 'Created Stores',
    };
  }

  get systemId() { return 'pf2e'; }

  get supportsImageGeneration() { return false; }

  get formConfig() { return { documentNoun: 'store' }; }

  get generationCost() { return STORE_COST; }

  get progressSteps() {
    return ['Sending request…', 'Stocking the shelves…', 'Hiring staff…', 'Creating journal…'];
  }

  /* ── Form handling ──────────────────────────────────────── */

  /** @returns {StoreFormData} */
  gatherFormData(form) {
    const fd = new FormData(form);
    const name        = (fd.get('name')?.toString()?.trim()) || 'New Shop';
    const storeTypeIn = (fd.get('storeType')?.toString() || 'general').trim();
    const priceTierIn = (fd.get('priceTier')?.toString() || 'standard').trim();
    const level       = Number(fd.get('level')) || 1;
    const stockCount  = Number(fd.get('stockCount')) || 10;
    const settlement  = (fd.get('settlement')?.toString()?.trim()) || '';
    const description = (fd.get('description')?.toString()?.trim()) || '';

    if (!description) throw new Error('Please provide a description for the store.');

    return {
      name,
      storeType:  STORE_TYPES.includes(storeTypeIn) ? storeTypeIn : 'general',
      priceTier:  PRICE_TIERS.includes(priceTierIn) ? priceTierIn : 'standard',
      level:      Math.min(30, Math.max(0, level)),
      stockCount: Math.min(40, Math.max(1, stockCount)),
      settlement,
      description,
    };
  }

  historyEntryFromForm(formData) {
    return {
      name:        formData.name,
      storeType:   formData.storeType,
      priceTier:   formData.priceTier,
      level:       formData.level,
      stockCount:  formData.stockCount,
      settlement:  formData.settlement,
      description: formData.description,
    };
  }

  historyMeta(entry) {
    const label = storeTypeLabel(entry.storeType);
    return `${label}&nbsp;·&nbsp;${entry.stockCount ?? 0}&nbsp;goods`;
  }

  populateForm(form, entry) {
    const set = (sel, val) => { const el = form.querySelector(sel); if (el != null && val != null) el.value = val; };
    set('[name="name"]',        entry.name);
    set('[name="storeType"]',   entry.storeType);
    set('[name="priceTier"]',   entry.priceTier);
    set('[name="level"]',       entry.level);
    set('[name="stockCount"]',  entry.stockCount);
    set('[name="settlement"]',  entry.settlement);
    set('[name="description"]', entry.description);
  }

  /* ── Generation ─────────────────────────────────────────── */

  /**
   * @param {import('./core/adapter.js').GenerateOptions & { formData: StoreFormData }} opts
   * @returns {Promise<import('./core/adapter.js').AdapterResult>}
   */
  async generate({ formData, key, devMode }) {
    const endpoint = devUrl(STORE_ENDPOINT, devMode);
    const payload  = {
      system:      'pf2e',
      name:        formData.name,
      storeType:   formData.storeType,
      priceTier:   formData.priceTier,
      level:       formData.level,
      stockCount:  formData.stockCount,
      settlement:  formData.settlement,
      description: formData.description,
    };

    const { response, responseText } = await postToN8n(endpoint, payload, key, this.moduleFolder);

    let data;
    try {
      data = JSON.parse(responseText);
    } catch (err) {
      throw new Error(`Invalid JSON response (${responseText.length} bytes): ${err.message}`);
    }

    if (!response.ok)     throw new Error(data?.message || `Server returned status ${response.status}`);
    if (data?.ok === false) throw new Error(data?.message || data?.error || 'Server rejected the request');

    const store = sanitizeStore(data, formData);

    const journal = await JournalEntry.create({
      name:  store.name || formData.name,
      flags: { [FLAG_SCOPE]: { [FLAG_KEY]: store } },
    });
    if (!journal) throw new Error('Foundry rejected the store journal entry.');

    try { new StoreSheet(journal).render(true); } catch (_) {}

    return {
      document:   journal,
      exportData: {
        content:  JSON.stringify(store, null, 2),
        filename: `${store.name || 'store'}.json`,
        mimeType: 'application/json',
      },
      message: `Store "${store.name}" created with ${store.inventory.length} goods!`,
    };
  }

  /** The store opens its own sheet, so skip the shared quick-edit dialog. */
  quickEditFields() { return null; }

  /* ── Cost hint ──────────────────────────────────────────── */

  onFormMount(form) {
    const hintEl = form.querySelector('#store-cost-hint');
    if (!hintEl) return;

    const updateHint = () => {
      const count = Number(form.querySelector('[name="stockCount"]')?.value) || 0;
      hintEl.textContent = `${STORE_COST} Patreon uses · ${count} goods on the shelves`;
    };

    form.querySelector('[name="stockCount"]')?.addEventListener('input', updateHint);
    updateHint();
  }
}
