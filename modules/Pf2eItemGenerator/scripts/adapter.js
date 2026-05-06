/**
 * PF2e Item SystemAdapter — generates PF2e items from a short description.
 *
 * Output lands directly in the world's Items directory (game.items).
 *
 * @typedef {object} Pf2eItemFormData
 * @property {string} name        Item name.
 * @property {number} level       Item level (0–20).
 * @property {string} itemType    One of the ITEM_TYPE_LABELS keys.
 * @property {string} description Free-text description for the AI.
 */

import { SystemAdapter, postToN8n } from './core/adapter.js';
import { N8N_BASE, devUrl }          from './core/n8n.js';
import { detectModuleFolder }        from './core/utils.js';
import { sanitizeItemDataPf2e,
         tryFixItemValidationError } from './sanitizer.js';

const MODULE_FOLDER = detectModuleFolder('Pf2eItemGenerator');
const ITEM_ENDPOINT = `${N8N_BASE}/webhook/pf2e-item-builder`;

const ITEM_TYPE_LABELS = {
  weapon:     'Weapon',
  armor:      'Armor',
  shield:     'Shield',
  equipment:  'Equipment',
  consumable: 'Consumable',
  treasure:   'Treasure',
  backpack:   'Container',
  book:       'Book',
};

export class Pf2eItemAdapter extends SystemAdapter {
  get moduleFolder() { return MODULE_FOLDER; }

  get module() {
    return {
      id:           'Pf2eItemGenerator',
      label:        'PF2e Item',
      icon:         'fa-solid fa-flask',
      githubUrl:    'https://github.com/JamesCfer/Pf2eItemMaker',
      historyLabel: 'Created Items',
    };
  }

  get systemId() { return 'pf2e'; }

  // Item generator does not use image generation.
  get supportsImageGeneration() { return false; }

  get formConfig() { return { documentNoun: 'item' }; }

  /* ── Form handling ──────────────────────────────────────── */

  /** @returns {Pf2eItemFormData} */
  gatherFormData(form) {
    const fd = new FormData(form);
    const name        = (fd.get('name')?.toString()?.trim()) || 'Generated Item';
    const level       = Number(fd.get('level')) || 0;
    const itemType    = (fd.get('itemType')?.toString() || 'equipment').trim();
    const description = (fd.get('description')?.toString()?.trim()) || '';

    if (!description) throw new Error('Please provide a description for the item.');
    return { name, level, itemType, description };
  }

  historyEntryFromForm(formData) {
    return {
      name:        formData.name,
      level:       formData.level,
      itemType:    formData.itemType,
      description: formData.description,
    };
  }

  historyMeta(entry) {
    const typeLabel = ITEM_TYPE_LABELS[entry.itemType] || 'Item';
    return `Lv.&nbsp;${entry.level}&nbsp;·&nbsp;${typeLabel}`;
  }

  populateForm(form, entry) {
    const nameInput    = form.querySelector('[name="name"]');
    const levelInput   = form.querySelector('[name="level"]');
    const typeSelect   = form.querySelector('[name="itemType"]');
    const descTextarea = form.querySelector('[name="description"]');
    if (nameInput)    nameInput.value    = entry.name ?? '';
    if (levelInput)   levelInput.value   = entry.level ?? 0;
    if (typeSelect)   typeSelect.value   = entry.itemType ?? 'equipment';
    if (descTextarea) descTextarea.value = entry.description ?? '';
  }

  /* ── Generation ─────────────────────────────────────────── */

  /**
   * @param {import('./core/adapter.js').GenerateOptions & { formData: Pf2eItemFormData }} opts
   * @returns {Promise<import('./core/adapter.js').AdapterResult>}
   */
  async generate({ formData, key, devMode }) {
    const endpoint = devUrl(ITEM_ENDPOINT, devMode);
    const payload  = {
      name:        formData.name,
      level:       formData.level,
      itemType:    formData.itemType,
      description: formData.description,
    };

    const { response, responseText } = await postToN8n(endpoint, payload, key);

    let data;
    try {
      data = JSON.parse(responseText);
    } catch (err) {
      throw new Error(`Invalid JSON response (${responseText.length} bytes): ${err.message}`);
    }

    if (!response.ok) throw new Error(data?.message || `Server returned status ${response.status}`);
    if (data?.ok === false) throw new Error(data?.message || data?.error || 'Server rejected the request');

    const itemData = data.foundryItem || data.item || data;
    if (!itemData || typeof itemData !== 'object') throw new Error('No valid item data returned from server');

    sanitizeItemDataPf2e(itemData, formData.itemType);

    let item, attempts = 0;
    const maxAttempts = 10;
    while (!item && attempts < maxAttempts) {
      attempts++;
      try {
        item = await Item.create(itemData);
      } catch (error) {
        const errorText = error.toString ? error.toString() : String(error.message || error);
        if (tryFixItemValidationError(itemData, errorText)) continue;
        throw error;
      }
    }
    if (!item) throw new Error('Failed to create item after maximum retry attempts');

    return {
      document:   item,
      exportData: {
        content:  JSON.stringify(itemData, null, 2),
        filename: `${item.name || 'item'}.json`,
        mimeType: 'application/json',
      },
      message: `Item "${item.name}" created successfully!`,
    };
  }
}
