/**
 * D&D 5e Item SystemAdapter — generates D&D 5e items from a short description.
 *
 * Output lands directly in the world's Items directory (game.items).
 */

import { SystemAdapter, postToN8n }       from './core/adapter.js';
import { N8N_BASE, devUrl }               from './core/n8n.js';
import { detectModuleFolder }             from './core/utils.js';
import { sanitizeItemDataDnd5e,
         tryFixItemValidationErrorDnd5e } from './sanitizer.js';

const MODULE_FOLDER = detectModuleFolder('Dnd5eItemGenerator');
const ITEM_ENDPOINT = `${N8N_BASE}/webhook/dnd5e-item-builder`;

const RARITY_LABELS = {
  '':        'Mundane',
  common:    'Common',
  uncommon:  'Uncommon',
  rare:      'Rare',
  veryRare:  'Very Rare',
  legendary: 'Legendary',
  artifact:  'Artifact',
};

const ITEM_TYPE_LABELS = {
  weapon:     'Weapon',
  equipment:  'Armor / Equipment',
  consumable: 'Consumable',
  tool:       'Tool',
  loot:       'Loot / Treasure',
  container:  'Container',
};

export class Dnd5eItemAdapter extends SystemAdapter {
  get moduleFolder() { return MODULE_FOLDER; }

  get module() {
    return {
      id:           'Dnd5eItemGenerator',
      label:        'D&D 5e Item',
      icon:         'fa-solid fa-hat-wizard',
      githubUrl:    'https://github.com/JamesCfer/D-Ditemmaker',
      historyLabel: 'Created Items',
    };
  }

  get systemId() { return 'dnd5e'; }

  get supportsImageGeneration() { return false; }

  get formConfig() { return { documentNoun: 'item' }; }

  /* ── Form handling ──────────────────────────────────────── */

  gatherFormData(form) {
    const fd = new FormData(form);
    const name        = (fd.get('name')?.toString()?.trim()) || 'Generated Item';
    const rarity      = (fd.get('rarity')?.toString() ?? '').trim();
    const itemType    = (fd.get('itemType')?.toString() || 'loot').trim();
    const subtype     = (fd.get('subtype')?.toString()?.trim()) || '';
    const description = (fd.get('description')?.toString()?.trim()) || '';

    if (!description) throw new Error('Please provide a description for the item.');
    return { name, rarity, itemType, subtype, description };
  }

  historyEntryFromForm(formData) {
    return {
      name:        formData.name,
      rarity:      formData.rarity,
      itemType:    formData.itemType,
      subtype:     formData.subtype,
      description: formData.description,
    };
  }

  historyMeta(entry) {
    const rarityLabel = RARITY_LABELS[entry.rarity] ?? 'Mundane';
    const typeLabel   = ITEM_TYPE_LABELS[entry.itemType] || 'Item';
    return `${rarityLabel}&nbsp;·&nbsp;${typeLabel}`;
  }

  populateForm(form, entry) {
    const nameInput    = form.querySelector('[name="name"]');
    const raritySelect = form.querySelector('[name="rarity"]');
    const typeSelect   = form.querySelector('[name="itemType"]');
    const subtypeInput = form.querySelector('[name="subtype"]');
    const descTextarea = form.querySelector('[name="description"]');
    if (nameInput)    nameInput.value    = entry.name ?? '';
    if (raritySelect) raritySelect.value = entry.rarity ?? '';
    if (typeSelect)   typeSelect.value   = entry.itemType ?? 'loot';
    if (subtypeInput) subtypeInput.value = entry.subtype ?? '';
    if (descTextarea) descTextarea.value = entry.description ?? '';
  }

  /* ── Generation ─────────────────────────────────────────── */

  async generate({ formData, key, devMode }) {
    const endpoint = devUrl(ITEM_ENDPOINT, devMode);
    const payload  = {
      name:        formData.name,
      rarity:      formData.rarity,
      itemType:    formData.itemType,
      subtype:     formData.subtype,
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

    sanitizeItemDataDnd5e(itemData, formData.itemType, formData.rarity);

    let item, attempts = 0;
    const maxAttempts = 10;
    while (!item && attempts < maxAttempts) {
      attempts++;
      try {
        item = await Item.create(itemData);
      } catch (error) {
        const errorText = error.toString ? error.toString() : String(error.message || error);
        if (tryFixItemValidationErrorDnd5e(itemData, errorText)) continue;
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
