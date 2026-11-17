/**
 * Stores (D&D 5e and PF2e) → one store handout with the shopkeeper and a stock
 * table, plus a give recipe per item so the GM can hand a purchase to a PC.
 */
import { makePlan } from '../plan.js';
import { escapeHtml, num, titleCase } from '../util.js';
import { dnd5eItemFacts, pf2eItemFacts, giveRecipe, itemCardHtml } from './items.js';

const LABELS = {
  general: 'General Store', blacksmith: 'Blacksmith', armorer: 'Armorer', alchemist: 'Alchemist', magic: 'Magic Shop',
  jeweler: 'Jeweler', fletcher: 'Fletcher', tavern: 'Tavern',
};

export function mapStore(system) {
  return (response, form = {}) => {
    const store = response.store || response.foundryStore || response;
    const inventory = Array.isArray(store.inventory) ? store.inventory
      : Array.isArray(store.items) ? store.items
      : Array.isArray(response.foundryItems) ? response.foundryItems : [];
    const name = store.name || form.name || 'Store';
    const owner = store.owner || store.shopkeeper || null;
    const staff = Array.isArray(store.staff) ? store.staff : [];
    const items = inventory.filter(i => i && typeof i === 'object' && i.name);

    const meta = [LABELS[store.storeType || form.storeType] || 'Store', titleCase(store.wealth || form.wealth || ''), titleCase(store.settlementSize || form.settlementSize || '')]
      .filter(Boolean).map(escapeHtml).join(' · ');
    const rows = items.map(item => {
      const f = system === 'pf2e' ? pf2eItemFacts(item) : dnd5eItemFacts(item);
      const second = system === 'pf2e' ? String(f.level) : (f.rarity || '-');
      return `<tr><td>${escapeHtml(item.name)}</td><td>${escapeHtml(second)}</td><td style="text-align:right">${escapeHtml(f.price || '-')}</td><td style="text-align:right">${num(item.system?.quantity, 1)}</td></tr>`;
    }).join('');
    const person = p => `<p><strong>${escapeHtml(p.name || 'Unnamed')}</strong>${p.race || p.ancestry ? ` (${escapeHtml(p.race || p.ancestry)})` : ''}${p.role ? `, ${escapeHtml(p.role)}` : ''}</p>${p.description ? `<p>${escapeHtml(p.description)}</p>` : ''}`;

    const notes = `<p><strong>${meta}</strong></p>`
      + `<p>${escapeHtml(store.description || form.description || '')}</p>`
      + (owner ? `<h3>Shopkeeper</h3>${person(owner)}` : '')
      + (staff.length ? `<h3>Staff</h3>${staff.map(person).join('')}` : '')
      + `<h3>Stock</h3><table><thead><tr><th>Item</th><th>${system === 'pf2e' ? 'Level' : 'Rarity'}</th><th>Price</th><th>Qty</th></tr></thead><tbody>${rows}</tbody></table>`
      + '<p><em>Item details are in the handouts in this folder. Use the Cfer chat menu or panel to give an item to a character.</em></p>';

    return makePlan('handouts', `${system}-store`, {
      summary: `${name} (${items.length} items)`,
      folder: name,
      handouts: [
        { name, notes, gmnotes: '' },
        ...items.map(item => ({ name: `${name}: ${item.name}`, notes: itemCardHtml(item, system), give: giveRecipe(item, system) })),
      ],
    });
  };
}
