/**
 * StoreSheet — custom sheet for a generated shop.
 *
 * The store payload lives in flags.<MODULE_ID>.store on a JournalEntry;
 * every mutation goes through doc.setFlag and the sheet re-renders.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY, getStore, setStore } from './constants.js';
import { sanitizeStore, storeTypeLabel, tieredPrice,
         inventoryValue, PRICE_TIERS, STORE_TYPES }             from './sanitizer.js';
import { escapeHtml }                                           from './core/utils.js';

const { HandlebarsApplicationMixin, ApplicationV2 } = foundry.applications.api;

export class StoreSheet extends HandlebarsApplicationMixin(ApplicationV2) {
  static DEFAULT_OPTIONS = {
    id: 'store-sheet-{id}',
    classes: ['cfer-store-sheet'],
    tag: 'form',
    window: { resizable: true },
    position: { width: 760, height: 640 },
    actions: {
      saveField:    function(ev) { this._onSaveField(ev); },
      addItem:      function()   { this._onAddItem(); },
      removeItem:   function(ev) { this._onRemoveItem(ev); },
      addStaff:     function()   { this._onAddStaff(); },
      removeStaff:  function(ev) { this._onRemoveStaff(ev); },
      toggleClosed: function()   { this._onToggleClosed(); },
      exportJson:   function()   { this._onExportJson(); },
      postToChat:   function()   { this._onPostToChat(); },
    },
  };

  static PARTS = {
    body: { template: `modules/${MODULE_ID}/templates/store-sheet.hbs` },
  };

  /**
   * @param {JournalEntry} document
   * @param {object}       [options]
   */
  constructor(document, options = {}) {
    super(options);
    this.document = document;
  }

  get title() { return `${this.document?.name || 'Store'} — Store Sheet`; }
  get id()    { return `store-sheet-${this.document?.id || 'unknown'}`; }

  async _prepareContext() {
    const store = sanitizeStore(getStore(this.document) || {});

    const inventory = store.inventory.map(i => ({
      ...i,
      displayPrice: tieredPrice(i.price, store.priceTier),
      lineValue:    tieredPrice(i.price, store.priceTier) * (i.stock || 0),
    }));

    return {
      isGM:        !!game.user?.isGM,
      docName:     this.document?.name || store.name,
      store,
      inventory,
      typeLabel:   storeTypeLabel(store.type),
      storeTypes:  STORE_TYPES.map(t => ({ value: t, label: storeTypeLabel(t), selected: t === store.type })),
      priceTiers:  PRICE_TIERS.map(t => ({
        value: t,
        label: t.charAt(0).toUpperCase() + t.slice(1),
        selected: t === store.priceTier,
      })),
      stockCount:  store.inventory.reduce((n, i) => n + (i.stock || 0), 0),
      totalValue:  inventoryValue(store),
      hasServices: store.services.length > 0,
      hasRumors:   store.rumors.length > 0,
      hasStaff:    store.staff.length > 0,
    };
  }

  /* ── Mutation helpers ───────────────────────────────────── */

  /**
   * Applies `mutate` to a deep clone of the stored payload and writes it back.
   * @param {(store: object) => void} mutate
   */
  async _update(mutate) {
    const store = sanitizeStore(getStore(this.document) || {});
    mutate(store);
    await setStore(this.document, store);
    this.render();
  }

  /* ── Actions ────────────────────────────────────────────── */

  /**
   * Writes one edited field back onto the store. The element carries
   * `data-path` (dot-path) and, for inventory/staff rows, `data-id`.
   */
  async _onSaveField(event) {
    const el = event?.currentTarget;
    if (!el?.dataset?.path || !game.user?.isGM) return;

    const path  = el.dataset.path;
    const rowId = el.dataset.id || null;
    const value = el.type === 'checkbox' ? el.checked
      : (el.type === 'number' ? Number(el.value) : el.value);

    await this._update(store => {
      if (rowId) {
        const list = path.startsWith('inventory.') ? store.inventory
          : path.startsWith('staff.') ? store.staff
          : null;
        if (!list) return;
        const row = list.find(r => r.id === rowId);
        if (row) row[path.split('.').pop()] = value;
        return;
      }
      foundry.utils.setProperty(store, path, value);
    });
  }

  async _onAddItem() {
    if (!game.user?.isGM) return;
    await this._update(store => {
      store.inventory.push({
        id: `inv-${foundry.utils.randomID(8)}`,
        name: 'New Good', description: '', price: 1, stock: 1,
        level: 0, rarity: 'common', itemId: null,
      });
    });
  }

  async _onRemoveItem(event) {
    const id = event?.currentTarget?.dataset?.id;
    if (!id || !game.user?.isGM) return;
    await this._update(store => {
      store.inventory = store.inventory.filter(i => i.id !== id);
    });
  }

  async _onAddStaff() {
    if (!game.user?.isGM) return;
    await this._update(store => {
      store.staff.push({
        id: `staff-${foundry.utils.randomID(8)}`,
        name: 'New Hire', role: 'Assistant', shift: 'day', actorId: null,
      });
    });
  }

  async _onRemoveStaff(event) {
    const id = event?.currentTarget?.dataset?.id;
    if (!id || !game.user?.isGM) return;
    await this._update(store => {
      store.staff = store.staff.filter(m => m.id !== id);
    });
  }

  async _onToggleClosed() {
    if (!game.user?.isGM) return;
    await this._update(store => { store.closed = !store.closed; });
  }

  _onExportJson() {
    const store = sanitizeStore(getStore(this.document) || {});
    const blob  = new Blob([JSON.stringify(store, null, 2)], { type: 'application/json' });
    const url   = URL.createObjectURL(blob);
    const a     = document.createElement('a');
    a.href      = url;
    a.download  = `${String(store.name).replace(/[^a-z0-9_-]/gi, '_')}.json`;
    a.click();
    URL.revokeObjectURL(url);
  }

  /** Whispers a short shop summary (and its price list) to the GMs. */
  _onPostToChat() {
    const store = sanitizeStore(getStore(this.document) || {});
    const rows  = store.inventory.map(i =>
      `<tr><td>${escapeHtml(i.name)}</td><td style="text-align:right">${i.stock}</td>` +
      `<td style="text-align:right">${tieredPrice(i.price, store.priceTier)} gp</td></tr>`).join('');

    const content = `
<h3>${escapeHtml(store.name)}</h3>
<p><em>${escapeHtml(storeTypeLabel(store.type))}${store.settlement ? ` · ${escapeHtml(store.settlement)}` : ''}</em></p>
<p>${escapeHtml(store.description)}</p>
<p><strong>Owner:</strong> ${escapeHtml(store.owner.name)} · <strong>Hours:</strong> ${escapeHtml(store.hours.open)}–${escapeHtml(store.hours.close)}</p>
${rows ? `<table><thead><tr><th>Item</th><th>Stock</th><th>Price</th></tr></thead><tbody>${rows}</tbody></table>` : ''}`.trim();

    ChatMessage.create({
      content,
      whisper: game.users.filter(u => u.isGM).map(u => u.id),
    });
  }
}

/**
 * Swaps in StoreSheet whenever a JournalEntry carrying a store payload is
 * opened, so the generated shop never shows Foundry's default page list.
 */
export function registerStoreSheetHooks() {
  Hooks.on('renderJournalSheet', (app) => {
    try {
      const doc = app?.document;
      if (!doc?.getFlag?.(FLAG_SCOPE, FLAG_KEY)) return;
      // Defer the swap so we don't recurse during render.
      app.close({ submit: false });
      queueMicrotask(() => new StoreSheet(doc).render(true));
    } catch (err) {
      console.error(`[${MODULE_ID}] store sheet swap failed:`, err);
    }
  });
}
