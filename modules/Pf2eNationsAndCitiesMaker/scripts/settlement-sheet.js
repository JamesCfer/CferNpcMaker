/**
 * SettlementSheet — PF2e-themed custom sheet for City/Town/Village JournalEntries.
 * ApplicationV2 + HandlebarsApplicationMixin.
 *
 * Data lives in flags.Pf2eNationsAndCitiesMaker.settlement; every mutation
 * goes through doc.setFlag and the sheet re-renders.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY, getSettlement, STORE_TYPES, storeTypeLabel } from './constants.js';
import { applyDailyTick, applyTax }                                                     from './economy.js';
import { generateStaffNpc, generateStoreItem, canGenerateNpc, canGenerateItem }         from './integrations.js';
import { sanitizeSettlement }                                                           from './sanitizer.js';

const { HandlebarsApplicationMixin, ApplicationV2 } = foundry.applications.api;

export class SettlementSheet extends HandlebarsApplicationMixin(ApplicationV2) {
  static DEFAULT_OPTIONS = {
    id: 'settlement-sheet-{id}',
    classes: ['pf2e-settlement-sheet'],
    tag: 'form',
    window: { resizable: true, contentClasses: ['pf2e-settlement-sheet-window'] },
    position: { width: 920, height: 720 },
    actions: {
      switchTab:           function(ev) { this._onSwitchTab(ev); },
      switchStoreTab:      function(ev) { this._onSwitchStoreTab(ev); },
      editField:           function(ev) { this._onEditField(ev); },
      tickDay:             function()   { this._onTickDay(1); },
      tickWeek:            function()   { this._onTickDay(7); },
      manualTax:           function()   { this._onManualTax(); },
      addStore:            function()   { this._onAddStore(); },
      removeStore:         function(ev) { this._onRemoveStore(ev); },
      addStaff:            function(ev) { this._onAddStaff(ev); },
      removeStaff:         function(ev) { this._onRemoveStaff(ev); },
      generateShopkeeper:  function(ev) { this._onGenerateShopkeeper(ev); },
      generateStoreItem:   function(ev) { this._onGenerateStoreItem(ev); },
      openLinkedActor:     function(ev) { this._onOpenLinkedActor(ev); },
      addLeader:           function()   { this._onAddLeader(); },
      removeLeader:        function(ev) { this._onRemoveLeader(ev); },
      generateLeader:      function(ev) { this._onGenerateLeader(ev); },
      addRank:             function()   { this._onAddRank(); },
      removeRank:          function(ev) { this._onRemoveRank(ev); },
      generateCommander:   function()   { this._onGenerateCommander(); },
      openBuilder:         function()   { this._onOpenBuilder(); },
      saveNotes:           function()   { this._onSaveNotes(); },
    },
  };

  static PARTS = {
    body: { template: `modules/${MODULE_ID}/templates/city-sheet.hbs` },
  };

  constructor(document, options = {}) {
    super(options);
    this.document = document;
    this.activeTab = 'overview';
    this.activeStoreTab = null;
  }

  get title() { return `${this.document?.name || 'Settlement'} — Settlement Sheet`; }
  get id()    { return `settlement-sheet-${this.document?.id || 'unknown'}`; }

  async _prepareContext() {
    const raw = getSettlement(this.document) || {};
    const settlement = sanitizeSettlement(raw); // fills defaults defensively

    // Group stores by type for the inner tabs.
    const storesByType = {};
    for (const store of settlement.stores) {
      const key = store.type || 'other';
      (storesByType[key] = storesByType[key] || []).push(store);
    }
    const storeTabs = Object.keys(storesByType)
      .sort((a, b) => a.localeCompare(b))
      .map(type => ({
        type, label: storeTypeLabel(type),
        count: storesByType[type].length,
        stores: storesByType[type],
        isActive: this.activeStoreTab ? type === this.activeStoreTab : false,
      }));
    if (storeTabs.length && !storeTabs.some(t => t.isActive)) {
      storeTabs[0].isActive = true;
      this.activeStoreTab = storeTabs[0].type;
    }

    const calendarActive = !!game.modules?.get('Pf2eCalendarTimeline')?.active;

    const tabs = [
      { key: 'overview',   label: 'Overview',     icon: 'fa-solid fa-flag' },
      { key: 'stores',     label: 'Stores',       icon: 'fa-solid fa-store' },
      { key: 'guards',     label: 'Guards & Military', icon: 'fa-solid fa-shield-halved' },
      { key: 'leadership', label: 'Leadership',   icon: 'fa-solid fa-crown' },
      { key: 'production', label: 'Production',   icon: 'fa-solid fa-wheat-awn' },
      { key: 'notes',      label: 'Notes',        icon: 'fa-solid fa-scroll' },
    ].map(t => ({ ...t, isActive: t.key === this.activeTab }));

    const activeKey = this.activeTab;

    return {
      doc:        this.document,
      docId:      this.document.id,
      settlement,
      tabs,
      isOverviewActive:   activeKey === 'overview',
      isStoresActive:     activeKey === 'stores',
      isGuardsActive:     activeKey === 'guards',
      isLeadershipActive: activeKey === 'leadership',
      isProductionActive: activeKey === 'production',
      isNotesActive:      activeKey === 'notes',
      storeTabs,
      storeTypeOptions: STORE_TYPES.map(t => ({ value: t, label: storeTypeLabel(t) })),
      hpPct: Math.round((settlement.stats.hp / Math.max(1, settlement.stats.maxHp)) * 100),
      moralePct: settlement.stats.morale,
      unrestPct: settlement.stats.unrest,
      calendarActive,
      canGenerateNpc:  canGenerateNpc(),
      canGenerateItem: canGenerateItem(),
    };
  }

  _onRender() {
    // Bind text/number inputs (delegated) — write to flags on blur/change.
    this.element.querySelectorAll('[data-settlement-path]').forEach(input => {
      input.addEventListener('change', (ev) => this._writePath(ev.currentTarget));
    });
    // Document.execCommand-free notes editor: just a textarea.
  }

  /* ── helpers ───────────────────────────────────────────── */

  async _patch(mutator) {
    const cur  = foundry.utils.deepClone(getSettlement(this.document) || {});
    mutator(cur);
    await this.document.setFlag(FLAG_SCOPE, FLAG_KEY, cur);
    this.render(false);
  }

  _writePath(input) {
    const path = input.dataset.settlementPath;
    if (!path) return;
    let value = input.value;
    if (input.type === 'number') value = Number(value);
    if (input.type === 'checkbox') value = input.checked;
    this._patch(s => foundry.utils.setProperty(s, path, value));
  }

  /* ── tab actions ───────────────────────────────────────── */

  _onSwitchTab(ev) {
    const key = ev.currentTarget?.dataset?.tab;
    if (!key) return;
    this.activeTab = key;
    this.render(false);
  }

  _onSwitchStoreTab(ev) {
    const key = ev.currentTarget?.dataset?.storeTab;
    if (!key) return;
    this.activeStoreTab = key;
    this.render(false);
  }

  _onEditField(ev) {
    const input = ev.currentTarget?.closest('label, div')?.querySelector('input, textarea, select');
    input?.focus?.();
  }

  /* ── economy ───────────────────────────────────────────── */

  async _onTickDay(days = 1) {
    await applyDailyTick(this.document, days);
    ui.notifications?.info?.(`Ticked ${days} day${days === 1 ? '' : 's'} for ${this.document.name}.`);
    this.render(false);
  }

  async _onManualTax() {
    const html = `
      <div style="display:flex;flex-direction:column;gap:0.5em;">
        <label>Tax type
          <select name="taxType" style="width:100%;margin-top:0.25em;">
            <option value="income">Income</option>
            <option value="poll">Poll</option>
            <option value="trade">Trade</option>
            <option value="property">Property</option>
          </select>
        </label>
        <label>Rate (%)
          <input type="number" name="ratePct" value="5" min="0" max="100" step="1" style="width:100%;margin-top:0.25em;" />
        </label>
      </div>`;
    const result = await foundry.applications.api.DialogV2.prompt({
      window: { title: 'Apply Tax' },
      content: html,
      ok: {
        label: 'Apply',
        callback: (_e, _b, dlg) => {
          const root = dlg.element;
          return {
            taxType: root.querySelector('[name="taxType"]')?.value || 'income',
            ratePct: Number(root.querySelector('[name="ratePct"]')?.value) || 0,
          };
        },
      },
      rejectClose: false,
    }).catch(() => null);
    if (!result) return;
    const { collected } = await applyTax(this.document, result);
    ui.notifications?.info?.(`Tax collected: ${collected} gp.`);
    this.render(false);
  }

  /* ── stores ────────────────────────────────────────────── */

  _findStore(s, storeId) { return s.stores?.find(x => x.id === storeId); }

  _onAddStore() {
    this._patch(s => {
      s.stores = s.stores || [];
      s.stores.push({
        id: `shop-${Math.random().toString(36).slice(2, 10)}`,
        name: 'New Shop',
        type: this.activeStoreTab || 'general',
        owner: { name: 'Owner', actorId: null },
        staff: [],
        hours: { open: '08:00', close: '20:00', daysClosed: [] },
        inventory: [],
        income: { balance: 0, dailyAvg: 5, lastTick: 0 },
      });
    });
  }

  _onRemoveStore(ev) {
    const storeId = ev.currentTarget?.dataset?.storeId;
    if (!storeId) return;
    this._patch(s => { s.stores = (s.stores || []).filter(x => x.id !== storeId); });
  }

  _onAddStaff(ev) {
    const storeId = ev.currentTarget?.dataset?.storeId;
    if (!storeId) return;
    this._patch(s => {
      const store = this._findStore(s, storeId);
      if (!store) return;
      store.staff = store.staff || [];
      store.staff.push({
        id: `staff-${Math.random().toString(36).slice(2, 10)}`,
        name: 'New Worker', role: 'clerk', shift: 'day', actorId: null,
      });
    });
  }

  _onRemoveStaff(ev) {
    const { storeId, staffId } = ev.currentTarget?.dataset || {};
    if (!storeId || !staffId) return;
    this._patch(s => {
      const store = this._findStore(s, storeId);
      if (!store) return;
      store.staff = (store.staff || []).filter(p => p.id !== staffId);
    });
  }

  _onGenerateShopkeeper(ev) {
    const { storeId, staffId } = ev.currentTarget?.dataset || {};
    const settlement = getSettlement(this.document) || {};
    const store = settlement.stores?.find(s => s.id === storeId);
    if (!store) return;
    const staff = staffId === 'owner' ? { id: 'owner', name: store.owner?.name || 'Owner', role: 'shopkeeper' }
                                       : store.staff?.find(p => p.id === staffId);
    if (!staff) return;
    generateStaffNpc({
      settlementDoc: this.document,
      storeId,
      staff,
      level: Math.max(1, Math.min(10, Math.ceil((settlement.population || 500) / 1000))),
      onCreate: () => this.render(false),
    });
  }

  _onGenerateStoreItem(ev) {
    const { storeId, itemType } = ev.currentTarget?.dataset || {};
    if (!storeId) return;
    generateStoreItem({
      settlementDoc: this.document,
      storeId,
      itemType: itemType || 'equipment',
      onCreate: () => this.render(false),
    });
  }

  _onOpenLinkedActor(ev) {
    const actorId = ev.currentTarget?.dataset?.actorId;
    if (!actorId) return;
    const actor = game.actors?.get(actorId);
    actor?.sheet?.render(true);
  }

  /* ── leadership ────────────────────────────────────────── */

  _onAddLeader() {
    this._patch(s => {
      s.leadership = s.leadership || [];
      s.leadership.push({ title: 'Council member', name: 'Unnamed', role: '', actorId: null });
    });
  }
  _onRemoveLeader(ev) {
    const idx = Number(ev.currentTarget?.dataset?.index);
    if (!Number.isFinite(idx)) return;
    this._patch(s => { s.leadership = (s.leadership || []).filter((_, i) => i !== idx); });
  }
  _onGenerateLeader(ev) {
    const idx = Number(ev.currentTarget?.dataset?.index);
    const settlement = getSettlement(this.document) || {};
    const leader = settlement.leadership?.[idx];
    if (!leader) return;
    Hooks.callAll('Pf2eNpcMaker.openWithPrefill', {
      name: leader.name,
      level: Math.max(2, Math.min(15, Math.ceil((settlement.population || 500) / 800))),
      description: `${leader.title || 'Leader'} of ${this.document.name}. ${leader.role || ''}`.trim(),
      onCreate: (actor) => {
        this._patch(s => { if (s.leadership?.[idx]) s.leadership[idx].actorId = actor.id; });
      },
    });
  }

  /* ── military ──────────────────────────────────────────── */

  _onAddRank() {
    this._patch(s => {
      s.military = s.military || { ranks: [], totalGuards: 0 };
      s.military.ranks = s.military.ranks || [];
      s.military.ranks.push({ rank: 'Guard', count: 1, leaderName: '' });
      s.military.totalGuards = (s.military.ranks).reduce((acc, r) => acc + (Number(r.count) || 0), 0);
    });
  }
  _onRemoveRank(ev) {
    const idx = Number(ev.currentTarget?.dataset?.index);
    if (!Number.isFinite(idx)) return;
    this._patch(s => {
      s.military.ranks = (s.military.ranks || []).filter((_, i) => i !== idx);
      s.military.totalGuards = s.military.ranks.reduce((acc, r) => acc + (Number(r.count) || 0), 0);
    });
  }
  _onGenerateCommander() {
    const settlement = getSettlement(this.document) || {};
    Hooks.callAll('Pf2eNpcMaker.openWithPrefill', {
      name: settlement.military?.commanderName || 'Captain of the Guard',
      level: Math.max(3, Math.min(15, Math.ceil((settlement.population || 500) / 700))),
      description: `Commander of the guard in ${this.document.name}. Leads ${settlement.military?.totalGuards || 'the'} guards.`,
      onCreate: (actor) => {
        this._patch(s => {
          s.military = s.military || {};
          s.military.commanderActorId = actor.id;
          s.military.commanderName    = actor.name;
        });
      },
    });
  }

  /* ── misc ──────────────────────────────────────────────── */

  _onOpenBuilder() {
    // Lazy import the adapter so we don't pull builder code on every sheet render.
    import('./adapter.js').then(({ SettlementAdapter }) => {
      const adapter = new SettlementAdapter();
      import('./core/app.js').then(({ openBuilder }) => openBuilder(adapter));
    });
  }

  _onSaveNotes() {
    const ta = this.element.querySelector('[data-settlement-path="notes"]');
    if (!ta) return;
    this._patch(s => { s.notes = ta.value; });
  }
}

/**
 * Replace the default journal sheet with our SettlementSheet when the journal
 * has a settlement flag set. Returns true if the swap happened.
 */
export function maybeOpenSettlementSheet(journal) {
  const s = getSettlement(journal);
  if (!s) return false;
  if (s.kind === 'nation') return false; // handled by NationSheet
  const sheet = new SettlementSheet(journal);
  sheet.render(true);
  return true;
}
