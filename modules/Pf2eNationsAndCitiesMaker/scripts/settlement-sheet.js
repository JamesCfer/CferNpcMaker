/**
 * SettlementSheet — PF2e-themed custom sheet for City/Town/Village JournalEntries.
 * ApplicationV2 + HandlebarsApplicationMixin.
 *
 * Data lives in flags.Pf2eNationsAndCitiesMaker.settlement; every mutation
 * goes through doc.setFlag and the sheet re-renders.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY, getSettlement, STORE_TYPES, storeTypeLabel } from './constants.js';
import { applyDailyTick, applyTax }                                                     from './economy.js';
import { generateStaffNpc, generateStoreItem, canGenerateNpc, canGenerateItem,
         rerollStores, rerollSingleStore }                                               from './integrations.js';
import { sanitizeSettlement }                                                           from './sanitizer.js';
import { goodsForProduction }                                                           from './trade-goods.js';

const { HandlebarsApplicationMixin, ApplicationV2 } = foundry.applications.api;

const PRICE_TIER_MULS = { low: 0.75, standard: 1.0, high: 1.5, luxury: 2.0 };

function resolveActor(actorId) {
  if (!actorId) return null;
  const actor = game.actors?.get(actorId);
  if (!actor) return null;
  return {
    portrait: actor.img || '',
    level:    actor.system?.details?.level?.value ?? actor.system?.details?.cr ?? null,
    name:     actor.name,
  };
}

function buildSparklineSvg(history) {
  if (!history || history.length < 2) return '';
  const W = 240, H = 36, PAD = 2;
  const vals = history.map(h => Number(h.gp) || 0);
  const min = Math.min(...vals);
  const max = Math.max(...vals);
  const range = max - min || 1;
  const n = vals.length;
  const xs = vals.map((_, i) => PAD + (i / (n - 1)) * (W - PAD * 2));
  const ys = vals.map(v => (H - PAD) - ((v - min) / range) * (H - PAD * 2));
  const pts = xs.map((x, i) => `${x.toFixed(1)},${ys[i].toFixed(1)}`).join(' ');
  const area = `${xs[0].toFixed(1)},${H - PAD} ${pts} ${xs[n - 1].toFixed(1)},${H - PAD}`;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${W} ${H}" width="${W}" height="${H}" class="pf2e-sparkline">` +
    `<polygon points="${area}" fill="rgba(155,105,35,0.15)" />` +
    `<polyline points="${pts}" fill="none" stroke="#9b6923" stroke-width="1.5" stroke-linejoin="round" />` +
    `</svg>`;
}

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
      toggleShowClosed:    function()   { this._onToggleShowClosed(); },
      toggleCompactStores: function()   { this._onToggleCompactStores(); },
      rerollAllStores:     function()   { this._onRerollAllStores(); },
      rerollStore:         function(ev) { this._onRerollStore(ev); },
      reopenStore:         function(ev) { this._onReopenStore(ev); },
      addTradeRoute:       function()   { this._onAddTradeRoute(); },
      removeTradeRoute:    function(ev) { this._onRemoveTradeRoute(ev); },
      addReligion:         function()   { this._onAddReligion(); },
      removeReligion:      function(ev) { this._onRemoveReligion(ev); },
      addDemographic:      function()   { this._onAddDemographic(); },
      removeDemographic:   function(ev) { this._onRemoveDemographic(ev); },
      levelUpActor:        function(ev) { this._onLevelUpActor(ev); },
      addDistrict:         function()   { this._onAddDistrict(); },
      removeDistrict:      function(ev) { this._onRemoveDistrict(ev); },
      exportHandout:       function()   { this._onExportHandout(); },
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
    this.showClosed = false;
    this.compactStores = false;
    this.districtFilter = null;
  }

  get title() { return `${this.document?.name || 'Settlement'} — Settlement Sheet`; }
  get id()    { return `settlement-sheet-${this.document?.id || 'unknown'}`; }

  async _prepareContext() {
    const raw = getSettlement(this.document) || {};
    const settlement = sanitizeSettlement(raw); // fills defaults defensively

    const isGM = !!game.user?.isGM;
    const showTreasury = !settlement.gmOnly?.treasury || isGM;
    const showMilitary = !settlement.gmOnly?.military || isGM;
    const showIncome   = !settlement.gmOnly?.income   || isGM;

    // Hide black market stores from non-GM players (#60)
    if (!isGM) {
      settlement.stores = settlement.stores.filter(st => !st.isBlackMarket);
    }

    const showClosed = this.showClosed;
    const closedStoreCount = settlement.stores.filter(s => s.closed).length;
    const districtFilter = this.districtFilter;
    const totalDailyWages = (settlement.military?.ranks || []).reduce(
      (sum, r) => sum + Number(r.dailyWage || 0) * Number(r.count || 0), 0
    );

    // Weekday names — resolved early so the store loop can label schedule entries (#43).
    let weekdayNames;
    try {
      const calState = game.settings.get('Pf2eCalendarTimeline', 'state');
      weekdayNames = calState?.calendarDef?.weekdays;
    } catch (_) {}
    if (!weekdayNames?.length) {
      weekdayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    }
    const weekdayOptions = weekdayNames.map((n, i) => ({ value: i, label: n }));

    // Group stores by type for the inner tabs; hide closed/filtered stores as appropriate.
    // Decorate each store with its combined effective price multiplier and global index.
    const storesByType = {};
    for (let gi = 0; gi < settlement.stores.length; gi++) {
      const store = settlement.stores[gi];
      if (store.closed && !showClosed) continue;
      if (districtFilter && store.districtId !== districtFilter) continue;
      const key = store.type || 'other';
      const effectiveMul = Math.round((PRICE_TIER_MULS[store.priceTier] ?? 1.0) * (settlement.priceMultiplier || 1) * 100) / 100;
      (storesByType[key] = storesByType[key] || []).push({
        ...store,
        _globalIndex: gi,
        effectiveMul,
        ownerActor: resolveActor(store.owner?.actorId),
        staff: (store.staff || []).map(p => ({ ...p, resolvedActor: resolveActor(p.actorId) })),
        hours: {
          ...store.hours,
          schedule: (store.hours?.schedule || []).map(entry => ({
            ...entry,
            dayLabel: weekdayNames[entry.day] ?? `Day ${entry.day}`,
          })),
        },
      });
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

    // Trade goods matching current production tags (#57)
    const tradeGoods = goodsForProduction(settlement.production).map(g => ({
      name: g.name,
      priceGp: g.priceGp,
      effectivePrice: Math.round(g.priceGp * (settlement.priceMultiplier || 1) * 100) / 100,
    }));

    const calendarActive = !!game.modules?.get('Pf2eCalendarTimeline')?.active;

    const sparklineSvg = buildSparklineSvg(raw.treasuryHistory);
    const settlementJournals = (game.journal?.contents || [])
      .filter(j => j.id !== this.document.id && getSettlement(j))
      .map(j => ({ id: j.id, name: j.name }))
      .sort((a, b) => a.name.localeCompare(b.name));

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
      canGenerateNpc:   canGenerateNpc(),
      canGenerateItem:  canGenerateItem(),
      showClosed,
      closedStoreCount,
      compactStores: this.compactStores,
      totalDailyWages,
      weekdayOptions,
      priceTierOptions: [
        { value: 'low',      label: 'Low (×0.75)' },
        { value: 'standard', label: 'Standard' },
        { value: 'high',     label: 'High (×1.5)' },
        { value: 'luxury',   label: 'Luxury (×2.0)' },
      ],
      shiftOptions: [
        { value: 'morning',   label: 'Morning' },
        { value: 'day',       label: 'Day' },
        { value: 'evening',   label: 'Evening' },
        { value: 'night',     label: 'Night' },
        { value: 'graveyard', label: 'Graveyard' },
      ],
      tradeGoods,
      priceMultiplier: settlement.priceMultiplier,
      sparklineSvg,
      settlementJournals,
      isGM,
      showTreasury,
      showMilitary,
      showIncome,
      districts: settlement.districts || [],
      districtFilter,
    };
  }

  _onRender() {
    this.element.querySelectorAll('[data-settlement-path]').forEach(input => {
      const tag  = input.tagName.toLowerCase();
      const type = input.getAttribute('type')?.toLowerCase();
      if (tag === 'select' || type === 'checkbox') {
        input.addEventListener('change', (ev) => this._writePath(ev.currentTarget));
      } else {
        input.addEventListener('input', (ev) => this._scheduleAutoSave(ev.currentTarget));
      }
    });

    // District filter select (in-memory, not persisted)
    const distFilterEl = this.element.querySelector('[data-district-filter]');
    if (distFilterEl) {
      distFilterEl.addEventListener('change', (ev) => {
        this.districtFilter = ev.currentTarget.value || null;
        this.render(false);
      });
    }

    // Drag-drop actor → staff row or owner field (#40)
    this.element.querySelectorAll('.pf2e-staff-name, [data-is-owner]').forEach(input => {
      input.addEventListener('dragover', (ev) => { ev.preventDefault(); ev.dataTransfer.dropEffect = 'link'; });
      input.addEventListener('drop', (ev) => this._onDropActorOnStaff(ev));
    });

    // Drag-drop PF2e item → store inventory (#41)
    this.element.querySelectorAll('.pf2e-store-inventory').forEach(section => {
      section.addEventListener('dragover',  (ev) => { ev.preventDefault(); ev.dataTransfer.dropEffect = 'copy'; section.classList.add('drag-over'); });
      section.addEventListener('dragleave', ()   => section.classList.remove('drag-over'));
      section.addEventListener('drop',      (ev) => { section.classList.remove('drag-over'); this._onDropItemOnInventory(ev); });
    });
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
    else if (input.type === 'checkbox') value = input.checked;
    else if ('nullableInt' in input.dataset) value = value === '' ? null : Number(value);
    else if ('nullableStr' in input.dataset) value = value === '' ? null : value;
    this._patch(s => foundry.utils.setProperty(s, path, value));
  }

  _scheduleAutoSave(input) {
    clearTimeout(this._autosaveTimer);
    this._autosaveTimer = setTimeout(() => this._writePathSilent(input), 400);
  }

  async _writePathSilent(input) {
    const path = input.dataset.settlementPath;
    if (!path) return;
    let value = input.value;
    if (input.type === 'number') value = Number(value);
    else if ('nullableInt' in input.dataset) value = value === '' ? null : Number(value);
    else if ('nullableStr' in input.dataset) value = value === '' ? null : value;
    const cur = foundry.utils.deepClone(getSettlement(this.document) || {});
    foundry.utils.setProperty(cur, path, value);
    await this.document.setFlag(FLAG_SCOPE, FLAG_KEY, cur);
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

  /* ── closed store actions ──────────────────────────────── */

  _onToggleShowClosed() {
    this.showClosed = !this.showClosed;
    this.render(false);
  }

  _onToggleCompactStores() {
    this.compactStores = !this.compactStores;
    this.render(false);
  }

  async _onRerollAllStores() {
    const confirmed = await foundry.applications.api.DialogV2.confirm({
      window:      { title: 'Re-roll All Stores' },
      content:     '<p>Replace all stores with a newly generated set? Existing store data will be lost.</p>',
      yes:         { label: 'Re-roll', icon: 'fa-solid fa-dice' },
      no:          { label: 'Cancel' },
      rejectClose: false,
    }).catch(() => false);
    if (!confirmed) return;
    ui.notifications?.info?.('Re-rolling stores…');
    try {
      await rerollStores(this.document);
      ui.notifications?.info?.('Stores re-rolled.');
    } catch (err) {
      ui.notifications?.error?.(`Store re-roll failed: ${err.message}`);
    }
    this.render(false);
  }

  async _onRerollStore(ev) {
    const storeId = ev.currentTarget?.dataset?.storeId;
    if (!storeId) return;
    ui.notifications?.info?.('Re-rolling store…');
    try {
      await rerollSingleStore(this.document, storeId);
      ui.notifications?.info?.('Store re-rolled.');
    } catch (err) {
      ui.notifications?.error?.(`Store re-roll failed: ${err.message}`);
    }
    this.render(false);
  }

  _onReopenStore(ev) {
    const storeId = ev.currentTarget?.dataset?.storeId;
    if (!storeId) return;
    this._patch(s => {
      const store = this._findStore(s, storeId);
      if (!store) return;
      store.closed = false;
      store.income.daysInDebt = 0;
    });
  }

  /* ── trade routes ──────────────────────────────────────── */

  _onAddTradeRoute() {
    this._patch(s => {
      s.tradeRoutes = s.tradeRoutes || [];
      s.tradeRoutes.push({ partnerId: '', goods: '', gpPerWeek: 0 });
    });
  }

  _onRemoveTradeRoute(ev) {
    const idx = Number(ev.currentTarget?.dataset?.index);
    if (!Number.isFinite(idx)) return;
    this._patch(s => { s.tradeRoutes = (s.tradeRoutes || []).filter((_, i) => i !== idx); });
  }

  /* ── religions ─────────────────────────────────────────────── */

  _onAddReligion() {
    this._patch(s => {
      s.religions = s.religions || [];
      s.religions.push({
        id: `rel-${Math.random().toString(36).slice(2, 10)}`,
        name: 'New Religion', followers: 0, templeStoreId: null, influence: 0,
      });
    });
  }

  _onRemoveReligion(ev) {
    const id = ev.currentTarget?.dataset?.religionId;
    if (!id) return;
    this._patch(s => { s.religions = (s.religions || []).filter(r => r.id !== id); });
  }

  /* ── districts (#45) ──────────────────────────────────── */

  _onAddDistrict() {
    this._patch(s => {
      s.districts = s.districts || [];
      s.districts.push({
        id:            `dist-${Math.random().toString(36).slice(2, 10)}`,
        name:          'New District',
        descriptor:    '',
        leaderActorId: null,
      });
    });
  }

  _onRemoveDistrict(ev) {
    const id = ev.currentTarget?.dataset?.districtId;
    if (!id) return;
    this._patch(s => {
      s.districts = (s.districts || []).filter(d => d.id !== id);
      (s.stores || []).forEach(st => { if (st.districtId === id) st.districtId = null; });
    });
    if (this.districtFilter === id) {
      this.districtFilter = null;
    }
  }

  /* ── demographics (#69) ────────────────────────────────── */

  _onAddDemographic() {
    this._patch(s => {
      s.demographics = s.demographics || [];
      s.demographics.push({ ancestry: 'Human', pct: 0 });
    });
  }

  _onRemoveDemographic(ev) {
    const idx = Number(ev.currentTarget?.dataset?.index);
    if (!Number.isFinite(idx)) return;
    this._patch(s => { s.demographics = (s.demographics || []).filter((_, i) => i !== idx); });
  }

  /* ── inline NPC chip (#39) ─────────────────────────────── */

  _onLevelUpActor(ev) {
    const actorId = ev.currentTarget?.dataset?.actorId;
    if (!actorId) return;
    game.actors?.get(actorId)?.sheet?.render(true);
  }

  /* ── drag-drop actor → staff (#40) ────────────────────── */

  async _onDropActorOnStaff(ev) {
    ev.preventDefault();
    let data;
    try { data = JSON.parse(ev.dataTransfer.getData('text/plain')); } catch { return; }
    if (data.type !== 'Actor') return;
    const actor = await fromUuid(data.uuid).catch(() => null);
    if (!actor) return;

    const card  = ev.currentTarget.closest('[data-store-id]');
    const storeId = card?.dataset?.storeId;
    if (!storeId) return;

    const staffId = ev.currentTarget.dataset?.staffId;
    const isOwner = !!ev.currentTarget.dataset?.isOwner;

    this._patch(s => {
      const store = this._findStore(s, storeId);
      if (!store) return;
      if (isOwner) {
        store.owner = store.owner || {};
        store.owner.name    = actor.name;
        store.owner.actorId = actor.id;
      } else if (staffId) {
        const person = store.staff?.find(p => p.id === staffId);
        if (person) { person.name = actor.name; person.actorId = actor.id; }
      }
    });
  }

  /* ── drag-drop PF2e item → inventory (#41) ────────────── */

  async _onDropItemOnInventory(ev) {
    ev.preventDefault();
    let data;
    try { data = JSON.parse(ev.dataTransfer.getData('text/plain')); } catch { return; }
    if (data.type !== 'Item') return;
    const item = await fromUuid(data.uuid).catch(() => null);
    if (!item) return;
    const card    = ev.currentTarget.closest('[data-store-id]');
    const storeId = card?.dataset?.storeId;
    if (!storeId) return;
    const pv      = item.system?.price?.value ?? {};
    const priceGp = Math.round(
      ((Number(pv.pp) || 0) * 10 + (Number(pv.gp) || 0) + (Number(pv.sp) || 0) / 10 + (Number(pv.cp) || 0) / 100) * 100
    ) / 100;
    this._patch(s => {
      const store = this._findStore(s, storeId);
      if (!store) return;
      store.inventory = store.inventory || [];
      store.inventory.push({
        id:     `inv-${Math.random().toString(36).slice(2, 10)}`,
        name:   item.name,
        itemId: item.id,
        price:  priceGp,
        stock:  1,
      });
    });
  }

  /* ── handout export (#49) ─────────────────────────────── */

  _onExportHandout() {
    const raw  = getSettlement(this.document) || {};
    const s    = sanitizeSettlement(raw);
    const name = this.document.name || 'Settlement';

    const lines = [];
    lines.push(`# ${name}`);
    lines.push('');
    lines.push(`**Kind:** ${s.kind}  ·  **Size:** ${s.size}  ·  **Population:** ${s.population.toLocaleString()}`);
    if (s.biome) lines.push(`**Biome:** ${s.biome}`);
    if (s.government?.type)       lines.push(`**Government:** ${s.government.type}`);
    if (s.government?.leaderName) lines.push(`**Leader:** ${s.government.leaderName}`);
    lines.push('');

    lines.push('## Stats');
    lines.push(`- HP: ${s.stats.hp} / ${s.stats.maxHp}  (DT ${s.stats.damageThreshold}, Hardness ${s.stats.hardness})`);
    lines.push(`- Morale: ${s.stats.morale}%  ·  Unrest: ${s.stats.unrest}%`);
    lines.push(`- Fort: +${s.stats.fortitude}  Ref: +${s.stats.reflex}  Will: +${s.stats.will}`);
    lines.push('');

    const { gp, sp, cp, pp } = s.treasury;
    if (gp || sp || cp || pp) {
      lines.push('## Treasury');
      const parts = [];
      if (pp) parts.push(`${pp} pp`);
      if (gp) parts.push(`${gp} gp`);
      if (sp) parts.push(`${sp} sp`);
      if (cp) parts.push(`${cp} cp`);
      lines.push(parts.join('  ·  '));
      lines.push('');
    }

    const activeStores = (s.stores || []).filter(st => !st.closed);
    if (activeStores.length) {
      lines.push('## Stores');
      for (const store of activeStores) {
        lines.push(`### ${store.name} (${store.type})`);
        if (store.owner?.name) lines.push(`**Owner:** ${store.owner.name}`);
        if (store.income?.dailyAvg) lines.push(`**Daily avg:** ${store.income.dailyAvg} gp`);
        if (store.inventory?.length) {
          lines.push('**Inventory:**');
          for (const item of store.inventory) {
            lines.push(`- ${item.name}: ${item.price} gp (×${item.stock})`);
          }
        }
      }
      lines.push('');
    }

    if (s.leadership?.length) {
      lines.push('## Leadership');
      for (const l of s.leadership) {
        lines.push(`- **${l.title}:** ${l.name}${l.role ? `  —  ${l.role}` : ''}`);
      }
      lines.push('');
    }

    const ranks = s.military?.ranks || [];
    if (ranks.length) {
      lines.push('## Military');
      for (const r of ranks) {
        lines.push(`- ${r.rank}: ${r.count}${r.leaderName ? `  (${r.leaderName})` : ''}`);
      }
      lines.push('');
    }

    if (s.notes) {
      lines.push('## Notes');
      lines.push(s.notes.trim());
      lines.push('');
    }

    const md   = lines.join('\n');
    const blob = new Blob([md], { type: 'text/markdown;charset=utf-8' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `${name.replace(/[^\w\s-]/g, '').trim().replace(/\s+/g, '_') || 'settlement'}.md`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
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

/**
 * GM-only QA helper (#122): creates a temporary fixture settlement, opens the
 * sheet for inspection, and deletes the journal when the sheet is closed.
 */
export async function openWithFixture() {
  const fixture = {
    kind: 'city', size: 'city', population: 12000, biome: 'temperate',
    stats: { hp: 180, maxHp: 180, damageThreshold: 20, hardness: 8, fortitude: 18, reflex: 14, will: 16, morale: 72, unrest: 12 },
    treasury: { cp: 250, sp: 800, gp: 4200, pp: 30 },
    production: ['grain', 'timber', 'iron', 'cloth'],
    government: { type: 'Council', leaderName: 'High Councillor Mira Ashvale' },
    military: {
      ranks: [
        { rank: 'City Watch',  count: 120, leaderName: 'Captain Brand',    dailyWage: 2 },
        { rank: 'Elite Guard', count: 20,  leaderName: 'Commander Syla',   dailyWage: 5 },
      ],
      totalGuards: 140, commanderName: 'Marshal Edric Stone',
    },
    stores: [
      {
        id: 'fix-shop-1', name: 'The Burnished Blade', type: 'blacksmith',
        owner: { name: 'Rolf Hammerborn', actorId: null },
        staff: [{ id: 'fix-staff-1', name: 'Apprentice Tane', role: 'smith', shift: 'day', actorId: null }],
        hours: { open: '07:00', close: '18:00', daysClosed: [] },
        inventory: [
          { id: 'fix-inv-1', name: 'Longsword',   price: 10, stock: 5, itemId: null },
          { id: 'fix-inv-2', name: 'Chain Shirt', price: 25, stock: 2, itemId: null },
        ],
        income: { balance: 320, dailyAvg: 18, lastTick: 0, daysInDebt: 0 },
        priceTier: 'standard', isBlackMarket: false,
      },
      {
        id: 'fix-shop-2', name: 'The Gilded Flask', type: 'alchemist',
        owner: { name: 'Priya Vellum', actorId: null },
        staff: [],
        hours: { open: '09:00', close: '20:00', daysClosed: [] },
        inventory: [
          { id: 'fix-inv-3', name: "Elixir of Life (Minor)", price: 3,  stock: 8,  itemId: null },
          { id: 'fix-inv-4', name: "Alchemist's Fire",       price: 3,  stock: 12, itemId: null },
        ],
        income: { balance: 150, dailyAvg: 22, lastTick: 0, daysInDebt: 0 },
        priceTier: 'high', isBlackMarket: false,
      },
    ],
    leadership: [
      { title: 'High Councillor', name: 'Mira Ashvale',   role: 'Civilian governance',  actorId: null },
      { title: 'Treasurer',       name: 'Owyn Brightcoin', role: 'Treasury & taxation', actorId: null },
    ],
    religions:    [{ id: 'fix-rel-1', name: 'Temple of Sarenrae', followers: 3200, influence: 45, templeStoreId: null }],
    demographics: [{ ancestry: 'Human', pct: 68 }, { ancestry: 'Halfling', pct: 14 }, { ancestry: 'Dwarf', pct: 10 }, { ancestry: 'Other', pct: 8 }],
    notes: 'Fixture settlement for QA. Data is illustrative — delete this journal when done.',
  };

  const doc = await JournalEntry.create({
    name: '[QA Fixture] Greymoor City',
    flags: { [FLAG_SCOPE]: { [FLAG_KEY]: fixture } },
  });
  if (!doc) return;

  const sheet      = new SettlementSheet(doc);
  const origClose  = sheet.close.bind(sheet);
  sheet.close      = async (...args) => { await origClose(...args); await doc.delete().catch(() => {}); };
  sheet.render(true);
}
