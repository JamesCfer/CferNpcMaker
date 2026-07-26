/**
 * ArmySheet (#70) — sheet for the standalone Army JournalEntry kind.
 * Shows the unit roster (type + count), total troop count, and a
 * stationed-at link to a settlement.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';
import { sanitizeArmy, totalUnitCount, UNIT_TYPES, UNIT_COSTS, recruitmentCost } from './army.js';

const { HandlebarsApplicationMixin, ApplicationV2 } = foundry.applications.api;

export class ArmySheet extends HandlebarsApplicationMixin(ApplicationV2) {
  static DEFAULT_OPTIONS = {
    id: 'army-sheet-{id}',
    classes: ['pf2e-army-sheet'],
    tag: 'form',
    window: { resizable: true },
    position: { width: 620, height: 560 },
    actions: {
      saveField:      function(ev) { this._onSaveField(ev); },
      saveNotes:      function()   { this._onSaveNotes(); },
      addUnit:        function()   { this._onAddUnit(); },
      removeUnit:     function(ev) { this._onRemoveUnit(ev); },
      openStationedAt: function()  { this._onOpenStationedAt(); },
      recruitUnits:   function()   { this._onRecruitUnits(); },
    },
  };

  static PARTS = {
    body: { template: `modules/${MODULE_ID}/templates/army-sheet.hbs` },
  };

  constructor(document, options = {}) {
    super(options);
    this.document = document;
  }

  get title() { return `${this.document?.name || 'Army'} — Army Sheet`; }
  get id()    { return `army-sheet-${this.document?.id || 'unknown'}`; }

  async _prepareContext() {
    const army = sanitizeArmy(getSettlement(this.document) || {});
    const stationedAtDoc = army.stationedAt ? game.journal?.get(army.stationedAt) : null;

    const availableJournals = (game.journal?.contents || [])
      .filter(j => {
        const s = getSettlement(j);
        return s && ['city', 'town', 'village'].includes(s.kind);
      })
      .map(j => ({ id: j.id, name: j.name }));

    return {
      doc: this.document,
      army,
      unitTypes: UNIT_TYPES,
      totalUnits: totalUnitCount(army),
      stationedAtName: stationedAtDoc?.name || '',
      availableJournals,
    };
  }

  _onRender() {
    this.element.classList.toggle('pf2e-high-contrast', !!game.settings.get(MODULE_ID, 'highContrastTheme'));
  }

  async _patch(mutator) {
    const cur = foundry.utils.deepClone(getSettlement(this.document) || {});
    mutator(cur);
    await this.document.setFlag(FLAG_SCOPE, FLAG_KEY, sanitizeArmy(cur));
    this.render(false);
  }

  _onSaveField(ev) {
    const path = ev.currentTarget?.dataset?.path;
    if (!path) return;
    let value = ev.currentTarget.value;
    if (ev.currentTarget.type === 'number') value = Number(value);
    if (value === '') value = null;
    this._patch(s => foundry.utils.setProperty(s, path, value));
  }

  _onSaveNotes() {
    const ta = this.element.querySelector('[data-path="notes"]');
    if (!ta) return;
    this._patch(s => { s.notes = ta.value; });
  }

  _onAddUnit() {
    this._patch(s => {
      s.units = s.units || [];
      s.units.push({ type: 'spearmen', count: 1, level: 1, equipment: '', morale: 100 });
    });
  }

  _onRemoveUnit(ev) {
    const id = ev.currentTarget?.dataset?.unitId;
    if (!id) return;
    this._patch(s => { s.units = (s.units || []).filter(u => u.id !== id); });
  }

  _onOpenStationedAt() {
    const army = sanitizeArmy(getSettlement(this.document) || {});
    const journal = army.stationedAt && game.journal?.get(army.stationedAt);
    if (!journal) { ui.notifications?.warn?.('This army is not stationed at a settlement.'); return; }
    journal.sheet.render(true);
  }

  async _onRecruitUnits() {
    const army = sanitizeArmy(getSettlement(this.document) || {});
    const settlementDoc = army.stationedAt && game.journal?.get(army.stationedAt);
    if (!settlementDoc) { ui.notifications?.warn?.('Station this army at a settlement before recruiting.'); return; }

    const settlement = getSettlement(settlementDoc) || {};
    const pop = Number(settlement.population) || 0;
    const gp = Number(settlement.treasury?.gp) || 0;

    const html = `
      <p>Recruiting from <strong>${settlementDoc.name}</strong> — pool: ${pop.toLocaleString()} population,
      ${gp.toLocaleString()} gp.</p>
      <div style="display:flex;flex-direction:column;gap:0.5em;">
        <label>Unit type
          <select name="type" style="width:100%;margin-top:0.25em;">
            ${UNIT_TYPES.map(t => `<option value="${t}">${t} (${UNIT_COSTS[t].gp} gp, ${UNIT_COSTS[t].pop} pop each)</option>`).join('')}
          </select>
        </label>
        <label>Count
          <input type="number" name="count" value="1" min="1" step="1" style="width:100%;margin-top:0.25em;" />
        </label>
      </div>`;
    const result = await foundry.applications.api.DialogV2.prompt({
      window: { title: 'Recruit Units' },
      content: html,
      ok: {
        label: 'Recruit',
        callback: (_e, _b, dlg) => {
          const root = dlg.element;
          return {
            type: root.querySelector('[name="type"]')?.value || 'spearmen',
            count: Math.max(1, Number(root.querySelector('[name="count"]')?.value) || 1),
          };
        },
      },
      rejectClose: false,
    }).catch(() => null);
    if (!result) return;

    const cost = recruitmentCost(result.type, result.count);
    if (cost.pop > pop || cost.gp > gp) {
      ui.notifications?.warn?.(`${settlementDoc.name} can't afford ${result.count} ${result.type} (needs ${cost.gp} gp, ${cost.pop} pop).`);
      return;
    }

    const s = foundry.utils.deepClone(settlement);
    s.population = pop - cost.pop;
    s.treasury = s.treasury || { cp: 0, sp: 0, gp: 0, pp: 0 };
    s.treasury.gp = gp - cost.gp;
    await settlementDoc.setFlag(FLAG_SCOPE, FLAG_KEY, s);

    this._patch(a => {
      a.units = a.units || [];
      const existing = a.units.find(u => u.type === result.type && Number(u.level) === 1);
      if (existing) existing.count = Number(existing.count || 0) + result.count;
      else a.units.push({ type: result.type, count: result.count, level: 1, equipment: '', morale: 100 });
    });

    ui.notifications?.info?.(`Recruited ${result.count} ${result.type} for ${cost.gp} gp.`);
  }
}
