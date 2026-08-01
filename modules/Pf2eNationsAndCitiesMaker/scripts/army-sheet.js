/**
 * ArmySheet (#70) — sheet for the standalone Army JournalEntry kind.
 * Shows the unit roster (type + count), total troop count, and a
 * stationed-at link to a settlement.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';
import { sanitizeArmy, totalUnitCount, UNIT_TYPES, UNIT_COSTS, recruitmentCost, computeArrivalDate } from './army.js';
import { resolveBattle, applyBattleResult, postBattleChatCard, TERRAIN_TYPES } from './battle.js';

function formatCalendarDate(date, cal) {
  const m = cal?.monthNames?.[date.month - 1] || `M${date.month}`;
  return `${m} ${date.day}, ${date.year}`;
}

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
      sendArmy:       function()   { this._onSendArmy(); },
      cancelMarch:    function()   { this._onCancelMarch(); },
      resolveBattle:  function()   { this._onResolveBattle(); },
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
    const destinationDoc = army.destination ? game.journal?.get(army.destination) : null;

    const availableJournals = (game.journal?.contents || [])
      .filter(j => {
        const s = getSettlement(j);
        return s && ['city', 'town', 'village'].includes(s.kind);
      })
      .map(j => ({ id: j.id, name: j.name }));

    let arrivalDateLabel = '';
    if (army.destination && army.arrivalDate) {
      let calendarDef = null;
      try { calendarDef = game.settings.get('Pf2eCalendarTimeline', 'state')?.calendarDef; } catch (_) {}
      arrivalDateLabel = formatCalendarDate(army.arrivalDate, calendarDef);
    }

    return {
      doc: this.document,
      army,
      unitTypes: UNIT_TYPES,
      totalUnits: totalUnitCount(army),
      stationedAtName: stationedAtDoc?.name || '',
      destinationName: destinationDoc?.name || '',
      arrivalDateLabel,
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

  async _onSendArmy() {
    const army = sanitizeArmy(getSettlement(this.document) || {});
    if (army.mode !== 'field') {
      ui.notifications?.warn?.('Switch this army to Field mode before sending it on campaign.');
      return;
    }
    if (!game.modules?.get('Pf2eCalendarTimeline')?.active) {
      ui.notifications?.warn?.('Enable Pf2eCalendarTimeline to compute travel time.');
      return;
    }

    const candidates = (game.journal?.contents || [])
      .filter(j => {
        const s = getSettlement(j);
        return s && ['city', 'town', 'village'].includes(s.kind) && j.id !== army.stationedAt;
      })
      .map(j => ({ id: j.id, name: j.name }));
    if (!candidates.length) { ui.notifications?.warn?.('No other settlements to march to.'); return; }

    const html = `
      <label>Destination
        <select name="destination" style="width:100%;margin-top:0.25em;">
          ${candidates.map(c => `<option value="${c.id}">${c.name}</option>`).join('')}
        </select>
      </label>`;
    const destination = await foundry.applications.api.DialogV2.prompt({
      window: { title: 'Send Army' },
      content: html,
      ok: {
        label: 'March',
        callback: (_e, _b, dlg) => dlg.element.querySelector('[name="destination"]')?.value || null,
      },
      rejectClose: false,
    }).catch(() => null);
    if (!destination) return;

    const calState = game.settings.get('Pf2eCalendarTimeline', 'state');
    const travelDays = game.settings.get(MODULE_ID, 'armyTravelDays');
    const arrivalDate = computeArrivalDate(calState.currentDate, travelDays, calState.calendarDef);

    await this._patch(a => { a.destination = destination; a.arrivalDate = arrivalDate; });
    ui.notifications?.info?.(`${this.document.name} is marching — arriving in ${travelDays} day(s).`);
  }

  _onCancelMarch() {
    this._patch(a => { a.destination = null; a.arrivalDate = null; });
  }

  async _onResolveBattle() {
    const attackerArmy = sanitizeArmy(getSettlement(this.document) || {});
    if (!totalUnitCount(attackerArmy)) { ui.notifications?.warn?.('This army has no units to fight with.'); return; }

    const candidates = (game.journal?.contents || [])
      .filter(j => { const s = getSettlement(j); return s?.kind === 'army' && j.id !== this.document.id; })
      .map(j => ({ id: j.id, name: j.name }));
    if (!candidates.length) { ui.notifications?.warn?.('No other armies to battle.'); return; }

    const html = `
      <div style="display:flex;flex-direction:column;gap:0.5em;">
        <label>Defending army
          <select name="defenderId" style="width:100%;margin-top:0.25em;">
            ${candidates.map(c => `<option value="${c.id}">${c.name}</option>`).join('')}
          </select>
        </label>
        <label>Terrain
          <select name="terrain" style="width:100%;margin-top:0.25em;">
            ${TERRAIN_TYPES.map(t => `<option value="${t}">${t}</option>`).join('')}
          </select>
        </label>
      </div>`;
    const picked = await foundry.applications.api.DialogV2.prompt({
      window: { title: 'Resolve Battle' },
      content: html,
      ok: {
        label: 'Resolve',
        callback: (_e, _b, dlg) => {
          const root = dlg.element;
          return {
            defenderId: root.querySelector('[name="defenderId"]')?.value || null,
            terrain: root.querySelector('[name="terrain"]')?.value || 'plains',
          };
        },
      },
      rejectClose: false,
    }).catch(() => null);
    if (!picked?.defenderId) return;

    const defenderDoc = game.journal?.get(picked.defenderId);
    if (!defenderDoc) return;
    const defenderArmy = sanitizeArmy(getSettlement(defenderDoc) || {});
    if (!totalUnitCount(defenderArmy)) { ui.notifications?.warn?.('The defending army has no units.'); return; }

    const result = resolveBattle(attackerArmy, defenderArmy, picked.terrain);
    await applyBattleResult(this.document, defenderDoc, result);
    postBattleChatCard(this.document, defenderDoc, result);

    const winnerName = result.winner === 'draw' ? 'Neither side' : (result.winner === 'attacker' ? this.document.name : defenderDoc.name);
    ui.notifications?.info?.(`Battle resolved — ${result.winner === 'draw' ? 'a stalemate' : `${winnerName} prevailed`}.`);
    this.render(false);
  }
}
