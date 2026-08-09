/**
 * ArmySheet (#70) — sheet for the standalone Army JournalEntry kind.
 * Shows the unit roster (type + count), total troop count, and a
 * stationed-at link to a settlement.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';
import { sanitizeArmy, totalUnitCount, UNIT_TYPES, UNIT_COSTS, recruitmentCost, computeArrivalDate,
         totalDailyFoodNeed, settlementFoodCapacity, isSupplied } from './army.js';
import { sanitizeSettlement } from './sanitizer.js';
import { resolveBattle, applyBattleResult, postBattleChatCard,
         resolveSiege, applySiegeResult, postSiegeChatCard, TERRAIN_TYPES } from './battle.js';

function formatCalendarDate(date, cal) {
  const m = cal?.monthNames?.[date.month - 1] || `M${date.month}`;
  return `${m} ${date.day}, ${date.year}`;
}

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

function commanderLevel(actorId) {
  return Number(resolveActor(actorId)?.level) || 0;
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
      siegeSettlement: function()  { this._onSiegeSettlement(); },
      openCommander:  function()   { this._onOpenCommander(); },
      clearCommander: function()   { this._onClearCommander(); },
      disbandCompany: function()   { this._onDisbandCompany(); },
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

    const availableNations = (game.journal?.contents || [])
      .filter(j => getSettlement(j)?.kind === 'nation')
      .map(j => ({ id: j.id, name: j.name }));

    let supplyStatus = null;
    if (army.supplySource) {
      const sourceDoc = game.journal?.get(army.supplySource);
      const sourceSettlement = sourceDoc ? sanitizeSettlement(getSettlement(sourceDoc) || {}) : null;
      supplyStatus = {
        sourceName: sourceDoc?.name || 'Unknown',
        need:     totalDailyFoodNeed(army),
        capacity: sourceSettlement ? settlementFoodCapacity(sourceSettlement) : 0,
        supplied: isSupplied(army, sourceSettlement),
      };
    }

    let arrivalDateLabel = '';
    if (army.destination && army.arrivalDate) {
      let calendarDef = null;
      try { calendarDef = game.settings.get('Pf2eCalendarTimeline', 'state')?.calendarDef; } catch (_) {}
      arrivalDateLabel = formatCalendarDate(army.arrivalDate, calendarDef);
    }

    let contractExpiresLabel = '';
    if (army.contract?.active && army.contract.expiresDate) {
      let calendarDef = null;
      try { calendarDef = game.settings.get('Pf2eCalendarTimeline', 'state')?.calendarDef; } catch (_) {}
      contractExpiresLabel = formatCalendarDate(army.contract.expiresDate, calendarDef);
    }

    return {
      doc: this.document,
      army,
      unitTypes: UNIT_TYPES,
      totalUnits: totalUnitCount(army),
      stationedAtName: stationedAtDoc?.name || '',
      destinationName: destinationDoc?.name || '',
      arrivalDateLabel,
      contractExpiresLabel,
      availableJournals,
      availableNations,
      supplyStatus,
      commanderActor: resolveActor(army.commanderActorId),
    };
  }

  _onRender() {
    this.element.classList.toggle('pf2e-high-contrast', !!game.settings.get(MODULE_ID, 'highContrastTheme'));

    const commanderZone = this.element.querySelector('[data-commander-drop]');
    if (commanderZone) {
      commanderZone.addEventListener('dragover',  (ev) => { ev.preventDefault(); ev.dataTransfer.dropEffect = 'link'; commanderZone.classList.add('drag-over'); });
      commanderZone.addEventListener('dragleave', ()   => commanderZone.classList.remove('drag-over'));
      commanderZone.addEventListener('drop',      (ev) => { commanderZone.classList.remove('drag-over'); this._onDropCommander(ev); });
    }
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

  async _onDropCommander(ev) {
    ev.preventDefault();
    let data;
    try { data = JSON.parse(ev.dataTransfer.getData('text/plain')); } catch { return; }
    if (data.type !== 'Actor') return;
    const actor = await fromUuid(data.uuid).catch(() => null);
    if (!actor) return;
    this._patch(a => { a.commanderActorId = actor.id; });
  }

  _onOpenCommander() {
    const army = sanitizeArmy(getSettlement(this.document) || {});
    game.actors?.get(army.commanderActorId)?.sheet?.render(true);
  }

  _onClearCommander() {
    this._patch(a => { a.commanderActorId = null; });
  }

  async _onDisbandCompany() {
    const confirmed = await foundry.applications.api.DialogV2.confirm({
      window:      { title: 'Disband Mercenary Company' },
      content:     `<p>Disband ${this.document.name} early? Their contract will not be refunded.</p>`,
      rejectClose: false,
    }).catch(() => false);
    if (!confirmed) return;
    await this.document.delete();
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

    const result = resolveBattle(
      attackerArmy, defenderArmy, picked.terrain,
      commanderLevel(attackerArmy.commanderActorId), commanderLevel(defenderArmy.commanderActorId),
    );
    await applyBattleResult(this.document, defenderDoc, result);
    postBattleChatCard(this.document, defenderDoc, result);

    const winnerName = result.winner === 'draw' ? 'Neither side' : (result.winner === 'attacker' ? this.document.name : defenderDoc.name);
    ui.notifications?.info?.(`Battle resolved — ${result.winner === 'draw' ? 'a stalemate' : `${winnerName} prevailed`}.`);
    this.render(false);
  }

  async _onSiegeSettlement() {
    const attackerArmy = sanitizeArmy(getSettlement(this.document) || {});
    if (!totalUnitCount(attackerArmy)) { ui.notifications?.warn?.('This army has no units to besiege with.'); return; }

    const candidates = (game.journal?.contents || [])
      .filter(j => { const s = getSettlement(j); return s && ['city', 'town', 'village'].includes(s.kind); })
      .map(j => ({ id: j.id, name: j.name }));
    if (!candidates.length) { ui.notifications?.warn?.('No settlements to besiege.'); return; }

    const html = `
      <div style="display:flex;flex-direction:column;gap:0.5em;">
        <label>Target settlement
          <select name="targetId" style="width:100%;margin-top:0.25em;">
            ${candidates.map(c => `<option value="${c.id}">${c.name}</option>`).join('')}
          </select>
        </label>
        <label>Terrain
          <select name="terrain" style="width:100%;margin-top:0.25em;">
            ${TERRAIN_TYPES.map(t => `<option value="${t}" ${t === 'urban' ? 'selected' : ''}>${t}</option>`).join('')}
          </select>
        </label>
      </div>`;
    const picked = await foundry.applications.api.DialogV2.prompt({
      window: { title: 'Siege Settlement' },
      content: html,
      ok: {
        label: 'Attack',
        callback: (_e, _b, dlg) => {
          const root = dlg.element;
          return {
            targetId: root.querySelector('[name="targetId"]')?.value || null,
            terrain: root.querySelector('[name="terrain"]')?.value || 'urban',
          };
        },
      },
      rejectClose: false,
    }).catch(() => null);
    if (!picked?.targetId) return;

    const targetDoc = game.journal?.get(picked.targetId);
    if (!targetDoc) return;
    const settlement = getSettlement(targetDoc) || {};

    const result = resolveSiege(attackerArmy, settlement, picked.terrain, commanderLevel(attackerArmy.commanderActorId));
    await applySiegeResult(targetDoc, result, attackerArmy.ownerNationId);
    const occupierNationName = attackerArmy.ownerNationId ? game.journal?.get(attackerArmy.ownerNationId)?.name : null;
    postSiegeChatCard(this.document, targetDoc, result, occupierNationName);

    ui.notifications?.info?.(result.occupied
      ? `${targetDoc.name} has fallen! ${result.damage} damage dealt.`
      : `Siege dealt ${result.damage} damage to ${targetDoc.name} (${result.hpAfter}/${settlement.stats?.maxHp ?? '?'} HP left).`);
    this.render(false);
  }
}
