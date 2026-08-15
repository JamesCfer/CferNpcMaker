/**
 * NationSheet — top-level sheet that aggregates child cities (linked by id)
 * and lets the user track nation-wide leadership, treasury, and population.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';
import { sanitizeSettlement }                              from './sanitizer.js';
import { computeArrivalDate }                              from './army.js';
import { gmWhisper }                                        from './diplomacy.js';

const { HandlebarsApplicationMixin, ApplicationV2 } = foundry.applications.api;

export class NationSheet extends HandlebarsApplicationMixin(ApplicationV2) {
  static DEFAULT_OPTIONS = {
    id: 'nation-sheet-{id}',
    classes: ['pf2e-nation-sheet'],
    tag: 'form',
    window: { resizable: true },
    position: { width: 820, height: 640 },
    actions: {
      addChildCity:    function()   { this._onAddChildCity(); },
      removeChildCity: function(ev) { this._onRemoveChildCity(ev); },
      openChildCity:   function(ev) { this._onOpenChildCity(ev); },
      saveField:       function(ev) { this._onSaveField(ev); },
      saveNotes:       function()   { this._onSaveNotes(); },
      setRelation:     function(ev) { this._onSetRelation(ev); },
      addTreaty:       function()   { this._onAddTreaty(); },
      removeTreaty:    function(ev) { this._onRemoveTreaty(ev); },
      addVassal:       function()   { this._onAddVassal(); },
      removeVassal:    function(ev) { this._onRemoveVassal(ev); },
      leaveSuzerain:   function()   { this._onLeaveSuzerain(); },
      addClaim:        function()   { this._onAddClaim(); },
      removeClaim:     function(ev) { this._onRemoveClaim(ev); },
      sendGift:        function()   { this._onSendGift(); },
    },
  };

  static PARTS = {
    body: { template: `modules/${MODULE_ID}/templates/nation-sheet.hbs` },
  };

  constructor(document, options = {}) {
    super(options);
    this.document = document;
  }

  get title() { return `${this.document?.name || 'Nation'} — Nation Sheet`; }
  get id()    { return `nation-sheet-${this.document?.id || 'unknown'}`; }

  async _prepareContext() {
    const nation = sanitizeSettlement(getSettlement(this.document) || {});

    const cityDocs = (nation.childCityIds || [])
      .map(id => game.journal?.get(id))
      .filter(Boolean);

    const cities = cityDocs.map(doc => {
      const cs = sanitizeSettlement(getSettlement(doc) || {});
      return {
        id: doc.id, name: doc.name,
        kind: cs.kind, population: cs.population,
        hpPct: Math.round((cs.stats.hp / Math.max(1, cs.stats.maxHp)) * 100),
        treasuryGp: cs.treasury.gp || 0,
        guards: cs.military?.totalGuards || 0,
      };
    });

    const totals = cities.reduce((acc, c) => {
      acc.population += c.population;
      acc.treasuryGp += c.treasuryGp;
      acc.guards     += c.guards;
      return acc;
    }, { population: 0, treasuryGp: 0, guards: 0 });

    const availableJournals = game.journal?.contents
      .filter(j => {
        const s = getSettlement(j);
        return s && s.kind !== 'nation' && !nation.childCityIds.includes(j.id);
      })
      .map(j => ({ id: j.id, name: j.name })) || [];

    const otherNations = (game.journal?.contents || [])
      .filter(j => j.id !== this.document.id && getSettlement(j)?.kind === 'nation')
      .map(j => ({ id: j.id, name: j.name }));

    const relationsView = otherNations.map(n => {
      const entry = nation.relations.find(r => r.nationId === n.id);
      return { id: n.id, name: n.name, relation: entry?.relation || 'neutral', score: entry?.score ?? 0 };
    });

    const treatiesView = (nation.treaties || []).map(t => ({
      ...t,
      partnerName:  game.journal?.get(t.partnerNationId)?.name || 'Unknown nation',
      signedLabel:  t.signedOn  ? `${t.signedOn.year}-${t.signedOn.month}-${t.signedOn.day}`   : '—',
      expiresLabel: t.expiresOn ? `${t.expiresOn.year}-${t.expiresOn.month}-${t.expiresOn.day}` : 'No expiry',
    }));

    // Vassal roll-up (#86): fold each vassal's own city totals into the nation's summary totals.
    const vassalDocs = (nation.vassalNationIds || []).map(id => game.journal?.get(id)).filter(Boolean);
    const vassals = vassalDocs.map(vDoc => {
      const vs = sanitizeSettlement(getSettlement(vDoc) || {});
      const vCityTotals = (vs.childCityIds || [])
        .map(id => game.journal?.get(id))
        .filter(Boolean)
        .reduce((acc, c) => {
          const cs = sanitizeSettlement(getSettlement(c) || {});
          acc.population += cs.population;
          acc.treasuryGp += cs.treasury.gp || 0;
          acc.guards     += cs.military?.totalGuards || 0;
          return acc;
        }, { population: 0, treasuryGp: 0, guards: 0 });
      return { id: vDoc.id, name: vDoc.name, ...vCityTotals };
    });
    for (const v of vassals) {
      totals.population += v.population;
      totals.treasuryGp += v.treasuryGp;
      totals.guards     += v.guards;
    }

    const suzerainDoc = nation.suzerainNationId ? game.journal?.get(nation.suzerainNationId) : null;
    const availableVassalCandidates = otherNations.filter(n =>
      !nation.vassalNationIds.includes(n.id) && n.id !== nation.suzerainNationId
    );

    // Claim targets (#87): any non-nation settlement this nation doesn't already hold as a city.
    const claimCandidates = (game.journal?.contents || [])
      .filter(j => {
        const s = getSettlement(j);
        return s && s.kind !== 'nation' && !nation.childCityIds.includes(j.id);
      })
      .map(j => ({ id: j.id, name: j.name }));

    const claimsView = (nation.claims || []).map(c => ({
      ...c,
      targetName: game.journal?.get(c.targetSettlementId)?.name || 'Unknown settlement',
    }));

    return {
      doc: this.document, nation, cities, totals, availableJournals,
      otherNations, relationsView, treatiesView, vassals, availableVassalCandidates,
      suzerainDoc: suzerainDoc ? { id: suzerainDoc.id, name: suzerainDoc.name } : null,
      claimCandidates, claimsView,
    };
  }

  _onRender() {
    this.element.classList.toggle('pf2e-high-contrast', !!game.settings.get(MODULE_ID, 'highContrastTheme'));
  }

  async _patch(mutator) {
    const cur = foundry.utils.deepClone(getSettlement(this.document) || {});
    mutator(cur);
    await this.document.setFlag(FLAG_SCOPE, FLAG_KEY, cur);
    this.render(false);
  }

  async _onAddChildCity() {
    const sel = this.element.querySelector('[name="addCityId"]');
    const id = sel?.value;
    if (!id) return;
    await this._patch(s => {
      s.childCityIds = s.childCityIds || [];
      if (!s.childCityIds.includes(id)) s.childCityIds.push(id);
    });
  }

  async _onRemoveChildCity(ev) {
    const id = ev.currentTarget?.dataset?.cityId;
    if (!id) return;
    await this._patch(s => { s.childCityIds = (s.childCityIds || []).filter(x => x !== id); });
  }

  _onOpenChildCity(ev) {
    const id = ev.currentTarget?.dataset?.cityId;
    const journal = game.journal?.get(id);
    journal?.sheet?.render(true);
  }

  _onSaveField(ev) {
    const path = ev.currentTarget?.dataset?.path;
    if (!path) return;
    let value = ev.currentTarget.value;
    if (ev.currentTarget.type === 'number') value = Number(value);
    this._patch(s => foundry.utils.setProperty(s, path, value));
  }

  _onSaveNotes() {
    const ta = this.element.querySelector('[data-path="notes"]');
    if (!ta) return;
    this._patch(s => { s.notes = ta.value; });
  }

  _onSetRelation(ev) {
    const nationId = ev.currentTarget?.dataset?.nationId;
    const field    = ev.currentTarget?.dataset?.field;
    if (!nationId || !field) return;
    const value = field === 'score' ? Number(ev.currentTarget.value) : ev.currentTarget.value;
    this._patch(s => {
      s.relations = s.relations || [];
      let entry = s.relations.find(r => r.nationId === nationId);
      if (!entry) { entry = { nationId, relation: 'neutral', score: 0 }; s.relations.push(entry); }
      entry[field] = value;
    });
  }

  _getCalendarState() {
    try {
      if (!game.modules?.get('Pf2eCalendarTimeline')?.active) return null;
      return game.settings.get('Pf2eCalendarTimeline', 'state') || null;
    } catch (_) { return null; }
  }

  async _onAddTreaty() {
    const root = this.element;
    const partnerNationId = root.querySelector('[name="treatyPartnerId"]')?.value;
    const kind             = root.querySelector('[name="treatyKind"]')?.value;
    const terms             = root.querySelector('[name="treatyTerms"]')?.value?.trim() || '';
    const durationDays     = Number(root.querySelector('[name="treatyDuration"]')?.value) || 0;
    if (!partnerNationId || !kind) return;

    const calState = this._getCalendarState();
    const signedOn = calState?.currentDate
      ? { year: calState.currentDate.year, month: calState.currentDate.month, day: calState.currentDate.day }
      : { year: 1, month: 1, day: 1 };
    const expiresOn = durationDays > 0 ? computeArrivalDate(signedOn, durationDays, calState?.calendarDef) : null;

    await this._patch(s => {
      s.treaties = s.treaties || [];
      s.treaties.push({
        id: `treaty-${Math.random().toString(36).slice(2, 10)}`,
        partnerNationId, kind, terms, signedOn, expiresOn,
      });
    });
  }

  async _onRemoveTreaty(ev) {
    const id = ev.currentTarget?.dataset?.treatyId;
    if (!id) return;
    await this._patch(s => { s.treaties = (s.treaties || []).filter(t => t.id !== id); });
  }

  async _onAddVassal() {
    const sel = this.element.querySelector('[name="addVassalId"]');
    const id = sel?.value;
    if (!id) return;
    await this._patch(s => {
      s.vassalNationIds = s.vassalNationIds || [];
      if (!s.vassalNationIds.includes(id)) s.vassalNationIds.push(id);
    });
    const vassalDoc = game.journal?.get(id);
    if (vassalDoc) {
      const vs = foundry.utils.deepClone(getSettlement(vassalDoc) || {});
      vs.suzerainNationId = this.document.id;
      vassalDoc.setFlag(FLAG_SCOPE, FLAG_KEY, vs).catch(() => {});
    }
  }

  async _onRemoveVassal(ev) {
    const id = ev.currentTarget?.dataset?.vassalId;
    if (!id) return;
    await this._patch(s => { s.vassalNationIds = (s.vassalNationIds || []).filter(x => x !== id); });
    const vassalDoc = game.journal?.get(id);
    if (vassalDoc) {
      const vs = foundry.utils.deepClone(getSettlement(vassalDoc) || {});
      if (vs.suzerainNationId === this.document.id) {
        vs.suzerainNationId = null;
        vassalDoc.setFlag(FLAG_SCOPE, FLAG_KEY, vs).catch(() => {});
      }
    }
  }

  async _onLeaveSuzerain() {
    const s = sanitizeSettlement(getSettlement(this.document) || {});
    const suzerainId = s.suzerainNationId;
    if (!suzerainId) return;
    await this._patch(cur => { cur.suzerainNationId = null; });
    const suzerainDoc = game.journal?.get(suzerainId);
    if (suzerainDoc) {
      const ss = foundry.utils.deepClone(getSettlement(suzerainDoc) || {});
      ss.vassalNationIds = (ss.vassalNationIds || []).filter(x => x !== this.document.id);
      suzerainDoc.setFlag(FLAG_SCOPE, FLAG_KEY, ss).catch(() => {});
    }
  }

  async _onAddClaim() {
    const root = this.element;
    const targetSettlementId = root.querySelector('[name="claimTargetId"]')?.value;
    const kind  = root.querySelector('[name="claimKind"]')?.value;
    const notes = root.querySelector('[name="claimNotes"]')?.value?.trim() || '';
    if (!targetSettlementId || !kind) return;

    await this._patch(s => {
      s.claims = s.claims || [];
      s.claims.push({
        id: `claim-${Math.random().toString(36).slice(2, 10)}`,
        targetSettlementId, kind, notes,
      });
    });
  }

  async _onRemoveClaim(ev) {
    const id = ev.currentTarget?.dataset?.claimId;
    if (!id) return;
    await this._patch(s => { s.claims = (s.claims || []).filter(c => c.id !== id); });
  }

  /** Transfers gp to a target nation and nudges relation scores both ways (#92). */
  async _onSendGift() {
    const root = this.element;
    const targetNationId = root.querySelector('[name="giftTargetId"]')?.value;
    const amount = Number(root.querySelector('[name="giftAmount"]')?.value) || 0;
    if (!targetNationId || amount <= 0) return;

    const targetDoc = game.journal?.get(targetNationId);
    if (!targetDoc) return;

    const nation = sanitizeSettlement(getSettlement(this.document) || {});
    if (amount > (nation.treasury.gp || 0)) return;

    const RELATION_BUMP = 5;

    await this._patch(s => {
      s.treasury = s.treasury || { cp: 0, sp: 0, gp: 0, pp: 0 };
      s.treasury.gp = (s.treasury.gp || 0) - amount;
      s.relations = s.relations || [];
      let entry = s.relations.find(r => r.nationId === targetNationId);
      if (!entry) { entry = { nationId: targetNationId, relation: 'neutral', score: 0 }; s.relations.push(entry); }
      entry.score = Math.min(100, entry.score + RELATION_BUMP);
    });

    const ts = foundry.utils.deepClone(getSettlement(targetDoc) || {});
    ts.treasury = ts.treasury || {};
    ts.treasury.gp = (ts.treasury.gp || 0) + amount;
    ts.relations = ts.relations || [];
    let targetEntry = ts.relations.find(r => r.nationId === this.document.id);
    if (!targetEntry) { targetEntry = { nationId: this.document.id, relation: 'neutral', score: 0 }; ts.relations.push(targetEntry); }
    targetEntry.score = Math.min(100, targetEntry.score + RELATION_BUMP);
    await targetDoc.setFlag(FLAG_SCOPE, FLAG_KEY, ts);

    ChatMessage.create({
      content: `<h3><i class="fa-solid fa-hand-holding-dollar"></i> Diplomatic Gift</h3>
        <p><strong>${this.document.name}</strong> sends <strong>${amount} gp</strong> to
        <strong>${targetDoc.name}</strong>, improving relations.</p>`,
      whisper: gmWhisper(),
    }).catch(() => {});
    Hooks.callAll('Pf2eNationsAndCitiesMaker.giftSent', {
      nationId: this.document.id, targetNationId, amount,
    });
  }
}

export function maybeOpenNationSheet(journal) {
  const s = getSettlement(journal);
  if (!s) return false;
  if (s.kind !== 'nation') return false;
  const sheet = new NationSheet(journal);
  sheet.render(true);
  return true;
}
