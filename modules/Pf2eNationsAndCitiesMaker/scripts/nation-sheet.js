/**
 * NationSheet — top-level sheet that aggregates child cities (linked by id)
 * and lets the user track nation-wide leadership, treasury, and population.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';
import { sanitizeSettlement }                              from './sanitizer.js';

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

    return { doc: this.document, nation, cities, totals, availableJournals };
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
}

export function maybeOpenNationSheet(journal) {
  const s = getSettlement(journal);
  if (!s) return false;
  if (s.kind !== 'nation') return false;
  const sheet = new NationSheet(journal);
  sheet.render(true);
  return true;
}
