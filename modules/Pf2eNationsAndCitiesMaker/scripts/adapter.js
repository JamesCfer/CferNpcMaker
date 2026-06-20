/**
 * SettlementAdapter — generates settlements and stores them as JournalEntries
 * with a settlement payload under flags.Pf2eNationsAndCitiesMaker.settlement.
 */

import { SystemAdapter, postToN8n } from './core/adapter.js';
import { N8N_BASE, devUrl }          from './core/n8n.js';
import { detectModuleFolder }        from './core/utils.js';
import { MODULE_ID, FLAG_SCOPE, FLAG_KEY } from './constants.js';
import { sanitizeSettlement }              from './sanitizer.js';
import { SettlementSheet }                 from './settlement-sheet.js';
import { NationSheet }                     from './nation-sheet.js';

const MODULE_FOLDER = detectModuleFolder(MODULE_ID);
const ENDPOINT      = `${N8N_BASE}/webhook/city-builder`;

const KIND_OPTIONS = ['city', 'town', 'village', 'nation'];
const SIZE_OPTIONS = ['thorp', 'hamlet', 'village', 'town', 'city', 'metropolis'];

export class SettlementAdapter extends SystemAdapter {
  get moduleFolder() { return MODULE_FOLDER; }

  get module() {
    return {
      id:           MODULE_ID,
      label:        'PF2e Settlement',
      icon:         'fa-solid fa-city',
      githubUrl:    'https://github.com/JamesCfer/Pf2eNationsAndCitiesMaker',
      historyLabel: 'Created Settlements',
    };
  }

  get systemId() { return 'pf2e'; }
  get supportsImageGeneration() { return false; }
  get formConfig() { return { documentNoun: 'settlement' }; }
  get progressSteps() {
    return ['Sending request…', 'Building town layout…', 'Generating stores…', 'Hiring guards…', 'Creating journal…'];
  }

  /* ── form ──────────────────────────────────────────────── */

  gatherFormData(form) {
    const fd = new FormData(form);
    const name           = (fd.get('name')?.toString()?.trim()) || 'New Settlement';
    const kindRaw        = (fd.get('kind')?.toString() || 'town').trim();
    const sizeRaw        = (fd.get('size')?.toString() || 'town').trim();
    const biome          = (fd.get('biome')?.toString()?.trim()) || '';
    const governmentHint = (fd.get('governmentHint')?.toString()?.trim()) || '';
    const populationHint = Number(fd.get('populationHint')) || 0;
    const description    = (fd.get('description')?.toString()?.trim()) || '';
    const includeStores     = fd.get('includeStores')     === 'on';
    const includeMilitary   = fd.get('includeMilitary')   === 'on';
    const includeLeadership = fd.get('includeLeadership') === 'on';
    const templateId        = (fd.get('template')?.toString() || 'custom').trim();

    if (!description) throw new Error('Please provide a description for the settlement.');

    return {
      name,
      kind: KIND_OPTIONS.includes(kindRaw) ? kindRaw : 'town',
      size: SIZE_OPTIONS.includes(sizeRaw) ? sizeRaw : 'town',
      biome, governmentHint, populationHint, description,
      includeStores, includeMilitary, includeLeadership,
      templateId,
    };
  }

  historyEntryFromForm(formData) {
    return {
      name: formData.name, kind: formData.kind, size: formData.size,
      biome: formData.biome, governmentHint: formData.governmentHint,
      populationHint: formData.populationHint, description: formData.description,
      includeStores: formData.includeStores, includeMilitary: formData.includeMilitary,
      includeLeadership: formData.includeLeadership, templateId: formData.templateId,
    };
  }

  historyMeta(entry) {
    const kind = (entry.kind || 'town');
    const cap = kind.charAt(0).toUpperCase() + kind.slice(1);
    const pop = entry.populationHint ? `&nbsp;·&nbsp;${entry.populationHint.toLocaleString()}` : '';
    return `${cap}${pop}`;
  }

  populateForm(form, entry) {
    const set = (sel, val) => { const el = form.querySelector(sel); if (el != null && val != null) el.value = val; };
    set('[name="name"]',           entry.name);
    set('[name="kind"]',           entry.kind);
    set('[name="size"]',           entry.size);
    set('[name="biome"]',          entry.biome);
    set('[name="governmentHint"]', entry.governmentHint);
    set('[name="populationHint"]', entry.populationHint);
    set('[name="description"]',    entry.description);
    set('[name="template"]',       entry.templateId);
    const setCheck = (sel, val) => { const el = form.querySelector(sel); if (el) el.checked = !!val; };
    setCheck('[name="includeStores"]',     entry.includeStores);
    setCheck('[name="includeMilitary"]',   entry.includeMilitary);
    setCheck('[name="includeLeadership"]', entry.includeLeadership);
  }

  /* ── generation ────────────────────────────────────────── */

  async generate({ formData, key, devMode }) {
    const endpoint = devUrl(ENDPOINT, devMode);
    const payload = {
      name:              formData.name,
      kind:              formData.kind,
      size:              formData.size,
      biome:             formData.biome,
      governmentHint:    formData.governmentHint,
      populationHint:    formData.populationHint,
      description:       formData.description,
      includeStores:     formData.includeStores,
      includeMilitary:   formData.includeMilitary,
      includeLeadership: formData.includeLeadership,
    };

    let settlementData;
    try {
      const { response, responseText } = await postToN8n(endpoint, payload, key);
      let data;
      try { data = JSON.parse(responseText); }
      catch (err) { throw new Error(`Invalid JSON response (${responseText.length} bytes): ${err.message}`); }
      if (!response.ok) throw new Error(data?.message || `Server returned status ${response.status}`);
      if (data?.ok === false) throw new Error(data?.message || data?.error || 'Server rejected the request');
      settlementData = data.settlement || data.foundrySettlement || data;
    } catch (netErr) {
      // For initial rollout, allow generation even when the AI endpoint isn't
      // wired up yet — fall back to a structured stub so the user can still
      // see the sheet work end-to-end.
      if (devMode || (game.settings?.get?.(MODULE_ID, 'allowOfflineStub') ?? true)) {
        console.warn(`[${MODULE_ID}] AI endpoint unavailable — using offline stub. ${netErr?.message || ''}`);
        settlementData = stubSettlementFromForm(formData);
      } else {
        throw netErr;
      }
    }

    const settlement = sanitizeSettlement(settlementData, formData);

    const journal = await JournalEntry.create({
      name: formData.name,
      flags: {
        [FLAG_SCOPE]: {
          [FLAG_KEY]: settlement,
          createdBy: MODULE_ID,
        },
      },
    });

    if (!journal) throw new Error('Failed to create the settlement journal entry.');

    // Open the matching custom sheet.
    if (settlement.kind === 'nation') {
      new NationSheet(journal).render(true);
    } else {
      new SettlementSheet(journal).render(true);
    }

    return {
      document: journal,
      exportData: {
        content:  JSON.stringify(settlement, null, 2),
        filename: `${journal.name || 'settlement'}.json`,
        mimeType: 'application/json',
      },
      message: `Settlement "${journal.name}" created.`,
    };
  }

  /**
   * Add a header button to the custom sheet for re-opening the Builder.
   * (The sheets themselves wire most actions; this is just a sidebar entry
   * point for journal directory headers.)
   */
  registerSheetHooks(_getApp) {
    // No-op: the custom sheets ship their own controls.
  }
}

/* ── Fallback stub used when the AI endpoint isn't reachable ─────────── */

function stubSettlementFromForm(formData) {
  const pop = formData.populationHint || (formData.kind === 'village' ? 250 : formData.kind === 'city' ? 8000 : 1500);
  const guards = Math.max(4, Math.round(pop / 90));
  return {
    kind: formData.kind, size: formData.size,
    population: pop, biome: formData.biome || 'temperate',
    stats: {
      maxHp: 100 + Math.round(pop / 25), hp: 100 + Math.round(pop / 25),
      damageThreshold: 10 + Math.round(pop / 800),
      hardness: 5 + Math.round(pop / 1500),
      fortitude: 10, reflex: 8, will: 10,
      morale: 65, unrest: 10,
    },
    treasury: { cp: 0, sp: 0, gp: Math.round(pop * 0.4), pp: 0 },
    production: ['grain', 'livestock'],
    government: {
      type: formData.governmentHint || 'Mayor + council',
      leaderName: 'TBD',
      leaderActorId: null,
    },
    leadership: formData.includeLeadership
      ? [{ title: 'Mayor', name: 'TBD', actorId: null, role: 'civic leader' }]
      : [],
    military: formData.includeMilitary ? {
      ranks: [
        { rank: 'Captain',  count: 1,                          leaderName: 'TBD' },
        { rank: 'Sergeant', count: Math.max(1, Math.round(guards / 10)), leaderName: '' },
        { rank: 'Guard',    count: guards,                     leaderName: '' },
      ],
      totalGuards: guards + 1 + Math.max(1, Math.round(guards / 10)),
      commanderActorId: null, commanderName: 'TBD',
    } : { ranks: [], totalGuards: 0, commanderActorId: null, commanderName: '' },
    stores: formData.includeStores ? [
      stubStore('The Whetstone', 'blacksmith'),
      stubStore('Greenleaf Provisions', 'general'),
      stubStore('The Tipsy Owl', 'tavern'),
    ] : [],
    notes: `Generated offline stub for "${formData.name}".`,
    ai: { endpoint: 'city-builder', model: 'stub', prompt: formData.description },
  };
}

function stubStore(name, type) {
  const id = `shop-${Math.random().toString(36).slice(2, 10)}`;
  return {
    id, name, type,
    owner: { name: 'Unnamed Owner', actorId: null },
    staff: [
      { id: `staff-${Math.random().toString(36).slice(2, 10)}`, name: 'Apprentice', role: 'apprentice', shift: 'day', actorId: null },
    ],
    hours: { open: '08:00', close: '20:00', daysClosed: [] },
    inventory: [],
    income: { balance: 50, dailyAvg: 6, lastTick: 0 },
  };
}
