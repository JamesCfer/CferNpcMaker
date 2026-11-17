/**
 * Every generator the Roll20 extension offers: the form it shows, the n8n
 * webhook it calls (the same ones the Foundry modules call), what it costs,
 * and which mapper turns the answer into an import plan.
 */
import { mapDnd5eNpc, dnd5eSpellMapping } from './mappers/dnd5e-npc.js';
import { mapPf2eNpc, pf2eSpellMapping } from './mappers/pf2e-npc.js';
import { mapItem } from './mappers/items.js';
import { mapStore } from './mappers/store.js';
import { mapSettlement } from './mappers/settlement.js';
import { mapHero6e } from './mappers/hero6e.js';

export const N8N_BASE = 'https://foundryrelay.dedicated2.com';
export const ENDPOINTS = {
  authLogin: `${N8N_BASE}/webhook/oauth/patreon/login`,
  authPoll: `${N8N_BASE}/webhook/oauth/patreon/poll`,
  authValidate: `${N8N_BASE}/webhook/oauth/patreon/validate`,
  usageCheck: `${N8N_BASE}/webhook/usage-check`,
  feedback: `${N8N_BASE}/webhook/feedback`,
  mapBuilder: `${N8N_BASE}/webhook/map-builder`,
};
export const PATREON_URL = 'https://www.patreon.com/cw/CelestiaTools';

const STORE_TYPES = ['general', 'blacksmith', 'armorer', 'alchemist', 'magic', 'jeweler', 'fletcher', 'tavern'];
const SIZES = ['village', 'town', 'city', 'metropolis'];

/** @typedef {{name:string,label:string,type:'text'|'number'|'textarea'|'select'|'checkbox',options?:string[],min?:number,max?:number,value?:any,required?:boolean,placeholder?:string}} Field */

export const GENERATORS = [
  {
    id: 'dnd5e-npc', label: 'D&D 5e NPC', system: 'dnd5e', sheet: 'D&D 5E by Roll20', cost: 1, endpoint: '/webhook/dnd5e-npc-builder',
    fields: [
      { name: 'name', label: 'Name', type: 'text', placeholder: 'Generated Creature' },
      { name: 'level', label: 'Challenge rating', type: 'number', min: 0, max: 30, value: 1 },
      { name: 'casterType', label: 'Spellcasting', type: 'select', options: ['none', 'half', 'full'], value: 'none' },
      { name: 'description', label: 'Description', type: 'textarea', required: true, placeholder: 'A grizzled orc warlord who leads from the front…' },
    ],
    payload: f => ({ name: f.name || 'Generated Creature', cr: Number(f.level) || 1, description: f.description, casterType: f.casterType || 'none',
      ...(f.casterType && f.casterType !== 'none' ? { spellMapping: dnd5eSpellMapping() } : {}) }),
    map: mapDnd5eNpc,
  },
  {
    id: 'pf2e-npc', label: 'PF2e NPC', system: 'pf2e', sheet: 'Pathfinder Second Edition by Roll20', cost: 1, endpoint: '/webhook/npc-builder',
    fields: [
      { name: 'name', label: 'Name', type: 'text', placeholder: 'Generated Creature' },
      { name: 'level', label: 'Level', type: 'number', min: -1, max: 25, value: 1 },
      { name: 'includeSpells', label: 'Spellcaster', type: 'checkbox', value: false },
      { name: 'description', label: 'Description', type: 'textarea', required: true },
    ],
    payload: f => ({ name: f.name || 'Generated Creature', level: Number(f.level) || 1, description: f.description, ...(f.includeSpells ? { spellMapping: pf2eSpellMapping() } : {}) }),
    map: mapPf2eNpc,
  },
  {
    id: 'dnd5e-item', label: 'D&D 5e item', system: 'dnd5e', sheet: 'D&D 5E by Roll20', cost: 1, endpoint: '/webhook/dnd5e-item-builder',
    fields: [
      { name: 'name', label: 'Name', type: 'text' },
      { name: 'rarity', label: 'Rarity', type: 'select', options: ['', 'common', 'uncommon', 'rare', 'veryRare', 'legendary', 'artifact'], value: '' },
      { name: 'itemType', label: 'Type', type: 'select', options: ['weapon', 'equipment', 'consumable', 'tool', 'loot', 'container'], value: 'weapon' },
      { name: 'subtype', label: 'Subtype', type: 'text', placeholder: 'longsword, ring, potion…' },
      { name: 'description', label: 'Description', type: 'textarea', required: true },
    ],
    payload: f => ({ name: f.name, rarity: f.rarity, itemType: f.itemType, subtype: f.subtype, description: f.description }),
    map: mapItem('dnd5e'),
  },
  {
    id: 'pf2e-item', label: 'PF2e item', system: 'pf2e', sheet: 'Pathfinder Second Edition by Roll20', cost: 1, endpoint: '/webhook/pf2e-item-builder',
    fields: [
      { name: 'name', label: 'Name', type: 'text' },
      { name: 'level', label: 'Level', type: 'number', min: 0, max: 25, value: 1 },
      { name: 'itemType', label: 'Type', type: 'select', options: ['weapon', 'armor', 'shield', 'equipment', 'consumable', 'treasure', 'backpack', 'book'], value: 'weapon' },
      { name: 'description', label: 'Description', type: 'textarea', required: true },
    ],
    payload: f => ({ name: f.name, level: Number(f.level) || 1, itemType: f.itemType, description: f.description }),
    map: mapItem('pf2e'),
  },
  {
    id: 'dnd5e-store', label: 'D&D 5e store', system: 'dnd5e', sheet: 'D&D 5E by Roll20', cost: 7, endpoint: '/webhook/dnd5e-store-builder',
    fields: [
      { name: 'name', label: 'Store name', type: 'text' },
      { name: 'storeType', label: 'Store type', type: 'select', options: STORE_TYPES, value: 'general' },
      { name: 'wealth', label: 'Wealth', type: 'select', options: ['poor', 'modest', 'wealthy', 'aristocratic'], value: 'modest' },
      { name: 'settlementSize', label: 'Settlement', type: 'select', options: SIZES, value: 'town' },
      { name: 'description', label: 'Description', type: 'textarea' },
    ],
    payload: f => ({ name: f.name, storeType: f.storeType, wealth: f.wealth, settlementSize: f.settlementSize, description: f.description }),
    map: mapStore('dnd5e'),
  },
  {
    id: 'pf2e-store', label: 'PF2e store', system: 'pf2e', sheet: 'Pathfinder Second Edition by Roll20', cost: 7, endpoint: '/webhook/pf2e-store-builder',
    fields: [
      { name: 'name', label: 'Store name', type: 'text' },
      { name: 'level', label: 'Settlement level', type: 'number', min: 0, max: 20, value: 3 },
      { name: 'storeType', label: 'Store type', type: 'select', options: STORE_TYPES, value: 'general' },
      { name: 'settlementSize', label: 'Settlement', type: 'select', options: SIZES, value: 'town' },
      { name: 'description', label: 'Description', type: 'textarea' },
    ],
    payload: f => ({ name: f.name, level: Number(f.level) || 3, storeType: f.storeType, settlementSize: f.settlementSize, description: f.description }),
    map: mapStore('pf2e'),
  },
  {
    id: 'pf2e-settlement', label: 'Nation, city or town', system: 'pf2e', sheet: 'Handouts (any sheet)', cost: 1,
    endpoint: f => (f.kind === 'nation' && f.includeChildCities ? '/webhook/nation-builder' : '/webhook/city-builder'),
    fields: [
      { name: 'name', label: 'Name', type: 'text', required: true },
      { name: 'kind', label: 'Kind', type: 'select', options: ['village', 'town', 'city', 'nation'], value: 'town' },
      { name: 'size', label: 'Size', type: 'select', options: ['thorp', 'hamlet', 'village', 'town', 'city', 'metropolis'], value: 'town' },
      { name: 'biome', label: 'Biome', type: 'text', placeholder: 'temperate forest' },
      { name: 'populationHint', label: 'Population (optional)', type: 'number', min: 0 },
      { name: 'governmentHint', label: 'Government (optional)', type: 'text' },
      { name: 'includeStores', label: 'Include stores', type: 'checkbox', value: true },
      { name: 'includeMilitary', label: 'Include military', type: 'checkbox', value: true },
      { name: 'includeLeadership', label: 'Include leadership', type: 'checkbox', value: true },
      { name: 'includeChildCities', label: 'Nation: also make its cities', type: 'checkbox', value: false },
      { name: 'description', label: 'Description', type: 'textarea' },
    ],
    payload: f => ({ name: f.name, kind: f.kind, size: f.size, biome: f.biome, governmentHint: f.governmentHint, populationHint: f.populationHint,
      description: f.description, includeStores: !!f.includeStores, includeMilitary: !!f.includeMilitary, includeLeadership: !!f.includeLeadership }),
    map: mapSettlement,
  },
  {
    id: 'hero6e-npc', label: 'Hero 6e NPC', system: 'hero6e', sheet: 'HeroSystem6eHeroic', cost: 1, endpoint: '/webhook/hero6e-hdc-builder',
    fields: [
      { name: 'name', label: 'Name', type: 'text' },
      { name: 'level', label: 'Total points', type: 'number', min: 25, max: 600, value: 150 },
      { name: 'genre', label: 'Genre', type: 'select', options: ['standard', 'superhero', 'pulp', 'dark_champions', 'fantasy', 'sci-fi'], value: 'standard' },
      { name: 'universe', label: 'Universe', type: 'select', options: ['standard', 'mha', 'dc', 'marvel'], value: 'standard' },
      { name: 'createGear', label: 'Create gear', type: 'checkbox', value: true },
      { name: 'description', label: 'Description', type: 'textarea', required: true },
    ],
    payload: f => ({ name: f.name, points: Number(f.level) || 150, genre: f.genre, description: f.description, universe: f.universe, createGear: !!f.createGear }),
    map: mapHero6e,
  },
];

export const getGenerator = id => GENERATORS.find(g => g.id === id);

export function endpointFor(gen, form) {
  const path = typeof gen.endpoint === 'function' ? gen.endpoint(form) : gen.endpoint;
  return `${N8N_BASE}${path}`;
}

/** Map-maker costs, mirroring shared/scripts/core/plans.js ACTION_COSTS. */
export const MAP_COSTS = { mockups: { 1: 1, 2: 2, 4: 3 }, final: { draft: 2, economy: 4, standard: 6, high: 6 }, walls: 1 };

/**
 * Read a form-values object against a generator's fields: trims text,
 * coerces numbers and checkboxes, and enforces required fields.
 */
export function readForm(gen, values) {
  const out = {};
  for (const f of gen.fields) {
    let v = values[f.name];
    if (f.type === 'checkbox') v = v === true || v === 'on' || v === 'true';
    else if (f.type === 'number') v = v === '' || v === undefined || v === null ? (f.value ?? '') : Number(v);
    else v = String(v ?? f.value ?? '').trim();
    if (f.required && (v === '' || v === undefined)) throw new Error(`Please fill in ${f.label.toLowerCase()}.`);
    out[f.name] = v;
  }
  return out;
}
