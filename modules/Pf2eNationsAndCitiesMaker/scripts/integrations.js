/**
 * Bridges to sibling modules (Pf2eNpcMaker, Pf2eItemGenerator).
 * Uses fire-and-forget Foundry Hooks so this module never needs to import
 * sibling code at build time and silently no-ops when they aren't installed.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';
import { N8N_BASE, devUrl, isDevMode } from './core/n8n.js';
import { postToN8n } from './core/adapter.js';
import { Storage } from './core/storage.js';
import { sanitizeSettlement } from './sanitizer.js';

function npcModuleActive() {
  return !!game.modules?.get('Pf2eNpcMaker')?.active;
}
function itemModuleActive() {
  return !!game.modules?.get('Pf2eItemGenerator')?.active;
}

export function canGenerateNpc()  { return npcModuleActive(); }
export function canGenerateItem() { return itemModuleActive(); }

/**
 * Prefill-launch the NPC builder for a shop staff entry.
 * onCreate: optional callback invoked with the resulting Actor.
 */
export function generateStaffNpc({ settlementDoc, storeId, staff, level = 1, onCreate }) {
  if (!npcModuleActive()) {
    ui.notifications?.warn('Install & enable Pf2eNpcMaker to generate NPCs for this shop.');
    return;
  }
  const settlement = getSettlement(settlementDoc);
  const settlementName = settlementDoc?.name || 'this town';
  const description = [
    `${staff?.role ? `${staff.role} ` : ''}at ${storeNameById(settlement, storeId)} in ${settlementName}.`,
    staff?.shift ? `Works the ${staff.shift} shift.` : '',
    settlement?.biome ? `Setting: ${settlement.biome}.` : '',
  ].filter(Boolean).join(' ');

  Hooks.callAll('Pf2eNpcMaker.openWithPrefill', {
    name: staff?.name || 'Shopkeeper',
    level,
    description,
    onCreate: async (actor) => {
      try {
        await attachActorToStaff(settlementDoc, storeId, staff?.id, actor.id);
        await tagHomeSettlement(actor, settlementDoc);
        if (typeof onCreate === 'function') onCreate(actor);
      } catch (err) {
        console.error(`[${MODULE_ID}] failed to attach actor to staff`, err);
      }
    },
  });
}

/**
 * Tag an NPC created through a settlement bridge with the settlement it
 * belongs to, so its actor sheet can link back. (#109, #110)
 */
export async function tagHomeSettlement(actor, settlementDoc) {
  if (!actor || !settlementDoc) return;
  try {
    await actor.setFlag(MODULE_ID, 'homeSettlementId', settlementDoc.id);
  } catch (err) {
    console.error(`[${MODULE_ID}] failed to tag home settlement`, err);
  }
}

/**
 * Prefill-launch the Item builder for a store inventory entry.
 * Drops the resulting Item id into the store's inventory list.
 */
export function generateStoreItem({ settlementDoc, storeId, itemType = 'equipment', hint = '', onCreate }) {
  if (!itemModuleActive()) {
    ui.notifications?.warn('Install & enable Pf2eItemGenerator to generate items for this shop.');
    return;
  }
  const settlement = getSettlement(settlementDoc);
  const store = settlement?.stores?.find(s => s.id === storeId);
  const description = [
    `Stocked at ${store?.name || 'a shop'} in ${settlementDoc?.name || 'town'}.`,
    hint,
  ].filter(Boolean).join(' ');

  const tierBias = { low: -2, standard: 0, high: 3, luxury: 6 };
  const bias = tierBias[store?.priceTier] ?? 0;

  Hooks.callAll('Pf2eItemGenerator.openWithPrefill', {
    name: '',
    level: Math.max(0, Math.min(20, Math.floor((settlement?.population || 500) / 1000) + bias)),
    itemType,
    description,
    onCreate: async (item) => {
      try {
        await appendItemToStore(settlementDoc, storeId, {
          name: item.name,
          itemId: item.id,
          price: item.system?.price?.value?.gp ?? 0,
          stock: 1,
        });
        if (typeof onCreate === 'function') onCreate(item);
      } catch (err) {
        console.error(`[${MODULE_ID}] failed to append item to store`, err);
      }
    },
  });
}

/* ── flag-mutation helpers (internal) ────────────────────── */

function storeNameById(settlement, storeId) {
  return settlement?.stores?.find(s => s.id === storeId)?.name || 'the shop';
}

async function attachActorToStaff(doc, storeId, staffId, actorId) {
  const s = foundry.utils.deepClone(getSettlement(doc) || {});
  const store = s.stores?.find(x => x.id === storeId);
  if (!store) return;
  const person = store.staff?.find(p => p.id === staffId);
  if (person) {
    person.actorId = actorId;
  } else if (store.owner && (!staffId || staffId === 'owner')) {
    store.owner.actorId = actorId;
  }
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);
}

async function appendItemToStore(doc, storeId, entry) {
  const s = foundry.utils.deepClone(getSettlement(doc) || {});
  const store = s.stores?.find(x => x.id === storeId);
  if (!store) return;
  store.inventory = store.inventory || [];
  store.inventory.push({
    id: `inv-${Math.random().toString(36).slice(2, 10)}`,
    ...entry,
  });
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);
}

/* ── store reroll (#35, #36) ─────────────────────────────── */

const CITY_BUILDER_ENDPOINT = `${N8N_BASE}/webhook/city-builder`;

function getKey() {
  return new Storage(MODULE_ID).getKey() || '';
}

function stubStoreForType(type) {
  const names = {
    blacksmith: 'The Iron Anvil',    armorer: 'Shield & Scale Armory',
    weapons:    'The Sharpened Edge', alchemist: 'The Alembic',
    tavern:     'The Crooked Flagon', inn: "The Wayfarer's Rest",
    general:    "Dower's Goods",      grocer: 'The Market Stall',
    tailor:     'Needle & Thread',    jeweler: 'The Gilded Gem',
    temple:     'The Sacred Flame',   magic: 'The Arcane Corner',
    stable:     'Ironhoof Stables',   apothecary: 'The Healing Touch',
    bookbinder: 'Scrolls & Tomes',   other: 'Miscellany',
  };
  return {
    name: names[type] || 'New Shop', type,
    owner: { name: 'Unknown Owner', actorId: null },
    staff: [],
    hours: { open: '08:00', close: '20:00', daysClosed: [] },
    inventory: [],
    income: { balance: 50, dailyAvg: 6, lastTick: 0 },
  };
}

const KIND_STUB_TYPES = {
  city:    ['blacksmith', 'tavern', 'general', 'alchemist', 'temple'],
  town:    ['general', 'tavern', 'blacksmith'],
  village: ['general', 'tavern'],
  nation:  ['general', 'tavern', 'magic'],
};

export async function rerollStores(doc) {
  if (!doc) return;
  const settlement = getSettlement(doc);
  if (!settlement) return;

  const endpoint = devUrl(CITY_BUILDER_ENDPOINT, isDevMode(MODULE_ID));
  const payload = {
    kind: settlement.kind, size: settlement.size,
    biome: settlement.biome || '', population: settlement.population || 0,
    description: settlement.ai?.prompt || '',
    requestType: 'storesOnly',
    includeStores: true, includeMilitary: false, includeLeadership: false,
  };

  let rawStores;
  try {
    const { response, responseText } = await postToN8n(endpoint, payload, getKey());
    let data;
    try { data = JSON.parse(responseText); } catch (e) { throw new Error(`Invalid JSON: ${e.message}`); }
    if (!response.ok) throw new Error(data?.message || `Status ${response.status}`);
    rawStores = data.stores || data.settlement?.stores;
    if (!Array.isArray(rawStores) || !rawStores.length) throw new Error('No stores in response');
  } catch (netErr) {
    if (!(game.settings?.get?.(MODULE_ID, 'allowOfflineStub') ?? true)) throw netErr;
    console.warn(`[${MODULE_ID}] rerollStores offline stub:`, netErr?.message);
    rawStores = (KIND_STUB_TYPES[settlement.kind] || KIND_STUB_TYPES.town).map(stubStoreForType);
  }

  const s = foundry.utils.deepClone(settlement);
  s.stores = sanitizeSettlement({ ...settlement, stores: rawStores }).stores;
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);
}

export async function rerollSingleStore(doc, storeId) {
  if (!doc) return;
  const settlement = getSettlement(doc);
  if (!settlement) return;

  const store = settlement.stores?.find(x => x.id === storeId);
  if (!store) return;

  const endpoint = devUrl(CITY_BUILDER_ENDPOINT, isDevMode(MODULE_ID));
  const payload = {
    kind: settlement.kind, size: settlement.size,
    biome: settlement.biome || '', population: settlement.population || 0,
    description: settlement.ai?.prompt || '',
    requestType: 'singleStore', storeType: store.type,
    includeStores: true, includeMilitary: false, includeLeadership: false,
  };

  let rawStore;
  try {
    const { response, responseText } = await postToN8n(endpoint, payload, getKey());
    let data;
    try { data = JSON.parse(responseText); } catch (e) { throw new Error(`Invalid JSON: ${e.message}`); }
    if (!response.ok) throw new Error(data?.message || `Status ${response.status}`);
    rawStore = data.store || data.stores?.[0];
    if (!rawStore) throw new Error('No store in response');
  } catch (netErr) {
    if (!(game.settings?.get?.(MODULE_ID, 'allowOfflineStub') ?? true)) throw netErr;
    console.warn(`[${MODULE_ID}] rerollSingleStore offline stub:`, netErr?.message);
    rawStore = stubStoreForType(store.type);
  }

  const sanitized = sanitizeSettlement({ ...settlement, stores: [rawStore] }).stores[0];
  if (!sanitized) return;

  const s = foundry.utils.deepClone(settlement);
  const idx = s.stores.findIndex(x => x.id === storeId);
  if (idx >= 0) s.stores[idx] = { ...sanitized, id: storeId };
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);
}
