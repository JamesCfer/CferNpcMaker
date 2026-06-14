/**
 * Bridges to sibling modules (Pf2eNpcMaker, Pf2eItemGenerator).
 * Uses fire-and-forget Foundry Hooks so this module never needs to import
 * sibling code at build time and silently no-ops when they aren't installed.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';

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
        if (typeof onCreate === 'function') onCreate(actor);
      } catch (err) {
        console.error(`[${MODULE_ID}] failed to attach actor to staff`, err);
      }
    },
  });
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
