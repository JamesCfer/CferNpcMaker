import { migrateSettlement } from './migrations.js';

export const MODULE_ID = 'Pf2eNationsAndCitiesMaker';
export const FLAG_SCOPE = MODULE_ID;
export const FLAG_KEY   = 'settlement';

export function getSettlement(doc) {
  const raw = doc?.getFlag?.(FLAG_SCOPE, FLAG_KEY) || null;
  return raw ? migrateSettlement(raw) : null;
}

export async function setSettlement(doc, data) {
  return doc?.setFlag?.(FLAG_SCOPE, FLAG_KEY, data);
}

/** Standard store types (used to build tabs on the city sheet). */
export const STORE_TYPES = [
  'blacksmith', 'armorer', 'weapons', 'alchemist',
  'tavern', 'inn', 'general', 'grocer',
  'tailor', 'jeweler', 'temple', 'magic',
  'stable', 'apothecary', 'bookbinder', 'other',
];

export function storeTypeLabel(t) {
  if (!t) return 'Other';
  return t.charAt(0).toUpperCase() + t.slice(1);
}
