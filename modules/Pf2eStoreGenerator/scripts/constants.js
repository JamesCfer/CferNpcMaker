/**
 * Module-wide identifiers and the store flag accessors.
 */

export const MODULE_ID  = 'Pf2eStoreGenerator';
export const FLAG_SCOPE = MODULE_ID;
export const FLAG_KEY   = 'store';

/**
 * @param {JournalEntry} doc
 * @returns {object|null} The stored store payload, or null when absent.
 */
export function getStore(doc) {
  return doc?.getFlag?.(FLAG_SCOPE, FLAG_KEY) || null;
}

/**
 * @param {JournalEntry} doc
 * @param {object}       data
 */
export async function setStore(doc, data) {
  return doc?.setFlag?.(FLAG_SCOPE, FLAG_KEY, data);
}
