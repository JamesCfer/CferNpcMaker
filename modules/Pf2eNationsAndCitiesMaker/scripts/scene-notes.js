/**
 * World-map scene pins (#104) — keeps any scene Note linked to a settlement
 * journal styled with the module's kind icon and a kind/population tooltip.
 */

import { MODULE_ID } from './constants.js';

const KIND_LABELS = { village: 'Village', town: 'Town', city: 'City', nation: 'Nation' };

const KIND_ICONS = {
  village: `modules/${MODULE_ID}/templates/icons/village.svg`,
  town:    `modules/${MODULE_ID}/templates/icons/town.svg`,
  city:    `modules/${MODULE_ID}/templates/icons/city.svg`,
  nation:  `modules/${MODULE_ID}/templates/icons/nation.svg`,
};

export function settlementNoteIcon(kind) {
  return KIND_ICONS[kind] || KIND_ICONS.city;
}

export function settlementNoteTooltip(settlement) {
  const label = KIND_LABELS[settlement?.kind] || 'Settlement';
  const population = Number(settlement?.population) || 0;
  return `${label} · Population ${population.toLocaleString()}`;
}

/**
 * Re-styles every Note on the settlement's linked scene that points at
 * `journal` so its icon and tooltip stay in sync with the current kind and
 * population (e.g. after population growth, a plague, or a manual edit).
 *
 * @param {JournalEntry} journal
 * @param {object}       settlement  Current (already-mutated) settlement payload.
 */
export async function syncSettlementNotes(journal, settlement) {
  if (!journal || !settlement?.sceneId) return;
  const scene = game.scenes?.get(settlement.sceneId);
  if (!scene) return;

  const icon = settlementNoteIcon(settlement.kind);
  const text = settlementNoteTooltip(settlement);
  const notes = scene.notes?.filter(n => n.entryId === journal.id) ?? [];

  for (const note of notes) {
    const updates = {};
    if (note.texture?.src !== icon) updates['texture.src'] = icon;
    if (note.text !== text) updates.text = text;
    if (Object.keys(updates).length) await note.update(updates).catch(() => {});
  }
}
