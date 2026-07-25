/**
 * Army document (#70) — a standalone JournalEntry kind holding a unit roster
 * and a stationed-at link to a settlement. Stored under the same
 * flags.Pf2eNationsAndCitiesMaker.settlement path as settlements/nations so
 * it reuses the existing getSettlement/setSettlement + sheet-swap plumbing.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY } from './constants.js';

export const UNIT_TYPES = ['spearmen', 'archers', 'cavalry', 'mages', 'siege'];

function safeNum(n, def, min = -Infinity, max = Infinity) {
  const v = Number(n);
  if (!Number.isFinite(v)) return def;
  return Math.min(max, Math.max(min, v));
}

function safeString(s, def = '') {
  return (typeof s === 'string' && s.trim()) ? s.trim() : def;
}

function shortId() { return Math.random().toString(36).slice(2, 10); }

export function sanitizeArmy(raw) {
  const a = (raw && typeof raw === 'object') ? raw : {};
  const units = Array.isArray(a.units) ? a.units.map(u => ({
    id:        safeString(u?.id, `unit-${shortId()}`),
    type:      UNIT_TYPES.includes(u?.type) ? u.type : 'spearmen',
    count:     safeNum(u?.count, 1, 0, 99999),
    level:     safeNum(u?.level, 1, 0, 30),
    equipment: safeString(u?.equipment, ''),
    morale:    safeNum(u?.morale, 100, 0, 100),
  })) : [];

  return {
    kind:        'army',
    stationedAt: (typeof a.stationedAt === 'string' && a.stationedAt) ? a.stationedAt : null,
    units,
    commanderActorId: a.commanderActorId || null,
    notes:       safeString(a.notes, ''),
  };
}

export function totalUnitCount(army) {
  return (army.units || []).reduce((sum, u) => sum + Number(u.count || 0), 0);
}

export async function createArmy(name, options = {}) {
  return JournalEntry.create({
    name: name || 'New Army',
    flags: {
      [FLAG_SCOPE]: {
        [FLAG_KEY]: sanitizeArmy({ stationedAt: options.stationedAt || null }),
        createdBy: MODULE_ID,
      },
    },
  });
}
