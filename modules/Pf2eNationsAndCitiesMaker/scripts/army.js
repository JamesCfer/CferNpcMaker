/**
 * Army document (#70) — a standalone JournalEntry kind holding a unit roster
 * and a stationed-at link to a settlement. Stored under the same
 * flags.Pf2eNationsAndCitiesMaker.settlement path as settlements/nations so
 * it reuses the existing getSettlement/setSettlement + sheet-swap plumbing.
 */

import { MODULE_ID, FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';
import { sanitizeSettlement } from './sanitizer.js';

export const UNIT_TYPES = ['spearmen', 'archers', 'cavalry', 'mages', 'siege'];

/** Recruitment cost per unit (#72) — gp drained from the treasury, pop drained from the populace. */
export const UNIT_COSTS = {
  spearmen: { gp: 5,   pop: 1 },
  archers:  { gp: 8,   pop: 1 },
  cavalry:  { gp: 25,  pop: 1 },
  mages:    { gp: 60,  pop: 1 },
  siege:    { gp: 150, pop: 2 },
};

/** Daily upkeep per unit, gp (#73) — drained from the stationed-at settlement's treasury. */
export const UNIT_WAGES = {
  spearmen: 1,
  archers:  1.5,
  cavalry:  3,
  mages:    5,
  siege:    8,
};

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

  const arrivalDate = a.arrivalDate;
  const validArrivalDate = arrivalDate && typeof arrivalDate === 'object'
    && Number.isFinite(Number(arrivalDate.year))
    && Number.isFinite(Number(arrivalDate.month))
    && Number.isFinite(Number(arrivalDate.day));

  return {
    kind:        'army',
    stationedAt: (typeof a.stationedAt === 'string' && a.stationedAt) ? a.stationedAt : null,
    mode:        ['garrison', 'field'].includes(a.mode) ? a.mode : 'garrison',
    destination: (typeof a.destination === 'string' && a.destination) ? a.destination : null,
    arrivalDate: validArrivalDate
      ? { year: Number(arrivalDate.year), month: Number(arrivalDate.month), day: Number(arrivalDate.day) }
      : null,
    units,
    commanderActorId: a.commanderActorId || null,
    ownerNationId: (typeof a.ownerNationId === 'string' && a.ownerNationId) ? a.ownerNationId : null,
    supplySource: (typeof a.supplySource === 'string' && a.supplySource) ? a.supplySource : null,
    notes:       safeString(a.notes, ''),
  };
}

/** Compare two { year, month, day } dates: returns -1/0/1. Mirrors Pf2eCalendarTimeline's scheduler.cmpDate. */
export function cmpDate(a, b) {
  if (a.year  !== b.year)  return a.year  < b.year  ? -1 : 1;
  if (a.month !== b.month) return a.month < b.month ? -1 : 1;
  if (a.day   !== b.day)   return a.day   < b.day   ? -1 : 1;
  return 0;
}

/**
 * Compute the arrival date for an army sent on a `travelDays`-day march (#75).
 * `calendarDef` is the Pf2eCalendarTimeline calendar shape ({ daysPerMonth[] });
 * falls back to a flat 30-day month if the calendar module isn't available.
 */
export function computeArrivalDate(fromDate, travelDays, calendarDef) {
  const daysPerMonth = calendarDef?.daysPerMonth?.length ? calendarDef.daysPerMonth : Array(12).fill(30);
  const monthsPerYear = daysPerMonth.length;
  let { year, month, day } = fromDate;
  day += Math.max(1, Math.floor(Number(travelDays) || 1));
  while (true) {
    const dim = daysPerMonth[month - 1];
    if (day <= dim) break;
    day -= dim;
    month += 1;
    if (month > monthsPerYear) { month = 1; year += 1; }
  }
  return { year, month, day };
}

export function totalUnitCount(army) {
  return (army.units || []).reduce((sum, u) => sum + Number(u.count || 0), 0);
}

/** Total gp + population needed to recruit `count` units of `type` (#72). */
export function recruitmentCost(type, count) {
  const c = UNIT_COSTS[type] || UNIT_COSTS.spearmen;
  const n = Math.max(0, Number(count) || 0);
  return { gp: c.gp * n, pop: c.pop * n };
}

/** Total daily gp upkeep for an army's full roster (#73). */
export function totalDailyWage(army) {
  return (army.units || []).reduce(
    (sum, u) => sum + (UNIT_WAGES[u.type] ?? UNIT_WAGES.spearmen) * Number(u.count || 0),
    0,
  );
}

/**
 * Drain an army's daily unit wages from its stationed-at settlement's treasury (#73).
 * Returns the gp amount drained (0 if the army isn't stationed or has no units).
 */
export async function applyArmyWages(armyDoc, settlementDoc, days = 1) {
  if (!settlementDoc) return 0;
  const army = sanitizeArmy(getSettlement(armyDoc) || {});
  const wage = Math.round(totalDailyWage(army) * days);
  if (wage <= 0) return 0;
  const s = foundry.utils.deepClone(getSettlement(settlementDoc) || {});
  s.treasury = s.treasury || { cp: 0, sp: 0, gp: 0, pp: 0 };
  s.treasury.gp = Math.max(-9_999_999, Math.round((s.treasury.gp || 0) - wage));
  await settlementDoc.setFlag(FLAG_SCOPE, FLAG_KEY, s);
  return wage;
}

/** Food consumed per unit per day (#79) — abstracted; settlements have no literal food stockpile. */
const FOOD_PER_UNIT = 0.05;
const SUPPLY_CUT_MORALE_LOSS = 10;
const SUPPLY_RESTORED_MORALE_GAIN = 5;
const STARVATION_ATTRITION_PCT = 0.05;

/** Total daily food an army's full roster needs (#79). */
export function totalDailyFoodNeed(army) {
  return totalUnitCount(army) * FOOD_PER_UNIT;
}

/** Food a settlement's population can spare per day, standing in for agricultural capacity (#79). */
export function settlementFoodCapacity(settlement) {
  return Math.floor((Number(settlement?.population) || 0) / 200);
}

/** Whether `sourceSettlement` can currently feed `army` (#79) — false if occupied or undersupplied. */
export function isSupplied(army, sourceSettlement) {
  if (!sourceSettlement || sourceSettlement.stats?.occupied) return false;
  return settlementFoodCapacity(sourceSettlement) >= totalDailyFoodNeed(army);
}

/**
 * Apply a day's supply chain (#79): an army with a `supplySource` fed by that
 * settlement recovers morale; cut off (no source, source occupied, or source
 * can't spare enough food), it loses morale each day, and once morale bottoms
 * out at 0 it starts losing troops to starvation attrition.
 * No-op for armies with no units or no assigned supply source.
 */
export async function applyArmySupply(armyDoc, sourceDoc, days = 1) {
  const army = sanitizeArmy(getSettlement(armyDoc) || {});
  if (!army.supplySource || !totalUnitCount(army)) return null;

  const sourceSettlement = sourceDoc ? sanitizeSettlement(getSettlement(sourceDoc) || {}) : null;
  const supplied = isSupplied(army, sourceSettlement);

  let starved = false;
  const units = army.units.map(u => {
    const count = Math.max(0, Number(u.count) || 0);
    let morale = Math.max(0, Math.min(100, Number(u.morale) || 0));
    if (supplied) {
      return { ...u, morale: Math.min(100, morale + SUPPLY_RESTORED_MORALE_GAIN * days) };
    }
    morale = Math.max(0, morale - SUPPLY_CUT_MORALE_LOSS * days);
    let nextCount = count;
    if (morale <= 0 && count > 0) {
      nextCount = Math.max(0, count - Math.ceil(count * STARVATION_ATTRITION_PCT * days));
      if (nextCount < count) starved = true;
    }
    return { ...u, morale, count: nextCount };
  });

  await armyDoc.setFlag(FLAG_SCOPE, FLAG_KEY, sanitizeArmy({ ...army, units }));
  return { supplied, starved };
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
