/**
 * Per-store income tick + tax application for a settlement.
 * Pure-ish: all mutation goes through doc.setFlag.
 */

import { FLAG_SCOPE, FLAG_KEY, getSettlement } from './constants.js';

function jitter(base, pct = 0.15) {
  const delta = base * pct;
  return Math.round((base - delta) + Math.random() * (delta * 2));
}

/**
 * Advance the per-store income by `days`. Adds a small RNG jitter so balances
 * feel alive across many ticks.
 */
export async function applyDailyTick(doc, days = 1) {
  if (!doc) return;
  const s = foundry.utils.deepClone(getSettlement(doc) || {});
  if (!s.stores) return;
  const now = Date.now();
  for (const store of s.stores) {
    const dailyAvg = Number(store.income?.dailyAvg ?? 0);
    const earned = jitter(dailyAvg) * days;
    store.income.balance = Math.max(-9_999_999, Math.round((store.income.balance || 0) + earned));
    store.income.lastTick = now;
  }
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);
}

/**
 * Apply a tax event to a settlement. Drains a percentage of every store's
 * balance (capped at the balance itself) into the town treasury (gp).
 *
 * @param {JournalEntry} doc
 * @param {{ taxType:string, ratePct:number }} payload
 * @returns {Promise<{ collected:number }>}
 */
export async function applyTax(doc, payload = {}) {
  if (!doc) return { collected: 0 };
  const ratePct = Math.max(0, Math.min(100, Number(payload.ratePct) || 0));
  if (!ratePct) return { collected: 0 };

  const s = foundry.utils.deepClone(getSettlement(doc) || {});
  if (!s.stores) return { collected: 0 };

  let collected = 0;
  const factor = ratePct / 100;

  for (const store of s.stores) {
    const bal = Number(store.income?.balance ?? 0);
    if (bal <= 0) continue;
    const take = Math.floor(bal * factor);
    if (take <= 0) continue;
    store.income.balance = bal - take;
    collected += take;
  }

  s.treasury = s.treasury || { cp: 0, sp: 0, gp: 0, pp: 0 };
  s.treasury.gp = (s.treasury.gp || 0) + collected;

  const log = `[${new Date().toLocaleString()}] ${payload.taxType || 'tax'} @${ratePct}% → collected ${collected} gp`;
  s.notes = `${log}<br>${s.notes || ''}`;

  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);
  return { collected };
}
