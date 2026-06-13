/**
 * Per-store income tick + tax application for a settlement.
 * Pure-ish: all mutation goes through doc.setFlag.
 */

import { FLAG_SCOPE, FLAG_KEY, MODULE_ID, getSettlement } from './constants.js';

const THIRTY_GAME_DAYS = 30;

function getJitterFactor() {
  try { return Math.min(1, Math.max(0, Number(game.settings?.get(MODULE_ID, 'incomeJitterPct') ?? 15))) / 100; }
  catch (_) { return 0.15; }
}

function jitter(base, factor) {
  const delta = base * factor;
  return Math.round((base - delta) + Math.random() * (delta * 2));
}

function gmWhisper() {
  return game.users?.filter(u => u.isGM).map(u => u.id) ?? [];
}

/**
 * Advance the per-store income by `days`. Also:
 * - drains guard wages from the treasury (#53)
 * - halves store income when unrest > 80 and posts a riot chat card (#55)
 * - marks stores `closed` after 30+ game-days in debt (#56)
 * - credits the treasury from production (#52)
 * - doubles income on a store's designated market weekday (#59)
 * - bumps settlement unrest for each active black market store (#60)
 */
export async function applyDailyTick(doc, days = 1, weekday = null) {
  if (!doc) return;
  const s = foundry.utils.deepClone(getSettlement(doc) || {});
  if (!s.stores) return;
  const now = Date.now();
  const factor = getJitterFactor();
  const unrest = Number(s.stats?.unrest ?? 0);
  const isRioting = unrest > 80;

  const newlyClosed = [];
  let blackMarketCount = 0;

  for (const store of s.stores) {
    if (store.closed) continue;
    if (store.isBlackMarket) blackMarketCount++;

    let effectiveAvg = Number(store.income?.dailyAvg ?? 0);
    if (isRioting) effectiveAvg = Math.floor(effectiveAvg / 2);

    // Market day: double income on the store's designated weekday (#59)
    if (days === 1 && store.marketWeekday !== null && weekday !== null
        && weekday === Number(store.marketWeekday)) {
      effectiveAvg *= 2;
    }

    const earned = jitter(effectiveAvg, factor) * days;
    store.income.balance = Math.max(-9_999_999, Math.round((store.income.balance || 0) + earned));
    store.income.lastTick = now;

    if (store.income.balance < 0) {
      store.income.daysInDebt = (Number(store.income.daysInDebt) || 0) + days;
      if (store.income.daysInDebt >= THIRTY_GAME_DAYS) {
        store.closed = true;
        newlyClosed.push(store.name);
      }
    } else {
      store.income.daysInDebt = 0;
    }
  }

  // Black market unrest: each active black market store draws +1 unrest per tick (#60)
  if (blackMarketCount > 0) {
    s.stats = s.stats || {};
    s.stats.unrest = Math.min(100, (Number(s.stats.unrest) || 0) + blackMarketCount);
  }

  // Guard wages (#53)
  const ranks = s.military?.ranks || [];
  const totalWages = Math.round(
    ranks.reduce((sum, r) => sum + Number(r.dailyWage ?? 0) * Number(r.count ?? 0), 0) * days
  );
  if (totalWages > 0) {
    s.treasury = s.treasury || { cp: 0, sp: 0, gp: 0, pp: 0 };
    s.treasury.gp = Math.max(-9_999_999, Math.round((s.treasury.gp || 0) - totalWages));
  }

  // Production credit (#52)
  const prodCredit = Math.round((s.production?.length ?? 0) * Number(s.population ?? 0) / 1000 * days);
  if (prodCredit > 0) {
    s.treasury = s.treasury || { cp: 0, sp: 0, gp: 0, pp: 0 };
    s.treasury.gp = (s.treasury.gp || 0) + prodCredit;
  }

  // Trade routes (#58) — credit this settlement's declared trade income, mirror to partner
  if (Array.isArray(s.tradeRoutes)) {
    for (const route of s.tradeRoutes) {
      if (!route.partnerId) continue;
      const gpPerDay = Math.round(Number(route.gpPerWeek) / 7 * days);
      if (gpPerDay <= 0) continue;
      s.treasury = s.treasury || { cp: 0, sp: 0, gp: 0, pp: 0 };
      s.treasury.gp = (s.treasury.gp || 0) + gpPerDay;
      const partnerDoc = game.journal?.get(route.partnerId);
      if (partnerDoc) {
        const ps = foundry.utils.deepClone(getSettlement(partnerDoc) || {});
        ps.treasury = ps.treasury || { cp: 0, sp: 0, gp: 0, pp: 0 };
        ps.treasury.gp = (ps.treasury.gp || 0) + gpPerDay;
        partnerDoc.setFlag(FLAG_SCOPE, FLAG_KEY, ps).catch(() => {});
      }
    }
  }

  // Crime leakage (#67) — each crime level point steals 0.5% of positive store balances per day
  const crimeLevel = Math.min(10, Math.max(0, Number(s.crimeLevel ?? 0)));
  if (crimeLevel > 0) {
    const leakFraction = crimeLevel * 0.005 * days;
    let totalLeaked = 0;
    for (const store of s.stores) {
      if (store.closed) continue;
      const bal = Number(store.income?.balance ?? 0);
      if (bal <= 0) continue;
      const leak = Math.floor(bal * leakFraction);
      if (leak > 0) {
        store.income.balance = bal - leak;
        totalLeaked += leak;
      }
    }
    if (Math.random() < (crimeLevel / 50) * days) {
      const incidents = [
        'A shipment of goods was stolen from a local merchant.',
        'Pickpockets have been reported near the market district.',
        'A protection racket is extorting several shop owners.',
        'Contraband was discovered in a warehouse.',
        'A prominent merchant was mugged on the way home.',
      ];
      const msg = incidents[Math.floor(Math.random() * incidents.length)];
      ChatMessage.create({
        content: `<h3><i class="fa-solid fa-user-ninja"></i> Crime Incident in ${doc.name}</h3>
          <p>${msg}</p>
          <p><em>Crime level: ${crimeLevel}/10${totalLeaked > 0 ? ` — ${totalLeaked} gp leaked from store balances.` : ''}</em></p>`,
        whisper: gmWhisper(),
      }).catch(() => {});
    }
  }

  // Famine (#65) — drains treasury, raises unrest, halts growth while active
  let famineActive = false;
  if (Number(s.famineDaysLeft ?? 0) > 0) {
    famineActive = true;
    const drainPerDay = Math.max(1, Math.round((Number(s.population) || 1000) * 0.01));
    s.treasury = s.treasury || { cp: 0, sp: 0, gp: 0, pp: 0 };
    s.treasury.gp = Math.max(-9_999_999, Math.round((s.treasury.gp || 0) - drainPerDay * days));
    s.stats = s.stats || {};
    s.stats.unrest = Math.min(100, (Number(s.stats.unrest) || 0) + days);
    s.famineDaysLeft = Math.max(0, Number(s.famineDaysLeft) - days);
    if (s.famineDaysLeft === 0) {
      ChatMessage.create({
        content: `<h3><i class="fa-solid fa-seedling"></i> Famine Ends in ${doc.name}</h3>
          <p>The famine has passed. Population growth and normal treasury income resume.</p>`,
        whisper: gmWhisper(),
      }).catch(() => {});
    }
  }

  // Population growth (#63) — halted during famine
  const growthRate = Number(s.growthRate ?? 0.001);
  const growthGain = famineActive ? 0 : Math.round((Number(s.population) || 0) * growthRate * (1 - Math.min(100, unrest) / 100) * days);
  if (growthGain > 0) s.population = Math.max(1, (Number(s.population) || 0) + growthGain);

  // Treasury history snapshot (#61) — append after all mutations, keep last 30
  s.treasuryHistory = Array.isArray(s.treasuryHistory) ? s.treasuryHistory : [];
  s.treasuryHistory.push({ gp: Math.round(s.treasury?.gp ?? 0) });
  if (s.treasuryHistory.length > 30) s.treasuryHistory = s.treasuryHistory.slice(-30);

  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);

  if (isRioting) {
    ChatMessage.create({
      content: `<h3><i class="fa-solid fa-fire"></i> Riots in ${doc.name}!</h3>
        <p>Unrest is <strong>${unrest}/100</strong> — store income is halved until unrest drops below 80.</p>`,
      whisper: gmWhisper(),
    }).catch(() => {});
  }

  if (newlyClosed.length) {
    ChatMessage.create({
      content: `<h3><i class="fa-solid fa-store-slash"></i> Stores Closed in ${doc.name}</h3>
        <p>The following stores have gone bankrupt after ${THIRTY_GAME_DAYS}+ days in debt:
        <strong>${newlyClosed.join(', ')}</strong>.</p>`,
      whisper: gmWhisper(),
    }).catch(() => {});
  }
}

/**
 * Apply a tax event to a settlement. Drains a percentage of every store's
 * balance into the town treasury and raises unrest by 1 (#54).
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
    if (store.closed) continue;
    const bal = Number(store.income?.balance ?? 0);
    if (bal <= 0) continue;
    const take = Math.floor(bal * factor);
    if (take <= 0) continue;
    store.income.balance = bal - take;
    collected += take;
  }

  s.treasury = s.treasury || { cp: 0, sp: 0, gp: 0, pp: 0 };
  s.treasury.gp = (s.treasury.gp || 0) + collected;

  // Taxation raises unrest (#54)
  s.stats = s.stats || {};
  s.stats.unrest = Math.min(100, (Number(s.stats.unrest) || 0) + 1);

  const log = `[${new Date().toLocaleString()}] ${payload.taxType || 'tax'} @${ratePct}% → collected ${collected} gp`;
  s.notes = `${log}<br>${s.notes || ''}`;

  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);
  return { collected };
}

/**
 * Apply a festival event — reduces unrest by 1 (#54).
 *
 * @param {JournalEntry} doc
 */
export async function applyFestival(doc) {
  if (!doc) return;
  const s = foundry.utils.deepClone(getSettlement(doc) || {});
  s.stats = s.stats || {};
  s.stats.unrest = Math.max(0, (Number(s.stats.unrest) || 0) - 1);
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);
}

/**
 * Apply a plague event — reduces population by `payload.ratePct` percent (#64).
 *
 * @param {JournalEntry} doc
 * @param {{ ratePct:number }} payload
 */
export async function applyPlague(doc, payload = {}) {
  if (!doc) return;
  const ratePct = Math.max(0, Math.min(100, Number(payload.ratePct) || 10));
  const s = foundry.utils.deepClone(getSettlement(doc) || {});
  const before = Number(s.population) || 0;
  const lost = Math.round(before * ratePct / 100);
  s.population = Math.max(0, before - lost);
  s.stats = s.stats || {};
  s.stats.unrest = Math.min(100, (Number(s.stats.unrest) || 0) + Math.round(ratePct / 5));
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);
  ChatMessage.create({
    content: `<h3><i class="fa-solid fa-biohazard"></i> Plague Strikes ${doc.name}!</h3>
      <p>A plague has swept through the settlement, claiming <strong>${lost.toLocaleString()}</strong> lives
      (${ratePct}% of the population).</p>
      <p>Population: ${s.population.toLocaleString()} (was ${before.toLocaleString()})</p>`,
    whisper: gmWhisper(),
  }).catch(() => {});
}

/**
 * Apply a famine event — drains treasury, raises unrest, halts growth for `payload.duration` days (#65).
 *
 * @param {JournalEntry} doc
 * @param {{ duration:number, unrestHit:number }} payload
 */
export async function applyFamine(doc, payload = {}) {
  if (!doc) return;
  const duration  = Math.max(1, Math.min(9999, Number(payload.duration)  || 30));
  const unrestHit = Math.max(0, Math.min(100,  Number(payload.unrestHit) || 15));
  const s = foundry.utils.deepClone(getSettlement(doc) || {});
  s.famineDaysLeft = duration;
  s.stats = s.stats || {};
  s.stats.unrest = Math.min(100, (Number(s.stats.unrest) || 0) + unrestHit);
  const immediateGpDrain = Math.round((Number(s.population) || 1000) * 0.05);
  s.treasury = s.treasury || { cp: 0, sp: 0, gp: 0, pp: 0 };
  s.treasury.gp = Math.max(-9_999_999, (s.treasury.gp || 0) - immediateGpDrain);
  await doc.setFlag(FLAG_SCOPE, FLAG_KEY, s);
  ChatMessage.create({
    content: `<h3><i class="fa-solid fa-wheat-awn-circle-exclamation"></i> Famine Begins in ${doc.name}!</h3>
      <p>Crops have failed and stores run low. The famine will last <strong>${duration} days</strong>,
      draining the treasury and raising unrest each day until resolved.</p>
      <p>Unrest raised by ${unrestHit}. Treasury docked ${immediateGpDrain} gp immediately.</p>`,
    whisper: gmWhisper(),
  }).catch(() => {});
}
