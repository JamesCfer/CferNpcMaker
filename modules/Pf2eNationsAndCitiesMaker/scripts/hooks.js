/**
 * Adventurer-hook generator (#68). Pure local table, no AI call — rolls 3
 * plot hooks tied to the settlement's current state (unrest, crime, black
 * markets, bankrupt stores, religions, districts, trade routes).
 */

import { rngFromSeed, pick } from './templates.js';

const GENERIC_HOOKS = [
  'A caravan bound for {name} has missed its expected arrival by three days; the last waystation reported nothing unusual.',
  'A local claims to have found an old boundary stone near {name} marked with a sigil no one recognises.',
  'A trusted courier serving {name} has started delivering sealed letters they refuse to explain.',
  'Livestock near {name} have been found drained of blood, with no tracks leading away from the pens.',
  'A stranger has been asking after a specific family in {name}, offering coin for any word of their whereabouts.',
];

const UNREST_HOOKS = [
  'Anonymous pamphlets accusing the leadership of {name} of corruption have started circulating in the market square.',
  'A crowd gathered outside the town hall of {name} last night; nothing was broken, but the mood was ugly.',
  'Someone is paying drinkers at every tavern in {name} to complain loudly about the local rulers.',
];

const CRIME_HOOKS = [
  'A protection racket has started collecting "insurance" from the smaller stalls in {name}, and the guard looks the other way.',
  'A string of burglaries in {name} has hit only houses whose owners were away at the same public event.',
  'A fence in {name} is offering an unusually good price for anything with a temple or noble crest on it.',
];

const BLACK_MARKET_HOOKS = [
  '{store} in {name} is rumoured to move goods that never touch an official ledger — for the right price and the right face.',
  'Someone has been underselling {store}\'s usual black-market rates in {name}, and the owner wants to know who.',
];

const BANKRUPTCY_HOOKS = [
  '{store} in {name} closed its doors owing coin to half the town, and its former owner has not been seen since.',
  'The shuttered {store} in {name} still has stock inside, and its creditors are arguing over who has the right to claim it.',
];

const RELIGION_HOOKS = [
  'A minor relic of the {religion} faith has gone missing from its shrine in {name}, and the temple is offering a quiet reward.',
  'Followers of {religion} in {name} have been holding closed-door meetings at odd hours, unusual even for the devout.',
];

const DISTRICT_HOOKS = [
  'Something is scaring residents out of the {district} district of {name} after dark, though no one can describe what.',
  'A dispute over who controls the {district} district of {name} has turned from words to hired muscle.',
];

const TRADE_ROUTE_HOOKS = [
  'The trade route serving {name} has gone quiet for a week longer than any delay should account for.',
  'Goods arriving in {name} along its trade route have started showing up short of what the manifests promise.',
];

function fill(template, ctx) {
  return template.replace(/\{(\w+)\}/g, (_, key) => ctx[key] ?? '');
}

function shuffle(arr, rng) {
  const out = arr.slice();
  for (let i = out.length - 1; i > 0; i--) {
    const j = Math.floor(rng() * (i + 1));
    [out[i], out[j]] = [out[j], out[i]];
  }
  return out;
}

export function generateHooks(settlement, name) {
  const rng = rngFromSeed(`${name}|${Date.now()}|${Math.random()}`);
  const ctx = { name };

  // Each entry is a pool of hooks tied to one piece of settlement state.
  // At most one hook per active pool makes it into the final 3, so a hook
  // tied to real state (unrest, a district, ...) is never crowded out by
  // the always-present generic filler pool.
  const statePools = [];

  if ((settlement.stats?.unrest ?? 0) >= 40) statePools.push(UNREST_HOOKS.map(t => fill(t, ctx)));
  if ((settlement.crimeLevel ?? 0) >= 5) statePools.push(CRIME_HOOKS.map(t => fill(t, ctx)));

  const blackMarketStores = (settlement.stores || []).filter(s => s.isBlackMarket);
  if (blackMarketStores.length) {
    statePools.push(BLACK_MARKET_HOOKS.map(t => fill(t, { ...ctx, store: pick(blackMarketStores, rng).name })));
  }

  const closedStores = (settlement.stores || []).filter(s => s.closed);
  if (closedStores.length) {
    statePools.push(BANKRUPTCY_HOOKS.map(t => fill(t, { ...ctx, store: pick(closedStores, rng).name })));
  }

  if ((settlement.religions || []).length) {
    statePools.push(RELIGION_HOOKS.map(t => fill(t, { ...ctx, religion: pick(settlement.religions, rng).name })));
  }

  if ((settlement.districts || []).length) {
    statePools.push(DISTRICT_HOOKS.map(t => fill(t, { ...ctx, district: pick(settlement.districts, rng).name })));
  }

  if ((settlement.tradeRoutes || []).length) {
    statePools.push(TRADE_ROUTE_HOOKS.map(t => fill(t, ctx)));
  }

  const hooks = [];
  for (const pool of shuffle(statePools, rng)) {
    if (hooks.length === 3) break;
    hooks.push(pick(pool, rng));
  }

  if (hooks.length < 3) {
    for (const hook of shuffle(GENERIC_HOOKS.map(t => fill(t, ctx)), rng)) {
      if (hooks.length === 3) break;
      if (!hooks.includes(hook)) hooks.push(hook);
    }
  }

  return hooks;
}
