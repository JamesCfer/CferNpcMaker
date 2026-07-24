import { describe, it, expect } from 'vitest';
import { sanitizeSettlement }  from '../../modules/Pf2eNationsAndCitiesMaker/scripts/sanitizer.js';
import { computeNext, DEFAULT_CALENDAR, GENERIC_FANTASY_CALENDAR,
         seasonForMonth, rollWeather } from '../../modules/Pf2eCalendarTimeline/scripts/scheduler.js';
import { applyDailyTick, applyFestival, applyTax } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/economy.js';
import { CURRENT_SCHEMA_VERSION, migrateSettlement } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/migrations.js';
import { generateHooks } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/hooks.js';
import { settlementNoteIcon, settlementNoteTooltip } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/scene-notes.js';

describe('sanitizeSettlement defaults', () => {
  it('fills in kind and size', () => {
    const s = sanitizeSettlement({});
    expect(s.kind).toBe('town');
    expect(s.size).toBe('town');
  });

  it('preserves valid kind', () => {
    expect(sanitizeSettlement({ kind: 'city' }).kind).toBe('city');
  });

  it('rejects invalid kind and falls back to town', () => {
    expect(sanitizeSettlement({ kind: 'hamlet' }).kind).toBe('town');
  });

  it('coerces stats to safe defaults', () => {
    const s = sanitizeSettlement({});
    expect(s.stats.morale).toBe(60);
    expect(s.stats.unrest).toBe(5);
    expect(s.stats.hp).toBeGreaterThanOrEqual(1);
  });

  it('defaults treasury to zero coins', () => {
    const { treasury } = sanitizeSettlement({});
    expect(treasury).toEqual({ cp: 0, sp: 0, gp: 0, pp: 0 });
  });

  it('defaults stores to empty array', () => {
    expect(sanitizeSettlement({}).stores).toEqual([]);
  });

  it('sanitizes store fields with safe defaults', () => {
    const s = sanitizeSettlement({ stores: [{}] });
    expect(s.stores[0].name).toBe('Unnamed Shop');
    expect(s.stores[0].income.dailyAvg).toBe(5);
  });

  it('stamps _schemaVersion', () => {
    expect(sanitizeSettlement({})._schemaVersion).toBe(CURRENT_SCHEMA_VERSION);
  });

  it('defaults bannerImage to null', () => {
    expect(sanitizeSettlement({}).bannerImage).toBeNull();
  });

  it('preserves a valid bannerImage path', () => {
    expect(sanitizeSettlement({ bannerImage: 'worlds/foo/banner.webp' }).bannerImage).toBe('worlds/foo/banner.webp');
  });
});

describe('computeNext', () => {
  const from = { year: 1, month: 1, day: 1 };

  it('advances by N days', () => {
    expect(computeNext({ every: 7, unit: 'day' }, from)).toEqual({ year: 1, month: 1, day: 8 });
  });

  it('advances by 1 week (7-day week)', () => {
    expect(computeNext({ every: 1, unit: 'week' }, from, DEFAULT_CALENDAR)).toEqual({ year: 1, month: 1, day: 8 });
  });

  it('advances by month, landing on dayOfMonth', () => {
    const result = computeNext({ every: 1, unit: 'month', dayOfMonth: 1 }, from);
    expect(result.month).toBe(2);
    expect(result.day).toBe(1);
  });

  it('advances multiple months', () => {
    const result = computeNext({ every: 3, unit: 'month', dayOfMonth: 15 }, from);
    expect(result.month).toBe(4);
    expect(result.day).toBe(15);
  });

  it('advances by year to a specific month+day', () => {
    const result = computeNext({ every: 1, unit: 'year', month: 6, day: 21 }, from);
    expect(result.month).toBe(6);
    expect(result.day).toBe(21);
  });

  it('wraps month overflow into next year', () => {
    const result = computeNext({ every: 2, unit: 'month', dayOfMonth: 1 }, { year: 1, month: 11, day: 1 });
    expect(result.year).toBe(2);
    expect(result.month).toBe(1);
  });

  it('defaults to day unit for unknown unit', () => {
    expect(computeNext({ every: 3, unit: 'unknown' }, from)).toEqual(computeNext({ every: 3, unit: 'day' }, from));
  });
});

describe('seasonForMonth', () => {
  it('assigns the four seasons across a 12-month calendar', () => {
    expect(seasonForMonth(1, DEFAULT_CALENDAR)).toBe('spring');
    expect(seasonForMonth(4, DEFAULT_CALENDAR)).toBe('summer');
    expect(seasonForMonth(7, DEFAULT_CALENDAR)).toBe('autumn');
    expect(seasonForMonth(10, DEFAULT_CALENDAR)).toBe('winter');
  });

  it('scales to calendars with a different month count', () => {
    expect(seasonForMonth(1, GENERIC_FANTASY_CALENDAR)).toBe('spring');
    expect(seasonForMonth(9, GENERIC_FANTASY_CALENDAR)).toBe('winter');
  });
});

describe('rollWeather', () => {
  it('always returns an option from the biome/season table', () => {
    for (let i = 0; i < 20; i++) {
      const w = rollWeather('arctic', 'winter', () => i / 20);
      expect(['clear', 'overcast', 'rain', 'storm', 'snow']).toContain(w);
    }
  });

  it('is deterministic for a fixed rng', () => {
    expect(rollWeather('temperate', 'summer', () => 0)).toBe(rollWeather('temperate', 'summer', () => 0));
  });

  it('falls back to the temperate table for an unknown biome', () => {
    expect(rollWeather('lunar', 'spring', () => 0)).toBe(rollWeather('temperate', 'spring', () => 0));
  });
});

describe('applyDailyTick jitter bounds', () => {
  function makeDoc(dailyAvg) {
    const settlement = {
      _schemaVersion: CURRENT_SCHEMA_VERSION,
      stores: [{ id: 's1', income: { balance: 0, dailyAvg, lastTick: 0 } }],
    };
    let stored = settlement;
    return {
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('earned income stays within ±15% of dailyAvg', async () => {
    const dailyAvg = 100;
    const doc = makeDoc(dailyAvg);
    for (let i = 0; i < 60; i++) {
      doc.getStored().stores[0].income.balance = 0;
      await applyDailyTick(doc, 1);
      const balance = doc.getStored().stores[0].income.balance;
      expect(balance).toBeGreaterThanOrEqual(85);
      expect(balance).toBeLessThanOrEqual(115);
    }
  });

  it('scales linearly with days', async () => {
    const doc = makeDoc(100);
    doc.getStored().stores[0].income.balance = 0;
    await applyDailyTick(doc, 3);
    const balance = doc.getStored().stores[0].income.balance;
    expect(balance).toBeGreaterThanOrEqual(255);
    expect(balance).toBeLessThanOrEqual(345);
  });

  it('leaves lastTick as a recent timestamp', async () => {
    const before = Date.now();
    const doc = makeDoc(10);
    await applyDailyTick(doc, 1);
    expect(doc.getStored().stores[0].income.lastTick).toBeGreaterThanOrEqual(before);
  });
});

describe('applyDailyTick production credit', () => {
  function makeDocWithProduction(production, population, startingGp = 0) {
    const settlement = {
      _schemaVersion: CURRENT_SCHEMA_VERSION,
      stores: [],
      production,
      population,
      treasury: { cp: 0, sp: 0, gp: startingGp, pp: 0 },
    };
    let stored = settlement;
    return {
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('credits treasury: production.length × population / 1000 gp per day', async () => {
    const doc = makeDocWithProduction(['grain', 'livestock'], 1000);
    await applyDailyTick(doc, 1);
    expect(doc.getStored().treasury.gp).toBe(2);
  });

  it('scales credit with days', async () => {
    const doc = makeDocWithProduction(['grain', 'livestock', 'ore'], 2000);
    await applyDailyTick(doc, 5);
    expect(doc.getStored().treasury.gp).toBe(30);
  });

  it('adds to existing treasury balance', async () => {
    const doc = makeDocWithProduction(['grain'], 1000, 100);
    await applyDailyTick(doc, 1);
    expect(doc.getStored().treasury.gp).toBe(101);
  });

  it('skips credit when production is empty', async () => {
    const doc = makeDocWithProduction([], 1000);
    await applyDailyTick(doc, 1);
    expect(doc.getStored().treasury.gp).toBe(0);
  });

  it('skips credit when population is zero', async () => {
    const doc = makeDocWithProduction(['grain', 'livestock'], 0);
    await applyDailyTick(doc, 1);
    expect(doc.getStored().treasury.gp).toBe(0);
  });
});

describe('applyFestival', () => {
  globalThis.ChatMessage = { create: async () => {} };
  globalThis.game = { users: [] };

  function makeDoc(morale, unrest, gp) {
    const settlement = {
      _schemaVersion: CURRENT_SCHEMA_VERSION,
      stats: { morale, unrest },
      treasury: { cp: 0, sp: 0, gp, pp: 0 },
    };
    let stored = settlement;
    return {
      name: 'Test Town',
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('boosts morale and eases unrest by the payload amounts', async () => {
    const doc = makeDoc(50, 20, 500);
    await applyFestival(doc, { moraleBoost: 15, unrestReduction: 4, gpCost: 100 });
    expect(doc.getStored().stats.morale).toBe(65);
    expect(doc.getStored().stats.unrest).toBe(16);
  });

  it('drains the gp cost from the treasury', async () => {
    const doc = makeDoc(50, 20, 500);
    await applyFestival(doc, { gpCost: 300 });
    expect(doc.getStored().treasury.gp).toBe(200);
  });

  it('clamps morale at 100 and unrest at 0', async () => {
    const doc = makeDoc(95, 2, 100);
    await applyFestival(doc, { moraleBoost: 20, unrestReduction: 10, gpCost: 0 });
    expect(doc.getStored().stats.morale).toBe(100);
    expect(doc.getStored().stats.unrest).toBe(0);
  });

  it('defaults to a 10 morale boost and 1 unrest reduction with no cost', async () => {
    const doc = makeDoc(50, 20, 500);
    await applyFestival(doc, {});
    expect(doc.getStored().stats.morale).toBe(60);
    expect(doc.getStored().stats.unrest).toBe(19);
    expect(doc.getStored().treasury.gp).toBe(500);
  });
});

describe('applyTax', () => {
  globalThis.ChatMessage = { create: async () => {} };
  globalThis.game = { users: [] };

  function makeDoc(balance, bannerImage = null) {
    const settlement = {
      _schemaVersion: CURRENT_SCHEMA_VERSION,
      stores: [{ id: 's1', closed: false, income: { balance, dailyAvg: 5, lastTick: 0 } }],
      treasury: { cp: 0, sp: 0, gp: 0, pp: 0 },
      stats: { unrest: 10 },
      bannerImage,
    };
    let stored = settlement;
    return {
      name: 'Test Town',
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('collects a percentage of positive store balances into the treasury', async () => {
    const doc = makeDoc(1000);
    const { collected } = await applyTax(doc, { taxType: 'income', ratePct: 10 });
    expect(collected).toBe(100);
    expect(doc.getStored().treasury.gp).toBe(100);
    expect(doc.getStored().stores[0].income.balance).toBe(900);
  });

  it('raises unrest by 1', async () => {
    const doc = makeDoc(1000);
    await applyTax(doc, { ratePct: 5 });
    expect(doc.getStored().stats.unrest).toBe(11);
  });

  it('posts a chat card including the banner thumbnail when one is set (#108)', async () => {
    const posted = [];
    globalThis.ChatMessage = { create: async (data) => { posted.push(data); } };
    const doc = makeDoc(1000, 'worlds/foo/banner.webp');
    await applyTax(doc, { ratePct: 10 });
    expect(posted).toHaveLength(1);
    expect(posted[0].content).toContain('worlds/foo/banner.webp');
  });

  it('omits the thumbnail when no banner is set', async () => {
    const posted = [];
    globalThis.ChatMessage = { create: async (data) => { posted.push(data); } };
    const doc = makeDoc(1000);
    await applyTax(doc, { ratePct: 10 });
    expect(posted[0].content).not.toContain('<img');
  });
});

describe('migrateSettlement', () => {
  it('adds bannerImage: null when upgrading from schema 4', () => {
    const migrated = migrateSettlement({ _schemaVersion: 4, kind: 'town' });
    expect(migrated._schemaVersion).toBe(CURRENT_SCHEMA_VERSION);
    expect(migrated.bannerImage).toBeNull();
  });

  it('leaves an existing bannerImage untouched', () => {
    const migrated = migrateSettlement({ _schemaVersion: 4, bannerImage: 'foo.webp' });
    expect(migrated.bannerImage).toBe('foo.webp');
  });
});

describe('settlementNoteIcon / settlementNoteTooltip', () => {
  it('maps each kind to a distinct icon path', () => {
    const icons = new Set(['village', 'town', 'city', 'nation'].map(settlementNoteIcon));
    expect(icons.size).toBe(4);
  });

  it('falls back to the city icon for an unknown kind', () => {
    expect(settlementNoteIcon('nonsense')).toBe(settlementNoteIcon('city'));
  });

  it('formats the tooltip as "<Kind> · Population <n>"', () => {
    expect(settlementNoteTooltip({ kind: 'town', population: 1500 })).toBe('Town · Population 1,500');
  });
});

describe('generateHooks', () => {
  it('returns 3 unique hooks for a bare settlement', () => {
    const hooks = generateHooks(sanitizeSettlement({}), 'Test Town');
    expect(hooks).toHaveLength(3);
    expect(new Set(hooks).size).toBe(3);
  });

  it('mentions a black-market store when one exists', () => {
    const settlement = sanitizeSettlement({ stores: [{ name: 'The Quiet Crate', isBlackMarket: true }] });
    const hooks = generateHooks(settlement, 'Test Town');
    expect(hooks.some(h => h.includes('The Quiet Crate'))).toBe(true);
  });

  it('mentions a closed store when one exists', () => {
    const settlement = sanitizeSettlement({ stores: [{ name: 'Old Mill', closed: true }] });
    const hooks = generateHooks(settlement, 'Test Town');
    expect(hooks.some(h => h.includes('Old Mill'))).toBe(true);
  });

  it('mentions a district when one exists', () => {
    const settlement = sanitizeSettlement({ districts: [{ name: 'Dockside' }] });
    const hooks = generateHooks(settlement, 'Test Town');
    expect(hooks.some(h => h.includes('Dockside'))).toBe(true);
  });
});
