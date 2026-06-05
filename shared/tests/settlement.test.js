import { describe, it, expect } from 'vitest';
import { sanitizeSettlement }  from '../../modules/Pf2eNationsAndCitiesMaker/scripts/sanitizer.js';
import { computeNext, DEFAULT_CALENDAR } from '../../modules/Pf2eCalendarTimeline/scripts/scheduler.js';
import { applyDailyTick }      from '../../modules/Pf2eNationsAndCitiesMaker/scripts/economy.js';
import { CURRENT_SCHEMA_VERSION } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/migrations.js';

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
