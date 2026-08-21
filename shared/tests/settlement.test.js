import { describe, it, expect } from 'vitest';
import { sanitizeSettlement }  from '../../modules/Pf2eNationsAndCitiesMaker/scripts/sanitizer.js';
import { computeNext, DEFAULT_CALENDAR, GENERIC_FANTASY_CALENDAR,
         seasonForMonth, rollWeather } from '../../modules/Pf2eCalendarTimeline/scripts/scheduler.js';
import { applyDailyTick, applyFestival, applyTax } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/economy.js';
import { CURRENT_SCHEMA_VERSION, migrateSettlement } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/migrations.js';
import { generateHooks } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/hooks.js';
import { processTreatyExpiry, declareWar, negotiatePeace, applySuccession } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/diplomacy.js';
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

  it('defaults occupied to false and preserves a truthy value (#77)', () => {
    expect(sanitizeSettlement({}).stats.occupied).toBe(false);
    expect(sanitizeSettlement({ stats: { occupied: true } }).stats.occupied).toBe(true);
  });

  it('defaults occupiedBy to null and preserves a valid nation id (#80)', () => {
    expect(sanitizeSettlement({}).stats.occupiedBy).toBeNull();
    expect(sanitizeSettlement({ stats: { occupiedBy: 'nation1' } }).stats.occupiedBy).toBe('nation1');
    expect(sanitizeSettlement({ stats: { occupiedBy: 42 } }).stats.occupiedBy).toBeNull();
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

  it('defaults borderColor and rejects malformed hex values', () => {
    expect(sanitizeSettlement({}).borderColor).toBe('#8b1a1a');
    expect(sanitizeSettlement({ borderColor: 'not-a-colour' }).borderColor).toBe('#8b1a1a');
  });

  it('preserves a valid borderColor hex value', () => {
    expect(sanitizeSettlement({ borderColor: '#2f6b3a' }).borderColor).toBe('#2f6b3a');
  });

  it('defaults borderDrawingId to null', () => {
    expect(sanitizeSettlement({}).borderDrawingId).toBeNull();
  });

  it('preserves a valid borderDrawingId', () => {
    expect(sanitizeSettlement({ borderDrawingId: 'drawing123' }).borderDrawingId).toBe('drawing123');
  });

  it('defaults a store jobs list to empty', () => {
    expect(sanitizeSettlement({ stores: [{}] }).stores[0].jobs).toEqual([]);
  });

  it('sanitizes job fields with safe defaults', () => {
    const s = sanitizeSettlement({ stores: [{ type: 'guild', jobs: [{ title: 'Clear the cellar', reward: '25' }] }] });
    const job = s.stores[0].jobs[0];
    expect(job.title).toBe('Clear the cellar');
    expect(job.reward).toBe(25);
    expect(job.status).toBe('open');
    expect(job.assignedActorId).toBeNull();
  });

  it('rejects an unknown job status and falls back to open', () => {
    const s = sanitizeSettlement({ stores: [{ jobs: [{ status: 'rampaging' }] }] });
    expect(s.stores[0].jobs[0].status).toBe('open');
  });

  it('defaults relations to an empty list (#83)', () => {
    expect(sanitizeSettlement({}).relations).toEqual([]);
  });

  it('sanitizes relation fields with safe defaults', () => {
    const s = sanitizeSettlement({ relations: [{ nationId: 'nation1', relation: 'ally', score: '50' }] });
    expect(s.relations).toEqual([{ nationId: 'nation1', relation: 'ally', score: 50 }]);
  });

  it('drops relation entries without a nationId', () => {
    expect(sanitizeSettlement({ relations: [{ relation: 'ally' }] }).relations).toEqual([]);
  });

  it('rejects an unknown relation value and falls back to neutral', () => {
    const s = sanitizeSettlement({ relations: [{ nationId: 'nation1', relation: 'rampaging' }] });
    expect(s.relations[0].relation).toBe('neutral');
  });

  it('clamps relation score to -100..100', () => {
    const s = sanitizeSettlement({ relations: [{ nationId: 'nation1', score: 500 }, { nationId: 'nation2', score: -500 }] });
    expect(s.relations[0].score).toBe(100);
    expect(s.relations[1].score).toBe(-100);
  });

  it('defaults treaties to an empty list (#84)', () => {
    expect(sanitizeSettlement({}).treaties).toEqual([]);
  });

  it('sanitizes treaty fields with safe defaults', () => {
    const s = sanitizeSettlement({
      treaties: [{ partnerNationId: 'nation1', kind: 'trade', terms: 'open borders',
        signedOn: { year: 1, month: 2, day: 3 }, expiresOn: { year: 2, month: 3, day: 4 } }],
    });
    expect(s.treaties).toHaveLength(1);
    const t = s.treaties[0];
    expect(t.partnerNationId).toBe('nation1');
    expect(t.kind).toBe('trade');
    expect(t.terms).toBe('open borders');
    expect(t.signedOn).toEqual({ year: 1, month: 2, day: 3 });
    expect(t.expiresOn).toEqual({ year: 2, month: 3, day: 4 });
  });

  it('drops treaty entries without a partnerNationId', () => {
    expect(sanitizeSettlement({ treaties: [{ kind: 'trade' }] }).treaties).toEqual([]);
  });

  it('rejects an unknown treaty kind and falls back to non-aggression', () => {
    const s = sanitizeSettlement({ treaties: [{ partnerNationId: 'nation1', kind: 'friendship' }] });
    expect(s.treaties[0].kind).toBe('non-aggression');
  });

  it('defaults a treaty with no expiry to a null expiresOn', () => {
    const s = sanitizeSettlement({ treaties: [{ partnerNationId: 'nation1' }] });
    expect(s.treaties[0].expiresOn).toBeNull();
  });

  it('defaults vassalNationIds to an empty list and suzerainNationId to null (#86)', () => {
    const s = sanitizeSettlement({});
    expect(s.vassalNationIds).toEqual([]);
    expect(s.suzerainNationId).toBeNull();
  });

  it('preserves valid vassalNationIds and suzerainNationId', () => {
    const s = sanitizeSettlement({ vassalNationIds: ['nation2', 'nation3'], suzerainNationId: 'nation1' });
    expect(s.vassalNationIds).toEqual(['nation2', 'nation3']);
    expect(s.suzerainNationId).toBe('nation1');
  });

  it('defaults claims to an empty list (#87)', () => {
    expect(sanitizeSettlement({}).claims).toEqual([]);
  });

  it('sanitizes claim fields with safe defaults', () => {
    const s = sanitizeSettlement({
      claims: [{ targetSettlementId: 'city1', kind: 'dynastic', notes: 'former royal seat' }],
    });
    expect(s.claims).toHaveLength(1);
    const c = s.claims[0];
    expect(c.targetSettlementId).toBe('city1');
    expect(c.kind).toBe('dynastic');
    expect(c.notes).toBe('former royal seat');
  });

  it('drops claim entries without a targetSettlementId', () => {
    expect(sanitizeSettlement({ claims: [{ kind: 'dynastic' }] }).claims).toEqual([]);
  });

  it('rejects an unknown claim kind and falls back to historical', () => {
    const s = sanitizeSettlement({ claims: [{ targetSettlementId: 'city1', kind: 'imaginary' }] });
    expect(s.claims[0].kind).toBe('historical');
  });

  it('defaults heir to a null actorId and empty name (#91)', () => {
    expect(sanitizeSettlement({}).heir).toEqual({ actorId: null, name: '' });
  });

  it('preserves a set heir', () => {
    const s = sanitizeSettlement({ heir: { actorId: 'actor1', name: 'Princess Yvaine' } });
    expect(s.heir).toEqual({ actorId: 'actor1', name: 'Princess Yvaine' });
  });

  it('defaults factions to an empty list (#90)', () => {
    expect(sanitizeSettlement({}).factions).toEqual([]);
  });

  it('sanitizes faction fields with safe defaults', () => {
    const s = sanitizeSettlement({ factions: [{ name: 'Thieves Guild', type: 'criminal', influence: '40' }] });
    expect(s.factions).toEqual([{ id: expect.any(String), name: 'Thieves Guild', type: 'criminal', influence: 40 }]);
  });

  it('rejects an unknown faction type and falls back to guild', () => {
    const s = sanitizeSettlement({ factions: [{ type: 'cabal' }] });
    expect(s.factions[0].type).toBe('guild');
  });

  it('clamps faction influence to 0..100', () => {
    const s = sanitizeSettlement({ factions: [{ influence: 500 }, { influence: -20 }] });
    expect(s.factions[0].influence).toBe(100);
    expect(s.factions[1].influence).toBe(0);
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

describe('applyDailyTick occupation drain (#80)', () => {
  function makeDoc(gp, occupied, occupiedBy) {
    const settlement = {
      _schemaVersion: CURRENT_SCHEMA_VERSION,
      stores: [],
      treasury: { cp: 0, sp: 0, gp, pp: 0 },
      stats: { occupied, occupiedBy },
    };
    let stored = settlement;
    return {
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('drains 5% of gp per day to the occupier nation treasury', async () => {
    const nationDoc = makeDoc(0, false, null);
    globalThis.game = { journal: { get: (id) => (id === 'nation1' ? nationDoc : null) } };
    const doc = makeDoc(1000, true, 'nation1');
    await applyDailyTick(doc, 1);
    expect(doc.getStored().treasury.gp).toBe(950);
    expect(nationDoc.getStored().treasury.gp).toBe(50);
  });

  it('does nothing when the settlement is not occupied', async () => {
    globalThis.game = { journal: { get: () => null } };
    const doc = makeDoc(1000, false, null);
    await applyDailyTick(doc, 1);
    expect(doc.getStored().treasury.gp).toBe(1000);
  });

  it('never drains below 0 gp', async () => {
    globalThis.game = { journal: { get: () => null } };
    const doc = makeDoc(0, true, 'nation1');
    await applyDailyTick(doc, 1);
    expect(doc.getStored().treasury.gp).toBe(0);
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

  it('adds an empty jobs list to every store when upgrading from schema 5', () => {
    const migrated = migrateSettlement({ _schemaVersion: 5, stores: [{ name: 'The Iron Anvil' }] });
    expect(migrated._schemaVersion).toBe(CURRENT_SCHEMA_VERSION);
    expect(migrated.stores[0].jobs).toEqual([]);
  });

  it('adds empty treaties/vassalNationIds and a null suzerainNationId when upgrading from schema 6', () => {
    const migrated = migrateSettlement({ _schemaVersion: 6, kind: 'nation' });
    expect(migrated._schemaVersion).toBe(CURRENT_SCHEMA_VERSION);
    expect(migrated.treaties).toEqual([]);
    expect(migrated.vassalNationIds).toEqual([]);
    expect(migrated.suzerainNationId).toBeNull();
  });

  it('adds a blank heir when upgrading from schema 7 (#91)', () => {
    const migrated = migrateSettlement({ _schemaVersion: 7, kind: 'nation' });
    expect(migrated._schemaVersion).toBe(CURRENT_SCHEMA_VERSION);
    expect(migrated.heir).toEqual({ actorId: null, name: '' });
  });

  it('leaves an existing heir untouched', () => {
    const migrated = migrateSettlement({ _schemaVersion: 7, heir: { actorId: 'actor1', name: 'Old Heir' } });
    expect(migrated.heir).toEqual({ actorId: 'actor1', name: 'Old Heir' });
  });
});

describe('processTreatyExpiry (#85)', () => {
  globalThis.ChatMessage = { create: async () => {} };
  globalThis.game = { users: [], journal: { get: () => null } };
  globalThis.Hooks = { callAll: () => {} };

  function makeDoc(treaties) {
    const settlement = { _schemaVersion: CURRENT_SCHEMA_VERSION, kind: 'nation', treaties };
    let stored = settlement;
    return {
      name: 'Test Nation',
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('drops a treaty once currentDate reaches its expiresOn', async () => {
    const doc = makeDoc([{ id: 't1', partnerNationId: 'nation2', kind: 'trade', terms: '', signedOn: null, expiresOn: { year: 1, month: 1, day: 10 } }]);
    await processTreatyExpiry(doc, doc.getStored(), { year: 1, month: 1, day: 10 });
    expect(doc.getStored().treaties).toEqual([]);
  });

  it('keeps a treaty that has not yet reached its expiresOn', async () => {
    const doc = makeDoc([{ id: 't1', partnerNationId: 'nation2', kind: 'trade', terms: '', signedOn: null, expiresOn: { year: 1, month: 2, day: 1 } }]);
    await processTreatyExpiry(doc, doc.getStored(), { year: 1, month: 1, day: 10 });
    expect(doc.getStored().treaties).toHaveLength(1);
  });

  it('keeps a treaty with no expiry indefinitely', async () => {
    const doc = makeDoc([{ id: 't1', partnerNationId: 'nation2', kind: 'non-aggression', terms: '', signedOn: null, expiresOn: null }]);
    await processTreatyExpiry(doc, doc.getStored(), { year: 99, month: 12, day: 31 });
    expect(doc.getStored().treaties).toHaveLength(1);
  });

  it('posts a chat reminder for each expired treaty', async () => {
    const posted = [];
    globalThis.ChatMessage = { create: async (data) => { posted.push(data); } };
    const doc = makeDoc([{ id: 't1', partnerNationId: 'nation2', kind: 'defensive', terms: '', signedOn: null, expiresOn: { year: 1, month: 1, day: 1 } }]);
    await processTreatyExpiry(doc, doc.getStored(), { year: 1, month: 1, day: 1 });
    expect(posted).toHaveLength(1);
    expect(posted[0].content).toContain('defensive');
  });
});

describe('declareWar (#88)', () => {
  globalThis.game = { users: [], journal: { get: () => null } };
  globalThis.Hooks = { callAll: () => {} };
  globalThis.ChatMessage = { create: async () => {} };

  function makeDoc(id, claims = [], relations = []) {
    const settlement = { _schemaVersion: CURRENT_SCHEMA_VERSION, kind: 'nation', claims, relations };
    let stored = settlement;
    return {
      id, name: `Nation ${id}`,
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('sets both nations relations to hostile', async () => {
    const doc = makeDoc('n1');
    const target = makeDoc('n2');
    await declareWar(doc, target);
    expect(doc.getStored().relations).toEqual([{ nationId: 'n2', relation: 'hostile', score: -80 }]);
    expect(target.getStored().relations).toEqual([{ nationId: 'n1', relation: 'hostile', score: -80 }]);
  });

  it('posts a chat card citing the chosen claims', async () => {
    const posted = [];
    globalThis.ChatMessage = { create: async (data) => { posted.push(data); } };
    const doc = makeDoc('n1', [{ id: 'c1', targetSettlementId: 'city1', kind: 'dynastic', notes: 'stolen throne' }]);
    const target = makeDoc('n2');
    await declareWar(doc, target, ['c1']);
    expect(posted[0].content).toContain('dynastic');
    expect(posted[0].content).toContain('stolen throne');
  });

  it('fires warDeclared with the cited claim ids', async () => {
    globalThis.ChatMessage = { create: async () => {} };
    let fired = null;
    globalThis.Hooks = { callAll: (name, payload) => { if (name === 'Pf2eNationsAndCitiesMaker.warDeclared') fired = payload; } };
    const doc = makeDoc('n1', [{ id: 'c1', targetSettlementId: 'city1', kind: 'historical' }]);
    const target = makeDoc('n2');
    await declareWar(doc, target, ['c1']);
    expect(fired).toEqual({ nationId: 'n1', targetNationId: 'n2', claimIds: ['c1'] });
    globalThis.Hooks = { callAll: () => {} };
  });
});

describe('negotiatePeace (#89)', () => {
  globalThis.game = { users: [], journal: { get: () => null } };
  globalThis.Hooks = { callAll: () => {} };
  globalThis.ChatMessage = { create: async () => {} };

  function makeDoc(id, claims = [], relations = []) {
    const settlement = { _schemaVersion: CURRENT_SCHEMA_VERSION, kind: 'nation', claims, relations, treaties: [] };
    let stored = settlement;
    return {
      id, name: `Nation ${id}`,
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('drops claims not chosen to be kept and signs a non-aggression treaty', async () => {
    const doc = makeDoc('n1', [
      { id: 'c1', targetSettlementId: 'city1', kind: 'historical' },
      { id: 'c2', targetSettlementId: 'city2', kind: 'dynastic' },
    ]);
    const target = makeDoc('n2');
    const { kept, givenUp } = await negotiatePeace(doc, target, ['c1']);
    expect(kept.map(c => c.id)).toEqual(['c1']);
    expect(givenUp.map(c => c.id)).toEqual(['c2']);
    expect(doc.getStored().claims.map(c => c.id)).toEqual(['c1']);
    expect(doc.getStored().treaties).toHaveLength(1);
    expect(doc.getStored().treaties[0].kind).toBe('non-aggression');
  });

  it('resets relations on both nations to neutral', async () => {
    const doc = makeDoc('n1', [], [{ nationId: 'n2', relation: 'hostile', score: -80 }]);
    const target = makeDoc('n2', [], [{ nationId: 'n1', relation: 'hostile', score: -80 }]);
    await negotiatePeace(doc, target);
    expect(doc.getStored().relations).toEqual([{ nationId: 'n2', relation: 'neutral', score: 0 }]);
    expect(target.getStored().relations).toEqual([{ nationId: 'n1', relation: 'neutral', score: 0 }]);
  });
});

describe('applySuccession (#91)', () => {
  globalThis.game = { users: [], journal: { get: () => null } };
  globalThis.Hooks = { callAll: () => {} };
  globalThis.ChatMessage = { create: async () => {} };

  function makeDoc(kind, government, heir) {
    const settlement = { _schemaVersion: CURRENT_SCHEMA_VERSION, kind, government, heir };
    let stored = settlement;
    return {
      id: 'n1', name: 'Test Nation',
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('promotes the heir to head of state and clears the heir slot', async () => {
    const doc = makeDoc('nation',
      { type: 'Monarchy', leaderName: 'King Aldric', leaderActorId: null },
      { actorId: 'actor9', name: 'Princess Yvaine' });
    await applySuccession(doc);
    expect(doc.getStored().government.leaderName).toBe('Princess Yvaine');
    expect(doc.getStored().government.leaderActorId).toBe('actor9');
    expect(doc.getStored().heir).toEqual({ actorId: null, name: '' });
  });

  it('does nothing when there is no heir', async () => {
    const doc = makeDoc('nation',
      { type: 'Monarchy', leaderName: 'King Aldric', leaderActorId: null },
      { actorId: null, name: '' });
    await applySuccession(doc);
    expect(doc.getStored().government.leaderName).toBe('King Aldric');
  });

  it('does nothing for a non-nation settlement', async () => {
    const doc = makeDoc('town', { leaderName: 'Mayor Jill' }, { actorId: null, name: 'Someone' });
    await applySuccession(doc);
    expect(doc.getStored().government.leaderName).toBe('Mayor Jill');
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
