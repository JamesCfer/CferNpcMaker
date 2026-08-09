import { describe, it, expect } from 'vitest';
import { sanitizeArmy, totalUnitCount, recruitmentCost, totalDailyWage, applyArmyWages, UNIT_TYPES,
         cmpDate, computeArrivalDate, totalDailyFoodNeed, settlementFoodCapacity, isSupplied, applyArmySupply,
         MERCENARY_COMPANIES, getMercenaryCompany, mercenaryHireCost, hireMercenaryCompany, isContractExpired }
  from '../../modules/Pf2eNationsAndCitiesMaker/scripts/army.js';
import { sanitizeSettlement } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/sanitizer.js';

describe('sanitizeArmy defaults', () => {
  it('defaults to an empty roster with kind army', () => {
    const a = sanitizeArmy({});
    expect(a.kind).toBe('army');
    expect(a.units).toEqual([]);
    expect(a.stationedAt).toBeNull();
    expect(a.mode).toBe('garrison');
    expect(a.destination).toBeNull();
    expect(a.arrivalDate).toBeNull();
    expect(a.ownerNationId).toBeNull();
    expect(a.supplySource).toBeNull();
    expect(a.contract).toEqual({ active: false, companyId: null, companyLabel: '', expiresDate: null });
  });

  it('accepts a field-army mode and rejects unknown modes', () => {
    expect(sanitizeArmy({ mode: 'field' }).mode).toBe('field');
    expect(sanitizeArmy({ mode: 'rampaging' }).mode).toBe('garrison');
  });

  it('keeps a well-formed destination and arrivalDate', () => {
    const a = sanitizeArmy({ destination: 'journal1', arrivalDate: { year: 4710, month: 3, day: 12 } });
    expect(a.destination).toBe('journal1');
    expect(a.arrivalDate).toEqual({ year: 4710, month: 3, day: 12 });
  });

  it('discards a malformed arrivalDate', () => {
    expect(sanitizeArmy({ arrivalDate: { year: 4710 } }).arrivalDate).toBeNull();
  });

  it('sanitizes unit fields with safe defaults', () => {
    const a = sanitizeArmy({ units: [{ type: 'archers', count: '4', level: '2', morale: '150' }] });
    expect(a.units[0].type).toBe('archers');
    expect(a.units[0].count).toBe(4);
    expect(a.units[0].level).toBe(2);
    expect(a.units[0].morale).toBe(100);
  });

  it('rejects an unknown unit type and falls back to spearmen', () => {
    expect(sanitizeArmy({ units: [{ type: 'dragons' }] }).units[0].type).toBe('spearmen');
  });
});

describe('totalUnitCount', () => {
  it('sums counts across all unit entries', () => {
    const army = sanitizeArmy({ units: [{ type: 'spearmen', count: 10 }, { type: 'archers', count: 5 }] });
    expect(totalUnitCount(army)).toBe(15);
  });
});

describe('recruitmentCost', () => {
  it('scales gp and pop cost by count', () => {
    expect(recruitmentCost('spearmen', 10)).toEqual({ gp: 50, pop: 10 });
  });

  it('uses the siege unit rate', () => {
    expect(recruitmentCost('siege', 2)).toEqual({ gp: 300, pop: 4 });
  });

  it('falls back to the spearmen rate for an unknown type', () => {
    expect(recruitmentCost('dragons', 1)).toEqual(recruitmentCost('spearmen', 1));
  });
});

describe('totalDailyWage', () => {
  it('sums wages across unit types', () => {
    const army = sanitizeArmy({ units: [{ type: 'spearmen', count: 10 }, { type: 'mages', count: 2 }] });
    expect(totalDailyWage(army)).toBe(10 * 1 + 2 * 5);
  });
});

describe('applyArmyWages', () => {
  function makeArmyDoc(units, stationedAt = 's1') {
    const stored = { kind: 'army', stationedAt, units };
    return { getFlag: () => stored, setFlag: async () => {} };
  }

  function makeSettlementDoc(gp) {
    const settlement = { treasury: { cp: 0, sp: 0, gp, pp: 0 } };
    let stored = settlement;
    return {
      name: 'Test Town',
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('drains the daily wage from the stationed settlement treasury', async () => {
    const armyDoc = makeArmyDoc([{ type: 'spearmen', count: 10 }]);
    const settlementDoc = makeSettlementDoc(1000);
    const drained = await applyArmyWages(armyDoc, settlementDoc, 1);
    expect(drained).toBe(10);
    expect(settlementDoc.getStored().treasury.gp).toBe(990);
  });

  it('scales with days', async () => {
    const armyDoc = makeArmyDoc([{ type: 'mages', count: 2 }]);
    const settlementDoc = makeSettlementDoc(1000);
    await applyArmyWages(armyDoc, settlementDoc, 5);
    expect(settlementDoc.getStored().treasury.gp).toBe(1000 - 2 * 5 * 5);
  });

  it('does nothing when the army has no units', async () => {
    const armyDoc = makeArmyDoc([]);
    const settlementDoc = makeSettlementDoc(1000);
    const drained = await applyArmyWages(armyDoc, settlementDoc, 1);
    expect(drained).toBe(0);
    expect(settlementDoc.getStored().treasury.gp).toBe(1000);
  });

  it('returns 0 and does not throw when there is no settlement', async () => {
    const armyDoc = makeArmyDoc([{ type: 'spearmen', count: 10 }], null);
    expect(await applyArmyWages(armyDoc, null, 1)).toBe(0);
  });
});

describe('UNIT_TYPES', () => {
  it('covers the five PF2e-flavoured unit types', () => {
    expect(UNIT_TYPES).toEqual(['spearmen', 'archers', 'cavalry', 'mages', 'siege']);
  });
});

describe('cmpDate', () => {
  it('orders by year, then month, then day', () => {
    expect(cmpDate({ year: 1, month: 1, day: 1 }, { year: 2, month: 1, day: 1 })).toBe(-1);
    expect(cmpDate({ year: 1, month: 2, day: 1 }, { year: 1, month: 1, day: 1 })).toBe(1);
    expect(cmpDate({ year: 1, month: 1, day: 1 }, { year: 1, month: 1, day: 1 })).toBe(0);
  });
});

describe('computeArrivalDate', () => {
  it('adds travel days within the same month', () => {
    const cal = { daysPerMonth: [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31] };
    expect(computeArrivalDate({ year: 4710, month: 1, day: 1 }, 7, cal)).toEqual({ year: 4710, month: 1, day: 8 });
  });

  it('rolls over into the next month and year', () => {
    const cal = { daysPerMonth: [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31] };
    expect(computeArrivalDate({ year: 4710, month: 12, day: 28 }, 7, cal)).toEqual({ year: 4711, month: 1, day: 4 });
  });

  it('falls back to a flat 30-day month with no calendarDef', () => {
    expect(computeArrivalDate({ year: 1, month: 1, day: 25 }, 10, null)).toEqual({ year: 1, month: 2, day: 5 });
  });
});

describe('totalDailyFoodNeed', () => {
  it('scales with total unit count', () => {
    const army = sanitizeArmy({ units: [{ type: 'spearmen', count: 100 }, { type: 'archers', count: 100 }] });
    expect(totalDailyFoodNeed(army)).toBeCloseTo(200 * 0.05);
  });
});

describe('settlementFoodCapacity', () => {
  it('scales with population', () => {
    expect(settlementFoodCapacity(sanitizeSettlement({ population: 2000 }))).toBe(10);
  });

  it('handles a missing settlement', () => {
    expect(settlementFoodCapacity(null)).toBe(0);
  });
});

describe('isSupplied', () => {
  it('is true when the source can spare enough food', () => {
    const army = sanitizeArmy({ units: [{ type: 'spearmen', count: 100 }] });
    const source = sanitizeSettlement({ population: 2000 });
    expect(isSupplied(army, source)).toBe(true);
  });

  it('is false when the source is undersupplied', () => {
    const army = sanitizeArmy({ units: [{ type: 'spearmen', count: 100000 }] });
    const source = sanitizeSettlement({ population: 2000 });
    expect(isSupplied(army, source)).toBe(false);
  });

  it('is false when the source settlement is occupied', () => {
    const army = sanitizeArmy({ units: [{ type: 'spearmen', count: 1 }] });
    const source = sanitizeSettlement({ population: 2000, stats: { occupied: true } });
    expect(isSupplied(army, source)).toBe(false);
  });

  it('is false with no source settlement', () => {
    const army = sanitizeArmy({ units: [{ type: 'spearmen', count: 1 }] });
    expect(isSupplied(army, null)).toBe(false);
  });
});

describe('applyArmySupply', () => {
  function makeArmyDoc(units, supplySource = 's1') {
    let stored = { kind: 'army', supplySource, units };
    return { getFlag: () => stored, setFlag: async (_scope, _key, data) => { stored = data; }, getStored: () => stored };
  }

  function makeSourceDoc(population, occupied = false) {
    const stored = { population, stats: { occupied } };
    return { getFlag: () => stored };
  }

  it('does nothing without a supply source', async () => {
    const armyDoc = makeArmyDoc([{ id: 'u1', type: 'spearmen', count: 10, morale: 50 }], null);
    expect(await applyArmySupply(armyDoc, makeSourceDoc(2000), 1)).toBeNull();
    expect(armyDoc.getStored().units[0].morale).toBe(50);
  });

  it('does nothing for an army with no units', async () => {
    const armyDoc = makeArmyDoc([]);
    expect(await applyArmySupply(armyDoc, makeSourceDoc(2000), 1)).toBeNull();
  });

  it('restores morale when supplied', async () => {
    const armyDoc = makeArmyDoc([{ id: 'u1', type: 'spearmen', count: 10, morale: 50 }]);
    const result = await applyArmySupply(armyDoc, makeSourceDoc(2000), 1);
    expect(result.supplied).toBe(true);
    expect(armyDoc.getStored().units[0].morale).toBe(55);
  });

  it('drains morale when cut off and does not lose troops above 0 morale', async () => {
    const armyDoc = makeArmyDoc([{ id: 'u1', type: 'spearmen', count: 10, morale: 50 }]);
    const result = await applyArmySupply(armyDoc, makeSourceDoc(2000, true), 1);
    expect(result.supplied).toBe(false);
    expect(result.starved).toBe(false);
    expect(armyDoc.getStored().units[0].morale).toBe(40);
    expect(armyDoc.getStored().units[0].count).toBe(10);
  });

  it('starves troops once morale bottoms out at 0', async () => {
    const armyDoc = makeArmyDoc([{ id: 'u1', type: 'spearmen', count: 10, morale: 5 }]);
    const result = await applyArmySupply(armyDoc, makeSourceDoc(2000, true), 1);
    expect(result.starved).toBe(true);
    expect(armyDoc.getStored().units[0].morale).toBe(0);
    expect(armyDoc.getStored().units[0].count).toBeLessThan(10);
  });

  it('treats a missing source settlement as cut off', async () => {
    const armyDoc = makeArmyDoc([{ id: 'u1', type: 'spearmen', count: 10, morale: 50 }]);
    const result = await applyArmySupply(armyDoc, null, 1);
    expect(result.supplied).toBe(false);
  });
});

describe('mercenaryHireCost', () => {
  it('multiplies the daily rate by contract length', () => {
    const company = getMercenaryCompany('skirmishers');
    expect(mercenaryHireCost(company, 10)).toBe(150);
  });

  it('floors the contract length at 1 day', () => {
    const company = getMercenaryCompany('skirmishers');
    expect(mercenaryHireCost(company, 0)).toBe(company.costPerDay);
  });
});

describe('getMercenaryCompany', () => {
  it('resolves every preset by id', () => {
    for (const c of MERCENARY_COMPANIES) expect(getMercenaryCompany(c.id)).toBe(c);
  });

  it('returns null for an unknown id', () => {
    expect(getMercenaryCompany('dragon-riders')).toBeNull();
  });
});

describe('hireMercenaryCompany', () => {
  function makeSettlementDoc(gp) {
    let stored = { treasury: { cp: 0, sp: 0, gp, pp: 0 } };
    return {
      id: 'settlement1',
      name: 'Test Town',
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  const calendarState = { currentDate: { year: 4710, month: 1, day: 1 }, calendarDef: { daysPerMonth: Array(12).fill(30) } };

  it('drains the lump-sum cost and creates a stationed army under contract', async () => {
    const created = [];
    globalThis.JournalEntry = { create: async (data) => { const doc = { id: 'army1', name: data.name, flags: data.flags }; created.push(doc); return doc; } };

    const settlementDoc = makeSettlementDoc(1000);
    const result = await hireMercenaryCompany(settlementDoc, 'skirmishers', 10, calendarState);

    expect(result.cost).toBe(150);
    expect(settlementDoc.getStored().treasury.gp).toBe(850);
    expect(created).toHaveLength(1);
    const army = created[0].flags.Pf2eNationsAndCitiesMaker.settlement;
    expect(army.stationedAt).toBe('settlement1');
    expect(army.contract.active).toBe(true);
    expect(army.contract.companyId).toBe('skirmishers');
    expect(army.contract.expiresDate).toEqual({ year: 4710, month: 1, day: 11 });
  });

  it('refuses to hire when the settlement cannot afford it', async () => {
    const settlementDoc = makeSettlementDoc(10);
    const result = await hireMercenaryCompany(settlementDoc, 'skirmishers', 10, calendarState);
    expect(result).toEqual({ error: 'insufficient-funds', cost: 150 });
    expect(settlementDoc.getStored().treasury.gp).toBe(10);
  });

  it('returns null for an unknown company', async () => {
    const settlementDoc = makeSettlementDoc(1000);
    expect(await hireMercenaryCompany(settlementDoc, 'dragon-riders', 10, calendarState)).toBeNull();
  });
});

describe('isContractExpired', () => {
  it('is false when the contract is not active', () => {
    const army = sanitizeArmy({});
    expect(isContractExpired(army, { year: 4711, month: 1, day: 1 })).toBe(false);
  });

  it('is true once currentDate reaches expiresDate', () => {
    const army = sanitizeArmy({ contract: { active: true, expiresDate: { year: 4710, month: 6, day: 1 } } });
    expect(isContractExpired(army, { year: 4710, month: 6, day: 1 })).toBe(true);
    expect(isContractExpired(army, { year: 4710, month: 5, day: 30 })).toBe(false);
  });
});
