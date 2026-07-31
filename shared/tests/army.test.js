import { describe, it, expect } from 'vitest';
import { sanitizeArmy, totalUnitCount, recruitmentCost, totalDailyWage, applyArmyWages, UNIT_TYPES,
         cmpDate, computeArrivalDate }
  from '../../modules/Pf2eNationsAndCitiesMaker/scripts/army.js';

describe('sanitizeArmy defaults', () => {
  it('defaults to an empty roster with kind army', () => {
    const a = sanitizeArmy({});
    expect(a.kind).toBe('army');
    expect(a.units).toEqual([]);
    expect(a.stationedAt).toBeNull();
    expect(a.mode).toBe('garrison');
    expect(a.destination).toBeNull();
    expect(a.arrivalDate).toBeNull();
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
