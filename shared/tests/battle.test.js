import { describe, it, expect } from 'vitest';
import { sanitizeArmy } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/army.js';
import { sanitizeSettlement } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/sanitizer.js';
import { armyPower, resolveBattle, applyBattleResult, resolveSiege, applySiegeResult, TERRAIN_TYPES }
  from '../../modules/Pf2eNationsAndCitiesMaker/scripts/battle.js';

describe('TERRAIN_TYPES', () => {
  it('lists the supported terrains', () => {
    expect(TERRAIN_TYPES).toEqual(['plains', 'forest', 'hills', 'mountains', 'urban', 'swamp']);
  });
});

describe('armyPower', () => {
  it('scales with count, level, and morale', () => {
    const army = sanitizeArmy({ units: [{ type: 'spearmen', count: 10, level: 2, morale: 100 }] });
    expect(armyPower(army, 'plains')).toBe(10 * 2 * 1);
  });

  it('halves power at 50% morale', () => {
    const army = sanitizeArmy({ units: [{ type: 'spearmen', count: 10, level: 1, morale: 50 }] });
    expect(armyPower(army, 'plains')).toBe(5);
  });

  it('applies the terrain modifier for the unit type', () => {
    const army = sanitizeArmy({ units: [{ type: 'cavalry', count: 10, level: 1, morale: 100 }] });
    expect(armyPower(army, 'plains')).toBeCloseTo(12.5);
    expect(armyPower(army, 'forest')).toBeCloseTo(6);
  });

  it('sums across mixed unit types', () => {
    const army = sanitizeArmy({ units: [
      { type: 'spearmen', count: 10, level: 1, morale: 100 },
      { type: 'archers', count: 5, level: 1, morale: 100 },
    ] });
    expect(armyPower(army, 'plains')).toBe(15);
  });

  it('adds a flat bonus per commander level (#78)', () => {
    const army = sanitizeArmy({ units: [{ type: 'spearmen', count: 10, level: 1, morale: 100 }] });
    expect(armyPower(army, 'plains')).toBe(10);
    expect(armyPower(army, 'plains', 5)).toBe(60);
  });

  it('ignores a negative commander level', () => {
    const army = sanitizeArmy({ units: [{ type: 'spearmen', count: 10, level: 1, morale: 100 }] });
    expect(armyPower(army, 'plains', -5)).toBe(10);
  });
});

describe('resolveBattle', () => {
  it('declares the stronger army the winner and inflicts heavier casualties on the loser', () => {
    const attacker = sanitizeArmy({ units: [{ type: 'spearmen', count: 100, level: 5, morale: 100 }] });
    const defender = sanitizeArmy({ units: [{ type: 'spearmen', count: 100, level: 1, morale: 100 }] });
    const result = resolveBattle(attacker, defender, 'plains');
    expect(result.winner).toBe('attacker');
    expect(result.defenderCasualtyPct).toBeGreaterThan(result.attackerCasualtyPct);
    expect(result.defenderCasualties[0].lost).toBeGreaterThan(result.attackerCasualties[0].lost);
  });

  it('declares a draw when powers are equal and inflicts equal casualties', () => {
    const attacker = sanitizeArmy({ units: [{ type: 'spearmen', count: 100, level: 1, morale: 100 }] });
    const defender = sanitizeArmy({ units: [{ type: 'spearmen', count: 100, level: 1, morale: 100 }] });
    const result = resolveBattle(attacker, defender, 'plains');
    expect(result.winner).toBe('draw');
    expect(result.attackerCasualtyPct).toBe(result.defenderCasualtyPct);
  });

  it('is deterministic for the same inputs', () => {
    const attacker = sanitizeArmy({ units: [{ type: 'archers', count: 40, level: 3, morale: 80 }] });
    const defender = sanitizeArmy({ units: [{ type: 'cavalry', count: 20, level: 4, morale: 90 }] });
    const a = resolveBattle(attacker, defender, 'forest');
    const b = resolveBattle(attacker, defender, 'forest');
    expect(a).toEqual(b);
  });

  it('never loses more units than an army has', () => {
    const attacker = sanitizeArmy({ units: [{ type: 'spearmen', count: 3, level: 1, morale: 100 }] });
    const defender = sanitizeArmy({ units: [{ type: 'spearmen', count: 500, level: 20, morale: 100 }] });
    const result = resolveBattle(attacker, defender, 'plains');
    expect(result.attackerCasualties[0].lost).toBeLessThanOrEqual(3);
    expect(result.attackerCasualties[0].remaining).toBeGreaterThanOrEqual(0);
  });

  it('results in no casualties when both armies are empty', () => {
    const attacker = sanitizeArmy({ units: [] });
    const defender = sanitizeArmy({ units: [] });
    const result = resolveBattle(attacker, defender, 'plains');
    expect(result.winner).toBe('draw');
    expect(result.attackerCasualtyPct).toBe(0);
    expect(result.defenderCasualtyPct).toBe(0);
  });

  it('lets a high-level commander swing an otherwise-even fight (#78)', () => {
    const attacker = sanitizeArmy({ units: [{ type: 'spearmen', count: 100, level: 1, morale: 100 }] });
    const defender = sanitizeArmy({ units: [{ type: 'spearmen', count: 100, level: 1, morale: 100 }] });
    const result = resolveBattle(attacker, defender, 'plains', 20, 0);
    expect(result.winner).toBe('attacker');
  });
});

describe('applyBattleResult', () => {
  function makeArmyDoc(units) {
    let stored = { kind: 'army', units };
    return {
      name: 'Test Army',
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('writes remaining unit counts back onto both army documents', async () => {
    const attackerDoc = makeArmyDoc([{ id: 'u1', type: 'spearmen', count: 100, level: 5, morale: 100 }]);
    const defenderDoc = makeArmyDoc([{ id: 'u2', type: 'spearmen', count: 100, level: 1, morale: 100 }]);
    const attacker = sanitizeArmy(attackerDoc.getStored());
    const defender = sanitizeArmy(defenderDoc.getStored());
    const result = resolveBattle(attacker, defender, 'plains');

    await applyBattleResult(attackerDoc, defenderDoc, result);

    expect(attackerDoc.getStored().units[0].count).toBe(result.attackerCasualties[0].remaining);
    expect(defenderDoc.getStored().units[0].count).toBe(result.defenderCasualties[0].remaining);
  });
});

describe('resolveSiege', () => {
  it('deals hardness-reduced damage and does not occupy an unbroken settlement', () => {
    const attacker = sanitizeArmy({ units: [{ type: 'spearmen', count: 10, level: 5, morale: 100 }] });
    const settlement = sanitizeSettlement({ stats: { hp: 500, maxHp: 500, hardness: 10, damageThreshold: 5 } });
    const result = resolveSiege(attacker, settlement, 'urban');
    expect(result.damage).toBeGreaterThan(0);
    expect(result.hpAfter).toBe(result.hpBefore - result.damage);
    expect(result.occupied).toBe(false);
  });

  it('deals no damage when it fails to clear the damage threshold', () => {
    const attacker = sanitizeArmy({ units: [{ type: 'spearmen', count: 1, level: 1, morale: 100 }] });
    const settlement = sanitizeSettlement({ stats: { hp: 500, maxHp: 500, hardness: 50, damageThreshold: 999 } });
    const result = resolveSiege(attacker, settlement, 'plains');
    expect(result.damage).toBe(0);
    expect(result.hpAfter).toBe(result.hpBefore);
  });

  it('occupies the settlement once hp reaches 0', () => {
    const attacker = sanitizeArmy({ units: [{ type: 'siege', count: 50, level: 10, morale: 100 }] });
    const settlement = sanitizeSettlement({ stats: { hp: 10, maxHp: 500, hardness: 0, damageThreshold: 0 } });
    const result = resolveSiege(attacker, settlement, 'urban');
    expect(result.hpAfter).toBe(0);
    expect(result.occupied).toBe(true);
  });

  it('never drops hp below 0', () => {
    const attacker = sanitizeArmy({ units: [{ type: 'siege', count: 999, level: 20, morale: 100 }] });
    const settlement = sanitizeSettlement({ stats: { hp: 5, maxHp: 500, hardness: 0, damageThreshold: 0 } });
    const result = resolveSiege(attacker, settlement, 'urban');
    expect(result.hpAfter).toBe(0);
  });

  it('folds in a commander bonus (#78)', () => {
    const attacker = sanitizeArmy({ units: [{ type: 'spearmen', count: 1, level: 1, morale: 100 }] });
    const settlement = sanitizeSettlement({ stats: { hp: 500, maxHp: 500, hardness: 5, damageThreshold: 1 } });
    const withoutCommander = resolveSiege(attacker, settlement, 'plains', 0);
    const withCommander = resolveSiege(attacker, settlement, 'plains', 10);
    expect(withCommander.damage).toBeGreaterThan(withoutCommander.damage);
  });
});

describe('applySiegeResult', () => {
  function makeSettlementDoc(stats) {
    let stored = sanitizeSettlement({ stats });
    return {
      name: 'Test City',
      getFlag: () => stored,
      setFlag: async (_scope, _key, data) => { stored = data; },
      getStored: () => stored,
    };
  }

  it('writes hp and occupied back onto the settlement document', async () => {
    const settlementDoc = makeSettlementDoc({ hp: 10, maxHp: 500, hardness: 0, damageThreshold: 0 });
    const attacker = sanitizeArmy({ units: [{ type: 'siege', count: 50, level: 10, morale: 100 }] });
    const settlement = sanitizeSettlement(settlementDoc.getFlag());
    const result = resolveSiege(attacker, settlement, 'urban');

    await applySiegeResult(settlementDoc, result);

    expect(settlementDoc.getStored().stats.hp).toBe(0);
    expect(settlementDoc.getStored().stats.occupied).toBe(true);
  });
});
