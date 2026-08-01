import { describe, it, expect } from 'vitest';
import { sanitizeArmy } from '../../modules/Pf2eNationsAndCitiesMaker/scripts/army.js';
import { armyPower, resolveBattle, applyBattleResult, TERRAIN_TYPES }
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
