import { describe, it, expect } from 'vitest';
import {
  coinsToCp, cpToCoins, cpLabel as cpLabelPf2e,
  parsePriceLabelCp as parsePf2e, effectiveCp as effectivePf2e,
} from '../../modules/Pf2eStoreGenerator/scripts/transactions.js';
import {
  priceToCp, cpLabel as cpLabel5e,
  parsePriceLabelCp as parse5e, effectiveCp as effective5e,
} from '../../modules/Dnd5eStoreGenerator/scripts/transactions.js';

describe('PF2e store currency', () => {
  it('converts coins objects to copper', () => {
    expect(coinsToCp({ pp: 1, gp: 2, sp: 3, cp: 4 })).toBe(1234);
    expect(coinsToCp({ gp: 15 })).toBe(1500);
    expect(coinsToCp(null)).toBe(0);
    expect(coinsToCp({})).toBe(0);
  });

  it('converts copper back to canonical coins', () => {
    expect(cpToCoins(1234)).toEqual({ gp: 12, sp: 3, cp: 4 });
    expect(cpToCoins(100)).toEqual({ gp: 1 });
    expect(cpToCoins(0)).toEqual({});
  });

  it('labels copper amounts', () => {
    expect(cpLabelPf2e(1234)).toBe('12 gp, 3 sp, 4 cp');
    expect(cpLabelPf2e(0)).toBe('free');
  });

  it('parses legacy price labels', () => {
    expect(parsePf2e('3 gp, 5 sp')).toBe(350);
    expect(parsePf2e('1 pp')).toBe(1000);
    expect(parsePf2e('—')).toBe(0);
    expect(parsePf2e('')).toBe(0);
  });

  it('applies the store price multiplier', () => {
    expect(effectivePf2e({ priceCp: 100 }, { priceMultiplier: 1.5 })).toBe(150);
    expect(effectivePf2e({ priceCp: 100 }, {})).toBe(100);
    expect(effectivePf2e({ price: '2 gp' }, { priceMultiplier: 2 })).toBe(400);
  });
});

describe('D&D 5e store currency', () => {
  it('converts price objects to copper', () => {
    expect(priceToCp({ value: 25, denomination: 'gp' })).toBe(2500);
    expect(priceToCp({ value: 5, denomination: 'sp' })).toBe(50);
    expect(priceToCp({ value: 1, denomination: 'pp' })).toBe(1000);
    expect(priceToCp({ value: 2, denomination: 'ep' })).toBe(100);
    expect(priceToCp(3)).toBe(300);
  });

  it('labels copper amounts', () => {
    expect(cpLabel5e(2510)).toBe('25 gp, 1 sp');
    expect(cpLabel5e(0)).toBe('free');
  });

  it('parses legacy price labels', () => {
    expect(parse5e('25 gp')).toBe(2500);
    expect(parse5e('1 sp, 5 cp')).toBe(15);
    expect(parse5e('—')).toBe(0);
  });

  it('applies the store price multiplier', () => {
    expect(effective5e({ priceCp: 2500 }, { priceMultiplier: 0.5 })).toBe(1250);
    expect(effective5e({ price: '10 gp' }, {})).toBe(1000);
  });
});
