import { describe, it, expect } from 'vitest';
import { sanitizeStore, tieredPrice, inventoryValue, storeTypeLabel,
         STORE_TYPES, PRICE_TIERS } from '../../modules/Pf2eStoreGenerator/scripts/sanitizer.js';
import { sanitizeStore as sanitizeStore5e }
  from '../../modules/Dnd5eStoreGenerator/scripts/sanitizer.js';

describe('sanitizeStore defaults', () => {
  it('fills every field from an empty payload', () => {
    const s = sanitizeStore({});
    expect(s.name).toBe('New Shop');
    expect(s.type).toBe('general');
    expect(s.priceTier).toBe('standard');
    expect(s.owner).toEqual({ name: 'Unknown Owner', description: '', actorId: null });
    expect(s.hours).toEqual({ open: '08:00', close: '20:00', daysClosed: [] });
    expect(s.inventory).toEqual([]);
    expect(s.staff).toEqual([]);
    expect(s.services).toEqual([]);
    expect(s.rumors).toEqual([]);
    expect(s.closed).toBe(false);
    expect(s.income.balance).toBe(50);
    expect(s.id).toMatch(/^store-/);
  });

  it('falls back to the form values when the payload omits them', () => {
    const s = sanitizeStore({}, {
      name: 'The Iron Anvil', storeType: 'blacksmith',
      priceTier: 'luxury', level: 5, settlement: 'Ashenford',
      description: 'A soot-stained smithy.',
    });
    expect(s.name).toBe('The Iron Anvil');
    expect(s.type).toBe('blacksmith');
    expect(s.priceTier).toBe('luxury');
    expect(s.level).toBe(5);
    expect(s.settlement).toBe('Ashenford');
    expect(s.ai.prompt).toBe('A soot-stained smithy.');
  });

  it('unwraps a { store: {...} } envelope', () => {
    expect(sanitizeStore({ store: { name: 'The Alembic', type: 'alchemist' } }).name).toBe('The Alembic');
  });

  it('rejects unknown store types and price tiers', () => {
    const s = sanitizeStore({ type: 'starbase', priceTier: 'free' });
    expect(s.type).toBe('general');
    expect(s.priceTier).toBe('standard');
  });

  it('keeps every documented store type', () => {
    for (const type of STORE_TYPES) expect(sanitizeStore({ type }).type).toBe(type);
    for (const tier of PRICE_TIERS) expect(sanitizeStore({ priceTier: tier }).priceTier).toBe(tier);
  });
});

describe('sanitizeStore inventory', () => {
  it('backfills malformed inventory rows', () => {
    const s = sanitizeStore({ inventory: [{}, { name: '  ', price: 'lots', stock: -3 }] });
    expect(s.inventory).toHaveLength(2);
    expect(s.inventory[0].name).toBe('Unnamed Good');
    expect(s.inventory[0].price).toBe(1);
    expect(s.inventory[1].price).toBe(1);
    expect(s.inventory[1].stock).toBe(0);
    expect(s.inventory[0].id).not.toBe(s.inventory[1].id);
  });

  it('drops a non-array inventory', () => {
    expect(sanitizeStore({ inventory: 'a sword' }).inventory).toEqual([]);
  });

  it('caps the shelves at 60 lines', () => {
    const inventory = Array.from({ length: 90 }, (_, i) => ({ name: `Good ${i}` }));
    expect(sanitizeStore({ inventory }).inventory).toHaveLength(60);
  });

  it('normalises rarity per system', () => {
    expect(sanitizeStore({ inventory: [{ rarity: 'legendary' }] }).inventory[0].rarity).toBe('common');
    expect(sanitizeStore5e({ inventory: [{ rarity: 'legendary' }] }).inventory[0].rarity).toBe('legendary');
  });
});

describe('sanitizeStore hours and staff', () => {
  it('accepts HH:MM and rejects anything else', () => {
    expect(sanitizeStore({ hours: { open: '06:30', close: '23:45' } }).hours)
      .toMatchObject({ open: '06:30', close: '23:45' });
    expect(sanitizeStore({ hours: { open: 'dawn', close: '25:99' } }).hours)
      .toMatchObject({ open: '08:00', close: '20:00' });
  });

  it('defaults an unknown shift to day', () => {
    const s = sanitizeStore({ staff: [{ name: 'Rell', shift: 'whenever' }] });
    expect(s.staff[0].shift).toBe('day');
    expect(s.staff[0].role).toBe('Assistant');
  });
});

describe('tieredPrice', () => {
  it('applies the wealth-tier multiplier', () => {
    expect(tieredPrice(100, 'low')).toBe(75);
    expect(tieredPrice(100, 'standard')).toBe(100);
    expect(tieredPrice(100, 'high')).toBe(150);
    expect(tieredPrice(100, 'luxury')).toBe(200);
  });

  it('never rounds a priced good down to free', () => {
    expect(tieredPrice(1, 'low')).toBe(1);
  });

  it('keeps free goods free and treats unknown tiers as standard', () => {
    expect(tieredPrice(0, 'luxury')).toBe(0);
    expect(tieredPrice(10, 'mystery')).toBe(10);
  });
});

describe('inventoryValue', () => {
  it('sums tiered price × stock', () => {
    const store = sanitizeStore({
      priceTier: 'high',
      inventory: [{ name: 'Longsword', price: 20, stock: 3 }, { name: 'Rope', price: 1, stock: 10 }],
    });
    expect(inventoryValue(store)).toBe(30 * 3 + 2 * 10);
  });

  it('is zero for an empty shop', () => {
    expect(inventoryValue(sanitizeStore({}))).toBe(0);
  });
});

describe('storeTypeLabel', () => {
  it('capitalises a type and falls back to Other', () => {
    expect(storeTypeLabel('blacksmith')).toBe('Blacksmith');
    expect(storeTypeLabel('')).toBe('Other');
  });
});
