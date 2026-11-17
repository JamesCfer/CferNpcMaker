// @vitest-environment jsdom
import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';
import { mapDnd5eNpc } from '../core/mappers/dnd5e-npc.js';
import { mapPf2eNpc } from '../core/mappers/pf2e-npc.js';
import { mapItem } from '../core/mappers/items.js';
import { mapStore } from '../core/mappers/store.js';
import { mapSettlement } from '../core/mappers/settlement.js';
import { mapHero6e, hdcToHdImport } from '../core/mappers/hero6e.js';
import { buildMapRequest, mapFinal, ROLL20_GRID } from '../core/mappers/map.js';
import { GENERATORS, readForm, endpointFor, getGenerator } from '../core/catalog.js';

const here = dirname(fileURLToPath(import.meta.url));
const fixture = name => JSON.parse(readFileSync(join(here, 'fixtures', name), 'utf8'));

describe('D&D 5e NPC', () => {
  const plan = mapDnd5eNpc(fixture('dnd5e-npc.json'));
  const drop = JSON.parse(plan.drop.attributes.drop_data);

  it('uses the sheet compendium drop with a direct fallback', () => {
    expect(plan.kind).toBe('character');
    expect(plan.drop.trigger).toEqual({ name: 'drop_category', value: 'Monsters' });
    expect(plan.drop.verify).toBe('npc_name');
    expect(plan.fallback.attributes.npc).toBe('1');
    expect(plan.fallback.attributes.npc_name).toBe(plan.character.name);
  });

  it('writes drop data in the format the sheet parses', () => {
    for (const k of ['STR', 'DEX', 'CON', 'INT', 'WIS', 'CHA']) expect(String(drop[k])).toMatch(/^\d+$/);
    expect(drop.AC).toMatch(/^\d+/);
    expect(drop.HP).toMatch(/^\d+ \(\d+d\d+/);
    expect(drop['Challenge Rating']).toBeTruthy();
    const actions = JSON.parse(drop['data-Actions']);
    expect(Array.isArray(actions)).toBe(true);
    expect(actions.length).toBeGreaterThan(0);
    for (const a of actions) expect(a.Name).toBeTruthy();
  });

  it('never leaks the builder key', () => {
    expect(JSON.stringify(plan).toLowerCase()).not.toContain('x-builder-key');
  });

  it('builds a D&D 2024 sheet block', () => {
    const b = plan.beacon;
    expect(b).toBeTruthy();
    const abilities = b.integrants.filter(i => i.type === 'Ability Score');
    expect(abilities.map(a => a.ability).sort()).toEqual(['Charisma', 'Constitution', 'Dexterity', 'Intelligence', 'Strength', 'Wisdom']);
    for (const a of abilities) expect(a.valueFormula.flatValue).toEqual(expect.any(Number));
    expect(b.integrants.find(i => i.type === 'Armor Class')._id).toBe('armor-class-override-total');
    expect(b.integrants.find(i => i.type === 'Hit Points').valueFormula.flatValue).toBeGreaterThan(0);
    const refs = new Set(b.integrants.map(i => i._ref).filter(Boolean));
    for (const i of b.integrants) if (i._parentRef) expect(refs.has(i._parentRef)).toBe(true);
    for (const r of b.traitRefs) expect(refs.has(r)).toBe(true);
    expect(b.store.npc.challengeRating).toBeTruthy();
  });
});

describe('PF2e NPC', () => {
  it('builds npcdrop_data for a plain creature', () => {
    const plan = mapPf2eNpc(fixture('pf2e-goblin-warrior.json'));
    expect(plan.drop.trigger.name).toBe('npcdrop_data');
    const d = JSON.parse(plan.drop.trigger.value);
    expect(d.level).toBe('-1');
    expect(d.armor_class).toBe('16');
    expect(d.hit_points_max).toBe('6');
    expect(d.speed).toBe('25');
    expect(JSON.parse(d['data-melee-strikes'])[0].weapon).toBe('Dogslicer');
    expect(plan.attributes.sheet_type).toBe('npc');
    expect(plan.fallback.attributes.hit_points_max).toBeDefined();
  });

  it('carries spellcasting', () => {
    const d = JSON.parse(mapPf2eNpc(fixture('pf2e-dragon-caster.json')).drop.trigger.value);
    expect(d.spell_dc).toBe('28');
    expect(JSON.parse(d['data-spell']).length).toBeGreaterThan(0);
  });
});

describe('items and stores', () => {
  const sword = { name: 'Flame Tongue', type: 'weapon', system: { rarity: 'rare', price: { value: 5000, denomination: 'gp' }, weight: { value: 3 },
    description: { value: '<p>A blade wreathed in fire.</p>' }, damage: { base: { number: 1, denomination: 8, types: ['slashing'] } } } };

  it('makes a 5e item handout with a give recipe', () => {
    const plan = mapItem('dnd5e')({ foundryItem: sword });
    expect(plan.kind).toBe('handouts');
    expect(plan.handouts[0].notes).toContain('Flame Tongue');
    expect(plan.handouts[0].notes).toContain('A blade wreathed in fire.');
    expect(plan.handouts[0].give.drop.trigger.value).toBe('Items');
  });

  it('makes a PF2e item give recipe for both sheet types', () => {
    const plan = mapItem('pf2e')({ foundryItem: { name: 'Healing Potion', type: 'consumable', system: { level: { value: 1 }, price: { value: { gp: 4 } }, traits: { value: ['healing'] } } } });
    const give = plan.handouts[0].give;
    expect(give.rowsBySheetType.npc.section).toBe('items-worn');
    expect(give.rowsBySheetType.pc.section).toBe('items-other');
  });

  it('makes a store handout plus one per item', () => {
    const plan = mapStore('dnd5e')({ store: { name: 'The Hot Anvil', storeType: 'blacksmith', owner: { name: 'Brakka', race: 'Dwarf' }, inventory: [sword, { name: 'Dagger', type: 'weapon', system: {} }] } });
    expect(plan.handouts).toHaveLength(3);
    expect(plan.handouts[0].notes).toContain('Brakka');
    expect(plan.handouts[0].notes).toContain('Blacksmith');
    expect(plan.handouts[1].give).toBeTruthy();
  });
});

describe('settlements', () => {
  it('renders a nation and its cities as handouts', () => {
    const plan = mapSettlement({ nation: { name: 'Varos', government: 'Monarchy', id: 'x1', economy: { exports: ['grain', 'iron'] } },
      childCities: [{ name: 'Port Hale', population: 12000 }] }, { kind: 'nation' });
    expect(plan.handouts.map(h => h.name)).toEqual(['Varos', 'Port Hale']);
    expect(plan.handouts[0].notes).toContain('Monarchy');
    expect(plan.handouts[0].notes).toContain('grain');
    expect(plan.handouts[0].notes).not.toContain('x1');
  });
});

describe('Hero 6e', () => {
  const xml = `<?xml version="1.0" encoding="UTF-16"?><CHARACTER version="6.0"><BASIC_CONFIGURATION/><CHARACTER_INFO CHARACTER_NAME="Test Hero" ALTERNATE_IDENTITIES="" PLAYER_NAME="" HEIGHT="" WEIGHT="" HAIR_COLOR="" EYE_COLOR="" CAMPAIGN_NAME="" GENRE=""><BACKGROUND>Born in the rain.</BACKGROUND></CHARACTER_INFO>
<CHARACTERISTICS><STR XMLID="STR" ID="1" LEVELS="15"/><DEX XMLID="DEX" ID="2" LEVELS="8"/><CON XMLID="CON" ID="3" LEVELS="5"/><SPD XMLID="SPD" ID="4" LEVELS="2"/><BODY XMLID="BODY" ID="5" LEVELS="0"/></CHARACTERISTICS>
<SKILLS/><PERKS/><TALENTS/><MARTIALARTS/><POWERS/><DISADVANTAGES/><EQUIPMENT/></CHARACTER>`;

  it('reads characteristics as base plus levels', () => {
    const data = hdcToHdImport(xml, new DOMParser());
    expect(data.character.character_name).toBe('Test Hero');
    const plan = mapHero6e({ hdcXml: xml }, { level: 150 }, new DOMParser());
    expect(plan.attributes.strength).toBe('25');
    expect(plan.attributes.dexterity).toBe('18');
    expect(plan.attributes.speed).toBe('4');
    expect(plan.attributes.body).toBe('10');
    expect(plan.heroImport.startsWith('!hero --import ')).toBe(true);
    expect(plan.hdc.xml).toBe(xml);
  });
});

describe('battle maps', () => {
  it('builds the same request shape as the Foundry module', () => {
    const req = buildMapRequest({ description: 'A forest road ambush', width: 30, height: 20 });
    expect(req).toMatchObject({ perspective: 'top-down', setting: 'outdoor', squares: [30, 20], dials: { detail: 3, clutter: 3, wear: 3 } });
    expect(() => buildMapRequest({})).toThrow();
  });

  it('turns the trace into Roll20 pixels', () => {
    const final = { image: { b64: 'AAAA', mediaType: 'image/webp' }, scene: { squares: [10, 10], grid: { distance: 5, units: 'ft' } },
      trace: { walls: [{ kind: 'wall', x1: 0, y1: 0, x2: 1000, y2: 0 }, { kind: 'door', x1: 500, y1: 0, x2: 500, y2: 100 }], lights: [{ kind: 'torch', x: 500, y: 500 }] } };
    const plan = mapFinal(final, { name: 'Road Ambush' });
    expect(plan.kind).toBe('map');
    expect(plan.map.image.filename).toBe('road-ambush.webp');
    expect(plan.map.squares).toEqual([10, 10]);
    expect(plan.map.walls).toHaveLength(2);
    expect(plan.map.walls[0].c).toEqual([0, 0, 700, 0]);
    expect(plan.map.lights).toHaveLength(1);
    const max = 10 * ROLL20_GRID;
    for (const w of plan.map.walls) for (const v of w.c) { expect(v).toBeGreaterThanOrEqual(0); expect(v).toBeLessThanOrEqual(max); }
    expect(JSON.stringify(plan)).not.toContain('AAAA');
  });
});

describe('catalog', () => {
  it('has a mapper and an endpoint for every generator', () => {
    for (const g of GENERATORS) {
      expect(typeof g.map).toBe('function');
      expect(endpointFor(g, { kind: 'town' })).toMatch(/^https:\/\/foundryrelay\.dedicated2\.com\/webhook\//);
    }
  });

  it('routes a nation with cities to the nation builder', () => {
    const g = getGenerator('pf2e-settlement');
    expect(endpointFor(g, { kind: 'nation', includeChildCities: true })).toContain('/nation-builder');
    expect(endpointFor(g, { kind: 'town' })).toContain('/city-builder');
  });

  it('hides the settlement maker until its server workflow exists', () => {
    expect(getGenerator('pf2e-settlement').available).toBe(false);
  });

  it('needs a description for stores, which the server requires', () => {
    expect(() => readForm(getGenerator('dnd5e-store'), { name: 'X' })).toThrow(/description/);
    expect(() => readForm(getGenerator('pf2e-store'), { name: 'X' })).toThrow(/description/);
  });

  it('rejects a store with no stock', () => {
    expect(() => mapStore('dnd5e')({ store: { name: 'Empty', inventory: [] } })).toThrow(/no stock/);
  });

  it('enforces required fields', () => {
    expect(() => readForm(getGenerator('dnd5e-npc'), { name: 'X' })).toThrow(/description/);
    const f = readForm(getGenerator('dnd5e-npc'), { description: 'An orc', level: '5' });
    expect(f.level).toBe(5);
    expect(getGenerator('dnd5e-npc').payload(f)).toMatchObject({ cr: 5, casterType: 'none' });
  });
});
