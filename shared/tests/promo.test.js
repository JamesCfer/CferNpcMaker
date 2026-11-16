import { describe, it, expect } from 'vitest';
import { ALL_MODULES } from '../scripts/core/home-data.js';
import { promoModulesFor, pickStripModule, isStripDismissed, dismissStrip,
         renderPromoList, PROMO_DISMISS_KEY } from '../scripts/core/promo.js';

const installed = entries => new Map(entries.map(([id, active]) => [id, { active }]));
const ids = list => list.map(m => m.id);

describe('promoModulesFor', () => {
  it('lists same-system siblings plus the system-agnostic map maker', () => {
    const out = ids(promoModulesFor({ currentId: 'DnD5eNpcMaker', systemId: 'dnd5e', modules: new Map() }));
    expect(out).toEqual(['Dnd5eItemGenerator', 'Dnd5eStoreGenerator', 'ai-battle-map-maker']);
  });

  it('hides active modules and flags installed-but-disabled ones', () => {
    const out = promoModulesFor({ currentId: 'Pf2eNpcMaker', systemId: 'pf2e',
      modules: installed([['Pf2eItemGenerator', true], ['Pf2eStoreGenerator', false]]) });
    expect(ids(out)).not.toContain('Pf2eItemGenerator');
    expect(out.find(m => m.id === 'Pf2eStoreGenerator').installedInactive).toBe(true);
  });

  it('gives the map maker every module for the running system', () => {
    const out = ids(promoModulesFor({ currentId: 'ai-battle-map-maker', systemId: 'hero6efoundryvttv2', modules: new Map() }));
    expect(out).toEqual(['Hero6eNpcMaker']);
  });

  it('never lists the running module', () => {
    for (const m of ALL_MODULES) {
      const sys = m.systems[0] || 'pf2e';
      expect(ids(promoModulesFor({ currentId: m.id, systemId: sys, modules: new Map() }))).not.toContain(m.id);
    }
  });

  it('has a manifest and pitch for every module', () => {
    for (const m of ALL_MODULES) {
      expect(m.manifest).toMatch(/^https:\/\/github\.com\/JamesCfer\/.+\/module\.json$/);
      expect(m.pitch).toBeTruthy();
    }
  });
});

describe('pickStripModule', () => {
  it('prefers the map maker, then rotates, and skips installed modules', () => {
    const pf2e = promoModulesFor({ currentId: 'Pf2eNpcMaker', systemId: 'pf2e', modules: new Map() });
    expect(pickStripModule(pf2e).id).toBe('ai-battle-map-maker');
    const noMaps = pf2e.filter(m => m.id !== 'ai-battle-map-maker');
    expect(pickStripModule(noMaps, 0).id).not.toBe(pickStripModule(noMaps, 86_400_000).id);
    expect(pickStripModule(noMaps.map(m => ({ ...m, installedInactive: true })))).toBeNull();
  });
});

describe('strip dismissal', () => {
  it('stays hidden for 30 days', () => {
    const store = new Map();
    const ls = { getItem: k => store.get(k) ?? null, setItem: (k, v) => store.set(k, v) };
    expect(isStripDismissed(0, ls)).toBe(false);
    dismissStrip(0, ls);
    expect(store.has(PROMO_DISMISS_KEY)).toBe(true);
    expect(isStripDismissed(29 * 86_400_000, ls)).toBe(true);
    expect(isStripDismissed(31 * 86_400_000, ls)).toBe(false);
  });
});

describe('renderPromoList', () => {
  it('escapes and shows a copy button or enable hint', () => {
    const html = renderPromoList([
      { ...ALL_MODULES[0], installedInactive: false },
      { ...ALL_MODULES[1], installedInactive: true },
    ]);
    expect(html).toContain('data-copy-manifest="https://github.com/JamesCfer/Pf2eNpcMaker');
    expect(html).toContain('Enable it in Manage Modules');
    expect(html).toContain('D&amp;D 5e NPC');
  });
});
