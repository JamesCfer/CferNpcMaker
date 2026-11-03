import { describe, it, expect } from 'vitest';
import { costSummary, buildMapRequest, buildSceneData, measuredLabel, sceneFileName } from '../../modules/BattleMapMaker/scripts/map-logic.js';

describe('Battle Map Maker logic', () => {
  it('uses the shared action costs', () => {
    expect(costSummary({ quality: 'standard' })).toMatchObject({ mockups: 3, final: 6, total: 9 });
    expect(costSummary({ quality: 'economy' })).toMatchObject({ mockups: 3, final: 4, total: 7 });
    expect(costSummary({ quality: 'standard', includeMockups: false })).toMatchObject({ mockups: 0, final: 6, total: 6 });
  });
  it('builds a content-only request with valid defaults and limits', () => {
    const map = buildMapRequest({ description: '  forest camp  ', perspective: 'top-down', genre: 'other', genreOther: '  mythic  ', style: 'painted', mood: 'day', width: 1, height: 99, dials: { detail: 7, clutter: 0, wear: '', color: 5 } });
    expect(map).toEqual({ description: 'forest camp', perspective: 'top-down', genre: 'mythic', style: 'painted', mood: 'day', squares: [5, 60], setting: 'outdoor', dials: { detail: 5, clutter: 1, wear: 3 } });
  });
  it('creates square-grid scene data in either unit system', () => {
    const input = { name: 'Camp', src: 'worlds/test/battle-maps/camp.webp', squares: [15, 12], imageSize: { width: 1400, height: 1100 } };
    expect(buildSceneData({ ...input, units: 'imperial' })).toMatchObject({ name: 'Camp', width: 1500, height: 1200, padding: 0, background: { src: input.src }, grid: { type: 1, size: 100, distance: 5, units: 'ft' } });
    expect(buildSceneData({ ...input, units: 'metric' }).grid).toEqual({ type: 1, size: 100, distance: 1.5, units: 'm' });
  });
  it('labels measured dimensions', () => {
    expect(measuredLabel({ squares: [15, 12], grid: { distance: 5, units: 'ft' } })).toBe('Measured 75 by 60 ft (15 x 12 squares)');
  });
  it('makes a safe dated file name', () => {
    expect(sceneFileName('Ruined Keep / River? 💥', new Date('2026-11-03T12:34:56Z'))).toBe('ruined-keep-river-20261103-123456Z.webp');
  });
});
