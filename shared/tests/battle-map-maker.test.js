import { describe, it, expect } from 'vitest';
import { costSummary, buildMapRequest, measuredLabel, sceneFileName } from '../../modules/ai-battle-map-maker/scripts/map-logic.js';

describe('AI Battle Map Maker logic', () => {
  it('uses the shared action costs', () => {
    expect(costSummary({ quality: 'standard' })).toMatchObject({ mockups: 3, final: 6, total: 9 });
    expect(costSummary({ quality: 'economy' })).toMatchObject({ mockups: 3, final: 4, total: 7 });
    expect(costSummary({ quality: 'standard', includeMockups: false })).toMatchObject({ mockups: 0, final: 6, total: 6 });
    expect(costSummary({ quality: 'standard', walls: true })).toMatchObject({ final: 6, total: 10 });
    expect(costSummary({ quality: 'economy', walls: true })).toMatchObject({ final: 4, total: 8 });
    expect(costSummary({ quality: 'draft', walls: true })).toMatchObject({ mockups: 3, final: 2, total: 6 });
    expect(costSummary({ quality: 'high', walls: true })).toMatchObject({ mockups: 3, final: 6, total: 10 });
    expect(costSummary({ walls: true, perspective: 'isometric' }).total).toBe(9);
    expect(costSummary({ walls: true }).text).toBe('Mockups 3 uses + Standard final 6 uses + walls 1 use = 10 uses');
  });
  it('summarizes every mockup count with each quality and walls choice', () => {
    for (const [count, mockupCost] of [[1, 1], [2, 2], [4, 3]]) {
      for (const [quality, final] of [['standard', 6], ['economy', 4], ['draft', 2], ['high', 6]]) {
        for (const walls of [false, true]) {
          const summary = costSummary({ mockups: count, quality, walls });
          expect(summary).toMatchObject({ mockups: mockupCost, final, total: mockupCost + final + Number(walls) });
          expect(summary.text).toContain(`Mockups ${mockupCost} uses`);
        }
      }
    }
  });
  it('builds a content-only request with valid defaults and limits', () => {
    const map = buildMapRequest({ description: '  forest camp  ', perspective: 'top-down', genre: 'other', genreOther: '  mythic  ', style: 'painted', mood: 'day', width: 1, height: 99, dials: { detail: 7, clutter: 0, wear: '', color: 5 } });
    expect(map).toEqual({ description: 'forest camp', perspective: 'top-down', genre: 'mythic', style: 'painted', mood: 'day', squares: [5, 60], setting: 'outdoor', dials: { detail: 5, clutter: 1, wear: 3 } });
  });
  it('trims and caps an editable layout', () => {
    expect(buildMapRequest({ description: 'Keep', layout: '  A'.repeat(700) + '  ' }).layout).toHaveLength(1200);
    expect(buildMapRequest({ description: 'Keep', layout: '  west gate  ' }).layout).toBe('west gate');
    expect(buildMapRequest({ description: 'Keep', layout: '  ' })).not.toHaveProperty('layout');
  });
  it('labels measured dimensions', () => {
    expect(measuredLabel({ squares: [15, 12], grid: { distance: 5, units: 'ft' } })).toBe('Measured 75 by 60 ft (15 x 12 squares)');
  });
  it('makes a safe dated file name', () => {
    expect(sceneFileName('Ruined Keep / River? 💥', new Date('2026-11-03T12:34:56Z'))).toBe('ruined-keep-river-20261103-123456Z.webp');
  });
});
