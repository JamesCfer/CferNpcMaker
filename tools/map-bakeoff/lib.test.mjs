import { describe, it, expect } from 'vitest';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { sceneScale, gridFit, toFoundryWalls, toFoundryLights, pickAspect, buildPrompt, imageSize, extFor, sessionCost, dialPhrases, DIALS, applyLook, lookFilters } from './lib.mjs';

const HERE = path.dirname(fileURLToPath(import.meta.url));
const prompts = JSON.parse(fs.readFileSync(path.join(HERE, 'prompts.json'), 'utf8'));
const edits = JSON.parse(fs.readFileSync(path.join(HERE, 'edits.json'), 'utf8'));

describe('map bake-off helpers', () => {
  it('picks the nearest supported aspect ratio', () => {
    expect(pickAspect([30, 20])).toBe('3:2');
    expect(pickAspect([20, 20])).toBe('1:1');
    expect(pickAspect([16, 36])).toBe('9:16');
    expect(pickAspect([24, 18])).toBe('4:3');
  });

  it('builds a prompt with the perspective, size, and no-grid rule', () => {
    const text = buildPrompt(prompts[0]);
    expect(text).toContain('top-down');
    expect(text).toContain('66 by 44');
    expect(text).toContain('Do not draw any grid lines');
  });

  it('keeps the old wording when every dial is on its default', () => {
    expect(buildPrompt(prompts[0])).toContain('Look: highly detailed; crisp textures.');
    expect(buildPrompt({ ...prompts[0], dials: { detail: 3, color: 3 } })).toBe(buildPrompt(prompts[0]));
  });

  it('turns dial values into look phrases', () => {
    const text = buildPrompt({ ...prompts[0], dials: { detail: 1, clutter: 5, wear: 5, color: 5 } });
    expect(text).toContain(DIALS.detail.levels[0]);
    expect(text).toContain(DIALS.clutter.levels[4]);
    expect(text).toContain(DIALS.wear.levels[4]);
    expect(buildPrompt({ ...prompts[0], dials: { color: 5, tint: 1 } })).toBe(buildPrompt(prompts[0]));
    expect(text).not.toContain('highly detailed;');
  });

  it('leaves pixels alone when every look dial is on 3', () => {
    const px = new Uint8ClampedArray([10, 120, 240, 255, 200, 50, 90, 255, 0, 0, 0, 255, 255, 255, 255, 255]);
    const copy = Uint8ClampedArray.from(px);
    applyLook(px, 2, 2, { detail: 5 });
    expect(px).toEqual(copy);
  });

  it('applies the look filters', () => {
    const gray = () => new Uint8ClampedArray([100, 100, 100, 255]);
    const warm = gray(); applyLook(warm, 1, 1, { tint: 5 });
    expect(warm[0]).toBeGreaterThan(warm[2]);
    const flat = new Uint8ClampedArray([200, 100, 50, 255]); applyLook(flat, 1, 1, { color: 1 });
    expect(flat[0] - flat[2]).toBeLessThan(150);
    const hi = new Uint8ClampedArray([200, 200, 200, 255]); applyLook(hi, 1, 1, { contrast: 5 });
    expect(hi[0]).toBeGreaterThan(200);
    expect(hi[3]).toBe(255);
    expect(lookFilters({ sharpness: 1 }).sharpen).toBeLessThan(0);
  });

  it('rejects unknown dials and out-of-range values', () => {
    expect(() => dialPhrases({ shininess: 2 })).toThrow(/Unknown dial/);
    expect(() => dialPhrases({ detail: 6 })).toThrow(/1 to 5/);
    expect(() => dialPhrases({ detail: 2.5 })).toThrow(/1 to 5/);
  });

  it('adds the water rule only when the scene has water', () => {
    expect(buildPrompt(prompts.find(p => p.id === 'p01-crypt'))).toContain('Never draw water as a flat black area');
    expect(buildPrompt(prompts.find(p => p.id === 'p05-spaceship'))).not.toContain('Never draw water');
  });

  it('converts vision lights into Foundry ambient lights', () => {
    const out = toFoundryLights([
      { kind: 'torch', x: 500, y: 250, color: '#ff8800' },
      { kind: 'disco ball', x: 0, y: 1000 },
      { kind: 'lamp', x: 1200, y: 10 },
    ], { squares: [30, 20] });
    expect(out).toHaveLength(2);
    expect(out[0]).toMatchObject({ kind: 'torch', x: 1500, y: 500, config: { bright: 10, dim: 20, color: '#ff8800' } });
    expect(out[0].config.animation.type).toBe('torch');
    expect(out[1]).toMatchObject({ kind: 'lamp', x: 0, y: 2000, config: { color: '#ffb35c' } });
  });

  it('converts vision walls into Foundry walls and closes corners', () => {
    const walls = toFoundryWalls([
      { kind: 'wall', x1: 100, y1: 100, x2: 500, y2: 100 },
      { kind: 'wall', x1: 505, y1: 103, x2: 505, y2: 900 },
      { kind: 'door', x1: 200, y1: 900, x2: 300, y2: 900 },
      { kind: 'window', x1: 600, y1: 100, x2: 700, y2: 100 },
      { kind: 'wall', x1: 10, y1: 10, x2: 11, y2: 11 },
      { kind: 'wall', x1: 'a', y1: 0, x2: 1, y2: 1 },
    ], { squares: [10, 10] });
    expect(walls).toHaveLength(4);
    expect(walls[1].c.slice(0, 2)).toEqual(walls[0].c.slice(2, 4));
    expect(walls[2]).toMatchObject({ door: 1, sight: 20 });
    expect(walls[3]).toMatchObject({ door: 0, sight: 0, light: 0, move: 20 });
  });

  it('fits maps to 100 pixels per square and sharpens big enlargements', () => {
    expect(gridFit(2400, 1792, [24, 18])).toEqual({ width: 2400, height: 1800, scale: 1, sharpen: 0 });
    const big = gridFit(2400, 1792, [40, 30]);
    expect(big).toMatchObject({ width: 4000, height: 3000, sharpen: 0.6 });
    expect(big.scale).toBeCloseTo(1.667, 2);
  });

  it('turns a size estimate into Foundry scene and grid settings', () => {
    // 36.6 m by 27.4 m is 120 ft by 90 ft: 24 by 18 five-foot squares.
    expect(sceneScale({ width_m: 36.6, height_m: 27.4 })).toEqual({ squares: [24, 18], width: 2400, height: 1800, grid: { size: 100, distance: 5, units: 'ft' } });
    expect(sceneScale({ width_m: 36.6, height_m: 27.4 }, { system: 'metric' })).toMatchObject({ squares: [24, 18], grid: { distance: 1.5, units: 'm' } });
    expect(sceneScale({ width_m: 30, height_m: 20 }, { system: 'metric', square: 1 }).squares).toEqual([30, 20]);
    expect(() => sceneScale({ width_m: 0, height_m: 5 })).toThrow();
  });

  it('rejects an unknown perspective', () => {
    expect(() => buildPrompt({ ...prompts[0], perspective: 'sideways' })).toThrow(/Unknown perspective/);
  });

  it('reads PNG and JPEG dimensions', () => {
    const png = new Uint8Array(32);
    png.set([0x89, 0x50, 0x4e, 0x47]);
    new DataView(png.buffer).setUint32(16, 2048);
    new DataView(png.buffer).setUint32(20, 1365);
    expect(imageSize(png)).toEqual({ width: 2048, height: 1365 });

    const jpg = new Uint8Array([0xff, 0xd8, 0xff, 0xc0, 0x00, 0x11, 0x08, 0x05, 0x55, 0x08, 0x00, 0, 0, 0]);
    expect(imageSize(jpg)).toEqual({ width: 2048, height: 1365 });
  });

  it('maps media types to extensions', () => {
    expect(extFor('image/jpeg')).toBe('jpg');
    expect(extFor('image/webp')).toBe('webp');
  });

  it('prices a session as 8 maps and 4 edits', () => {
    expect(sessionCost(0.045, 0.025)).toBeCloseTo(0.46);
  });

  it('has 20 prompts with unique ids, valid perspectives, and edits that point at real prompts', () => {
    expect(prompts).toHaveLength(20);
    expect(new Set(prompts.map(p => p.id)).size).toBe(20);
    for (const p of prompts) expect(() => buildPrompt(p)).not.toThrow();
    for (const e of edits) expect(prompts.some(p => p.id === e.source)).toBe(true);
  });
});

it('sizeBoost asks for a larger area indoors than outdoors', async () => {
  const { sizeBoost, buildPrompt, SIZE_BOOST } = await import('./lib.mjs');
  const inside = { description: 'a two-room tavern', perspective: 'top-down', genre: 'fantasy', style: 'painted', mood: 'day', squares: [24, 18] };
  expect(sizeBoost(inside)).toBe(SIZE_BOOST.indoor);
  expect(sizeBoost({ ...inside, description: 'a walled market town' })).toBe(SIZE_BOOST.outdoor);
  expect(sizeBoost({ ...inside, setting: 'outdoor' })).toBe(SIZE_BOOST.outdoor);
  expect(buildPrompt({ ...inside, sizeBoost: 1 })).toContain('24 by 18 five-foot squares');
  expect(buildPrompt(inside)).toContain('53 by 40 five-foot squares');
});

it('mapDefects retries on grid, text, or frame, and on a margin only for top-down', async () => {
  const { mapDefects } = await import('./lib.mjs');
  expect(mapDefects({ grid: false, text: false, frame: false, margin: 0 })).toEqual({ retry: false, reasons: [] });
  expect(mapDefects({ grid: true, text: true, frame: false, margin: 0 }).reasons).toEqual(['grid', 'text']);
  expect(mapDefects({ margin: 0.2 }).reasons).toEqual(['margin']);
  expect(mapDefects({ margin: 0.2 }, { perspective: 'isometric' }).retry).toBe(false);
});
