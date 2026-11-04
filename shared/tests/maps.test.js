import { describe, it, expect } from 'vitest';
import { sceneScale, gridFit, toFoundryWalls, toFoundryLights, pickAspect, buildPrompt, dialPhrases, DIALS, applyLook, lookFilters, sizeBoost, SIZE_BOOST, mapDefects, averageScale, cropBox, LIGHTS_PROMPT, WALLS_PROMPT, WALLS_MASK_PROMPT, SCALE_PROMPT, CHECK_PROMPT, PLAN_PROMPT, parsePlan, planText } from '../scripts/core/maps/index.js';

const prompts = [
  {
    "id": "p01-crypt",
    "perspective": "top-down",
    "genre": "dark fantasy",
    "style": "painted",
    "mood": "night",
    "squares": [
      30,
      20
    ],
    "description": "a flooded crypt with stone sarcophagi, a collapsed section of ceiling letting in moonlight, and shallow black water across the floor"
  },
  {
    "id": "p05-spaceship",
    "perspective": "top-down",
    "genre": "sci-fi",
    "style": "realistic",
    "mood": "night",
    "squares": [
      30,
      20
    ],
    "description": "the interior deck of a small freighter: cockpit at the front, a cargo bay with crates, crew bunks, an engine room with glowing reactor"
  }
];

describe('shared map helpers', () => {
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
  it('inserts layout immediately after the scene and preserves old prompts without it', () => {
    const base = buildPrompt(prompts[0]);
    const laid = buildPrompt({ ...prompts[0], layout: 'Layout: fire in the centre.' });
    expect(laid).toContain(`Scene: ${prompts[0].description}. Layout: fire in the centre.`);
    expect(laid).toContain('Do not write any room names or labels on the map.');
    expect(base).not.toContain('Do not write any room names or labels on the map.');
  });

  it('parses a fenced plan, keeps described areas, and caps layout prose', () => {
    const area = (what, where) => ({ what, where, size: 'small' });
    const plan = parsePlan('```json\n' + JSON.stringify({ areas: [area('broken portcullis', 'south'), area('fire in the centre, ash and rubble', 'centre'), area('stone tower', 'north'), area('stable stalls', 'west'), area('invalid', 'up'), area('rubble', 'south')], connections: ['a passage through a broken portcullis', 5], notes: 'An uneven ruined keep.' }) + '\n```');
    expect(plan.areas.map(a => a.what)).toEqual(['broken portcullis', 'fire in the centre, ash and rubble', 'stone tower', 'stable stalls', 'invalid', 'rubble']);
    expect(plan.areas[4].where).toBe('up');
    expect(plan.connections).toEqual(['a passage through a broken portcullis']);
    expect(parsePlan(JSON.stringify({ areas: plan.areas.slice(0, 2) })).areas).toHaveLength(2);
    expect(parsePlan(JSON.stringify({ areas: [area('kitchen', 'north'), area('', 'up')] }))).toMatchObject({ areas: [area('kitchen', 'north')], rawAreas: 2, warning: 'The layout plan was thin; check it before generating.' });
    expect(plan).not.toHaveProperty('warning');
    expect(planText(plan, 'fire in the centre.')).toContain('GM request: fire in the centre. Layout:');
    expect(planText(plan, 'fire in the centre.')).not.toContain('centre..');
    expect(planText(plan, 'fire in the centre')).toContain('fire in the centre, ash and rubble');
    expect(planText(plan).endsWith('Do not write any room names or labels on the map.')).toBe(true);
    const long = planText({ areas: Array.from({ length: 8 }, (_, i) => area('long ruined chamber, broken stone, ash, rubble, old banners ' + i, 'centre')), connections: Array(12).fill('long corridor connecting two chambers'), notes: 'A huge keep.' }, 'fire in the centre');
    expect(long.split(/\s+/).length).toBeLessThanOrEqual(180);
  });

  it('plans natural scenes as one continuous area and keeps dominant features prominent', () => {
    expect(PLAN_PROMPT).toContain('ONE continuous space');
    expect(PLAN_PROMPT).toContain('exactly one area');
    expect(PLAN_PROMPT).toContain('Rewrite every request as polished visual description');
    expect(PLAN_PROMPT).toContain('never paste GM wording');
    expect(PLAN_PROMPT).toContain('A double door is ONE doorway with two leaves');
    expect(PLAN_PROMPT).toContain('mention each once and early');
    expect(PLAN_PROMPT).toContain('jagged natural rock edges and boulders');
    expect(PLAN_PROMPT).toContain('never say walls, brick, masonry, or blocks unless the GM specifies built, carved, masonry, or brick construction');
    expect(PLAN_PROMPT).toContain('ring of wedge-shaped steps around a central column');
    expect(PLAN_PROMPT).toContain('four triangular faces meeting at a central point');
    expect(PLAN_PROMPT).toContain('on a wider square base with a long shadow on the floor');
    expect(PLAN_PROMPT).toContain('Describe tall things only by their top outline and shadow on the floor, never their height or sides');
    expect(PLAN_PROMPT).toContain('two thick wooden leaves meeting in the middle, lying in a wide gap in the north rock edge');
    expect(PLAN_PROMPT).toContain('Doors and gates are thin strips in the wall line seen from above');
    expect(PLAN_PROMPT).toContain('one requested staircase means exactly one');
    expect(PLAN_PROMPT).toContain('Ban ceilings, stalactites, hanging things, wall faces, front-view arches');
    const description = 'A cave where the walls are made of stone, there is a oblisk in the centre, to the north there is a huge double door that is shut, to the south there is a spiral stone staircase that is also huge.';
    const plan = parsePlan(JSON.stringify({ areas: [
      { what: 'a stone cave with a oblisk in the centre', where: 'centre' },
      { what: 'a vaulted stone threshold where to the north there is a huge double door that is shut', where: 'north' },
      { what: 'a sunken cavern tier where to the south there is a spiral stone staircase that is also huge', where: 'south' },
    ], connections: ['a passage to the threshold', 'stairs down to the tier'], notes: 'Split floor levels.' }));
    const layout = planText(plan, description);
    expect(layout).toMatch(/^GM request: /);
    expect(layout).toContain('GM request: ' + description);
    expect(layout).toContain('obelisk');
    expect(layout).toContain('double door');
    expect(layout).toContain('spiral stone staircase');
    expect(layout.slice(layout.indexOf('Layout:'))).not.toMatch(/threshold|tier|Connections:|\d/);
    expect(layout).toContain('A single continuous space');
    expect((layout.match(/ at centre\b/g) || [])).toHaveLength(1);
    expect(layout).not.toMatch(/\d/);
  });

  it('strips measurements from the layout plan', () => {
    const plan = parsePlan(JSON.stringify({ areas: [{ what: 'a hall 30 ft long with pillars', where: 'north', size: 'about 5 by 4 squares' }, { what: 'two braziers and 3 crates', where: 'centre', size: 'the largest space (roughly 8x6 squares)' }], connections: ['a 10-foot corridor'], notes: 'A keep 60 by 40 feet.' }));
    expect(plan.areas[0]).toEqual({ what: 'a hall with pillars', where: 'north' });
    expect(plan.areas[1]).toEqual({ what: 'two braziers and crates', where: 'centre', size: 'the largest space' });
    expect(plan.connections).toEqual(['a corridor']);
    expect(planText(plan)).not.toMatch(/\d+\s*(?:ft|feet|squares?)|\d+\s*(?:x|by)\s*\d+/i);
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

});

it('sizeBoost asks for a larger area indoors than outdoors', async () => {
  const inside = { description: 'a two-room tavern', perspective: 'top-down', genre: 'fantasy', style: 'painted', mood: 'day', squares: [24, 18] };
  expect(sizeBoost(inside)).toBe(SIZE_BOOST.indoor);
  expect(sizeBoost({ ...inside, description: 'a walled market town' })).toBe(SIZE_BOOST.outdoor);
  expect(sizeBoost({ ...inside, setting: 'outdoor' })).toBe(SIZE_BOOST.outdoor);
  expect(buildPrompt({ ...inside, sizeBoost: 1 })).toContain('24 by 18 five-foot squares');
  expect(buildPrompt(inside)).toContain('53 by 40 five-foot squares');
});

it('mapDefects retries on grid, text, or frame, and on a margin only for top-down', async () => {
  expect(mapDefects({ grid: false, text: false, frame: false, margin: 0 })).toEqual({ retry: false, reasons: [] });
  expect(mapDefects({ grid: true, text: true, frame: false, margin: 0 }).reasons).toEqual(['grid', 'text']);
  expect(mapDefects({ margin: 0.2 }).reasons).toEqual(['margin']);
  expect(mapDefects({ margin: 0.2 }, { perspective: 'isometric' }).retry).toBe(false);
});


describe('new shared map helpers', () => {
  it('averages two valid size estimates', () => {
    expect(averageScale({ width_m: 10, height_m: 20 }, { width_m: 30, height_m: 40 })).toEqual({ width_m: 20, height_m: 30 });
  });

  it('uses the one valid estimate and drops extra fields', () => {
    expect(averageScale({ width_m: 10, height_m: 20, note: 'first' }, { width_m: Infinity, height_m: 40 })).toEqual({ width_m: 10, height_m: 20 });
    expect(averageScale({ width_m: 0, height_m: 20 }, { width_m: 30, height_m: 40, note: 'second' })).toEqual({ width_m: 30, height_m: 40 });
    expect(averageScale({ width_m: 10, height_m: 20, note: 'first' }, { width_m: 30, height_m: 40, note: 'second' })).toEqual({ width_m: 20, height_m: 30 });
  });

  it('rejects two invalid estimates', () => {
    expect(() => averageScale({ width_m: NaN, height_m: 20 }, { width_m: 10, height_m: -1 })).toThrow();
  });

  it('pads a centre box and truncates fractional pixel coordinates', () => {
    // Python padded_box([.25, .25, .75, .75], 1000, 800): box 500x400;
    // margins max(32, 75)=75 and max(32, 60)=60, so [175, 140, 825, 660].
    expect(cropBox([0.25, 0.25, 0.75, 0.75], 1000, 800)).toEqual([175, 140, 825, 660]);
    const result = cropBox([0.123, 0.234, 0.567, 0.678], 333, 277);
    expect(result.every(Number.isInteger)).toBe(true);
  });

  it('uses a minimum 32 pixel margin and clamps to map edges', () => {
    expect(cropBox([0.45, 0.45, 0.55, 0.55], 200, 200)).toEqual([58, 58, 142, 142]);
    expect(cropBox([0, 0, 1, 1], 100, 80)).toEqual([0, 0, 100, 80]);
  });

  it('exports nonempty vision prompts with the expected wall JSON instructions', () => {
    for (const prompt of [LIGHTS_PROMPT, WALLS_PROMPT, WALLS_MASK_PROMPT, SCALE_PROMPT, CHECK_PROMPT]) {
      expect(typeof prompt).toBe('string');
      expect(prompt.length).toBeGreaterThan(0);
    }
    expect(WALLS_PROMPT).toContain('JSON');
    expect(WALLS_PROMPT).toContain('walls');
  });
});
