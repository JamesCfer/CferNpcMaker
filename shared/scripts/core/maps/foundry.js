/** Foundry light, wall, grid, and map size helpers. */

/** Light kinds the vision model may return, with Foundry light settings in squares. */
export const LIGHT_PRESETS = {
  candle:    { bright: 0.5, dim: 1.5, animation: 'torch',  intensity: 2 },
  lamp:      { bright: 1,   dim: 3,   animation: null,     intensity: 3 },
  torch:     { bright: 2,   dim: 4,   animation: 'torch',  intensity: 4 },
  brazier:   { bright: 2,   dim: 5,   animation: 'flame',  intensity: 4 },
  fireplace: { bright: 3,   dim: 6,   animation: 'flame',  intensity: 5 },
  campfire:  { bright: 4,   dim: 8,   animation: 'flame',  intensity: 5 },
  crystal:   { bright: 1,   dim: 3,   animation: 'pulse',  intensity: 3 },
  magic:     { bright: 2,   dim: 4,   animation: 'energy', intensity: 4 },
  screen:    { bright: 0.5, dim: 2,   animation: null,     intensity: 2 },
  lava:      { bright: 2,   dim: 5,   animation: 'fog',    intensity: 4 },
  window:    { bright: 2,   dim: 5,   animation: null,     intensity: 3 },
};
export const LIGHT_KINDS = Object.keys(LIGHT_PRESETS);

/**
 * Turn the vision model's lights (x and y from 0 to 1000) into Foundry
 * AmbientLight data. Positions are in pixels at 100 pixels per square, and
 * radii are in grid units at 5 feet per square. Unknown kinds become lamps,
 * and points outside the map are dropped.
 * @param {{kind: string, x: number, y: number, color?: string}[]} raw
 * @param {{squares: [number, number], gridPx?: number, feetPerSquare?: number}} opts
 */
export function toFoundryLights(raw, { squares, gridPx = 100, feetPerSquare = 5 }) {
  const [w, h] = squares;
  return raw
    .filter(l => Number.isFinite(l.x) && Number.isFinite(l.y) && l.x >= 0 && l.x <= 1000 && l.y >= 0 && l.y <= 1000)
    .map(l => {
      const kind = LIGHT_PRESETS[l.kind] ? l.kind : 'lamp';
      const p = LIGHT_PRESETS[kind];
      return {
        kind,
        x: Math.round((l.x / 1000) * w * gridPx),
        y: Math.round((l.y / 1000) * h * gridPx),
        config: {
          bright: p.bright * feetPerSquare,
          dim: p.dim * feetPerSquare,
          color: /^#[0-9a-f]{6}$/i.test(l.color || '') ? l.color : '#ffb35c',
          alpha: 0.5,
          animation: p.animation ? { type: p.animation, speed: 3, intensity: p.intensity } : { type: null },
        },
      };
    });
}

/** Foundry wall sense levels (CONST.WALL_SENSE_TYPES) and door types (CONST.WALL_DOOR_TYPES). */
const SENSE = { none: 0, normal: 20 };
const DOOR = { none: 0, door: 1 };

/**
 * Turn the vision model's wall segments (0 to 1000 on each axis) into Foundry
 * Wall data at 100 pixels per square. Windows block movement but not sight or
 * light. Endpoints within a third of a square snap together so corners close,
 * and segments shorter than a fifth of a square are dropped.
 * @param {{kind?: string, x1: number, y1: number, x2: number, y2: number}[]} raw
 * @param {{squares: [number, number], gridPx?: number}} opts
 */
export function toFoundryWalls(raw, { squares, gridPx = 100 }) {
  const [w, h] = squares;
  const clamp = v => Math.min(1000, Math.max(0, v));
  const segs = raw
    .filter(s => ['x1', 'y1', 'x2', 'y2'].every(k => Number.isFinite(s[k])))
    .map(s => ({
      kind: ['door', 'window'].includes(s.kind) ? s.kind : 'wall',
      c: [clamp(s.x1) / 1000 * w * gridPx, clamp(s.y1) / 1000 * h * gridPx, clamp(s.x2) / 1000 * w * gridPx, clamp(s.y2) / 1000 * h * gridPx],
    }))
    .filter(s => Math.hypot(s.c[2] - s.c[0], s.c[3] - s.c[1]) >= gridPx / 5);
  const pts = segs.flatMap(s => [[s, 0], [s, 2]]);
  for (let i = 0; i < pts.length; i++) {
    for (let j = i + 1; j < pts.length; j++) {
      const [s1, k1] = pts[i], [s2, k2] = pts[j];
      if (Math.hypot(s1.c[k1] - s2.c[k2], s1.c[k1 + 1] - s2.c[k2 + 1]) < gridPx / 3) { s2.c[k2] = s1.c[k1]; s2.c[k2 + 1] = s1.c[k1 + 1]; }
    }
  }
  return segs.map(s => ({
    c: s.c.map(Math.round),
    kind: s.kind,
    door: s.kind === 'door' ? DOOR.door : DOOR.none,
    move: SENSE.normal,
    sight: s.kind === 'window' ? SENSE.none : SENSE.normal,
    light: s.kind === 'window' ? SENSE.none : SENSE.normal,
    sound: SENSE.normal,
  }));
}

/**
 * How to fit a generated map to the Foundry grid. The map is resized to
 * exactly 100 pixels per square with the browser's high-quality smoothing,
 * then lightly sharpened (applyLook sharpen 0.6) when it was enlarged by more
 * than 20%. Tested Nov 3, 2026: this looked close to a 4K render, which costs
 * 50% more, so no AI upscaler is used.
 * @param {number} width  generated image width in pixels
 * @param {number} height generated image height in pixels
 * @param {[number, number]} squares map size in squares
 */
export function gridFit(width, height, [sw, sh], gridPx = 100) {
  const scale = (sw * gridPx) / width;
  return { width: sw * gridPx, height: sh * gridPx, scale, sharpen: scale > 1.2 ? 0.6 : 0 };
}

/** Grid square sizes the builder offers. Foundry's grid.distance is the size of one square in these units. */
export const UNIT_SYSTEMS = {
  imperial: { units: 'ft', square: 5 },
  metric:   { units: 'm',  square: 1.5 },
};
const FEET_PER_METRE = 3.28084;

/**
 * Foundry scene settings from the vision model's size estimate. The map is
 * split into whole squares of the chosen size and fitted at 100 pixels per
 * square, so the image stretches slightly to land exactly on the grid.
 * @param {{width_m: number, height_m: number}} estimate real size of the whole image, in metres
 * @param {{system?: 'imperial'|'metric', square?: number, gridPx?: number}} [opts] square overrides the default square size
 */
export function sceneScale({ width_m, height_m }, { system = 'imperial', square, gridPx = 100 } = {}) {
  if (!(width_m > 0 && height_m > 0)) throw new Error('The size estimate needs a positive width and height.');
  const u = UNIT_SYSTEMS[system];
  if (!u) throw new Error(`Unknown unit system "${system}"`);
  const size = square || u.square;
  const toUnits = m => (u.units === 'ft' ? m * FEET_PER_METRE : m);
  const squares = [Math.max(1, Math.round(toUnits(width_m) / size)), Math.max(1, Math.round(toUnits(height_m) / size))];
  return {
    squares,
    width: squares[0] * gridPx,
    height: squares[1] * gridPx,
    grid: { size: gridPx, distance: size, units: u.units },
  };
}

/**
 * Average valid map size estimates in metres. An estimate needs two finite,
 * positive dimensions; an invalid estimate is ignored.
 * @param {{width_m: number, height_m: number}} a
 * @param {{width_m: number, height_m: number}} b
 * @returns {{width_m: number, height_m: number}}
 */
export function averageScale(a, b) {
  const valid = value => Number.isFinite(value?.width_m) && value.width_m > 0 && Number.isFinite(value?.height_m) && value.height_m > 0;
  const estimates = [a, b].filter(valid);
  if (!estimates.length) throw new Error('At least one valid size estimate is required.');
  return {
    width_m: estimates.reduce((sum, value) => sum + value.width_m, 0) / estimates.length,
    height_m: estimates.reduce((sum, value) => sum + value.height_m, 0) / estimates.length,
  };
}

/**
 * Convert a fractional map box into a padded, clamped pixel crop. Matches
 * regions.py padded_box, including Python int() truncation.
 * @param {[number, number, number, number]} box fractional left, top, right, bottom
 * @param {number} width map width in pixels
 * @param {number} height map height in pixels
 * @param {number} [pad=0.15] padding as a share of the box size
 * @returns {[number, number, number, number]} pixel left, top, right, bottom
 */
export function cropBox(box, width, height, pad = 0.15) {
  const [left, top, right, bottom] = [box[0] * width, box[1] * height, box[2] * width, box[3] * height];
  const px = Math.max(32, (right - left) * pad);
  const py = Math.max(32, (bottom - top) * pad);
  return [Math.max(0, left - px), Math.max(0, top - py), Math.min(width, right + px), Math.min(height, bottom + py)].map(Math.trunc);
}
