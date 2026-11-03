/**
 * Pure helpers for the battle map model bake-off. Kept separate from run.mjs
 * so they can be unit tested without network access.
 */

export const PERSPECTIVES = {
  // Chosen Nov 3, 2026 after testing four wordings ("combo" won: flattest walls, no labels, no white borders).
  'top-down':      'flat top-down orthographic floor plan seen from directly overhead, like a satellite view from far above. Only the top surfaces of walls, furniture, and objects are visible, and nothing has visible height or sides. Walls are flat strips, never showing their inner or outer faces. Tables read as flat circles and rectangles. Shadows are short and fall straight down',
  'isometric':     'isometric view at a 30 degree angle, 2:1 isometric projection, walls and objects shown from above at an angle',
  'three-quarter': 'three-quarter overhead view, camera tilted slightly so the fronts of walls and objects are visible, no horizon',
};

/** Other top-down wordings tested on Nov 3, 2026. "plan" drew dimension labels and white borders, "sat" tilted the crypt. Pass one as p.view to use it. */
export const TOP_DOWN_VARIANTS = {
  plan:  'orthographic floor plan seen from directly overhead, like an architect\'s plan drawn in full color. Only the top surfaces of walls, furniture, and objects are visible. Walls are flat bands with no visible height or faces. Shadows are short and fall straight down',
  sat:   'satellite view from far above, looking straight down, so nothing has visible height or sides. Tables read as flat circles and rectangles, chairs as small squares, and walls as flat strips',
  previous: 'flat top-down orthographic view with the camera pointing straight down at 90 degrees, like a classic virtual tabletop battle map. Walls are flat thick outlines seen only from above, with no wall faces or sides of furniture visible. No tilt, no depth, no horizon, no vanishing point',
  combo: 'orthographic floor plan seen from directly overhead, like a satellite view from far above. Only the top surfaces of walls, furniture, and objects are visible, and nothing has visible height or sides. Walls are flat strips, never showing their inner or outer faces. Tables read as flat circles and rectangles. Shadows are short and fall straight down',
};

export const MOODS = {
  day:   'bright daylight',
  dusk:  'warm dusk light with long shadows',
  night: 'night time, dark with warm light from lamps, torches, and other light sources',
};

/**
 * Dials the GM can set from 1 to 5. Level 3 is the default.
 *
 * Content dials change what is drawn, so they go into the mockup prompt (the
 * final copies the mockup's content). Look dials are image filters applied
 * after the map is made: free, exact, and changeable later without new uses.
 * A null level adds nothing to the prompt.
 */
export const DIALS = {
  detail:    { label: 'Detail',            stage: 'mockup', levels: ['simple and clean, with large plain areas and only the key features', 'lightly detailed, with simple textures', 'highly detailed', 'very highly detailed, with rich textures on every surface', 'extremely intricate, with fine texture and small details everywhere'] },
  clutter:   { label: 'Props',             stage: 'mockup', levels: ['almost empty, with only essential furniture', 'sparse props', null, 'plenty of props and furnishings', 'crowded with props, debris, and small objects, while keeping clear floor space to walk on'] },
  wear:      { label: 'Wear',              stage: 'mockup', levels: ['everything pristine and brand new', 'everything clean and well kept', null, 'surfaces worn, dirty, and cracked', 'ruined, broken, and overgrown surroundings'] },
  color:     { label: 'Color',             stage: 'filter', levels: [0.35, 0.7, 1, 1.3, 1.6] },   // saturation
  tint:      { label: 'Color temperature', stage: 'filter', levels: [-1, -0.5, 0, 0.5, 1] },      // cool to warm
  contrast:  { label: 'Lighting contrast', stage: 'filter', levels: [0.75, 0.88, 1, 1.15, 1.3] },
  sharpness: { label: 'Sharpness',         stage: 'filter', levels: [-1, -0.5, 0, 0.6, 1.2] },    // blur to sharpen
};

function checkDials(dials) {
  for (const [k, v] of Object.entries(dials)) {
    if (!DIALS[k]) throw new Error(`Unknown dial "${k}"`);
    if (!Number.isInteger(v) || v < 1 || v > 5) throw new Error(`Dial "${k}" must be a whole number from 1 to 5`);
  }
}

/**
 * Prompt phrases for the content dials, in a fixed order. Look dials are ignored.
 * @param {Record<string, number>} [dials] e.g. { detail: 4, wear: 5 }
 * @returns {string[]}
 */
export function dialPhrases(dials = {}) {
  checkDials(dials);
  return Object.entries(DIALS).filter(([, d]) => d.stage === 'mockup').map(([k, d]) => d.levels[(dials[k] ?? 3) - 1]).filter(Boolean);
}

/**
 * Filter amounts for the look dials.
 * @param {Record<string, number>} [dials]
 * @returns {{saturation: number, temperature: number, contrast: number, sharpen: number}}
 */
export function lookFilters(dials = {}) {
  checkDials(dials);
  const v = k => DIALS[k].levels[(dials[k] ?? 3) - 1];
  return { saturation: v('color'), temperature: v('tint'), contrast: v('contrast'), sharpen: v('sharpness') };
}

/**
 * Apply the look dials to RGBA pixels in place. Works on canvas ImageData in
 * the browser (Foundry) and on plain arrays in Node, so the module bakes the
 * same result the preview shows. With every look dial on 3 nothing changes.
 * @param {Uint8ClampedArray} px RGBA pixels
 * @param {number} width
 * @param {number} height
 * @param {Record<string, number>} [dials]
 */
export function applyLook(px, width, height, dials = {}) {
  const { saturation, temperature, contrast, sharpen } = lookFilters(dials);
  if (sharpen) {
    // 3x3 box blur, then push away from it (sharpen) or blend toward it (soften).
    const src = Uint8ClampedArray.from(px);
    for (let y = 0; y < height; y++) {
      for (let x = 0; x < width; x++) {
        const i = (y * width + x) * 4;
        for (let c = 0; c < 3; c++) {
          let sum = 0, n = 0;
          for (let dy = -1; dy <= 1; dy++) {
            const yy = y + dy; if (yy < 0 || yy >= height) continue;
            for (let dx = -1; dx <= 1; dx++) {
              const xx = x + dx; if (xx < 0 || xx >= width) continue;
              sum += src[(yy * width + xx) * 4 + c]; n++;
            }
          }
          const blur = sum / n;
          px[i + c] = sharpen > 0 ? src[i + c] + (src[i + c] - blur) * sharpen : src[i + c] + (blur - src[i + c]) * -sharpen;
        }
      }
    }
  }
  if (saturation === 1 && temperature === 0 && contrast === 1) return;
  for (let i = 0; i < px.length; i += 4) {
    let r = px[i], g = px[i + 1], b = px[i + 2];
    r = (r - 128) * contrast + 128; g = (g - 128) * contrast + 128; b = (b - 128) * contrast + 128;
    const l = 0.299 * r + 0.587 * g + 0.114 * b;
    r = l + (r - l) * saturation; g = l + (g - l) * saturation; b = l + (b - l) * saturation;
    r += 18 * temperature; g += 4 * temperature; b -= 18 * temperature;
    px[i] = r; px[i + 1] = g; px[i + 2] = b;
  }
}

/** Aspect ratios most image providers accept. */
export const ASPECTS = [
  ['1:1', 1], ['4:3', 4 / 3], ['3:4', 3 / 4], ['3:2', 3 / 2], ['2:3', 2 / 3], ['16:9', 16 / 9], ['9:16', 9 / 16],
];

/**
 * Closest supported aspect ratio for a map of the given size in squares.
 * @param {[number, number]} squares
 */
export function pickAspect([w, h]) {
  const target = w / h;
  let best = ASPECTS[0];
  for (const a of ASPECTS) {
    if (Math.abs(Math.log(a[1] / target)) < Math.abs(Math.log(best[1] / target))) best = a;
  }
  return best[0];
}

/**
 * The full text prompt sent to every model for one test map.
 * @param {{description: string, perspective: string, genre: string, style: string, mood: string, squares: [number, number], dials?: Record<string, number>}} p
 */
export function buildPrompt(p) {
  const view = p.view ? TOP_DOWN_VARIANTS[p.view] : PERSPECTIVES[p.perspective];
  if (!view) throw new Error(`Unknown perspective "${p.perspective}"`);
  const mood = MOODS[p.mood] || p.mood;
  const [w, h] = p.squares;
  const boost = sizeBoost(p);
  return [
    `Tabletop RPG battle map, ${view}.`,
    `${cap(p.genre)} setting, ${p.style} art style, ${mood}.`,
    `Scene: ${p.description}.`,
    // Dark water came out as a flat black hole on night maps (Nov 3 crypt test).
    ...(/water|river|lake|pond|pool|flood|sea|stream|canal|moat|swamp|marsh/i.test(p.description) ? ['Draw any water, even water described as dark or black, in deep teal or green-blue with visible ripples, reflections of nearby lights, and the floor or stones showing through where it is shallow. Never draw water as a flat black area.'] : []),
    `The map covers an area of ${Math.round(w * boost)} by ${Math.round(h * boost)} five-foot squares, so a doorway is about one square wide and a person fits in one square.`,
    `Look: ${[...dialPhrases(p.dials), 'crisp textures'].join('; ')}.`,
    'Clear readable layout, the whole area fills the frame edge to edge.',
    'Do not draw any grid lines, text, labels, legend, compass, border, frame, or watermark. No people or creatures.',
  ].join(' ');
}

/**
 * Image models draw furniture and doors too big for the area we ask for, so
 * the map comes out smaller than requested. Measured Nov 3, 2026: indoor maps
 * came out at about 0.45 of the asked width, outdoor (town) at about 0.88.
 * The boost helped the tavern and freighter a little; listing object sizes
 * in squares did not help and was dropped.
 * The prompt asks for a larger area by the inverse of that.
 */
export const SIZE_BOOST = { indoor: 2.2, outdoor: 1.15 };
const OUTDOOR = /\b(town|village|city|street|road|forest|alley|market|camp|wasteland|highway|courtyard|bridge|field|jungle|outpost|main deck|square)\b/i;

/** Prompt size multiplier: p.sizeBoost if set, else by p.setting, else guessed from the description. */
export function sizeBoost(p) {
  if (typeof p.sizeBoost === 'number') return p.sizeBoost;
  const setting = p.setting || (OUTDOOR.test(p.description) ? 'outdoor' : 'indoor');
  return SIZE_BOOST[setting] ?? 1;
}

function cap(s) {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

/**
 * Width and height of a PNG, JPEG, or WebP image from its bytes.
 * @param {Uint8Array} b
 * @returns {{width: number, height: number} | null}
 */
export function imageSize(b) {
  if (b.length > 24 && b[0] === 0x89 && b[1] === 0x50) {
    const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
    return { width: v.getUint32(16), height: v.getUint32(20) };
  }
  if (b.length > 4 && b[0] === 0xff && b[1] === 0xd8) {
    let i = 2;
    while (i + 9 < b.length) {
      if (b[i] !== 0xff) { i++; continue; }
      const marker = b[i + 1];
      const len = (b[i + 2] << 8) | b[i + 3];
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { height: (b[i + 5] << 8) | b[i + 6], width: (b[i + 7] << 8) | b[i + 8] };
      }
      i += 2 + len;
    }
    return null;
  }
  if (b.length > 30 && b[0] === 0x52 && b[8] === 0x57) {
    const tag = String.fromCharCode(b[12], b[13], b[14], b[15]);
    if (tag === 'VP8X') return { width: 1 + (b[24] | b[25] << 8 | b[26] << 16), height: 1 + (b[27] | b[28] << 8 | b[29] << 16) };
    if (tag === 'VP8 ') return { width: (b[26] | b[27] << 8) & 0x3fff, height: (b[28] | b[29] << 8) & 0x3fff };
    if (tag === 'VP8L') {
      const bits = b[21] | b[22] << 8 | b[23] << 16 | b[24] << 24;
      return { width: (bits & 0x3fff) + 1, height: ((bits >> 14) & 0x3fff) + 1 };
    }
  }
  return null;
}

/** File extension for an image media type. */
export function extFor(mediaType) {
  return { 'image/png': 'png', 'image/jpeg': 'jpg', 'image/webp': 'webp' }[mediaType] || 'img';
}

/**
 * Model cost of one typical session (8 new maps and 4 edits) from measured averages.
 * @param {number} mapUsd  average cost of one new map
 * @param {number} editUsd average cost of one edit
 */
export function sessionCost(mapUsd, editUsd) {
  return 8 * mapUsd + 4 * editUsd;
}

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
 * Turns the vision check's answers into a retry decision. A drawn grid, text,
 * or frame always means a retry. A blank margin only counts on top-down maps,
 * since isometric and three-quarter maps sit on a background by design.
 * @param {{grid?: boolean, text?: boolean, frame?: boolean, margin?: number}} raw
 * @param {{perspective?: string, maxMargin?: number}} [opts]
 */
export function mapDefects(raw, { perspective = 'top-down', maxMargin = 0.08 } = {}) {
  const reasons = ['grid', 'text', 'frame'].filter(k => raw[k] === true);
  if (perspective === 'top-down' && Number(raw.margin) > maxMargin) reasons.push('margin');
  return { retry: reasons.length > 0, reasons };
}
