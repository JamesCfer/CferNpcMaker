/** Pure map prompt builders and options. */
import { LOOK_DIALS } from './dials.js';

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
  ...LOOK_DIALS,
};
export function checkDials(dials) {
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
    ...(typeof p.layout === 'string' && p.layout.trim() ? [p.layout.trim()] : []),
    // Dark water came out as a flat black hole on night maps (Nov 3 crypt test).
    ...(/water|river|lake|pond|pool|flood|sea|stream|canal|moat|swamp|marsh/i.test(p.description) ? ['Draw any water, even water described as dark or black, in deep teal or green-blue with visible ripples, reflections of nearby lights, and the floor or stones showing through where it is shallow. Never draw water as a flat black area.'] : []),
    `The map covers an area of ${Math.round(w * boost)} by ${Math.round(h * boost)} five-foot squares, so a doorway is about one square wide and a person fits in one square.`,
    `Look: ${[...dialPhrases(p.dials), 'crisp textures'].join('; ')}.`,
    'Clear readable layout, the whole area fills the frame edge to edge.',
    // Long layout plans pulled the camera into a tilted view (Nov 4 cave test), so repeat it last.
    ...(p.perspective === 'top-down' && !p.view ? ['Camera: straight down from directly overhead, like a satellite photo. No wall faces, no ceilings, no tilt, no perspective.'] : []),
    `Do not draw any grid lines, text, labels, legend, compass, border, frame, or watermark.${typeof p.layout === 'string' && p.layout.trim() ? ' Do not write any room names or labels on the map.' : ''} No people or creatures.`,
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
export const OUTDOOR = /\b(town|village|city|street|road|forest|alley|market|camp|wasteland|highway|courtyard|bridge|field|jungle|outpost|main deck|square)\b/i;

/** Prompt size multiplier: p.sizeBoost if set, else by p.setting, else guessed from the description. */
export function sizeBoost(p) {
  if (typeof p.sizeBoost === 'number') return p.sizeBoost;
  const setting = p.setting || (OUTDOOR.test(p.description) ? 'outdoor' : 'indoor');
  return SIZE_BOOST[setting] ?? 1;
}

export function cap(s) {
  return s.charAt(0).toUpperCase() + s.slice(1);
}

/** Turn a selected rough mockup into the final image prompt. */
export function finalPrompt(p) {
  return [
    'The attached image is a rough layout mockup. Redraw it as a finished battle map.',
    'Keep the same layout: room shapes, walls, doors, paths, and where every feature sits, and keep the same camera angle.',
    'Ignore and remove any grid lines, text, or labels in the mockup.',
    ...(p.perspective === 'top-down' ? ['If the mockup is tilted or shows the sides of walls, flatten it into a true straight-down view.'] : []),
    buildPrompt(p),
  ].join(' ');
}
