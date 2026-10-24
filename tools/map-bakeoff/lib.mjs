/**
 * Pure helpers for the battle map model bake-off. Kept separate from run.mjs
 * so they can be unit tested without network access.
 */

export const PERSPECTIVES = {
  'top-down':      'strict top-down orthographic view, looking straight down, no horizon, no perspective distortion',
  'isometric':     'isometric view at a 30 degree angle, 2:1 isometric projection, walls and objects shown from above at an angle',
  'three-quarter': 'three-quarter overhead view, camera tilted slightly so the fronts of walls and objects are visible, no horizon',
};

export const MOODS = {
  day:   'bright daylight',
  dusk:  'warm dusk light with long shadows',
  night: 'night time, dark with warm light from lamps, torches, and other light sources',
};

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
 * @param {{description: string, perspective: string, genre: string, style: string, mood: string, squares: [number, number]}} p
 */
export function buildPrompt(p) {
  const view = PERSPECTIVES[p.perspective];
  if (!view) throw new Error(`Unknown perspective "${p.perspective}"`);
  const mood = MOODS[p.mood] || p.mood;
  const [w, h] = p.squares;
  return [
    `Tabletop RPG battle map, ${view}.`,
    `${cap(p.genre)} setting, ${p.style} art style, ${mood}.`,
    `Scene: ${p.description}.`,
    `The map covers an area of ${w} by ${h} five-foot squares, so a doorway is about one square wide and a person fits in one square.`,
    'Highly detailed, crisp textures, clear readable layout, the whole area fills the frame edge to edge.',
    'Do not draw any grid lines, text, labels, legend, compass, border, frame, or watermark. No people or creatures.',
  ].join(' ');
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
