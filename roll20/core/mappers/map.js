/**
 * AI Battle Map Maker on Roll20.
 *
 * Generation is unchanged (the same /webhook/map-builder actions: plan,
 * mockups, final). This module turns a finished map into a plan: the image the
 * GM drops on the page, the page size, and walls, doors, windows and lights in
 * Roll20 pixels (70 per square).
 */
import { makePlan } from '../plan.js';
import { toFoundryWalls, toFoundryLights, LIGHT_PRESETS } from '../vendor/maps-foundry.js';

export const ROLL20_GRID = 70;
const OUTDOOR = /\b(town|village|city|street|road|forest|alley|market|camp|wasteland|highway|courtyard|bridge|field|jungle|outpost|main deck|square)\b/i;
const clamp = (v, lo, hi, d) => (Number.isFinite(Number(v)) && v !== '' ? Math.min(hi, Math.max(lo, Math.round(Number(v)))) : d);

/** Same request shape the Foundry module sends (see map-logic.js buildMapRequest). */
export function buildMapRequest(values) {
  const description = String(values.description ?? '').trim();
  if (!description) throw new Error('Describe the map before generating mockups.');
  const perspective = ['top-down', 'isometric', 'three-quarter'].includes(values.perspective) ? values.perspective : 'top-down';
  const mood = ['day', 'dusk', 'night'].includes(values.mood) ? values.mood : 'day';
  const genre = String(values.genre === 'other' ? values.genreOther : values.genre ?? 'fantasy').trim() || 'fantasy';
  const style = String(values.style === 'other' ? values.styleOther : values.style ?? 'painted').trim() || 'painted';
  const squares = [clamp(values.width, 5, 60, 24), clamp(values.height, 5, 60, 18)];
  const setting = ['indoor', 'outdoor'].includes(values.setting) ? values.setting : (OUTDOOR.test(description) ? 'outdoor' : 'indoor');
  const dials = Object.fromEntries(['detail', 'clutter', 'wear'].map(k => [k, clamp(values[k], 1, 5, 3)]));
  const layout = String(values.layout ?? '').trim().slice(0, 1200);
  return { description, perspective, genre, style, mood, squares, setting, dials, ...(layout ? { layout } : {}) };
}

/**
 * @param {object} final  The /map-builder `final` response: { image:{b64,mediaType}, scene:{squares,grid}, trace:{walls,lights} }
 * @param {{name?:string, units?:'imperial'|'metric'}} opts
 */
export function mapFinal(final, { name = 'Battle map', walls: wantWalls = true } = {}) {
  if (!final?.image?.b64 || !final?.scene?.squares) throw new Error('The final map response was incomplete.');
  const squares = final.scene.squares;
  const distance = Number(final.scene.grid?.distance) || 5;
  const units = final.scene.grid?.units || 'ft';
  const trace = final.trace || {};
  const walls = wantWalls ? toFoundryWalls(Array.isArray(trace.walls) ? trace.walls : [], { squares, gridPx: ROLL20_GRID }) : [];
  const lights = wantWalls ? toFoundryLights(Array.isArray(trace.lights) ? trace.lights : [], { squares, gridPx: ROLL20_GRID, feetPerSquare: distance }) : [];
  return makePlan('map', 'battle-map', {
    summary: `${name} (${squares[0]} x ${squares[1]} squares, ${walls.length} walls, ${lights.length} lights)`,
    map: {
      name,
      squares,
      grid: { distance, units },
      image: { mediaType: final.image.mediaType || 'image/png', filename: `${name.replace(/[^\w-]+/g, '-').toLowerCase() || 'battle-map'}.${(final.image.mediaType || 'image/png').split('/')[1] || 'png'}` },
      walls: walls.map(w => ({ kind: w.kind, c: w.c })),
      lights: lights.map(l => ({ kind: l.kind, x: l.x, y: l.y, bright: l.config.bright, dim: l.config.dim, color: l.config.color })),
      darkness: final.scene.mood === 'night' || final.mood === 'night',
    },
  });
}

export { LIGHT_PRESETS };
