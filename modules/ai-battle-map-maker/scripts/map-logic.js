import { ACTION_COSTS } from './core/plans.js';

const OUTDOOR = /\b(town|village|city|street|road|forest|alley|market|camp|wasteland|highway|courtyard|bridge|field|jungle|outpost|main deck|square)\b/i;

const price = id => ACTION_COSTS.find(action => action.id === id).cost;
export const plural = (n, one, many) => n === 1 ? one : many;
const clamp = (value, min, max, fallback) => Number.isFinite(Number(value)) && String(value).trim() !== ''
  ? Math.max(min, Math.min(max, Math.round(Number(value)))) : fallback;

export function costSummary({ quality = 'standard', mockups: mockupCount = 4, includeMockups = true, walls = false, perspective = 'top-down' } = {}) {
  const mockups = includeMockups ? price(mockupCount === 1 ? 'map-mockups-1' : mockupCount === 2 ? 'map-mockups-2' : 'map-mockups') : 0;
  const qualities = { standard: ['map-final', 'Standard'], economy: ['map-final-economy', 'Economy'], draft: ['map-final-draft', 'Draft'], high: ['map-final-high', 'High detail'] };
  const [action, label] = qualities[quality] || qualities.standard;
  const final = price(action);
  const wallCost = walls && perspective === 'top-down' ? price('map-walls') : 0;
  const total = mockups + final + wallCost;
  return { mockups, final, total, text: includeMockups
    ? `${plural(mockupCount, 'Mockup', 'Mockups')} ${mockups} ${plural(mockups, 'use', 'uses')} + ${label} final ${final} ${plural(final, 'use', 'uses')}${wallCost ? ` + walls ${wallCost} use` : ''} = ${total} ${plural(total, 'use', 'uses')}`
    : `${label} final ${final} uses${wallCost ? ` + walls ${wallCost} use` : ''}` };
}

export function buildMapRequest(values) {
  const description = String(values.description ?? '').trim();
  if (!description) throw new Error('Describe the map before generating mockups.');
  const perspective = ['top-down', 'isometric', 'three-quarter'].includes(values.perspective) ? values.perspective : 'top-down';
  const mood = ['day', 'dusk', 'night'].includes(values.mood) ? values.mood : 'day';
  const genre = String(values.genre === 'other' ? values.genreOther : values.genre ?? 'fantasy').trim();
  const style = String(values.style === 'other' ? values.styleOther : values.style ?? 'painted').trim();
  if (!genre || !style) throw new Error('Enter a genre and style for Other.');
  const squares = [clamp(values.width ?? values.squares?.[0], 5, 60, 24), clamp(values.height ?? values.squares?.[1], 5, 60, 18)];
  const setting = ['indoor', 'outdoor'].includes(values.setting) ? values.setting : (OUTDOOR.test(description) ? 'outdoor' : 'indoor');
  const source = values.dials ?? values;
  const dials = Object.fromEntries(['detail', 'clutter', 'wear'].map(key => [key, clamp(source[key], 1, 5, 3)]));
  const layout = String(values.layout ?? '').trim().slice(0, 1200);
  return { description, perspective, genre, style, mood, squares, setting, dials, ...(layout ? { layout } : {}) };
}

export function measuredLabel(scene) {
  const [w, h] = scene.squares;
  const distance = scene.grid.distance;
  return `Measured ${w * distance} by ${h * distance} ${scene.grid.units} (${w} x ${h} squares)`;
}

export function sceneFileName(description, date = new Date(), extension = 'webp') {
  const slug = String(description).toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 50).replace(/-$/g, '') || 'battle-map';
  const stamp = new Date(date).toISOString().replace(/[-:]/g, '').replace(/\.\d{3}Z$/, 'Z').replace('T', '-');
  return `${slug}-${stamp}.${extension}`;
}

export function drawingBox(drawing, scene, dimensions) {
  const shape = drawing?.document?.shape || drawing?.shape;
  const x = drawing?.document?.x ?? drawing?.x;
  const y = drawing?.document?.y ?? drawing?.y;
  if (shape?.type !== 'r' || ![x,y,shape.width,shape.height,scene?.width,scene?.height].every(Number.isFinite) || shape.width <= 0 || shape.height <= 0) return null;
  const sx = dimensions?.sceneX || 0, sy = dimensions?.sceneY || 0;
  const l = Math.max(0, Math.min(1, (x - sx) / scene.width));
  const t = Math.max(0, Math.min(1, (y - sy) / scene.height));
  const r = Math.max(0, Math.min(1, (x + shape.width - sx) / scene.width));
  const b = Math.max(0, Math.min(1, (y + shape.height - sy) / scene.height));
  return r - l >= 0.02 && b - t >= 0.02 ? [l,t,r,b] : null;
}

const lastCreatedDrawingIds = new Map();

export function rememberCreatedDrawing(drawing, userId, currentUserId) {
  const sceneId = drawing?.parent?.id;
  if (userId === currentUserId && sceneId && drawing?.id && drawing?.shape?.type === 'r') {
    lastCreatedDrawingIds.set(sceneId, drawing.id);
  }
}

export function selectRegionDrawing(scene, controlled = [], userId, recentIds = lastCreatedDrawingIds) {
  const rectangles = controlled.filter(drawing => (drawing?.document?.shape || drawing?.shape)?.type === 'r');
  if (rectangles.length === 1) return rectangles[0];
  const drawings = Array.from(scene?.drawings?.contents || scene?.drawings || []);
  const tracked = drawings.find(drawing => drawing.id === recentIds.get(scene?.id) && drawing?.shape?.type === 'r');
  if (tracked) return tracked;
  const owned = drawings.filter(drawing => {
    const author = drawing?.author?.id ?? drawing?.author ?? drawing?.user?.id ?? drawing?.user;
    return drawing?.shape?.type === 'r' && author === userId;
  });
  return owned.sort((a, b) => (Number(b.sort) || 0) - (Number(a.sort) || 0)
    || (Number(b._stats?.createdTime) || 0) - (Number(a._stats?.createdTime) || 0))[0] || null;
}

export function sceneTraceOptions(scene) {
  const size = scene?.grid?.size;
  if (!(size > 0) || !(scene.width > 0) || !(scene.height > 0)) throw new Error('Scene grid is unavailable.');
  return { squares: [Math.round(scene.width / size), Math.round(scene.height / size)], units: scene.grid.units === 'm' ? 'metric' : 'imperial' };
}

/* ── Scene lighting and weather ─────────────────────────────── */

/**
 * Foundry scene environment for a map mood. Global light is on for day and
 * dusk so tokens can see the whole map; without it a walled day map is black
 * away from placed lights. Night relies on the placed lights, and keeps
 * global light only when there are none, so the map is never fully black.
 */
export function sceneEnvironment(mood, lightCount = 0) {
  const lit = { enabled: true, darkness: { min: 0, max: 1 } };
  if (mood === 'night') return { darknessLevel: 0.75, globalLight: lightCount > 0 ? { enabled: false } : lit };
  if (mood === 'dusk') return { darknessLevel: 0.35, globalLight: lit };
  return { darknessLevel: 0, globalLight: lit };
}

/** Foundry's built-in weather effect ids (CONFIG.weatherEffects). */
export const WEATHER = { rain: 'Rain', rainStorm: 'Rain storm', snow: 'Snow', blizzard: 'Blizzard', fog: 'Fog', leaves: 'Autumn leaves' };
const WEATHER_WORDS = [
  ['blizzard', /\b(blizzard|snowstorm|whiteout)\b/i],
  ['rainStorm', /\b(storm|thunder|lightning|downpour|tempest|hurricane|monsoon)\w*\b/i],
  ['snow', /\b(snow|snowy|snowing|snowfall)\b/i],
  ['rain', /\b(rain|rainy|raining|drizzle|drizzling)\b/i],
  ['fog', /\b(fog|foggy|mist|misty|haze|smog)\b/i],
  ['leaves', /\b(autumn|falling leaves)\b/i],
];

/**
 * Weather effect for a scene: an explicit choice, or for "auto" a guess from
 * weather words in the description. Indoor maps never get auto weather.
 * @returns {string} a WEATHER id, or '' for none
 */
export function pickWeather(choice, description = '', setting = 'outdoor') {
  if (choice && choice !== 'auto') return Object.prototype.hasOwnProperty.call(WEATHER, choice) ? choice : '';
  if (setting !== 'outdoor') return '';
  return WEATHER_WORDS.find(([, words]) => words.test(description))?.[0] || '';
}

/* ── Wall and light preview ─────────────────────────────────── */

/**
 * The server's raw trace (0 to 1000 on each axis) as shapes in canvas pixels
 * for the preview overlay.
 */
export function overlayShapes(trace, width, height) {
  const ok = (...values) => values.every(Number.isFinite);
  const sx = v => Math.max(0, Math.min(1000, v)) / 1000 * width;
  const sy = v => Math.max(0, Math.min(1000, v)) / 1000 * height;
  const walls = (Array.isArray(trace?.walls) ? trace.walls : []).filter(w => w && ok(w.x1, w.y1, w.x2, w.y2))
    .map(w => ({ kind: ['door', 'window'].includes(w.kind) ? w.kind : 'wall', x1: sx(w.x1), y1: sy(w.y1), x2: sx(w.x2), y2: sy(w.y2) }));
  const lights = (Array.isArray(trace?.lights) ? trace.lights : []).filter(l => l && ok(l.x, l.y) && l.x >= 0 && l.x <= 1000 && l.y >= 0 && l.y <= 1000)
    .map(l => ({ x: sx(l.x), y: sy(l.y), color: /^#[0-9a-f]{6}$/i.test(l.color || '') ? l.color : '#ffb35c' }));
  return { walls, lights };
}

/* ── Saved maps (Created Maps) ──────────────────────────────── */

const WALL_KINDS = ['wall', 'door', 'window'];

/** Shrink a trace for localStorage: walls as [kind, x1, y1, x2, y2], lights as [kind, x, y, color]. */
export function packTrace(trace) {
  const n = v => Math.round(Number(v));
  return {
    w: (Array.isArray(trace?.walls) ? trace.walls : []).filter(w => w && [w.x1, w.y1, w.x2, w.y2].every(Number.isFinite))
      .map(w => [Math.max(0, WALL_KINDS.indexOf(w.kind)), n(w.x1), n(w.y1), n(w.x2), n(w.y2)]),
    l: (Array.isArray(trace?.lights) ? trace.lights : []).filter(l => l && [l.x, l.y].every(Number.isFinite))
      .map(l => [String(l.kind || 'lamp'), n(l.x), n(l.y), ...(typeof l.color === 'string' ? [l.color] : [])]),
  };
}

export function unpackTrace(packed) {
  return {
    walls: (Array.isArray(packed?.w) ? packed.w : []).map(([kind, x1, y1, x2, y2]) => ({ kind: WALL_KINDS[kind] || 'wall', x1, y1, x2, y2 })),
    lights: (Array.isArray(packed?.l) ? packed.l : []).map(([kind, x, y, color]) => ({ kind, x, y, ...(color ? { color } : {}) })),
  };
}

/** Form fields a saved map or a style preset remembers. */
export const STYLE_FIELDS = ['perspective', 'genre', 'genreOther', 'style', 'styleOther', 'mood', 'quality', 'mockups', 'detail', 'clutter', 'wear', 'color', 'tint', 'contrast', 'sharpness'];
export const MAP_FIELDS = ['description', 'layout', 'width', 'height', 'setting', 'weather', ...STYLE_FIELDS];
const pick = (values, fields) => Object.fromEntries(fields.filter(key => values?.[key] != null && values[key] !== '').map(key => [key, String(values[key])]));

/** The Created Maps entry for a final map saved in the world. */
export function mapRecord({ values, path, final, units, now = Date.now(), id }) {
  const description = String(values?.description ?? '').trim();
  return {
    id, name: description.slice(0, 60) || 'Battle map', description, status: 'success', createdAt: now, error: null,
    map: { path, units, squares: final.scene.squares, trace: packTrace(final.trace), ...(final.warning ? { warning: String(final.warning).slice(0, 300) } : {}),
      values: { ...pick(values, MAP_FIELDS), walls: !!values?.walls } },
  };
}

/** True when a Created Maps entry can be reopened. */
export const isSavedMap = entry => typeof entry?.map?.path === 'string' && Array.isArray(entry.map.squares) && entry.map.squares.length === 2;

/* ── Style presets ──────────────────────────────────────────── */

export const BUILT_IN_PRESETS = [
  { name: 'Classic painted fantasy', values: { perspective: 'top-down', genre: 'fantasy', style: 'painted', mood: 'day', detail: '3', clutter: '3', wear: '3' } },
  { name: 'Inked old-school dungeon', values: { perspective: 'top-down', genre: 'fantasy', style: 'hand-drawn ink', mood: 'day', detail: '2', clutter: '2', wear: '4', color: '2' } },
  { name: 'Grim horror night', values: { perspective: 'top-down', genre: 'horror', style: 'realistic', mood: 'night', detail: '4', clutter: '4', wear: '5', color: '2', contrast: '4' } },
  { name: 'Clean sci-fi deck plan', values: { perspective: 'top-down', genre: 'sci-fi', style: 'realistic', mood: 'day', detail: '4', clutter: '2', wear: '2', tint: '2' } },
  { name: 'Isometric showpiece', values: { perspective: 'isometric', genre: 'fantasy', style: 'painted', mood: 'dusk', detail: '5', clutter: '4', wear: '3', quality: 'high' } },
];

export const presetFromValues = (name, values) => ({ name: String(name).trim().slice(0, 40), values: pick(values, STYLE_FIELDS) });

/** Add or replace a preset by name (case-insensitive), newest last, at most 20. */
export function upsertPreset(presets, preset) {
  if (!preset.name) throw new Error('Give the preset a name.');
  const same = p => p.name.toLowerCase() === preset.name.toLowerCase();
  if (BUILT_IN_PRESETS.some(same)) throw new Error('That name is used by a built-in preset.');
  return [...(Array.isArray(presets) ? presets : []).filter(p => p && typeof p.name === 'string' && !same(p)), preset].slice(-20);
}
