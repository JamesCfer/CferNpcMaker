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
