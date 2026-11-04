import { ACTION_COSTS } from './core/plans.js';

const OUTDOOR = /\b(town|village|city|street|road|forest|alley|market|camp|wasteland|highway|courtyard|bridge|field|jungle|outpost|main deck|square)\b/i;

const price = id => ACTION_COSTS.find(action => action.id === id).cost;
const clamp = (value, min, max, fallback) => Number.isFinite(Number(value)) && String(value).trim() !== ''
  ? Math.max(min, Math.min(max, Math.round(Number(value)))) : fallback;

export function costSummary({ quality = 'standard', includeMockups = true, walls = false, perspective = 'top-down' } = {}) {
  const mockups = includeMockups ? price('map-mockups') : 0;
  const final = price(quality === 'economy' ? 'map-final-economy' : 'map-final');
  const wallCost = walls && perspective === 'top-down' ? price('map-walls') : 0;
  const total = mockups + final + wallCost;
  return { mockups, final, total, text: includeMockups
    ? `Mockups ${mockups} uses + ${quality === 'economy' ? 'Economy' : 'Standard'} final ${final} uses${wallCost ? ` + walls ${wallCost} use` : ''} = ${total} uses`
    : `${quality === 'economy' ? 'Economy' : 'Standard'} final ${final} uses${wallCost ? ` + walls ${wallCost} use` : ''}` };
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
  return { description, perspective, genre, style, mood, squares, setting, dials };
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
