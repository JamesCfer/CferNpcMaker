import { UNIT_SYSTEMS, toFoundryWalls, toFoundryLights } from '../../shared/scripts/core/maps/foundry.js';

export const EDIT_QUALITIES = {
  standard: { model: 'google/gemini-3.1-flash-image', cost: 4 },
  economy: { model: 'google/gemini-3.1-flash-image', cost: 4 },
  draft: { model: 'black-forest-labs/flux.2-klein-4b', cost: 4 },
  high: { model: 'bytedance-seed/seedream-5-0-pro', cost: 6 },
};

export function editSettingsCore(quality, crop, qualities) {
  const chosen = qualities[quality];
  if (!chosen) throw new Error('Unknown edit quality.');
  return { ...chosen, resolution: quality === 'high' || Math.max(crop.width, crop.height) > 1100 ? '2K' : '1K' };
}
export const editSettings = (quality, crop) => editSettingsCore(quality, crop, EDIT_QUALITIES);

export function convertTraceCore(trace, squares, units, systems, convertWalls, convertLights) {
  const unit = systems[units];
  const walls = convertWalls(trace.walls, { squares, gridPx: 100 }).map(({ c, door, move, sight, light, sound }) => ({ c, door, move, sight, light, sound }));
  const lights = convertLights(trace.lights, { squares, gridPx: 100, feetPerSquare: unit.square }).map(({ x, y, config }) => ({ x, y, config }));
  const counts = { walls: walls.filter(w => w.door === 0 && w.sight !== 0).length, doors: walls.filter(w => w.door === 1).length, windows: walls.filter(w => w.sight === 0).length, lights: lights.length };
  return { walls, lights, counts };
}
export const convertTrace = (trace, squares, units) => convertTraceCore(trace, squares, units, UNIT_SYSTEMS, toFoundryWalls, toFoundryLights);
