/** Bake-off image and cost helpers, plus the shared map API. */

export * from '../../shared/scripts/core/maps/index.js';
export { imageSize } from '../../shared/scripts/core/maps/image.js';

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
