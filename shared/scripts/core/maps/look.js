/** Pixel look controls for battle maps. */

import { LOOK_DIALS, checkLookDials } from './dials.js';

/**
 * Filter amounts for the look dials.
 * @param {Record<string, number>} [dials]
 * @returns {{saturation: number, temperature: number, contrast: number, sharpen: number}}
 */
export function lookFilters(dials = {}) {
  checkLookDials(dials);
  const v = k => LOOK_DIALS[k].levels[(dials[k] ?? 3) - 1];
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
