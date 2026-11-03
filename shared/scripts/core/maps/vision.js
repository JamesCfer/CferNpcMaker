/** Vision prompts and map defect checks. */

import { LIGHT_KINDS } from './foundry.js';

/** Prompt for the map vision model. */
export const LIGHTS_PROMPT = [
  'You are placing light sources on a tabletop RPG battle map for Foundry VTT.',
  'Find every visible light source: candles, lamps, lanterns, torches, braziers, fireplaces, campfires, glowing crystals, magic glows, screens and console lights, lava, and windows or holes letting in daylight or moonlight.',
  'Return JSON only, in this shape:',
  `{"lights":[{"kind":"<one of: ${LIGHT_KINDS.join(', ')}>","x":<0-1000, left to right>,"y":<0-1000, top to bottom>,"color":"#rrggbb"}]}`,
  'x and y are the centre of the light source on the image. Group candles that are touching into one light. Do not add lights where there is no visible source.',
].join('\n');

/** Prompt for the map vision model. */
export const WALLS_PROMPT = [
  'You are tracing walls on a top-down tabletop RPG battle map for Foundry VTT.',
  'List every wall as straight line segments along the centre of the wall. Split walls at doors and windows, and list doors and windows as their own segments.',
  'Return JSON only: {"walls":[{"kind":"wall|door|window","x1":0-1000,"y1":0-1000,"x2":0-1000,"y2":0-1000}]}',
  'x runs left to right and y top to bottom, as a share of the image width and height out of 1000. Do not include furniture.',
].join('\n');

/** Prompt for the map vision model. */
export const WALLS_MASK_PROMPT = [
  'Convert the attached battle map into a wall mask for a virtual tabletop.',
  'Keep every wall exactly where it is in the input, at the same size and position, and keep the same framing.',
  'Fill the whole image with pure black. Paint every wall that blocks movement as solid pure white.',
  'Paint every door as a solid pure red bar in the doorway. Paint every window as a solid pure blue bar.',
  'Draw nothing else: no furniture, floor texture, shading, outlines, text, or grid.',
].join(' ');

/** Prompt for the map vision model. */
export const SCALE_PROMPT = [
  'This is a tabletop RPG battle map. Estimate the real-world size of the whole area shown in the image, edge to edge.',
  'Use objects of known size as rulers: a single door is about 0.9 m wide, a chair about 0.5 m, a bed about 2 m long, a round tavern table about 1.2 m across, a cart about 3 m long, a person about 0.5 m across from above.',
  'Measure at least three different objects, work out metres per pixel from each, and combine them.',
  'Return JSON only: {"references":[{"object":"...","size_m":0.9,"share_of_width":0.03}],"width_m":<number>,"height_m":<number>}',
  'share_of_width is how much of the image width the object spans.',
].join('\n');

/** Prompt for the map vision model. */
export const CHECK_PROMPT = [
  'You are checking an AI-generated tabletop RPG battle map before it goes into Foundry VTT, which adds its own grid.',
  'Answer each question about the image:',
  'grid: are there square or hex grid lines drawn over the map? Floor tiles, paving stones, deck planks, or floor grates that are part of the scenery do not count. Only a regular overlay of thin lines across the map counts, including faint ones.',
  'text: is there any label, number, title, legend, scale bar, dimension, signature, or watermark? Writing that is part of the scene, such as a shop sign or a neon sign, does not count.',
  'frame: is there a drawn decorative border, picture frame, parchment edge, or blueprint sheet around the map? The outer walls of a building or room do not count.',
  'margin: is the map drawn on plain paper, parchment, or a white, grey, or light background, with that blank area showing around it? Dark empty space or rock outside the walls of an interior is normal and does not count. Give the share of the image that is such blank background, from 0 to 1.',
  'Return JSON only: {"grid":true|false,"text":true|false,"frame":true|false,"margin":0.0,"notes":"one short sentence"}',
].join('\n');

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
