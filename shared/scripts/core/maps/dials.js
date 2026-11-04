/** Look controls shipped with the Foundry module. No prompt text belongs here. */
export const LOOK_DIALS = {
  color:     { label: 'Color',             stage: 'filter', levels: [0.35, 0.7, 1, 1.3, 1.6] },
  tint:      { label: 'Color temperature', stage: 'filter', levels: [-1, -0.5, 0, 0.5, 1] },
  contrast:  { label: 'Lighting contrast', stage: 'filter', levels: [0.75, 0.88, 1, 1.15, 1.3] },
  sharpness: { label: 'Sharpness',         stage: 'filter', levels: [-1, -0.5, 0, 0.6, 1.2] },
};

export function checkLookDials(dials = {}) {
  for (const [key, value] of Object.entries(dials)) {
    if (!LOOK_DIALS[key] && !['detail', 'clutter', 'wear'].includes(key)) throw new Error(`Unknown dial "${key}"`);
    if (!Number.isInteger(value) || value < 1 || value > 5) throw new Error(`Dial "${key}" must be a whole number from 1 to 5`);
  }
}
