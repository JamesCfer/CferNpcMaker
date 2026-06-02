import { MODULE_ID } from './constants.js';

const LEVELS = { error: 0, warn: 1, info: 2, debug: 3 };

function getLevel() {
  try { return game.settings?.get(MODULE_ID, 'logLevel') ?? 'info'; } catch (_) { return 'info'; }
}

export function log(level, ...args) {
  if ((LEVELS[level] ?? 0) > (LEVELS[getLevel()] ?? LEVELS.info)) return;
  (console[level] ?? console.log)(`[${MODULE_ID}]`, ...args);
}
