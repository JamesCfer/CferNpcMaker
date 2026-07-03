/**
 * World-setting backed state for the calendar.
 * Wraps a single 'state' setting containing currentDate, calendarDef, events.
 */

import { DEFAULT_CALENDAR, seasonForMonth, rollWeather } from './scheduler.js';

export const MODULE_ID = 'Pf2eCalendarTimeline';
export const STATE_KEY = 'state';

export function defaultState() {
  const currentDate = { year: 4725, month: 6, day: 1, hour: 8 };
  const season = seasonForMonth(currentDate.month, DEFAULT_CALENDAR);
  return {
    currentDate,
    calendarDef: DEFAULT_CALENDAR,
    events: [],
    biome: 'temperate',
    season,
    weather: rollWeather('temperate', season),
  };
}

export function getState() {
  try { return game.settings?.get(MODULE_ID, STATE_KEY) || defaultState(); }
  catch (_) { return defaultState(); }
}

export async function setState(state) {
  return game.settings.set(MODULE_ID, STATE_KEY, state);
}

export async function patchState(mutator) {
  const cur = foundry.utils.deepClone(getState());
  mutator(cur);
  await setState(cur);
  return cur;
}
