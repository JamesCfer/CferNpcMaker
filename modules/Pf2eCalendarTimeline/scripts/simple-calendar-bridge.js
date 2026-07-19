/**
 * Optional bridge (#113): when `foundryvtt-simple-calendar` is active, drive
 * our own dayAdvanced/eventFired/seasonChanged/weatherChanged hooks off its
 * clock instead of requiring the GM to also click through our own UI.
 */

import { MODULE_ID } from './state.js';
import { advanceCalendarByDays } from './calendar-app.js';

const SIMPLE_CALENDAR_MODULE_ID = 'foundryvtt-simple-calendar';
const SECONDS_PER_DAY = 86400;

export function registerSimpleCalendarBridge() {
  game.settings.register(MODULE_ID, 'syncWithSimpleCalendar', {
    name: 'Sync with Simple Calendar',
    hint: 'When the Simple Calendar module is active, advance this calendar (and fire its events) whenever Simple Calendar\'s date changes.',
    scope: 'world', config: true, type: Boolean, default: true,
  });

  Hooks.once('ready', () => {
    if (!game.modules?.get(SIMPLE_CALENDAR_MODULE_ID)?.active) return;

    Hooks.on('simple-calendar-date-time-change', ({ diff } = {}) => {
      if (!game.user?.isGM || !game.settings.get(MODULE_ID, 'syncWithSimpleCalendar')) return;
      // Simple Calendar reports elapsed time in seconds; our own calendar
      // only tracks whole days, so round rather than drop sub-day ticks.
      const days = Math.round((diff || 0) / SECONDS_PER_DAY);
      if (days <= 0) return;
      advanceCalendarByDays(days).catch(err => console.warn(`[${MODULE_ID}] Simple Calendar bridge tick failed`, err));
    });

    console.log(`[${MODULE_ID}] Simple Calendar bridge active`);
  });
}
