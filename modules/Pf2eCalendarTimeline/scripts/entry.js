/**
 * Pf2eCalendarTimeline — entry point.
 * Registers the world setting, sidebar button, and opens the calendar UI.
 */

import { registerSidebar }      from './core/sidebar.js';
import { startHeartbeat }       from './core/heartbeat.js';
import { MODULE_ID, STATE_KEY, defaultState } from './state.js';
import { openCalendar }         from './calendar-app.js';
import { CalendarAdapter }      from './adapter.js';

const adapter = new CalendarAdapter();

registerSidebar(MODULE_ID, () => openCalendar(), {
  buttonLabel: 'Calendar',
  buttonIcon:  adapter.module.icon,
  directories: ['journal'],
});

Hooks.once('init', () => {
  game.settings.register(MODULE_ID, STATE_KEY, {
    scope: 'world', config: false, type: Object, default: defaultState(),
  });
  game.settings.register(MODULE_ID, 'welcomeMessageShown', {
    scope: 'world', config: false, type: Boolean, default: false,
  });
});

Hooks.once('ready', () => {
  foundry.applications.handlebars.loadTemplates([
    `modules/${MODULE_ID}/templates/calendar.hbs`,
  ]).catch(err => console.warn(`[${MODULE_ID}] template preload`, err));

  console.log(`Pf2eCalendarTimeline ready.`);
  startHeartbeat(MODULE_ID);

  if (game.user.isGM && !game.settings.get(MODULE_ID, 'welcomeMessageShown')) {
    ChatMessage.create({
      content: `<h3>PF2e Calendar & Timeline installed</h3>
        <p>Open it from the <strong>Journal</strong> sidebar header. Schedule taxes, paydays and festivals — other modules can listen for fired events.</p>`,
      whisper: game.users.filter(u => u.isGM).map(u => u.id),
    });
    game.settings.set(MODULE_ID, 'welcomeMessageShown', true);
  }
});
