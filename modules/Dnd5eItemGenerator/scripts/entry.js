/**
 * D&D 5e Item Generator — module entry point.
 * Registers in the Items directory (not Actors) since output goes to game.items.
 */

import { openBuilder }          from './core/app.js';
import { checkForModuleUpdate } from './core/update-check.js';
import { registerSidebar }      from './core/sidebar.js';
import { startHeartbeat }       from './core/heartbeat.js';
import { Dnd5eItemAdapter }     from './adapter.js';

const adapter   = new Dnd5eItemAdapter();
const MODULE_ID = adapter.module.id;

const openFn = () => {
  openBuilder(adapter);
  checkForModuleUpdate(MODULE_ID, adapter.module.githubUrl).catch(() => {});
};

registerSidebar(MODULE_ID, openFn, {
  buttonLabel: 'Item Generator',
  buttonIcon:  '✦',
  directories: ['items', 'compendium'],
});

Hooks.once('init', () => {
  game.settings.register(MODULE_ID, 'devMode', {
    name:   'Developer Mode',
    hint:   'When enabled, all webhook URLs are routed to the -dev endpoints. Disable before going live.',
    scope:  'world', config: true, type: Boolean, default: false,
  });

  game.settings.register(MODULE_ID, 'welcomeMessageShown', {
    scope:  'world', config: false, type: Boolean, default: false,
  });
});

Hooks.once('ready', () => {
  const mod = game.modules?.get(MODULE_ID);
  const currentVersion = mod?.version || '';

  const storedVersionKey = `${MODULE_ID}.module-version`;
  let storedVersion = '';
  try { storedVersion = localStorage.getItem(storedVersionKey) || ''; } catch (_) {}

  if (currentVersion && storedVersion && currentVersion !== storedVersion) {
    try {
      localStorage.removeItem(`${MODULE_ID}.key`);
      localStorage.removeItem(`${MODULE_ID}:key`);
    } catch (_) {}
    ui.notifications?.info?.('D&D 5e Item Generator was updated — please sign in again.');
  }
  if (currentVersion) {
    try { localStorage.setItem(storedVersionKey, currentVersion); } catch (_) {}
  }

  foundry.applications.handlebars.loadTemplates([
    `modules/${MODULE_ID}/templates/builder.html`,
  ]);
  console.log(`D&D 5e Item Generator ready (version: ${currentVersion}).`);

  startHeartbeat(MODULE_ID);

  if (game.user.isGM && !game.settings.get(MODULE_ID, 'welcomeMessageShown')) {
    const welcomeContent = `
<h3>Welcome to the D&D 5e Item Generator!</h3>
<p>The generator has opened automatically — you're ready to start generating items right away.</p>
<p>You can reopen it any time from the <em>Item Generator</em> button in the <strong>Items</strong> or <strong>Compendium</strong> sidebar header.</p>`.trim();

    ChatMessage.create({
      content: welcomeContent,
      whisper: game.users.filter(u => u.isGM).map(u => u.id),
    });
    game.settings.set(MODULE_ID, 'welcomeMessageShown', true);
    openBuilder(adapter);
    checkForModuleUpdate(MODULE_ID, adapter.module.githubUrl).catch(() => {});
  }
});
