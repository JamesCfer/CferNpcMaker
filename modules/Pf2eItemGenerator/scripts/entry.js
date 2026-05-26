/**
 * PF2e Item Generator — module entry point.
 * Registers in the Items directory (not Actors) since output goes to game.items.
 */

import { openBuilder, ensureBuilder } from './core/app.js';
import { checkForModuleUpdate } from './core/update-check.js';
import { registerSidebar }      from './core/sidebar.js';
import { startHeartbeat }       from './core/heartbeat.js';
import { Storage }              from './core/storage.js';
import { Pf2eItemAdapter }      from './adapter.js';

const adapter   = new Pf2eItemAdapter();
const MODULE_ID = adapter.module.id;

const openFn = () => {
  openBuilder(adapter);
  checkForModuleUpdate(MODULE_ID, adapter.module.githubUrl).catch(() => {});
};

registerSidebar(MODULE_ID, openFn, {
  buttonLabel: 'Item Generator',
  buttonIcon:  adapter.module.icon,
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

// Cross-module bridge — lets sibling modules pre-fill the item builder.
Hooks.on('Pf2eItemGenerator.openWithPrefill', (payload = {}) => {
  try {
    const app = ensureBuilder(adapter);
    app.render({ force: true });
    const tryFill = (attemptsLeft) => {
      const form = app.element?.querySelector?.('.npc-form');
      if (!form) {
        if (attemptsLeft > 0) setTimeout(() => tryFill(attemptsLeft - 1), 100);
        return;
      }
      const set = (sel, val) => { const el = form.querySelector(sel); if (el != null && val != null) el.value = val; };
      set('[name="name"]',        payload.name);
      set('[name="level"]',       payload.level);
      set('[name="itemType"]',    payload.itemType);
      set('[name="description"]', payload.description);
    };
    tryFill(20);
    if (typeof payload.onCreate === 'function') {
      const handler = (item) => { try { payload.onCreate(item); } catch (e) { console.error(e); } };
      const id = Hooks.on('createItem', (item) => { handler(item); Hooks.off('createItem', id); });
    }
  } catch (err) {
    console.error('[Pf2eItemGenerator] openWithPrefill failed', err);
  }
});

Hooks.once('ready', () => {
  const mod = game.modules?.get(MODULE_ID);
  const currentVersion = mod?.version || '';

  const storage = new Storage(MODULE_ID);
  const storedVersion = storage.getVersion();

  if (currentVersion && storedVersion && currentVersion !== storedVersion) {
    storage.setKey('');
    ui.notifications?.info?.('Item Generator was updated — please sign in again.');
  }
  if (currentVersion) storage.setVersion(currentVersion);

  foundry.applications.handlebars.loadTemplates([
    `modules/${MODULE_ID}/templates/builder.html`,
  ]);
  console.log(`PF2e Item Generator ready (version: ${currentVersion}).`);

  startHeartbeat(MODULE_ID);

  if (game.user.isGM && !game.settings.get(MODULE_ID, 'welcomeMessageShown')) {
    const welcomeContent = `
<h3>Welcome to the PF2e Item Generator!</h3>
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
