/**
 * Pf2eNationsAndCitiesMaker — entry point.
 * Wires sidebar button, settings, custom journal sheet swap, calendar listener.
 */

import { openBuilder, ensureBuilder } from './core/app.js';
import { checkForModuleUpdate }       from './core/update-check.js';
import { registerSidebar }            from './core/sidebar.js';
import { startHeartbeat }             from './core/heartbeat.js';
import { Storage }                    from './core/storage.js';
import { SettlementAdapter }          from './adapter.js';
import { MODULE_ID, getSettlement }   from './constants.js';
import { SettlementSheet }            from './settlement-sheet.js';
import { NationSheet }                from './nation-sheet.js';
import { applyDailyTick, applyTax }   from './economy.js';
import { getTemplate, randomName, randomSettlement } from './templates.js';
import { log }                        from './logger.js';

const adapter = new SettlementAdapter();

const openFn = () => {
  openBuilder(adapter);
  checkForModuleUpdate(MODULE_ID, adapter.module.githubUrl).catch(() => {});
};

registerSidebar(MODULE_ID, openFn, {
  buttonLabel: 'Settlement Builder',
  buttonIcon:  adapter.module.icon,
  directories: ['journal'],
});

// Register a couple of small Handlebars helpers used by the sheet partials.
// (Avoids depending on whichever core helpers a given Foundry build ships.)
try {
  const Hb = (globalThis.Handlebars || foundry?.applications?.handlebars?.Handlebars);
  if (Hb && !Hb.helpers?.eqStr) {
    Hb.registerHelper('eqStr', (a, b) => String(a) === String(b));
  }
  if (Hb && !Hb.helpers?.pfPad) {
    Hb.registerHelper('pfPad', (n, w) => String(n ?? '').padStart(w || 2, '0'));
  }
} catch (_) { /* deferred to init below */ }

Hooks.once('init', () => {
  // Re-register helpers at init in case Handlebars wasn't ready at module-load.
  try {
    const Hb = (globalThis.Handlebars || foundry?.applications?.handlebars?.Handlebars);
    if (Hb && !Hb.helpers?.eqStr) {
      Hb.registerHelper('eqStr', (a, b) => String(a) === String(b));
    }
    if (Hb && !Hb.helpers?.pfPad) {
      Hb.registerHelper('pfPad', (n, w) => String(n ?? '').padStart(w || 2, '0'));
    }
  } catch (err) { log('warn', 'helper registration', err); }

  game.settings.register(MODULE_ID, 'logLevel', {
    name: 'Log Level',
    hint: 'Controls how much Pf2eNationsAndCitiesMaker writes to the browser console.',
    scope: 'world', config: true, type: String,
    choices: { error: 'Error', warn: 'Warn', info: 'Info', debug: 'Debug' },
    default: 'info',
  });
  game.settings.register(MODULE_ID, 'devMode', {
    name: 'Developer Mode',
    hint: 'When enabled, all webhook URLs are routed to the -dev endpoints.',
    scope: 'world', config: true, type: Boolean, default: false,
  });
  game.settings.register(MODULE_ID, 'allowOfflineStub', {
    name: 'Allow offline stub settlements',
    hint: 'If the AI endpoint is unreachable, generate a structured stub instead of failing. Useful while the backend is being built out.',
    scope: 'world', config: true, type: Boolean, default: true,
  });
  game.settings.register(MODULE_ID, 'welcomeMessageShown', {
    scope: 'world', config: false, type: Boolean, default: false,
  });
});

// Intercept JournalEntry.sheet rendering — when the journal carries a
// settlement flag, open our custom sheet instead of the stock journal sheet.
Hooks.on('renderJournalSheet', (app, html) => {
  try {
    const journal = app?.document;
    if (!journal) return;
    const s = getSettlement(journal);
    if (!s) return;
    // Defer the swap so we don't recurse during render.
    app.close({ submit: false });
    queueMicrotask(() => {
      if (s.kind === 'nation') new NationSheet(journal).render(true);
      else                     new SettlementSheet(journal).render(true);
    });
  } catch (err) {
    log('error', 'sheet swap failed', err);
  }
});

// Calendar integration — listen for events fired by Pf2eCalendarTimeline.
Hooks.on('Pf2eCalendarTimeline.dayAdvanced', async ({ days = 1 } = {}) => {
  try {
    const journals = game.journal?.contents || [];
    for (const j of journals) {
      const s = getSettlement(j);
      if (!s || s.kind === 'nation') continue;
      await applyDailyTick(j, days);
    }
  } catch (err) {
    log('error', 'dayAdvanced handler failed', err);
  }
});

Hooks.on('Pf2eCalendarTimeline.eventFired', async (event = {}) => {
  try {
    if (event.kind !== 'tax') return;
    const ids = event.payload?.targetSettlementIds || [];
    for (const id of ids) {
      const j = game.journal?.get(id);
      if (!j) continue;
      await applyTax(j, event.payload);
    }
  } catch (err) {
    log('error', 'eventFired handler failed', err);
  }
});

// Delegated handlers for the Settlement Builder form's template / dice buttons.
// We register once on document so they survive any open/close render cycle.
function applyTemplateFields(form, fields) {
  const set = (sel, val) => {
    const el = form.querySelector(sel);
    if (!el || val == null) return;
    if (el.type === 'checkbox') el.checked = !!val;
    else el.value = val;
  };
  if (fields.name)              set('[name="name"]',           fields.name);
  if (fields.kind)              set('[name="kind"]',           fields.kind);
  if (fields.size)              set('[name="size"]',           fields.size);
  if (fields.biome)             set('[name="biome"]',          fields.biome);
  if (fields.governmentHint)    set('[name="governmentHint"]', fields.governmentHint);
  if (fields.populationHint)    set('[name="populationHint"]', fields.populationHint);
  if (fields.description)       set('[name="description"]',    fields.description);
  if ('includeStores'     in fields) set('[name="includeStores"]',     fields.includeStores);
  if ('includeMilitary'   in fields) set('[name="includeMilitary"]',   fields.includeMilitary);
  if ('includeLeadership' in fields) set('[name="includeLeadership"]', fields.includeLeadership);
}

document.addEventListener('click', (ev) => {
  const form = ev.target?.closest?.('.settlement-form');
  if (!form) return;
  if (ev.target.closest('.btn-roll-name')) {
    ev.preventDefault();
    const kind  = form.querySelector('[name="kind"]')?.value || 'town';
    const biome = form.querySelector('[name="biome"]')?.value || '';
    const nameInput = form.querySelector('[name="name"]');
    if (nameInput) nameInput.value = randomName({ kind, biome });
  }
  if (ev.target.closest('.btn-randomize-settlement')) {
    ev.preventDefault();
    const { templateId, fields } = randomSettlement();
    const sel = form.querySelector('.settlement-template-select');
    if (sel) sel.value = templateId;
    applyTemplateFields(form, fields);
  }
});

document.addEventListener('change', (ev) => {
  const sel = ev.target?.closest?.('.settlement-template-select');
  if (!sel) return;
  const form = sel.closest('.settlement-form');
  if (!form) return;
  const t = getTemplate(sel.value);
  if (t && t.id !== 'custom') applyTemplateFields(form, t.fields);
});

Hooks.once('ready', () => {
  const mod = game.modules?.get(MODULE_ID);
  const currentVersion = mod?.version || '';
  const storage = new Storage(MODULE_ID);
  const storedVersion = storage.getVersion();
  if (currentVersion && storedVersion && currentVersion !== storedVersion) {
    storage.setKey('');
    ui.notifications?.info?.('Settlement Builder was updated — please sign in again.');
  }
  if (currentVersion) storage.setVersion(currentVersion);

  foundry.applications.handlebars.loadTemplates([
    `modules/${MODULE_ID}/templates/builder.html`,
    `modules/${MODULE_ID}/templates/city-sheet.hbs`,
    `modules/${MODULE_ID}/templates/nation-sheet.hbs`,
    `modules/${MODULE_ID}/templates/partials/statblock.hbs`,
    `modules/${MODULE_ID}/templates/partials/stores-tab.hbs`,
    `modules/${MODULE_ID}/templates/partials/guards-tab.hbs`,
    `modules/${MODULE_ID}/templates/partials/leadership-tab.hbs`,
    `modules/${MODULE_ID}/templates/partials/production-tab.hbs`,
  ]).catch(err => log('warn', 'template preload', err));

  log('info', `ready (version: ${currentVersion})`);
  startHeartbeat(MODULE_ID);

  if (game.user.isGM && !game.settings.get(MODULE_ID, 'welcomeMessageShown')) {
    ChatMessage.create({
      content: `<h3>Welcome to the PF2e Settlement Builder!</h3>
        <p>Open the builder from the <strong>Journal</strong> sidebar header to generate cities, towns and nations.</p>
        <p>Generated settlements live as JournalEntries with a custom, PF2e-themed sheet.</p>`,
      whisper: game.users.filter(u => u.isGM).map(u => u.id),
    });
    game.settings.set(MODULE_ID, 'welcomeMessageShown', true);
    openBuilder(adapter);
  }
});
