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
import { SettlementSheet, openWithFixture } from './settlement-sheet.js';
import { NationSheet }                from './nation-sheet.js';
import { applyDailyTick, applyTax, applyFestival, applyPlague, applyFamine } from './economy.js';
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
  if (Hb && !Hb.helpers?.mulNum) {
    Hb.registerHelper('mulNum', (a, b) => Math.round(Number(a) * Number(b) * 100) / 100);
  }
  if (Hb && !Hb.helpers?.neqNum) {
    Hb.registerHelper('neqNum', (a, b) => Number(a) !== Number(b));
  }
} catch (_) { /* deferred to init below */ }

class ResetWelcomeMessageMenu {
  render() {
    foundry.applications.api.DialogV2.confirm({
      window:      { title: game.i18n.localize('SettlementBuilder.Settings.ResetWelcome.Name') },
      content:     `<p>${game.i18n.localize('SettlementBuilder.Settings.ResetWelcome.ConfirmContent')}</p>`,
      yes:         { label: game.i18n.localize('SettlementBuilder.Settings.ResetWelcome.ConfirmLabel'), icon: 'fa-solid fa-rotate-left' },
      no:          { label: 'Cancel' },
      rejectClose: false,
    }).then(ok => {
      if (ok) {
        game.settings.set(MODULE_ID, 'welcomeMessageShown', false);
        ui.notifications.info(game.i18n.localize('SettlementBuilder.Settings.ResetWelcome.Success'));
      }
    }).catch(() => {});
    return this;
  }
}

class ClearSettlementsMenu {
  render() {
    foundry.applications.api.DialogV2.confirm({
      window:      { title: game.i18n.localize('SettlementBuilder.Settings.ClearSettlements.Name') },
      content:     `<p>${game.i18n.localize('SettlementBuilder.Settings.ClearSettlements.ConfirmContent')}</p>`,
      yes:         { label: game.i18n.localize('SettlementBuilder.Settings.ClearSettlements.ConfirmLabel'), icon: 'fa-solid fa-trash' },
      no:          { label: 'Cancel' },
      rejectClose: false,
    }).then(ok => {
      if (ok) {
        new Storage(MODULE_ID).setKey('');
        ui.notifications.info(game.i18n.localize('SettlementBuilder.Settings.ClearSettlements.Success'));
      }
    }).catch(() => {});
    return this;
  }
}

class FixtureSheetMenu {
  async render() {
    await openWithFixture();
    return this;
  }
}

class ResetCalendarMenu {
  render() {
    if (!game.modules?.get('Pf2eCalendarTimeline')?.active) {
      ui.notifications.warn(game.i18n.localize('SettlementBuilder.Settings.ResetCalendar.NotActive'));
      return this;
    }
    foundry.applications.api.DialogV2.confirm({
      window:      { title: game.i18n.localize('SettlementBuilder.Settings.ResetCalendar.Name') },
      content:     `<p>${game.i18n.localize('SettlementBuilder.Settings.ResetCalendar.ConfirmContent')}</p>`,
      yes:         { label: game.i18n.localize('SettlementBuilder.Settings.ResetCalendar.ConfirmLabel'), icon: 'fa-solid fa-calendar-xmark' },
      no:          { label: 'Cancel' },
      rejectClose: false,
    }).then(ok => {
      if (!ok) return;
      game.settings.set('Pf2eCalendarTimeline', 'state', {
        currentDate: { year: 4725, month: 6, day: 1, hour: 8 },
        calendarDef: {
          monthsPerYear: 12,
          daysPerMonth:  [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31],
          weekdays:      ['Moonday', 'Toilday', 'Wealday', 'Oathday', 'Fireday', 'Starday', 'Sunday'],
          monthNames:    ['Abadius', 'Calistril', 'Pharast', 'Gozran', 'Desnus', 'Sarenith',
                          'Erastus', 'Arodus', 'Rova', 'Lamashan', 'Neth', 'Kuthona'],
        },
        events: [],
      });
      ui.notifications.info(game.i18n.localize('SettlementBuilder.Settings.ResetCalendar.Success'));
    }).catch(() => {});
    return this;
  }
}

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
    if (Hb && !Hb.helpers?.mulNum) {
      Hb.registerHelper('mulNum', (a, b) => Math.round(Number(a) * Number(b) * 100) / 100);
    }
    if (Hb && !Hb.helpers?.neqNum) {
      Hb.registerHelper('neqNum', (a, b) => Number(a) !== Number(b));
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
  game.settings.register(MODULE_ID, 'incomeJitterPct', {
    name: 'SettlementBuilder.Settings.IncomeJitter.Name',
    hint: 'SettlementBuilder.Settings.IncomeJitter.Hint',
    scope: 'world', config: true, type: Number,
    range: { min: 0, max: 50, step: 1 },
    default: 15,
  });

  game.settings.registerMenu(MODULE_ID, 'resetWelcome', {
    name:       'SettlementBuilder.Settings.ResetWelcome.Name',
    label:      'SettlementBuilder.Settings.ResetWelcome.Label',
    hint:       'SettlementBuilder.Settings.ResetWelcome.Hint',
    icon:       'fa-solid fa-rotate-left',
    type:       ResetWelcomeMessageMenu,
    restricted: true,
  });
  game.settings.registerMenu(MODULE_ID, 'clearSettlements', {
    name:       'SettlementBuilder.Settings.ClearSettlements.Name',
    label:      'SettlementBuilder.Settings.ClearSettlements.Label',
    hint:       'SettlementBuilder.Settings.ClearSettlements.Hint',
    icon:       'fa-solid fa-trash',
    type:       ClearSettlementsMenu,
    restricted: true,
  });
  game.settings.registerMenu(MODULE_ID, 'resetCalendar', {
    name:       'SettlementBuilder.Settings.ResetCalendar.Name',
    label:      'SettlementBuilder.Settings.ResetCalendar.Label',
    hint:       'SettlementBuilder.Settings.ResetCalendar.Hint',
    icon:       'fa-solid fa-calendar-xmark',
    type:       ResetCalendarMenu,
    restricted: true,
  });
  game.settings.registerMenu(MODULE_ID, 'fixtureSheet', {
    name:       'SettlementBuilder.Settings.FixtureSheet.Name',
    label:      'SettlementBuilder.Settings.FixtureSheet.Label',
    hint:       'SettlementBuilder.Settings.FixtureSheet.Hint',
    icon:       'fa-solid fa-flask-vial',
    type:       FixtureSheetMenu,
    restricted: true,
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

// Reverse-link (#110): NPCs generated through a settlement bridge carry a
// homeSettlementId flag (set in integrations.js). Show a button on their
// actor sheet that jumps back to the settlement journal.
function injectHomeSettlementLink(app, html) {
  const actor = app?.actor || app?.document;
  const homeId = actor?.getFlag?.(MODULE_ID, 'homeSettlementId');
  if (!homeId) return;

  const root = html instanceof HTMLElement ? html : html?.[0];
  if (!root || root.querySelector('.settlement-builder-home-link')) return;

  const anchor = root.querySelector('.sheet-header .tags') || root.querySelector('.sheet-header') || root.querySelector('header');
  if (!anchor) return;

  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'settlement-builder-home-link';
  btn.innerHTML = '<i class="fa-solid fa-house"></i> Open Home Settlement';
  btn.title = 'Open the settlement this NPC was generated for';
  btn.addEventListener('click', (ev) => {
    ev.preventDefault();
    ev.stopPropagation();
    const journal = game.journal?.get(homeId);
    if (journal) journal.sheet.render(true);
    else ui.notifications?.warn('The home settlement for this NPC no longer exists.');
  });
  anchor.appendChild(btn);
}

Hooks.on('renderActorSheetPF2e', injectHomeSettlementLink);
Hooks.on('renderActorSheet', (app, html) => {
  if (game.system?.id !== 'pf2e') return;
  injectHomeSettlementLink(app, html);
});

// Settlement icons by kind (#107): give village/town/city/nation journals a
// distinct icon + colour in the journal directory listing.
const KIND_ICONS = {
  village: 'fa-house-chimney',
  town:    'fa-shop',
  city:    'fa-city',
  nation:  'fa-flag',
};

function applySettlementDirectoryIcons(app, html) {
  const root = html instanceof HTMLElement ? html : html?.[0];
  if (!root) return;
  const rows = root.querySelectorAll('li[data-entry-id], li[data-document-id]');
  for (const row of rows) {
    const id = row.dataset.entryId || row.dataset.documentId;
    const journal = id && game.journal?.get(id);
    const s = journal && getSettlement(journal);
    if (!s) continue;
    const icon = row.querySelector('.document-name i, .entry-name i, a i');
    if (!icon) continue;
    icon.className = `fa-solid ${KIND_ICONS[s.kind] || 'fa-city'} settlement-kind-icon settlement-kind-icon--${s.kind}`;
    icon.title = s.kind.charAt(0).toUpperCase() + s.kind.slice(1);
  }
}

Hooks.on('renderJournalDirectory', applySettlementDirectoryIcons);
Hooks.on('renderJournalDirectoryPF2e', applySettlementDirectoryIcons);

function getCurrentWeekday() {
  try {
    const state = game.settings.get('Pf2eCalendarTimeline', 'state');
    const { year, month, day } = state?.currentDate || {};
    const { weekdays, daysPerMonth } = state?.calendarDef || {};
    if (!year || !weekdays?.length || !Array.isArray(daysPerMonth)) return null;
    const yearDays = daysPerMonth.reduce((s, d) => s + d, 0);
    let total = (year - 1) * yearDays;
    for (let m = 0; m < month - 1; m++) total += daysPerMonth[m];
    total += day;
    return total % weekdays.length;
  } catch (_) { return null; }
}

// Calendar integration — listen for events fired by Pf2eCalendarTimeline.
Hooks.on('Pf2eCalendarTimeline.dayAdvanced', async ({ days = 1 } = {}) => {
  try {
    const weekday = getCurrentWeekday();
    const journals = game.journal?.contents || [];
    for (const j of journals) {
      const s = getSettlement(j);
      if (!s || s.kind === 'nation') continue;
      await applyDailyTick(j, days, weekday);
    }
  } catch (err) {
    log('error', 'dayAdvanced handler failed', err);
  }
});

Hooks.on('Pf2eCalendarTimeline.eventFired', async (event = {}) => {
  try {
    const handled = new Set(['tax', 'festival', 'plague', 'famine']);
    if (!handled.has(event.kind)) return;
    const ids = event.payload?.targetSettlementIds || [];
    for (const id of ids) {
      const j = game.journal?.get(id);
      if (!j) continue;
      if (event.kind === 'tax')      await applyTax(j, event.payload);
      if (event.kind === 'festival') await applyFestival(j, event.payload);
      if (event.kind === 'plague')   await applyPlague(j, event.payload);
      if (event.kind === 'famine')   await applyFamine(j, event.payload);
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
