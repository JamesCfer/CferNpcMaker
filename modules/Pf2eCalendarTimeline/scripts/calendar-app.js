/**
 * CalendarApp — the in-world calendar UI.
 * Lets the GM view the date, advance days, and manage scheduled events
 * (taxes, paydays, festivals) targeted at settlement journals.
 */

import { MODULE_ID, getState, patchState } from './state.js';
import { DEFAULT_CALENDAR, CALENDAR_PRESETS, BIOMES, addDays, cmpDate, computeNext,
         dayOrdinal, weekdayIndex, describeRule, formatDate,
         seasonForMonth, rollWeather } from './scheduler.js';

const WEATHER_ICONS = {
  clear: 'fa-sun', overcast: 'fa-cloud', rain: 'fa-cloud-rain',
  storm: 'fa-cloud-bolt', snow: 'fa-snowflake',
};

const { HandlebarsApplicationMixin, ApplicationV2 } = foundry.applications.api;

const TAX_TYPES = [
  { value: 'income',   label: 'Income tax' },
  { value: 'poll',     label: 'Poll tax' },
  { value: 'trade',    label: 'Trade tax' },
  { value: 'property', label: 'Property tax' },
  { value: 'festival', label: 'Festival / payout' },
];

export class CalendarApp extends HandlebarsApplicationMixin(ApplicationV2) {
  static DEFAULT_OPTIONS = {
    id: 'pf2e-calendar-app',
    classes: ['pf2e-calendar-app'],
    tag: 'form',
    window: { title: 'PF2e Calendar & Timeline', resizable: true },
    position: { width: 760, height: 640 },
    actions: {
      tickDay:     function() { this._advance(1); },
      tickWeek:    function() { this._advance(7); },
      advanceN:    function() { this._advancePrompt(); },
      prevMonth:   function() { this._navMonth(-1); },
      nextMonth:   function() { this._navMonth(1); },
      addEvent:    function() { this._onAddEvent(); },
      removeEvent: function(ev) { this._onRemoveEvent(ev); },
      saveDate:    function(ev) { this._onSaveDate(ev); },
      loadPreset:  function(ev) { this._onLoadPreset(ev); },
      changeBiome: function(ev) { this._onChangeBiome(ev); },
    },
  };

  static PARTS = {
    body: { template: `modules/${MODULE_ID}/templates/calendar.hbs` },
  };

  constructor(options = {}) {
    super(options);
    const s = getState();
    this.viewYear  = s.currentDate.year;
    this.viewMonth = s.currentDate.month;
  }

  async _prepareContext() {
    const state = getState();
    const cal   = state.calendarDef || DEFAULT_CALENDAR;
    const current = state.currentDate;

    const month = this.viewMonth;
    const year  = this.viewYear;
    const dim   = cal.daysPerMonth[month - 1];

    const firstWeekday = weekdayIndex({ year, month, day: 1 }, cal);
    const grid = [];
    for (let i = 0; i < firstWeekday; i++) grid.push({ blank: true });
    for (let d = 1; d <= dim; d++) {
      const date = { year, month, day: d };
      const isCurrent = cmpDate(date, current) === 0;
      const events = (state.events || []).filter(e => cmpDate(e.nextRun || {}, date) === 0);
      grid.push({ blank: false, day: d, isCurrent, eventCount: events.length });
    }

    // Settlement journals (from sibling module) — used for event targeting select.
    const settlementJournals = (game.journal?.contents || [])
      .filter(j => j.getFlag?.('Pf2eNationsAndCitiesMaker', 'settlement'))
      .map(j => ({ id: j.id, name: j.name }));

    const upcoming = (state.events || [])
      .slice()
      .sort((a, b) => cmpDate(a.nextRun || {}, b.nextRun || {}))
      .map(e => ({
        ...e,
        ruleLabel:  describeRule(e.rule, cal),
        whenLabel:  e.nextRun ? formatDate(e.nextRun, cal) : '—',
        targetLabel: (e.payload?.targetSettlementIds || [])
          .map(id => settlementJournals.find(s => s.id === id)?.name || id)
          .join(', '),
      }));

    const biome = state.biome || 'temperate';
    const season = state.season || seasonForMonth(current.month, cal);
    const weather = state.weather || 'clear';

    return {
      state, cal,
      calendarPresets: CALENDAR_PRESETS,
      monthName: cal.monthNames?.[month - 1] || `M${month}`,
      year, month,
      currentLabel: formatDate(current, cal),
      currentWeekday: cal.weekdays[weekdayIndex(current, cal)],
      grid, weekdays: cal.weekdays,
      taxTypes: TAX_TYPES,
      settlementJournals,
      upcoming,
      biomes: BIOMES.map(b => ({ value: b, selected: b === biome })),
      season,
      seasonLabel: season.charAt(0).toUpperCase() + season.slice(1),
      weather,
      weatherLabel: weather.charAt(0).toUpperCase() + weather.slice(1),
      weatherIcon: WEATHER_ICONS[weather] || 'fa-cloud-question',
    };
  }

  _navMonth(delta) {
    const cal = getState().calendarDef || DEFAULT_CALENDAR;
    this.viewMonth += delta;
    while (this.viewMonth < 1) { this.viewMonth += cal.monthsPerYear; this.viewYear -= 1; }
    while (this.viewMonth > cal.monthsPerYear) { this.viewMonth -= cal.monthsPerYear; this.viewYear += 1; }
    this.render(false);
  }

  /**
   * Advance the world clock by N days, firing dayAdvanced once and eventFired
   * for each scheduled event whose nextRun falls within the window.
   */
  async _advance(days) {
    days = Math.max(1, Math.floor(Number(days) || 1));
    const cur = getState();
    const cal = cur.calendarDef || DEFAULT_CALENDAR;
    const fromDate = cur.currentDate;
    const toDate   = addDays(fromDate, days, cal);

    const fired = [];
    for (const ev of cur.events || []) {
      // Fire events whose nextRun is strictly > fromDate and <= toDate.
      let nextRun = ev.nextRun || fromDate;
      while (cmpDate(nextRun, toDate) <= 0 && cmpDate(nextRun, fromDate) > 0) {
        fired.push({ ...ev, nextRun });
        nextRun = computeNext(ev.rule, nextRun, cal);
      }
      ev.nextRun = nextRun;
    }

    const biome = cur.biome || 'temperate';
    const fromSeason = seasonForMonth(fromDate.month, cal);
    const toSeason   = seasonForMonth(toDate.month, cal);
    const weather    = rollWeather(biome, toSeason);

    cur.currentDate = toDate;
    cur.season  = toSeason;
    cur.weather = weather;
    await game.settings.set(MODULE_ID, 'state', cur);

    Hooks.callAll('Pf2eCalendarTimeline.dayAdvanced', {
      days, currentDate: toDate, fromDate,
    });
    for (const ev of fired) {
      Hooks.callAll('Pf2eCalendarTimeline.eventFired', ev);
    }
    if (toSeason !== fromSeason) {
      Hooks.callAll('Pf2eCalendarTimeline.seasonChanged', {
        season: toSeason, previousSeason: fromSeason, currentDate: toDate,
      });
    }
    Hooks.callAll('Pf2eCalendarTimeline.weatherChanged', {
      weather, season: toSeason, biome, currentDate: toDate,
    });
    ui.notifications?.info?.(`Advanced ${days} day${days === 1 ? '' : 's'} → ${formatDate(toDate, cal)}. ${fired.length} event(s) fired.`);
    this.viewYear  = toDate.year;
    this.viewMonth = toDate.month;
    this.render(false);
  }

  async _advancePrompt() {
    const res = await foundry.applications.api.DialogV2.prompt({
      window: { title: 'Advance Time' },
      content: `<label>Days <input name="days" type="number" min="1" value="1" /></label>`,
      ok: { label: 'Advance', callback: (_e, _b, dlg) => Number(dlg.element.querySelector('[name="days"]').value) || 1 },
      rejectClose: false,
    }).catch(() => null);
    if (res) this._advance(res);
  }

  async _onSaveDate(ev) {
    const form = ev.currentTarget?.closest('.pf2e-calendar-current');
    if (!form) return;
    const y = Number(form.querySelector('[name="curYear"]').value)  || 1;
    const m = Math.max(1, Number(form.querySelector('[name="curMonth"]').value) || 1);
    const d = Math.max(1, Number(form.querySelector('[name="curDay"]').value)   || 1);
    await patchState(s => { s.currentDate = { year: y, month: m, day: d, hour: 8 }; });
    this.viewYear = y; this.viewMonth = m;
    this.render(false);
  }

  async _onAddEvent() {
    const root = this.element;
    const label    = root.querySelector('[name="evLabel"]').value.trim() || 'Untitled event';
    const taxType  = root.querySelector('[name="evTaxType"]').value || 'income';
    const unit     = root.querySelector('[name="evUnit"]').value || 'day';
    const every    = Math.max(1, Number(root.querySelector('[name="evEvery"]').value) || 1);
    const ratePct  = Math.max(0, Number(root.querySelector('[name="evRate"]').value) || 0);
    const dayOfMonth = Math.max(1, Number(root.querySelector('[name="evDom"]').value) || 1);
    const month    = Math.max(1, Number(root.querySelector('[name="evMonth"]').value) || 1);
    const day      = Math.max(1, Number(root.querySelector('[name="evDay"]').value)   || 1);
    const targets  = Array.from(root.querySelectorAll('[name="evTargets"] option:checked')).map(o => o.value);

    const cur = getState();
    const cal = cur.calendarDef || DEFAULT_CALENDAR;
    const rule = { every, unit };
    if (unit === 'month') rule.dayOfMonth = dayOfMonth;
    if (unit === 'year')  { rule.month = month; rule.day = day; }
    const event = {
      id: foundry.utils.randomID(12),
      label, kind: 'tax',
      rule,
      nextRun: computeNext(rule, cur.currentDate, cal),
      payload: { taxType, ratePct, targetSettlementIds: targets },
    };
    await patchState(s => { s.events = [...(s.events || []), event]; });
    this.render(false);
  }

  async _onRemoveEvent(ev) {
    const id = ev.currentTarget?.dataset?.eventId;
    if (!id) return;
    await patchState(s => { s.events = (s.events || []).filter(e => e.id !== id); });
    this.render(false);
  }

  async _onChangeBiome(ev) {
    const biome = this.element.querySelector('[name="calBiome"]')?.value;
    if (!biome) return;
    await patchState(s => { s.biome = biome; });
    this.render(false);
  }

  async _onLoadPreset(ev) {
    const key = this.element.querySelector('[name="calPreset"]')?.value;
    if (!key) return;
    const preset = CALENDAR_PRESETS.find(p => p.key === key);
    if (!preset) return;
    await patchState(s => { s.calendarDef = foundry.utils.deepClone(preset.def); });
    this.viewMonth = Math.min(this.viewMonth, preset.def.monthsPerYear);
    this.render(false);
    ui.notifications?.info?.(`Calendar preset loaded: ${preset.label}`);
  }
}

let _instance = null;
export function openCalendar() {
  if (!_instance || !_instance.rendered) _instance = new CalendarApp();
  _instance.render({ force: true });
  return _instance;
}
