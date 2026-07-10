/**
 * DayEventsPopover — small window listing every event scheduled for a
 * clicked calendar day, with per-row edit (label) and cancel (delete).
 */

import { MODULE_ID, getState, patchState } from './state.js';
import { cmpDate, describeRule, formatDate, DEFAULT_CALENDAR } from './scheduler.js';
import { escapeHtml } from './core/utils.js';

const { HandlebarsApplicationMixin, ApplicationV2 } = foundry.applications.api;

export class DayEventsPopover extends HandlebarsApplicationMixin(ApplicationV2) {
  static DEFAULT_OPTIONS = {
    id: 'pf2e-cal-day-popover',
    classes: ['pf2e-cal-day-popover'],
    window: { title: 'Day Events', resizable: false },
    position: { width: 380, height: 'auto' },
    actions: {
      editEvent:   function(ev) { this._onEditEvent(ev); },
      cancelEvent: function(ev) { this._onCancelEvent(ev); },
    },
  };

  static PARTS = {
    body: { template: `modules/${MODULE_ID}/templates/day-popover.hbs` },
  };

  constructor(date, parentApp, options = {}) {
    super(options);
    this.date = date;
    this.parentApp = parentApp;
  }

  get title() {
    const cal = getState().calendarDef || DEFAULT_CALENDAR;
    return `Events — ${formatDate(this.date, cal)}`;
  }

  async _prepareContext() {
    const state = getState();
    const cal = state.calendarDef || DEFAULT_CALENDAR;
    const events = (state.events || [])
      .filter(e => cmpDate(e.nextRun || {}, this.date) === 0)
      .map(e => ({ ...e, ruleLabel: describeRule(e.rule, cal) }));
    return { events, dateLabel: formatDate(this.date, cal) };
  }

  async _onEditEvent(ev) {
    const id = ev.currentTarget?.dataset?.eventId;
    const event = (getState().events || []).find(e => e.id === id);
    if (!event) return;
    const label = await foundry.applications.api.DialogV2.prompt({
      window: { title: 'Edit Event Label' },
      content: `<label>Label <input type="text" name="label" value="${escapeHtml(event.label ?? '')}" /></label>`,
      ok: { label: 'Save', callback: (_e, _b, dlg) => dlg.element.querySelector('[name="label"]').value.trim() },
      rejectClose: false,
    }).catch(() => null);
    if (!label) return;
    await patchState(s => {
      const e = (s.events || []).find(e => e.id === id);
      if (e) e.label = label;
    });
    this.render(false);
    this.parentApp?.render(false);
  }

  async _onCancelEvent(ev) {
    const id = ev.currentTarget?.dataset?.eventId;
    if (!id) return;
    await patchState(s => { s.events = (s.events || []).filter(e => e.id !== id); });
    this.parentApp?.render(false);
    const remaining = (getState().events || []).filter(e => cmpDate(e.nextRun || {}, this.date) === 0);
    if (remaining.length === 0) this.close();
    else this.render(false);
  }
}
