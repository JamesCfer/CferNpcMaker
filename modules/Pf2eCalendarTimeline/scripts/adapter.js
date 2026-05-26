/**
 * Minimal SystemAdapter so the calendar module still shows in the shared
 * home grid even though it has no AI generation flow.
 */

import { SystemAdapter } from './core/adapter.js';
import { detectModuleFolder } from './core/utils.js';

const MODULE_FOLDER = detectModuleFolder('Pf2eCalendarTimeline');

export class CalendarAdapter extends SystemAdapter {
  get moduleFolder() { return MODULE_FOLDER; }
  get module() {
    return {
      id:           'Pf2eCalendarTimeline',
      label:        'PF2e Calendar',
      icon:         'fa-solid fa-calendar-days',
      githubUrl:    'https://github.com/JamesCfer/Pf2eCalendarTimeline',
      historyLabel: 'Recent Calendar Activity',
    };
  }
  get systemId() { return 'pf2e'; }
  get supportsImageGeneration() { return false; }
  get formConfig() { return { documentNoun: 'calendar event' }; }
  get progressSteps() { return ['Saving event…']; }

  gatherFormData(_form) { throw new Error('Calendar uses its own UI — not the BuilderApp form.'); }
  historyEntryFromForm(_) { return { name: 'event' }; }
  historyMeta() { return ''; }
  async generate(_) { throw new Error('Calendar uses its own UI.'); }
}
