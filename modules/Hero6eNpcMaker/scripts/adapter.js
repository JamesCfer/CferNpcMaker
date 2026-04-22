/**
 * Hero System 6e NPC SystemAdapter — implements the BuilderApp contract for HDC import.
 *
 * The n8n hero6e endpoint returns HDC XML (either wrapped in JSON or raw).
 * On success we download the .hdc to the user's machine, then import it into
 * Foundry via the hero6efoundryvttv2 system's uploadFromXml API.
 */

import { SystemAdapter, postToN8n } from './core/adapter.js';
import { N8N_BASE, devUrl }          from './core/n8n.js';
import { detectModuleFolder }        from './core/utils.js';
import { sanitizeActorDataHero6e }   from './sanitizer.js';

const MODULE_FOLDER = detectModuleFolder('Hero6eNpcMaker');
const NPC_ENDPOINT  = `${N8N_BASE}/webhook/hero6e-hdc-builder`;

export const HERO6E_GENRES    = ['standard', 'superhero', 'pulp', 'dark_champions', 'fantasy', 'sci-fi'];
export const HERO6E_UNIVERSES = ['standard', 'mha', 'dc', 'marvel'];

export class Hero6eNpcAdapter extends SystemAdapter {
  get moduleFolder() { return MODULE_FOLDER; }

  get module() {
    return {
      id:           'Hero6eNpcMaker',
      label:        'HERO 6e',
      icon:         'fa-solid fa-bolt',
      githubUrl:    'https://github.com/JamesCfer/Hero6eNpcMaker',
      historyLabel: 'Created Characters',
    };
  }

  get systemId() { return 'hero6e'; }

  get supportsImageGeneration() { return true; }

  get formConfig() { return { documentNoun: 'character' }; }

  /* ── Form handling ──────────────────────────────────────── */

  gatherFormData(form) {
    const fd = new FormData(form);
    const name        = (fd.get('name')?.toString()?.trim()) || 'Generated Character';
    const level       = Number(fd.get('level')) || 150;
    const description = (fd.get('description')?.toString()?.trim()) || '';

    if (!description) throw new Error('Please provide a description for the character.');

    let universe = (fd.get('hero6eUniverse') || '').toLowerCase().trim();
    if (!HERO6E_UNIVERSES.includes(universe)) {
      const m = description.match(/\buniverse\s*:\s*([\w-]+)/i);
      if (m && HERO6E_UNIVERSES.includes(m[1].toLowerCase())) universe = m[1].toLowerCase();
    }
    if (!HERO6E_UNIVERSES.includes(universe)) universe = 'standard';

    let genre = (fd.get('hero6eGenre') || '').toLowerCase().trim();
    if (!HERO6E_GENRES.includes(genre)) {
      const m = description.match(/\bgenre\s*:\s*([\w_-]+)/i);
      if (m && HERO6E_GENRES.includes(m[1].toLowerCase())) genre = m[1].toLowerCase();
    }
    if (!HERO6E_GENRES.includes(genre)) genre = 'standard';

    const gearEl = form.querySelector('[name="hero6eCreateGear"]');
    let createGear = gearEl?.checked === true || fd.get('hero6eCreateGear') === 'on';
    if (!createGear) createGear = /\bgear\s*:\s*(yes|true|1)\b/i.test(description);

    return { name, level, description, universe, genre, createGear };
  }

  historyEntryFromForm(formData) {
    return {
      name:        formData.name,
      level:       formData.level,
      description: formData.description,
      universe:    formData.universe,
      genre:       formData.genre,
      createGear:  formData.createGear,
    };
  }

  historyMeta(entry) {
    const universePart = (entry.universe && entry.universe !== 'standard')
      ? `&nbsp;<span class="history-entry-universe">[${entry.universe.toUpperCase()}]</span>`
      : '';
    return `${entry.level}&nbsp;pts${universePart}`;
  }

  populateForm(form, entry) {
    const nameInput      = form.querySelector('[name="name"]');
    const levelInput     = form.querySelector('[name="level"]');
    const descTextarea   = form.querySelector('[name="description"]');
    const universeSelect = form.querySelector('[name="hero6eUniverse"]');
    const genreSelect    = form.querySelector('[name="hero6eGenre"]');
    const gearCheckbox   = form.querySelector('[name="hero6eCreateGear"]');
    if (nameInput)      nameInput.value        = entry.name ?? '';
    if (levelInput)     levelInput.value       = entry.level ?? 150;
    if (descTextarea)   descTextarea.value     = entry.description ?? '';
    if (universeSelect) universeSelect.value   = entry.universe ?? 'standard';
    if (genreSelect)    genreSelect.value      = entry.genre ?? 'standard';
    if (gearCheckbox)   gearCheckbox.checked   = !!entry.createGear;
  }

  /* ── Generation ─────────────────────────────────────────── */

  async generate({ formData, key, devMode, builderApp }) {
    const endpoint = devUrl(NPC_ENDPOINT, devMode);
    const payload  = {
      name:       formData.name,
      points:     formData.level,
      genre:      formData.genre,
      description: formData.description,
      universe:   formData.universe,
      createGear: formData.createGear,
    };

    const { response, responseText } = await postToN8n(endpoint, payload, key);

    let data;
    try { data = JSON.parse(responseText); } catch { data = null; }

    if (!response.ok) {
      throw new Error(data?.message || `Server returned status ${response.status}`);
    }
    if (data?.ok === false) throw new Error(data?.message || data?.error || 'Server rejected the request');

    // HDC XML — either in a JSON wrapper or raw body
    const hdcXml = data?.hdcXml || (!data ? responseText : null);
    if (!hdcXml || typeof hdcXml !== 'string') {
      throw new Error('No valid HDC data returned from server');
    }

    const parser   = new DOMParser();
    const xmlDoc   = parser.parseFromString(hdcXml, 'text/xml');
    const parseErr = xmlDoc.querySelector('parsererror');
    if (parseErr) {
      throw new Error(`Failed to parse HDC file: ${parseErr.textContent.substring(0, 200)}`);
    }

    const charName = xmlDoc.querySelector('CHARACTER_NAME')?.textContent?.trim()
                  || data?.name
                  || formData.name
                  || 'npc';

    // 1. Download the .hdc to the user's machine
    {
      const blob    = new Blob([hdcXml], { type: 'application/xml' });
      const blobUrl = URL.createObjectURL(blob);
      const anchor  = document.createElement('a');
      anchor.href     = blobUrl;
      anchor.download = `${charName}.hdc`;
      anchor.click();
      URL.revokeObjectURL(blobUrl);
      ui.notifications.info(`Downloaded "${charName}.hdc" to your device.`);
    }

    // 2. Import HDC via hero6e's uploadFromXml
    const templateEl   = xmlDoc.querySelector('TEMPLATE');
    const templateName = templateEl?.getAttribute('name')
                      || templateEl?.getAttribute('EXTENDS')
                      || 'builtIn.Heroic6E.hdt';

    const actor = await Actor.create({
      name: charName,
      type: 'pc',
      system: {
        is5e: false,
        CHARACTER: {
          TEMPLATE: { name: templateName, EXTENDS: templateName },
          BASIC_CONFIGURATION: {},
          CHARACTER_INFO: { CHARACTER_NAME: charName },
        },
      },
    });
    if (!actor) throw new Error('Failed to create actor for HDC import');

    // Monkey-patch _templateType to be null-safe during upload
    const origDescriptor = Object.getOwnPropertyDescriptor(
      Object.getPrototypeOf(actor), '_templateType'
    );
    Object.defineProperty(actor, '_templateType', {
      get() {
        const val = actor.system?.CHARACTER?.TEMPLATE?.name;
        return val || templateName;
      },
      configurable: true,
    });

    await actor.uploadFromXml(xmlDoc, {});

    if (origDescriptor) {
      Object.defineProperty(actor, '_templateType', origDescriptor);
    } else {
      delete actor._templateType;
    }

    // Optional post-import sanitization pass — runs silently
    try { sanitizeActorDataHero6e(actor.toObject()); } catch (_) {}

    return {
      document:   actor,
      exportData: {
        content:  hdcXml,
        filename: `${charName}.hdc`,
        mimeType: 'application/xml',
      },
      message: `Character "${actor.name}" created successfully via HDC import!`,
    };
  }
}
