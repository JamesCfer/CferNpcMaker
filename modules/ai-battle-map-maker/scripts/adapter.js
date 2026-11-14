import { SystemAdapter, postToN8n, AuthError } from './core/adapter.js';
import { N8N_ENDPOINTS, PATREON_URL } from './core/n8n.js';
import { buildMapRequest, costSummary, measuredLabel, sceneFileName, drawingBox, selectRegionDrawing, sceneTraceOptions, plural,
  sceneEnvironment, pickWeather, overlayShapes, mapRecord, isSavedMap, unpackTrace, MAP_FIELDS, STYLE_FIELDS, BUILT_IN_PRESETS, presetFromValues, upsertPreset } from './map-logic.js';
import { applyLook } from './core/maps/look.js';
import { saveToWorld, sceneData, wallData, lightData, createMapScene, viewMapScene, notifyMapCreated, getSceneBackground, setSceneBackground,
  mapSceneFolder, setSceneThumbnail, hasWeather, startProgress } from './compat.js';
import { pasteRegion } from './region-paste.js';
import { Storage } from './core/storage.js';

const imageUrl = image => `data:${image.mediaType || 'image/png'};base64,${image.b64}`;
const loadImage = image => new Promise((resolve, reject) => {
  const img = new Image();
  img.onload = () => resolve(img);
  img.onerror = () => reject(new Error('The map image could not be decoded.'));
  img.src = imageUrl(image);
});
const encodeCanvas = (canvas, type, quality = 0.9) => new Promise(resolve => canvas.toBlob(resolve, type, type === 'image/webp' ? quality : undefined));
const loadUrl = (url, message) => new Promise((resolve, reject) => {
  const img = new Image();
  if (!url.startsWith('blob:') && !url.startsWith('data:')) img.crossOrigin = 'anonymous';
  img.onload = () => resolve(img);
  img.onerror = () => reject(new Error(message));
  img.src = url;
});
const imageCanvas = image => {
  const canvas = document.createElement('canvas');
  canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
  canvas.getContext('2d').drawImage(image, 0, 0);
  return canvas;
};
const pngPayload = image => ({ b64: imageCanvas(image).toDataURL('image/png').split(',')[1], mediaType: 'image/png' });
const MAX_MOCKUPS = 12;
const CHECKBOXES = ['autoPlan', 'walls', 'showTrace'];
const FORM_KEY = 'ai-battle-map-maker.form';
const PRESET_KEY = 'ai-battle-map-maker.presets';
const FLAG_SCOPE = 'ai-battle-map-maker';
const readStore = (key, fallback) => { try { return JSON.parse(localStorage.getItem(key)) ?? fallback; } catch { return fallback; } };
const writeStore = (key, value) => { try { localStorage.setItem(key, JSON.stringify(value)); } catch { /* storage full or blocked */ } };

/** Show an image large over the Foundry window. Click or Escape closes it. */
export function zoomImage(src, alt = 'Battle map') {
  const overlay = document.createElement('div');
  overlay.className = 'abmm-zoom';
  overlay.setAttribute('role', 'dialog');
  overlay.setAttribute('aria-label', `${alt}, enlarged. Click or press Escape to close.`);
  overlay.tabIndex = -1;
  const img = document.createElement('img');
  img.src = src; img.alt = alt;
  overlay.append(img);
  const close = () => { overlay.remove(); document.removeEventListener('keydown', onKey, true); if (src.startsWith('blob:')) URL.revokeObjectURL(src); };
  const onKey = event => { if (event.key === 'Escape') { event.stopPropagation(); close(); } };
  overlay.addEventListener('click', close);
  document.addEventListener('keydown', onKey, true);
  document.body.append(overlay);
  overlay.focus();
  return overlay;
}

/** Draw traced walls, doors, windows and lights over a preview canvas. */
export function drawTrace(ctx, trace, width, height) {
  const { walls, lights } = overlayShapes(trace, width, height);
  const line = Math.max(2, width / 320);
  ctx.save();
  ctx.lineCap = 'round';
  for (const [stroke, grow] of [['rgba(0,0,0,.75)', 2.2], [null, 1]]) {
    for (const wall of walls) {
      ctx.strokeStyle = stroke || { wall: '#fff36b', door: '#ff8a1f', window: '#5fd4ff' }[wall.kind];
      ctx.lineWidth = line * grow * (wall.kind === 'wall' ? 1 : 1.6);
      ctx.beginPath(); ctx.moveTo(wall.x1, wall.y1); ctx.lineTo(wall.x2, wall.y2); ctx.stroke();
    }
  }
  for (const light of lights) {
    ctx.beginPath(); ctx.arc(light.x, light.y, line * 3, 0, Math.PI * 2);
    ctx.fillStyle = light.color; ctx.globalAlpha = 0.9; ctx.fill();
    ctx.globalAlpha = 1; ctx.lineWidth = line * 0.8; ctx.strokeStyle = '#000'; ctx.stroke();
    ctx.beginPath(); ctx.arc(light.x, light.y, line * 7, 0, Math.PI * 2);
    ctx.strokeStyle = light.color; ctx.setLineDash([line * 2, line * 2]); ctx.stroke(); ctx.setLineDash([]);
  }
  ctx.restore();
}
const DIAL_LABELS = { detail: 'Detail', clutter: 'Props', wear: 'Wear', color: 'Color', tint: 'Color temperature', contrast: 'Lighting contrast', sharpness: 'Sharpness' };
const UNIT_LABELS = { imperial: { distance: 5, units: 'ft' }, metric: { distance: 1.5, units: 'm' } };

export function renderMapThumbnails(box, images, aspect, onSelect, { selected = -1, onZoom } = {}) {
  box.replaceChildren();
  box.style.setProperty('--cols', images.length === 1 ? '1' : '2');
  const ratio = typeof aspect === 'string' && /^\d+:\d+$/.test(aspect) ? aspect.replace(':', ' / ') : '4 / 3';
  box.style.setProperty('--map-aspect', ratio);
  images.forEach((image, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = 'map-thumbnail';
    button.setAttribute('aria-label', `Select mockup ${index + 1}`);
    const img = document.createElement('img');
    img.src = imageUrl(image);
    img.alt = `Mockup ${index + 1}`;
    const caption = document.createElement('span');
    caption.className = 'map-picked';
    const idle = image.earlier ? 'Earlier' : '';
    caption.textContent = index === selected ? 'Picked' : idle;
    caption.dataset.idle = idle;
    button.classList.toggle('is-selected', index === selected);
    button.classList.toggle('is-earlier', !!image.earlier);
    if (image.aspect && /^\d+:\d+$/.test(image.aspect)) button.style.setProperty('--map-aspect', image.aspect.replace(':', ' / '));
    button.append(img, caption);
    button.addEventListener('click', () => {
      box.querySelectorAll('.map-thumbnail').forEach((node, i) => {
        node.classList.toggle('is-selected', i === index);
        const label = node.querySelector('.map-picked');
        label.textContent = i === index ? 'Picked' : label.dataset.idle || '';
      });
      onSelect(index);
    });
    const wrap = document.createElement('div');
    wrap.className = 'map-thumb-wrap';
    wrap.append(button);
    if (onZoom) {
      const zoom = document.createElement('button');
      zoom.type = 'button';
      zoom.className = 'map-zoom-btn';
      zoom.title = 'View larger';
      zoom.setAttribute('aria-label', `View mockup ${index + 1} larger`);
      zoom.innerHTML = '<i class="fa-solid fa-magnifying-glass-plus" aria-hidden="true"></i>';
      zoom.addEventListener('click', () => onZoom(index));
      wrap.append(zoom);
    }
    box.append(wrap);
  });
}

export class BattleMapAdapter extends SystemAdapter {
  get moduleFolder() { return 'ai-battle-map-maker'; }
  get module() { return { id: 'ai-battle-map-maker', label: 'AI Battle Map Maker', icon: 'fa-solid fa-map',
    githubUrl: 'https://github.com/JamesCfer/CferNpcMaker', historyLabel: 'Created Maps' }; }
  get windowTitle() { return 'AI Battle Map Maker'; }
  get systemId() { return 'agnostic'; }
  get formConfig() { return { documentNoun: 'map' }; }
  gatherFormData(form) { return buildMapRequest(Object.fromEntries(new FormData(form))); }
  historyEntryFromForm(data) { return { name: data.description?.slice(0, 60) || 'Battle map', description: data.description }; }
  historyMeta() { return 'Battle map'; }
  populateForm(form, entry) {
    if (isSavedMap(entry)) { this.openSaved(entry); return; }
    form.elements.description.value = entry.description || '';
  }
  async generate() { throw new Error('Use the map mockup and final controls.'); }

  onFormMount(form) {
    this.form = form;
    // Mockups and the final map outlive the window: closing the builder must not lose paid work.
    this.mockups ||= [];
    this.selected = this.mockups.length ? this.selected ?? -1 : -1;
    this.final ||= null;
    this.sourceImage ||= null;
    // A request still running from before the window was closed finishes into this form.
    this.busy = !!this.busy;
    const field = name => form.elements.namedItem(name);
    const status = message => { form.querySelector('.map-status').textContent = message; };
    const error = message => { const el = form.querySelector('.map-error'); el.textContent = message || ''; el.hidden = !message; };
    this.status = status;
    this.error = error;
    const makeDials = (container, keys) => {
      for (const key of keys) {
        const wrapper = document.createElement('div');
        wrapper.className = 'field';
        const label = document.createElement('label');
        label.textContent = DIAL_LABELS[key];
        const slider = document.createElement('input');
        slider.type = 'range'; slider.name = key; slider.min = '1'; slider.max = '5'; slider.step = '1'; slider.value = '3';
        slider.id = `map-dial-${key}`;
        label.htmlFor = slider.id;
        const value = document.createElement('output'); value.textContent = '3';
        value.htmlFor = slider.id;
        slider.addEventListener('input', () => { value.textContent = slider.value; if (['color', 'tint', 'contrast', 'sharpness'].includes(key)) this.renderPreview(); });
        slider.addEventListener('change', () => { value.textContent = slider.value; });
        wrapper.classList.add('map-dial');
        const heading = document.createElement('div'); heading.className = 'map-dial-heading';
        heading.append(label, value);
        wrapper.append(heading, slider);
        container.append(wrapper);
      }
    };
    makeDials(form.querySelector('.map-content-dials'), ['detail', 'clutter', 'wear']);
    makeDials(form.querySelector('.map-look-dials'), ['color', 'tint', 'contrast', 'sharpness']);
    const update = () => {
      for (const kind of ['genre', 'style']) field(`${kind}Other`).hidden = field(kind).value !== 'other';
      const perspective = field('perspective').value;
      form.querySelector('.map-walls-option').hidden = perspective !== 'top-down';
      const summary = costSummary({ quality: field('quality').value, mockups: Number(field('mockups').value), walls: field('walls').checked, perspective });
      form.querySelector('.map-cost').textContent = summary.text;
      form.querySelector('.map-mockups').textContent = `${this.mockups.length ? 'Regenerate' : 'Generate mockups'} (${summary.mockups} ${summary.mockups === 1 ? 'use' : 'uses'})`;
      form.querySelector('.map-final').textContent = `Make final map (${summary.final + (field('walls').checked && perspective === 'top-down' ? 1 : 0)} uses)`;
      const signed = !!this.app?.authenticated && !this.app?._isOffline;
      form.querySelector('.map-mockups').disabled = this.busy || !signed;
      form.querySelector('.map-plan').disabled = this.busy || !signed;
      form.querySelector('.map-final').disabled = this.busy || !signed || this.selected < 0;
      form.querySelector('.map-create').disabled = this.busy || !this.final;
    };
    this.update = update;
    // The layout plan belongs to one description, so it is kept for this session only.
    const remember = () => { this.lastValues = this.snapshot(); writeStore(FORM_KEY, { ...this.lastValues, layout: '' }); };
    this.remember = remember;
    form.addEventListener('input', () => { update(); remember(); });
    form.addEventListener('change', () => { update(); remember(); });
    const description = field('description');
    let settingEdited = false;
    field('setting').addEventListener('change', () => { settingEdited = true; });
    description.addEventListener('input', () => {
      if (!settingEdited) field('setting').value = /\b(town|village|city|street|road|forest|alley|market|camp|wasteland|highway|courtyard|bridge|field|jungle|outpost|main deck|square)\b/i.test(description.value) ? 'outdoor' : 'indoor';
    });
    form.querySelector('.map-mockups').addEventListener('click', () => this.generateMockups());
    form.querySelector('.map-plan').addEventListener('click', () => this.planLayout());
    form.querySelector('.map-final').addEventListener('click', () => this.generateFinal());
    form.querySelector('.map-create').addEventListener('click', () => this.createScene());
    field('sceneWidth').addEventListener('input', () => this.updateMeasured());
    field('sceneHeight').addEventListener('input', () => this.updateMeasured());
    field('showTrace').addEventListener('change', () => this.renderPreview());
    form.querySelector('.map-preview-zoom').addEventListener('click', () => this.zoomFinal());
    const templates = form.querySelector('.map-template-select');
    templates.addEventListener('change', () => {
      const option = templates.selectedOptions[0];
      if (!option?.value) return;
      description.value = option.value;
      if (option.dataset.genre) field('genre').value = option.dataset.genre;
      if (option.dataset.setting) { field('setting').value = option.dataset.setting; settingEdited = true; }
      field('layout').value = '';
      templates.selectedIndex = 0;
      update(); remember();
    });
    this.mountPresets(form);
    this.lastValues ||= readStore(FORM_KEY, null);
    if (this.lastValues) { this.fill(this.lastValues); settingEdited = !!this.lastValues.description; }
    this.renderMockups();
    if (this.final && this.sourceImage) this.showFinal();
    const root = form.closest('.npc-builder');
    if (root) new MutationObserver(update).observe(root, { attributes: true, attributeFilter: ['class'] });
    update();
    queueMicrotask(update);
  }

  values() { return Object.fromEntries(new FormData(this.form)); }
  /** Form values plus the true state of every checkbox (FormData drops unchecked ones). */
  snapshot() {
    const values = this.values();
    for (const name of CHECKBOXES) values[name] = !!this.form.elements[name]?.checked;
    delete values.sceneWidth; delete values.sceneHeight;
    return values;
  }
  /** Put saved values back into the form, ignoring anything the form no longer offers. */
  fill(values, fields = [...MAP_FIELDS, ...CHECKBOXES]) {
    for (const name of fields) {
      const el = this.form.elements[name];
      if (!el || values?.[name] == null) continue;
      if (el.type === 'checkbox') { el.checked = values[name] === true || values[name] === 'on' || values[name] === 'true'; continue; }
      const before = el.value;
      el.value = String(values[name]);
      if (el.tagName === 'SELECT' && el.value !== String(values[name])) el.value = before;
    }
    for (const dial of this.form.querySelectorAll('.map-dial')) dial.querySelector('output').textContent = dial.querySelector('input').value;
    this.update?.();
  }

  /* ── Style presets ─────────────────────────────────────── */

  mountPresets(form) {
    const select = form.querySelector('.map-preset-select');
    const remove = form.querySelector('.map-preset-delete');
    const render = chosen => {
      const mine = readStore(PRESET_KEY, []).filter(p => p && typeof p.name === 'string' && p.values);
      select.replaceChildren(new Option('— Choose a preset… —', ''));
      for (const [label, list, prefix] of [['Built in', BUILT_IN_PRESETS, 'b'], ['Yours', mine, 'u']]) {
        if (!list.length) continue;
        const group = document.createElement('optgroup'); group.label = label;
        list.forEach((preset, index) => group.append(new Option(preset.name, `${prefix}${index}`)));
        select.append(group);
      }
      if (chosen) select.value = chosen;
      remove.disabled = !select.value.startsWith('u');
      return mine;
    };
    render();
    select.addEventListener('change', event => {
      event.stopPropagation();
      const mine = readStore(PRESET_KEY, []);
      const preset = (select.value[0] === 'b' ? BUILT_IN_PRESETS : mine)[Number(select.value.slice(1))];
      remove.disabled = !select.value.startsWith('u');
      if (!preset) return;
      this.fill(preset.values, STYLE_FIELDS);
      this.renderPreview();
      this.remember();
    });
    form.querySelector('.map-preset-save').addEventListener('click', async () => {
      const name = await this.askText('Save style preset', 'Preset name', 40);
      if (!name) return;
      try {
        const mine = upsertPreset(readStore(PRESET_KEY, []), presetFromValues(name, this.snapshot()));
        writeStore(PRESET_KEY, mine);
        render(`u${mine.length - 1}`);
        this.status(`Saved preset “${mine.at(-1).name}”.`);
      } catch (err) { this.error(err.message); }
    });
    remove.addEventListener('click', () => {
      if (!select.value.startsWith('u')) return;
      const mine = readStore(PRESET_KEY, []);
      const [gone] = mine.splice(Number(select.value.slice(1)), 1);
      writeStore(PRESET_KEY, mine);
      render();
      if (gone) this.status(`Deleted preset “${gone.name}”.`);
    });
  }

  async askText(title, label, maxLength) {
    const Dialog = globalThis.foundry?.applications?.api?.DialogV2;
    if (!Dialog?.wait) return (globalThis.prompt?.(label) || '').trim().slice(0, maxLength);
    const field = document.createElement('input');
    field.type = 'text'; field.name = 'value'; field.maxLength = maxLength; field.setAttribute('autofocus', '');
    const wrapper = document.createElement('label');
    wrapper.append(label, field);
    const answer = await Dialog.wait({ window: { title }, content: wrapper.outerHTML, rejectClose: false, buttons: [
      { action: 'save', label: 'Save', default: true, callback: (_event, _button, dialog) => dialog.element.querySelector('[name="value"]')?.value?.trim() || '' },
      { action: 'cancel', label: 'Cancel', callback: () => '' },
    ] });
    return typeof answer === 'string' ? answer.slice(0, maxLength) : '';
  }
  key() {
    const key = this.app?.accessKey || this.app?.storage.getKey() || new Storage(this.module.id).getKey();
    if (!key) throw new Error('Please sign in with Patreon first.');
    return key;
  }
  async request(payload, timeout = 150000) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeout);
    try {
      // This workflow has no -dev webhook, even when the shared app is in dev mode.
      const { response, responseText } = await postToN8n(N8N_ENDPOINTS.mapBuilder, payload, this.key(), { signal: controller.signal });
      let data;
      try { data = JSON.parse(responseText); } catch { throw new Error('The map service returned an invalid response.'); }
      if (!response.ok || data.ok === false) throw new Error(data.message || `Map service returned ${response.status}.`);
      return data;
    } catch (err) {
      if (err.name === 'AbortError') throw new Error(`The map request timed out after ${Math.round(timeout / 1000)} seconds. Please try again.`);
      if (err instanceof AuthError) {
        new Storage(this.module.id).setKey('');
        if (this.app) { this.app.accessKey = ''; this.app.authenticated = false; this.app._applyAuthStateUI(); }
        window.open(PATREON_URL, '_blank');
      }
      throw err;
    } finally { clearTimeout(timer); }
  }

  /** Free actions are safe to send twice, so one dropped connection does not fail the job. */
  async requestFree(payload, timeout) {
    try { return await this.request(payload, timeout); }
    catch (err) {
      if (err instanceof AuthError || !/failed to fetch|networkerror|load failed|invalid response|timed out/i.test(err?.message || '')) throw err;
      return this.request(payload, timeout);
    }
  }

  async run(task) {
    if (this.busy) return;
    this.busy = true; this.error(''); this.update();
    try { await task(); } catch (err) {
      const message = err?.message || String(err);
      this.error(message);
      this.form?.querySelector('.map-error')?.scrollIntoView?.({ block: 'nearest' });
      if (typeof ui !== 'undefined') ui.notifications?.error?.(message);
      this.status('');
    }
    finally { this.busy = false; this.update(); }
  }

  planLayout() { return this.run(async () => {
    const map = buildMapRequest(this.values());
    this.status('Planning the layout…');
    const data = await this.requestFree({ action: 'plan', map });
    if (typeof data.layout !== 'string' || !data.layout.trim()) throw new Error('No layout plan was returned.');
    this.form.elements.layout.value = data.layout;
    this.planned = { description: map.description, layout: buildMapRequest(this.values()).layout };
    this.form.querySelector('.map-layout-plan').open = true;
    this.remember?.();
    this.status('Layout plan ready. You can edit it before drawing.');
  }); }

  generateMockups() { return this.run(async () => {
    const map = buildMapRequest(this.values());
    // A plan made for another description, and not edited since, would draw the old map: plan again.
    if (map.layout && this.planned && map.layout === this.planned.layout && map.description !== this.planned.description && this.form.elements.autoPlan.checked) delete map.layout;
    const count = Number(this.form.elements.mockups.value);
    this.status(`Generating ${count} ${count === 1 ? 'mockup' : 'mockups'}…`);
    const data = await this.request({ action: 'mockups', map, count, ...(!map.layout && this.form.elements.autoPlan.checked ? { autoPlan: true } : {}) });
    if (typeof data.layout === 'string' && data.layout.trim()) {
      this.form.elements.layout.value = data.layout;
      this.planned = { description: map.description, layout: buildMapRequest(this.values()).layout };
      this.form.querySelector('.map-layout-plan').open = true;
    }
    if (!Array.isArray(data.images) || !data.images.length) throw new Error('No mockups were returned.');
    // New mockups go first; earlier sets were paid for too, so they stay pickable.
    const made = { ...map, ...(typeof data.layout === 'string' && data.layout.trim() ? { layout: data.layout.trim().slice(0, 1200) } : {}) };
    const earlier = this.mockups.map(mockup => ({ ...mockup, earlier: true }));
    this.mockups = [...data.images.map(image => ({ b64: image.b64, mediaType: image.mediaType, aspect: data.aspect, map: made })), ...earlier].slice(0, MAX_MOCKUPS);
    this.selected = -1; this.final = null; this.sourceImage = null; this.savedEntry = null;
    this.form.querySelector('.map-result').hidden = true;
    this.renderMockups();
    this.status(data.failed ? `${data.images.length} mockups ready; ${data.failed} failed. Choose one.` : 'Choose a mockup to make the final map.');
  }); }

  renderMockups() {
    const box = this.form.querySelector('.map-thumbnails');
    renderMapThumbnails(box, this.mockups, this.mockups[0]?.aspect, index => { this.selected = index; this.update(); },
      { selected: this.selected, onZoom: index => zoomImage(imageUrl(this.mockups[index]), `Mockup ${index + 1}`) });
    this.form.querySelector('.map-mockup-note').hidden = !this.mockups.some(mockup => mockup.earlier);
  }

  generateFinal() { return this.run(async () => {
    const mockup = this.mockups[this.selected];
    if (!mockup) throw new Error('Choose a mockup first.');
    // An earlier mockup is finished with the description it was drawn from, not whatever the form says now.
    const map = mockup.earlier && mockup.map ? mockup.map : buildMapRequest(this.values());
    const quality = this.form.elements.quality.value;
    const units = game.settings.get(this.module.id, 'units');
    this.finalUnits = units;
    const started = Date.now();
    const tick = () => this.status(`Making final map… ${Math.floor((Date.now() - started) / 1000)} seconds elapsed. This can take several minutes.`);
    tick(); const clock = setInterval(tick, 1000);
    let data;
    try { data = await this.request({ action: 'final', quality, map, mockup: { b64: mockup.b64, mediaType: mockup.mediaType }, units, walls: this.form.elements.walls.checked && map.perspective === 'top-down' }, 330000); }
    finally { clearInterval(clock); }
    if (!data.image || !data.scene) throw new Error('The final map response was incomplete.');
    this.sourceImage = await loadImage(data.image);
    this.final = data;
    this.savedEntry = null;
    this.finalValues = { ...this.snapshot(), description: map.description, ...(map.layout ? { layout: map.layout } : {}) };
    this.showFinal();
    this.status('Final map ready. Adjust its look and create a scene.');
    this.saveFinal();
  }); }

  /** Show this.final in the result panel. */
  showFinal() {
    const data = this.final;
    const warning = this.form.querySelector('.map-warning');
    warning.textContent = data.warning || ''; warning.hidden = !data.warning;
    const result = this.form.querySelector('.map-result'); result.hidden = false;
    const trace = data.trace || {};
    const walls = Array.isArray(trace.walls) ? trace.walls : [];
    const lights = Array.isArray(trace.lights) ? trace.lights : [];
    const wallCount = walls.filter(w => w.kind !== 'door' && w.kind !== 'window').length;
    const doorCount = walls.filter(w => w.kind === 'door').length;
    this.form.querySelector('.map-trace-counts').textContent = `Found ${wallCount} ${plural(wallCount, 'wall', 'walls')}, ${doorCount} ${plural(doorCount, 'door', 'doors')}, ${lights.length} ${plural(lights.length, 'light', 'lights')}`;
    this.form.elements.sceneWidth.value = data.scene.squares[0];
    this.form.elements.sceneHeight.value = data.scene.squares[1];
    this.form.querySelector('.map-trace-toggle').hidden = !walls.length && !lights.length;
    this.updateMeasured();
    this.renderPreview();
    this.showSaved();
    this.update();
  }

  showSaved() {
    const note = this.form.querySelector('.map-saved');
    note.hidden = !this.savedEntry;
    note.textContent = this.savedEntry ? 'Saved to Created Maps. You can reopen it there and create the scene any time, at no extra cost.' : '';
  }

  /**
   * Keep the paid final map: store the image in the world and list it under
   * Created Maps. Runs in the background and never fails the final itself.
   */
  async saveFinal() {
    const final = this.final, image = this.sourceImage, values = this.finalValues;
    try {
      const canvas = imageCanvas(image);
      const blob = await encodeCanvas(canvas, 'image/webp', 0.95) || await encodeCanvas(canvas, 'image/png');
      if (!blob) throw new Error('The map image could not be encoded.');
      const extension = blob.type === 'image/png' ? 'png' : 'webp';
      const filename = sceneFileName(`${values.description}`, new Date(), extension).replace(`.${extension}`, `-original.${extension}`);
      const path = await saveToWorld(new File([blob], filename, { type: blob.type }));
      const id = globalThis.foundry?.utils?.randomID?.(16) || crypto.randomUUID().replaceAll('-', '').slice(0, 16);
      const entry = mapRecord({ values, path, final, units: this.finalUnits, id });
      const app = this.app;
      if (app?.history) {
        app.history.push(entry);
        app.storage?.saveHistory(app.history);
        app._renderHistory?.();
      }
      if (this.final === final) { this.savedEntry = entry; this.showSaved(); }
    } catch (err) {
      console.warn('[AI Battle Map Maker] Could not save the final map to the world:', err);
      if (this.final === final) {
        const note = this.form.querySelector('.map-saved');
        note.hidden = false;
        note.textContent = 'This map could not be saved to Created Maps. Create the scene before closing Foundry.';
      }
    }
  }

  /** Reopen a map from Created Maps so its scene can be made (again) for free. */
  openSaved(entry) { return this.run(async () => {
    this.status('Opening saved map…');
    const saved = entry.map;
    const image = await loadUrl(saved.path, 'The saved map image is missing from the world’s battle-maps folder.');
    this.fill({ ...saved.values, walls: saved.values?.walls });
    this.sourceImage = image;
    this.final = { scene: { squares: saved.squares }, trace: unpackTrace(saved.trace), ...(saved.warning ? { warning: saved.warning } : {}) };
    this.finalUnits = saved.units === 'metric' ? 'metric' : 'imperial';
    this.finalValues = this.snapshot();
    this.savedEntry = entry;
    this.selected = -1;
    this.renderMockups();
    this.showFinal();
    this.form.querySelector('.map-result').scrollIntoView?.({ block: 'nearest' });
    this.status('Saved map opened. Adjust its look and create a scene.');
  }); }

  lookDials() { return Object.fromEntries(['color', 'tint', 'contrast', 'sharpness'].map(key => [key, Number(this.form.elements[key].value)])); }
  renderPreview() {
    if (!this.sourceImage) return;
    const canvas = this.form.querySelector('.map-preview');
    const scale = Math.min(1, 900 / this.sourceImage.naturalWidth);
    canvas.width = Math.max(1, Math.round(this.sourceImage.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(this.sourceImage.naturalHeight * scale));
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    ctx.drawImage(this.sourceImage, 0, 0, canvas.width, canvas.height);
    const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height);
    applyLook(pixels.data, canvas.width, canvas.height, this.lookDials());
    ctx.putImageData(pixels, 0, 0);
    if (this.form.elements.showTrace?.checked && this.final?.trace) drawTrace(ctx, this.final.trace, canvas.width, canvas.height);
  }
  async zoomFinal() {
    if (!this.sourceImage) return;
    const canvas = imageCanvas(this.sourceImage);
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height);
    applyLook(pixels.data, canvas.width, canvas.height, this.lookDials());
    ctx.putImageData(pixels, 0, 0);
    if (this.form.elements.showTrace?.checked && this.final?.trace) drawTrace(ctx, this.final.trace, canvas.width, canvas.height);
    const blob = await encodeCanvas(canvas, 'image/webp');
    zoomImage(blob ? URL.createObjectURL(blob) : canvas.toDataURL('image/png'), 'Final map');
  }
  squares() {
    return [Number(this.form.elements.sceneWidth.value), Number(this.form.elements.sceneHeight.value)].map(v =>
      Number.isFinite(v) ? Math.max(1, Math.min(200, Math.round(v))) : 1);
  }
  updateMeasured() {
    if (!this.final) return;
    const system = UNIT_LABELS[this.finalUnits];
    this.form.querySelector('.map-measured').textContent = measuredLabel({ squares: this.squares(), grid: system });
  }

  createScene() { return this.run(async () => {
    if (!this.final) throw new Error('Make a final map first.');
    const description = String(this.finalValues?.description || this.form.elements.description.value).trim();
    const name = description.slice(0, 60) || 'Battle map';
    const mood = this.form.elements.mood.value;
    const progress = startProgress('Creating battle map scene…');
    const step = (text, pct) => { this.status(text); progress.step(text, pct); };
    try {
      step('Fitting the map to the grid…', 0.1);
      const image = this.final.image || pngPayload(this.sourceImage);
      const bake = await this.requestFree({ action: 'bake', image, squares: this.squares(), units: this.finalUnits,
        trace: this.final.trace, mood, name });
      if (!bake.image?.b64 || !bake.scene || !Array.isArray(bake.walls) || !Array.isArray(bake.lights)) throw new Error('The baked map response was incomplete.');
      step('Applying the look…', 0.4);
      const resized = await loadImage(bake.image);
      const canvas = document.createElement('canvas');
      canvas.width = bake.scene.width; canvas.height = bake.scene.height;
      const ctx = canvas.getContext('2d', { willReadFrequently: true });
      ctx.drawImage(resized, 0, 0, canvas.width, canvas.height);
      const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height);
      applyLook(pixels.data, canvas.width, canvas.height, this.lookDials());
      if (bake.sharpen > 0) applyLook(pixels.data, canvas.width, canvas.height, { sharpness: 4 });
      ctx.putImageData(pixels, 0, 0);
      const blob = await encodeCanvas(canvas, 'image/webp') || await encodeCanvas(canvas, 'image/png');
      if (!blob) throw new Error('The map image could not be encoded.');
      const extension = blob.type === 'image/png' ? 'png' : 'webp';
      step('Uploading the map image…', 0.55);
      const path = await saveToWorld(new File([blob], sceneFileName(description, new Date(), extension), { type: blob.type }));
      step('Creating the Foundry scene…', 0.75);
      const weather = pickWeather(this.form.elements.weather?.value || 'auto', description, this.form.elements.setting.value);
      const scene = await createMapScene(sceneData(bake.scene, path, globalThis.game, {
        environment: sceneEnvironment(mood, bake.lights.length),
        ...(hasWeather(weather) ? { weather } : {}),
        folder: await mapSceneFolder(),
      }));
      if (!scene) throw new Error('Foundry did not create the scene.');
      await setSceneBackground(scene, path);
      if (!getSceneBackground(scene)) throw new Error('Foundry did not attach the map image to the scene.');
      if (bake.walls.length) await scene.createEmbeddedDocuments('Wall', bake.walls.map(wallData));
      if (bake.lights.length) await scene.createEmbeddedDocuments('AmbientLight', bake.lights.map(lightData));
      step('Opening the scene…', 0.9);
      await viewMapScene(scene);
      await setSceneThumbnail(scene);
      this.status(`Created scene “${name}”.`);
      notifyMapCreated(`Created battle map scene “${name}”.`);
    } finally { progress.done(); }
  }); }

  async sceneImage(scene) {
    const src = getSceneBackground(scene);
    if (!src) throw new Error('This scene has no background image.');
    const response = await fetch(src);
    if (!response.ok) throw new Error('The scene background could not be loaded.');
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    try {
      const image = await loadUrl(url, 'The scene background could not be decoded.');
      return { image, payload: pngPayload(image) };
    } finally { URL.revokeObjectURL(url); }
  }

  async editRegion() {
    const scene = canvas?.scene;
    const selected = selectRegionDrawing(scene, canvas?.drawings?.controlled || [], game.user?.id);
    const box = drawingBox(selected, scene, canvas?.dimensions);
    if (!box) {
      ui.notifications?.warn?.('Click Draw edit box, drag a box over the area, then click Edit region.');
      ui.controls?.activate({ control: 'drawings', tool: 'rect' });
      return;
    }
    const Dialog = foundry?.applications?.api?.DialogV2;
    if (!Dialog?.wait) { ui.notifications?.error?.('Foundry DialogV2 is unavailable.'); return; }
    const choice = await Dialog.wait({ window: { title: 'AI edit region' }, content: '<label>Change<textarea name="change" maxlength="500" rows="4" style="width:100%" placeholder="The bridge has collapsed and is still smouldering"></textarea></label><label>Quality<select name="quality"><option value="standard">Standard (4 uses)</option><option value="economy">Economy (4 uses)</option><option value="draft">Draft (4 uses)</option><option value="high">High detail (6 uses)</option></select></label>', buttons: [
      { action: 'apply', label: 'Apply', default: true, callback: (_event, _button, dialog) => ({ change: dialog.element.querySelector('[name="change"]')?.value?.trim(), quality: dialog.element.querySelector('[name="quality"]')?.value }) },
      { action: 'cancel', label: 'Cancel', callback: () => null },
    ] });
    if (!choice) return;
    if (!choice.change || choice.change.length > 500) { ui.notifications?.error?.('Enter a change of up to 500 characters.'); return; }
    if (this.editing) { ui.notifications?.warn?.('A region edit is already running.'); return; }
    this.editing = true;
    const progress = startProgress('Editing map region…');
    try {
      progress.step('Reading the map…', 0.05);
      const before = getSceneBackground(scene);
      const { image, payload } = await this.sceneImage(scene);
      progress.step('Drawing the change. This usually takes under a minute…', 0.2);
      const result = await this.request({ action: 'edit', image: payload, box, change: choice.change, quality: choice.quality }, 330000);
      if (!result.patch?.b64 || !result.crop) throw new Error('The region edit response was incomplete.');
      progress.step('Blending the change into the map…', 0.8);
      const patch = await loadImage(result.patch);
      const composed = pasteRegion(image, patch, result.crop);
      const blob = await encodeCanvas(composed, 'image/webp') || await encodeCanvas(composed, 'image/png');
      if (!blob) throw new Error('The edited map could not be encoded.');
      progress.step('Saving the edited map…', 0.9);
      const extension = blob.type === 'image/png' ? 'png' : 'webp';
      const filename = sceneFileName(scene.name || 'battle-map', new Date(), extension).replace(`.${extension}`, `-${crypto.randomUUID().slice(0,8)}.${extension}`);
      const path = await saveToWorld(new File([blob], filename, { type: blob.type }));
      await setSceneBackground(scene, path);
      // The map before this edit stays on disk, so the edit can be undone for free.
      try { await scene.setFlag?.(FLAG_SCOPE, 'editUndo', { src: before, redo: false }); } catch { /* undo is a convenience */ }
      await scene.deleteEmbeddedDocuments('Drawing', [selected?.document?.id || selected.id]);
      ui.notifications?.success?.(`Region edited (${result.usesCharged} uses). Use “Undo or redo last edit” in the AI Battle Map tools to compare.`);
    } catch (err) { ui.notifications?.error?.(err.message); }
    finally { progress.done(); this.editing = false; }
  }

  /** Swap the scene between the map before and after the last region edit. Free. */
  async undoEdit() {
    const scene = canvas?.scene;
    const saved = scene?.getFlag?.(FLAG_SCOPE, 'editUndo');
    if (!saved?.src) { ui.notifications?.warn?.('There is no region edit to undo on this scene.'); return; }
    try {
      const current = getSceneBackground(scene);
      await setSceneBackground(scene, saved.src);
      await scene.setFlag(FLAG_SCOPE, 'editUndo', { src: current, redo: !saved.redo });
      ui.notifications?.info?.(saved.redo ? 'Region edit restored.' : 'Region edit undone. Click again to bring it back.');
    } catch (err) { ui.notifications?.error?.(err.message); }
  }

  async retraceWalls() {
    const scene = canvas?.scene;
    if (!scene) { ui.notifications?.error?.('Open a scene first.'); return; }
    const Dialog = foundry?.applications?.api?.DialogV2;
    if (!Dialog?.confirm) { ui.notifications?.error?.('Foundry DialogV2 is unavailable.'); return; }
    if (!await Dialog.confirm({ window: { title: 'Re-trace walls' }, content: 'Replace all walls and lights on this scene? 1 use' })) return;
    const progress = startProgress('Tracing walls and lights…');
    try {
      progress.step('Reading the map…', 0.1);
      const { payload } = await this.sceneImage(scene);
      progress.step('Tracing walls and lights…', 0.3);
      const result = await this.request({ action: 'retrace', image: payload, ...sceneTraceOptions(scene) }, 330000);
      if (!Array.isArray(result.walls) || !Array.isArray(result.lights)) throw new Error('The re-trace response was incomplete.');
      progress.step('Placing walls and lights…', 0.85);
      await scene.deleteEmbeddedDocuments('Wall', scene.walls.map(w => w.id));
      await scene.deleteEmbeddedDocuments('AmbientLight', scene.lights.map(l => l.id));
      if (result.walls.length) await scene.createEmbeddedDocuments('Wall', result.walls.map(wallData));
      if (result.lights.length) await scene.createEmbeddedDocuments('AmbientLight', result.lights.map(lightData));
      ui.notifications?.success?.(`Re-traced walls and lights (${result.usesCharged} use).`);
    } catch (err) { ui.notifications?.error?.(err.message); }
    finally { progress.done(); }
  }
}
