import { SystemAdapter, postToN8n, AuthError } from './core/adapter.js';
import { N8N_ENDPOINTS, PATREON_URL } from './core/n8n.js';
import { buildMapRequest, costSummary, measuredLabel, sceneFileName, drawingBox, sceneTraceOptions, plural } from './map-logic.js';
import { applyLook } from './core/maps/look.js';
import { filePicker, directoryExists, uploadMap, sceneData, wallData, lightData, createMapScene, viewMapScene, notifyMapCreated } from './compat.js';
import { pasteRegion } from './region-paste.js';
import { Storage } from './core/storage.js';

const imageUrl = image => `data:${image.mediaType || 'image/png'};base64,${image.b64}`;
const loadImage = image => new Promise((resolve, reject) => {
  const img = new Image();
  img.onload = () => resolve(img);
  img.onerror = () => reject(new Error('The map image could not be decoded.'));
  img.src = imageUrl(image);
});
const encodeCanvas = (canvas, type) => new Promise(resolve => canvas.toBlob(resolve, type, type === 'image/webp' ? 0.9 : undefined));
const DIAL_LABELS = { detail: 'Detail', clutter: 'Props', wear: 'Wear', color: 'Color', tint: 'Color temperature', contrast: 'Lighting contrast', sharpness: 'Sharpness' };
const UNIT_LABELS = { imperial: { distance: 5, units: 'ft' }, metric: { distance: 1.5, units: 'm' } };

export function renderMapThumbnails(box, images, aspect, onSelect) {
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
    button.append(img, caption);
    button.addEventListener('click', () => {
      box.querySelectorAll('.map-thumbnail').forEach((node, i) => {
        node.classList.toggle('is-selected', i === index);
        node.querySelector('.map-picked').textContent = i === index ? 'Picked' : '';
      });
      onSelect(index);
    });
    box.append(button);
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
  populateForm(form, entry) { form.elements.description.value = entry.description || ''; }
  async generate() { throw new Error('Use the map mockup and final controls.'); }

  onFormMount(form) {
    this.form = form;
    this.mockups = [];
    this.selected = -1;
    this.final = null;
    this.sourceImage = null;
    this.busy = false;
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
    form.addEventListener('input', update);
    form.addEventListener('change', update);
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
    const root = form.closest('.npc-builder');
    if (root) new MutationObserver(update).observe(root, { attributes: true, attributeFilter: ['class'] });
    update();
    queueMicrotask(update);
  }

  values() { return Object.fromEntries(new FormData(this.form)); }
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

  async run(task) {
    if (this.busy) return;
    this.busy = true; this.error(''); this.update();
    try { await task(); } catch (err) { this.error(err.message); this.status(''); }
    finally { this.busy = false; this.update(); }
  }

  planLayout() { return this.run(async () => {
    const map = buildMapRequest(this.values());
    this.status('Planning the layout…');
    const data = await this.request({ action: 'plan', map });
    if (typeof data.layout !== 'string' || !data.layout.trim()) throw new Error('No layout plan was returned.');
    this.form.elements.layout.value = data.layout;
    this.form.querySelector('.map-layout-plan').open = true;
    this.status('Layout plan ready. You can edit it before drawing.');
  }); }

  generateMockups() { return this.run(async () => {
    const map = buildMapRequest(this.values());
    const count = Number(this.form.elements.mockups.value);
    this.status(`Generating ${count} ${count === 1 ? 'mockup' : 'mockups'}…`);
    this.mockups = []; this.selected = -1; this.final = null; this.sourceImage = null;
    this.form.querySelector('.map-result').hidden = true;
    const box = this.form.querySelector('.map-thumbnails'); box.replaceChildren();
    const data = await this.request({ action: 'mockups', map, count, ...(!map.layout && this.form.elements.autoPlan.checked ? { autoPlan: true } : {}) });
    if (typeof data.layout === 'string' && data.layout.trim()) {
      this.form.elements.layout.value = data.layout;
      this.form.querySelector('.map-layout-plan').open = true;
    }
    if (!Array.isArray(data.images) || !data.images.length) throw new Error('No mockups were returned.');
    this.mockups = data.images;
    renderMapThumbnails(box, data.images, data.aspect, index => { this.selected = index; this.update(); });
    this.status(data.failed ? `${data.images.length} mockups ready; ${data.failed} failed. Choose one.` : 'Choose a mockup to make the final map.');
  }); }

  generateFinal() { return this.run(async () => {
    if (this.selected < 0) throw new Error('Choose a mockup first.');
    const map = buildMapRequest(this.values());
    const quality = this.form.elements.quality.value;
    const units = game.settings.get(this.module.id, 'units');
    this.finalUnits = units;
    const started = Date.now();
    const tick = () => this.status(`Making final map… ${Math.floor((Date.now() - started) / 1000)} seconds elapsed. This can take several minutes.`);
    tick(); const clock = setInterval(tick, 1000);
    let data;
    try { data = await this.request({ action: 'final', quality, map, mockup: this.mockups[this.selected], units, walls: this.form.elements.walls.checked && map.perspective === 'top-down' }, 330000); }
    finally { clearInterval(clock); }
    if (!data.image || !data.scene) throw new Error('The final map response was incomplete.');
    this.sourceImage = await loadImage(data.image);
    this.final = data;
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
    this.updateMeasured();
    this.renderPreview();
    this.status('Final map ready. Adjust its look and create a scene.');
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
    const description = this.form.elements.description.value.trim();
    const name = description.slice(0, 60) || 'Battle map';
    this.status('Fitting image and preparing the scene…');
    const image = this.final.image;
    const bake = await this.request({ action: 'bake', image, squares: this.squares(), units: this.finalUnits,
      trace: this.final.trace, mood: this.form.elements.mood.value, name });
    if (!bake.image?.b64 || !bake.scene || !Array.isArray(bake.walls) || !Array.isArray(bake.lights)) throw new Error('The baked map response was incomplete.');
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
    const filename = sceneFileName(description, new Date(), extension);
    const picker = filePicker();
    if (!picker) throw new Error('Foundry FilePicker is unavailable.');
    const folder = `worlds/${game.world.id}/battle-maps`;
    this.status('Creating battle-maps folder…');
    try { await picker.createDirectory('data', folder); }
    catch (err) { if (!directoryExists(err)) throw err; }
    this.status('Uploading map image…');
    const path = await uploadMap(picker, folder, new File([blob], filename, { type: blob.type }));
    this.status('Creating Foundry scene…');
    const scene = await createMapScene(sceneData(bake.scene, path));
    if (!scene) throw new Error('Foundry did not create the scene.');
    if (bake.walls.length) await scene.createEmbeddedDocuments('Wall', bake.walls.map(wallData));
    if (bake.lights.length) await scene.createEmbeddedDocuments('AmbientLight', bake.lights.map(lightData));
    this.status('Opening scene…'); await viewMapScene(scene);
    this.status(`Created scene “${name}”.`);
    notifyMapCreated(`Created battle map scene “${name}”.`);
  }); }

  async sceneImage(scene) {
    if (!scene?.background?.src) throw new Error('This scene has no background image.');
    const response = await fetch(scene.background.src);
    if (!response.ok) throw new Error('The scene background could not be loaded.');
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    try {
      const image = await new Promise((resolve, reject) => {
        const img = new Image(); img.onload = () => resolve(img); img.onerror = () => reject(new Error('The scene background could not be decoded.')); img.src = url;
      });
      const canvas = document.createElement('canvas');
      canvas.width = image.naturalWidth; canvas.height = image.naturalHeight;
      canvas.getContext('2d').drawImage(image, 0, 0);
      return { image, payload: { b64: canvas.toDataURL('image/png').split(',')[1], mediaType: 'image/png' } };
    } finally { URL.revokeObjectURL(url); }
  }

  async editRegion() {
    const scene = canvas?.scene;
    const selected = canvas?.drawings?.controlled || [];
    const box = selected.length === 1 ? drawingBox(selected[0], scene, canvas.dimensions) : null;
    if (!box) { ui.notifications?.warn?.('Draw a rectangle with the Drawings tool, select it, then click AI edit region.'); return; }
    const Dialog = foundry?.applications?.api?.DialogV2;
    if (!Dialog?.wait) { ui.notifications?.error?.('Foundry DialogV2 is unavailable.'); return; }
    const choice = await Dialog.wait({ window: { title: 'AI edit region' }, content: '<label>Change<textarea name="change" maxlength="500" rows="4" style="width:100%"></textarea></label><label>Quality<select name="quality"><option value="standard">Standard (4 uses)</option><option value="economy">Economy (4 uses)</option><option value="draft">Draft (4 uses)</option><option value="high">High detail (6 uses)</option></select></label>', buttons: [
      { action: 'apply', label: 'Apply', default: true, callback: (_event, _button, dialog) => ({ change: dialog.element.querySelector('[name="change"]')?.value?.trim(), quality: dialog.element.querySelector('[name="quality"]')?.value }) },
      { action: 'cancel', label: 'Cancel', callback: () => null },
    ] });
    if (!choice) return;
    if (!choice.change || choice.change.length > 500) { ui.notifications?.error?.('Enter a change of up to 500 characters.'); return; }
    try {
      ui.notifications?.info?.('Preparing region edit…');
      const { image, payload } = await this.sceneImage(scene);
      ui.notifications?.info?.('Generating region patch…');
      const result = await this.request({ action: 'edit', image: payload, box, change: choice.change, quality: choice.quality }, 330000);
      if (!result.patch?.b64 || !result.crop) throw new Error('The region edit response was incomplete.');
      const patch = await loadImage(result.patch);
      const composed = pasteRegion(image, patch, result.crop);
      const blob = await encodeCanvas(composed, 'image/webp') || await encodeCanvas(composed, 'image/png');
      if (!blob) throw new Error('The edited map could not be encoded.');
      const picker = filePicker(); if (!picker) throw new Error('Foundry FilePicker is unavailable.');
      const folder = `worlds/${game.world.id}/battle-maps`;
      try { await picker.createDirectory('data', folder); } catch (err) { if (!directoryExists(err)) throw err; }
      ui.notifications?.info?.('Uploading edited map…');
      const extension = blob.type === 'image/png' ? 'png' : 'webp';
      const filename = sceneFileName(scene.name || 'battle-map', new Date(), extension).replace(`.${extension}`, `-${crypto.randomUUID().slice(0,8)}.${extension}`);
      const path = await uploadMap(picker, folder, new File([blob], filename, { type: blob.type }));
      await scene.update({ 'background.src': path });
      await scene.deleteEmbeddedDocuments('Drawing', [selected[0].document.id]);
      ui.notifications?.success?.(`Region edited (${result.usesCharged} uses).`);
    } catch (err) { ui.notifications?.error?.(err.message); }
  }

  async retraceWalls() {
    const scene = canvas?.scene;
    if (!scene) { ui.notifications?.error?.('Open a scene first.'); return; }
    const Dialog = foundry?.applications?.api?.DialogV2;
    if (!Dialog?.confirm) { ui.notifications?.error?.('Foundry DialogV2 is unavailable.'); return; }
    if (!await Dialog.confirm({ window: { title: 'Re-trace walls' }, content: 'Replace all walls and lights on this scene? 1 use' })) return;
    try {
      ui.notifications?.info?.('Tracing walls and lights…');
      const { payload } = await this.sceneImage(scene);
      const result = await this.request({ action: 'retrace', image: payload, ...sceneTraceOptions(scene) }, 330000);
      if (!Array.isArray(result.walls) || !Array.isArray(result.lights)) throw new Error('The re-trace response was incomplete.');
      await scene.deleteEmbeddedDocuments('Wall', scene.walls.map(w => w.id));
      await scene.deleteEmbeddedDocuments('AmbientLight', scene.lights.map(l => l.id));
      if (result.walls.length) await scene.createEmbeddedDocuments('Wall', result.walls.map(wallData));
      if (result.lights.length) await scene.createEmbeddedDocuments('AmbientLight', result.lights.map(lightData));
      ui.notifications?.success?.(`Re-traced walls and lights (${result.usesCharged} use).`);
    } catch (err) { ui.notifications?.error?.(err.message); }
  }
}
