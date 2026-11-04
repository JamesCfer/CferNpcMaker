import { SystemAdapter, postToN8n, AuthError } from './core/adapter.js';
import { N8N_ENDPOINTS, PATREON_URL } from './core/n8n.js';
import { buildMapRequest, costSummary, measuredLabel, sceneFileName } from './map-logic.js';
import { applyLook } from './core/maps/look.js';
import { filePicker, directoryExists, uploadMap, sceneData, wallData, lightData, createMapScene, viewMapScene, notifyMapCreated } from './compat.js';

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

export class BattleMapAdapter extends SystemAdapter {
  get moduleFolder() { return 'ai-battle-map-maker'; }
  get module() { return { id: 'ai-battle-map-maker', label: 'AI Battle Map Maker', icon: 'fa-solid fa-map',
    githubUrl: 'https://github.com/JamesCfer/CferNpcMaker', historyLabel: 'Created Maps' }; }
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
        const value = document.createElement('output'); value.textContent = '3';
        slider.addEventListener('input', () => { value.textContent = slider.value; if (['color', 'tint', 'contrast', 'sharpness'].includes(key)) this.renderPreview(); });
        wrapper.append(label, slider, value);
        container.append(wrapper);
      }
    };
    makeDials(form.querySelector('.map-content-dials'), ['detail', 'clutter', 'wear']);
    makeDials(form.querySelector('.map-look-dials'), ['color', 'tint', 'contrast', 'sharpness']);
    const update = () => {
      for (const kind of ['genre', 'style']) field(`${kind}Other`).hidden = field(kind).value !== 'other';
      const perspective = field('perspective').value;
      form.querySelector('.map-walls-option').hidden = perspective !== 'top-down';
      const summary = costSummary({ quality: field('quality').value, walls: field('walls').checked, perspective });
      form.querySelector('.map-cost').textContent = summary.text;
      form.querySelector('.map-final').textContent = `Make final map (${summary.final + (field('walls').checked && perspective === 'top-down' ? 1 : 0)} uses)`;
      const signed = !!this.app?.authenticated && !this.app?._isOffline;
      form.querySelector('.map-mockups').disabled = this.busy || !signed;
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
    const key = this.app?.accessKey || this.app?.storage.getKey();
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
        this.app.storage.setKey(''); this.app.accessKey = ''; this.app.authenticated = false;
        this.app._applyAuthStateUI();
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

  generateMockups() { return this.run(async () => {
    const map = buildMapRequest(this.values());
    this.status('Generating four mockups…');
    this.mockups = []; this.selected = -1; this.final = null; this.sourceImage = null;
    this.form.querySelector('.map-result').hidden = true;
    const box = this.form.querySelector('.map-thumbnails'); box.replaceChildren();
    const data = await this.request({ action: 'mockups', map });
    if (!Array.isArray(data.images) || !data.images.length) throw new Error('No mockups were returned.');
    this.mockups = data.images;
    data.images.forEach((image, index) => {
      const button = document.createElement('button'); button.type = 'button'; button.className = 'map-thumbnail';
      button.setAttribute('aria-label', `Select mockup ${index + 1}`);
      const img = document.createElement('img'); img.src = imageUrl(image); img.alt = `Mockup ${index + 1}`;
      button.append(img);
      button.addEventListener('click', () => {
        this.selected = index;
        box.querySelectorAll('.map-thumbnail').forEach((node, i) => node.classList.toggle('is-selected', i === index));
        this.update();
      });
      box.append(button);
    });
    this.form.querySelector('.map-mockups').textContent = 'Regenerate (3 uses)';
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
    try { data = await this.request({ action: 'final', quality, map, mockup: this.mockups[this.selected], units, walls: this.form.elements.walls.checked && map.perspective === 'top-down' }, 900000); }
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
    this.form.querySelector('.map-trace-counts').textContent = `Found ${walls.filter(w => w.kind !== 'door' && w.kind !== 'window').length} walls, ${walls.filter(w => w.kind === 'door').length} doors, ${lights.length} lights`;
    this.form.elements.sceneWidth.value = data.scene.squares[0];
    this.form.elements.sceneHeight.value = data.scene.squares[1];
    this.updateMeasured();
    this.renderPreview();
    this.status('Final map ready. Adjust its look and create a scene.');
  }); }

  lookDials() { return Object.fromEntries(['color', 'tint', 'contrast', 'sharpness'].map(key => [key, Number(this.form.elements[key].value)])); }
  renderPreview() {
    if (!this.sourceImage) return;
    const canvas = document.createElement('canvas');
    const scale = Math.min(1, 900 / this.sourceImage.naturalWidth);
    canvas.width = Math.max(1, Math.round(this.sourceImage.naturalWidth * scale));
    canvas.height = Math.max(1, Math.round(this.sourceImage.naturalHeight * scale));
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    ctx.drawImage(this.sourceImage, 0, 0, canvas.width, canvas.height);
    const pixels = ctx.getImageData(0, 0, canvas.width, canvas.height);
    applyLook(pixels.data, canvas.width, canvas.height, this.lookDials());
    ctx.putImageData(pixels, 0, 0);
    this.form.querySelector('.map-preview').src = canvas.toDataURL('image/png');
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
}
