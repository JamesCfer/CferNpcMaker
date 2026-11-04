import { SystemAdapter, postToN8n, AuthError } from './core/adapter.js';
import { N8N_ENDPOINTS, PATREON_URL } from './core/n8n.js';
import { DIALS } from './core/maps/prompt.js';
import { applyLook } from './core/maps/look.js';
import { gridFit, UNIT_SYSTEMS } from './core/maps/foundry.js';
import { buildMapRequest, buildSceneData, costSummary, measuredLabel, sceneFileName } from './map-logic.js';

const imageUrl = image => `data:${image.mediaType || 'image/png'};base64,${image.b64}`;
const loadImage = image => new Promise((resolve, reject) => {
  const img = new Image();
  img.onload = () => resolve(img);
  img.onerror = () => reject(new Error('The returned image could not be opened.'));
  img.src = imageUrl(image);
});
const canvasBlob = (canvas, type) => new Promise(resolve => canvas.toBlob(resolve, type, .9));

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
        label.textContent = DIALS[key].label;
        const slider = document.createElement('input');
        slider.type = 'range'; slider.name = key; slider.min = '1'; slider.max = '5'; slider.step = '1'; slider.value = '3';
        const value = document.createElement('output'); value.textContent = '3';
        slider.addEventListener('input', () => { value.textContent = slider.value; if (DIALS[key].stage === 'filter') this.renderPreview().catch(e => error(e.message)); });
        wrapper.append(label, slider, value);
        container.append(wrapper);
      }
    };
    makeDials(form.querySelector('.map-content-dials'), ['detail', 'clutter', 'wear']);
    makeDials(form.querySelector('.map-look-dials'), ['color', 'tint', 'contrast', 'sharpness']);
    const update = () => {
      for (const kind of ['genre', 'style']) field(`${kind}Other`).hidden = field(kind).value !== 'other';
      const summary = costSummary({ quality: field('quality').value });
      form.querySelector('.map-cost').textContent = summary.text;
      form.querySelector('.map-final').textContent = `Make final map (${summary.final} uses)`;
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
      if (err.name === 'AbortError') throw new Error('The map request timed out after 150 seconds. Please try again.');
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
    this.mockups = []; this.selected = -1; this.final = null;
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
    const tick = () => this.status(`Making final map… ${Math.floor((Date.now() - started) / 1000)} seconds elapsed. This can take about 100 seconds.`);
    tick(); const clock = setInterval(tick, 1000);
    let data;
    try { data = await this.request({ action: 'final', quality, map, mockup: this.mockups[this.selected], units }); }
    finally { clearInterval(clock); }
    if (!data.image || !data.scene) throw new Error('The final map response was incomplete.');
    this.final = data;
    this.fullImage = await loadImage(data.image);
    const warning = this.form.querySelector('.map-warning');
    warning.textContent = data.warning || ''; warning.hidden = !data.warning;
    const result = this.form.querySelector('.map-result'); result.hidden = false;
    this.form.elements.sceneWidth.value = data.scene.squares[0];
    this.form.elements.sceneHeight.value = data.scene.squares[1];
    this.updateMeasured();
    await this.renderPreview();
    this.status('Final map ready. Adjust its look and create a scene.');
  }); }

  lookDials() { return Object.fromEntries(['color', 'tint', 'contrast', 'sharpness'].map(key => [key, Number(this.form.elements[key].value)])); }
  async renderPreview() {
    if (!this.fullImage) return;
    const img = this.fullImage; const scale = Math.min(1, 900 / img.width);
    const canvas = this.form.querySelector('.map-preview');
    canvas.width = Math.round(img.width * scale); canvas.height = Math.round(img.height * scale);
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    ctx.imageSmoothingQuality = 'high'; ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
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
    const system = UNIT_SYSTEMS[this.finalUnits];
    this.form.querySelector('.map-measured').textContent = measuredLabel({ squares: this.squares(), grid: { distance: system.square, units: system.units } });
  }

  createScene() { return this.run(async () => {
    if (!this.final || !this.fullImage) throw new Error('Make a final map first.');
    const img = this.fullImage, squares = this.squares();
    const fit = gridFit(img.width, img.height, squares);
    this.status('Applying look at full resolution…');
    const source = document.createElement('canvas'); source.width = img.width; source.height = img.height;
    const srcCtx = source.getContext('2d', { willReadFrequently: true }); srcCtx.drawImage(img, 0, 0);
    const original = srcCtx.getImageData(0, 0, img.width, img.height);
    applyLook(original.data, img.width, img.height, this.lookDials()); srcCtx.putImageData(original, 0, 0);
    this.status('Fitting image to the scene grid…');
    const canvas = document.createElement('canvas'); canvas.width = fit.width; canvas.height = fit.height;
    const ctx = canvas.getContext('2d', { willReadFrequently: !!fit.sharpen }); ctx.imageSmoothingEnabled = true; ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(source, 0, 0, fit.width, fit.height);
    if (fit.sharpen) { const pixels = ctx.getImageData(0, 0, fit.width, fit.height); applyLook(pixels.data, fit.width, fit.height, { sharpness: 4 }); ctx.putImageData(pixels, 0, 0); }
    let blob = await canvasBlob(canvas, 'image/webp'); let extension = 'webp';
    if (!blob || blob.type !== 'image/webp') { blob = await canvasBlob(canvas, 'image/png'); extension = 'png'; }
    if (!blob) throw new Error('Could not encode the map image.');
    const description = this.form.elements.description.value.trim();
    const name = description.slice(0, 60) || 'Battle map';
    const filename = sceneFileName(description, new Date(), extension);
    const picker = foundry.applications.apps.FilePicker?.implementation || globalThis.FilePicker;
    if (!picker) throw new Error('Foundry FilePicker is unavailable.');
    const folder = `worlds/${game.world.id}/battle-maps`;
    this.status('Creating battle-maps folder…');
    try { await picker.createDirectory('data', folder); }
    catch (err) { if (!/already exists|EEXIST/i.test(String(err?.message || err))) throw err; }
    this.status('Uploading map image…');
    const upload = await picker.upload('data', folder, new File([blob], filename, { type: blob.type }), {});
    const path = upload?.path || `${folder}/${filename}`;
    this.status('Creating Foundry scene…');
    const data = buildSceneData({ name, src: path, squares, units: this.finalUnits,
      imageSize: { width: img.width, height: img.height } });
    const scene = await Scene.create(data);
    if (!scene) throw new Error('Foundry did not create the scene.');
    this.status('Opening scene…'); await scene.view();
    this.status(`Created scene “${name}”.`);
    ui.notifications.success(`Created battle map scene “${name}”.`);
  }); }
}
