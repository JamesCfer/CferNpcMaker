/**
 * Foundry compatibility boundary. Source of truth available offline: this repo's
 * shared/scripts/core/app.js (ApplicationV2 and Handlebars mixin),
 * shared/scripts/core/sidebar.js (directory hooks), and the v13-shaped bake
 * payload in tools/n8n/build-map-workflow.mjs. The v14 Scene, Wall, and
 * AmbientLight schemas and FilePicker return shape are unverified offline.
 */
export const generation = gameObject => Number(gameObject?.release?.generation || 13);

export function filePicker(foundryObject = globalThis.foundry, legacy = globalThis.FilePicker) {
  return foundryObject?.applications?.apps?.FilePicker?.implementation || legacy;
}

export const directoryExists = error => /already exists|EEXIST/i.test(String(error?.message || error));

export async function uploadMap(picker, folder, file) {
  const result = await picker.upload('data', folder, file, {});
  if (!result?.path) throw new Error('Foundry did not return the uploaded map path.');
  return result.path;
}

export function sceneData(baked, path, gameObject = globalThis.game) {
  const scene = { ...baked, background: { src: path } };
  // Both supported generations use the v13 environment field in the offline
  // payload. Keep this boundary for a verified v14 mapping if one is needed.
  if (generation(gameObject) >= 14 && baked.environment?.darknessLevel != null) {
    scene.environment = { ...baked.environment, darknessLevel: baked.environment.darknessLevel };
  }
  return scene;
}

export const wallData = wall => ({ c: wall.c, door: wall.door, move: wall.move, sight: wall.sight, light: wall.light, sound: wall.sound });
export const lightData = light => ({ x: light.x, y: light.y, config: { bright: light.config.bright, dim: light.config.dim, color: light.config.color, alpha: light.config.alpha, animation: light.config.animation } });

export function loadMapTemplates(paths, foundryObject = globalThis.foundry, legacy = globalThis.loadTemplates) {
  const load = foundryObject?.applications?.handlebars?.loadTemplates || legacy;
  if (!load) throw new Error('Foundry template loader is unavailable.');
  return load(paths);
}

export async function createMapScene(data, config = globalThis.CONFIG, legacy = globalThis.Scene) {
  const SceneClass = config?.Scene?.documentClass || legacy;
  if (!SceneClass?.create) throw new Error('Foundry Scene document class is unavailable.');
  return SceneClass.create(data);
}

export async function viewMapScene(scene) {
  if (typeof scene?.view !== 'function') throw new Error('Foundry Scene view is unavailable.');
  return scene.view();
}

export function notifyMapCreated(message, uiObject = globalThis.ui) {
  uiObject?.notifications?.success?.(message);
}

export function registerMapSceneControls(controls, edit, retrace, isGM = globalThis.game?.user?.isGM) {
  if (!isGM || !controls || typeof controls !== 'object') return false;
  const groups = Array.isArray(controls) ? controls : Object.values(controls);
  const drawings = controls.drawings || groups.find(group => group?.name === 'drawings');
  const walls = controls.walls || groups.find(group => group?.name === 'walls');
  const add = (group, name, title, icon, handler) => {
    if (!group?.tools) return false;
    const tool = { name, title, icon, button: true, visible: true, onChange: active => { if (active !== false) handler(); } };
    if (Array.isArray(group.tools)) { if (!group.tools.some(t => t.name === name)) group.tools.push(tool); return true; }
    if (typeof group.tools === 'object') { group.tools[name] = tool; return true; }
    return false;
  };
  const a = add(drawings, 'ai-edit-region', 'AI edit region', 'fa-solid fa-wand-magic-sparkles', edit);
  const b = add(walls, 'ai-retrace-walls', 'Re-trace walls (1 use)', 'fa-solid fa-draw-polygon', retrace);
  return a || b;
}
