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

function sceneLevel(scene) {
  const levels = scene?.levels;
  if (!levels || (typeof levels.get !== 'function' && !Array.isArray(levels.contents))) return null;
  const initialId = scene.initialLevel?.id || scene.initialLevel;
  const level = (typeof levels.get === 'function' && initialId ? levels.get(initialId) : null)
    || levels.contents?.find(item => item.id === initialId || item._id === initialId)
    || levels.contents?.[0];
  return { level, id: level?.id || level?._id || initialId || 'defaultLevel0000' };
}

export function getSceneBackground(scene) {
  const target = sceneLevel(scene);
  return target ? target.level?.background?.src : scene?.background?.src;
}

export async function setSceneBackground(scene, src) {
  const target = sceneLevel(scene);
  if (!target) return scene.update({ 'background.src': src });
  if (typeof target.level?.update === 'function') return target.level.update({ 'background.src': src });
  return scene.updateEmbeddedDocuments('Level', [{ _id: target.id, 'background.src': src }]);
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

export function registerMapSceneControls(controls, edit, retrace, open, isGM = globalThis.game?.user?.isGM) {
  if (!isGM || !controls || typeof controls !== 'object') return false;
  const legacy = Array.isArray(controls);
  const groups = legacy ? controls : Object.values(controls);
  const order = Math.max(-1, ...groups.filter(group => group?.name !== 'ai-battle-map').map((group, index) =>
    Number.isFinite(group?.order) ? group.order : index)) + 1;
  const button = (name, title, icon, order, handler) => ({
    name, title, icon, order, button: true,
    onChange: active => { if (active !== false) handler(); },
  });
  const tools = [
    { name: 'ai-pointer', title: 'AI Battle Map tools', icon: 'fa-solid fa-map', order: 0 },
    button('ai-draw-box', 'Draw edit box', 'fa-regular fa-square', 1, () => {
      globalThis.ui?.controls?.activate({ control: 'drawings', tool: 'rect' });
      globalThis.ui?.notifications?.info('Drag a box over the area you want to change, then click Edit region.');
    }),
    button('ai-edit-region', 'Edit region (4 uses)', 'fa-solid fa-wand-magic-sparkles', 2, edit),
    button('ai-retrace-walls', 'Re-trace walls (1 use)', 'fa-solid fa-draw-polygon', 3, retrace),
    button('ai-open-builder', 'Open AI Battle Map Maker', 'fa-solid fa-map-location-dot', 4, open),
  ];
  const group = { name: 'ai-battle-map', title: 'AI Battle Map', icon: 'fa-solid fa-map', layer: 'drawings',
    order, activeTool: 'ai-pointer', visible: true,
    onChange: () => {},
    tools: legacy ? tools : Object.fromEntries(tools.map(tool => [tool.name, tool])),
  };
  const drawings = legacy ? controls.find(existing => existing?.name === 'drawings') : controls.drawings;
  if (drawings?.tools) {
    const existingTools = Array.isArray(drawings.tools) ? drawings.tools : Object.values(drawings.tools);
    const lastOrder = Math.max(-1, ...existingTools.filter(tool => tool?.name !== 'ai-edit-region')
      .map((tool, index) => Number.isFinite(tool?.order) ? tool.order : index));
    const editButton = button('ai-edit-region', 'Edit region (4 uses)', 'fa-solid fa-wand-magic-sparkles',
      lastOrder + 1, edit);
    if (Array.isArray(drawings.tools)) {
      const index = drawings.tools.findIndex(tool => tool?.name === editButton.name);
      if (index < 0) drawings.tools.push(editButton);
      else drawings.tools[index] = editButton;
    } else drawings.tools[editButton.name] = editButton;
  }
  if (legacy) {
    const index = controls.findIndex(existing => existing?.name === group.name);
    if (index < 0) controls.push(group);
    else controls[index] = group;
  } else controls[group.name] = group;
  return true;
}
