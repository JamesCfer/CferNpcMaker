import { openBuilder } from './core/app.js';
import { checkForModuleUpdate } from './core/update-check.js';
import { registerSidebar } from './core/sidebar.js';
import { startHeartbeat } from './core/heartbeat.js';
import { Storage } from './core/storage.js';
import { BattleMapAdapter } from './adapter.js';
import { loadMapTemplates, registerMapSceneControls } from './compat.js';

const adapter = new BattleMapAdapter();
const ID = adapter.module.id;
const open = () => {
  adapter.app = openBuilder(adapter);
  checkForModuleUpdate(ID, adapter.module.githubUrl).catch(() => {});
};

registerSidebar(ID, open, { buttonLabel: 'AI Battle Map Maker', buttonIcon: adapter.module.icon,
  directories: ['scenes'] });
Hooks.on('getSceneControlButtons', controls => registerMapSceneControls(controls, () => adapter.editRegion(), () => adapter.retraceWalls()));

Hooks.once('init', () => {
  game.settings.register(ID, 'devMode', { name: 'Developer Mode', hint: 'Other shared service calls use development endpoints. Map generation always uses the production webhook.', scope: 'world', config: true, type: Boolean, default: false });
  game.settings.register(ID, 'units', { name: 'Battle map units', hint: 'Imperial uses 5 ft squares; metric uses 1.5 m squares.', scope: 'world', config: true,
    type: String, choices: { imperial: 'Imperial (feet)', metric: 'Metric (metres)' }, default: 'imperial' });
  game.settings.register(ID, 'welcomeMessageShown', { scope: 'world', config: false, type: Boolean, default: false });
});

Hooks.once('ready', () => {
  const version = game.modules?.get(ID)?.version || '';
  const storage = new Storage(ID);
  if (version && storage.getVersion() && storage.getVersion() !== version) {
    storage.setKey('');
    ui.notifications?.info?.('AI Battle Map Maker was updated, please sign in again.');
  }
  if (version) storage.setVersion(version);
  loadMapTemplates([`modules/${ID}/templates/builder.html`]);
  startHeartbeat(ID);
  if (game.user.isGM && !game.settings.get(ID, 'welcomeMessageShown')) {
    ChatMessage.create({ content: '<h3>Welcome to AI Battle Map Maker!</h3><p>Create four mockups for 3 uses, then make a final map for 2 to 6 uses. Open the maker from the Scenes sidebar.</p>',
      whisper: game.users.filter(user => user.isGM).map(user => user.id) });
    game.settings.set(ID, 'welcomeMessageShown', true);
    open();
  }
});
