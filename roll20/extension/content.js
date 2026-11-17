/**
 * Isolated-world content script on the Roll20 editor. Content scripts cannot
 * be ES modules, so this loads the panel module from the extension bundle.
 */
(() => {
  if (window.__cferPanel) return;
  window.__cferPanel = true;
  const api = globalThis.browser ?? globalThis.chrome;
  import(api.runtime.getURL('panel.js'))
    .then(m => m.start())
    .catch(err => console.error('[Cfer] Could not start the panel:', err));
})();
