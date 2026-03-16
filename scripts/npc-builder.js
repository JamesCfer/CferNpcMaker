/*
 * PF2E NPC Auto-Builder — UNIFIED VERSION (ApplicationV2)
 *
 * Features:
 * - Patreon OAuth authentication
 * - Tier-based rate limiting (monthly NPC limits)
 * - Spell ID mapping for proper spell linking
 * - Full error handling and validation
 * - Automatic retry on validation errors
 * - Sidebar buttons and header controls
 * - Bulk / concurrent NPC generation
 * - Local history panel (stored in localStorage)
 *
 * Authentication Flow:
 * - Click "Sign in with Patreon" → opens popup to n8n login endpoint
 * - n8n redirects to Patreon; on success the callback window postMessages:
 *     { type:'patreon-auth', ok:true, key:'<32+ char session key>' }
 * - We accept the message (from the popup window), store key, enable UI
 *
 * Generation Flow:
 * - When generating an NPC, we POST to n8n /webhook/npc-builder with:
 *     headers: { 'X-Builder-Key': <key>, 'X-Foundry-Origin': window.location.origin }
 *     body: { name, level, description, spellMapping (optional) }
 * - The server re-validates key + origin and runs the full generation pipeline
 * - Multiple NPCs can be generated concurrently; each gets its own history entry
 *
 * Rate Limiting:
 * - Free tier: 3 NPCs/month
 * - Local Adventurer: 15 NPCs/month
 * - Standard: 50 NPCs/month
 * - Champion: 80 NPCs/month
 */

const { HandlebarsApplicationMixin, ApplicationV2 } = foundry.applications.api;

class NPCBuilderApp extends HandlebarsApplicationMixin(ApplicationV2) {
  /** n8n endpoints */
  static N8N_AUTH_URL   = 'https://foundryrelay.dedicated2.com/webhook/oauth/patreon/login';
  static N8N_NPC_URL    = 'https://foundryrelay.dedicated2.com/webhook/npc-builder';
  static N8N_DND5E_URL  = 'https://foundryrelay.dedicated2.com/webhook/dnd5e-npc-builder';
  static N8N_HERO6E_URL = 'https://foundryrelay.dedicated2.com/webhook/hero6e-npc-builder';
  static PATREON_URL    = 'https://www.patreon.com/cw/CelestiaTools';

  /** localStorage slots */
  static STORAGE_KEYS = ['pf2e-npc-builder.key', 'pf2e-npc-builder:key'];

  /** localStorage slot for NPC history */
  static HISTORY_KEY = 'pf2e-npc-builder.history';

  /** localStorage slot for selected game system */
  static SYSTEM_KEY = 'pf2e-npc-builder.system';

  /** localStorage slot for last-seen module version (used to force sign-out on updates) */
  static VERSION_KEY = 'pf2e-npc-builder.module-version';

  /** Max history entries to retain */
  static MAX_HISTORY = 50;

  /** Supported game systems */
  static SYSTEMS = ['pf2e', 'dnd5e', 'hero6e'];

  /**
   * Valid Hero System 6e point tiers.
   * Used to snap the input value to the nearest recognised budget.
   */
  static HERO6E_POINT_TIERS = [25, 50, 75, 100, 150, 175, 200, 250, 300, 350, 400, 500, 600];

  /**
   * Valid Hero System 6e genre values accepted by the n8n workflow.
   */
  static HERO6E_GENRES = ['standard', 'superhero', 'pulp', 'dark_champions', 'fantasy', 'sci-fi'];

  static DEFAULT_OPTIONS = {
    id: 'pf2e-npc-builder',
    classes: ['pf2e', 'npc-builder'],
    window: {
      title: 'NPC Builder',
      resizable: true,
    },
    position: {
      width: 800,
    },
    actions: {
      signin:        function(event) { this._signIn(event); },
      signout:       function(event) { this._signOut(event); },
      generate:      function(event) { this._generateNPC(event); },
      export:        function(event) { this._exportJSON(event); },
      patreon:       function()      { window.open(this.constructor.PATREON_URL, '_blank'); },
      selectsystem:  function(event) { this._selectSystem(event); },
    },
  };

  static get PARTS() {
    const modId = game.modules?.get('Pf2eNpcMaker') ? 'Pf2eNpcMaker' : 'pf2e-npc-auto-builder';
    return {
      form: { template: `modules/${modId}/templates/builder.html` },
    };
  }

  /* ── Key storage helpers ─────────────────────────────────── */

  static getStoredKey() {
    for (const k of NPCBuilderApp.STORAGE_KEYS) {
      try {
        const v = localStorage.getItem(k);
        if (v) return v;
      } catch (_) {}
    }
    return '';
  }

  static setStoredKey(value) {
    try {
      if (value) {
        for (const k of NPCBuilderApp.STORAGE_KEYS) localStorage.setItem(k, value);
      } else {
        for (const k of NPCBuilderApp.STORAGE_KEYS) localStorage.removeItem(k);
      }
    } catch (_) {}
  }

  /* ── System storage helpers ──────────────────────────────── */

  static getStoredSystem() {
    try {
      const v = localStorage.getItem(NPCBuilderApp.SYSTEM_KEY);
      if (v && NPCBuilderApp.SYSTEMS.includes(v)) return v;
    } catch (_) {}
    // Default to whatever system is active in Foundry
    try {
      const gameSystem = game?.system?.id;
      if (gameSystem && NPCBuilderApp.SYSTEMS.includes(gameSystem)) return gameSystem;
    } catch (_) {}
    return 'pf2e';
  }

  static setStoredSystem(system) {
    try { localStorage.setItem(NPCBuilderApp.SYSTEM_KEY, system); } catch (_) {}
  }

  /* ── Module version storage helpers ─────────────────────── */

  static getStoredVersion() {
    try { return localStorage.getItem(NPCBuilderApp.VERSION_KEY) || ''; } catch (_) { return ''; }
  }

  static setStoredVersion(version) {
    try { localStorage.setItem(NPCBuilderApp.VERSION_KEY, version); } catch (_) {}
  }

  /* ── History storage helpers ─────────────────────────────── */

  static loadHistory() {
    try {
      const raw = localStorage.getItem(NPCBuilderApp.HISTORY_KEY);
      if (raw) return JSON.parse(raw);
    } catch (_) {}
    return [];
  }

  static saveHistory(history) {
    try {
      const trimmed = history.slice(-NPCBuilderApp.MAX_HISTORY);
      localStorage.setItem(NPCBuilderApp.HISTORY_KEY, JSON.stringify(trimmed));
    } catch (_) {}
  }

  /* ── Constructor ─────────────────────────────────────────── */

  constructor(options = {}) {
    super(options);
    this.accessKey         = NPCBuilderApp.getStoredKey() || '';
    this.authenticated     = !!this.accessKey;
    this.lastGeneratedNPC  = null;
    this.selectedHistoryId = null;
    this.selectedSystem    = NPCBuilderApp.getStoredSystem();

    // Load history; clean up any entries stuck in "generating" from a prior session
    this.npcHistory = NPCBuilderApp.loadHistory();
    let hadStale = false;
    for (const entry of this.npcHistory) {
      if (entry.status === 'generating') {
        entry.status = 'error';
        entry.error  = 'Session was interrupted';
        hadStale = true;
      }
    }
    if (hadStale) NPCBuilderApp.saveHistory(this.npcHistory);
  }

  /* ── Template data ───────────────────────────────────────── */

  async _prepareContext(options) {
    return {
      authenticated: this.authenticated,
      patreonUrl:    NPCBuilderApp.PATREON_URL,
    };
  }

  /* ── Render hook ─────────────────────────────────────────── */

  _onRender(context, options) {
    // Bind tab clicks directly — bypasses ApplicationV2 action delegation
    // which can be intercepted by system-level CSS/JS (e.g. PF2e overrides).
    this.element.querySelectorAll('.system-tab[data-system]').forEach(btn => {
      btn.addEventListener('click', ev => {
        ev.preventDefault();
        ev.stopPropagation();
        this._selectSystem(ev);
      });
    });

    this._applySystemUI();  // also calls _applyAuthStateUI
    this._renderHistory();
  }

  /* ── Auth state UI ───────────────────────────────────────── */

  _applyAuthStateUI() {
    const root = this.element;
    if (!root) return;

    root.classList.toggle('is-authenticated', !!this.authenticated);

    const anyGenerating = this.npcHistory.some(e => e.status === 'generating');
    root.classList.toggle('is-generating', anyGenerating);

    const genBtn = root.querySelector('button[data-action="generate"]');
    if (genBtn) {
      genBtn.disabled = !this.authenticated;
      const label = genBtn.querySelector('.btn-label');
      if (label) label.textContent = 'Generate NPC';
    }

    const expBtn = root.querySelector('button[data-action="export"]');
    if (expBtn) expBtn.disabled = !this.authenticated;
  }

  /* ── System selection ────────────────────────────────────── */

  _selectSystem(event) {
    const btn    = event.currentTarget || event.target;
    const system = btn?.dataset?.system;
    if (!system) return;

    if (system === 'home') {
      this.selectedSystem = 'home';
      // Don't persist 'home' — remember the last real system for next open
    } else if (NPCBuilderApp.SYSTEMS.includes(system)) {
      this.selectedSystem = system;
      NPCBuilderApp.setStoredSystem(system);
    } else {
      return;
    }
    this._applySystemUI();
  }

  _applySystemUI() {
    const root = this.element;
    if (!root) return;

    const system = this.selectedSystem || 'pf2e';

    // Update system tab active state (includes 'home')
    root.querySelectorAll('.system-tab').forEach(btn => {
      btn.classList.toggle('is-active', btn.dataset.system === system);
    });

    // Toggle home panel vs builder inner
    const homePanel    = root.querySelector('.home-panel');
    const builderInner = root.querySelector('.npc-builder-inner');
    const isHome       = system === 'home';
    if (homePanel)    homePanel.style.display    = isHome ? 'flex'  : 'none';
    if (builderInner) builderInner.style.display = isHome ? 'none'  : 'flex';

    if (isHome) {
      // Clear all system classes, hide warnings, update auth buttons and return
      NPCBuilderApp.SYSTEMS.forEach(s => root.classList.remove(`system-${s}`));
      root.querySelectorAll('.system-warning').forEach(el => { el.style.display = 'none'; });
      this._applyAuthStateUI();
      return;
    }

    // Apply system class to root for CSS-driven theming
    NPCBuilderApp.SYSTEMS.forEach(s => root.classList.remove(`system-${s}`));
    root.classList.add(`system-${system}`);

    // Show/hide system-specific elements via data-system-only attribute
    root.querySelectorAll('[data-system-only]').forEach(el => {
      el.style.display = el.dataset.systemOnly === system ? '' : 'none';
    });

    // Show/hide warning banners
    root.querySelectorAll('.system-warning').forEach(el => { el.style.display = 'none'; });
    const warning = root.querySelector(`.system-warning--${system}`);
    if (warning) warning.style.display = 'flex';

    // Update field labels and input constraints per system
    const configs = {
      pf2e: {
        levelLabel:      'Level',
        levelMin:        '0',
        levelMax:        '25',
        levelStep:       '1',
        levelDefault:    '1',
        namePlaceholder: 'e.g. Goblin Warchief',
        descPlaceholder: 'Describe this NPC: their role, fighting style, special abilities, equipment, personality traits…',
        historyLabel:    'Created NPCs',
      },
      dnd5e: {
        levelLabel:      'Challenge Rating',
        levelMin:        '0',
        levelMax:        '30',
        levelStep:       '0.125',
        levelDefault:    '1',
        namePlaceholder: 'e.g. Bandit Captain',
        descPlaceholder: 'Describe this creature: their role, attacks, special abilities, legendary actions, lore…',
        historyLabel:    'Created Creatures',
      },
      hero6e: {
        levelLabel:      'Point Value',
        levelMin:        '25',
        levelMax:        '600',
        levelStep:       '25',
        levelDefault:    '150',
        namePlaceholder: 'e.g. Ironclad',
        descPlaceholder: 'Describe this character: their powers, combat style, skills, limitations, background…\n\nOptionally add "genre: superhero/standard/pulp/dark_champions/fantasy/sci-fi" to set the genre.',
        historyLabel:    'Created Characters',
      },
    };

    const cfg = configs[system] || configs.pf2e;

    const levelLabel = root.querySelector('label[for="npc-level"]');
    if (levelLabel) levelLabel.textContent = cfg.levelLabel;

    const levelInput = root.querySelector('#npc-level');
    if (levelInput) {
      levelInput.min  = cfg.levelMin;
      levelInput.max  = cfg.levelMax;
      levelInput.step = cfg.levelStep;
      // Snap current value to a valid Hero tier if switching to hero6e
      if (system === 'hero6e') {
        const raw = parseInt(levelInput.value) || 150;
        levelInput.value = NPCBuilderApp._snapToHero6eTier(raw);
      }
    }

    const nameInput = root.querySelector('#npc-name');
    if (nameInput) nameInput.placeholder = cfg.namePlaceholder;

    const descTextarea = root.querySelector('#npc-desc');
    if (descTextarea) descTextarea.placeholder = cfg.descPlaceholder;

    const historyLabel = root.querySelector('.history-header-label');
    if (historyLabel) historyLabel.textContent = cfg.historyLabel;

    // Sync auth UI
    this._applyAuthStateUI();
  }

  /**
   * Snap a raw point value to the nearest valid Hero System 6e tier.
   */
  static _snapToHero6eTier(raw) {
    const tiers = NPCBuilderApp.HERO6E_POINT_TIERS;
    return tiers.reduce((prev, curr) =>
      Math.abs(curr - raw) < Math.abs(prev - raw) ? curr : prev
    );
  }

  /* ── History rendering ───────────────────────────────────── */

  _renderHistory() {
    const list = this.element?.querySelector('.history-list');
    if (!list) return;

    list.innerHTML = '';

    if (this.npcHistory.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'history-empty';
      empty.textContent = 'No NPCs created yet.\nGenerate one to see it here.';
      list.appendChild(empty);
      return;
    }

    // Newest first
    for (const entry of [...this.npcHistory].reverse()) {
      list.appendChild(this._createHistoryEntryElement(entry));
    }
  }

  _createHistoryEntryElement(entry) {
    const el = document.createElement('div');
    el.className = `history-entry history-entry--${entry.status}`;
    el.dataset.entryId = entry.id;
    if (this.selectedHistoryId === entry.id) el.classList.add('is-selected');

    const statusIcon = {
      generating: '<i class="fa-solid fa-circle-notch fa-spin"></i>',
      success:    '<i class="fa-solid fa-circle-check"></i>',
      error:      '<i class="fa-solid fa-circle-xmark"></i>',
    }[entry.status] ?? '<i class="fa-solid fa-circle"></i>';

    const escapedName  = this._escapeHtml(entry.name);
    const escapedError = entry.error ? this._escapeHtml(entry.error) : '';

    // Pick the right label for the secondary meta column
    let metaLabel;
    if (entry.system === 'dnd5e') {
      metaLabel = `CR&nbsp;${entry.level}`;
    } else if (entry.system === 'hero6e') {
      metaLabel = `${entry.level}&nbsp;pts`;
    } else {
      metaLabel = `Lv.&nbsp;${entry.level}`;
    }

    el.innerHTML = `
      <div class="history-entry-main">
        <div class="history-entry-info">
          <span class="history-entry-name">${escapedName}</span>
          <span class="history-entry-meta">${metaLabel}</span>
        </div>
        <div class="history-entry-icon">${statusIcon}</div>
      </div>
      ${entry.status === 'generating'
        ? '<div class="history-progress"><div class="history-progress-bar"></div></div>'
        : ''}
      ${entry.status === 'error' && escapedError
        ? `<div class="history-entry-error">${escapedError}</div>`
        : ''}
    `;

    el.addEventListener('click', () => this._selectHistoryEntry(entry));
    return el;
  }

  /* ── Select a history entry → populate form ──────────────── */

  _selectHistoryEntry(entry) {
    this.selectedHistoryId = entry.id;

    // Switch to the system the entry was generated with (if specified)
    if (entry.system && entry.system !== this.selectedSystem && NPCBuilderApp.SYSTEMS.includes(entry.system)) {
      this.selectedSystem = entry.system;
      NPCBuilderApp.setStoredSystem(entry.system);
      this._applySystemUI();
    }

    // Populate the form with the saved prompt values
    const form = this.element?.querySelector('.npc-form');
    if (form) {
      const nameInput        = form.querySelector('[name="name"]');
      const levelInput       = form.querySelector('[name="level"]');
      const descTextarea     = form.querySelector('[name="description"]');
      const spellsCheckbox   = form.querySelector('[name="includeSpells"]');
      const casterTypeSelect = form.querySelector('[name="casterType"]');

      if (nameInput)        nameInput.value         = entry.name;
      if (levelInput)       levelInput.value        = entry.level;
      if (descTextarea)     descTextarea.value      = entry.description;
      if (spellsCheckbox)   spellsCheckbox.checked  = !!entry.includeSpells;
      if (casterTypeSelect) casterTypeSelect.value  = entry.casterType || 'none';
    }

    // Show the "editing from history" banner
    const banner = this.element?.querySelector('.history-selected-banner');
    if (banner) {
      banner.style.display = 'flex';
      const strong = banner.querySelector('strong');
      if (strong) strong.textContent = entry.name;
    }

    // Update highlighted entry
    this.element?.querySelectorAll('.history-entry').forEach(el => {
      el.classList.toggle('is-selected', el.dataset.entryId === entry.id);
    });
  }

  /* ── Update a single history entry (in memory + DOM) ─────── */

  _updateHistoryEntry(id, changes) {
    const entry = this.npcHistory.find(e => e.id === id);
    if (!entry) return;

    Object.assign(entry, changes);
    NPCBuilderApp.saveHistory(this.npcHistory);

    // Patch the DOM element in-place
    const el = this.element?.querySelector(`.history-entry[data-entry-id="${id}"]`);
    if (el) {
      const newEl = this._createHistoryEntryElement(entry);
      el.parentNode.replaceChild(newEl, el);
    }

    // Sync is-generating class on root
    const anyGenerating = this.npcHistory.some(e => e.status === 'generating');
    this.element?.classList.toggle('is-generating', anyGenerating);
  }

  /* ── HTML escape helper ──────────────────────────────────── */

  _escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = String(str);
    return div.innerHTML;
  }

  /* ── Sign-in (Patreon OAuth — popup + poll fallback) ─────── */

  async _signIn(event) {
    event?.preventDefault?.();

    const N8N_ORIGIN  = new URL(NPCBuilderApp.N8N_AUTH_URL).origin;
    const POLL_URL    = N8N_ORIGIN + '/webhook/oauth/patreon/poll';
    const POLL_MS     = 2500;   // poll every 2.5 s
    const TIMEOUT_MS  = 5 * 60 * 1000; // give up after 5 min

    console.log('[NPC Builder] starting Patreon sign-in, poll URL:', POLL_URL);

    try {
      // Generate a nonce here in the module and pass it to n8n via the login URL.
      // n8n will embed it in the OAuth state and store it with the session on callback.
      // This lets us poll /oauth/patreon/poll?nonce=<nonce> from any environment
      // (browser popup OR Electron external browser) without relying on postMessage.
      const nonce = Array.from(crypto.getRandomValues(new Uint8Array(12)))
        .map(b => b.toString(16).padStart(2,'0')).join('');

      const authUrl = NPCBuilderApp.N8N_AUTH_URL
        + '?origin=' + encodeURIComponent(window.location.origin)
        + '&nonce='  + encodeURIComponent(nonce);

      const w = 520, h = 720;
      const Y = (window.top?.outerHeight || window.outerHeight);
      const X = (window.top?.outerWidth  || window.outerWidth);
      const y = (Y / 2) + (window.top?.screenY || window.screenY) - (h / 2);
      const x = (X / 2) + (window.top?.screenX || window.screenX) - (w / 2);

      console.log('[NPC Builder] generated poll nonce:', nonce);

      // Open the auth window for the user to log in
      const win = window.open(
        authUrl,
        'patreon-login',
        `toolbar=0,location=1,status=0,menubar=0,scrollbars=1,resizable=1,width=${w},height=${h},left=${x},top=${y}`
      );

      let resolved = false;

      const onSuccess = (key) => {
        if (resolved) return;
        resolved = true;
        clearInterval(pollTimer);
        window.removeEventListener('message', msgHandler);
        try { win?.close?.(); } catch {}
        this.accessKey     = String(key);
        NPCBuilderApp.setStoredKey(this.accessKey);
        this.authenticated = true;
        this._applyAuthStateUI();
        ui.notifications?.info?.('Patreon sign-in complete.');
      };

      const onFailure = (errMsg) => {
        if (resolved) return;
        resolved = true;
        clearInterval(pollTimer);
        window.removeEventListener('message', msgHandler);
        ui.notifications?.error?.(errMsg || 'Patreon membership required to use the NPC Builder.', { permanent: true });
        setTimeout(() => window.open(NPCBuilderApp.PATREON_URL, '_blank'), 800);
      };

      // Method A: postMessage (works in browser popup flow)
      const msgHandler = (ev) => {
        // Log every message so we can diagnose whether postMessage arrives at all
        console.log('[NPC Builder] window message received — origin:', ev.origin, 'data:', ev.data);
        const okOrigins = new Set([N8N_ORIGIN, window.location.origin, 'null', '*']);
        if (!okOrigins.has(ev.origin) && ev.origin !== '') {
          console.log('[NPC Builder] postMessage ignored (origin not in allowlist):', ev.origin);
          return;
        }
        let data = ev.data;
        if (typeof data === 'string') {
          try { data = JSON.parse(data); } catch { return; }
        }
        if (!data || data.type !== 'patreon-auth') return;
        console.log('[NPC Builder] postMessage auth received:', data);
        if (data.ok && data.key && String(data.key).length >= 32) {
          console.log('[NPC Builder] postMessage success — key length:', data.key.length);
          onSuccess(data.key);
        } else {
          console.warn('[NPC Builder] postMessage auth failed:', data);
          onFailure(data?.error);
        }
      };
      window.addEventListener('message', msgHandler);

      // Method B: polling (works in Electron / external browser where opener is null)
      let pollTimer = null;
      let consecutiveErrors = 0;
      const MAX_ERRORS = 10; // ~25 s of consecutive failures before giving up
      const deadline = Date.now() + TIMEOUT_MS;
      pollTimer = setInterval(async () => {
          if (resolved) { clearInterval(pollTimer); return; }
          if (Date.now() > deadline) {
            clearInterval(pollTimer);
            if (!resolved) onFailure('Sign-in timed out. Please try again.');
            return;
          }
          try {
            const resp = await fetch(POLL_URL, {
              method: 'POST',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ nonce }),
            });
            console.log('[NPC Builder] poll status:', resp.status);
            if (resp.status === 500) {
              consecutiveErrors++;
              if (consecutiveErrors === 3) {
                console.error(
                  '[NPC Builder] Poll endpoint returning 500 errors.',
                  'The patreon_sessions table may be missing the "nonce" column.',
                  'Add it in n8n under Data → patreon_sessions → Add column → nonce (string).'
                );
              }
              if (consecutiveErrors >= MAX_ERRORS) {
                clearInterval(pollTimer);
                onFailure('Sign-in server error — please contact support or check n8n logs.');
              }
              return;
            }
            consecutiveErrors = 0;
            if (!resp.ok) {
              console.log('[NPC Builder] poll not ready yet (status', resp.status, ')');
              return; // 404 / other = not ready yet, keep polling
            }
            const data = await resp.json();
            console.log('[NPC Builder] poll response data:', data);
            if (data.ok && data.key && String(data.key).length >= 32) {
              console.log('[NPC Builder] poll success — key length:', data.key.length);
              onSuccess(data.key);
            } else if (data.error && data.error !== 'not_found') {
              console.warn('[NPC Builder] poll auth failed:', data.error);
              onFailure(data.error);
            }
            // data.pending === true means still waiting — keep polling
          } catch (err) {
            consecutiveErrors++;
            console.warn(
              `[NPC Builder] poll fetch error (${consecutiveErrors}/${MAX_ERRORS}):`, err.message,
              '— likely CORS headers missing from n8n poll endpoint error responses'
            );
            if (consecutiveErrors >= MAX_ERRORS) {
              clearInterval(pollTimer);
              onFailure('Sign-in server error — CORS or network issue on the poll endpoint.');
            }
          }
        }, POLL_MS);

      ui.notifications?.info?.('Opening Patreon sign-in…');

    } catch (err) {
      console.error('[NPC Builder] sign-in error', err);
      ui.notifications?.error?.('Failed to start Patreon sign-in.');
    }
  }

  /* ── Sign-out ────────────────────────────────────────────── */

  async _signOut(event) {
    event?.preventDefault?.();
    NPCBuilderApp.setStoredKey('');
    this.accessKey     = '';
    this.authenticated = false;
    this._applyAuthStateUI();
    ui.notifications?.info?.('Signed out.');
  }

  /* ── Build spell mapping from Foundry compendium ─────────── */

  async _buildSpellMapping() {
    console.log('[NPC Builder] Building spell mapping...');

    const spellMapping = [];
    const spellPacks   = game.packs.filter(pack =>
      pack.documentName === 'Item' &&
      pack.metadata.type === 'Item' &&
      (pack.metadata.id?.includes('spell') || pack.metadata.label?.toLowerCase().includes('spell'))
    );

    console.log(`[NPC Builder] Found ${spellPacks.length} spell packs`);

    for (const pack of spellPacks) {
      const index = await pack.getIndex({ fields: ['name', 'system.level.value', 'type'] });
      for (const entry of index) {
        if (entry.type === 'spell') {
          spellMapping.push({
            name:   entry.name,
            id:     entry._id,
            packId: pack.collection,
            level:  entry.system?.level?.value ?? 0,
          });
        }
      }
    }

    console.log(`[NPC Builder] Mapped ${spellMapping.length} spells`);
    return spellMapping;
  }

  /* ── Generate NPC (concurrent — each request gets its own history entry) ── */

  async _generateNPC(event) {
    event?.preventDefault?.();

    if (!this.authenticated) {
      ui.notifications.warn('Please sign in with Patreon before generating an NPC.');
      return;
    }

    if (this.selectedSystem === 'home') return;

    const form = this.element?.querySelector?.('.npc-form');
    if (!form) {
      ui.notifications.error('Builder form not found.');
      return;
    }

    const fd            = new FormData(form);
    const name          = (fd.get('name')?.toString()?.trim()) || 'Generated NPC';
    const level         = Number(fd.get('level')) || 1;
    const description   = (fd.get('description')?.toString()?.trim()) || '';
    const includeSpells = fd.get('includeSpells') === 'on';
    const casterType    = fd.get('casterType') || 'none';

    if (!description) {
      ui.notifications.warn('Please provide a description for the NPC.');
      return;
    }

    const key = this.accessKey || NPCBuilderApp.getStoredKey() || '';
    if (!key) {
      this.authenticated = false;
      this._applyAuthStateUI();
      ui.notifications.error('Session missing. Please sign in again.');
      return;
    }

    // ── Create history entry ──────────────────────────────────
    const historyEntry = {
      id:            foundry.utils.randomID(16),
      name,
      level,
      description,
      includeSpells,
      casterType,
      system:        this.selectedSystem,
      status:        'generating',
      createdAt:     Date.now(),
      error:         null,
    };

    this.npcHistory.push(historyEntry);
    NPCBuilderApp.saveHistory(this.npcHistory);

    // Insert at the top of the history list (newest first)
    const list = this.element?.querySelector('.history-list');
    if (list) {
      const emptyEl = list.querySelector('.history-empty');
      if (emptyEl) emptyEl.remove();
      list.insertBefore(this._createHistoryEntryElement(historyEntry), list.firstChild);
    }

    // Set is-generating on root
    this.element?.classList.add('is-generating');

    // ── Clear the selected-entry banner since we're starting fresh ──
    const banner = this.element?.querySelector('.history-selected-banner');
    if (banner) banner.style.display = 'none';
    this.selectedHistoryId = null;
    this.element?.querySelectorAll('.history-entry.is-selected').forEach(el => el.classList.remove('is-selected'));

    // ── Run generation (no await on outer scope — truly concurrent) ──
    this._runGeneration(historyEntry, key, name, level, description, includeSpells, casterType, this.selectedSystem);
  }

  /** Internal async worker for a single NPC generation. */
  async _runGeneration(historyEntry, key, name, level, description, includeSpells, casterType = 'none', system = 'pf2e') {
    try {
      let endpoint, payload;

      if (system === 'dnd5e') {
        // ── D&D 5e ───────────────────────────────────────────────────────────
        endpoint = NPCBuilderApp.N8N_DND5E_URL;
        payload  = { name, cr: level, description, casterType };
        console.log('[NPC Builder] D&D 5e generation request:', { name, cr: level, casterType });

        if (casterType !== 'none') {
          ui.notifications.info('Building spell mapping… (this may take 5–10 seconds)');
          payload.spellMapping = await this._buildSpellMapping();
          console.log(`[NPC Builder] Added ${payload.spellMapping.length} D&D 5e spells to payload`);
        }

      } else if (system === 'hero6e') {
        // ── Hero System 6e ───────────────────────────────────────────────────
        endpoint = NPCBuilderApp.N8N_HERO6E_URL;

        // Snap points to nearest valid tier
        const points = NPCBuilderApp._snapToHero6eTier(level);

        // Extract optional genre tag from description (e.g. "genre: superhero")
        let genre = 'standard';
        const genreMatch = description.match(/\bgenre\s*:\s*([\w_-]+)/i);
        if (genreMatch) {
          const extracted = genreMatch[1].toLowerCase();
          if (NPCBuilderApp.HERO6E_GENRES.includes(extracted)) genre = extracted;
        }

        payload = { name, points, genre, description };
        console.log('[NPC Builder] Hero System 6e generation request:', { name, points, genre });

      } else {
        // ── Pathfinder 2e (default) ───────────────────────────────────────────
        endpoint = NPCBuilderApp.N8N_NPC_URL;
        payload  = { name, level, description };

        if (includeSpells) {
          ui.notifications.info('Building spell mapping… (this may take 5–10 seconds)');
          payload.spellMapping = await this._buildSpellMapping();
          console.log(`[NPC Builder] Added ${payload.spellMapping.length} spells to payload`);
        }

        console.log('[NPC Builder] PF2e generation request:', {
          name,
          level,
          hasSpellMapping: !!payload.spellMapping,
          spellCount:      payload.spellMapping?.length || 0,
        });
      }

      const response = await fetch(endpoint, {
        method:  'POST',
        headers: {
          'Content-Type':     'application/json',
          'X-Builder-Key':    key,
          'X-Foundry-Origin': window.location.origin,
        },
        body: JSON.stringify(payload),
      });

      const responseText = await response.text();
      console.log('[NPC Builder] Raw response length:', responseText.length, 'bytes');

      let data;
      try {
        data = JSON.parse(responseText);
        console.log('[NPC Builder] Response from n8n:', { status: response.status, data });
      } catch (err) {
        console.error('[NPC Builder] Failed to parse JSON response:', err);
        console.error('[NPC Builder] Response text preview (first 500):', responseText.substring(0, 500));
        console.error('[NPC Builder] Response text preview (last 500):', responseText.substring(Math.max(0, responseText.length - 500)));

        const foundryNpcMatch = responseText.match(/"foundryNpc"\s*:\s*({[\s\S]*)/);
        if (foundryNpcMatch) {
          try {
            let depth = 0, inString = false, escape = false;
            const npcJson = foundryNpcMatch[1];
            for (let i = 0; i < npcJson.length; i++) {
              const char = npcJson[i];
              if (escape)        { escape = false; continue; }
              if (char === '\\') { escape = true;  continue; }
              if (char === '"')  { inString = !inString; continue; }
              if (!inString) {
                if (char === '{') depth++;
                if (char === '}') {
                  depth--;
                  if (depth === 0) {
                    data = { ok: true, foundryNpc: JSON.parse(npcJson.substring(0, i + 1)) };
                    break;
                  }
                }
              }
            }
          } catch (extractErr) {
            console.error('[NPC Builder] Failed to extract foundryNpc:', extractErr);
          }
        }

        if (!data) throw new Error(`Invalid JSON response (${responseText.length} bytes): ${err.message}`);
      }

      if (response.status === 401 || response.status === 403) {
        NPCBuilderApp.setStoredKey('');
        this.accessKey     = '';
        this.authenticated = false;
        this._applyAuthStateUI();
        this._updateHistoryEntry(historyEntry.id, { status: 'error', error: 'Authentication failed' });

        const message = data?.message || 'Unauthorized. Please sign in with Patreon.';
        ui.notifications.error(message, { permanent: true });
        setTimeout(() => window.open(NPCBuilderApp.PATREON_URL, '_blank'), 800);

      } else if (response.status === 429 || data?.error === 'RATE_LIMIT_EXCEEDED') {
        this._updateHistoryEntry(historyEntry.id, { status: 'error', error: 'Rate limit exceeded' });
        const message      = data?.message || 'Monthly NPC limit reached.';
        const currentUsage = data?.currentUsage || 0;
        const limit        = data?.limit || 0;
        ui.notifications.error(message, { permanent: true });
        ui.notifications.warn(
          `You've used ${currentUsage}/${limit} NPCs this month. Opening Patreon to upgrade…`,
          { permanent: true }
        );
        setTimeout(() => window.open(NPCBuilderApp.PATREON_URL, '_blank'), 1200);

      } else if (response.ok) {
        if (data?.ok === false) throw new Error(data?.message || data?.error || 'Server rejected the request');

        const actorData    = data.foundryNpc || data.npcDesign || data.actor || data;
        const chosenSpells = Array.isArray(data.chosenSpells) ? data.chosenSpells : [];

        if (!actorData || typeof actorData !== 'object') throw new Error('No valid actor data returned from server');
        if (!actorData.name || !actorData.type) throw new Error(`Invalid actor data: missing ${!actorData.name ? 'name' : 'type'}`);

        console.log('[NPC Builder] Creating actor in Foundry...', actorData);

        if (system === 'dnd5e') {
          this._sanitizeActorDataDnd5e(actorData);
        } else if (system === 'hero6e') {
          this._sanitizeActorDataHero6e(actorData);
        } else {
          this._sanitizeActorData(actorData);
        }

        // prototypeToken was deleted by _sanitizeActorDataDnd5e. Save name/img to
        // restore on the live document after Actor.create() completes.
        const _dnd5eTokenName = actorData.name;
        const _dnd5eTokenImg  = actorData.img || 'icons/svg/mystery-man.svg';

        // ── dnd5e 5.x / Foundry v14: merge system data against the blank NPC schema ──
        if (system === 'dnd5e') {
          try {
            const blankSchema = foundry.utils.deepClone(
              game.system.model?.Actor?.npc ?? {}
            );
            actorData.system = foundry.utils.mergeObject(
              blankSchema,
              actorData.system ?? {},
              { inplace: false, insertKeys: true, insertValues: true, overwrite: true }
            );
            console.log('[NPC Builder] D&D 5e: system merged against blank NPC schema');
            if (!actorData.system.token || typeof actorData.system.token !== 'object') {
              actorData.system.token = {};
            }
          } catch (mergeErr) {
            console.warn('[NPC Builder] D&D 5e: schema merge failed (non-fatal):', mergeErr);
          }
        }

        let actor, attempts = 0;
        const maxAttempts = 10;

        while (!actor && attempts < maxAttempts) {
          attempts++;
          try {
            actor = await Actor.create(actorData);
          } catch (error) {
            const errorText = error.toString ? error.toString() : String(error.message || error);
            if (system !== 'dnd5e' && system !== 'hero6e' && this._tryFixValidationError(actorData, errorText)) {
              console.warn(`[NPC Builder] Fixed validation error, retrying (attempt ${attempts})...`);
              continue;
            }
            throw error;
          }
        }

        if (actor) {
          if (system === 'dnd5e') {
            // Patch token name + img now that the DataModel is fully initialized
            await actor.update({
              'prototypeToken.name': _dnd5eTokenName,
              'prototypeToken.texture.src': _dnd5eTokenImg,
            });

            // Embed chosen spells from compendium
            if (chosenSpells.length > 0) {
              ui.notifications.info(`Adding ${chosenSpells.length} spells…`);
              const spellItems = [];
              for (const spell of chosenSpells) {
                try {
                  const pack = game.packs.get(spell.packId);
                  if (!pack) { console.warn('[NPC Builder] Pack not found:', spell.packId); continue; }
                  const doc = await pack.getDocument(spell.id);
                  if (doc) spellItems.push(doc.toObject());
                } catch (e) {
                  console.warn('[NPC Builder] Failed to load spell:', spell.name, e.message);
                }
              }
              if (spellItems.length > 0) {
                await actor.createEmbeddedDocuments('Item', spellItems);
                console.log(`[NPC Builder] Embedded ${spellItems.length} spells on actor`);
              }
            }
          }

          this.lastGeneratedNPC = actorData;
          this._updateHistoryEntry(historyEntry.id, { status: 'success' });
          ui.notifications.success(`NPC "${actor.name}" created successfully!`);
          actor.sheet.render(true);
        } else {
          throw new Error('Failed to create actor after maximum retry attempts');
        }

      } else {
        throw new Error(data?.message || `Server returned status ${response.status}`);
      }

    } catch (err) {
      console.error('[NPC Builder] NPC generation error', err);
      this._updateHistoryEntry(historyEntry.id, { status: 'error', error: err.message });
      ui.notifications.error(`Failed to generate "${name}": ${err.message}`);
    }
  }

  /* ── Sanitize actor data to fix common validation issues ──── */

  _sanitizeActorData(actorData) {
    const generateId = () => foundry.utils.randomID(16);

    if (!actorData._id || actorData._id.length !== 16 || !/^[a-zA-Z0-9]{16}$/.test(actorData._id)) {
      console.warn('[NPC Builder] Fixing invalid actor _id:', actorData._id);
      actorData._id = generateId();
    }

    if (actorData._stats?.exportSource?.uuid) {
      actorData._stats.exportSource.uuid = `Actor.${actorData._id}`;
    }

    const invalidWeaponTraits = [
      'melee', 'ranged', 'skirmisher', 'concealed', 'stabbing', 'light',
      'piercing', 'slashing', 'bludgeoning', 'defense', 'mobility', 'curved', 'special',
    ];
    const invalidItemTypes = new Set(['loot', 'ranged']);

    if (Array.isArray(actorData.items)) {
      actorData.items = actorData.items.filter(item => {
        if (invalidItemTypes.has(item.type)) {
          console.warn(`[NPC Builder] Removing invalid item type "${item.type}":`, item.name);
          return false;
        }
        return true;
      });

      actorData.items = actorData.items.map(item => {
        if (item.type === 'feat') {
          console.warn('[NPC Builder] Converting feat to action:', item.name);
          const description   = item.system?.description?.value || '';
          const hasActionCost = item.system?.actions?.value !== null && item.system?.actions?.value !== undefined;
          const actionType    = hasActionCost ? 'action'    : 'passive';
          const category      = hasActionCost ? 'offensive' : 'defensive';

          const action = {
            ...item,
            type: 'action',
            system: {
              ...item.system,
              description: { value: description },
              category,
              actionType: { value: actionType },
              actions:    item.system?.actions || { value: null },
            },
          };

          if (action.system.prerequisites) delete action.system.prerequisites;
          if (action.system.level && typeof action.system.level === 'object') delete action.system.level;
          if (action.system.selfEffect) {
            console.warn('[NPC Builder] Removing selfEffect from converted feat:', item.name);
            delete action.system.selfEffect;
          }
          return action;
        }

        if (item.type === 'ranged') {
          console.warn('[NPC Builder] Converting invalid "ranged" to "weapon":', item.name);
          return { ...item, type: 'weapon' };
        }

        return item;
      });

      actorData.items.forEach(item => {
        if (!item._id || item._id.length !== 16 || !/^[a-zA-Z0-9]{16}$/.test(item._id)) {
          console.warn('[NPC Builder] Fixing invalid item _id:', item._id, 'for', item.name);
          item._id = generateId();
        }

        if ((item.type === 'melee' || item.type === 'weapon') && item.system?.traits?.value) {
          const orig = item.system.traits.value;
          item.system.traits.value = orig.filter(t => !invalidWeaponTraits.includes(t.toLowerCase()));
          if (item.system.traits.value.length !== orig.length)
            console.warn('[NPC Builder] Removed invalid traits from', item.name);
        }

        if (item.system?.traits?.value && Array.isArray(item.system.traits.value)) {
          const orig = item.system.traits.value;
          item.system.traits.value = orig.filter(t => t.toLowerCase() !== 'special');
          if (item.system.traits.value.length !== orig.length)
            console.warn('[NPC Builder] Removed "special" trait from', item.name);
        }
      });
    }

    console.log('[NPC Builder] Actor data sanitized successfully');
  }

  /* ── Sanitize D&D 5e actor data ──────────────────────────── */

  _sanitizeActorDataDnd5e(actorData) {
    const generateId = () => foundry.utils.randomID(16);

    if (!actorData._id || actorData._id.length !== 16 || !/^[a-zA-Z0-9]{16}$/.test(actorData._id)) {
      console.warn('[NPC Builder] D&D 5e: Fixing invalid actor _id:', actorData._id);
      actorData._id = generateId();
    }

    if (actorData.type !== 'npc') {
      console.warn('[NPC Builder] D&D 5e: Correcting actor type to "npc" (was:', actorData.type, ')');
      actorData.type = 'npc';
    }

    if (Array.isArray(actorData.items)) {
      actorData.items.forEach(item => {
        if (!item._id || item._id.length !== 16 || !/^[a-zA-Z0-9]{16}$/.test(item._id)) {
          console.warn('[NPC Builder] D&D 5e: Fixing invalid item _id:', item._id, 'for', item.name);
          item._id = generateId();
        }
        if (!item.system) item.system = {};
        if (!item.system.description) item.system.description = { value: '', chat: '', unidentified: '' };
      });
    }

    if (!actorData.flags) actorData.flags = {};
    if (!actorData.img) actorData.img = 'icons/svg/mystery-man.svg';

    const _abilities = actorData.system?.abilities || {};

    if (actorData.system?.save && typeof actorData.system.save === 'object') {
      console.warn('[NPC Builder] D&D 5e: Fixing misplaced system.save');
      for (const [k, v] of Object.entries(actorData.system.save)) {
        if (_abilities[k]) _abilities[k].proficient = v?.proficient ?? 0;
      }
      delete actorData.system.save;
    }
    if (actorData.system?.['attributes.save'] && typeof actorData.system['attributes.save'] === 'object') {
      console.warn('[NPC Builder] D&D 5e: Fixing misplaced system["attributes.save"]');
      for (const [k, v] of Object.entries(actorData.system['attributes.save'])) {
        if (_abilities[k]) _abilities[k].proficient = v?.proficient ?? 0;
      }
      delete actorData.system['attributes.save'];
    }
    if (actorData.system?.attributes?.save && typeof actorData.system.attributes.save === 'object') {
      console.warn('[NPC Builder] D&D 5e: Fixing misplaced system.attributes.save');
      for (const [k, v] of Object.entries(actorData.system.attributes.save)) {
        if (_abilities[k]) _abilities[k].proficient = v?.proficient ?? 0;
      }
      delete actorData.system.attributes.save;
    }

    const _traits = actorData.system?.traits;
    if (_traits) {
      const toTraitSet = (val) => {
        if (val && typeof val === 'object' && !Array.isArray(val) && Array.isArray(val.value)) {
          if (!val.custom) val.custom = '';
          return val;
        }
        if (Array.isArray(val)) return { value: val, custom: '' };
        return { value: [], custom: '' };
      };
      if (_traits.di        !== undefined) _traits.di        = toTraitSet(_traits.di);
      if (_traits.dr        !== undefined) _traits.dr        = toTraitSet(_traits.dr);
      if (_traits.dv        !== undefined) _traits.dv        = toTraitSet(_traits.dv);
      if (_traits.ci        !== undefined) _traits.ci        = toTraitSet(_traits.ci);
      if (_traits.languages !== undefined) _traits.languages = toTraitSet(_traits.languages);
    }

    const _attrs = actorData.system?.attributes;
    if (_attrs && !_attrs.movement) {
      const spd = _attrs.speed;
      const walkVal = (spd && typeof spd === 'object' ? spd.value : spd) || 30;
      _attrs.movement = {
        burrow: 0, climb: 0, fly: 0, swim: 0,
        walk: typeof walkVal === 'number' ? walkVal : (parseInt(walkVal) || 30),
        units: 'ft', hover: false,
      };
    }

    if (actorData.system && !actorData.system.currency) {
      actorData.system.currency = { pp: 0, gp: 0, ep: 0, sp: 0, cp: 0 };
    }

    delete actorData.prototypeToken;

    console.log('[NPC Builder] D&D 5e actor data sanitized:', actorData.name, '| items:', actorData.items?.length || 0);
  }

  /* ── Sanitize Hero System 6e actor data ──────────────────── */

  /**
   * Cleans up a hero6efoundryvttv2 actor object returned from the n8n workflow
   * before passing it to Actor.create().
   *
   * IMPORTANT: n8n Step 4 injects critical fields onto each item.system before
   * returning the response. This function MUST preserve those fields:
   *
   *   INPUT       — required by _getNonCharacteristicsBasedRollComponents; calling
   *                 INPUT.includes() on undefined causes a TypeError that crashes
   *                 the entire actor sheet render.
   *   OPTIONID    — required by COMBAT_LEVELS, FLASHDEFENSE, STRIKING_APPEARANCE,
   *                 SKILL_LEVELS, etc.; missing causes a crash on sheet open.
   *   OPTION      — companion to OPTIONID.
   *   OPTION_ALIAS — display label for the selected option.
   *   CHARACTERISTIC — required by skills (DEX, INT, PRE, etc.) and some talents.
   *   ADDER       — required by REPUTATION, ACCIDENTALCHANGE, HUNTED, SOCIALLIMITATION,
   *                 ENRAGED, PSYCHOLOGICALLIMITATION; missing causes a fatal TypeError.
   *   xmlTag      — required by getPowerInfo() to resolve the correct power definition.
   *   is5e        — must be false on every item to suppress 5e-mode warnings.
   *
   * CHARACTERISTIC STORAGE — two paths must both be populated:
   *
   *   system.characteristics[KEY] — what this sanitizer reads (LEVELS → max/value).
   *
   *   system[KEY] e.g. system.STR  — EmbeddedDataField HeroItemCharacteristic.
   *     hero6efoundryvttv2's _preCreate iterates all uppercase system keys and, when
   *     XMLID is MISSING, does:
   *       actorChanges.system[KEY] = { XMLID: KEY, xmlTag: KEY }
   *     then calls updateSource(actorChanges) which REPLACES the embedded object,
   *     wiping LEVELS back to 0.
   *     Fix: always write { LEVELS, XMLID, xmlTag } onto every direct uppercase key
   *     so the `!char.XMLID` guard is false and _preCreate skips it entirely.
   *
   * This function only handles structural concerns (IDs, type coercion, required
   * defaults). It trusts n8n's validated XMLIDs and injected fields completely.
   */
  _sanitizeActorDataHero6e(actorData) {
    const generateId = () => foundry.utils.randomID(16);

    // ── Actor-level fields ────────────────────────────────────────────────────
    if (!actorData._id || actorData._id.length !== 16 || !/^[a-zA-Z0-9]{16}$/.test(actorData._id)) {
      console.warn('[NPC Builder] Hero 6e: Fixing invalid actor _id:', actorData._id);
      actorData._id = generateId();
    }

    if (actorData.type !== 'npc') {
      console.warn('[NPC Builder] Hero 6e: Correcting actor type to "npc" (was:', actorData.type, ')');
      actorData.type = 'npc';
    }

    if (!actorData.img)     actorData.img     = 'icons/svg/mystery-man.svg';
    if (!actorData.flags)   actorData.flags   = {};
    if (!actorData.effects) actorData.effects = [];

    // ── System-level fields ───────────────────────────────────────────────────
    if (!actorData.system) actorData.system = {};
    const sys = actorData.system;

    if (typeof sys.is5e === 'undefined') sys.is5e = false;

    // ── Characteristics ───────────────────────────────────────────────────────
    // hero6efoundryvttv2 stores characteristics in TWO places that must both be set:
    //
    // 1. system.characteristics[KEY]  (HeroCharacteristicsModel — what THIS code reads)
    //    { LEVELS, value, max } where value = max = base + LEVELS.
    //
    // 2. system[KEY]  e.g. system.STR  (EmbeddedDataField HeroItemCharacteristic)
    //    _preCreate reads these. If XMLID is absent it REPLACES the whole object with
    //    { XMLID, xmlTag }, zeroing LEVELS. We must include XMLID so _preCreate skips
    //    the field. After Actor.create() the _onUpdate hook reads system[KEY].LEVELS
    //    and propagates it into system.characteristics[key].max / .value.
    const CHAR_BASES = {
      STR:10, DEX:10, CON:10, INT:10, EGO:10, PRE:10,
      OCV:3,  DCV:3,  OMCV:3, DMCV:3,
      SPD:2,  PD:2,   ED:2,
      REC:4,  END:20, BODY:10, STUN:20,
    };

    if (!sys.characteristics) sys.characteristics = {};
    const chars = sys.characteristics;

    for (const [key, base] of Object.entries(CHAR_BASES)) {
      // 1. Validate / compute system.characteristics[key]
      if (!chars[key]) {
        chars[key] = { value: base, max: base, LEVELS: 0 };
      } else {
        const c      = chars[key];
        const levels = Math.max(0, parseInt(c.LEVELS) || 0);
        c.LEVELS     = levels;
        c.max        = base + levels;
        c.value      = Math.min(parseInt(c.value) || c.max, c.max);
      }

      // 2. Mirror onto the direct uppercase field with XMLID present.
      //    This prevents _preCreate from replacing the object and zeroing LEVELS.
      const levels = chars[key].LEVELS;
      sys[key] = { LEVELS: levels, XMLID: key, xmlTag: key };
    }

    // Remove any non-standard characteristic keys (e.g. "Natural", lowercase dupes).
    // The Hero system iterates ALL keys in characteristics and calls getPowerInfo()
    // on each — any unknown key logs "Unable to find 6e power entry" and can
    // interfere with point calculations and sheet rendering.
    for (const k of Object.keys(chars)) {
      if (!CHAR_BASES[k]) {
        console.warn('[NPC Builder] Hero 6e: Removing non-standard characteristic key:', k);
        delete chars[k];
      }
    }

    // ── Items: coerce LEVELS to string, ensure required fields ───────────────
    if (!Array.isArray(actorData.items)) actorData.items = [];

    const VALID_ITEM_TYPES = new Set([
      'power', 'skill', 'talent', 'complication', 'equipment',
      'perk', 'martialart', 'maneuver', 'characteristic',
    ]);

    actorData.items = actorData.items.map(item => {
      if (!item._id || item._id.length !== 16 || !/^[a-zA-Z0-9]{16}$/.test(item._id)) {
        item._id = generateId();
      }

      // Default type to 'power' if missing or invalid
      if (!VALID_ITEM_TYPES.has(item.type)) {
        console.warn('[NPC Builder] Hero 6e: Unknown item type', item.type, '— defaulting to "power"');
        item.type = 'power';
      }

      if (!item.system) item.system = {};
      const s = item.system;

      // XMLID: trust n8n's validated value entirely.
      // Only set a safe fallback if the field is completely absent.
      // Do NOT use invalid 5e-era defaults like 'PSYCHOLOGICAL_LIMITATION' (broken —
      // the correct 6e XMLID has no underscores: PSYCHOLOGICALLIMITATION) or 'SKILL'
      // (doesn't exist in hero6efoundryvttv2).
      if (!s.XMLID) {
        const xmlidDefaults = {
          power:        'CUSTOMPOWER',
          skill:        'CUSTOMSKILL',
          talent:       'CUSTOMTALENT',
          complication: 'GENERICDISADVANTAGE',
        };
        s.XMLID = xmlidDefaults[item.type] || 'CUSTOMPOWER';
        console.warn('[NPC Builder] Hero 6e: Missing XMLID on', item.type, '→ defaulting to', s.XMLID);
      }

      // LEVELS must be a string (hero6e system reads it that way)
      if (typeof s.LEVELS === 'number') s.LEVELS = String(s.LEVELS);
      if (!s.LEVELS) s.LEVELS = '1';

      // ALIAS doubles as the display name
      if (!s.ALIAS)       s.ALIAS       = item.name || s.XMLID;
      if (!s.description) s.description = item.name || '';

      // Numeric fields
      if (s.active_points !== undefined) s.active_points = parseInt(s.active_points) || 0;
      if (s.real_cost     !== undefined) s.real_cost     = parseInt(s.real_cost)     || 0;
      if (s.ENDCOST       !== undefined) s.ENDCOST       = parseInt(s.ENDCOST)       || 0;

      // Complications need a numeric POINTS value
      if (item.type === 'complication') {
        s.POINTS = parseInt(s.POINTS) || 10;
      }

      // Debug: log key fields for every power so we can confirm INPUT/OPTIONID survive
      if (item.type === 'power') {
        console.log('[NPC Builder] Hero 6e item:', s.XMLID,
          '| INPUT:', s.INPUT        || 'MISSING',
          '| OPTIONID:', s.OPTIONID  || 'none',
          '| CHARACTERISTIC:', s.CHARACTERISTIC || 'none');
      }

      item.system = s;
      return item;
    });

    console.log('[NPC Builder] Hero 6e actor data sanitized:', actorData.name,
      '| items:', actorData.items.length,
      '| powers:', actorData.items.filter(i => i.type === 'power').length,
      '| skills:', actorData.items.filter(i => i.type === 'skill').length,
      '| complications:', actorData.items.filter(i => i.type === 'complication').length);
  }

  /**
   * Try to fix a validation error by parsing the error message and modifying actorData.
   * Returns true if the error was fixed and a retry should be attempted.
   */
  _tryFixValidationError(actorData, errorMessage) {
    console.log('[NPC Builder] Attempting to fix validation error:', errorMessage);

    const invalidTraitMatch = errorMessage.match(/(\w+) is not a valid choice/);
    if (invalidTraitMatch) {
      const badTrait = invalidTraitMatch[1];
      let removed    = false;
      if (Array.isArray(actorData.items)) {
        actorData.items.forEach(item => {
          if (item.system?.traits?.value && Array.isArray(item.system.traits.value)) {
            const before = item.system.traits.value.length;
            item.system.traits.value = item.system.traits.value.filter(
              t => t.toLowerCase() !== badTrait.toLowerCase()
            );
            if (item.system.traits.value.length < before) removed = true;
          }
        });
      }
      return removed;
    }

    const invalidDocIdMatch = errorMessage.match(/Invalid document ID "([^"]+)"/);
    if (invalidDocIdMatch) {
      const invalidId = invalidDocIdMatch[1];
      let removed     = false;
      if (Array.isArray(actorData.items)) {
        actorData.items.forEach(item => {
          if (item.system?.selfEffect?.uuid?.includes(invalidId)) {
            delete item.system.selfEffect;
            removed = true;
          }
        });
      }
      return removed;
    }

    const invalidTypeMatch = errorMessage.match(/"(\w+)" is not a valid type/);
    if (invalidTypeMatch) {
      const invalidType = invalidTypeMatch[1];
      let fixed         = false;
      if (Array.isArray(actorData.items)) {
        if (invalidType === 'loot') {
          const before = actorData.items.length;
          actorData.items = actorData.items.filter(i => i.type !== 'loot');
          fixed = actorData.items.length < before;
        } else {
          actorData.items.forEach(item => {
            if (item.type === invalidType) { item.type = 'weapon'; fixed = true; }
          });
        }
      }
      return fixed;
    }

    return false;
  }

  /* ── Export last generated NPC as JSON ───────────────────── */

  async _exportJSON(event) {
    event?.preventDefault?.();

    if (!this.lastGeneratedNPC) {
      ui.notifications.warn('No NPC has been generated yet.');
      return;
    }

    const json = JSON.stringify(this.lastGeneratedNPC, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = `${this.lastGeneratedNPC.name || 'npc'}.json`;
    a.click();
    URL.revokeObjectURL(url);

    ui.notifications.info('NPC exported to JSON file.');
  }
}

/* -----------------------------------------------------------------------------
   Singleton helper — reuses an existing window instead of spawning duplicates
----------------------------------------------------------------------------- */

let _npcBuilderApp = null;

function openNPCBuilder() {
  if (_npcBuilderApp?.rendered && _npcBuilderApp?.element?.isConnected) {
    _npcBuilderApp.bringToTop?.();
  } else {
    _npcBuilderApp = null;
    _npcBuilderApp = new NPCBuilderApp();
    _npcBuilderApp.render({ force: true }).catch(err => {
      console.error('[NPC Builder] Failed to open:', err);
      ui.notifications?.error?.('NPC Builder failed to open. Check the console (F12) for details.');
      _npcBuilderApp = null;
    });
  }

  // Non-blocking update check — runs every time the builder is opened
  _checkForModuleUpdate().catch(() => {});
}

/* -----------------------------------------------------------------------------
   Update checker — fetches the manifest and shows a popup when outdated
----------------------------------------------------------------------------- */

/** Returns true if `a` is strictly newer than `b` (simple semver). */
function _isNewerVersion(a, b) {
  const parse = v => (v || '0.0.0').split('.').map(n => parseInt(n, 10) || 0);
  const [a1, a2, a3] = parse(a);
  const [b1, b2, b3] = parse(b);
  if (a1 !== b1) return a1 > b1;
  if (a2 !== b2) return a2 > b2;
  return a3 > b3;
}

async function _checkForModuleUpdate() {
  try {
    const modId          = game.modules?.get('Pf2eNpcMaker') ? 'Pf2eNpcMaker' : 'pf2e-npc-auto-builder';
    const mod            = game.modules?.get(modId);
    const manifestUrl    = mod?.manifest;
    const currentVersion = mod?.version || '';

    if (!manifestUrl || !currentVersion) return;

    const response = await fetch(manifestUrl, { cache: 'no-cache' });
    if (!response.ok) return;

    const data          = await response.json();
    const latestVersion = data?.version || '';

    if (!latestVersion || !_isNewerVersion(latestVersion, currentVersion)) return;

    // Build the popup content
    const content = `
      <div style="display:flex;flex-direction:column;gap:0.6em;padding:0.25em 0;">
        <p style="margin:0;">
          <strong>NPC Builder v${latestVersion}</strong> is available.
          You are running <strong>v${currentVersion}</strong>.
        </p>
        <p style="margin:0;color:#555;font-size:0.92em;">
          Update via the Foundry <em>Add-on Modules</em> manager or from GitHub to get the
          latest features and bug fixes.
        </p>
      </div>`;

    // Prefer DialogV2 (Foundry v13+), fall back to classic Dialog
    const DialogV2 = foundry.applications?.api?.DialogV2;
    if (DialogV2) {
      DialogV2.prompt({
        window:  { title: 'NPC Builder — Update Available' },
        content,
        ok: {
          label:    'View on GitHub',
          icon:     'fa-brands fa-github',
          callback: () => window.open('https://github.com/JamesCfer/Pf2eNpcMaker/releases/latest', '_blank'),
        },
        rejectClose: false,
      }).catch(() => {});
    } else {
      new Dialog({
        title:   'NPC Builder — Update Available',
        content,
        buttons: {
          github: {
            label:    '<i class="fa-brands fa-github"></i> View on GitHub',
            callback: () => window.open('https://github.com/JamesCfer/Pf2eNpcMaker/releases/latest', '_blank'),
          },
          dismiss: { label: 'Dismiss' },
        },
        default: 'dismiss',
      }).render(true);
    }
  } catch (err) {
    // Network errors are expected offline — log quietly and move on
    console.debug('[NPC Builder] Update check failed (offline?):', err);
  }
}

/* -----------------------------------------------------------------------------
   Header controls + sidebar injection
----------------------------------------------------------------------------- */

function registerNPCBuilderControl(app, controls) {
  if (!game.user?.isGM) return;
  const exists = controls.some(c => c.action === 'pf2e-npc-builder');
  if (exists) return;
  controls.push({
    action:  'pf2e-npc-builder',
    icon:    'fa-solid fa-star',
    label:   'NPC Builder',
    onClick: () => openNPCBuilder(),
    onclick: () => openNPCBuilder(),
    visible: true,
  });
}

function injectSidebarButton(app, html) {
  if (!game.user?.isGM) return;
  const root = html instanceof HTMLElement ? html : html?.[0];
  if (!root) return;
  if (root.querySelector('.npc-builder-button')) return;

  const button = document.createElement('button');
  button.type  = 'button';
  button.classList.add('npc-builder-button');
  button.style.marginLeft = '4px';
  button.innerHTML = '★ NPC Builder ★';
  button.addEventListener('click', () => openNPCBuilder());

  const header = root.querySelector('header') || root.querySelector('.directory-header');
  if (header) header.appendChild(button); else root.prepend(button);
}

// Support common hook names across versions
Hooks.on('getHeaderControlsActorDirectory',          registerNPCBuilderControl);
Hooks.on('getHeaderControlsCompendiumDirectory',      registerNPCBuilderControl);
Hooks.on('getHeaderControlsActorDirectoryPF2e',       registerNPCBuilderControl);
Hooks.on('getHeaderControlsCompendiumDirectoryPF2e',  registerNPCBuilderControl);
Hooks.on('getHeaderControlsApplicationV2', (app, controls) => {
  try {
    const name = app?.constructor?.name;
    if (
      name === 'ActorDirectory' || name === 'CompendiumDirectory' ||
      name === 'ActorDirectoryPF2e' || name === 'CompendiumDirectoryPF2e'
    ) registerNPCBuilderControl(app, controls);
  } catch (err) {
    console.warn('PF2E NPC Builder: generic header control hook failed', err);
  }
});

Hooks.on('renderActorDirectory',             injectSidebarButton);
Hooks.on('renderCompendiumDirectory',        injectSidebarButton);
Hooks.on('renderActorDirectoryPF2e',         injectSidebarButton);
Hooks.on('renderCompendiumDirectoryPF2e',    injectSidebarButton);

Hooks.once('ready', () => {
  const modId         = game.modules?.get('Pf2eNpcMaker') ? 'Pf2eNpcMaker' : 'pf2e-npc-auto-builder';
  const currentVersion = game.modules?.get(modId)?.version || '';
  const storedVersion  = NPCBuilderApp.getStoredVersion();

  // Sign users out when the module updates so stale sessions don't persist
  if (currentVersion && storedVersion && currentVersion !== storedVersion) {
    NPCBuilderApp.setStoredKey('');
    console.log(`[NPC Builder] Module updated ${storedVersion} → ${currentVersion}. Session cleared.`);
    ui.notifications?.info?.('NPC Builder was updated — please sign in again.', { permanent: false });
  }

  if (currentVersion) NPCBuilderApp.setStoredVersion(currentVersion);

  (foundry.applications.handlebars?.loadTemplates ?? loadTemplates)([`modules/${modId}/templates/builder.html`]);
  console.log(`PF2E NPC Auto-Builder ready (module folder: ${modId}, version: ${currentVersion}).`);
});
