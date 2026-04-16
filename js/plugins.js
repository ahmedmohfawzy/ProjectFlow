/**
 * ProjectFlow™ Plugin System
 * © 2026 Ahmed M. Fawzy
 *
 * Browser-based Project Management Application
 * E.5 — Plugin System Implementation
 *
 * Provides extensible plugin architecture with:
 * - Plugin registration and lifecycle management
 * - Event/hook system for plugin interaction
 * - Scoped localStorage for plugin data persistence
 * - Built-in plugins (Pomodoro, Risk Register, Activity Log)
 * - Plugin manager UI
 */


    // PRIVATE STATE
    // ============================================================================

    const _plugins = new Map();           // id → pluginDef
    const _hooks = new Map();             // eventName → [callback]
    const _activityLog = [];              // Activity log for logging plugin
    const _activityLogMaxSize = 50;       // Max entries in activity log

    // ============================================================================
    // PRIVATE UTILITIES
    // ============================================================================

    /**
     * Generate a scoped localStorage key for a plugin
     * @param {string} pluginId
     * @returns {string}
     */
    function _makeScopedKey(pluginId) {
        return `pf_plugin_${pluginId}`;
    }

    /**
     * Log an entry to the activity log
     * @param {string} event
     * @param {string} summary
     */
    function _logActivity(event, summary) {
        const entry = {
            timestamp: new Date().toISOString(),
            event: event,
            summary: summary
        };
        _activityLog.unshift(entry);
        if (_activityLog.length > _activityLogMaxSize) {
            _activityLog.pop();
        }
    }

    // ============================================================================
    // CORE FUNCTIONS
    // ============================================================================

    /**
     * Register a plugin
     * @param {Object} pluginDef - { id, name, version, description, init?, destroy? }
     */
    function register(pluginDef) {
        if (!pluginDef || !pluginDef.id) {
            console.warn('PluginSystem.register: plugin must have an id');
            return false;
        }
        if (_plugins.has(pluginDef.id)) {
            console.warn(`PluginSystem.register: plugin "${pluginDef.id}" already registered`);
            return false;
        }
        _plugins.set(pluginDef.id, pluginDef);
        _logActivity('plugin:registered', `Plugin "${pluginDef.name || pluginDef.id}" registered`);
        if (typeof pluginDef.init === 'function') {
            try { pluginDef.init({ storage: getStorage(pluginDef.id), hooks: { addHook, removeHook } }); }
            catch (e) { console.error(`Plugin "${pluginDef.id}" init error:`, e); }
        }
        return true;
    }

    /**
     * Unregister a plugin
     * @param {string} pluginId
     */
    function unregister(pluginId) {
        const pluginDef = _plugins.get(pluginId);
        if (!pluginDef) return false;
        if (typeof pluginDef.destroy === 'function') {
            try { pluginDef.destroy(); }
            catch (e) { console.error(`Plugin "${pluginId}" destroy error:`, e); }
        }
        _plugins.delete(pluginId);
        _logActivity('plugin:unregistered', `Plugin "${pluginDef.name || pluginId}" unregistered`);
        return true;
    }

    /**
     * List all registered plugins
     * @returns {Object[]}
     */
    function list() {
        return Array.from(_plugins.values()).map(p => ({
            id: p.id, name: p.name || p.id, version: p.version || '1.0',
            description: p.description || ''
        }));
    }

    /**
     * Add a hook callback for an event
     * @param {string} eventName
     * @param {Function} callback
     */
    function addHook(eventName, callback) {
        if (!_hooks.has(eventName)) _hooks.set(eventName, []);
        _hooks.get(eventName).push(callback);
    }

    /**
     * Remove a hook callback
     * @param {string} eventName
     * @param {Function} callback
     */
    function removeHook(eventName, callback) {
        if (!_hooks.has(eventName)) return;
        const arr = _hooks.get(eventName);
        const idx = arr.indexOf(callback);
        if (idx >= 0) arr.splice(idx, 1);
    }

    /**
     * Emit an event to all registered hooks
     * @param {string} eventName
     * @param {*} data
     */
    function emit(eventName, data) {
        const callbacks = _hooks.get(eventName);
        if (!callbacks) return;
        callbacks.forEach(cb => {
            try { cb(data); }
            catch (e) { console.error(`PluginSystem hook error [${eventName}]:`, e); }
        });
    }

    /**
     * Get scoped storage helpers for a plugin
     * @param {string} pluginId
     * @returns {{ get, set, remove }}
     */
    function getStorage(pluginId) {
        const key = _makeScopedKey(pluginId);
        return {
            get(prop) {
                try {
                    const data = JSON.parse(localStorage.getItem(key) || '{}');
                    return prop ? data[prop] : data;
                } catch { return prop ? undefined : {}; }
            },
            set(prop, value) {
                try {
                    const data = JSON.parse(localStorage.getItem(key) || '{}');
                    data[prop] = value;
                    localStorage.setItem(key, JSON.stringify(data));
                } catch (e) { console.error('PluginSystem storage set error:', e); }
            },
            remove(prop) {
                try {
                    const data = JSON.parse(localStorage.getItem(key) || '{}');
                    delete data[prop];
                    localStorage.setItem(key, JSON.stringify(data));
                } catch (e) { console.error('PluginSystem storage remove error:', e); }
            }
        };
    }

    /**
     * Get activity log entries
     * @returns {Object[]}
     */
    function getActivityLog() {
        return [..._activityLog];
    }

    // ============================================================================
    // BUILT-IN PLUGINS
    // ============================================================================

    const builtins = {
        pomodoroPlugin: {
            id: 'pomodoro', name: 'Pomodoro Timer', version: '1.0',
            description: 'Focus timer with 25-min work / 5-min break cycles.',
            init(ctx) { _logActivity('pomodoro:init', 'Pomodoro timer activated'); },
            destroy() { _logActivity('pomodoro:destroy', 'Pomodoro timer deactivated'); }
        },
        riskRegisterPlugin: {
            id: 'risk-register', name: 'Risk Register', version: '1.0',
            description: 'Track and manage project risks with probability and impact scoring.',
            init(ctx) { _logActivity('risk:init', 'Risk Register activated'); },
            destroy() { _logActivity('risk:destroy', 'Risk Register deactivated'); }
        },
        activityLogPlugin: {
            id: 'activity-log', name: 'Activity Log', version: '1.0',
            description: 'Records all project actions and plugin events for audit trail.',
            init(ctx) { _logActivity('actlog:init', 'Activity Log activated'); },
            destroy() { _logActivity('actlog:destroy', 'Activity Log deactivated'); }
        }
    };

    // ============================================================================
    // PUBLIC API
    // ============================================================================

    export const PluginSystem = {
        register,
        unregister,
        list,
        addHook,
        removeHook,
        emit,
        getStorage,
        getActivityLog,
        builtins
    };

/**
 * Render the Plugin Manager UI
 *
 * @param {HTMLElement} container - Container to render into
 */
export function renderPluginManager(container) {
    if (!container) {
        console.warn('renderPluginManager: container is required');
        return;
    }

    // Clear container
    container.innerHTML = '';

    // Main panel
    const panel = document.createElement('div');
    panel.className = 'pf-plugin-manager';
    panel.style.cssText = `
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        padding: 20px;
        max-width: 600px;
    `;

    // Header
    const header = document.createElement('h2');
    header.textContent = 'Plugin Manager';
    header.style.marginTop = '0';
    panel.appendChild(header);

    // Active plugins section
    const activeSection = document.createElement('div');
    activeSection.style.marginBottom = '30px';

    const activeTitle = document.createElement('h3');
    activeTitle.textContent = 'Active Plugins';
    activeTitle.style.borderBottom = '2px solid #e9ecef';
    activeTitle.style.paddingBottom = '8px';
    activeSection.appendChild(activeTitle);

    const activePlugins = PluginSystem.list();

    if (activePlugins.length === 0) {
        const empty = document.createElement('p');
        empty.textContent = 'No plugins active.';
        empty.style.color = '#868e96';
        activeSection.appendChild(empty);
    } else {
        activePlugins.forEach(plugin => {
            const card = document.createElement('div');
            card.style.cssText = `
                border: 1px solid #dee2e6;
                border-radius: 4px;
                padding: 12px;
                margin-bottom: 10px;
                display: flex;
                justify-content: space-between;
                align-items: center;
            `;

            const info = document.createElement('div');
            info.innerHTML = `
                <strong>${plugin.name}</strong> <span style="color: #868e96;">v${plugin.version}</span><br>
                <small style="color: #495057;">${plugin.description}</small>
            `;

            const unregBtn = document.createElement('button');
            unregBtn.textContent = 'Unregister';
            unregBtn.className = 'pf-btn pf-btn-danger';
            unregBtn.style.cssText = `
                padding: 6px 12px;
                background: #dc3545;
                color: white;
                border: none;
                border-radius: 3px;
                cursor: pointer;
                font-size: 12px;
            `;
            unregBtn.addEventListener('click', () => {
                if (confirm(`Unregister plugin "${plugin.name}"?`)) {
                    PluginSystem.unregister(plugin.id);
                    renderPluginManager(container);
                }
            });

            card.appendChild(info);
            card.appendChild(unregBtn);
            activeSection.appendChild(card);
        });
    }

    panel.appendChild(activeSection);

    // Available built-ins section
    const builtinSection = document.createElement('div');

    const builtinTitle = document.createElement('h3');
    builtinTitle.textContent = 'Available Built-in Plugins';
    builtinTitle.style.borderBottom = '2px solid #e9ecef';
    builtinTitle.style.paddingBottom = '8px';
    builtinSection.appendChild(builtinTitle);

    const builtinIds = ['pomodoro', 'risk-register', 'activity-log'];
    const activeIds = new Set(activePlugins.map(p => p.id));

    builtinIds.forEach(builtinId => {
        const builtinPlugin = PluginSystem.builtins[builtinId === 'pomodoro' ? 'pomodoroPlugin' :
                                                     builtinId === 'risk-register' ? 'riskRegisterPlugin' :
                                                     'activityLogPlugin'];

        const isActive = activeIds.has(builtinId);

        const card = document.createElement('div');
        card.style.cssText = `
            border: 1px solid #dee2e6;
            border-radius: 4px;
            padding: 12px;
            margin-bottom: 10px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            background: ${isActive ? '#f0f9ff' : 'white'};
        `;

        const info = document.createElement('div');
        info.innerHTML = `
            <strong>${builtinPlugin.name}</strong> <span style="color: #868e96;">v${builtinPlugin.version}</span><br>
            <small style="color: #495057;">${builtinPlugin.description}</small>
        `;

        const actionBtn = document.createElement('button');
        actionBtn.className = 'pf-btn';
        actionBtn.textContent = isActive ? 'Active' : 'Activate';
        actionBtn.disabled = isActive;
        actionBtn.style.cssText = `
            padding: 6px 12px;
            background: ${isActive ? '#28a745' : '#007bff'};
            color: white;
            border: none;
            border-radius: 3px;
            cursor: ${isActive ? 'default' : 'pointer'};
            font-size: 12px;
            opacity: ${isActive ? '0.7' : '1'};
        `;

        if (!isActive) {
            actionBtn.addEventListener('click', () => {
                PluginSystem.register(builtinPlugin);
                renderPluginManager(container);
            });
        }

        card.appendChild(info);
        card.appendChild(actionBtn);
        builtinSection.appendChild(card);
    });

    panel.appendChild(builtinSection);

    container.appendChild(panel);
}
