/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 * Proprietary Software — Unauthorized use prohibited.
 * https://www.linkedin.com/in/ahmed-m-fawzy
 */
/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — Central EventBus
 * Sprint B.3 — Modularize Architecture
 * Lightweight pub/sub for decoupled module communication
 * ═══════════════════════════════════════════════════════
 */



    /** @type {Object.<string, Function[]>} */
    const _listeners = {};

    /**
     * Subscribe to an event
     * @param {string} event
     * @param {Function} callback
     */
    function on(event, callback) {
        if (typeof callback !== 'function') return;
        if (!_listeners[event]) _listeners[event] = [];
        if (!_listeners[event].includes(callback)) {
            _listeners[event].push(callback);
        }
    }

    /**
     * Unsubscribe from an event
     * @param {string} event
     * @param {Function} callback
     */
    function off(event, callback) {
        if (!_listeners[event]) return;
        _listeners[event] = _listeners[event].filter(cb => cb !== callback);
    }

    /**
     * Subscribe to an event — fires once then auto-removes
     * @param {string} event
     * @param {Function} callback
     */
    function once(event, callback) {
        const wrapper = (data) => { callback(data); off(event, wrapper); };
        on(event, wrapper);
    }

    /**
     * Emit an event to all subscribers
     * @param {string} event
     * @param {*} data — payload passed to callbacks
     */
    function emit(event, data) {
        const callbacks = _listeners[event];
        if (!callbacks || callbacks.length === 0) return;
        // Iterate over a snapshot to avoid mutation issues
        [...callbacks].forEach(cb => {
            try { cb(data); } catch (err) {
                console.error(`[EventBus] Error in handler for "${event}":`, err);
            }
        });
    }

    /**
     * Remove all listeners for an event (or all events)
     * @param {string} [event] — omit to clear all
     */
    function clear(event) {
        if (event) delete _listeners[event];
        else Object.keys(_listeners).forEach(k => delete _listeners[k]);
    }

    /**
     * List all registered event names (for debugging)
     * @returns {string[]}
     */
    function events() {
        return Object.keys(_listeners).filter(k => _listeners[k].length > 0);
    }

    // ─── Standard ProjectFlow Events ──────────────────────────────
    // 'project:loaded'      → { project }
    // 'project:changed'     → { project, source }   (any data mutation)
    // 'project:saved'       → { id }
    // 'project:switched'    → { id }
    // 'task:selected'       → { task }
    // 'task:updated'        → { task }
    // 'view:changed'        → { view }
    // 'settings:changed'    → { settings }
    // ──────────────────────────────────────────────────────────────

    export const EventBus = { on, off, once, emit, clear, events };
