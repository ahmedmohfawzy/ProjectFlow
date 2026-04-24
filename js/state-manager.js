/**
 * ProjectFlow™ State Manager Module
 * © 2026 Ahmed M. Fawzy
 *
 * Provides a clean API over window.PF state management.
 * Handles project data, settings, selection, and undo operations.
 */



  // Private constants
  const SETTINGS_KEY = 'pf_settings';
  const DEFAULT_SETTINGS = {
    theme: 'light',
    viewMode: 'timeline',
    autoSave: true,
    showWeekends: false,
    groupBy: 'phase'
  };

  /**
   * Get the current project object
   * @returns {Object|null} Project object or null if not available
   */
  function getProject() {
    if (!window.PF) return null;
    return window.PF.project;
  }

  /**
   * Set the project object
   * @param {Object} project - The project to set
   */
  function setProject(project) {
    if (!window.PF) return;
    window.PF.project = project;
  }

  /**
   * Get all settings
   * @returns {Object} Settings object
   */
  function getSettings() {
    if (!window.PF) return DEFAULT_SETTINGS;
    return window.PF.settings || DEFAULT_SETTINGS;
  }

  /**
   * Update a single setting and persist to localStorage
   * @param {string} key - Setting key
   * @param {*} value - Setting value
   */
  function updateSetting(key, value) {
    if (!window.PF) return;
    const current = getSettings();
    current[key] = value;
    try {
      localStorage.setItem(SETTINGS_KEY, JSON.stringify(current));
    } catch (e) {
      console.warn('Failed to save settings to localStorage:', e);
    }
  }

  /**
   * Get currently selected task IDs
   * @returns {Set<string>|null} Set of selected task UIDs or null
   */
  function getSelectedIds() {
    if (!window.PF) return new Set();
    return window.PF.selectedTaskIds || new Set();
  }

  /**
   * Select a task by UID
   * @param {string} uid - Task UID
   */
  function selectTask(uid) {
    if (!window.PF) return;
    const ids = getSelectedIds();
    ids.add(uid);
  }

  /**
   * Deselect a task by UID
   * @param {string} uid - Task UID
   */
  function deselectTask(uid) {
    if (!window.PF) return;
    const ids = getSelectedIds();
    ids.delete(uid);
  }

  /**
   * Clear all selections
   */
  function clearSelection() {
    if (!window.PF) return;
    const ids = getSelectedIds();
    ids.clear();
  }

  /**
   * Select a range of task UIDs
   * @param {string[]} uids - Array of task UIDs
   */
  function selectRange(uids) {
    if (!window.PF || !Array.isArray(uids)) return;
    clearSelection();
    uids.forEach(uid => selectTask(uid));
  }

  /**
   * Wrap a function with undo/redo and auto-refresh
   * Equivalent to mutation() in app.js
   * @param {Function} fn - Function to execute
   * @returns {*} Return value of fn
   */
  function undoable(fn) {
    if (!window.PF) return null;

    // Save state before mutation
    if (window.PF.saveUndoState) {
      window.PF.saveUndoState();
    }

    // Execute the mutation
    const result = fn();

    // Recalculate dependent values
    if (window.PF.recalculate) {
      window.PF.recalculate();
    }

    // Render UI
    if (window.PF.renderAll) {
      window.PF.renderAll();
    }

    // Auto-save to storage
    if (window.PF.autoSave) {
      window.PF.autoSave();
    }

    return result;
  }

  /**
   * Check if a project is currently open
   * @returns {boolean} True if project exists and has tasks
   */
  function isProjectOpen() {
    const project = getProject();
    return !!(project && project.tasks && project.tasks.length > 0);
  }

  /**
   * Get cached analytics data
   * @returns {Object|null} Analytics cache object or null
   */
  function getAnalytics() {
    if (!window.PF || !window.PF.ProjectAnalytics) return null;
    return window.PF.ProjectAnalytics.getCache ? window.PF.ProjectAnalytics.getCache() : null;
  }

  // Public API
  export const StateManager = Object.freeze({
    getProject,
    setProject,
    getSettings,
    updateSetting,
    getSelectedIds,
    selectTask,
    deselectTask,
    clearSelection,
    selectRange,
    undoable,
    isProjectOpen,
    getAnalytics
  });
