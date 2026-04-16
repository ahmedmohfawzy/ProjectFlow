/**
 * ProjectFlow™ Task Editor Module
 * © 2026 Ahmed M. Fawzy
 *
 * Task manipulation helpers and utilities.
 * All mutations wrapped with window.PF.mutation() for undo/redo.
 */



  /**
   * Create a new task with default values
   * @param {Object} overrides - Properties to override defaults
   * @returns {Object} New task object
   */
  function createTask(overrides) {
    const defaults = {
      uid: generateUID(),
      name: 'New Task',
      outlineLevel: 1,
      percentComplete: 0,
      predecessors: [],
      resourceNames: [],
      tags: [],
      start: new Date().toISOString().split('T')[0],
      finish: addDaysToDate(new Date(), 1).toISOString().split('T')[0],
      durationDays: 1,
      isExpanded: true,
      isVisible: true,
      cost: 0,
      critical: false
    };

    return Object.assign({}, defaults, overrides);
  }

  /**
   * Deep clone a task with a new UID
   * @param {Object} task - Task to clone
   * @returns {Object} Cloned task with new UID
   */
  function cloneTask(task) {
    if (!task) return null;

    const cloned = JSON.parse(JSON.stringify(task));
    cloned.uid = generateUID();
    cloned.name = (task.name || 'Task') + ' (Copy)';

    return cloned;
  }

  /**
   * Find a task by UID
   * @param {string} uid - Task UID
   * @returns {Object|null} Task object or null
   */
  function findById(uid) {
    if (!window.PF || !uid) return null;

    const project = window.PF.project;
    if (!project || !project.tasks) return null;

    return project.tasks.find(t => t.uid === uid) || null;
  }

  /**
   * Find tasks by name (case-insensitive partial match)
   * @param {string} query - Search query
   * @returns {Object[]} Array of matching tasks
   */
  function findByName(query) {
    if (!window.PF || !query) return [];

    const project = window.PF.project;
    if (!project || !project.tasks) return [];

    const lower = query.toLowerCase();
    return project.tasks.filter(t => (t.name || '').toLowerCase().includes(lower));
  }

  /**
   * Get direct children of a summary task
   * @param {Object} summaryTask - Parent task
   * @returns {Object[]} Array of child tasks
   */
  function getChildren(summaryTask) {
    if (!window.PF || !summaryTask) return [];

    const project = window.PF.project;
    if (!project || !project.tasks) return [];

    const parentLevel = summaryTask.outlineLevel || 0;
    const parentIndex = project.tasks.indexOf(summaryTask);

    if (parentIndex === -1) return [];

    const children = [];
    for (let i = parentIndex + 1; i < project.tasks.length; i++) {
      const task = project.tasks[i];
      const taskLevel = task.outlineLevel || 0;

      // Stop if we hit a task at same level or higher
      if (taskLevel <= parentLevel) break;

      // Only include direct children (one level deeper)
      if (taskLevel === parentLevel + 1) {
        children.push(task);
      }
    }

    return children;
  }

  /**
   * Get ancestor tasks (parents up to root)
   * @param {Object} task - Task
   * @returns {Object[]} Array of ancestor tasks (closest first)
   */
  function getAncestors(task) {
    if (!window.PF || !task) return [];

    const project = window.PF.project;
    if (!project || !project.tasks) return [];

    const ancestors = [];
    const taskLevel = task.outlineLevel || 1;
    const taskIndex = project.tasks.indexOf(task);

    if (taskIndex === -1) return [];

    // Search backwards for parents
    for (let i = taskIndex - 1; i >= 0; i--) {
      const candidate = project.tasks[i];
      const candidateLevel = candidate.outlineLevel || 1;

      // Stop if we go too far back
      if (candidateLevel >= taskLevel) continue;

      // If we find the immediate parent level, add it and continue searching up
      if (candidateLevel === taskLevel - 1) {
        ancestors.unshift(candidate);
      }
    }

    return ancestors;
  }

  /**
   * Calculate working days between dates (excludes weekends)
   * @param {string|Date} startDate - Start date
   * @param {string|Date} endDate - End date
   * @returns {number} Number of working days
   */
  function calculateDuration(startDate, endDate) {
    if (!startDate || !endDate) return 0;

    const start = new Date(startDate);
    const end = new Date(endDate);

    if (end <= start) return 0;

    let workDays = 0;
    const current = new Date(start);

    while (current <= end) {
      const dayOfWeek = current.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        workDays++;
      }
      current.setDate(current.getDate() + 1);
    }

    return workDays;
  }

  /**
   * Shift task dates by delta days
   * @param {Object} task - Task to shift
   * @param {number} deltaDays - Days to shift
   * @returns {Object} Modified task
   */
  function shiftDates(task, deltaDays) {
    if (!task || deltaDays === 0) return task;

    const shifted = Object.assign({}, task);

    if (task.start) {
      shifted.start = addDaysToDate(new Date(task.start), deltaDays).toISOString().split('T')[0];
    }

    if (task.finish) {
      shifted.finish = addDaysToDate(new Date(task.finish), deltaDays).toISOString().split('T')[0];
    }

    // Recalculate duration
    if (shifted.start && shifted.finish) {
      shifted.durationDays = calculateDuration(shifted.start, shifted.finish);
    }

    return shifted;
  }

  /**
   * Shift multiple tasks by delta days (wrapped in mutation)
   * @param {string[]} uids - Array of task UIDs
   * @param {number} deltaDays - Days to shift
   * @returns {boolean} True if successful
   */
  function batchShiftDates(uids, deltaDays) {
    if (!window.PF || !Array.isArray(uids) || deltaDays === 0) return false;

    if (!window.PF.mutation) return false;

    return window.PF.mutation(() => {
      const project = window.PF.project;
      if (!project || !project.tasks) return false;

      uids.forEach(uid => {
        const task = project.tasks.find(t => t.uid === uid);
        if (task) {
          const shifted = shiftDates(task, deltaDays);
          Object.assign(task, shifted);
        }
      });

      return true;
    });
  }

  /**
   * Set task progress percentage
   * @param {string} uid - Task UID
   * @param {number} pct - Progress percentage (0-100)
   * @returns {boolean} True if successful
   */
  function setProgress(uid, pct) {
    if (!window.PF || !uid) return false;

    if (!window.PF.mutation) return false;

    const clamped = Math.max(0, Math.min(100, Math.round(pct || 0)));

    return window.PF.mutation(() => {
      const task = findById(uid);
      if (task) {
        task.percentComplete = clamped;
        return true;
      }
      return false;
    });
  }

  /**
   * Set task resource names
   * @param {string} uid - Task UID
   * @param {string[]} resourceNames - Array of resource names
   * @returns {boolean} True if successful
   */
  function setResource(uid, resourceNames) {
    if (!window.PF || !uid || !Array.isArray(resourceNames)) return false;

    if (!window.PF.mutation) return false;

    return window.PF.mutation(() => {
      const task = findById(uid);
      if (task) {
        task.resourceNames = resourceNames;
        return true;
      }
      return false;
    });
  }

  /**
   * Apply a tag to a task
   * @param {string} uid - Task UID
   * @param {string} tagName - Tag to apply
   * @returns {boolean} True if successful
   */
  function applyTag(uid, tagName) {
    if (!window.PF || !uid || !tagName) return false;

    if (!window.PF.mutation) return false;

    return window.PF.mutation(() => {
      const task = findById(uid);
      if (task) {
        if (!task.tags) task.tags = [];
        if (!task.tags.includes(tagName)) {
          task.tags.push(tagName);
        }
        return true;
      }
      return false;
    });
  }

  /**
   * Remove a tag from a task
   * @param {string} uid - Task UID
   * @param {string} tagName - Tag to remove
   * @returns {boolean} True if successful
   */
  function removeTag(uid, tagName) {
    if (!window.PF || !uid || !tagName) return false;

    if (!window.PF.mutation) return false;

    return window.PF.mutation(() => {
      const task = findById(uid);
      if (task && task.tags) {
        const idx = task.tags.indexOf(tagName);
        if (idx >= 0) {
          task.tags.splice(idx, 1);
          return true;
        }
      }
      return false;
    });
  }

  /**
   * Get hierarchical path to task (e.g., "Phase 1 > Task Name")
   * @param {Object} task - Task
   * @returns {string} Path string
   */
  function getTaskPath(task) {
    if (!task) return '';

    const ancestors = getAncestors(task);
    const names = ancestors.map(a => a.name || 'Unnamed').concat([task.name || 'Unnamed']);

    return names.join(' > ');
  }

  /**
   * Get summary of late tasks
   * @param {Object} project - The project
   * @returns {Object[]} Array of {name, daysLate, uid} sorted by daysLate descending
   */
  function getLateTasksSummary(project) {
    if (!project || !project.tasks) return [];

    const now = new Date();
    const late = [];

    project.tasks.forEach(task => {
      if (!task.name || !task.finish) return;

      // Only include incomplete tasks
      if (task.percentComplete === 100) return;

      const finish = new Date(task.finish);
      if (finish < now) {
        const daysLate = Math.ceil((now - finish) / (1000 * 60 * 60 * 24));
        late.push({
          name: task.name,
          daysLate,
          uid: task.uid
        });
      }
    });

    // Sort by daysLate descending
    late.sort((a, b) => b.daysLate - a.daysLate);

    return late;
  }

  // ==================== Helper Functions ====================

  /**
   * Generate unique ID
   * @private
   */
  function generateUID() {
    return 'uid_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
  }

  /**
   * Add days to a date
   * @private
   */
  function addDaysToDate(date, days) {
    const result = new Date(date);
    result.setDate(result.getDate() + days);
    return result;
  }

  export const TaskEditor = {
    createTask,
    cloneTask,
    findById,
    findByName,
    getChildren,
    getAncestors,
    calculateDuration,
    shiftDates,
    batchShiftDates,
    setProgress,
    setResource,
    applyTag,
    removeTag,
    getTaskPath,
    getLateTasksSummary
  };
