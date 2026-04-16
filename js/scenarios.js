/**
 * ProjectFlow™ © 2026 Ahmed M. Fawzy
 *
 * What-If Scenarios System
 * Enables users to create, manage, and compare project snapshots
 * for scenario planning and analysis.
 */
  const STORAGE_KEY = 'pf_scenarios';
  const MAX_SCENARIOS = 20;

  /**
   * Deep clone via JSON serialization, stripping attachments
   * @param {Object} obj - Object to clone
   * @returns {Object} Cloned object without attachment blobs
   */
  function deepCloneWithoutAttachments(obj) {
    const serialized = JSON.stringify(obj, (key, value) => {
      // Strip attachment data to keep localStorage size small
      if (key === 'attachments' && Array.isArray(value)) {
        return [];
      }
      return value;
    });
    return JSON.parse(serialized);
  }

  /**
   * Retrieve all scenarios from localStorage
   * @returns {Array} Array of scenario objects
   */
  function getScenarios() {
    try {
      const stored = localStorage.getItem(STORAGE_KEY);
      return stored ? JSON.parse(stored) : [];
    } catch (e) {
      console.error('Error reading scenarios from localStorage:', e);
      return [];
    }
  }

  /**
   * Save scenarios to localStorage with eviction policy
   * @param {Array} scenarios - Array of scenario objects
   */
  function saveScenarios(scenarios) {
    try {
      // Keep only the most recent MAX_SCENARIOS
      const sorted = scenarios.sort((a, b) => b.timestamp - a.timestamp);
      const toKeep = sorted.slice(0, MAX_SCENARIOS);
      localStorage.setItem(STORAGE_KEY, JSON.stringify(toKeep));
    } catch (e) {
      console.error('Error saving scenarios to localStorage:', e);
    }
  }

  /**
   * Generate unique scenario ID
   * @returns {string} Unique ID
   */
  function generateId() {
    return 'scenario_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
  }

  /**
   * Save current project state as a scenario
   * @param {Object} project - Project object
   * @param {string} name - Scenario name
   * @returns {Object} Saved scenario {id, name, timestamp, projectSnap}
   */
  function save(project, name) {
    if (!project) {
      throw new Error('Project is required');
    }
    if (!name || typeof name !== 'string') {
      throw new Error('Scenario name is required and must be a string');
    }

    const scenario = {
      id: generateId(),
      name: name.trim(),
      timestamp: Date.now(),
      projectSnap: deepCloneWithoutAttachments(project)
    };

    const scenarios = getScenarios();
    scenarios.push(scenario);
    saveScenarios(scenarios);

    return {
      id: scenario.id,
      name: scenario.name,
      timestamp: scenario.timestamp,
      projectSnap: scenario.projectSnap
    };
  }

  /**
   * List all saved scenarios (metadata only)
   * @returns {Array} Array of scenario metadata {id, name, timestamp, taskCount, projectName}
   */
  function list() {
    const scenarios = getScenarios();
    return scenarios.map(scenario => ({
      id: scenario.id,
      name: scenario.name,
      timestamp: scenario.timestamp,
      taskCount: scenario.projectSnap?.tasks ? scenario.projectSnap.tasks.length : 0,
      projectName: scenario.projectSnap?.name || 'Unnamed Project'
    }));
  }

  /**
   * Load full project snapshot for a scenario
   * @param {string} id - Scenario ID
   * @returns {Object|null} Full project snapshot or null if not found
   */
  function load(id) {
    const scenarios = getScenarios();
    const scenario = scenarios.find(s => s.id === id);
    return scenario ? scenario.projectSnap : null;
  }

  /**
   * Delete scenario by ID
   * @param {string} id - Scenario ID
   * @returns {boolean} True if deleted, false if not found
   */
  function deleteFn(id) {
    const scenarios = getScenarios();
    const filtered = scenarios.filter(s => s.id !== id);
    if (filtered.length < scenarios.length) {
      saveScenarios(filtered);
      return true;
    }
    return false;
  }

  /**
   * Compare two project snapshots
   * @param {Object} snapA - First project snapshot
   * @param {Object} snapB - Second project snapshot
   * @returns {Object} Comparison result {added, removed, changed, summary}
   */
  function compare(snapA, snapB) {
    const tasksA = (snapA?.tasks || []).reduce((acc, t) => {
      acc[t.uid] = t;
      return acc;
    }, {});

    const tasksB = (snapB?.tasks || []).reduce((acc, t) => {
      acc[t.uid] = t;
      return acc;
    }, {});

    const added = [];
    const removed = [];
    const changed = [];

    // Find added and changed tasks
    Object.keys(tasksB).forEach(uid => {
      if (!tasksA[uid]) {
        added.push(tasksB[uid]);
      } else {
        const taskA = tasksA[uid];
        const taskB = tasksB[uid];
        const isChanged = (
          taskA.name !== taskB.name ||
          taskA.duration !== taskB.duration ||
          taskA.percentComplete !== taskB.percentComplete ||
          taskA.start !== taskB.start ||
          taskA.finish !== taskB.finish ||
          taskA.critical !== taskB.critical
        );
        if (isChanged) {
          changed.push({
            uid: uid,
            before: taskA,
            after: taskB
          });
        }
      }
    });

    // Find removed tasks
    Object.keys(tasksA).forEach(uid => {
      if (!tasksB[uid]) {
        removed.push(tasksA[uid]);
      }
    });

    // Calculate summary metrics
    const durationA = snapA?.tasks?.reduce((sum, t) => sum + (t.duration || 0), 0) || 0;
    const durationB = snapB?.tasks?.reduce((sum, t) => sum + (t.duration || 0), 0) || 0;
    const progressA = snapA?.tasks?.reduce((sum, t) => sum + (t.percentComplete || 0), 0) || 0;
    const progressB = snapB?.tasks?.reduce((sum, t) => sum + (t.percentComplete || 0), 0) || 0;
    const criticalA = snapA?.tasks?.filter(t => t.critical).length || 0;
    const criticalB = snapB?.tasks?.filter(t => t.critical).length || 0;
    const taskCountA = snapA?.tasks?.length || 0;
    const taskCountB = snapB?.tasks?.length || 0;
    const lateA = snapA?.tasks?.filter(t => t.isLate).length || 0;
    const lateB = snapB?.tasks?.filter(t => t.isLate).length || 0;

    return {
      added,
      removed,
      changed,
      summary: {
        durationDeltaDays: durationB - durationA,
        progressDelta: (progressB - progressA) / Math.max(taskCountB, 1),
        criticalDelta: criticalB - criticalA,
        lateDelta: lateB - lateA
      }
    };
  }

  /**
   * Render comparison UI into a container
   * @param {HTMLElement} container - Target container
   * @param {Object} snapA - First project snapshot
   * @param {Object} snapB - Second project snapshot
   * @param {string} nameA - Name of first scenario
   * @param {string} nameB - Name of second scenario
   */
  function renderCompareUI(container, snapA, snapB, nameA, nameB) {
    if (!container || !(container instanceof HTMLElement)) {
      throw new Error('Invalid container element');
    }

    // Clear container
    container.innerHTML = '';

    // Get comparison data
    const comp = compare(snapA, snapB);

    // Create wrapper
    const wrapper = document.createElement('div');
    wrapper.className = 'scenarios-compare-wrapper';

    // Summary bar
    const summaryBar = document.createElement('div');
    summaryBar.className = 'scenarios-summary-bar';
    summaryBar.style.cssText = 'padding: 12px; background: #f5f5f5; border-bottom: 1px solid #ddd; margin-bottom: 12px;';

    const summaryText = document.createElement('div');
    summaryText.className = 'scenarios-summary-text';
    summaryText.style.cssText = 'font-size: 14px; line-height: 1.5;';

    const summaryLines = [
      `Added: ${comp.added.length} tasks`,
      `Removed: ${comp.removed.length} tasks`,
      `Changed: ${comp.changed.length} tasks`,
      `Duration delta: ${comp.summary.durationDeltaDays > 0 ? '+' : ''}${comp.summary.durationDeltaDays} days`,
      `Critical delta: ${comp.summary.criticalDelta > 0 ? '+' : ''}${comp.summary.criticalDelta}`,
      `Late delta: ${comp.summary.lateDelta > 0 ? '+' : ''}${comp.summary.lateDelta}`
    ];

    summaryLines.forEach(line => {
      const p = document.createElement('p');
      p.style.cssText = 'margin: 4px 0;';
      p.textContent = line;
      summaryText.appendChild(p);
    });

    summaryBar.appendChild(summaryText);
    wrapper.appendChild(summaryBar);

    // Comparison table
    const table = document.createElement('table');
    table.className = 'scenarios-compare-table';
    table.style.cssText = 'width: 100%; border-collapse: collapse; font-size: 13px;';

    // Table header
    const thead = document.createElement('thead');
    const headerRow = document.createElement('tr');
    headerRow.style.cssText = 'background: #e8e8e8; border-bottom: 2px solid #999;';

    const headers = ['Task UID', 'Name', 'Duration (days)', 'Progress (%)', 'Status'];
    headers.forEach(headerText => {
      const th = document.createElement('th');
      th.style.cssText = 'padding: 8px; text-align: left; font-weight: bold;';
      th.textContent = headerText;
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
    table.appendChild(thead);

    const tbody = document.createElement('tbody');

    // Added tasks (green)
    if (comp.added.length > 0) {
      const addedHeader = document.createElement('tr');
      addedHeader.style.cssText = 'background: #e8f5e9; font-weight: bold;';
      const addedCell = document.createElement('td');
      addedCell.colSpan = 5;
      addedCell.style.cssText = 'padding: 8px; color: #2e7d32;';
      addedCell.textContent = `ADDED (${comp.added.length} tasks)`;
      addedHeader.appendChild(addedCell);
      tbody.appendChild(addedHeader);

      comp.added.forEach(task => {
        const row = document.createElement('tr');
        row.style.cssText = 'background: #f1f8f6; border-bottom: 1px solid #ddd;';

        const cells = [
          task.uid || '',
          task.name || '',
          task.duration || '',
          task.percentComplete || '0',
          task.critical ? 'Critical' : 'Normal'
        ];

        cells.forEach(cellText => {
          const td = document.createElement('td');
          td.style.cssText = 'padding: 8px; border-right: 1px solid #eee;';
          td.textContent = String(cellText);
          row.appendChild(td);
        });

        tbody.appendChild(row);
      });
    }

    // Removed tasks (red)
    if (comp.removed.length > 0) {
      const removedHeader = document.createElement('tr');
      removedHeader.style.cssText = 'background: #ffebee; font-weight: bold;';
      const removedCell = document.createElement('td');
      removedCell.colSpan = 5;
      removedCell.style.cssText = 'padding: 8px; color: #c62828;';
      removedCell.textContent = `REMOVED (${comp.removed.length} tasks)`;
      removedHeader.appendChild(removedCell);
      tbody.appendChild(removedHeader);

      comp.removed.forEach(task => {
        const row = document.createElement('tr');
        row.style.cssText = 'background: #fdeaea; border-bottom: 1px solid #ddd;';

        const cells = [
          task.uid || '',
          task.name || '',
          task.duration || '',
          task.percentComplete || '0',
          task.critical ? 'Critical' : 'Normal'
        ];

        cells.forEach(cellText => {
          const td = document.createElement('td');
          td.style.cssText = 'padding: 8px; border-right: 1px solid #eee;';
          td.textContent = String(cellText);
          row.appendChild(td);
        });

        tbody.appendChild(row);
      });
    }

    // Changed tasks (yellow)
    if (comp.changed.length > 0) {
      const changedHeader = document.createElement('tr');
      changedHeader.style.cssText = 'background: #fffde7; font-weight: bold;';
      const changedCell = document.createElement('td');
      changedCell.colSpan = 5;
      changedCell.style.cssText = 'padding: 8px; color: #f57f17;';
      changedCell.textContent = `CHANGED (${comp.changed.length} tasks)`;
      changedHeader.appendChild(changedCell);
      tbody.appendChild(changedHeader);

      comp.changed.forEach(item => {
        const before = item.before;
        const after = item.after;

        // Before row
        const beforeRow = document.createElement('tr');
        beforeRow.style.cssText = 'background: #fffce6; border-bottom: 1px solid #ddd;';

        const beforeCells = [
          before.uid || '',
          before.name || '',
          before.duration || '',
          before.percentComplete || '0',
          before.critical ? 'Critical' : 'Normal'
        ];

        beforeCells.forEach(cellText => {
          const td = document.createElement('td');
          td.style.cssText = 'padding: 8px; border-right: 1px solid #eee; color: #666;';
          td.textContent = String(cellText);
          beforeRow.appendChild(td);
        });

        tbody.appendChild(beforeRow);

        // After row
        const afterRow = document.createElement('tr');
        afterRow.style.cssText = 'background: #fffff9; border-bottom: 2px solid #ddd;';

        const afterCells = [
          after.uid || '',
          after.name || '',
          after.duration || '',
          after.percentComplete || '0',
          after.critical ? 'Critical' : 'Normal'
        ];

        afterCells.forEach(cellText => {
          const td = document.createElement('td');
          td.style.cssText = 'padding: 8px; border-right: 1px solid #eee; font-weight: bold; color: #333;';
          td.textContent = String(cellText);
          afterRow.appendChild(td);
        });

        tbody.appendChild(afterRow);
      });
    }

    table.appendChild(tbody);
    wrapper.appendChild(table);
    container.appendChild(wrapper);
  }

  /**
   * Render scenarios management panel
   * @param {HTMLElement} container - Target container
   * @param {Object} project - Current project object
   * @param {Object} callbacks - {onLoad(snapshot), onCompare(snapA, snapB, nameA, nameB)}
   */
  function renderScenariosPanel(container, project, callbacks) {
    if (!container || !(container instanceof HTMLElement)) {
      throw new Error('Invalid container element');
    }

    const cb = callbacks || {};

    // Clear container
    container.innerHTML = '';

    const panel = document.createElement('div');
    panel.className = 'scenarios-panel';
    panel.style.cssText = 'padding: 12px; background: #fafafa; border-radius: 4px;';

    // Header
    const header = document.createElement('h3');
    header.style.cssText = 'margin: 0 0 12px 0; font-size: 16px; color: #333;';
    header.textContent = 'What-If Scenarios';
    panel.appendChild(header);

    // Controls section
    const controlsDiv = document.createElement('div');
    controlsDiv.style.cssText = 'margin-bottom: 12px; padding-bottom: 12px; border-bottom: 1px solid #ddd;';

    // Save button
    const saveBtn = document.createElement('button');
    saveBtn.textContent = 'Save Current Scenario';
    saveBtn.style.cssText = 'padding: 6px 12px; margin-right: 6px; background: #4CAF50; color: white; border: none; border-radius: 3px; cursor: pointer; font-size: 12px;';
    saveBtn.addEventListener('click', () => {
      const name = prompt('Enter scenario name:');
      if (name) {
        try {
          save(project, name);
          refreshScenariosList();
        } catch (e) {
          alert('Error saving scenario: ' + e.message);
        }
      }
    });

    // Load button
    const loadBtn = document.createElement('button');
    loadBtn.textContent = 'Load Selected';
    loadBtn.style.cssText = 'padding: 6px 12px; margin-right: 6px; background: #2196F3; color: white; border: none; border-radius: 3px; cursor: pointer; font-size: 12px;';
    loadBtn.addEventListener('click', () => {
      const selected = scenariosList.querySelector('input[type="radio"]:checked');
      if (selected) {
        const snapshot = load(selected.value);
        if (snapshot && cb.onLoad) {
          cb.onLoad(snapshot);
        }
      } else {
        alert('Please select a scenario to load');
      }
    });

    // Delete button
    const deleteBtn = document.createElement('button');
    deleteBtn.textContent = 'Delete Selected';
    deleteBtn.style.cssText = 'padding: 6px 12px; margin-right: 6px; background: #f44336; color: white; border: none; border-radius: 3px; cursor: pointer; font-size: 12px;';
    deleteBtn.addEventListener('click', () => {
      const selected = scenariosList.querySelector('input[type="radio"]:checked');
      if (selected) {
        if (confirm('Delete this scenario?')) {
          deleteFn(selected.value);
          refreshScenariosList();
        }
      } else {
        alert('Please select a scenario to delete');
      }
    });

    // Compare button
    const compareBtn = document.createElement('button');
    compareBtn.textContent = 'Compare Selected (2)';
    compareBtn.style.cssText = 'padding: 6px 12px; background: #FF9800; color: white; border: none; border-radius: 3px; cursor: pointer; font-size: 12px;';
    compareBtn.addEventListener('click', () => {
      const checkboxes = scenariosList.querySelectorAll('input[type="checkbox"]:checked');
      if (checkboxes.length === 2) {
        const id1 = checkboxes[0].value;
        const id2 = checkboxes[1].value;
        const snap1 = load(id1);
        const snap2 = load(id2);
        const name1 = checkboxes[0].dataset.name;
        const name2 = checkboxes[1].dataset.name;
        if (snap1 && snap2 && cb.onCompare) {
          cb.onCompare(snap1, snap2, name1, name2);
        }
      } else {
        alert('Please select exactly 2 scenarios to compare');
      }
    });

    controlsDiv.appendChild(saveBtn);
    controlsDiv.appendChild(loadBtn);
    controlsDiv.appendChild(deleteBtn);
    controlsDiv.appendChild(compareBtn);
    panel.appendChild(controlsDiv);

    // Scenarios list
    const listLabel = document.createElement('label');
    listLabel.style.cssText = 'display: block; margin-bottom: 8px; font-weight: bold; font-size: 12px; color: #555;';
    listLabel.textContent = 'Saved Scenarios:';
    panel.appendChild(listLabel);

    const scenariosList = document.createElement('div');
    scenariosList.className = 'scenarios-list';
    scenariosList.style.cssText = 'max-height: 300px; overflow-y: auto; border: 1px solid #ddd; border-radius: 3px; background: white; padding: 6px;';

    function refreshScenariosList() {
      scenariosList.innerHTML = '';
      const scenarios = list();

      if (scenarios.length === 0) {
        const empty = document.createElement('div');
        empty.style.cssText = 'padding: 12px; color: #999; font-size: 12px; text-align: center;';
        empty.textContent = 'No scenarios saved yet';
        scenariosList.appendChild(empty);
      } else {
        scenarios.forEach(scenario => {
          const itemDiv = document.createElement('div');
          itemDiv.style.cssText = 'padding: 6px; border-bottom: 1px solid #f0f0f0; display: flex; align-items: center; gap: 8px;';

          const radio = document.createElement('input');
          radio.type = 'radio';
          radio.name = 'scenario-select';
          radio.value = scenario.id;
          radio.style.cssText = 'cursor: pointer;';

          const checkbox = document.createElement('input');
          checkbox.type = 'checkbox';
          checkbox.value = scenario.id;
          checkbox.dataset.name = scenario.name;
          checkbox.style.cssText = 'cursor: pointer;';

          const label = document.createElement('label');
          label.style.cssText = 'flex: 1; cursor: pointer; font-size: 12px; margin: 0;';
          const nameSpan = document.createElement('span');
          nameSpan.textContent = scenario.name;
          nameSpan.style.cssText = 'font-weight: 500;';

          const metaSpan = document.createElement('span');
          metaSpan.style.cssText = 'display: block; font-size: 11px; color: #999; margin-top: 2px;';
          const date = new Date(scenario.timestamp).toLocaleString();
          metaSpan.textContent = `${scenario.projectName} • ${scenario.taskCount} tasks • ${date}`;

          label.appendChild(nameSpan);
          label.appendChild(metaSpan);

          itemDiv.appendChild(radio);
          itemDiv.appendChild(checkbox);
          itemDiv.appendChild(label);
          scenariosList.appendChild(itemDiv);
        });
      }
    }

    refreshScenariosList();
    panel.appendChild(scenariosList);
    container.appendChild(panel);
  }

  // Public API
  export const ScenariosManager = {
    save,
    list,
    load,
    delete: deleteFn,
    compare,
    renderCompareUI,
    renderScenariosPanel
  };
