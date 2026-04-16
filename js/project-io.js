/**
 * ProjectFlow™ Project I/O Module
 * © 2026 Ahmed M. Fawzy
 *
 * Import/Export utilities for project data.
 * Handles JSON, CSV, and markdown formats with clipboard support.
 */



  /**
   * Export project as JSON and trigger download
   * @param {Object} project - The project to export
   * @returns {boolean} True if export successful
   */
  function exportJSON(project) {
    if (!project) return false;

    try {
      const json = JSON.stringify(project, null, 2);
      downloadFile(json, `${project.name || 'project'}.json`, 'application/json');
      return true;
    } catch (e) {
      console.error('JSON export failed:', e);
      return false;
    }
  }

  /**
   * Import project from JSON string
   * @param {string} jsonString - JSON string to import
   * @returns {boolean} True if import successful
   */
  function importJSON(jsonString) {
    if (!window.PF || !jsonString) return false;

    try {
      const project = JSON.parse(jsonString);

      // Validate basic structure
      if (!project.name || !Array.isArray(project.tasks)) {
        console.error('Invalid project format');
        return false;
      }

      // Set the project
      window.PF.project = project;

      // Reindex tasks
      if (window.PF.reindexTasks) {
        window.PF.reindexTasks();
      }

      // Recalculate
      if (window.PF.recalculate) {
        window.PF.recalculate();
      }

      // Render UI
      if (window.PF.renderAll) {
        window.PF.renderAll();
      }

      // Auto-save
      if (window.PF.autoSave) {
        window.PF.autoSave();
      }

      return true;
    } catch (e) {
      console.error('JSON import failed:', e);
      return false;
    }
  }

  /**
   * Export project as markdown report
   * @param {Object} project - The project to export
   * @returns {string} Markdown text
   */
  function exportMarkdownReport(project) {
    if (!project) return '';

    const lines = [];

    // Header
    lines.push(`# ${project.name || 'Project Report'}`);
    lines.push('');

    // Project info
    if (project.description) {
      lines.push(project.description);
      lines.push('');
    }

    if (project.startDate || project.finishDate) {
      lines.push('## Project Timeline');
      if (project.startDate) {
        lines.push(`- **Start:** ${formatDateMarkdown(project.startDate)}`);
      }
      if (project.finishDate) {
        lines.push(`- **Finish:** ${formatDateMarkdown(project.finishDate)}`);
      }
      lines.push('');
    }

    // Tasks table
    if (project.tasks && project.tasks.length > 0) {
      lines.push('## Tasks');
      lines.push('');
      lines.push('| Name | Duration | Progress | Status |');
      lines.push('|------|----------|----------|--------|');

      project.tasks.forEach(task => {
        const duration = task.durationDays ? `${task.durationDays}d` : '—';
        const progress = `${Math.round(task.percentComplete || 0)}%`;
        const status = task.percentComplete === 100 ? '✓ Complete' : (task.percentComplete > 0 ? 'In Progress' : 'Not Started');
        lines.push(`| ${escape(task.name || 'Unnamed')} | ${duration} | ${progress} | ${status} |`);
      });
      lines.push('');
    }

    // Analytics summary
    const stats = getProjectStats(project);
    if (stats) {
      lines.push('## Summary');
      lines.push(`- **Total Tasks:** ${stats.total}`);
      lines.push(`- **Complete:** ${stats.complete}`);
      lines.push(`- **In Progress:** ${stats.inProgress}`);
      lines.push(`- **Late Tasks:** ${stats.late}`);
      lines.push(`- **Overall Progress:** ${Math.round(stats.overallProgress)}%`);
      if (stats.totalCost) {
        lines.push(`- **Total Cost:** ${formatCurrency(stats.totalCost)}`);
      }
      if (stats.daysRemaining !== null) {
        lines.push(`- **Days Remaining:** ${stats.daysRemaining}`);
      }
      lines.push('');
    }

    return lines.join('\n');
  }

  /**
   * Copy plain text project summary to clipboard
   * @param {Object} project - The project
   * @returns {boolean} True if copy successful
   */
  function copyProjectSummaryText(project) {
    if (!project) return false;

    try {
      const stats = getProjectStats(project);
      const lines = [
        `PROJECT: ${project.name || 'Untitled'}`,
        ''
      ];

      if (stats) {
        lines.push(`Progress: ${Math.round(stats.overallProgress)}%`);
        lines.push(`Tasks: ${stats.total} total (${stats.complete} done, ${stats.inProgress} in progress)`);
        if (stats.late > 0) {
          lines.push(`Late: ${stats.late} tasks`);
        }
        if (stats.critical > 0) {
          lines.push(`Critical: ${stats.critical} tasks`);
        }
      }

      const text = lines.join('\n');
      return copyToClipboard(text);
    } catch (e) {
      console.error('Failed to copy summary:', e);
      return false;
    }
  }

  /**
   * Generate CSV representation of tasks
   * @param {Object} project - The project
   * @returns {string} CSV text
   */
  function generateShareableCSV(project) {
    if (!project || !project.tasks) return '';

    const rows = [];

    // Header row
    rows.push(['Name', 'Duration (days)', 'Start', 'Finish', 'Progress %', 'Status', 'Resources', 'Tags'].map(escapeCSV).join(','));

    // Data rows
    project.tasks.forEach(task => {
      const row = [
        escapeCSV(task.name || ''),
        task.durationDays || '',
        formatDateISO(task.start),
        formatDateISO(task.finish),
        Math.round(task.percentComplete || 0),
        getTaskStatus(task),
        escapeCSV((task.resourceNames || []).join('; ')),
        escapeCSV((task.tags || []).join('; '))
      ];
      rows.push(row.join(','));
    });

    return rows.join('\n');
  }

  /**
   * Detect CSV format from header row
   * @param {string[]} headerRow - Array of column names
   * @returns {string} 'jira' | 'ms-project' | 'generic'
   */
  function detectCSVFormat(headerRow) {
    if (!Array.isArray(headerRow)) return 'generic';

    const lower = headerRow.map(h => (h || '').toLowerCase());

    // Jira format detection
    if (lower.includes('key') && lower.includes('summary')) {
      return 'jira';
    }

    // MS Project format detection
    if (lower.includes('id') && (lower.includes('task name') || lower.includes('task'))) {
      return 'ms-project';
    }

    return 'generic';
  }

  /**
   * Get project statistics
   * @param {Object} project - The project
   * @returns {Object} Statistics object
   */
  function getProjectStats(project) {
    if (!project || !project.tasks) {
      return {
        total: 0,
        complete: 0,
        inProgress: 0,
        late: 0,
        critical: 0,
        overallProgress: 0,
        totalCost: 0,
        earnedCost: 0,
        daysRemaining: null
      };
    }

    const tasks = project.tasks;
    const now = new Date();

    let total = 0;
    let complete = 0;
    let inProgress = 0;
    let late = 0;
    let critical = 0;
    let totalCost = 0;
    let earnedCost = 0;

    tasks.forEach(task => {
      // Skip summary tasks in counts (optional logic)
      if (!task.name) return;

      total++;

      const pct = task.percentComplete || 0;
      if (pct === 100) {
        complete++;
      } else if (pct > 0) {
        inProgress++;
      }

      // Check if late
      if (task.finish && new Date(task.finish) < now && pct < 100) {
        late++;
      }

      // Check if critical (based on critical flag or predecessor count)
      if (task.critical || (task.predecessors && task.predecessors.length > 2)) {
        critical++;
      }

      // Sum costs
      if (task.cost) {
        totalCost += parseFloat(task.cost) || 0;
        earnedCost += (parseFloat(task.cost) || 0) * (pct / 100);
      }
    });

    // Calculate overall progress
    const overallProgress = total > 0 ? (complete / total) * 100 : 0;

    // Days remaining
    let daysRemaining = null;
    if (project.finishDate) {
      const finish = new Date(project.finishDate);
      const daysDiff = Math.ceil((finish - now) / (1000 * 60 * 60 * 24));
      daysRemaining = Math.max(-999, daysDiff); // Cap at -999 for display
    }

    return {
      total,
      complete,
      inProgress,
      late,
      critical,
      overallProgress,
      totalCost,
      earnedCost,
      daysRemaining
    };
  }

  // ==================== Helper Functions ====================

  /**
   * Download file to user's device
   * @private
   */
  function downloadFile(content, filename, mimeType) {
    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }

  /**
   * Copy text to clipboard
   * @private
   */
  function copyToClipboard(text) {
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(text);
        return true;
      } else {
        // Fallback for older browsers
        const textarea = document.createElement('textarea');
        textarea.value = text;
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        document.body.removeChild(textarea);
        return true;
      }
    } catch (e) {
      console.error('Clipboard copy failed:', e);
      return false;
    }
  }

  /**
   * Escape CSV field value
   * @private
   */
  function escapeCSV(value) {
    if (!value) return '';
    const str = String(value);
    if (str.includes(',') || str.includes('"') || str.includes('\n')) {
      return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
  }

  /**
   * Format date for markdown
   * @private
   */
  function formatDateMarkdown(dateStr) {
    if (!dateStr) return '—';
    try {
      const date = new Date(dateStr);
      return date.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
    } catch {
      return dateStr;
    }
  }

  /**
   * Format date as ISO string
   * @private
   */
  function formatDateISO(dateStr) {
    if (!dateStr) return '';
    try {
      return new Date(dateStr).toISOString().split('T')[0];
    } catch {
      return '';
    }
  }

  /**
   * Format currency value
   * @private
   */
  function formatCurrency(amount) {
    if (typeof amount !== 'number') return '$0';
    return '$' + amount.toLocaleString('en-US', { maximumFractionDigits: 0 });
  }

  /**
   * Get task status string
   * @private
   */
  function getTaskStatus(task) {
    const pct = task.percentComplete || 0;
    if (pct === 100) return 'Complete';
    if (pct > 0) return 'In Progress';
    return 'Not Started';
  }

  /**
   * HTML escape helper
   * @private
   */
  function escape(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // Public API
  export const ProjectIO = Object.freeze({
    exportJSON,
    importJSON,
    exportMarkdownReport,
    copyProjectSummaryText,
    generateShareableCSV,
    detectCSVFormat,
    getProjectStats
  });
