/**
 * ProjectFlow™ UI Helpers Module
 * © 2026 Ahmed M. Fawzy
 *
 * Pure UI utilities for formatting, DOM creation, and user interactions.
 * No state dependencies - safe to use independently.
 */



  /**
   * Format date with various formats
   * @param {string|Date} date - Date to format
   * @param {string} format - 'short' | 'iso' | 'long' | 'rel'
   * @returns {string} Formatted date string
   */
  function fmtDate(date, format) {
    if (!date) return '—';

    const d = new Date(date);
    if (isNaN(d.getTime())) return '—';

    switch (format) {
      case 'short':
        // "Jan 5"
        return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });

      case 'iso':
        // "2026-01-05"
        return d.toISOString().split('T')[0];

      case 'long':
        // "5 January 2026"
        return d.toLocaleDateString('en-US', { day: 'numeric', month: 'long', year: 'numeric' });

      case 'rel':
        // "3d ago" / "in 5d" / "today" / "tomorrow"
        return formatRelativeDate(d);

      default:
        return d.toLocaleDateString();
    }
  }

  /**
   * Format duration in days
   * @param {number} days - Number of days
   * @returns {string} Formatted duration (e.g., "5d", "2w 1d", "3m")
   */
  function fmtDuration(days) {
    if (!days || days < 0) return '—';

    days = Math.round(days);

    if (days < 7) {
      return days + 'd';
    }

    if (days < 30) {
      const weeks = Math.floor(days / 7);
      const remainder = days % 7;
      if (remainder === 0) {
        return weeks + 'w';
      }
      return weeks + 'w ' + remainder + 'd';
    }

    const months = Math.floor(days / 30);
    const remainder = days % 30;
    if (remainder === 0) {
      return months + 'm';
    }
    return months + 'm ' + Math.floor(remainder / 7) + 'w';
  }

  /**
   * Format currency amount
   * @param {number} amount - Amount to format
   * @param {string} currency - Currency code (default: 'USD')
   * @returns {string} Formatted currency string
   */
  function fmtCurrency(amount, currency) {
    if (typeof amount !== 'number') return '$0';

    const currencySymbol = {
      'USD': '$',
      'EUR': '€',
      'GBP': '£',
      'JPY': '¥'
    }[currency] || '$';

    return currencySymbol + amount.toLocaleString('en-US', {
      minimumFractionDigits: 0,
      maximumFractionDigits: 2
    });
  }

  /**
   * Format percentage
   * @param {number} pct - Percentage value (0-100)
   * @returns {string} Formatted percentage string
   */
  function fmtPercent(pct) {
    if (typeof pct !== 'number') return '0%';
    return Math.round(pct) + '%';
  }

  /**
   * Get status emoji for a task
   * @param {Object} task - Task object
   * @returns {string} Status emoji
   */
  function statusIcon(task) {
    if (!task) return '⬜';

    const pct = task.percentComplete || 0;
    const now = new Date();
    const finish = task.finish ? new Date(task.finish) : null;
    const isLate = finish && finish < now && pct < 100;
    const isCritical = task.critical || false;

    if (pct === 100) {
      return '✅';
    }

    if (isLate && isCritical) {
      return '🔴';
    }

    if (isLate) {
      return '⏰';
    }

    if (isCritical) {
      return '⚠️';
    }

    if (pct > 0) {
      return '🔵';
    }

    return '⬜';
  }

  /**
   * Get health color based on score
   * @param {number} score - Score from 0-100
   * @returns {string} CSS color string
   */
  function healthColor(score) {
    if (typeof score !== 'number') return '#999';

    if (score >= 80) return '#22c55e'; // green
    if (score >= 60) return '#84cc16'; // lime
    if (score >= 40) return '#eab308'; // yellow
    if (score >= 20) return '#f97316'; // orange
    return '#ef4444'; // red
  }

  /**
   * Generate HTML for a progress bar
   * @param {number} pct - Percentage (0-100)
   * @param {string} color - CSS color (optional)
   * @param {string} width - CSS width (default: '100%')
   * @returns {string} HTML string
   */
  function progressBarHTML(pct, color, width) {
    const clamped = Math.max(0, Math.min(100, pct || 0));
    const barColor = color || '#3b82f6';
    const containerWidth = width || '100%';

    return `<div style="width: ${containerWidth}; height: 20px; background-color: #e5e7eb; border-radius: 4px; overflow: hidden;">
      <div style="width: ${clamped}%; height: 100%; background-color: ${barColor}; transition: width 0.3s ease;"></div>
    </div>`;
  }

  /**
   * Pluralize a noun
   * @param {number} n - Count
   * @param {string} singular - Singular form
   * @param {string} plural - Plural form
   * @returns {string} Pluralized string
   */
  function pluralize(n, singular, plural) {
    if (!singular || !plural) return '';
    return n === 1 ? `${n} ${singular}` : `${n} ${plural}`;
  }

  /**
   * Truncate string with ellipsis
   * @param {string} str - String to truncate
   * @param {number} maxLen - Maximum length
   * @returns {string} Truncated string
   */
  function truncate(str, maxLen) {
    if (!str || typeof str !== 'string') return '';
    if (str.length <= maxLen) return str;
    return str.substring(0, maxLen - 1) + '…';
  }

  /**
   * Escape HTML special characters
   * @param {string} str - String to escape
   * @returns {string} Escaped string
   */
  function escapeHTML(str) {
    if (!str || typeof str !== 'string') return '';

    const map = {
      '&': '&amp;',
      '<': '&lt;',
      '>': '&gt;',
      '"': '&quot;',
      "'": '&#39;'
    };

    return str.replace(/[&<>"']/g, char => map[char]);
  }

  /**
   * Create a DOM element with shorthand
   * @param {string} tag - HTML tag name
   * @param {string} cls - CSS class names (space-separated)
   * @param {string} text - Element text content
   * @returns {HTMLElement} Created element
   */
  function createEl(tag, cls, text) {
    const el = document.createElement(tag || 'div');

    if (cls) {
      el.className = cls;
    }

    if (text) {
      el.textContent = text;
    }

    return el;
  }

  /**
   * Create a styled badge element
   * @param {string} text - Badge text
   * @param {string} color - Background color (e.g., '#ff6b6b', 'red')
   * @returns {HTMLElement} Badge span element
   */
  function createBadge(text, color) {
    const badge = document.createElement('span');
    badge.textContent = text || '';
    badge.style.cssText = `
      display: inline-block;
      padding: 2px 6px;
      border-radius: 3px;
      background-color: ${color || '#999'};
      color: white;
      font-size: 12px;
      font-weight: 500;
      white-space: nowrap;
    `;
    return badge;
  }

  /**
   * Show a confirmation dialog (native confirm for now)
   * @param {string} message - Confirmation message
   * @returns {Promise<boolean>} True if user confirms
   */
  function showConfirmDialog(message) {
    return Promise.resolve(window.confirm(message || 'Are you sure?'));
  }

  /**
   * Debounce function execution
   * @param {Function} fn - Function to debounce
   * @param {number} ms - Delay in milliseconds
   * @returns {Function} Debounced function
   */
  function debounce(fn, ms) {
    if (typeof fn !== 'function') return () => {};

    let timeoutId;

    return function debounced(...args) {
      clearTimeout(timeoutId);
      timeoutId = setTimeout(() => fn.apply(this, args), ms || 300);
    };
  }

  /**
   * Throttle function execution
   * @param {Function} fn - Function to throttle
   * @param {number} ms - Interval in milliseconds
   * @returns {Function} Throttled function
   */
  function throttle(fn, ms) {
    if (typeof fn !== 'function') return () => {};

    let lastRun = 0;

    return function throttled(...args) {
      const now = Date.now();
      if (now - lastRun >= (ms || 300)) {
        lastRun = now;
        fn.apply(this, args);
      }
    };
  }

  // ==================== Private Helpers ====================

  /**
   * Format date relative to now
   * @private
   */
  function formatRelativeDate(date) {
    const now = new Date();
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const tomorrow = new Date(today);
    tomorrow.setDate(tomorrow.getDate() + 1);
    const yesterday = new Date(today);
    yesterday.setDate(yesterday.getDate() - 1);

    const dateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());

    if (dateOnly.getTime() === today.getTime()) {
      return 'today';
    }
    if (dateOnly.getTime() === tomorrow.getTime()) {
      return 'tomorrow';
    }
    if (dateOnly.getTime() === yesterday.getTime()) {
      return 'yesterday';
    }

    const diff = Math.floor((dateOnly.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));

    if (diff > 0) {
      return `in ${diff}d`;
    } else {
      return `${Math.abs(diff)}d ago`;
    }
  }

  // Public API
  export const UIHelpers = Object.freeze({
    fmtDate,
    fmtDuration,
    fmtCurrency,
    fmtPercent,
    statusIcon,
    healthColor,
    progressBarHTML,
    pluralize,
    truncate,
    escapeHTML,
    createEl,
    createBadge,
    showConfirmDialog,
    debounce,
    throttle
  });
