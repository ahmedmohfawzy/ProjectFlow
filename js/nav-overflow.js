/**
 * NavOverflow — Smart Tab Overflow Navigation
 * ────────────────────────────────────────────
 * Automatically collapses tabs that don't fit into a "More ▾" dropdown.
 * Works with any number of tabs and updates on window resize.
 *
 * Usage: NavOverflow.init() once on DOMContentLoaded.
 *        NavOverflow.setActive(view) whenever the active view changes.
 */

const NavOverflow = (() => {
    let _tabs    = [];
    let _menuEl  = null;
    let _wrapEl  = null;
    let _moreBtn = null;
    let _ro      = null;

    // ── Icon SVG for each view (used inside the dropdown) ──────────────────
    const VIEW_ICONS = {
        dashboard : `<svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14"><path d="M2 10a8 8 0 018-8v8h8a8 8 0 11-16 0z"/><path d="M12 2.252A8.014 8.014 0 0117.748 8H12V2.252z"/></svg>`,
        table     : `<svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14"><path d="M0 3a2 2 0 012-2h16a2 2 0 012 2v14a2 2 0 01-2 2H2a2 2 0 01-2-2V3zm6 2H2v4h4V5zm0 6H2v4h4v-4zm2 4h4v-4H8v4zm6 0h4v-4h-4v4zm4-6v-4h-4v4h4zm-6 0h4V5h-4v4zM8 5v4h4V5H8z"/></svg>`,
        gantt     : `<svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14"><path d="M2 4h6v3H2zM4 9h8v3H4zM6 14h10v3H6z"/></svg>`,
        split     : `<svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14"><path d="M1 3h8v14H1zM11 3h8v14h-8z" fill-opacity="0.6"/><path d="M1 3h8v14H1z"/></svg>`,
        resources : `<svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14"><path d="M9 6a3 3 0 11-6 0 3 3 0 016 0zM17 6a3 3 0 11-6 0 3 3 0 016 0zM12.93 17c.046-.327.07-.66.07-1a6.97 6.97 0 00-1.5-4.33A5 5 0 0119 16v1h-6.07zM6 11a5 5 0 015 5v1H1v-1a5 5 0 015-5z"/></svg>`,
        calendar  : `<svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14"><path fill-rule="evenodd" d="M6 2a1 1 0 00-1 1v1H4a2 2 0 00-2 2v10a2 2 0 002 2h12a2 2 0 002-2V6a2 2 0 00-2-2h-1V3a1 1 0 10-2 0v1H7V3a1 1 0 00-1-1zm0 5a1 1 0 000 2h8a1 1 0 100-2H6z" clip-rule="evenodd"/></svg>`,
        network   : `<svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14"><path d="M3 4a1 1 0 011-1h3a1 1 0 011 1v3a1 1 0 01-1 1H4a1 1 0 01-1-1V4zm9 0a1 1 0 011-1h3a1 1 0 011 1v3a1 1 0 01-1 1h-3a1 1 0 01-1-1V4zM6 12a1 1 0 011-1h3a1 1 0 011 1v3a1 1 0 01-1 1H7a1 1 0 01-1-1v-3z"/><path d="M7 7v2h2M14 7v4H10" stroke="currentColor" stroke-width="1.5" fill="none"/></svg>`,
        board     : `<svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14"><path d="M2 3h4v14H2V3zm6 0h4v14H8V3zm6 0h4v14h-4V3z" fill-opacity="0.7"/></svg>`,
        portfolio : `<svg viewBox="0 0 20 20" fill="currentColor" width="14" height="14"><path d="M2 4a1 1 0 011-1h5a1 1 0 011 1v5a1 1 0 01-1 1H3a1 1 0 01-1-1V4zm9 0a1 1 0 011-1h5a1 1 0 011 1v5a1 1 0 01-1 1h-5a1 1 0 01-1-1V4zM2 13a1 1 0 011-1h5a1 1 0 011 1v4a1 1 0 01-1 1H3a1 1 0 01-1-1v-4zm9 0a1 1 0 011-1h5a1 1 0 011 1v4a1 1 0 01-1 1h-5a1 1 0 01-1-1v-4z"/></svg>`,
    };

    // Check SVG icon
    const CHECK_ICON = `<svg class="nav-more-check" viewBox="0 0 16 16" fill="currentColor" width="12" height="12"><path d="M13.78 4.22a.75.75 0 010 1.06l-7.25 7.25a.75.75 0 01-1.06 0L2.22 9.28a.75.75 0 011.06-1.06L6 10.94l6.72-6.72a.75.75 0 011.06 0z"/></svg>`;

    // ── init ───────────────────────────────────────────────────────────────
    function init() {
        _wrapEl  = document.getElementById('navMoreWrap');
        _moreBtn = document.getElementById('navMoreBtn');
        _menuEl  = document.getElementById('navMoreMenu');

        if (!_wrapEl || !_moreBtn || !_menuEl) return;

        // Collect all tabs EXCEPT the More button itself
        _tabs = Array.from(
            document.querySelectorAll('#viewTabs .tab-btn:not(.nav-more-btn)')
        );

        // Toggle dropdown open/close
        _moreBtn.addEventListener('click', (e) => {
            e.stopPropagation();
            const isOpen = _menuEl.classList.toggle('open');
            _moreBtn.setAttribute('aria-expanded', isOpen);
        });

        // Close on outside click
        document.addEventListener('click', () => {
            _menuEl.classList.remove('open');
            _moreBtn.setAttribute('aria-expanded', 'false');
        });

        // Close on Escape
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                _menuEl.classList.remove('open');
                _moreBtn.setAttribute('aria-expanded', 'false');
            }
        });

        // Observe the header-center for width changes
        const headerCenter = document.querySelector('.header-center');
        if (!headerCenter) return;

        _ro = new ResizeObserver(() => _recalculate());
        _ro.observe(headerCenter);

        // Run immediately
        _recalculate();
    }

    // ── setActive ─────────────────────────────────────────────────────────
    function setActive(view) {
        // Update overflow menu items if visible
        document.querySelectorAll('.nav-more-item').forEach(item => {
            item.classList.toggle('active', item.dataset.view === view);
        });
        // Reflect active-in-overflow on the More button itself
        const activeInOverflow = _tabs.some(
            t => t.dataset.view === view && t.getAttribute('data-hidden') === 'overflow'
        );
        _moreBtn && _moreBtn.classList.toggle('has-active', activeInOverflow);

        // Re-run layout so the active tab is always promoted to bar
        _recalculate();
    }

    // ── _recalculate ──────────────────────────────────────────────────────
    function _recalculate() {
        if (!_wrapEl) return;

        // 1. Temporarily reveal all tabs to measure natural widths
        _tabs.forEach(t => {
            t.style.display = 'flex';
            t.removeAttribute('data-hidden');
        });
        _wrapEl.style.display = 'none';

        const container  = document.querySelector('.view-tabs');
        if (!container) return;

        // 2. Measure available width (container width minus a buffer for the More button ~80px)
        const MORE_BTN_W = 80;
        const GAP        = 1;
        const available  = container.offsetWidth - MORE_BTN_W - 8;

        // 3. Walk tabs in order; hide when cumulative width exceeds available
        let used     = 0;
        let overflow = [];

        for (const tab of _tabs) {
            const w = tab.offsetWidth + GAP;
            if (used + w > available) {
                overflow.push(tab);
            } else {
                used += w;
            }
        }

        // 4. If the active tab is in overflow → swap it with the last visible tab
        const activeTab = _tabs.find(t => t.classList.contains('active'));
        if (activeTab && overflow.includes(activeTab)) {
            const lastVisible = _tabs.filter(t => !overflow.includes(t)).at(-1);
            if (lastVisible) {
                // Swap: hide lastVisible, show activeTab
                overflow.splice(overflow.indexOf(activeTab), 1);
                overflow.push(lastVisible);
            }
        }

        // 5. Apply visibility
        overflow.forEach(t => {
            t.style.display = 'none';
            t.setAttribute('data-hidden', 'overflow');
        });

        // 6. Show or hide the More button
        if (overflow.length > 0) {
            _wrapEl.style.display = 'flex';
            _buildMenu(overflow);
            const hasActive = overflow.some(t => t.classList.contains('active'));
            _moreBtn.classList.toggle('has-active', hasActive);
        } else {
            _wrapEl.style.display = 'none';
            _moreBtn.classList.remove('has-active');
        }
    }

    // ── _buildMenu ────────────────────────────────────────────────────────
    function _buildMenu(overflowTabs) {
        _menuEl.innerHTML = '';

        overflowTabs.forEach((tab, idx) => {
            const view     = tab.dataset.view;
            const label    = tab.textContent.trim();
            const isActive = tab.classList.contains('active');
            const icon     = VIEW_ICONS[view] || '';

            const btn = document.createElement('button');
            btn.className      = 'nav-more-item' + (isActive ? ' active' : '');
            btn.dataset.view   = view;
            btn.setAttribute('role', 'menuitem');
            btn.innerHTML = `
                <span class="nav-more-icon">${icon}</span>
                <span>${label}</span>
                ${CHECK_ICON}
            `;

            btn.addEventListener('click', (e) => {
                e.stopPropagation();
                _menuEl.classList.remove('open');
                _moreBtn.setAttribute('aria-expanded', 'false');
                // Fire the original tab click so existing app.js logic runs
                tab.click();
            });

            _menuEl.appendChild(btn);

            // Divider between groups (optional cosmetic touch)
            if (idx === 0 && overflowTabs.length > 1) {
                // no divider at top
            }
        });
    }

    return { init, setActive, recalculate: _recalculate };
})();
