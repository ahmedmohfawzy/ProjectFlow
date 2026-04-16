/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 * Proprietary Software — Unauthorized use prohibited.
 * https://www.linkedin.com/in/ahmed-m-fawzy
 */
/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — Main Application Controller v2
 * ALL features: MPP import, full editing, auto-save,
 * search/filter, detail panel, resources, CPM, baseline
 * ═══════════════════════════════════════════════════════
 */
import { EventBus } from './event-bus.js';
import { StateManager } from './state-manager.js';
import { ProjectIO } from './project-io.js';
import { CPMEngine } from './critical-path.js';
import { EVMEngine } from './evm.js';
import { WorkCalendar } from './calendar.js';
import { ProjectAnalytics } from './project-analytics.js';
import { ResourceManager } from './resource-manager.js';
import { UIHelpers } from './ui-helpers.js';
import { MSProjectXML } from './xml-parser.js';
import { PlannerParser } from './planner-parser.js';
import { D365Client } from './d365.js';
import { MSGraphClient } from './ms-graph.js';
import { ScenariosManager } from './scenarios.js';
import { PluginSystem, renderPluginManager } from './plugins.js';
import { Reports } from './reports.js';
import { Dashboard } from './dashboard.js';
import { GanttChart } from './gantt.js';
import { NetworkDiagram } from './network.js';
import { BoardView, ResourceHeatmap } from './board.js';
import { TaskEditor } from './task-editor.js';
import { TeamsBridge } from './teams-bridge.js';



    // ══════ CONSTANTS ══════
    const MS_PER_DAY = 86400000;
    const SERVER_DETECT_TIMEOUT_MS = 1500;
    const MAX_ATTACHMENT_SIZE = 500 * 1024; // 500KB
    const AT_RISK_THRESHOLD = -15; // % behind to mark as at-risk
    const MAX_CANVAS_DIM = 16000; // px
    const CANVAS_FALLBACK_WIDTH = 800; // px

    // ══════ STATE ══════
    let project = null;
    let activeProjectId = null;
    let selectedTaskIds = new Set();
    let activeView = 'split';
    let _isDirty = false; // Track unsaved changes for beforeunload
    let autoSaveTimer = null;
    let sortColumn = null, sortDir = 'asc';
    let activeFilter = 'all', searchQuery = '';
    let detailTask = null;
    let calViewYear, calViewMonth;
    let sidebarOpen = false;
    let deleteTargetId = null;

    let settings = {
        dateFormat: 'YYYY-MM-DD', hoursPerDay: 8, currency: '$',
        showWBS: true, showCost: true, showFloat: true, showColor: false,
        showBaseline: true, showCritical: true, showLinks: true
    };

    let undoStack = [], redoStack = [];
    const MAX_UNDO = 50;

    // ══════ SPRINT B: RENDER THROTTLING (B.2) ══════
    let _renderPending = false;      // RAF de-duplication flag
    let _ganttResizeObserver = null; // ResizeObserver instance

    // ══════ PROJECT STORE (IndexedDB via Dexie.js) ══════
    const PROJECT_COLORS = ['#6366f1','#8b5cf6','#ec4899','#f59e0b','#22c55e','#3b82f6','#ef4444','#14b8a6','#f97316','#06b6d4'];

    // Dexie database
    const db = new Dexie('ProjectFlowDB');
    db.version(1).stores({
        projects: 'id',       // full project JSON blobs
        meta: 'id, name, lastModified, progress, pinned'  // lightweight index
    });

    // In-memory cache for sync access (loaded once at init)
    let _indexCache = [];
    let _dbReady = false;

    // ══════ SERVER BACKEND ══════
    let _serverMode = false;
    const DEFAULT_SERVER = 'http://localhost:3456';
    let _serverURL = DEFAULT_SERVER;
    try {
        const saved = localStorage.getItem('pf_server_url');
        if (saved && /^https?:\/\/(localhost|127\.0\.0\.1)(:\d+)?$/.test(saved)) {
            _serverURL = saved;
        }
    } catch(e) {}
    let _activeDB = localStorage.getItem('pf_active_db') || 'default.db';

    async function serverAPI(method, path, body) {
        const opts = { method, headers: { 'Content-Type': 'application/json' } };
        if (body) opts.body = JSON.stringify(body);
        const resp = await fetch(_serverURL + path, opts);
        if (!resp.ok) { const e = await resp.json().catch(() => ({})); throw new Error(e.error || resp.statusText); }
        return resp.json();
    }

    async function detectServer() {
        try {
            const r = await fetch(_serverURL + '/api/ping', { signal: AbortSignal.timeout(1500) });
            if (r.ok) { _serverMode = true; return true; }
        } catch(e) {}
        _serverMode = false;
        return false;
    }

    const ProjectStore = {
        // ── Init: load index into cache ──
        async init() {
            // Try server first
            if (await detectServer()) {
                try {
                    const cfg = await serverAPI('GET', '/api/config');
                    _activeDB = cfg.activeDatabase || 'default.db';
                    localStorage.setItem('pf_active_db', _activeDB);
                    const projects = await serverAPI('GET', `/api/db/${_activeDB}/projects`);
                    _indexCache = projects.map(p => ({
                        id: p.id, name: p.name, color: p.color || '#6366f1',
                        pinned: !!p.pinned, archived: !!p.archived,
                        description: p.description || '',
                        taskCount: p.task_count || 0, progress: p.progress || 0,
                        startDate: p.start_date, finishDate: p.finish_date,
                        lastModified: p.last_modified
                    }));
                    console.log(`[ProjectStore] Server mode — ${_indexCache.length} projects in ${_activeDB}`);
                    return;
                } catch(e) {
                    console.warn('[ProjectStore] Server connect failed, falling back:', e);
                    _serverMode = false;
                }
            }
            // Fallback: IndexedDB
            try {
                _indexCache = await db.meta.toArray();
                _dbReady = true;
                console.log(`[ProjectStore] IndexedDB ready — ${_indexCache.length} projects`);
            } catch (e) {
                console.warn('[ProjectStore] IndexedDB failed, falling back to localStorage', e);
                _dbReady = false;
                try { _indexCache = JSON.parse(localStorage.getItem('pf_projects_index') || '[]'); } catch(e2) { _indexCache = []; }
            }
        },

        // ── Sync reads (from cache) ──
        getIndex() { return _indexCache; },
        generateId() { return 'proj_' + Date.now() + '_' + Math.random().toString(36).substr(2,5); },
        getActive() { return localStorage.getItem('pf_active_project') || null; },
        setActive(id) { localStorage.setItem('pf_active_project', id); },

        // ── Async writes ──
        async save(id, proj) {
            if (_serverMode) {
                try {
                    const meta = this._buildMeta(id, proj);
                    await serverAPI('POST', `/api/db/${_activeDB}/projects/${id}`, { meta, data: proj });
                    return;
                } catch(e) { console.warn('[ProjectStore] Server save failed:', e); }
            }
            try {
                if (_dbReady) {
                    await db.projects.put({ id, data: proj });
                } else {
                    localStorage.setItem('pf_project_' + id, JSON.stringify(proj));
                }
            } catch (e) {
                console.warn('[ProjectStore] Save failed, trying localStorage fallback', e);
                try { localStorage.setItem('pf_project_' + id, JSON.stringify(proj)); } catch(e2) {
                    console.error('[ProjectStore] Both saves failed!', e2);
                }
            }
        },

        async load(id) {
            try {
                let raw;
                if (_serverMode) {
                    const result = await serverAPI('GET', `/api/db/${_activeDB}/projects/${id}`);
                    raw = result.data;
                } else if (_dbReady) {
                    const rec = await db.projects.get(id);
                    raw = rec ? rec.data : null;
                } else {
                    const str = localStorage.getItem('pf_project_' + id);
                    raw = str ? JSON.parse(str) : null;
                }
                if (!raw) return null;
                const p = typeof raw === 'string' ? JSON.parse(raw) : raw;
                // Restore dates & defaults
                (p.tasks || []).forEach(t => {
                    t.start = new Date(t.start); t.finish = new Date(t.finish);
                    if (t.baselineStart) t.baselineStart = new Date(t.baselineStart);
                    if (t.baselineFinish) t.baselineFinish = new Date(t.baselineFinish);
                    t.isExpanded = t.isExpanded !== false; t.isVisible = true;
                    if (!t.predecessors) t.predecessors = [];
                    if (!t.resourceNames) t.resourceNames = [];
                    if (!t.tags) t.tags = [];
                    if (!t.comments) t.comments = [];
                    if (!t.attachments) t.attachments = [];
                });
                if (p.startDate) p.startDate = new Date(p.startDate);
                if (p.finishDate) p.finishDate = new Date(p.finishDate);
                if (!p.resources) p.resources = [];
                if (!p.assignments) p.assignments = [];
                return p;
            } catch (e) { console.warn('[ProjectStore] Load failed', e); return null; }
        },

        async delete(id) {
            console.log('[ProjectStore] Deleting project:', id);
            try {
                if (_serverMode) {
                    await serverAPI('DELETE', `/api/db/${_activeDB}/projects/${id}`);
                    console.log('[ProjectStore] Deleted from server');
                }
                // Also clean local copies
                if (_dbReady) {
                    try { await db.projects.delete(id); await db.meta.delete(id); } catch(e) {}
                }
                localStorage.removeItem('pf_project_' + id);
                if (localStorage.getItem('pf_active_project') === id) {
                    localStorage.removeItem('pf_active_project');
                }
                localStorage.removeItem('pf_autosave');
                _indexCache = _indexCache.filter(p => p.id !== id);
                this._syncIndexToLS();
                console.log('[ProjectStore] Delete complete. Remaining:', _indexCache.length);
            } catch (e) { console.warn('[ProjectStore] Delete failed', e); }
        },

        async addToIndex(id, proj) {
            const existing = _indexCache.findIndex(p => p.id === id);
            const entry = {
                id, name: proj.name || 'Untitled',
                color: PROJECT_COLORS[_indexCache.length % PROJECT_COLORS.length],
                pinned: false,
                lastModified: new Date().toISOString(),
                taskCount: (proj.tasks || []).length,
                progress: this._calcProgress(proj),
                startDate: proj.startDate ? new Date(proj.startDate).toISOString() : null,
                finishDate: proj.finishDate ? new Date(proj.finishDate).toISOString() : null
            };
            if (existing >= 0) {
                entry.color = _indexCache[existing].color;
                entry.pinned = _indexCache[existing].pinned;
                entry.archived = _indexCache[existing].archived || false;
                entry.description = _indexCache[existing].description || '';
                _indexCache[existing] = entry;
            } else { _indexCache.push(entry); }

            // Persist
            try {
                if (_dbReady) await db.meta.put(entry);
            } catch(e) { console.warn('[ProjectStore] Index save failed', e); }
            this._syncIndexToLS();
        },

        _buildMeta(id, proj) {
            const existing = _indexCache.find(p => p.id === id);
            return {
                name: proj.name || 'Untitled',
                color: (existing && existing.color) || PROJECT_COLORS[_indexCache.length % PROJECT_COLORS.length],
                pinned: existing ? !!existing.pinned : false,
                archived: existing ? !!existing.archived : false,
                description: (existing && existing.description) || '',
                task_count: (proj.tasks || []).length,
                progress: this._calcProgress(proj),
                start_date: proj.startDate ? new Date(proj.startDate).toISOString() : null,
                finish_date: proj.finishDate ? new Date(proj.finishDate).toISOString() : null
            };
        },

        _calcProgress(proj) {
            if (!proj || !proj.tasks || !proj.tasks.length) return 0;
            const leafTasks = proj.tasks.filter(t => !t.summary);
            if (!leafTasks.length) return 0;
            return Math.round(leafTasks.reduce((s,t) => s + (t.percentComplete||0), 0) / leafTasks.length);
        },

        // Keep localStorage in sync as backup
        _syncIndexToLS() {
            try { localStorage.setItem('pf_projects_index', JSON.stringify(_indexCache)); } catch(e) {}
        },

        // ── Migration: localStorage → IndexedDB ──
        async migrate() {
            // Step 1: Migrate legacy pf_autosave
            const legacy = localStorage.getItem('pf_autosave');
            if (legacy && _indexCache.length === 0) {
                try {
                    const id = this.generateId();
                    const proj = JSON.parse(legacy);
                    await this.save(id, proj);
                    await this.addToIndex(id, proj);
                    this.setActive(id);
                    console.log('[ProjectStore] Migrated legacy pf_autosave →', id);
                    // Remove legacy key so it never re-triggers
                    localStorage.removeItem('pf_autosave');
                } catch(e) { console.warn('[ProjectStore] Legacy migration failed', e); }
            }

            // Step 2: Migrate localStorage projects → IndexedDB
            if (_dbReady) {
                // Read the localStorage index (which reflects deletions via _syncIndexToLS)
                let lsIndex = [];
                try { lsIndex = JSON.parse(localStorage.getItem('pf_projects_index') || '[]'); } catch(e) {}

                for (const meta of lsIndex) {
                    const lsKey = 'pf_project_' + meta.id;
                    const lsData = localStorage.getItem(lsKey);
                    if (lsData) {
                        const exists = await db.projects.get(meta.id);
                        if (!exists) {
                            try {
                                await db.projects.put({ id: meta.id, data: JSON.parse(lsData) });
                                await db.meta.put(meta);
                                console.log('[ProjectStore] Migrated to IDB:', meta.name);
                            } catch(e) { console.warn('[ProjectStore] Project migration failed:', meta.name, e); }
                        }
                    }
                }
                // Reload cache from IDB
                _indexCache = await db.meta.toArray();

                // CRITICAL: Cross-check — remove IDB entries that aren't in localStorage index
                // This catches cases where IDB delete silently failed (e.g., file:// protocol)
                const lsIdSet = new Set(lsIndex.map(m => m.id));
                const orphans = _indexCache.filter(m => !lsIdSet.has(m.id));
                for (const orphan of orphans) {
                    console.log('[ProjectStore] Cleaning orphaned IDB entry:', orphan.name);
                    try {
                        await db.projects.delete(orphan.id);
                        await db.meta.delete(orphan.id);
                    } catch(e) {}
                }
                if (orphans.length > 0) {
                    _indexCache = _indexCache.filter(m => lsIdSet.has(m.id));
                }
            }
        },

        // ── Storage Info ──
        async getStorageInfo() {
            if (navigator.storage && navigator.storage.estimate) {
                const est = await navigator.storage.estimate();
                return { used: est.usage || 0, quota: est.quota || 0, pct: Math.round((est.usage / est.quota) * 100) };
            }
            return { used: 0, quota: 0, pct: 0 };
        },

        // ── Portfolio Export / Import ──
        async exportAll() {
            const allProjects = [];
            for (const meta of _indexCache) {
                const proj = await this.load(meta.id);
                if (proj) allProjects.push({ meta, data: proj });
            }
            return { version: 2, exportDate: new Date().toISOString(), projects: allProjects };
        },

        async importAll(bundle) {
            if (!bundle || !bundle.projects) throw new Error('Invalid backup file');
            let imported = 0;
            for (const item of bundle.projects) {
                const newId = this.generateId();
                const proj = item.data;
                // Restore dates
                (proj.tasks || []).forEach(t => {
                    t.start = new Date(t.start); t.finish = new Date(t.finish);
                    if (t.baselineStart) t.baselineStart = new Date(t.baselineStart);
                    if (t.baselineFinish) t.baselineFinish = new Date(t.baselineFinish);
                });
                if (proj.startDate) proj.startDate = new Date(proj.startDate);
                if (proj.finishDate) proj.finishDate = new Date(proj.finishDate);
                await this.save(newId, proj);
                const meta = item.meta || {};
                meta.id = newId;
                meta.lastModified = new Date().toISOString();
                if (!meta.name) meta.name = proj.name || 'Imported Project';
                if (!meta.color) meta.color = PROJECT_COLORS[_indexCache.length % PROJECT_COLORS.length];
                meta.taskCount = (proj.tasks || []).length;
                meta.progress = this._calcProgress(proj);
                _indexCache.push(meta);
                if (_dbReady) await db.meta.put(meta);
                imported++;
            }
            this._syncIndexToLS();
            return imported;
        }
    };

    // ══════ DOM REFS ══════
    const $ = (id) => document.getElementById(id);
    const els = {
        welcomeScreen: $('welcomeScreen'), workspace: $('workspace'),
        splitContainer: $('splitContainer'), taskTableBody: $('taskTableBody'),
        ganttCanvas: $('ganttCanvas'), ganttHeader: $('ganttHeader'),
        ganttBody: $('ganttBody'), tableWrapper: $('tableWrapper'),
        projectNameDisplay: $('projectNameDisplay'), taskCount: $('taskCount'),
        projectDuration: $('projectDuration'), criticalInfo: $('criticalInfo'),
        fileInput: $('fileInput'),
        toastContainer: $('toastContainer'), zoomLevel: $('zoomLevel'),
        statusIndicator: $('statusIndicator'), serverStatus: $('serverStatus'),
        autoSaveStatus: $('autoSaveStatus'),
        searchInput: $('searchInput'), searchClear: $('searchClear'),
        filterSelect: $('filterSelect'),
        resourceTableBody: $('resourceTableBody'), resourceView: $('resourceView'),
        // recentProjects / recentList — legacy, replaced by ProjectHub (kept as null-safe)
        detailPanel: $('taskDetailPanel'),
        calendarView: $('calendarView'), calendarGrid: $('calendarGrid'),
        calMonthLabel: $('calMonthLabel'), calendarHolidaysList: $('calendarHolidaysList'),
        dashboardView: $('dashboardView'), notifPanel: $('notifPanel'),
        notifBadge: $('notifBadge'), notifBody: $('notifBody'),
        filePlannerInput: $('filePlannerInput'),
        networkView: $('networkView'), networkCanvas: $('networkCanvas'),
        // Phase 6a
        projectSidebar: $('projectSidebar'), sidebarProjectList: $('sidebarProjectList'),
        sidebarSearch: $('sidebarSearch'),
        hubProjectGrid: $('hubProjectGrid'), hubEmptyState: $('hubEmptyState'),
        hubKpiBar: $('hubKpiBar'),
        portfolioView: $('portfolioView'),
    };

    // ══════ INIT ══════
    async function init() {
        await ProjectStore.init();
        if (!_serverMode) await ProjectStore.migrate(); // Skip migration for server mode
        bindEvents();
        bindMultiProjectEvents();
        loadSettings();
        initCalendar();
        $('inputStartDate').value = new Date().toISOString().split('T')[0];
        setStatus('Ready');
        // Update storage indicator
        const si = $('storageIndicator');
        if (si) si.textContent = _serverMode ? `🗄️ SQLite: ${_activeDB}` : '📦 Browser';
        renderProjectHub();
        renderSidebar();
        setView(activeView);
        // ── Smart overflow nav (must run after DOM is laid out) ──
        if (window.NavOverflow) NavOverflow.init();
        // ── Header dropdowns (Settings / Scenarios) ──
        initHeaderDropdowns();
    }

    // ══════ EVENT BINDING ══════
    function bindEvents() {
        $('btnImport').addEventListener('click', () => els.fileInput.click());
        $('btnWelcomeImport').addEventListener('click', () => { els.fileInput.click(); });
        els.fileInput.addEventListener('change', handleFileImport);

        // MPP removed — no server dependency

        $('btnImportPlanner').addEventListener('click', () => els.filePlannerInput.click());
        els.filePlannerInput.addEventListener('change', handlePlannerFileSelected);

        $('btnExport').addEventListener('click', handleExportXML);
        $('btnExportExcel').addEventListener('click', handleExportCSV);

        $('btnNewProject').addEventListener('click', showNewProjectModal);
        $('btnWelcomeNew').addEventListener('click', showNewProjectModal);
        $('btnCreateProject').addEventListener('click', handleCreateProject);
        $('btnCancelNewProject').addEventListener('click', () => toggleModal('modalNewProject', false));
        $('btnCloseNewProject').addEventListener('click', () => toggleModal('modalNewProject', false));

        $('btnSettings').addEventListener('click', () => { populateSettingsModal(); toggleModal('modalSettings', true); });
        $('btnCloseSettings').addEventListener('click', () => toggleModal('modalSettings', false));
        $('btnAbout').addEventListener('click', () => toggleModal('modalAbout', true));
        $('btnCloseAbout').addEventListener('click', () => toggleModal('modalAbout', false));
        $('btnSaveSettings').addEventListener('click', handleSaveSettings);

        // ── Logo picker ──
        $('btnChooseLogo').addEventListener('click', () => $('inputLogoFile').click());
        $('inputLogoFile').addEventListener('change', function() {
            const file = this.files[0];
            if (!file) return;
            if (file.size > 500 * 1024) { showToast('warning', 'Logo must be under 500KB'); return; }
            const reader = new FileReader();
            reader.onload = (e) => {
                const dataUrl = e.target.result;
                localStorage.setItem('pf_report_logo', dataUrl);
                window.PROART_LOGO = dataUrl;
                $('logoPreviewImg').src = dataUrl;
                $('logoPreviewImg').style.display = 'block';
                $('logoPlaceholder').style.display = 'none';
                showToast('success', 'Logo saved for reports');
            };
            reader.readAsDataURL(file);
        });
        $('btnClearLogo').addEventListener('click', () => {
            localStorage.removeItem('pf_report_logo');
            window.PROART_LOGO = undefined;
            $('logoPreviewImg').src = '';
            $('logoPreviewImg').style.display = 'none';
            $('logoPlaceholder').style.display = '';
            showToast('info', 'Logo cleared');
        });

        // Server settings
        $('btnTestServer').addEventListener('click', async () => {
            const url = $('settingServerURL').value.trim() || 'http://localhost:3456';
            _serverURL = url;
            localStorage.setItem('pf_server_url', url);
            const badge = $('serverStatusBadge');
            badge.textContent = '⏳ Testing…';
            badge.className = 'server-badge offline';
            if (await detectServer()) {
                badge.textContent = '🟢 Server Connected';
                badge.className = 'server-badge online';
                $('serverSettingsPanel').classList.remove('hidden');
                // Load databases list
                try {
                    const data = await serverAPI('GET', '/api/databases');
                    const sel = $('settingActiveDB');
                    sel.innerHTML = '';
                    data.databases.forEach(db => {
                        const opt = document.createElement('option');
                        opt.value = db.name; opt.textContent = `${db.name} (${(db.size/1024).toFixed(1)} KB)`;
                        if (db.active) opt.selected = true;
                        sel.appendChild(opt);
                    });
                    $('settingDataDir').value = data.dataDir || '';
                } catch(e) {}
            } else {
                badge.textContent = '🔴 Not Connected';
                badge.className = 'server-badge offline';
                $('serverSettingsPanel').classList.add('hidden');
                showToast('error', 'Cannot connect to server at ' + url);
            }
        });

        $('btnNewDB').addEventListener('click', async () => {
            const name = prompt('Database name (e.g., "Work Projects"):');
            if (!name) return;
            try {
                await serverAPI('POST', '/api/databases', { name });
                showToast('success', `Database "${name}" created`);
                $('btnTestServer').click(); // Refresh list
            } catch(e) { showToast('error', e.message); }
        });

        // Folder Browser
        let _browseCurrentPath = '';

        async function loadFolderBrowser(dirPath) {
            try {
                const data = await serverAPI('GET', `/api/browse?path=${encodeURIComponent(dirPath || '')}`);
                _browseCurrentPath = data.current;
                $('folderBreadcrumb').textContent = data.current;
                const list = $('folderList');
                list.innerHTML = '';
                if (data.folders.length === 0) {
                    list.innerHTML = '<div style="padding:12px;text-align:center;color:var(--text-muted);font-size:0.8rem">📭 No subfolders here</div>';
                }
                data.folders.forEach(f => {
                    const row = document.createElement('div');
                    row.style.cssText = 'display:flex;align-items:center;gap:8px;padding:6px 10px;border-radius:6px;cursor:pointer;font-size:0.8rem;transition:background 0.15s';
                    row.innerHTML = `<span style="font-size:1.1rem">📁</span><span>${escapeHTML(f.name)}</span>`;
                    row.addEventListener('mouseenter', () => row.style.background = 'var(--bg-hover)');
                    row.addEventListener('mouseleave', () => row.style.background = '');
                    row.addEventListener('dblclick', () => loadFolderBrowser(f.path));
                    row.addEventListener('click', () => {
                        list.querySelectorAll('div').forEach(d => d.style.background = '');
                        row.style.background = 'var(--accent-primary-alpha)';
                        _browseCurrentPath = f.path;
                        $('folderBreadcrumb').textContent = f.path;
                    });
                    list.appendChild(row);
                });
                // Set up the Up button
                $('btnFolderUp').onclick = () => loadFolderBrowser(data.parent);
            } catch(e) {
                showToast('error', 'Browse failed: ' + e.message);
            }
        }

        $('btnChangeDataDir').addEventListener('click', async () => {
            try {
                showToast('info', 'Opening folder picker…');
                const result = await serverAPI('GET', '/api/pick-folder');
                if (result.success && result.path) {
                    await serverAPI('PUT', '/api/config', { dataDir: result.path });
                    $('settingDataDir').value = result.path;
                    showToast('success', 'Data directory: ' + result.path);
                    await ProjectStore.init();
                    const si = $('storageIndicator');
                    if (si) si.textContent = `🗄️ SQLite: ${_activeDB}`;
                    renderProjectHub();
                    renderSidebar();
                    $('btnTestServer').click(); // refresh DB list
                } else {
                    // User cancelled — do nothing
                }
            } catch(e) { showToast('error', e.message); }
        });

        $('settingActiveDB').addEventListener('change', async () => {
            const name = $('settingActiveDB').value;
            try {
                await serverAPI('PUT', '/api/config', { activeDatabase: name });
                _activeDB = name;
                localStorage.setItem('pf_active_db', name);
                await ProjectStore.init();
                const si = $('storageIndicator');
                if (si) si.textContent = `🗄️ SQLite: ${name}`;
                renderProjectHub();
                renderSidebar();
                showToast('success', `Switched to database: ${name}`);
            } catch(e) { showToast('error', e.message); }
        });

        $('btnShutdownServer').addEventListener('click', async () => {
            if (!confirm('Stop the server? You will need to restart it from "Start Server.command" to use server features again.')) return;
            try {
                await serverAPI('POST', '/api/shutdown');
                _serverMode = false;
                const si = $('storageIndicator');
                if (si) si.textContent = '📦 Browser';
                const badge = $('serverStatusBadge');
                badge.textContent = '🔴 Server Stopped';
                badge.className = 'server-badge offline';
                $('serverSettingsPanel').classList.add('hidden');
                showToast('info', 'Server stopped');
            } catch(e) {}
        });

        $('btnAddTask').addEventListener('click', handleAddTask);
        $('btnAddMilestone').addEventListener('click', handleAddMilestone);
        $('btnDeleteTask').addEventListener('click', handleDeleteTasks);
        $('btnIndent').addEventListener('click', () => handleIndent(1));
        $('btnOutdent').addEventListener('click', () => handleIndent(-1));

        $('btnSetBaseline').addEventListener('click', handleSetBaseline);
        $('btnAutoLevel').addEventListener('click', handleAutoLevel);

        $('viewTabs').addEventListener('click', (e) => {
            const btn = e.target.closest('.tab-btn');
            if (!btn) return;
            const view = btn.dataset.view;
            // Portfolio, Dashboard, Calendar, Resources, Network are workspace-level views
            // that don't require an open project — make sure workspace is visible
            const noProjectViews = ['portfolio', 'dashboard', 'calendar', 'resources', 'network'];
            if (noProjectViews.includes(view) && !project) {
                els.welcomeScreen.classList.add('hidden');
                els.workspace.classList.remove('hidden');
            }
            setView(view);
        });

        $('btnUndo').addEventListener('click', handleUndo);
        $('btnRedo').addEventListener('click', handleRedo);

        $('btnZoomIn').addEventListener('click', () => { els.zoomLevel.textContent = GanttChart.zoomIn(); });
        $('btnZoomOut').addEventListener('click', () => { els.zoomLevel.textContent = GanttChart.zoomOut(); });

        $('selectAll').addEventListener('change', (e) => { if (!project) return; selectedTaskIds.clear(); if (e.target.checked) project.tasks.forEach(t => selectedTaskIds.add(t.uid)); renderTable(); });

        // Search
        els.searchInput.addEventListener('input', debounce(() => { searchQuery = els.searchInput.value.trim().toLowerCase(); els.searchClear.classList.toggle('hidden', !searchQuery); renderTable(); }, 200));
        els.searchClear.addEventListener('click', () => { els.searchInput.value = ''; searchQuery = ''; els.searchClear.classList.add('hidden'); renderTable(); });

        // Filter
        els.filterSelect.addEventListener('change', () => { activeFilter = els.filterSelect.value; renderTable(); });

        // ── Toolbar dropdown menus ────────────────────────────
        /**
         * Position the menu using fixed coords from trigger's bounding rect.
         * This works even when the toolbar has overflow:hidden/auto.
         */
        function _positionDropdown(drop) {
            const trigger = drop.querySelector('.tb-drop-trigger');
            const menu    = drop.querySelector('.tb-drop-menu');
            if (!trigger || !menu) return;

            const rect = trigger.getBoundingClientRect();
            const menuW = menu.offsetWidth  || 200;
            const menuH = menu.offsetHeight || 200;

            let top  = rect.bottom + 4;
            let left = rect.left;

            // Keep menu inside viewport horizontally
            if (left + menuW > window.innerWidth - 8) left = window.innerWidth - menuW - 8;
            if (left < 8) left = 8;

            // Flip up if not enough space below
            if (top + menuH > window.innerHeight - 8) top = rect.top - menuH - 4;

            menu.style.top  = Math.round(top)  + 'px';
            menu.style.left = Math.round(left) + 'px';
        }

        function _initDropdown(dropId) {
            const drop = $(dropId);
            if (!drop) return;
            const trigger = drop.querySelector('.tb-drop-trigger');
            const menu    = drop.querySelector('.tb-drop-menu');

            if (trigger) {
                trigger.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const isOpen = drop.classList.contains('open');

                    // Close all other dropdowns
                    document.querySelectorAll('.tb-dropdown.open').forEach(d => {
                        if (d !== drop) d.classList.remove('open');
                    });

                    if (isOpen) {
                        drop.classList.remove('open');
                    } else {
                        drop.classList.add('open');
                        // Position AFTER display:flex kicks in (next tick)
                        requestAnimationFrame(() => _positionDropdown(drop));
                    }
                });
            }

            // Close dropdown when any item is clicked, BEFORE the item's own handler fires
            if (menu) {
                menu.addEventListener('click', (e) => {
                    // Let the click reach the item's own listeners first, then close
                    requestAnimationFrame(() => drop.classList.remove('open'));
                });
            }
        }

        _initDropdown('tbImportDrop');
        _initDropdown('tbExportDrop');

        // Close all on outside click
        document.addEventListener('click', (e) => {
            if (!e.target.closest('.tb-dropdown')) {
                document.querySelectorAll('.tb-dropdown.open').forEach(d => d.classList.remove('open'));
            }
        });

        // Close all on Escape
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape') {
                document.querySelectorAll('.tb-dropdown.open').forEach(d => d.classList.remove('open'));
            }
        });

        // Reposition on scroll/resize
        window.addEventListener('resize', () => {
            document.querySelectorAll('.tb-dropdown.open').forEach(drop => _positionDropdown(drop));
        });

        // Export trigger: sync disabled state with btnExport
        const _syncExportTrigger = () => {
            const t = $('btnExportTrigger'), b = $('btnExport');
            if (t && b) t.disabled = b.disabled;
        };
        const _exportObserver = new MutationObserver(_syncExportTrigger);
        const _btnExportEl = $('btnExport');
        if (_btnExportEl) _exportObserver.observe(_btnExportEl, { attributes: true, attributeFilter: ['disabled'] });

        // Sort
        document.querySelectorAll('.task-table th.sortable').forEach(th => {
            th.addEventListener('click', () => {
                const col = th.dataset.sort;
                if (sortColumn === col) sortDir = sortDir === 'asc' ? 'desc' : 'asc';
                else { sortColumn = col; sortDir = 'asc'; }
                document.querySelectorAll('.task-table th').forEach(h => h.classList.remove('sort-asc', 'sort-desc'));
                th.classList.add(sortDir === 'asc' ? 'sort-asc' : 'sort-desc');
                renderTable();
            });
        });

        // Detail panel
        $('btnCloseDetail').addEventListener('click', closeDetailPanel);
        $('btnSaveDetail').addEventListener('click', saveDetailPanel);

        // Resources
        $('btnAddResource').addEventListener('click', () => toggleModal('modalAddResource', true));
        $('btnCloseAddResource').addEventListener('click', () => toggleModal('modalAddResource', false));
        $('btnCancelAddResource').addEventListener('click', () => toggleModal('modalAddResource', false));
        $('btnConfirmAddResource').addEventListener('click', handleAddResource);

        // Calendar
        $('btnCalendarSettings').addEventListener('click', () => { populateCalendarModal(); toggleModal('modalCalendar', true); });
        $('btnCloseCalendar').addEventListener('click', () => toggleModal('modalCalendar', false));
        $('btnSaveCalendar').addEventListener('click', handleSaveCalendar);
        $('btnCalFetchHolidays').addEventListener('click', handleFetchHolidays);
        $('btnAddCustomHoliday').addEventListener('click', handleAddCustomHoliday);
        $('calPreset').addEventListener('change', (e) => {
            $('customDaysGroup').style.display = e.target.value === 'custom' ? '' : 'none';
        });
        $('calPrev').addEventListener('click', () => { calViewMonth--; if (calViewMonth < 0) { calViewMonth = 11; calViewYear--; } renderCalendarView(); });
        $('calNext').addEventListener('click', () => { calViewMonth++; if (calViewMonth > 11) { calViewMonth = 0; calViewYear++; } renderCalendarView(); });

        // MPP removed — no server-related modals or buttons

        // Notifications
        $('btnNotifications').addEventListener('click', toggleNotifPanel);
        $('btnCloseNotif').addEventListener('click', () => els.notifPanel.classList.add('hidden'));

        // Reports
        $('btnReports').addEventListener('click', () => { if (project) toggleModal('modalReport', true); });
        $('btnCloseReport').addEventListener('click', () => toggleModal('modalReport', false));
        $('modalReport').addEventListener('click', (e) => { if (e.target === $('modalReport')) toggleModal('modalReport', false); });

        $('rptPDF').addEventListener('click', handleReportPDF);
        $('rptExcel').addEventListener('click', handleReportExcel);
        $('rptGanttPNG').addEventListener('click', handleReportGanttPNG);
        $('rptSummary').addEventListener('click', handleReportSummary);
        $('rptPrint').addEventListener('click', handleReportPrint);
        $('rptDashboard').addEventListener('click', handleReportDashboard);

        document.addEventListener('keydown', handleKeyboard);
        els.tableWrapper.addEventListener('scroll', syncScroll);
        els.ganttBody.addEventListener('scroll', syncScrollReverse);
        initResizeHandle();

        // Phase 4: Theme, RTL, Shortcuts
        $('btnThemeToggle').addEventListener('click', toggleTheme);
        $('btnRTLToggle').addEventListener('click', toggleRTL);
        $('btnShortcuts').addEventListener('click', () => { populateShortcuts(); toggleModal('modalShortcuts', true); });
        if ($('shortcutSearch')) $('shortcutSearch').addEventListener('input', filterShortcuts);

        // ── Sprint C & D: New Feature Listeners ──────────────
        initAdvancedFilter();   // C.2
        initTableDragDrop();    // C.3
        initBatchOps();         // C.4
        initBoardView();        // D.1
        initCustomFields();     // D.4
        initCSVImport();        // D.5

        // ── Sprint E & F + Phase 6b ───────────────────────────
        initShareLink();        // 6b.1
        initScenarios();        // 6b.2 + E.4
        initPWA();              // E.1
        initPluginSystem();     // E.5
        initWebhooks();         // F.3
        initRoles();            // F.2

        // ── MS Integrations ───────────────────────────────────
        initPlannerSync();      // MS Planner Live Sync
        initD365Sync();         // D365 Project Accounting

        // ── Teams Bridge: Auto-connect Planner & D365 ─────────
        // Listen for Planner events BEFORE init so we don't miss them
        EventBus.on('planner:connected', ({ plans, lastPlanId }) => {
            if (!plans || plans.length === 0) return;
            // If only one plan, auto-import it silently
            if (plans.length === 1) {
                _autoImportPlan(plans[0].id, plans[0].title);
                return;
            }
            // Multiple plans: show picker modal
            _showPlanPickerModal(plans, lastPlanId);
        });

        EventBus.on('planner:needs-signin', () => {
            // Only show the sign-in toast if we're in a Teams context
            if (TeamsBridge.isInTeams()) {
                showToast('info', '🔗 Click "MS Planner Live" to sign in and load your projects.');
            }
        });

        TeamsBridge.init().then(result => {
            if (result.isInTeams) console.log('[PF] Running inside Microsoft Teams');
            if (result.planner) console.log('[PF] Planner auto-connected:', result.planner.plans?.length, 'plans');
            if (result.d365) console.log('[PF] D365 auto-connected:', result.d365.projects?.length, 'projects');
        }).catch(err => console.warn('[PF] Teams bridge init:', err.message));

        // Restore saved theme/RTL
        const savedTheme = localStorage.getItem('pf_theme');
        if (savedTheme) document.documentElement.setAttribute('data-theme', savedTheme);
        if (savedTheme === 'light') $('themeIcon').textContent = '☀️';
        const savedDir = localStorage.getItem('pf_dir');
        if (savedDir) document.documentElement.setAttribute('dir', savedDir);
    }

    // ══════ MPP SERVER — REMOVED ══════
    // MPP import functionality has been removed.
    // The app now relies on XML and MS Planner import only.
    function checkMPPServer() { /* no-op */ }

    function handlePlannerFileSelected(e) {
        const file = e.target.files[0]; if (!file) return;
        setStatus('Parsing Planner Excel…');
        
        PlannerParser.parse(file)
            .then(data => {
                project = data;
                // Pre-process project data
                project.tasks.forEach(t => {
                    if (t.start && !(t.start instanceof Date)) t.start = new Date(t.start);
                    if (t.finish && !(t.finish instanceof Date)) t.finish = new Date(t.finish);
                    t.isExpanded = true; t.isVisible = true;
                    if (!t.predecessors) t.predecessors = [];
                    if (!t.resourceNames) t.resourceNames = [];
                });
                if (project.startDate && !(project.startDate instanceof Date)) project.startDate = new Date(project.startDate);
                if (project.finishDate && !(project.finishDate instanceof Date)) project.finishDate = new Date(project.finishDate);
                
                reindexTasks();
                // Multi-project: assign new ID
                activeProjectId = ProjectStore.generateId();
                onProjectLoaded();
                showToast('success', `Imported Planner: "${project.name}" — ${project.tasks.length} items`);
            })
            .catch(err => {
                const msg = err.message || 'Unknown error';
                showToast('error', 'Planner import failed: ' + msg);
                console.error('[Planner Import Error]', err);
                if (msg.includes('library not loaded')) {
                    showToast('warning', 'Please refresh the page and try again.');
                }
            })
            .finally(() => {
                setStatus('Ready');
                els.filePlannerInput.value = '';
            });
    }

    // ══════ FILE IMPORT (XML) ══════
    function handleFileImport(e) {
        const file = e.target.files[0]; if (!file) return;
        setStatus('Importing…');
        const reader = new FileReader();
        reader.onload = (ev) => {
            try {
                project = MSProjectXML.parse(ev.target.result);
                project.tasks.forEach(t => { t.isExpanded = true; t.isVisible = true; });
                if (!project.resources) project.resources = [];
                if (!project.assignments) project.assignments = [];
                reindexTasks();
                // Multi-project: assign new ID
                activeProjectId = ProjectStore.generateId();
                onProjectLoaded();
                showToast('success', `Imported "${project.name}" — ${project.tasks.length} tasks`);
            } catch (err) { showToast('error', 'XML parse error: ' + err.message); }
            setStatus('Ready');
        };
        reader.readAsText(file);
        els.fileInput.value = '';
    }

    // ══════ EXPORT ══════
    function handleExportXML() {
        if (!project) return;
        try {
            const xml = MSProjectXML.exportXML(project);
            downloadFile(xml, sanitize(project.name) + '.xml', 'application/xml');
            showToast('success', 'Exported XML');
        } catch (err) { showToast('error', 'Export failed: ' + err.message); }
    }

    function handleExportCSV() {
        if (!project) return;
        const h = ['ID','WBS','Task Name','Duration','Start','Finish','%Complete','Predecessors','Resources','Cost','Critical','Status','Notes'];
        const rows = project.tasks.map(t => [
            t.id, t.wbs||'', `"${(t.name||'').replace(/"/g,'""')}"`, t.durationDays,
            formatDate(t.start), formatDate(t.finish), t.percentComplete,
            (t.predecessors||[]).map(p=>p.predecessorUID).join(';'),
            `"${(t.resourceNames||[]).join(', ')}"`, t.cost||0,
            t.critical?'Yes':'No', t.status||'', `"${(t.notes||'').replace(/"/g,'""')}"`
        ]);
        const csv = [h.join(','), ...rows.map(r => r.join(','))].join('\n');
        downloadFile(csv, sanitize(project.name) + '.csv', 'text/csv');
        showToast('success', 'Exported CSV');
    }

    // ══════ NEW PROJECT ══════
    function showNewProjectModal() {
        $('inputProjectName').value = ''; $('inputProjectManager').value = '';
        $('inputStartDate').value = new Date().toISOString().split('T')[0];
        toggleModal('modalNewProject', true);
        setTimeout(() => $('inputProjectName').focus(), 200);
    }

    function handleCreateProject() {
        const name = $('inputProjectName').value.trim() || 'New Project';
        const manager = $('inputProjectManager').value.trim();
        const startDate = $('inputStartDate').value || new Date().toISOString().split('T')[0];
        const start = new Date(startDate);
        const d2 = addDays(start, 5), d3 = addDays(d2, 3), d4 = addDays(d3, 4), d5 = addDays(d3, 0);

        project = {
            name, manager, startDate: start, finishDate: d4,
            minutesPerDay: 480, minutesPerWeek: 2400, daysPerMonth: 20,
            currencySymbol: settings.currency,
            tasks: [
                mkTask(1, 'Phase 1: Planning', start, d4, 1, true),
                mkTask(2, 'Define project scope', start, d2, 2),
                mkTask(3, 'Create project plan', d2, d3, 2),
                mkTask(4, 'Kickoff meeting', d5, d5, 2, false, true),
                mkTask(5, 'Phase 2: Execution', d4, addDays(d4, 10), 1, true),
                mkTask(6, 'Task A', d4, addDays(d4, 5), 2),
                mkTask(7, 'Task B', addDays(d4, 3), addDays(d4, 8), 2),
                mkTask(8, 'Project Complete', addDays(d4, 10), addDays(d4, 10), 1, false, true),
            ],
            resources: [
                { uid: 1, id: 1, name: 'Team Lead', type: 1, maxUnits: 1, cost: 0 },
                { uid: 2, id: 2, name: 'Developer', type: 1, maxUnits: 1, cost: 0 },
            ],
            assignments: [
                { taskUID: 2, resourceUID: 1, units: 1 },
                { taskUID: 3, resourceUID: 1, units: 1 },
                { taskUID: 6, resourceUID: 2, units: 1 },
                { taskUID: 7, resourceUID: 2, units: 1 },
            ]
        };

        // B.1: Resource names resolved automatically by recalculate() via ProjectAnalytics Maps
        // (removed O(n²) nested filter+find — onProjectLoaded() → recalculate() handles this)

        project.tasks[2].predecessors = [{ predecessorUID: 2, type: 1, typeName: 'FS', lag: 0 }];
        project.tasks[3].predecessors = [{ predecessorUID: 3, type: 1, typeName: 'FS', lag: 0 }];
        project.tasks[6].predecessors = [{ predecessorUID: 6, type: 1, typeName: 'FS', lag: 0 }];
        project.tasks[7].predecessors = [{ predecessorUID: 7, type: 1, typeName: 'FS', lag: 0 }];

        reindexTasks();
        toggleModal('modalNewProject', false);
        // Multi-project: assign ID
        activeProjectId = ProjectStore.generateId();
        onProjectLoaded();
        showToast('success', `Created "${name}"`);
    }

    function mkTask(uid, name, start, finish, level, isSummary, isMilestone) {
        const s = new Date(start), f = new Date(finish);
        const dur = Math.max(Math.round((f - s) / 86400000), 0);
        return {
            uid, id: uid, name, wbs: '', outlineLevel: level, outlineNumber: '',
            start: s, finish: f, duration: MSProjectXML.daysToDuration(dur), durationDays: dur,
            percentComplete: 0, summary: isSummary || false, milestone: isMilestone || false,
            critical: false, cost: 0, notes: '', color: null,
            predecessors: [], resourceNames: [],
            baselineStart: null, baselineFinish: null, baselineDuration: null,
            totalFloat: null, freeFloat: null,
            status: 'not-started', statusIcon: '⬜', statusColor: '#64748b',
            isExpanded: true, isVisible: true,
            tags: [], comments: [], attachments: [],
            plannedHours: dur * 8, actualHours: 0, actualCost: null
        };
    }

    // ══════ PROJECT LOADED ══════
    function onProjectLoaded() {
        els.welcomeScreen.classList.add('hidden');
        els.workspace.classList.remove('hidden');
        els.projectNameDisplay.textContent = project.name;

        ['btnExport','btnExportExcel','btnAddTask','btnAddMilestone','btnDeleteTask','btnIndent','btnOutdent','btnZoomIn','btnZoomOut','btnSetBaseline','btnAutoLevel','btnReports','btnCustomFields','btnExportTrigger'].forEach(id => $(id) && ($(id).disabled = false));

        GanttChart.init(els.ganttCanvas, els.ganttHeader, {
            onTaskSelect: (task) => { selectedTaskIds.clear(); selectedTaskIds.add(task.uid); renderTable(); openDetailPanel(task); },
            onTaskUpdate: (task) => { saveUndoState(); recalculate(); renderAll(); autoSave(); }
        });

        // B.4: Reset analytics cache on new project load
        ProjectAnalytics.reset();

        undoStack = []; redoStack = []; updateUndoButtons();
        recalculate();
        renderAll();

        // B.3: Emit project:loaded event for modules listening on EventBus
        EventBus.emit('project:loaded', { project });
        // E.5: Notify Plugin System
        if (typeof PluginSystem !== 'undefined') PluginSystem.emit('project:loaded', { project });

        // Multi-project: save & update
        if (!activeProjectId) activeProjectId = ProjectStore.generateId();
        ProjectStore.setActive(activeProjectId);
        autoSave();
        renderSidebar();
    }

    // ══════ RECALCULATE ══════
    /**
     * Recalculate CPM + analytics after any data change.
     * Sprint B.1 — resource Maps built here in O(n) instead of O(n²)
     * Sprint B.4 — single ProjectAnalytics pass, cached until next mutation
     */
    function recalculate() {
        if (!project) return;

        // B.1: Apply resource names via O(n) Maps (before CPM so names are fresh)
        ProjectAnalytics.invalidate();
        const rMaps = ProjectAnalytics.buildResourceMaps(project);
        ProjectAnalytics.applyResourceNames(project, rMaps.taskResourceNames);

        // CPM engine (modifies task.critical, totalFloat, etc.)
        try {
            CPMEngine.compute(project.tasks, project.minutesPerDay || 480);
            CPMEngine.calculateVariance(project.tasks);
            CPMEngine.calculateStatus(project.tasks);
        } catch (err) {
            console.error('CPM Error:', err);
            showToast('error', 'CPM Error: ' + err.message);
        }

        // B.4: Full analytics pass after CPM (status flags are now current)
        ProjectAnalytics.invalidate();
        ProjectAnalytics.compute(project);
    }

    // ══════ RENDERING ══════
    /**
     * Schedule a full UI refresh (table + gantt + footer).
     * Uses requestAnimationFrame to batch rapid calls into a single frame.
     * Sprint B.2 — Render Throttling
     */
    function renderAll() {
        if (_renderPending) return; // already queued for this frame
        _renderPending = true;
        requestAnimationFrame(() => {
            _renderPending = false;
            _doRenderAll();
        });
    }

    /** Immediate render — used internally after RAF fires */
    function _doRenderAll() {
        updateVisibility();
        renderTable();
        renderGantt();
        updateFooter();
    }

    function getVisibleTasks() {
        if (!project) return [];
        let tasks = project.tasks.filter(t => t.isVisible !== false);

        // Search filter
        if (searchQuery) {
            tasks = tasks.filter(t => t.name.toLowerCase().includes(searchQuery) || (t.notes || '').toLowerCase().includes(searchQuery) || (t.resourceNames || []).join(' ').toLowerCase().includes(searchQuery));
        }

        // Status filter
        if (activeFilter !== 'all') {
            switch (activeFilter) {
                case 'critical': tasks = tasks.filter(t => t.critical); break;
                case 'late': tasks = tasks.filter(t => t.status === 'late'); break;
                case 'at-risk': tasks = tasks.filter(t => t.status === 'at-risk'); break;
                case 'complete': tasks = tasks.filter(t => t.percentComplete >= 100); break;
                case 'milestone': tasks = tasks.filter(t => t.milestone); break;
                case 'summary': tasks = tasks.filter(t => t.summary); break;
            }
        }

        // C.2: Advanced filter
        if (typeof _applyAdvFilter === 'function') tasks = _applyAdvFilter(tasks);

        // Sort
        if (sortColumn) {
            tasks = [...tasks].sort((a, b) => {
                let va, vb;
                switch (sortColumn) {
                    case 'id': va = a.id; vb = b.id; break;
                    case 'name': va = a.name.toLowerCase(); vb = b.name.toLowerCase(); break;
                    case 'duration': va = a.durationDays; vb = b.durationDays; break;
                    case 'start': va = new Date(a.start).getTime(); vb = new Date(b.start).getTime(); break;
                    case 'finish': va = new Date(a.finish).getTime(); vb = new Date(b.finish).getTime(); break;
                    case 'pct': va = a.percentComplete; vb = b.percentComplete; break;
                    default: return 0;
                }
                if (va < vb) return sortDir === 'asc' ? -1 : 1;
                if (va > vb) return sortDir === 'asc' ? 1 : -1;
                return 0;
            });
        }

        return tasks;
    }

    function renderTable() {
        if (!project) return;
        const tbody = els.taskTableBody; tbody.innerHTML = '';
        const visibleTasks = getVisibleTasks();

        // C.4: Update batch toolbar
        if (typeof _updateBatchToolbar === 'function') _updateBatchToolbar();

        // D.4: Custom field headers — add/remove th elements
        _syncCustomFieldHeaders();

        for (const task of visibleTasks) {
            const tr = document.createElement('tr');
            tr.dataset.uid = task.uid;
            // C.3: Make rows draggable
            tr.draggable = true;
            if (task.summary) tr.classList.add('summary-task');
            if (task.milestone) tr.classList.add('milestone');
            if (selectedTaskIds.has(task.uid)) tr.classList.add('selected');
            if (task.critical) tr.classList.add('task-critical');
            if (task.status === 'late') tr.classList.add('task-late');

            // Select
            const tdSel = document.createElement('td'); tdSel.className = 'col-select'; tdSel.style.textAlign = 'center';
            const cb = document.createElement('input'); cb.type = 'checkbox'; cb.checked = selectedTaskIds.has(task.uid);
            cb.addEventListener('change', () => { if (cb.checked) selectedTaskIds.add(task.uid); else selectedTaskIds.delete(task.uid); tr.classList.toggle('selected', cb.checked); });
            tdSel.appendChild(cb); tr.appendChild(tdSel);

            // Status
            const tdStatus = document.createElement('td'); tdStatus.className = 'col-status'; tdStatus.style.textAlign = 'center';
            tdStatus.textContent = task.statusIcon || '⬜'; tdStatus.title = task.status || ''; tr.appendChild(tdStatus);

            // ID
            addCell(tr, task.id, 'col-id');

            // WBS
            if (settings.showWBS) addCell(tr, task.wbs || '', 'col-wbs');

            // Task Name
            const tdName = document.createElement('td'); tdName.className = 'col-name';
            const nw = document.createElement('div'); nw.className = 'task-name-cell';
            const indent = document.createElement('span'); indent.className = 'task-indent'; indent.style.width = ((task.outlineLevel - 1) * 16) + 'px'; nw.appendChild(indent);
            if (task.summary) {
                const eb = document.createElement('button'); eb.className = 'task-expand-btn' + (task.isExpanded ? '' : ' collapsed');
                eb.innerHTML = '<svg viewBox="0 0 12 12" fill="currentColor"><path d="M3 4.5l3 3 3-3"/></svg>';
                eb.addEventListener('click', (e) => { e.stopPropagation(); task.isExpanded = !task.isExpanded; renderAll(); });
                nw.appendChild(eb);
            } else { const sp = document.createElement('span'); sp.style.cssText = 'width:16px;display:inline-block;flex-shrink:0'; nw.appendChild(sp); }
            const nt = document.createElement('span'); nt.className = 'task-name-text'; nt.textContent = task.name; nt.title = task.name; nw.appendChild(nt);
            tdName.appendChild(nw); tr.appendChild(tdName);

            // Duration (editable) — C.1: cellType 'duration' for smart parsing
            const durText = task.milestone ? '0d' : task.durationDays + 'd';
            addEditableCell(tr, durText, 'col-duration', (val) => {
                saveUndoState(); const d = parseInt(val); if (!isNaN(d) && d >= 0) { task.durationDays = d; task.finish = addDays(new Date(task.start), d); recalculate(); renderAll(); autoSave(); }
            }, 'duration');

            // Start (editable) — C.1: cellType 'date' + baseDate for smart parsing
            addEditableCell(tr, formatDate(task.start), 'col-start', (val) => {
                saveUndoState(); const d = parseInputDate(val); if (d) { task.start = d; task.finish = addDays(d, task.durationDays); recalculate(); renderAll(); autoSave(); }
            }, 'date', new Date(task.start));

            // Finish (editable) — C.1: smart date
            addEditableCell(tr, formatDate(task.finish), 'col-finish', (val) => {
                saveUndoState(); const d = parseInputDate(val); if (d) { task.finish = d; task.durationDays = Math.max(0, Math.round((d - new Date(task.start)) / 86400000)); recalculate(); renderAll(); autoSave(); }
            }, 'date', new Date(task.finish));

            // % Complete (editable) — C.1: with Tab support via dataset
            const tdPct = document.createElement('td'); tdPct.className = 'col-pct editable'; tdPct.dataset.cellType = 'pct';
            const progressColor = task.percentComplete >= 100 ? '#22c55e' : (task.status === 'late' ? '#ef4444' : '#6366f1');
            const progInner = document.createElement('div'); progInner.className = 'progress-cell';
            progInner.innerHTML = `<div class="progress-bar-mini"><div class="fill" style="width:${task.percentComplete}%;background:${progressColor}"></div></div><span class="progress-text">${task.percentComplete}%</span>`;
            tdPct.appendChild(progInner);
            tdPct.addEventListener('dblclick', () => startCellEdit(tdPct, task.percentComplete + '', (val) => {
                saveUndoState(); const p = Math.min(100, Math.max(0, parseInt(val) || 0)); task.percentComplete = p; recalculate(); renderAll(); autoSave();
            }, 'pct'));
            tr.appendChild(tdPct);

            // Float
            if (settings.showFloat) {
                const fv = task.totalFloat;
                const f = (fv != null && isFinite(fv)) ? fv + 'd' : '—';
                addCell(tr, f, 'col-float');
            }

            // Predecessors
            const predText = (task.predecessors || []).map(p => { const tn = p.typeName || 'FS'; return p.predecessorUID + (tn !== 'FS' ? tn : ''); }).join(', ');
            addCell(tr, predText, 'col-predecessors');

            // Resources — C.1: now editable inline
            addEditableCell(tr, (task.resourceNames || []).join(', '), 'col-resources', (val) => {
                saveUndoState();
                task.resourceNames = val.split(',').map(s => s.trim()).filter(Boolean);
                autoSave(); renderTable();
            }, 'text');

            // Cost
            if (settings.showCost) {
                addEditableCell(tr, task.cost > 0 ? task.cost.toFixed(0) : '', 'col-cost', (val) => {
                    saveUndoState(); task.cost = parseFloat(val) || 0; renderAll(); autoSave();
                }, 'cost');
            }

            // Custom Fields — D.4
            if (project.customFields && project.customFields.length) {
                project.customFields.forEach(cf => {
                    const cfVal = (task.customData || {})[cf.id];
                    const displayVal = cfVal != null ? String(cfVal) : '';
                    if (cf.type === 'dropdown') {
                        const tdCF = document.createElement('td'); tdCF.className = 'col-cf editable'; tdCF.textContent = displayVal;
                        tdCF.addEventListener('dblclick', () => {
                            const sel = document.createElement('select');
                            sel.innerHTML = `<option value="">—</option>` + (cf.options||[]).map(o => `<option${cfVal===o?' selected':''}>${escapeHTML(o)}</option>`).join('');
                            sel.className = 'cf-inline-select';
                            tdCF.textContent = '';
                            tdCF.appendChild(sel); sel.focus();
                            const done = () => { const v = sel.value; if (!task.customData) task.customData = {}; task.customData[cf.id] = v; saveUndoState(); autoSave(); renderTable(); };
                            sel.addEventListener('change', done);
                            sel.addEventListener('blur', done, { once: true });
                        });
                        tr.appendChild(tdCF);
                    } else {
                        addEditableCell(tr, displayVal, 'col-cf', (val) => {
                            saveUndoState(); if (!task.customData) task.customData = {}; task.customData[cf.id] = val; autoSave();
                        }, cf.type === 'number' ? 'cost' : 'text');
                    }
                });
            }

            // Color
            if (settings.showColor) {
                const tdColor = document.createElement('td'); tdColor.className = 'col-color'; tdColor.style.textAlign = 'center';
                const dot = document.createElement('span'); dot.className = 'color-dot';
                dot.style.background = task.color || '#6366f1';
                dot.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const input = document.createElement('input'); input.type = 'color'; input.value = task.color || '#6366f1';
                    input.style.cssText = 'position:absolute;opacity:0;width:0;height:0';
                    document.body.appendChild(input);
                    input.addEventListener('change', () => { saveUndoState(); task.color = input.value; dot.style.background = input.value; renderGantt(); autoSave(); document.body.removeChild(input); });
                    input.click();
                });
                tdColor.appendChild(dot); tr.appendChild(tdColor);
            }

            // Row click — C.4: Shift+Click range select
            tr.addEventListener('click', (e) => {
                if (e.target.tagName === 'INPUT' || e.target.classList.contains('color-dot')) return;
                if (e.shiftKey && selectedTaskIds.size > 0) {
                    // Range select: find last selected task and select range
                    const lastUid = [...selectedTaskIds].pop();
                    const lastIdx = visibleTasks.findIndex(t => t.uid === lastUid);
                    const thisIdx = visibleTasks.findIndex(t => t.uid === task.uid);
                    const lo = Math.min(lastIdx, thisIdx), hi = Math.max(lastIdx, thisIdx);
                    for (let i = lo; i <= hi; i++) selectedTaskIds.add(visibleTasks[i].uid);
                } else if (!e.metaKey && !e.ctrlKey) {
                    selectedTaskIds.clear();
                    selectedTaskIds.add(task.uid);
                } else {
                    // Cmd/Ctrl: toggle
                    if (selectedTaskIds.has(task.uid)) selectedTaskIds.delete(task.uid);
                    else selectedTaskIds.add(task.uid);
                }
                renderTable();
            });

            // Double-click name → edit
            tdName.addEventListener('dblclick', () => startCellEdit(nt, task.name, (val) => { saveUndoState(); task.name = val; renderAll(); autoSave(); }));

            // Right-click → detail panel
            tr.addEventListener('contextmenu', (e) => { e.preventDefault(); openDetailPanel(task); });

            tbody.appendChild(tr);
        }
    }

    /** D.4: Add/remove <th> elements for custom fields in the table header */
    function _syncCustomFieldHeaders() {
        const thead = els.taskTableBody ? els.taskTableBody.closest('table')?.querySelector('thead tr') : null;
        if (!thead) return;
        // Remove old custom field headers
        thead.querySelectorAll('th.col-cf').forEach(th => th.remove());
        // Add new ones
        (project && project.customFields || []).forEach(cf => {
            const th = document.createElement('th');
            th.className = 'col-cf';
            th.textContent = cf.name;
            th.title = cf.name + ' (' + cf.type + ')';
            thead.appendChild(th);
        });
    }

    /** D.2: Ensure heatmap container exists in resources view */
    function _ensureHeatmapContainer() {
        const rv = els.resourceView;
        if (!rv) return null;
        let wrap = rv.querySelector('#resourceHeatmapWrap');
        if (!wrap) {
            const section = document.createElement('div');
            section.className = 'resource-heatmap-section';
            section.innerHTML = '<div class="resource-section-title">📊 Resource Load Heatmap</div>';
            wrap = document.createElement('div');
            wrap.id = 'resourceHeatmapWrap';
            section.appendChild(wrap);
            rv.appendChild(section);
        }
        return wrap;
    }

    function renderResourceHeatmap() {
        const wrap = _ensureHeatmapContainer() || $('resourceHeatmapWrap');
        if (!wrap || !project) return;
        ResourceHeatmap.render(wrap, project);
    }

    function renderGantt() {
        if (!project) return;
        const visibleTasks = getVisibleTasks();
        const sel = Array.from(selectedTaskIds);
        const idx = sel.length === 1 ? visibleTasks.findIndex(t => t.uid === sel[0]) : -1;
        GanttChart.update(visibleTasks, {
            selectedIndex: idx,
            showBaseline: settings.showBaseline,
            showCritical: settings.showCritical,
            showLinks: settings.showLinks
        });
    }

    // ══════ CELL EDITING  (C.1 — Smart Inline Editing) ══════

    /**
     * Parse smart shorthand inputs into values:
     *   Durations : "5d" → 5 | "2w" → 14 | "3m" → 90
     *   Dates     : "+5" → today+5d | "mon" → next Monday
     *              | "2w" → today+14d | standard ISO / locale dates
     * @param {string} val
     * @param {'duration'|'date'|'pct'|'cost'} type
     * @param {Date} [baseDate] — current value, used for relative offsets
     * @returns {string} normalised value ready for the existing parsers
     */
    function parseSmartInput(val, type, baseDate) {
        val = (val || '').trim();
        if (!val) return val;

        if (type === 'duration') {
            const m = val.match(/^(\d+)\s*([dwmh])?$/i);
            if (m) {
                const n = parseInt(m[1]);
                const u = (m[2] || 'd').toLowerCase();
                if (u === 'w') return String(n * 7);
                if (u === 'm') return String(Math.round(n * 30.44));
                if (u === 'h') return String(Math.round(n / (settings.hoursPerDay || 8)));
                return String(n);
            }
            return val;
        }

        if (type === 'date') {
            const base = baseDate instanceof Date && !isNaN(baseDate) ? baseDate : new Date();
            const low = val.toLowerCase();
            // +N / -N days relative
            const rel = val.match(/^([+-])(\d+)([dwm]?)$/i);
            if (rel) {
                const sign = rel[1] === '+' ? 1 : -1;
                let n = parseInt(rel[2]);
                const u = (rel[3] || 'd').toLowerCase();
                if (u === 'w') n *= 7;
                if (u === 'm') n = Math.round(n * 30.44);
                const r = addDays(base, sign * n);
                return r.toISOString().split('T')[0];
            }
            // Weekday names: mon, tue, wed, thu, fri
            const days = { sun:0, mon:1, tue:2, wed:3, thu:4, fri:5, sat:6 };
            const dname = low.slice(0, 3);
            if (days[dname] !== undefined) {
                const target = days[dname];
                const d = new Date(); d.setHours(0,0,0,0);
                let diff = target - d.getDay();
                if (diff <= 0) diff += 7;
                d.setDate(d.getDate() + diff);
                return d.toISOString().split('T')[0];
            }
            // "today", "now"
            if (low === 'today' || low === 'now') return new Date().toISOString().split('T')[0];
            // "tomorrow"
            if (low === 'tomorrow' || low === 'tom') return addDays(new Date(), 1).toISOString().split('T')[0];
            // Relative duration offset: "2w", "5d" from base
            const relDur = val.match(/^(\d+)\s*([dwm])$/i);
            if (relDur) {
                let n = parseInt(relDur[1]);
                const u = relDur[2].toLowerCase();
                if (u === 'w') n *= 7;
                if (u === 'm') n = Math.round(n * 30.44);
                return addDays(base, n).toISOString().split('T')[0];
            }
        }

        return val;
    }

    /**
     * Create an editable TD cell.
     * @param {'duration'|'date'|'pct'|'cost'|'text'} [cellType] — for smart parsing
     * @param {Date} [baseDate] — reference date for relative parsing
     */
    function addEditableCell(tr, text, className, onCommit, cellType, baseDate) {
        const td = document.createElement('td');
        td.className = className + ' editable';
        td.dataset.cellType = cellType || 'text';
        td.textContent = text; td.title = String(text);
        td.addEventListener('dblclick', () => startCellEdit(td, text, onCommit, cellType, baseDate));
        tr.appendChild(td);
        return td;
    }

    /**
     * Activate inline editing on a cell element.
     * C.1: Tab moves to the next .editable cell in the same row.
     * C.1: Smart shorthand input is parsed before commit.
     */
    function startCellEdit(element, currentValue, onCommit, cellType, baseDate) {
        if (element.contentEditable === 'true') return; // already editing
        const old = String(currentValue ?? '');
        element.contentEditable = true;
        element.focus();
        element.classList.add('cell-editing');
        element.textContent = old;

        // Select all text
        try {
            const sel = window.getSelection(); const range = document.createRange();
            range.selectNodeContents(element); sel.removeAllRanges(); sel.addRange(range);
        } catch(_) {}

        let committed = false;
        const commit = () => {
            if (committed) return; committed = true;
            element.contentEditable = false;
            element.classList.remove('cell-editing');
            let nv = element.textContent.trim();
            // C.1: Apply smart parsing if a cellType hint is provided
            if (cellType && cellType !== 'text') nv = parseSmartInput(nv, cellType, baseDate);
            if (nv !== old && onCommit) onCommit(nv);
        };

        element.addEventListener('blur', commit, { once: true });
        element.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); element.blur(); }
            else if (e.key === 'Escape') { element.textContent = old; committed = true; element.contentEditable = false; element.classList.remove('cell-editing'); }
            else if (e.key === 'Tab') {
                // C.1: Tab navigation — move to next/prev editable cell in the row
                e.preventDefault();
                element.blur(); // commit current
                const row = element.closest('tr');
                if (!row) return;
                const editables = Array.from(row.querySelectorAll('td.editable'));
                const idx = editables.indexOf(element);
                const next = e.shiftKey ? editables[idx - 1] : editables[idx + 1];
                if (next) next.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
            }
        });
    }

    // ══════ DETAIL PANEL ══════
    function openDetailPanel(task) {
        detailTask = task;
        els.detailPanel.classList.remove('hidden');
        $('detailTaskName').textContent = task.name;
        $('detailName').value = task.name;
        $('detailStart').value = toDateInput(task.start);
        $('detailFinish').value = toDateInput(task.finish);
        $('detailDuration').value = task.durationDays;
        $('detailPct').value = task.percentComplete;
        $('detailCost').value = task.cost || 0;
        $('detailColor').value = task.color || '#6366f1';
        $('detailPredecessors').value = (task.predecessors || []).map(p => p.predecessorUID + (p.typeName || 'FS')).join(', ');
        $('detailResources').value = (task.resourceNames || []).join(', ');
        $('detailNotes').value = task.notes || '';
        $('detailStatus').textContent = (task.statusIcon || '') + ' ' + (task.status || '—');
        $('detailCritical').textContent = task.critical ? '🔴 Yes' : '🟢 No';
        $('detailFloat').textContent = task.totalFloat != null ? task.totalFloat + ' days' : '—';
        $('detailBaseStart').textContent = task.baselineStart ? formatDate(task.baselineStart) : '—';
        $('detailBaseFinish').textContent = task.baselineFinish ? formatDate(task.baselineFinish) : '—';
        $('detailStartVar').textContent = task.startVariance ? task.startVariance + 'd' : '0d';
        $('detailFinishVar').textContent = task.finishVariance ? task.finishVariance + 'd' : '0d';

        // Time Tracking (Phase 3.3)
        $('detailPlannedHours').value = task.plannedHours || (task.durationDays * (settings.hoursPerDay || 8));
        $('detailActualHours').value = task.actualHours || 0;

        // Phase 5: Tags, Comments, Attachments
        renderTagsInDetailPanel(task);
        renderCommentsInDetailPanel(task);
        renderAttachmentsInDetailPanel(task);
    }

    function closeDetailPanel() { els.detailPanel.classList.add('hidden'); detailTask = null; }

    function saveDetailPanel() {
        if (!detailTask) return;
        saveUndoState();
        detailTask.name = $('detailName').value.trim() || detailTask.name;
        const newStart = new Date($('detailStart').value);
        const newFinish = new Date($('detailFinish').value);
        if (!isNaN(newStart.getTime())) detailTask.start = newStart;
        if (!isNaN(newFinish.getTime())) detailTask.finish = newFinish;
        detailTask.durationDays = parseInt($('detailDuration').value) || 0;
        detailTask.percentComplete = Math.min(100, Math.max(0, parseInt($('detailPct').value) || 0));
        detailTask.cost = parseFloat($('detailCost').value) || 0;
        detailTask.color = $('detailColor').value !== '#6366f1' ? $('detailColor').value : null;
        detailTask.notes = $('detailNotes').value;
        detailTask.resourceNames = $('detailResources').value.split(',').map(s => s.trim()).filter(Boolean);

        // Time Tracking (Phase 3.3)
        detailTask.plannedHours = parseFloat($('detailPlannedHours').value) || 0;
        detailTask.actualHours = parseFloat($('detailActualHours').value) || 0;

        recalculate(); renderAll(); autoSave();
        openDetailPanel(detailTask); // Refresh
        showToast('info', 'Task updated');
    }

    // ══════ RESOURCES ══════
    function handleAddResource() {
        if (!project) return;
        const name = $('inputResName').value.trim();
        if (!name) { showToast('warning', 'Enter resource name'); return; }
        const uid = project.resources.reduce((max, r) => Math.max(max, r.uid || 0), 0) + 1;
        project.resources.push({
            uid, id: uid, name,
            type: 1, maxUnits: parseFloat($('inputResUnits').value) || 1,
            cost: parseFloat($('inputResCost').value) || 0
        });
        toggleModal('modalAddResource', false);
        $('inputResName').value = '';
        renderResources();
        showToast('info', `Added resource "${name}"`);
        autoSave();
    }

    function renderResources() {
        if (!project) return;
        const tbody = els.resourceTableBody; tbody.innerHTML = '';
        const { resourceLoads, overAllocations } = ResourceManager.calculateResourceLoad(project.tasks, project.resources, project.assignments || []);
        const summary = ResourceManager.getUtilizationSummary(resourceLoads, project.startDate, project.finishDate);

        for (const res of project.resources) {
            const s = summary.find(su => su.uid === res.uid) || {};
            const tr = document.createElement('tr');
            addCell(tr, res.uid); addCell(tr, res.name);
            addCell(tr, res.type === 1 ? 'Work' : 'Material');
            addCell(tr, (res.maxUnits || 1) * 100 + '%');
            addCell(tr, settings.currency + (res.cost || 0));
            addCell(tr, (s.totalHours || 0) + 'h');
            addCell(tr, (s.utilization || 0) + '%');
            const tdSt = document.createElement('td');
            tdSt.innerHTML = s.overAllocated ? '<span class="res-over">Over-allocated</span>' : '<span class="res-ok">OK</span>';
            tr.appendChild(tdSt);
            tbody.appendChild(tr);
        }

        // Time Tracking Summary (Phase 3.3)
        renderTimeTrackingSummary();
    }

    function renderTimeTrackingSummary() {
        if (!project || !project.resources || project.resources.length === 0) return;
        let container = $('timeTrackingSummary');
        if (!container) {
            // Create if doesn't exist yet — append after resource table
            container = document.createElement('div');
            container.id = 'timeTrackingSummary';
            container.className = 'time-tracking-section';
            els.resourceTableBody.closest('.resource-panel')?.appendChild(container);
            if (!container.parentElement) {
                // Fallback: append after the resource table body's parent
                els.resourceTableBody.closest('table')?.parentElement?.appendChild(container);
            }
        }
        if (!container) return;

        const summary = ResourceManager.getTimeTrackingSummary(
            project.tasks, project.resources, project.assignments || []
        );

        let html = '<h4 style="margin:16px 0 8px;color:var(--text-secondary)">⏱️ Time Tracking Summary</h4>';
        html += '<table class="data-table" style="margin-bottom:12px"><thead><tr>';
        html += '<th>Resource</th><th>Planned Hours</th><th>Actual Hours</th><th>Variance</th><th>Efficiency</th>';
        html += '</tr></thead><tbody>';

        for (const s of summary) {
            const varClass = s.variance > 0 ? 'color:#ef4444' : s.variance < 0 ? 'color:#22c55e' : '';
            const effClass = s.efficiency >= 90 ? 'color:#22c55e' : s.efficiency >= 70 ? 'color:#eab308' : 'color:#ef4444';
            html += `<tr>`;
            html += `<td>${escapeHTML(s.name)}</td>`;
            html += `<td>${s.plannedHours}h</td>`;
            html += `<td>${s.actualHours}h</td>`;
            html += `<td style="${varClass}">${s.variance > 0 ? '+' : ''}${s.variance}h</td>`;
            html += `<td style="${effClass}">${s.efficiency}%</td>`;
            html += `</tr>`;
        }
        html += '</tbody></table>';
        container.innerHTML = html;
    }

    // ══════ BASELINE ══════
    function handleSetBaseline() {
        if (!project) return;
        saveUndoState();
        CPMEngine.setBaseline(project.tasks);
        CPMEngine.calculateVariance(project.tasks);
        renderAll();
        showToast('success', 'Baseline saved for all tasks');
        autoSave();
    }

    // ══════ RESOURCE LEVELING (Phase 3.2) ══════
    function handleAutoLevel() {
        if (!project) return;
        if (!project.resources || project.resources.length === 0) {
            showToast('warning', 'No resources defined — add resources first');
            return;
        }
        if (!project.assignments || project.assignments.length === 0) {
            showToast('warning', 'No assignments found — assign resources to tasks first');
            return;
        }

        saveUndoState();
        const changes = ResourceManager.autoLevel(
            project.tasks, project.resources, project.assignments
        );

        if (changes.length === 0) {
            showToast('info', 'No over-allocations found — resources are balanced');
            return;
        }

        recalculate();
        renderAll();
        autoSave();

        // Build summary message
        const names = changes.map(c => `"${c.taskName}" → +${c.delayed}d`).join('\n');
        showToast('success', `Leveled ${changes.length} task(s). Undo to revert.`);
        console.log('Resource Leveling Results:', changes);
    }

    // ══════ TASK ACTIONS ══════
    function handleAddTask() {
        if (!project) return;
        saveUndoState();
        const last = project.tasks[project.tasks.length - 1];
        const uid = project.tasks.reduce((max, t) => Math.max(max, t.uid || 0), 0) + 1;
        const start = last ? new Date(last.finish) : new Date();
        const task = mkTask(uid, 'New Task', start, addDays(start, 5), 1);
        project.tasks.push(task);
        reindexTasks(); recalculate(); renderAll(); autoSave();
        selectedTaskIds.clear(); selectedTaskIds.add(task.uid); renderTable();
        setTimeout(() => { els.tableWrapper.scrollTop = els.tableWrapper.scrollHeight; }, 50);
        showToast('info', 'Task added');
    }

    function handleAddMilestone() {
        if (!project) return;
        saveUndoState();
        const last = project.tasks[project.tasks.length - 1];
        const uid = project.tasks.reduce((max, t) => Math.max(max, t.uid || 0), 0) + 1;
        const date = last ? new Date(last.finish) : new Date();
        const task = mkTask(uid, 'New Milestone', date, date, 1, false, true);
        task.durationDays = 0;
        project.tasks.push(task);
        reindexTasks(); recalculate(); renderAll(); autoSave();
        showToast('info', 'Milestone added');
    }

    function handleDeleteTasks() {
        if (!project || selectedTaskIds.size === 0) return;
        saveUndoState();
        project.tasks = project.tasks.filter(t => !selectedTaskIds.has(t.uid));
        selectedTaskIds.clear();
        reindexTasks(); recalculate(); renderAll(); autoSave();
        showToast('info', 'Deleted');
    }

    function handleIndent(dir) {
        if (!project || selectedTaskIds.size === 0) return;
        saveUndoState();
        project.tasks.forEach(t => { if (selectedTaskIds.has(t.uid)) t.outlineLevel = Math.max(1, (t.outlineLevel || 1) + dir); });
        for (let i = 0; i < project.tasks.length; i++) {
            const c = project.tasks[i], n = project.tasks[i + 1];
            c.summary = !!(n && n.outlineLevel > c.outlineLevel);
        }
        reindexTasks(); recalculate(); renderAll(); autoSave();
    }

    // ══════ VISIBILITY ══════
    function updateVisibility() {
        if (!project) return;
        let hideBelow = Infinity;
        for (const t of project.tasks) {
            if (t.outlineLevel <= hideBelow) {
                t.isVisible = true;
                hideBelow = (t.summary && !t.isExpanded) ? t.outlineLevel : Infinity;
            } else { t.isVisible = false; }
        }
    }

    // ══════ UNDO/REDO ══════
    function saveUndoState() { if (!project) return; undoStack.push(JSON.stringify(project)); if (undoStack.length > MAX_UNDO) undoStack.shift(); redoStack = []; updateUndoButtons(); }

    // Centralized mutation wrapper — ensures undo state is always saved before changes (TD-04)
    // B.4: Also invalidates analytics cache so next recalculate() recomputes fresh
    function mutation(fn) {
        saveUndoState();
        fn();
        ProjectAnalytics.invalidate();
        recalculate();
        renderAll();
        autoSave();
        EventBus.emit('project:changed', { project, source: 'mutation' });
    }
    function handleUndo() { if (!undoStack.length) return; redoStack.push(JSON.stringify(project)); restoreProject(undoStack.pop()); }
    function handleRedo() { if (!redoStack.length) return; undoStack.push(JSON.stringify(project)); restoreProject(redoStack.pop()); }
    function restoreProject(json) {
        project = JSON.parse(json);
        project.tasks.forEach(t => { t.start = new Date(t.start); t.finish = new Date(t.finish); if (t.baselineStart) t.baselineStart = new Date(t.baselineStart); if (t.baselineFinish) t.baselineFinish = new Date(t.baselineFinish); });
        if (project.startDate) project.startDate = new Date(project.startDate);
        if (project.finishDate) project.finishDate = new Date(project.finishDate);
        ProjectAnalytics.reset();
        updateUndoButtons();
        recalculate();
        renderAll();
    }
    function updateUndoButtons() { $('btnUndo').disabled = !undoStack.length; $('btnRedo').disabled = !redoStack.length; }

    // ══════ VIEW ══════
    function setView(view) {
        activeView = view;
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.view === view));
        if (window.NavOverflow) NavOverflow.setActive(view);
        els.splitContainer.className = 'split-container view-' + view;

        // Hide all extra views
        els.resourceView.classList.add('hidden');
        if (els.calendarView)   els.calendarView.classList.add('hidden');
        if (els.dashboardView)  els.dashboardView.classList.add('hidden');
        if (els.networkView)    els.networkView.classList.add('hidden');
        if (els.portfolioView)  els.portfolioView.classList.add('hidden');
        const boardView = $('boardView');
        if (boardView) boardView.classList.add('hidden');

        if (view === 'resources') {
            els.resourceView.classList.remove('hidden');
            els.splitContainer.style.display = 'none';
            renderResources();
            // D.2: Also render heatmap inside resources view
            setTimeout(() => renderResourceHeatmap(), 50);
        } else if (view === 'calendar') {
            els.calendarView.classList.remove('hidden');
            els.splitContainer.style.display = 'none';
            renderCalendarView();
        } else if (view === 'dashboard') {
            els.dashboardView.classList.remove('hidden');
            els.splitContainer.style.display = 'none';
            renderDashboard();
        } else if (view === 'network') {
            els.networkView.classList.remove('hidden');
            els.splitContainer.style.display = 'none';
            renderNetworkView();
        } else if (view === 'portfolio') {
            els.portfolioView.classList.remove('hidden');
            els.splitContainer.style.display = 'none';
            renderPortfolioView();
        } else if (view === 'board') {
            // D.1: Kanban board
            if (!project) { showToast('info', 'Open a project first'); return; }
            if (boardView) boardView.classList.remove('hidden');
            els.splitContainer.style.display = 'none';
            renderBoardView();
        } else {
            // Project-specific views (gantt, table, split) need an open project
            if (!project) {
                els.workspace.classList.add('hidden');
                els.welcomeScreen.classList.remove('hidden');
                renderProjectHub();
                return;
            }
            els.splitContainer.style.display = '';
            if (view === 'gantt' || view === 'split') setTimeout(() => { GanttChart.resize(); GanttChart.render(); }, 50);
        }
    }

    // ══════ SETTINGS ══════
    function populateSettingsModal() {
        $('settingDateFormat').value = settings.dateFormat;
        $('settingCurrency').value = settings.currency;
        $('settingHoursPerDay').value = settings.hoursPerDay;
        $('settingShowWBS').checked = settings.showWBS;
        $('settingShowCost').checked = settings.showCost;
        $('settingShowFloat').checked = settings.showFloat;
        $('settingShowColor').checked = settings.showColor;
        $('settingShowBaseline').checked = settings.showBaseline;
        $('settingShowCritical').checked = settings.showCritical;
        $('settingShowLinks').checked = settings.showLinks;
        // Logo preview
        const savedLogo = localStorage.getItem('pf_report_logo');
        if (savedLogo) {
            $('logoPreviewImg').src = savedLogo;
            $('logoPreviewImg').style.display = 'block';
            $('logoPlaceholder').style.display = 'none';
        } else {
            $('logoPreviewImg').style.display = 'none';
            $('logoPlaceholder').style.display = '';
        }
        // Server settings
        $('settingServerURL').value = _serverURL;
        const badge = $('serverStatusBadge');
        if (_serverMode) {
            badge.textContent = '🟢 Server Connected';
            badge.className = 'server-badge online';
            $('btnTestServer').click(); // Refresh DB list
        } else {
            badge.textContent = '🔴 Browser Mode';
            badge.className = 'server-badge offline';
            $('serverSettingsPanel').classList.add('hidden');
        }
    }

    function handleSaveSettings() {
        settings.dateFormat = $('settingDateFormat').value;
        settings.currency = $('settingCurrency').value || '$';
        settings.hoursPerDay = parseInt($('settingHoursPerDay').value) || 8;
        settings.showWBS = $('settingShowWBS').checked;
        settings.showCost = $('settingShowCost').checked;
        settings.showFloat = $('settingShowFloat').checked;
        settings.showColor = $('settingShowColor').checked;
        settings.showBaseline = $('settingShowBaseline').checked;
        settings.showCritical = $('settingShowCritical').checked;
        settings.showLinks = $('settingShowLinks').checked;
        settings.serverUrl = undefined; // MPP removed
        localStorage.setItem('pf_settings', JSON.stringify(settings));
        toggleModal('modalSettings', false);
        applyColumnVisibility();
        if (project) { GanttChart.setOptions({ showBaseline: settings.showBaseline, showCritical: settings.showCritical, showLinks: settings.showLinks }); renderAll(); }
        showToast('info', 'Settings saved');
    }

    function loadSettings() {
        try { const s = localStorage.getItem('pf_settings'); if (s) Object.assign(settings, JSON.parse(s)); } catch (e) {}
        applyColumnVisibility();
        // Load saved report logo
        const savedLogo = localStorage.getItem('pf_report_logo');
        if (savedLogo) window.PROART_LOGO = savedLogo;
    }

    function applyColumnVisibility() {
        const toggle = (cls, show) => document.querySelectorAll('.' + cls).forEach(el => el.style.display = show ? '' : 'none');
        toggle('col-wbs', settings.showWBS); toggle('col-cost', settings.showCost);
        toggle('col-float', settings.showFloat); toggle('col-color', settings.showColor);
    }

    // ══════ AUTO-SAVE (Fixed race condition — TD-03) ══════
    let _saveInFlight = false;
    let _saveRetries = 0;
    const MAX_SAVE_RETRIES = 3;
    function autoSave() {
        if (!project || !activeProjectId) return;
        _isDirty = true;
        clearTimeout(autoSaveTimer);
        autoSaveTimer = setTimeout(async () => {
            if (_saveInFlight) return; // Single-writer queue
            _saveInFlight = true;
            try {
                await ProjectStore.save(activeProjectId, project);
                await ProjectStore.addToIndex(activeProjectId, project);
                _isDirty = false;
                _saveRetries = 0;
                els.autoSaveStatus.textContent = '✓ Saved';
                els.autoSaveStatus.classList.add('visible');
                setTimeout(() => els.autoSaveStatus.classList.remove('visible'), 2000);
            } catch (e) {
                console.warn('Auto-save failed', e);
                _saveRetries++;
                if (_saveRetries < MAX_SAVE_RETRIES) {
                    els.autoSaveStatus.textContent = `⚠ Retry ${_saveRetries}/${MAX_SAVE_RETRIES}…`;
                    els.autoSaveStatus.classList.add('visible');
                    // Retry with exponential backoff
                    _saveInFlight = false;
                    setTimeout(() => autoSave(), 1000 * Math.pow(2, _saveRetries));
                    return;
                }
                els.autoSaveStatus.textContent = '❌ Save failed';
                els.autoSaveStatus.classList.add('visible');
            } finally {
                _saveInFlight = false;
            }
        }, 800);
    }

    // Warn before closing with unsaved changes
    window.addEventListener('beforeunload', (e) => {
        if (_isDirty) { e.preventDefault(); e.returnValue = ''; }
    });

    function saveToRecent() {
        // Legacy — no longer primary, but keep for compat
    }

    function loadRecentProjects() {
        // Replaced by renderProjectHub()
    }

    // ══════ KEYBOARD ══════
    function handleKeyboard(e) {
        // Allow escape even in inputs (for global search)
        if (e.key === 'Escape') {
            if (!$('globalSearchOverlay').classList.contains('hidden')) {
                closeGlobalSearch(); e.preventDefault(); return;
            }
            if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'TEXTAREA') {
                selectedTaskIds.clear(); closeDetailPanel(); renderTable();
            }
            return;
        }
        // Global search shortcut: ⌘/ or ⌘G
        if ((e.metaKey || e.ctrlKey) && (e.key === '/' || e.key === 'g')) {
            e.preventDefault(); openGlobalSearch(); return;
        }
        // Sidebar toggle: ⌘B
        if ((e.metaKey || e.ctrlKey) && e.key === 'b') { e.preventDefault(); toggleSidebar(); return; }
        // ⌘1-9: switch projects by index
        if ((e.metaKey || e.ctrlKey) && e.key >= '1' && e.key <= '9') {
            e.preventDefault();
            const idx = parseInt(e.key) - 1;
            const index = ProjectStore.getIndex();
            if (idx < index.length) { switchProject(index[idx].id); }
            return;
        }
        if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA' || e.target.isContentEditable) return;
        if ((e.metaKey || e.ctrlKey) && e.key === 'z' && !e.shiftKey) { e.preventDefault(); handleUndo(); }
        if ((e.metaKey || e.ctrlKey) && e.key === 'z' && e.shiftKey) { e.preventDefault(); handleRedo(); }
        if ((e.metaKey || e.ctrlKey) && e.key === 'f') { e.preventDefault(); els.searchInput.focus(); }
        if ((e.metaKey || e.ctrlKey) && e.key === 's') { e.preventDefault(); autoSave(); showToast('info', 'Saved'); }
        if (e.key === 'Delete' || (e.key === 'Backspace' && !e.target.closest('td'))) { if (selectedTaskIds.size > 0) { e.preventDefault(); handleDeleteTasks(); } }
    }

    // ══════ RESIZE HANDLE ══════
    function initResizeHandle() {
        const handle = $('resizeHandle'); let isR = false, sX, sW;
        handle.addEventListener('mousedown', (e) => { isR = true; sX = e.clientX; sW = $('tablePanel').offsetWidth; handle.classList.add('active'); document.body.style.cursor = 'col-resize'; document.body.style.userSelect = 'none'; });
        document.addEventListener('mousemove', (e) => { if (!isR) return; const nw = Math.max(300, Math.min(sW + e.clientX - sX, window.innerWidth - 300)); $('tablePanel').style.flex = `0 0 ${nw}px`; });
        document.addEventListener('mouseup', () => { if (isR) { isR = false; handle.classList.remove('active'); document.body.style.cursor = ''; document.body.style.userSelect = ''; if (project) { GanttChart.resize(); GanttChart.render(); } } });

        // B.2: ResizeObserver on Gantt panel — replaces window.onresize polling
        const ganttPanel = $('ganttPanel');
        if (ganttPanel && 'ResizeObserver' in window) {
            _ganttResizeObserver = new ResizeObserver(debounce(() => {
                if (project) { GanttChart.resize(); GanttChart.render(); }
            }, 120));
            _ganttResizeObserver.observe(ganttPanel);
        }
    }

    function syncScroll() {
        if (_syncLock) return; _syncLock = true;
        els.ganttBody.scrollTop = els.tableWrapper.scrollTop;
        requestAnimationFrame(() => _syncLock = false);
    }
    function syncScrollReverse() {
        if (_syncLock) return; _syncLock = true;
        els.tableWrapper.scrollTop = els.ganttBody.scrollTop;
        requestAnimationFrame(() => _syncLock = false);
    }
    let _syncLock = false;

    // ══════ HELPERS ══════
    function reindexTasks() {
        if (!project) return;

        // Check if tasks already have valid outline numbers (e.g. from Planner import)
        const hasOriginalOutline = project.tasks.some(t => t.outlineNumber && t.outlineNumber.length > 0);

        if (!hasOriginalOutline) {
            // Recalculate WBS from scratch
            let wbs = [0];
            for (let i = 0; i < project.tasks.length; i++) {
                const t = project.tasks[i]; t.id = i + 1;
                const lv = t.outlineLevel || 1;
                while (wbs.length < lv) wbs.push(0); wbs.length = lv;
                wbs[lv - 1] = (wbs[lv - 1] || 0) + 1;
                t.wbs = wbs.join('.'); t.outlineNumber = t.wbs;
            }
        } else {
            // Just reassign sequential IDs, keep original WBS
            for (let i = 0; i < project.tasks.length; i++) {
                project.tasks[i].id = i + 1;
            }
        }

        if (project.tasks.length > 0) {
            let minS = Infinity, maxF = -Infinity;
            for (const t of project.tasks) { const s = new Date(t.start).getTime(); const f = new Date(t.finish).getTime(); if (s < minS) minS = s; if (f > maxF) maxF = f; }
            project.startDate = new Date(minS); project.finishDate = new Date(maxF);
        }
    }

    function addCell(tr, text, cls) { const td = document.createElement('td'); if (cls) td.className = cls; td.textContent = text; td.title = text; tr.appendChild(td); return td; }

    function formatDate(date) {
        if (!date) return ''; const d = new Date(date); if (isNaN(d.getTime())) return '';
        const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, '0'), day = String(d.getDate()).padStart(2, '0');
        switch (settings.dateFormat) {
            case 'DD/MM/YYYY': return `${day}/${m}/${y}`;
            case 'MM/DD/YYYY': return `${m}/${day}/${y}`;
            case 'DD MMM YYYY': return d.toLocaleDateString('en-US', { day: '2-digit', month: 'short', year: 'numeric' });
            default: return `${y}-${m}-${day}`;
        }
    }

    function toDateInput(date) { const d = new Date(date); return d.toISOString().split('T')[0]; }
    function parseInputDate(str) { const d = new Date(str); return isNaN(d.getTime()) ? null : d; }
    function addDays(d, n) { const r = new Date(d); r.setDate(r.getDate() + n); return r; }

    function updateFooter() {
        if (!project) return;
        const c = project.tasks.length;
        els.taskCount.textContent = c + (c === 1 ? ' task' : ' tasks');
        if (c > 0) {
            const days = Math.round((new Date(project.finishDate) - new Date(project.startDate)) / 86400000);
            els.projectDuration.textContent = days + 'd';
            // B.4: Use cached analytics instead of re-filtering
            const analytics = ProjectAnalytics.getCache();
            const critCount = analytics ? analytics.critCount : project.tasks.filter(t => t.critical && !t.summary).length;
            els.criticalInfo.textContent = critCount > 0 ? `${critCount} critical` : 'No critical';
            els.criticalInfo.style.color = critCount > 0 ? '#ef4444' : '';
        } else { els.projectDuration.textContent = '—'; els.criticalInfo.textContent = '—'; }
    }

    function setStatus(text) {
        const dot = document.createElement('span'); dot.className = 'status-dot';
        els.statusIndicator.textContent = '';
        els.statusIndicator.appendChild(dot);
        els.statusIndicator.appendChild(document.createTextNode(' ' + text));
    }
    function toggleModal(id, show) { $(id).classList.toggle('hidden', !show); }

    function showToast(type, msg) {
        const t = document.createElement('div'); t.className = `toast ${type}`;
        const icons = { success: '✅', error: '❌', info: 'ℹ️', warning: '⚠️' };
        // Fix XSS: use textContent instead of innerHTML (TD-01)
        const iconSpan = document.createElement('span'); iconSpan.className = 'toast-icon'; iconSpan.textContent = icons[type] || '';
        const msgSpan = document.createElement('span'); msgSpan.textContent = msg;
        t.appendChild(iconSpan); t.appendChild(msgSpan);
        els.toastContainer.appendChild(t);
        setTimeout(() => { t.classList.add('toast-out'); setTimeout(() => t.remove(), 300); }, 3000);
    }

    function downloadFile(content, filename, mime) { const b = new Blob([content], { type: mime }); const u = URL.createObjectURL(b); const a = document.createElement('a'); a.href = u; a.download = filename; document.body.appendChild(a); a.click(); document.body.removeChild(a); URL.revokeObjectURL(u); }
    function sanitize(n) { return (n || 'project').replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 50); }
    function debounce(fn, ms) { let t; return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); }; }

    // ══════ PHASE 2: REPORT HANDLERS ══════
    async function handleReportPDF() {
        if (!project) return;
        const btn = $('rptPDF');
        btn.classList.add('loading');
        setStatus('Generating PDF…');
        try {
            if (typeof window.jspdf === 'undefined') {
                showToast('error', 'PDF library not loaded. Check internet connection.');
                return;
            }
            const canvas = $('ganttCanvas');
            const filename = await Reports.generatePDF(project, settings, canvas);
            toggleModal('modalReport', false);
            showToast('success', `PDF saved: ${filename}`);
        } catch (err) {
            showToast('error', 'PDF generation failed: ' + err.message);
        } finally {
            btn.classList.remove('loading');
            setStatus('Ready');
        }
    }

    function handleReportExcel() {
        if (!project) return;
        const btn = $('rptExcel');
        btn.classList.add('loading');
        setStatus('Generating Excel…');
        try {
            if (typeof XLSX === 'undefined') {
                showToast('error', 'Excel library not loaded. Check internet connection.');
                return;
            }
            const wb = Reports.exportExcel(project, settings);
            if (wb) {
                Reports.downloadExcel(wb, sanitize(project.name) + '.xlsx');
                toggleModal('modalReport', false);
                showToast('success', 'Excel file exported successfully');
            }
        } catch (err) {
            showToast('error', 'Excel export failed: ' + err.message);
        } finally {
            btn.classList.remove('loading');
            setStatus('Ready');
        }
    }

    function handleReportGanttPNG() {
        if (!project) return;
        const canvas = $('ganttCanvas');
        if (!canvas || canvas.width === 0) {
            showToast('warning', 'Switch to Gantt or Split view first');
            return;
        }
        try {
            Reports.exportGanttPNG(canvas);
            toggleModal('modalReport', false);
            showToast('success', 'Gantt chart saved as PNG');
        } catch (err) {
            showToast('error', 'PNG export failed: ' + err.message);
        }
    }

    async function handleReportSummary(e) {
        if (!project) return;
        // Don't trigger if clicking the select dropdown
        if (e.target.id === 'summaryFormat') return;
        const fmt = $('summaryFormat') ? $('summaryFormat').value : 'text';
        try {
            const ok = await Reports.copySummary(project, settings, fmt);
            if (ok) {
                toggleModal('modalReport', false);
                showToast('success', 'Summary copied to clipboard!');
            }
        } catch (err) {
            showToast('error', 'Copy failed: ' + err.message);
        }
    }

    function handleReportPrint() {
        if (!project) return;
        toggleModal('modalReport', false);
        setTimeout(() => {
            Reports.printProject();
        }, 200);
    }

    function handleReportDashboard() {
        if (!project) return;
        toggleModal('modalReport', false);

        const kpis = Dashboard.computeKPIs(project);
        const tasks = project.tasks.filter(t => !t.summary);
        const dur = project.startDate ? Math.ceil((project.finishDate - project.startDate) / 864e5) : 0;

        // Build task rows using concat (avoids nested template literal issues)
        var rowsHTML = '';
        tasks.forEach(function(t, i) {
            var pct = t.percentComplete || 0;
            var cls = pct >= 100 ? 'p100' : '';
            var status = pct >= 100 ? '\u2705' : pct > 0 ? '\ud83d\udfe2' : (t.critical ? '\ud83d\udd34' : '\u2b1c');
            var isLate = t.finish && new Date(t.finish) < new Date() && pct < 100;
            var lateAttr = isLate ? ' class="late"' : '';
            rowsHTML += '<tr><td>' + (i+1) + '</td><td' + lateAttr + '>' + t.name +
                '</td><td>' + (t.durationDays||0) + 'd</td><td>' + formatDate(t.start) +
                '</td><td>' + formatDate(t.finish) + '</td><td><div class="pbar ' + cls +
                '"><div class="pfill" style="width:' + pct + '%"></div></div> ' + pct +
                '%</td><td>' + status + '</td></tr>';
        });

        var critColor = (kpis && kpis.critical > 0) ? '#ef4444' : '#1a1d2e';
        var html = '<!DOCTYPE html><html><head><title>' + project.name + '</title>' +
            '<style>body{font-family:Inter,sans-serif;color:#1a1d2e;padding:24px;max-width:900px;margin:0 auto}' +
            'h1{font-size:1.3rem;border-bottom:3px solid #6366f1;padding-bottom:8px}' +
            '.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:16px 0}' +
            '.kpi{text-align:center;padding:14px 8px;border:1px solid #e2e4ec;border-radius:8px}' +
            '.kpi-label{font-size:.65rem;font-weight:600;text-transform:uppercase;color:#8690a7}' +
            '.kpi-val{font-size:1.4rem;font-weight:700;color:#1a1d2e}' +
            'table{width:100%;border-collapse:collapse;margin-top:16px;font-size:.72rem}' +
            'th{background:#f0f1f5;padding:6px 8px;text-align:left;font-weight:600;border-bottom:2px solid #d5d8e3}' +
            'td{padding:5px 8px;border-bottom:1px solid #e8e9f0}tr:nth-child(even) td{background:#f9fafb}' +
            '.pbar{width:50px;height:6px;background:#e8e9f0;border-radius:3px;display:inline-block}' +
            '.pfill{height:100%;border-radius:3px;background:#6366f1}.p100 .pfill{background:#22c55e}' +
            '.late{color:#ef4444;font-weight:600}.meta{font-size:.65rem;color:#8690a7;margin-top:4px}' +
            '</style></head><body>' +
            '<h1>📊 ' + escapeHTML(project.name) + '</h1>' +
            '<p class="meta">Generated: ' + new Date().toLocaleDateString() + ' | Tasks: ' + (kpis ? kpis.total : tasks.length) + ' | Duration: ' + dur + 'd</p>' +
            '<div class="kpi-grid">' +
            '<div class="kpi"><div class="kpi-label">Progress</div><div class="kpi-val">' + (kpis ? kpis.overallProgress : 0) + '%</div></div>' +
            '<div class="kpi"><div class="kpi-label">Tasks</div><div class="kpi-val">' + (kpis ? kpis.total : tasks.length) + '</div></div>' +
            '<div class="kpi"><div class="kpi-label">Critical</div><div class="kpi-val" style="color:' + critColor + '">' + (kpis ? kpis.critical : 0) + '</div></div>' +
            '<div class="kpi"><div class="kpi-label">Days Left</div><div class="kpi-val">' + (kpis ? kpis.daysRemaining : 0) + '</div></div>' +
            '</div>' +
            '<table><tr><th>#</th><th>Task Name</th><th>Duration</th><th>Start</th><th>Finish</th><th>% Done</th><th>Status</th></tr>' +
            rowsHTML + '</table></body></html>';

        var printWin = window.open('', '_blank', 'width=900,height=700');
        printWin.document.write(html);
        printWin.document.close();
        setTimeout(function() { printWin.print(); }, 500);
    }

    // ══════ CALENDAR SYSTEM ══════
    function initCalendar() {
        WorkCalendar.load();
        const now = new Date();
        calViewYear = now.getFullYear();
        calViewMonth = now.getMonth();
    }

    function populateCalendarModal() {
        const cfg = WorkCalendar.getConfig();
        $('calPreset').value = cfg.preset || 'western';
        $('customDaysGroup').style.display = cfg.preset === 'custom' ? '' : 'none';

        // Day checkboxes
        document.querySelectorAll('#dayCheckboxes input').forEach(cb => {
            cb.checked = cfg.workDays.includes(parseInt(cb.value));
        });

        // Country grid
        const grid = $('countryGrid'); grid.innerHTML = '';
        for (const c of WorkCalendar.getCountries()) {
            const chip = document.createElement('div'); chip.className = 'country-chip';
            if (cfg.countries.includes(c.code)) chip.classList.add('selected');
            chip.innerHTML = `<span class="flag">${escapeHTML(c.flag)}</span><span>${escapeHTML(c.code)}</span>`;
            chip.title = c.name;
            chip.addEventListener('click', () => chip.classList.toggle('selected'));
            grid.appendChild(chip);
        }

        // Custom holidays
        renderCustomHolidaysList();
    }

    function renderCustomHolidaysList() {
        const list = $('customHolidaysList'); list.innerHTML = '';
        const cfg = WorkCalendar.getConfig();
        for (const hol of cfg.customHolidays) {
            const div = document.createElement('div'); div.className = 'custom-hol-item';
            div.innerHTML = `<span>📌 ${escapeHTML(hol.date)}</span><span>${escapeHTML(hol.name)}</span><button class="btn-remove">&times;</button>`;
            div.querySelector('.btn-remove').addEventListener('click', () => {
                WorkCalendar.removeCustomHoliday(hol.date);
                renderCustomHolidaysList();
            });
            list.appendChild(div);
        }
    }

    function handleAddCustomHoliday() {
        const date = $('customHolDate').value;
        const name = $('customHolName').value.trim();
        if (!date || !name) { showToast('warning', 'Enter date and name'); return; }
        WorkCalendar.addCustomHoliday(date, name);
        $('customHolDate').value = ''; $('customHolName').value = '';
        renderCustomHolidaysList();
        showToast('info', `Added custom holiday: ${name}`);
    }

    async function handleFetchHolidays() {
        // Gather selected countries
        const selectedCountries = [];
        document.querySelectorAll('#countryGrid .country-chip.selected').forEach(chip => {
            selectedCountries.push(chip.querySelector('span:last-child').textContent);
        });
        if (selectedCountries.length === 0) { showToast('warning', 'Select at least one country'); return; }

        WorkCalendar.setCountries(selectedCountries);
        setStatus('Fetching holidays…');
        try {
            const year = calViewYear || new Date().getFullYear();
            await WorkCalendar.fetchHolidays(year);
            // Also fetch next year
            await WorkCalendar.fetchHolidays(year + 1);
            const cfg = WorkCalendar.getConfig();
            const count = Object.keys(cfg.holidays[year] || {}).length;
            showToast('success', `Loaded ${count} holidays for ${year}`);
        } catch (err) { showToast('error', 'Failed to fetch: ' + err.message); }
        setStatus('Ready');
    }

    function handleSaveCalendar() {
        const preset = $('calPreset').value;
        if (preset === 'custom') {
            const days = [];
            document.querySelectorAll('#dayCheckboxes input:checked').forEach(cb => days.push(parseInt(cb.value)));
            WorkCalendar.setWorkDays(days);
        } else {
            WorkCalendar.setPreset(preset);
        }

        // Save selected countries
        const selectedCountries = [];
        document.querySelectorAll('#countryGrid .country-chip.selected').forEach(chip => {
            selectedCountries.push(chip.querySelector('span:last-child').textContent);
        });
        WorkCalendar.setCountries(selectedCountries);
        WorkCalendar.save();

        toggleModal('modalCalendar', false);
        renderCalendarView();
        showToast('success', 'Calendar settings saved');
    }

    function renderCalendarView() {
        if (!els.calendarGrid) return;
        const cfg = WorkCalendar.getConfig();
        const presetName = WorkCalendar.getPresets()[cfg.preset]?.name || 'Custom';

        els.calMonthLabel.textContent = new Date(calViewYear, calViewMonth).toLocaleDateString('en-US', { month: 'long', year: 'numeric' });

        $('calPresetLabel').textContent = presetName;
        const holidays = cfg.holidays?.[calViewYear] || {};
        const holCount = Object.keys(holidays).length;
        $('calHolidayCount').textContent = holCount + ' holidays loaded';
        $('calCountryFlags').textContent = cfg.countries.map(c => {
            const found = WorkCalendar.getCountries().find(cc => cc.code === c);
            return found ? found.flag : '';
        }).join(' ');

        const grid = els.calendarGrid; grid.innerHTML = '';
        const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
        for (const dn of dayNames) {
            const h = document.createElement('div'); h.className = 'cal-day-header';
            h.textContent = dn; grid.appendChild(h);
        }

        const firstDay = new Date(calViewYear, calViewMonth, 1).getDay();
        for (let i = 0; i < firstDay; i++) {
            const e = document.createElement('div'); e.className = 'cal-day empty'; grid.appendChild(e);
        }

        const days = WorkCalendar.getMonthView(calViewYear, calViewMonth);
        const today = new Date(); today.setHours(0,0,0,0);

        for (const day of days) {
            const cell = document.createElement('div'); cell.className = 'cal-day';
            if (day.isWeekend) cell.classList.add('weekend');
            if (day.isHoliday) cell.classList.add('holiday');
            if (day.date.getTime() === today.getTime()) cell.classList.add('today');

            const num = document.createElement('div'); num.className = 'day-num';
            num.textContent = day.day; cell.appendChild(num);

            if (day.isHoliday && day.holidayInfo) {
                const lbl = document.createElement('div'); lbl.className = 'hol-label';
                lbl.textContent = day.holidayInfo.flag + ' ' + (day.holidayInfo.localName || day.holidayInfo.name);
                cell.appendChild(lbl);
            }

            // Phase 4.4: Show task pills on each day
            if (project && project.tasks) {
                const dayDate = day.date.getTime();
                const tasksOnDay = project.tasks.filter(t => {
                    if (t.summary || !t.isVisible) return false;
                    const ts = new Date(t.start); ts.setHours(0,0,0,0);
                    const tf = new Date(t.finish); tf.setHours(0,0,0,0);
                    return dayDate >= ts.getTime() && dayDate <= tf.getTime();
                });

                const maxPills = 3;
                for (let pi = 0; pi < Math.min(tasksOnDay.length, maxPills); pi++) {
                    const t = tasksOnDay[pi];
                    const pill = document.createElement('div');
                    pill.className = 'cal-task-pill';
                    if (t.critical) pill.classList.add('critical');
                    if (t.milestone) pill.classList.add('milestone');
                    pill.textContent = t.name.substring(0, 14) + (t.name.length > 14 ? '…' : '');
                    pill.style.backgroundColor = t.color || (t.critical ? '#ef4444' : '#6366f1');
                    pill.title = `${t.name} (${t.percentComplete || 0}%)`;

                    // Drag & Drop: make pill draggable
                    pill.draggable = true;
                    pill.addEventListener('dragstart', (e) => {
                        e.dataTransfer.setData('text/plain', String(t.uid));
                        e.dataTransfer.effectAllowed = 'move';
                        pill.classList.add('dragging');
                    });
                    pill.addEventListener('dragend', () => pill.classList.remove('dragging'));
                    cell.appendChild(pill);
                }
                if (tasksOnDay.length > maxPills) {
                    const more = document.createElement('div');
                    more.className = 'cal-task-more';
                    more.textContent = `+${tasksOnDay.length - maxPills} more`;
                    cell.appendChild(more);
                }
            }

            // Phase 4.4: Drop target for dragged task pills
            cell.addEventListener('dragover', (e) => { e.preventDefault(); e.dataTransfer.dropEffect = 'move'; cell.classList.add('drag-over'); });
            cell.addEventListener('dragleave', () => cell.classList.remove('drag-over'));
            cell.addEventListener('drop', (e) => {
                e.preventDefault();
                cell.classList.remove('drag-over');
                const uid = parseInt(e.dataTransfer.getData('text/plain'));
                if (!uid || !project) return;
                const task = project.tasks.find(t => t.uid === uid);
                if (!task || task.summary) return;

                saveUndoState();
                const newStart = new Date(day.date);
                const newFinish = addDays(newStart, task.durationDays || 1);
                task.start = newStart;
                task.finish = newFinish;
                recalculate(); renderAll(); autoSave();
                showToast('info', `Moved "${task.name}" to ${toDateInput(newStart)}`);
            });

            grid.appendChild(cell);
        }

        // Holidays list below calendar
        const holList = els.calendarHolidaysList; holList.innerHTML = '';
        const yearHols = Object.values(holidays).filter(h => {
            const m = parseInt(h.date.split('-')[1]) - 1;
            return m === calViewMonth;
        });
        if (yearHols.length > 0) {
            const title = document.createElement('h4'); title.textContent = 'Holidays this month: ' + yearHols.length;
            holList.appendChild(title);
            for (const h of yearHols) {
                const item = document.createElement('div'); item.className = 'hol-list-item';
                item.innerHTML = `<span class="hol-flag">${escapeHTML(h.flag || '')}</span><span class="hol-date">${escapeHTML(h.date)}</span><span class="hol-name">${escapeHTML(h.localName || h.name)}</span>`;
                holList.appendChild(item);
            }
        } else {
            holList.innerHTML = '<h4 style="color:var(--text-muted)">No holidays this month</h4>';
        }
    }

    // ══════ DASHBOARD ══════
    function renderDashboard() {
        if (!project) return;
        const kpis = Dashboard.computeKPIs(project);
        if (!kpis) return;

        // KPI Cards
        $('kpiProgress').textContent = kpis.overallProgress + '%';
        $('kpiProgressSub').textContent = `${kpis.complete}/${kpis.total} complete`;
        $('kpiTasks').textContent = kpis.total;
        $('kpiTasksSub').textContent = `${kpis.inProgress} active · ${kpis.late} late`;
        $('kpiCritical').textContent = kpis.critical;
        $('kpiCritical').style.color = kpis.critical > 0 ? '#ef4444' : '';
        $('kpiCriticalSub').textContent = kpis.critical > 0 ? 'tasks on critical path' : 'no critical tasks';
        $('kpiCost').textContent = settings.currency + kpis.totalCost.toLocaleString();
        $('kpiCostSub').textContent = `Earned: ${settings.currency}${Math.round(kpis.earnedCost).toLocaleString()}`;
        $('kpiDaysLeft').textContent = kpis.daysRemaining;
        $('kpiDaysLeft').style.color = kpis.daysRemaining < 7 ? '#ef4444' : '';
        $('kpiTimeSub').textContent = `${kpis.timeProgress}% time elapsed`;

        // Donut Chart
        const donutData = [
            { label: 'Complete', value: kpis.complete, color: '#22c55e' },
            { label: 'In Progress', value: kpis.inProgress, color: '#3b82f6' },
            { label: 'Not Started', value: kpis.notStarted, color: '#64748b' },
            { label: 'Late', value: kpis.late, color: '#ef4444' },
        ];
        Dashboard.drawDonut($('dashDonut'), donutData, {
            centerText: kpis.overallProgress + '%',
            centerSub: 'complete'
        });

        // Donut legend
        const legend = $('dashDonutLegend'); legend.innerHTML = '';
        donutData.forEach(d => {
            if (d.value > 0) {
                const item = document.createElement('span'); item.className = 'dash-legend-item';
                item.innerHTML = `<span class="dash-legend-dot" style="background:${escapeHTML(d.color)}"></span>${escapeHTML(d.label)}: ${escapeHTML(String(d.value))}`;
                legend.appendChild(item);
            }
        });

        // Phase Bars
        Dashboard.drawBars($('dashBars'), kpis.phases);

        // S-Curve
        Dashboard.drawSCurve($('dashSCurve'), project);

        // Timeline
        Dashboard.drawTimeline($('dashTimeline'), project);

        // Attention list
        const notifs = Dashboard.generateNotifications(project);
        updateNotifBadge(notifs.length);
        const attList = $('dashAttentionList'); attList.innerHTML = '';

        if (notifs.length === 0) {
            attList.innerHTML = '<div style="text-align:center;padding:16px;color:var(--text-muted)">✅ All tasks on track!</div>';
        } else {
            // Build a detailed table
            const tbl = document.createElement('table');
            tbl.className = 'attention-table';
            tbl.innerHTML = '<thead><tr><th></th><th>Task Name</th><th>Start</th><th>Finish</th><th>% Done</th><th>Note</th></tr></thead>';
            const tbody = document.createElement('tbody');

            notifs.slice(0, 12).forEach(n => {
                // Find task
                const t = project.tasks.find(t => t.uid === n.taskUid);
                const tr = document.createElement('tr');
                tr.className = 'att-row ' + n.type;
                const startStr = t ? formatDate(t.start) : '—';
                const finishStr = t ? formatDate(t.finish) : '—';
                const pct = t ? (t.percentComplete || 0) : 0;
                const taskName = t ? t.name : n.title;
                tr.innerHTML = '<td class="att-icon-cell">' + escapeHTML(n.icon) + '</td>' +
                    '<td class="att-name">' + escapeHTML(taskName) + '</td>' +
                    '<td class="att-date">' + escapeHTML(startStr) + '</td>' +
                    '<td class="att-date">' + escapeHTML(finishStr) + '</td>' +
                    '<td class="att-pct"><div class="progress-bar-mini"><div class="fill" style="width:' + pct + '%;background:' + (pct >= 100 ? '#22c55e' : n.type === 'danger' ? '#ef4444' : '#6366f1') + '"></div></div> ' + pct + '%</td>' +
                    '<td class="att-note">' + escapeHTML(n.subtitle) + '</td>';
                tbody.appendChild(tr);
            });

            tbl.appendChild(tbody);
            attList.appendChild(tbl);
        }

        // EVM
        renderEVM();
    }

    // ══════ NOTIFICATIONS PANEL ══════
    function toggleNotifPanel() {
        const panel = els.notifPanel;
        const wasHidden = panel.classList.contains('hidden');
        panel.classList.toggle('hidden');

        if (wasHidden) {
            const notifs = project ? Dashboard.generateNotifications(project) : [];
            renderNotifPanel(notifs);
        }
    }

    function renderNotifPanel(notifs) {
        const body = els.notifBody; body.innerHTML = '';
        if (notifs.length === 0) {
            body.innerHTML = '<div class="notif-empty">✅ No notifications</div>';
            return;
        }
        notifs.forEach((n, i) => {
            const item = document.createElement('div');
            item.className = 'notif-item ' + n.type;
            item.innerHTML = `<span class="nf-icon">${escapeHTML(n.icon)}</span><div class="nf-content"><div class="nf-title">${escapeHTML(n.title)}</div><div class="nf-sub">${escapeHTML(n.subtitle)}</div></div><button class="nf-dismiss">&times;</button>`;
            item.querySelector('.nf-dismiss').addEventListener('click', (e) => {
                e.stopPropagation();
                Dashboard.dismissNotification(i);
                item.remove();
                updateNotifBadge(Dashboard.getNotificationCount());
            });
            body.appendChild(item);
        });
    }

    function updateNotifBadge(count) {
        const badge = els.notifBadge;
        if (count > 0) {
            badge.textContent = count > 9 ? '9+' : count;
            badge.classList.remove('hidden');
        } else {
            badge.classList.add('hidden');
        }
    }

    // ══════════════════════════════════════════════
    // PHASE 3: EVM RENDERING IN DASHBOARD
    // ══════════════════════════════════════════════
    function renderEVM() {
        if (!project || typeof EVMEngine === 'undefined') return;
        const evm = EVMEngine.compute(project);
        if (!evm) return;

        const setVal = (id, val, health) => {
            const el = $(id);
            if (!el) return;
            el.textContent = val;
            el.className = 'evm-val' + (health ? ' ' + health : '');
        };

        setVal('evmSPI', EVMEngine.fmt(evm.SPI, 'index'), evm.scheduleHealth);
        setVal('evmCPI', EVMEngine.fmt(evm.CPI, 'index'), evm.costHealth);
        setVal('evmSV', EVMEngine.fmt(evm.SV, 'number'), evm.SV >= 0 ? 'good' : 'danger');
        setVal('evmCV', EVMEngine.fmt(evm.CV, 'number'), evm.CV >= 0 ? 'good' : 'danger');
        setVal('evmEAC', evm.hasCosts ? EVMEngine.fmt(evm.EAC, 'currency') : EVMEngine.fmt(evm.EAC, 'percent'));
        setVal('evmVAC', evm.hasCosts ? EVMEngine.fmt(evm.VAC, 'currency') : EVMEngine.fmt(evm.VAC, 'number'), evm.VAC >= 0 ? 'good' : 'danger');

        const evmCanvas = $('evmChart');
        if (evmCanvas && evm.timeData) EVMEngine.drawEVMChart(evmCanvas, evm.timeData);
    }

    // ══════════════════════════════════════════════
    // PHASE 4: THEME TOGGLE
    // ══════════════════════════════════════════════
    function toggleTheme() {
        const html = document.documentElement;
        const current = html.getAttribute('data-theme');
        const next = current === 'light' ? '' : 'light';
        if (next) html.setAttribute('data-theme', next); else html.removeAttribute('data-theme');
        $('themeIcon').textContent = next === 'light' ? '\u2600\ufe0f' : '\ud83c\udf19';
        localStorage.setItem('pf_theme', next || 'dark');
        // Re-render Gantt since Canvas colors are hardcoded
        if (project) setTimeout(() => { GanttChart.resize(); GanttChart.render(); }, 100);
    }

    // ══════════════════════════════════════════════
    // PHASE 4: RTL TOGGLE
    // ══════════════════════════════════════════════
    function toggleRTL() {
        const html = document.documentElement;
        const current = html.getAttribute('dir');
        const next = current === 'rtl' ? 'ltr' : 'rtl';
        html.setAttribute('dir', next);
        localStorage.setItem('pf_dir', next);
        showToast('info', next === 'rtl' ? 'RTL mode activated' : 'LTR mode activated');
    }

    // ══════════════════════════════════════════════
    // PHASE 4: KEYBOARD SHORTCUTS
    // ══════════════════════════════════════════════
    const SHORTCUTS = [
        { keys: ['\u2318', 'Z'], label: 'Undo', action: handleUndo },
        { keys: ['\u2318', '\u21e7', 'Z'], label: 'Redo', action: handleRedo },
        { keys: ['\u2318', 'S'], label: 'Save (auto-save)', action: () => autoSave() },
        { keys: ['Del'], label: 'Delete selected task(s)', action: () => $('btnDeleteTask')?.click() },
        { keys: ['T'], label: 'Switch to Table view', action: () => setView('table') },
        { keys: ['G'], label: 'Switch to Gantt view', action: () => setView('gantt') },
        { keys: ['S'], label: 'Switch to Split view', action: () => setView('split') },
        { keys: ['D'], label: 'Switch to Dashboard', action: () => setView('dashboard') },
        { keys: ['R'], label: 'Switch to Resources', action: () => setView('resources') },
        { keys: ['C'], label: 'Switch to Calendar', action: () => setView('calendar') },
        { keys: ['N'], label: 'Switch to Network', action: () => setView('network') },
        { keys: ['+'], label: 'Zoom In', action: () => { els.zoomLevel.textContent = GanttChart.zoomIn(); } },
        { keys: ['-'], label: 'Zoom Out', action: () => { els.zoomLevel.textContent = GanttChart.zoomOut(); } },
        { keys: ['?'], label: 'Show this panel', action: () => { populateShortcuts(); toggleModal('modalShortcuts', true); } },
        { keys: ['\u2318', 'P'], label: 'Print', action: () => window.print() },
    ];

    function populateShortcuts() {
        const list = $('shortcutList');
        if (!list) return;
        list.innerHTML = '';
        SHORTCUTS.forEach(sc => {
            const item = document.createElement('div');
            item.className = 'shortcut-item';
            item.innerHTML = `<span class="sc-label">${escapeHTML(sc.label)}</span><span class="sc-keys">${(sc.keys||[]).map(k => `<span class="sc-key">${escapeHTML(k)}</span>`).join('')}</span>`;
            list.appendChild(item);
        });
    }

    function filterShortcuts() {
        const q = ($('shortcutSearch')?.value || '').toLowerCase();
        document.querySelectorAll('.shortcut-item').forEach(item => {
            item.style.display = item.querySelector('.sc-label').textContent.toLowerCase().includes(q) ? '' : 'none';
        });
    }

    // ══════════════════════════════════════════════
    // PHASE 5: NETWORK VIEW
    // ══════════════════════════════════════════════
    let networkInitialized = false;
    function renderNetworkView() {
        if (!project) return;
        requestAnimationFrame(() => {
            if (!networkInitialized && els.networkCanvas) {
                NetworkDiagram.init(els.networkCanvas);
                networkInitialized = true;

                // Zoom / fit / export
                $('ndZoomIn') ?.addEventListener('click', () => NetworkDiagram.zoomIn());
                $('ndZoomOut')?.addEventListener('click', () => NetworkDiagram.zoomOut());
                $('ndFit')    ?.addEventListener('click', () => NetworkDiagram.fitToScreen());
                $('ndExport') ?.addEventListener('click', () => NetworkDiagram.exportPNG());

                // Mode buttons (Normal / Compact / Micro)
                document.querySelectorAll('.nd-mode-btn').forEach(btn => {
                    btn.addEventListener('click', () => {
                        document.querySelectorAll('.nd-mode-btn').forEach(b => b.classList.remove('active'));
                        btn.classList.add('active');
                        NetworkDiagram.setMode(btn.dataset.mode);
                    });
                });

                // Filter dropdown
                $('ndFilter')?.addEventListener('change', e => {
                    NetworkDiagram.setFilter(e.target.value);
                });

                // Search
                const ndSearch = $('ndSearch');
                if (ndSearch) {
                    ndSearch.addEventListener('input', debounce(e => {
                        NetworkDiagram.setSearch(e.target.value);
                    }, 200));
                    ndSearch.addEventListener('keydown', e => {
                        if (e.key === 'Escape') { ndSearch.value = ''; NetworkDiagram.setSearch(''); }
                    });
                }

                // Double-click → open detail panel
                els.networkCanvas.addEventListener('nodeDoubleClick', e => {
                    if (e.detail?.task) openDetailPanel(e.detail.task);
                });
            }

            if (typeof NetworkDiagram !== 'undefined') {
                NetworkDiagram.update(project.tasks);
                setTimeout(() => NetworkDiagram.fitToScreen(), 80);
            }
        });
    }

    // ══════════════════════════════════════════════
    // PHASE 5: TAGS SYSTEM
    // ══════════════════════════════════════════════
    const PREDEFINED_TAGS = [
        { name: 'Urgent', css: 'tag-urgent' },
        { name: 'Review', css: 'tag-review' },
        { name: 'Blocked', css: 'tag-blocked' },
        { name: 'Design', css: 'tag-design' },
        { name: 'Dev', css: 'tag-dev' },
        { name: 'Test', css: 'tag-test' },
    ];

    function renderTagsInDetailPanel(task) {
        let container = $('detailTags');
        if (!container) return;
        const tags = task.tags || [];

        // Current tags
        let html = '<div class="tags-cell" style="margin-bottom:6px">';
        tags.forEach(t => {
            const preset = PREDEFINED_TAGS.find(p => p.name === t);
            html += `<span class="tag ${preset ? preset.css : 'tag-custom'}">${t} <span style="cursor:pointer;margin-left:2px" data-remove-tag="${t}">\u00d7</span></span>`;
        });
        if (tags.length === 0) html += '<span style="font-size:0.72rem;color:var(--text-muted)">No tags</span>';
        html += '</div>';

        // Picker
        html += '<div class="tags-picker">';
        PREDEFINED_TAGS.forEach(pt => {
            const active = tags.includes(pt.name) ? 'active' : '';
            html += `<span class="tag ${pt.css} ${active}" data-toggle-tag="${pt.name}">${pt.name}</span>`;
        });
        html += '</div>';

        container.innerHTML = html;

        // Tag toggle events
        container.querySelectorAll('[data-toggle-tag]').forEach(el => {
            el.addEventListener('click', () => {
                saveUndoState();
                const tagName = el.dataset.toggleTag;
                if (!task.tags) task.tags = [];
                const idx = task.tags.indexOf(tagName);
                if (idx >= 0) task.tags.splice(idx, 1);
                else task.tags.push(tagName);
                renderTagsInDetailPanel(task);
                renderAll(); autoSave();
            });
        });
        container.querySelectorAll('[data-remove-tag]').forEach(el => {
            el.addEventListener('click', (e) => {
                e.stopPropagation();
                saveUndoState();
                const tagName = el.dataset.removeTag;
                if (task.tags) task.tags = task.tags.filter(t => t !== tagName);
                renderTagsInDetailPanel(task);
                renderAll(); autoSave();
            });
        });
    }

    // ══════════════════════════════════════════════
    // PHASE 5: COMMENTS SYSTEM
    // ══════════════════════════════════════════════
    function renderCommentsInDetailPanel(task) {
        let container = $('detailComments');
        if (!container) return;
        const comments = task.comments || [];

        let html = '<div class="comments-list">';
        if (comments.length === 0) {
            html += '<div style="font-size:0.72rem;color:var(--text-muted);padding:8px 0">No comments yet</div>';
        }
        comments.forEach((c, i) => {
            html += `<div class="comment-item"><div class="comment-meta">${c.date} <span style="cursor:pointer;float:right" data-del-comment="${i}">\u00d7</span></div><div class="comment-text">${escapeHTML(c.text)}</div></div>`;
        });
        html += '</div>';
        html += '<div class="comment-input-row"><input type="text" class="form-input" id="commentInput" placeholder="Add a comment..."><button class="btn btn-primary btn-xs" id="btnAddComment">Add</button></div>';

        container.innerHTML = html;

        $('btnAddComment')?.addEventListener('click', () => {
            const input = $('commentInput');
            const text = input?.value.trim();
            if (!text) return;
            if (!task.comments) task.comments = [];
            task.comments.push({ text, date: new Date().toLocaleString() });
            input.value = '';
            renderCommentsInDetailPanel(task);
            autoSave();
        });
        $('commentInput')?.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') $('btnAddComment')?.click();
        });
        container.querySelectorAll('[data-del-comment]').forEach(el => {
            el.addEventListener('click', () => {
                const idx = parseInt(el.dataset.delComment);
                if (task.comments) task.comments.splice(idx, 1);
                renderCommentsInDetailPanel(task);
                autoSave();
            });
        });
    }

    // ══════════════════════════════════════════════
    // PHASE 5: ATTACHMENTS SYSTEM
    // ══════════════════════════════════════════════
    function renderAttachmentsInDetailPanel(task) {
        let container = $('detailAttachments');
        if (!container) return;
        const attachments = task.attachments || [];

        let html = '<div class="attachment-list">';
        attachments.forEach((att, i) => {
            // Validate data URL — only allow data: protocol
            const safeData = (att.data && att.data.startsWith('data:')) ? att.data : '';
            if (att.type && att.type.startsWith('image/') && safeData) {
                html += `<div class="attachment-thumb"><img src="${safeData}" title="${escapeHTML(att.name)}"><button class="att-remove" data-del-att="${i}">\u00d7</button></div>`;
            } else {
                html += `<div class="attachment-thumb attachment-file-icon" title="${escapeHTML(att.name)}">📄<button class="att-remove" data-del-att="${i}">\u00d7</button></div>`;
            }
        });
        html += '</div>';
        html += '<label class="btn btn-secondary btn-xs" style="cursor:pointer">\ud83d\udcce Attach File<input type="file" id="attachFileInput" style="display:none" accept="image/*,.pdf,.doc,.docx,.txt"></label>';
        html += '<span style="font-size:0.62rem;color:var(--text-muted);margin-left:6px">Max 500KB per file</span>';

        container.innerHTML = html;

        $('attachFileInput')?.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (!file) return;
            if (file.size > 512000) { showToast('warning', 'File too large (max 500KB)'); return; }
            const reader = new FileReader();
            reader.onload = () => {
                if (!task.attachments) task.attachments = [];
                task.attachments.push({ name: file.name, type: file.type, data: reader.result });
                renderAttachmentsInDetailPanel(task);
                autoSave();
                showToast('info', `Attached "${file.name}"`);
            };
            reader.readAsDataURL(file);
        });
        container.querySelectorAll('[data-del-att]').forEach(el => {
            el.addEventListener('click', () => {
                const idx = parseInt(el.dataset.delAtt);
                if (task.attachments) task.attachments.splice(idx, 1);
                renderAttachmentsInDetailPanel(task);
                autoSave();
            });
        });
    }

    function escapeHTML(str) {
        const div = document.createElement('div');
        div.textContent = str;
        return div.innerHTML;
    }

    // ══════════════════════════════════════════════
    // TIER 1: HEALTH SCORE ENGINE
    // B.4: Delegates to ProjectAnalytics (eliminates duplicate calcs)
    // ══════════════════════════════════════════════
    function calculateHealthScore(projData) {
        if (!projData || !projData.tasks || !projData.tasks.length) return { score: 0, label: 'Unknown', icon: '⬜', class: 'not-started' };

        // Use cached analytics if available for the active project (fast path)
        if (projData === project) {
            const analytics = ProjectAnalytics.getCache();
            if (analytics && analytics.health) {
                const h = analytics.health;
                // Bonus: completed projects always show 100
                const completePct = analytics.total > 0 ? analytics.completeCount / analytics.total : 0;
                if (completePct >= 1) return { score: 100, label: 'Complete', icon: '✅', class: 'healthy' };
                return { score: h.score, label: h.label, icon: h.icon, class: h.cssClass };
            }
        }

        // Fallback: compute fresh for non-active projects (e.g. Hub cards)
        const tasks = projData.tasks.filter(t => !t.summary);
        if (!tasks.length) return { score: 0, label: 'Empty', icon: '⬜', class: 'not-started' };

        const total = tasks.length;
        const completePct = tasks.filter(t => t.percentComplete >= 100).length / total;
        let score = 100;

        const lateTasks = tasks.filter(t => t.percentComplete < 100 && new Date(t.finish) < new Date());
        score -= Math.min(30, lateTasks.length * 5);

        const critStalled = tasks.filter(t => t.critical && t.percentComplete < 50 && new Date(t.start) < new Date());
        score -= Math.min(20, critStalled.length * 3);

        if (projData.startDate && projData.finishDate) {
            const now = Date.now();
            const ps = new Date(projData.startDate).getTime();
            const pf = new Date(projData.finishDate).getTime();
            const duration = pf - ps;
            if (duration > 0 && now > ps) {
                const elapsed = Math.min(1, (now - ps) / duration);
                const avgPct = tasks.reduce((s, t) => s + (t.percentComplete || 0), 0) / total / 100;
                const spi = elapsed > 0 ? avgPct / elapsed : 1;
                if (spi < 0.7) score -= 20;
                else if (spi < 0.85) score -= 10;
                else if (spi < 0.95) score -= 5;
            }
        }

        if (completePct >= 0.9) score = Math.max(score, 90);
        if (completePct >= 1) score = 100;
        score = Math.max(0, Math.min(100, Math.round(score)));

        let label, icon, cls;
        if (score >= 75) { label = 'Healthy'; icon = '🟢'; cls = 'healthy'; }
        else if (score >= 45) { label = 'At Risk'; icon = '🟡'; cls = 'at-risk'; }
        else { label = 'Critical'; icon = '🔴'; cls = 'critical'; }
        if (completePct >= 1) { label = 'Complete'; icon = '✅'; cls = 'healthy'; score = 100; }

        return { score, label, icon, class: cls };
    }

    // Lightweight version using only index metadata
    function estimateHealth(meta) {
        const pct = meta.progress || 0;
        if (pct >= 100) return { score: 100, label: 'Complete', icon: '✅', class: 'healthy' };
        if (pct >= 75) return { score: 80, label: 'Healthy', icon: '🟢', class: 'healthy' }; // Aligned with calculateHealthScore
        if (pct >= 45) return { score: 55, label: 'At Risk', icon: '🟡', class: 'at-risk' }; // Aligned with calculateHealthScore
        if (pct > 0) return { score: 35, label: 'Critical', icon: '🔴', class: 'critical' };
        return { score: 0, label: 'Not Started', icon: '⬜', class: 'not-started' };
    }

    function getHealthRingSVG(score, color, size) {
        size = size || 36;
        const r = (size - 6) / 2;
        const circ = 2 * Math.PI * r;
        const offset = circ * (1 - score / 100);
        const bg = 'rgba(128,128,128,0.15)';
        return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}"><circle cx="${size/2}" cy="${size/2}" r="${r}" fill="none" stroke="${bg}" stroke-width="3"/><circle cx="${size/2}" cy="${size/2}" r="${r}" fill="none" stroke="${color}" stroke-width="3" stroke-dasharray="${circ}" stroke-dashoffset="${offset}" stroke-linecap="round" style="transition:stroke-dashoffset 0.6s ease"/></svg>`;
    }

    // ══════════════════════════════════════════════
    // TIER 1: COLOR PICKER
    // ══════════════════════════════════════════════
    function showColorPicker(element, currentColor, onChange) {
        // Remove existing popover
        document.querySelectorAll('.color-picker-popover').forEach(p => p.remove());

        const pop = document.createElement('div');
        pop.className = 'color-picker-popover';
        pop.innerHTML = '<div class="color-picker-grid">' +
            PROJECT_COLORS.map(c => `<div class="color-picker-swatch${c === currentColor ? ' active' : ''}" data-color="${c}" style="background:${c}"></div>`).join('') +
            '</div>';

        // Position near the element
        const rect = element.getBoundingClientRect();
        pop.style.position = 'fixed';
        pop.style.top = (rect.bottom + 4) + 'px';
        pop.style.left = rect.left + 'px';

        pop.addEventListener('click', (e) => {
            const swatch = e.target.closest('.color-picker-swatch');
            if (swatch) {
                const color = swatch.dataset.color;
                onChange(color);
                pop.remove();
            }
        });

        document.body.appendChild(pop);

        // Close on outside click
        const closeHandler = (e) => {
            if (!pop.contains(e.target) && e.target !== element) {
                pop.remove();
                document.removeEventListener('click', closeHandler);
            }
        };
        setTimeout(() => document.addEventListener('click', closeHandler), 10);
    }

    // ══════════════════════════════════════════════
    // TIER 1: INLINE RENAME
    // ══════════════════════════════════════════════
    function startInlineRename(nameEl, projectId) {
        const index = ProjectStore.getIndex();
        const meta = index.find(p => p.id === projectId);
        if (!meta) return;

        const currentName = meta.name;
        const input = document.createElement('input');
        input.className = 'inline-rename-input';
        input.value = currentName;
        input.setAttribute('maxlength', '60');

        nameEl.textContent = '';
        nameEl.appendChild(input);
        input.focus();
        input.select();

        const commit = () => {
            const newName = input.value.trim() || currentName;
            meta.name = newName;
            ProjectStore._syncIndexToLS();
            if (_dbReady) db.meta.put(meta).catch(() => {});
            // Also update project data name if it's the active project
            if (projectId === activeProjectId && project) {
                project.name = newName;
                els.projectNameDisplay.textContent = newName;
            }
            renderProjectHub();
            renderSidebar();
        };

        input.addEventListener('blur', commit, { once: true });
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') { e.preventDefault(); input.blur(); }
            if (e.key === 'Escape') { input.value = currentName; input.blur(); }
        });
    }

    // ══════════════════════════════════════════════
    // TIER 1: PIN FAVORITE
    // ══════════════════════════════════════════════
    function togglePin(projectId) {
        const index = ProjectStore.getIndex();
        const meta = index.find(p => p.id === projectId);
        if (!meta) return;
        meta.pinned = !meta.pinned;
        ProjectStore._syncIndexToLS();
        if (_dbReady) db.meta.put(meta).catch(() => {});
        renderProjectHub();
        renderSidebar();
    }

    // ══════════════════════════════════════════════
    // TIER 1: GLOBAL SEARCH
    // ══════════════════════════════════════════════
    function openGlobalSearch() {
        $('globalSearchOverlay').classList.remove('hidden');
        const input = $('globalSearchInput');
        input.value = '';
        $('globalSearchResults').innerHTML = '<div class="gs-empty">Type to search across all projects</div>';
        setTimeout(() => input.focus(), 100);
    }

    function closeGlobalSearch() {
        $('globalSearchOverlay').classList.add('hidden');
        $('globalSearchInput').value = '';
    }

    let _searchAbort = null;
    async function performGlobalSearch(query) {
        // Cancel previous search
        if (_searchAbort) _searchAbort.abort();
        _searchAbort = new AbortController();
        const signal = _searchAbort.signal;

        const results = $('globalSearchResults');
        if (!query || query.length < 2) {
            results.innerHTML = '<div class="gs-empty">Type at least 2 characters…</div>';
            return;
        }

        const q = query.toLowerCase();
        const index = ProjectStore.getIndex();
        let html = '';
        let totalResults = 0;

        for (const meta of index) {
            if (signal.aborted) return; // Cancelled — stop searching
            const proj = await ProjectStore.load(meta.id);
            if (signal.aborted) return;
            if (!proj || !proj.tasks) continue;

            const matches = proj.tasks.filter(t =>
                t.name.toLowerCase().includes(q) ||
                (t.notes || '').toLowerCase().includes(q) ||
                (t.resourceNames || []).join(' ').toLowerCase().includes(q)
            );

            if (matches.length > 0) {
                html += `<div class="gs-section-label"><span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${meta.color||'#6366f1'};margin-right:4px"></span>${escapeHTML(meta.name)} (${matches.length})</div>`;
                for (const t of matches.slice(0, 8)) {
                    const pct = t.percentComplete || 0;
                    html += `<div class="gs-result-item" data-proj="${meta.id}" data-task="${t.uid}">`;
                    html += `<span class="gs-result-name">${escapeHTML(t.name)}</span>`;
                    html += `<span class="gs-result-meta">${pct}% · ${t.durationDays || 0}d</span>`;
                    html += `</div>`;
                }
                if (matches.length > 8) {
                    html += `<div class="gs-result-meta" style="padding:4px 12px;font-style:italic">+${matches.length - 8} more results</div>`;
                }
                totalResults += matches.length;
            }
        }

        if (totalResults === 0) {
            results.innerHTML = '<div class="gs-empty">No matching tasks found</div>';
        } else {
            results.innerHTML = html;
            // Bind click handlers
            results.querySelectorAll('.gs-result-item').forEach(el => {
                el.addEventListener('click', () => {
                    const projId = el.dataset.proj;
                    const taskUid = parseInt(el.dataset.task);
                    closeGlobalSearch();
                    // Switch to the project and select the task
                    if (projId !== activeProjectId) {
                        switchProject(projId);
                    }
                    setTimeout(() => {
                        if (project) {
                            selectedTaskIds.clear();
                            selectedTaskIds.add(taskUid);
                            renderTable();
                            const task = project.tasks.find(t => t.uid === taskUid);
                            if (task) openDetailPanel(task);
                            showToast('info', `Found: "${task ? task.name : 'task'}"`);
                        }
                    }, 300);
                });
            });
        }
    }

    // ══════════════════════════════════════════════
    // PHASE 6a: MULTI-PROJECT SYSTEM
    // ══════════════════════════════════════════════

    function bindMultiProjectEvents() {
        // Sidebar toggle
        $('btnSidebarToggle').addEventListener('click', (e) => { e.stopPropagation(); toggleSidebar(); });
        $('btnSidebarClose').addEventListener('click', () => toggleSidebar(false));

        // Logo click → back to Hub
        const logoEl = document.querySelector('.logo');
        if (logoEl) {
            logoEl.style.cursor = 'pointer';
            logoEl.addEventListener('click', goToHub);
        }

        // Sidebar actions
        $('btnSidebarNew').addEventListener('click', () => { toggleSidebar(false); showNewProjectModal(); });
        $('btnSidebarImportPlanner').addEventListener('click', () => { toggleSidebar(false); openPlannerSyncModal(); });
        $('btnSidebarImportXML').addEventListener('click', () => { $('fileInputSidebar').click(); });
        $('fileInputSidebar').addEventListener('change', handleFileImport);
        $('btnSidebarPortfolio').addEventListener('click', () => { toggleSidebar(false); if (!project) { els.welcomeScreen.classList.add('hidden'); els.workspace.classList.remove('hidden'); } setView('portfolio'); });

        // Sidebar search
        if (els.sidebarSearch) {
            els.sidebarSearch.addEventListener('input', debounce(() => { renderSidebar(els.sidebarSearch.value.trim().toLowerCase()); }, 200));
        }

        // Hub buttons
        if ($('btnHubImportPlanner')) {
            $('btnHubImportPlanner').addEventListener('click', () => { openPlannerSyncModal(); });
        }

        // Portfolio import
        if ($('btnPortfolioImport')) {
            $('btnPortfolioImport').addEventListener('click', () => { $('filePlannerInputPortfolio').click(); });
            $('filePlannerInputPortfolio').addEventListener('change', handlePlannerFileForHub);
        }

        // Delete modal
        $('btnCloseDeleteProject').addEventListener('click', () => toggleModal('modalDeleteProject', false));
        $('btnCancelDelete').addEventListener('click', () => toggleModal('modalDeleteProject', false));
        $('btnConfirmDelete').addEventListener('click', async () => {
            if (deleteTargetId) {
                const wasActive = (deleteTargetId === activeProjectId);
                await ProjectStore.delete(deleteTargetId);
                if (wasActive) {
                    project = null;
                    activeProjectId = null;
                    els.workspace.classList.add('hidden');
                    els.welcomeScreen.classList.remove('hidden');
                }
                toggleModal('modalDeleteProject', false);
                renderProjectHub();
                renderSidebar();
                showToast('info', 'Project deleted');
                deleteTargetId = null;
            }
        });

        // Click outside sidebar to close
        document.addEventListener('click', (e) => {
            if (sidebarOpen && !els.projectSidebar.contains(e.target) && !e.target.closest('#btnSidebarToggle')) {
                toggleSidebar(false);
            }
        });

        // ── Tier 1: Hub Toolbar (Search / Sort / Filter) ──
        const hubSearchInput = $('hubSearchInput');
        const hubSortSelect = $('hubSortSelect');
        const hubFilterSelect = $('hubFilterSelect');
        if (hubSearchInput) hubSearchInput.addEventListener('input', debounce(() => renderProjectHub(), 200));
        if (hubSortSelect) hubSortSelect.addEventListener('change', () => renderProjectHub());
        if (hubFilterSelect) hubFilterSelect.addEventListener('change', () => renderProjectHub());

        // ── Tier 1: Global Search ──
        const gsOverlay = $('globalSearchOverlay');
        const gsInput = $('globalSearchInput');
        if (gsOverlay) {
            gsOverlay.addEventListener('click', (e) => { if (e.target === gsOverlay) closeGlobalSearch(); });
        }
        if (gsInput) {
            gsInput.addEventListener('input', debounce(() => performGlobalSearch(gsInput.value.trim()), 300));
        }
    }

    function handlePlannerFileForHub(e) {
        const file = e.target.files[0]; if (!file) return;
        setStatus('Parsing Planner Excel…');
        toggleSidebar(false);

        PlannerParser.parse(file)
            .then(data => {
                project = data;
                project.tasks.forEach(t => {
                    if (t.start && !(t.start instanceof Date)) t.start = new Date(t.start);
                    if (t.finish && !(t.finish instanceof Date)) t.finish = new Date(t.finish);
                    t.isExpanded = true; t.isVisible = true;
                    if (!t.predecessors) t.predecessors = [];
                    if (!t.resourceNames) t.resourceNames = [];
                });
                if (project.startDate && !(project.startDate instanceof Date)) project.startDate = new Date(project.startDate);
                if (project.finishDate && !(project.finishDate instanceof Date)) project.finishDate = new Date(project.finishDate);

                reindexTasks();
                activeProjectId = ProjectStore.generateId();
                onProjectLoaded();
                showToast('success', `Imported: "${project.name}" — ${project.tasks.length} items`);
            })
            .catch(err => {
                showToast('error', 'Import failed: ' + (err.message || 'Unknown error'));
            })
            .finally(() => {
                setStatus('Ready');
                e.target.value = '';
            });
    }

    function toggleSidebar(forceState) {
        if (forceState instanceof Event) forceState = undefined;
        sidebarOpen = typeof forceState === 'boolean' ? forceState : !sidebarOpen;
        els.projectSidebar.classList.toggle('open', sidebarOpen);
    }

    function renderSidebar(filter) {
        if (!els.sidebarProjectList) return;
        const index = ProjectStore.getIndex();
        const q = filter || '';
        let filtered = q ? index.filter(p => p.name.toLowerCase().includes(q)) : index.slice();
        // Hide archived from sidebar
        filtered = filtered.filter(p => !p.archived);

        els.sidebarProjectList.innerHTML = '';
        if (filtered.length === 0) {
            els.sidebarProjectList.innerHTML = '<div style="text-align:center;padding:20px;color:var(--text-muted);font-size:0.78rem">No projects yet</div>';
            return;
        }

        // Separate pinned and unpinned
        const pinned = filtered.filter(p => p.pinned);
        const unpinned = filtered.filter(p => !p.pinned);

        if (pinned.length > 0 && unpinned.length > 0) {
            const sep1 = document.createElement('div');
            sep1.className = 'sidebar-pin-separator';
            sep1.textContent = '⭐ Pinned';
            els.sidebarProjectList.appendChild(sep1);
        }

        const renderGroup = (group) => {
            group.forEach((p, i) => {
                const isActive = p.id === activeProjectId;
                const pct = p.progress || 0;
                const health = estimateHealth(p);
                const healthColor = health.class === 'healthy' ? '#22c55e' : health.class === 'at-risk' ? '#f59e0b' : health.class === 'critical' ? '#ef4444' : '#64748b';
                const barColor = pct >= 100 ? '#22c55e' : pct > 0 ? (p.color || '#6366f1') : '#64748b';

                const div = document.createElement('div');
                div.className = 'sidebar-project' + (isActive ? ' active' : '');
                div.style.cssText = `--proj-color: ${p.color || '#6366f1'}`;
                div.innerHTML = `
                    <div class="sidebar-project-info">
                        <div class="sidebar-project-name" style="display:flex;align-items:center;gap:4px">
                            <span style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escapeHTML(p.name)}</span>
                            <span class="health-badge ${health.class}" style="font-size:0.55rem;padding:1px 5px">${health.icon}</span>
                        </div>
                        <div class="sidebar-project-meta">${p.taskCount || 0} tasks · ${pct}% · ${health.label}</div>
                        <div class="sidebar-project-progress"><div class="sidebar-project-progress-fill" style="width:${pct}%;background:${barColor}"></div></div>
                    </div>
                    <div class="sidebar-project-actions">
                        <button title="${p.pinned ? 'Unpin' : 'Pin'}" class="pin-btn ${p.pinned ? 'pinned' : ''}" data-pin="${p.id}">${p.pinned ? '⭐' : '☆'}</button>
                        <button title="Duplicate" data-dup="${p.id}">📋</button>
                        <button title="Delete" data-del="${p.id}">🗑</button>
                    </div>
                `;

                div.addEventListener('click', (e) => {
                    if (e.target.closest('[data-dup]') || e.target.closest('[data-del]') || e.target.closest('[data-pin]')) return;
                    switchProject(p.id);
                    toggleSidebar(false);
                });

                const pinBtn = div.querySelector('[data-pin]');
                if (pinBtn) pinBtn.addEventListener('click', (e) => { e.stopPropagation(); togglePin(p.id); });
                const dupBtn = div.querySelector('[data-dup]');
                if (dupBtn) dupBtn.addEventListener('click', (e) => { e.stopPropagation(); duplicateProject(p.id); });
                const delBtn = div.querySelector('[data-del]');
                if (delBtn) delBtn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    deleteTargetId = p.id;
                    $('deleteProjectMsg').textContent = `Are you sure you want to delete "${p.name}"? This cannot be undone.`;
                    toggleModal('modalDeleteProject', true);
                });

                els.sidebarProjectList.appendChild(div);
            });
        };

        renderGroup(pinned);

        if (pinned.length > 0 && unpinned.length > 0) {
            const sep2 = document.createElement('div');
            sep2.className = 'sidebar-pin-separator';
            sep2.textContent = 'All Projects';
            els.sidebarProjectList.appendChild(sep2);
        }

        renderGroup(unpinned);
    }

    function renderProjectHub() {
        let index = ProjectStore.getIndex().slice();
        const grid = els.hubProjectGrid;
        const empty = els.hubEmptyState;
        const kpiBar = els.hubKpiBar;
        const toolbar = $('hubToolbar');

        if (!grid || !empty) return;

        if (index.length === 0) {
            grid.innerHTML = '';
            empty.classList.remove('hidden');
            if (kpiBar) kpiBar.classList.add('hidden');
            if (toolbar) toolbar.classList.add('hidden');
            return;
        }

        empty.classList.add('hidden');
        if (kpiBar) kpiBar.classList.remove('hidden');
        if (toolbar) toolbar.classList.remove('hidden');

        // KPIs
        const totalProj = index.length;
        const totalTasks = index.reduce((s, p) => s + (p.taskCount || 0), 0);
        const avgProgress = Math.round(index.reduce((s, p) => s + (p.progress || 0), 0) / totalProj);
        const healthyCount = index.filter(p => estimateHealth(p).class === 'healthy').length;
        const atRiskCount = index.filter(p => { const h = estimateHealth(p); return h.class === 'at-risk' || h.class === 'critical'; }).length;

        if ($('hubKpiProjects')) $('hubKpiProjects').textContent = totalProj;
        if ($('hubKpiAvgProgress')) $('hubKpiAvgProgress').textContent = avgProgress + '%';
        if ($('hubKpiTotalTasks')) $('hubKpiTotalTasks').textContent = totalTasks;
        if ($('hubKpiAtRisk')) $('hubKpiAtRisk').textContent = atRiskCount;

        // ── Apply Hub Search/Sort/Filter ──
        const hubSearch = ($('hubSearchInput') || {}).value || '';
        const hubSort = ($('hubSortSelect') || {}).value || 'recent';
        const hubFilter = ($('hubFilterSelect') || {}).value || 'all';

        // Filter by search
        if (hubSearch.trim()) {
            const q = hubSearch.trim().toLowerCase();
            index = index.filter(p => p.name.toLowerCase().includes(q));
        }

        // Filter by status
        if (hubFilter === 'archived') {
            index = index.filter(p => p.archived);
        } else {
            // Hide archived by default
            index = index.filter(p => !p.archived);
            if (hubFilter !== 'all') {
                index = index.filter(p => {
                    const h = estimateHealth(p);
                    switch (hubFilter) {
                        case 'on-track': return h.class === 'healthy' && p.progress < 100;
                        case 'at-risk': return h.class === 'at-risk';
                        case 'critical': return h.class === 'critical';
                        case 'complete': return (p.progress || 0) >= 100;
                        case 'not-started': return (p.progress || 0) === 0;
                        default: return true;
                    }
                });
            }
        }

        // Sort
        const pinned = index.filter(p => p.pinned);
        const unpinned = index.filter(p => !p.pinned);

        const sortFn = (arr) => {
            switch (hubSort) {
                case 'name': return arr.sort((a,b) => a.name.localeCompare(b.name));
                case 'name-desc': return arr.sort((a,b) => b.name.localeCompare(a.name));
                case 'progress': return arr.sort((a,b) => (a.progress||0) - (b.progress||0));
                case 'progress-desc': return arr.sort((a,b) => (b.progress||0) - (a.progress||0));
                case 'tasks': return arr.sort((a,b) => (a.taskCount||0) - (b.taskCount||0));
                case 'health': return arr.sort((a,b) => estimateHealth(a).score - estimateHealth(b).score);
                case 'recent': default: return arr.sort((a,b) => new Date(b.lastModified||0) - new Date(a.lastModified||0));
            }
        };

        sortFn(pinned);
        sortFn(unpinned);
        index = [...pinned, ...unpinned];

        // Cards
        grid.innerHTML = '';
        index.forEach(p => {
            const pct = p.progress || 0;
            const health = estimateHealth(p);
            const healthColor = health.class === 'healthy' ? '#22c55e' : health.class === 'at-risk' ? '#f59e0b' : health.class === 'critical' ? '#ef4444' : '#64748b';
            const barColor = pct >= 100 ? '#22c55e' : (p.color || '#6366f1');
            const startStr = p.startDate ? new Date(p.startDate).toLocaleDateString('en-US', {month:'short',year:'numeric'}) : '—';
            const finishStr = p.finishDate ? new Date(p.finishDate).toLocaleDateString('en-US', {month:'short',year:'numeric'}) : '—';

            const card = document.createElement('div');
            card.className = 'hub-project-card';
            card.style.cssText = `--card-color: ${p.color || '#6366f1'}`;
            card.innerHTML = `
                <div style="position:absolute;top:0;left:0;right:0;height:4px;background:${p.color || '#6366f1'}"></div>
                <div class="card-header">
                    <span class="card-name" data-rename="${p.id}">${p.pinned ? '⭐ ' : ''}📁 ${escapeHTML(p.name)}</span>
                    <div style="display:flex;gap:2px;align-items:center">
                        <button class="pin-btn ${p.pinned ? 'pinned' : ''}" data-pin="${p.id}" title="${p.pinned ? 'Unpin' : 'Pin to top'}">${p.pinned ? '⭐' : '☆'}</button>
                        <button class="pin-btn" data-color="${p.id}" title="Change color" style="font-size:0.7rem">🎨</button>
                    </div>
                </div>
                <div class="card-health">
                    <div class="health-ring">${getHealthRingSVG(health.score, healthColor)}<span class="health-score-text">${health.score}</span></div>
                    <span class="health-badge ${health.class}">${health.icon} ${health.label}</span>
                </div>
                <div class="card-progress">
                    <div class="card-progress-bar"><div class="card-progress-fill" style="width:${pct}%;background:${barColor}"></div></div>
                    <span class="card-progress-text" style="color:${barColor}">${pct}%</span>
                </div>
                <div class="card-stats">
                    <div class="card-stat">📋 <strong>${p.taskCount || 0}</strong> tasks</div>
                    <div class="card-stat">📅 ${startStr} — ${finishStr}</div>
                </div>
                <div class="card-description" data-desc="${p.id}" title="Click to edit description">${escapeHTML(p.description || 'Click to add notes…')}</div>
                <div class="card-footer">
                    <span class="card-status ${health.class === 'healthy' ? 'on-track' : health.class === 'at-risk' ? 'at-risk' : health.class === 'critical' ? 'behind' : 'not-started'}">${health.icon} ${health.label}</span>
                    <span class="card-date">Updated ${p.lastModified ? new Date(p.lastModified).toLocaleDateString() : '—'}</span>
                </div>
                <div class="card-actions-row">
                    <button class="btn btn-primary btn-xs" data-open="${p.id}">Open</button>
                    <button class="btn btn-ghost btn-xs" data-dup="${p.id}">📋 Clone</button>
                    <button class="btn btn-ghost btn-xs" data-archive="${p.id}" title="${p.archived ? 'Unarchive' : 'Archive'}">${p.archived ? '📤' : '📥'} ${p.archived ? 'Unarchive' : 'Archive'}</button>
                    <button class="btn btn-ghost btn-xs btn-danger" data-del="${p.id}">🗑 Delete</button>
                </div>
            `;

            // Pin
            card.querySelector('[data-pin]').addEventListener('click', (e) => { e.stopPropagation(); togglePin(p.id); });

            // Color picker
            card.querySelector('[data-color]').addEventListener('click', (e) => {
                e.stopPropagation();
                showColorPicker(e.target, p.color || '#6366f1', (newColor) => {
                    const meta = ProjectStore.getIndex().find(m => m.id === p.id);
                    if (meta) {
                        meta.color = newColor;
                        ProjectStore._syncIndexToLS();
                        if (_dbReady) db.meta.put(meta).catch(() => {});
                        renderProjectHub();
                        renderSidebar();
                    }
                });
            });

            // Inline rename (double-click on name)
            const nameEl = card.querySelector('[data-rename]');
            if (nameEl) {
                nameEl.addEventListener('dblclick', (e) => { e.stopPropagation(); startInlineRename(nameEl, p.id); });
            }

            // Open
            card.querySelector('[data-open]').addEventListener('click', (e) => { e.stopPropagation(); switchProject(p.id); });
            card.querySelector('[data-dup]').addEventListener('click', (e) => { e.stopPropagation(); duplicateProject(p.id); });
            card.querySelector('[data-del]').addEventListener('click', (e) => {
                e.stopPropagation();
                deleteTargetId = p.id;
                $('deleteProjectMsg').textContent = `Are you sure you want to delete "${p.name}"? This cannot be undone.`;
                toggleModal('modalDeleteProject', true);
            });

            // Double-click to open
            card.addEventListener('dblclick', () => switchProject(p.id));

            // Archive
            card.querySelector('[data-archive]').addEventListener('click', (e) => { e.stopPropagation(); archiveProject(p.id); });

            // Description inline edit
            const descEl = card.querySelector('[data-desc]');
            if (descEl) {
                descEl.addEventListener('click', (e) => {
                    e.stopPropagation();
                    const input = document.createElement('textarea');
                    input.className = 'inline-desc-input';
                    input.value = p.description || '';
                    input.placeholder = 'Add project notes…';
                    input.rows = 2;
                    descEl.textContent = '';
                    descEl.appendChild(input);
                    input.focus();

                    const commitDesc = () => {
                        const val = input.value.trim();
                        const meta = ProjectStore.getIndex().find(m => m.id === p.id);
                        if (meta) {
                            meta.description = val;
                            ProjectStore._syncIndexToLS();
                            if (_dbReady) db.meta.put(meta).catch(() => {});
                        }
                        renderProjectHub();
                    };
                    input.addEventListener('blur', commitDesc, { once: true });
                    input.addEventListener('keydown', (ev) => {
                        if (ev.key === 'Escape') { input.value = p.description || ''; input.blur(); }
                    });
                });
            }

            grid.appendChild(card);
        });
    }

    let _switchLock = false;
    async function switchProject(id) {
        if (_switchLock) return; // Prevent concurrent switches
        _switchLock = true;
        try {
            // Save current project first (await to ensure completion)
            if (project && activeProjectId) {
                await ProjectStore.save(activeProjectId, project);
                await ProjectStore.addToIndex(activeProjectId, project);
            }

            // Load the new project
            const loaded = await ProjectStore.load(id);
            if (!loaded) {
                showToast('error', 'Could not load project');
                return;
            }

            project = loaded;
            activeProjectId = id;
            reindexTasks();
            // B.4: Wipe analytics cache before loading new project
            ProjectAnalytics.reset();
            onProjectLoaded();
            setView('split');
            // B.3: Notify listeners of project switch
            EventBus.emit('project:switched', { id });
            showToast('success', `Switched to "${project.name}"`);
        } finally {
            _switchLock = false;
        }
    }

    async function duplicateProject(id) {
        const proj = await ProjectStore.load(id);
        if (!proj) { showToast('error', 'Source project not found'); return; }

        const newId = ProjectStore.generateId();
        proj.name = proj.name + ' (Copy)';
        // Reset progress
        (proj.tasks || []).forEach(t => { t.percentComplete = 0; });

        await ProjectStore.save(newId, proj);
        await ProjectStore.addToIndex(newId, proj);
        renderProjectHub();
        renderSidebar();
        showToast('success', `Duplicated as "${proj.name}"`);
    }

    // ══════ ARCHIVE PROJECT ══════
    function archiveProject(projectId) {
        const index = ProjectStore.getIndex();
        const meta = index.find(p => p.id === projectId);
        if (!meta) return;
        meta.archived = !meta.archived;
        ProjectStore._syncIndexToLS();
        if (_dbReady) db.meta.put(meta).catch(() => {});
        renderProjectHub();
        renderSidebar();
        showToast('info', meta.archived ? 'Project archived' : 'Project unarchived');
    }

    function goToHub() {
        // Save current & go back to hub
        if (project && activeProjectId) {
            ProjectStore.save(activeProjectId, project);
            ProjectStore.addToIndex(activeProjectId, project);
        }
        project = null;
        activeProjectId = null;
        els.workspace.classList.add('hidden');
        els.welcomeScreen.classList.remove('hidden');
        renderProjectHub();
        renderSidebar();
    }

    // ══════════════════════════════════════════════
    // PHASE 6a: PORTFOLIO VIEW  (redesigned)
    // ══════════════════════════════════════════════

    /** Main entry – called whenever portfolio tab is activated */
    function renderPortfolioView() {
        const index = ProjectStore.getIndex();
        _pfSetupCommandBar();
        _pfRenderKpis(index);
        _pfRenderHealthBar(index);
        // Dashboard sub-panels
        _pfRenderProgressBars(index);
        _pfRenderDonut(index);
        _pfRenderTimeline(index);
        _pfRenderAttention(index);
        _pfRenderHeatmap(index);
        // E.3: Bubble Chart
        setTimeout(() => renderPortfolioBubbleChart(index), 60);
        // Table view
        _pfRenderTable(index);
        // Canvas timeline (for export fallback)
        renderPortfolioTimeline(index);
    }

    /** Wire command-bar buttons (idempotent – uses a flag) */
    function _pfSetupCommandBar() {
        if ($('portfolioView') && $('portfolioView').dataset.cbInit) return;
        if ($('portfolioView')) $('portfolioView').dataset.cbInit = '1';

        // View-toggle buttons
        document.querySelectorAll('.pf-vtoggle').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.pf-vtoggle').forEach(b => b.classList.remove('active'));
                btn.classList.add('active');
                _pfSwitchView(btn.dataset.pfview);
            });
        });

        // Export PDF
        const btnExport = $('btnPortfolioExport');
        if (btnExport) btnExport.addEventListener('click', handlePortfolioExport);

        // Compare
        const btnCompare = $('btnCompareSelected');
        if (btnCompare) btnCompare.addEventListener('click', handleCompareMode);

        // Close compare panel
        const btnClose = $('btnCloseCompare');
        if (btnClose) btnClose.addEventListener('click', () => {
            const cr = $('compareResults');
            if (cr) cr.classList.add('hidden');
        });

        // Select-all checkbox
        const selectAll = $('pfSelectAll');
        if (selectAll) {
            selectAll.addEventListener('change', () => {
                document.querySelectorAll('.pf-compare-cb').forEach(cb => cb.checked = selectAll.checked);
            });
        }
    }

    /** Show one of: dashboard | table | timeline | resources */
    function _pfSwitchView(view) {
        const views = {
            dashboard: 'pfDashboardView',
            table:     'pfTableView',
            timeline:  'pfTimelineView',
            resources: 'pfResourceView',
        };
        Object.entries(views).forEach(([k, id]) => {
            const el = $(id);
            if (el) el.classList.toggle('hidden', k !== view);
        });
        if (view === 'timeline')  renderPortfolioTimeline(ProjectStore.getIndex());
        if (view === 'resources') _pfRenderResources();
    }

    // ──────────────────────────────────────────────
    // RESOURCE ALLOCATION ENGINE
    // ──────────────────────────────────────────────

    /** Entry point — triggers async data load then renders all panels */
    async function _pfRenderResources() {
        const index = ProjectStore.getIndex();
        const badge = $('pfResBadge');
        const tlWrap = $('pfResTimeline');
        const cardWrap = $('pfResCards');

        if (index.length === 0) {
            if (tlWrap) tlWrap.innerHTML = '<div class="pf-empty">No projects yet</div>';
            return;
        }
        if (badge) badge.textContent = 'Loading…';
        if (tlWrap) tlWrap.innerHTML = '<div class="pf-empty" style="padding:24px 0">Analysing resource assignments…</div>';
        if (cardWrap) cardWrap.innerHTML = '';

        const resMap = await _pfBuildResourceMap(index);
        if (badge) badge.textContent = resMap.sortedResources.length + ' resources';

        _pfRenderResKpis(resMap);
        _pfRenderResTimeline(resMap);
        _pfRenderResCards(resMap);
        _pfRenderConflicts(resMap);
    }

    /**
     * Load all projects from IndexedDB and build a rich resource map.
     * Returns: { resources, sortedResources, conflicts, allProjects,
     *            gStart, gEnd, today, totalPortfolioDays }
     */
    async function _pfBuildResourceMap(index) {
        const resources = {};   // key: resource name
        const allProjects = [];
        let gStart = Infinity, gEnd = -Infinity;
        const today = new Date(); today.setHours(0, 0, 0, 0);

        for (const meta of index) {
            const proj = await ProjectStore.load(meta.id);
            if (!proj) continue;
            allProjects.push({ meta, proj });

            if (proj.startDate)  gStart = Math.min(gStart, new Date(proj.startDate).getTime());
            if (proj.finishDate) gEnd   = Math.max(gEnd,   new Date(proj.finishDate).getTime());

            const leafTasks = (proj.tasks || []).filter(t => !t.summary);
            leafTasks.forEach(t => {
                // Normalise resourceNames to an array
                let rNames = t.resourceNames || [];
                if (typeof rNames === 'string')
                    rNames = rNames.split(',').map(s => s.trim()).filter(Boolean);
                if (rNames.length === 0) rNames = ['Unassigned'];

                const tStart = t.start  ? new Date(t.start)  : null;
                const tEnd   = t.finish ? new Date(t.finish) : null;
                if (!tStart || !tEnd || tEnd <= tStart) return;

                const pct       = t.percentComplete || 0;
                const durDays   = Math.max(1, Math.round((tEnd - tStart) / 86400000));
                const remaining = Math.round(durDays * (1 - pct / 100));
                const isDone    = pct >= 100;
                const isActive  = !isDone && tStart <= today && tEnd >= today;
                const isOverdue = !isDone && tEnd < today;
                const isUpcoming= !isDone && tStart > today;

                rNames.forEach(rName => {
                    if (!resources[rName]) {
                        resources[rName] = {
                            name: rName,
                            assignments: [],
                            totalTasks: 0, completedTasks: 0, activeTasks: 0,
                            remainingDays: 0, overdueTasks: 0,
                            projects: new Set(),
                            conflicts: [],
                        };
                    }
                    const r = resources[rName];
                    r.assignments.push({
                        taskName:  t.name || 'Untitled',
                        projName:  meta.name,
                        projColor: meta.color || '#6366f1',
                        projId:    meta.id,
                        start: tStart, end: tEnd,
                        pct, durDays, remaining,
                        isDone, isActive, isOverdue, isUpcoming,
                        critical: t.critical || false,
                    });
                    r.totalTasks++;
                    if (isDone)    r.completedTasks++;
                    if (isActive)  r.activeTasks++;
                    if (isOverdue) r.overdueTasks++;
                    if (!isDone)   r.remainingDays += remaining;
                    r.projects.add(meta.name);
                });
            });
        }

        if (!isFinite(gStart)) gStart = today.getTime() - 30 * 86400000;
        if (!isFinite(gEnd))   gEnd   = today.getTime() + 90 * 86400000;
        const totalPortfolioDays = Math.max(1, Math.round((gEnd - today.getTime()) / 86400000));

        // ── Conflict detection ──────────────────────────────
        // Two assignments conflict when they overlap in time, are both incomplete,
        // and belong to DIFFERENT projects.
        const allConflicts = [];
        Object.values(resources).forEach(r => {
            const pending = r.assignments.filter(a => !a.isDone);
            for (let i = 0; i < pending.length; i++) {
                for (let j = i + 1; j < pending.length; j++) {
                    const a = pending[i], b = pending[j];
                    if (a.projName === b.projName) continue;
                    const oStart = new Date(Math.max(a.start.getTime(), b.start.getTime()));
                    const oEnd   = new Date(Math.min(a.end.getTime(),   b.end.getTime()));
                    if (oStart >= oEnd) continue;          // no overlap
                    const overlapDays = Math.round((oEnd - oStart) / 86400000);
                    // Avoid duplicate pair entries
                    const already = r.conflicts.find(c =>
                        (c.proj1 === a.projName && c.proj2 === b.projName) ||
                        (c.proj1 === b.projName && c.proj2 === a.projName)
                    );
                    if (already) {
                        // Extend if this overlap is longer
                        if (overlapDays > already.overlapDays) {
                            already.overlapDays = overlapDays;
                            already.oStart = oStart; already.oEnd = oEnd;
                        }
                        continue;
                    }
                    const severity = overlapDays >= 14 ? 'high' : overlapDays >= 5 ? 'medium' : 'low';
                    const conflict = {
                        resource: r.name,
                        proj1: a.projName, proj1Color: a.projColor,
                        proj2: b.projName, proj2Color: b.projColor,
                        task1: a.taskName, task2: b.taskName,
                        oStart, oEnd, overlapDays, severity,
                    };
                    r.conflicts.push(conflict);
                    allConflicts.push(conflict);
                }
            }
            // Convert Set to sorted array
            r.projects = [...r.projects];
            // Utilisation: remaining days as % of remaining portfolio days
            r.utilizationPct = Math.min(200, Math.round((r.remainingDays / totalPortfolioDays) * 100));
        });

        allConflicts.sort((a, b) => b.overlapDays - a.overlapDays);

        // Sort resources: overloaded first → active → by remaining load
        const sortedResources = Object.values(resources).sort((a, b) => {
            if (b.conflicts.length !== a.conflicts.length) return b.conflicts.length - a.conflicts.length;
            if (b.activeTasks    !== a.activeTasks)    return b.activeTasks    - a.activeTasks;
            return b.remainingDays - a.remainingDays;
        });

        return { resources, sortedResources, conflicts: allConflicts, allProjects, gStart, gEnd, today, totalPortfolioDays };
    }

    /** Render the 5 KPI cards for the resource view */
    function _pfRenderResKpis(resMap) {
        const { sortedResources, conflicts } = resMap;
        const total       = sortedResources.length;
        const active      = sortedResources.filter(r => r.activeTasks > 0).length;
        const overloaded  = sortedResources.filter(r => r.conflicts.length > 0).length;
        const openTasks   = sortedResources.reduce((s, r) => s + (r.totalTasks - r.completedTasks), 0);
        const totalRemDays= sortedResources.reduce((s, r) => s + r.remainingDays, 0);

        const set = (id, val) => { if ($(id)) $(id).textContent = val; };
        set('pfResKpiTotal',     total);
        set('pfResKpiActive',    active);
        set('pfResKpiConflicts', overloaded > 0 ? overloaded + ' resource' + (overloaded > 1 ? 's' : '') : '✔ None');
        set('pfResKpiTasks',     openTasks);
        set('pfResKpiDays',      totalRemDays + 'd');
    }

    /** HTML allocation timeline — one row per resource */
    function _pfRenderResTimeline(resMap) {
        const wrap = $('pfResTimeline');
        if (!wrap) return;
        const { sortedResources, gStart, gEnd, today } = resMap;
        const visible = sortedResources.slice(0, 24);

        if (visible.length === 0) {
            wrap.innerHTML = '<div class="pf-empty">No resource assignments found. Add resource names to tasks to see the timeline.</div>';
            return;
        }

        const tStart = new Date(gStart), tEnd = new Date(gEnd);
        const range  = tEnd - tStart || 1;
        const toLeft  = dt => Math.min(100, Math.max(0, (dt - tStart) / range * 100));
        const toWidth = (s, e) => Math.max(0.3, toLeft(e) - toLeft(s));
        const todayPct = toLeft(today).toFixed(2);

        // Month tick marks
        const months = [];
        const md = new Date(tStart); md.setDate(1);
        for (let m = new Date(md); m <= tEnd; m.setMonth(m.getMonth() + 1)) {
            months.push({ left: toLeft(new Date(m)).toFixed(2), label: m.toLocaleDateString(undefined, { month: 'short', year: '2-digit' }) });
        }

        // Unique project→colour map for legend
        const projMap = {};
        sortedResources.forEach(r => r.assignments.forEach(a => { projMap[a.projName] = a.projColor; }));

        const rowsHTML = visible.map(r => {
            const statusColor = r.conflicts.length > 0 ? '#ef4444'
                              : r.activeTasks > 0      ? '#6366f1'
                              : r.remainingDays > 0    ? '#f59e0b'
                              : '#22c55e';
            const initials = r.name.split(' ').map(w => w[0] || '').join('').substring(0, 2).toUpperCase() || '?';

            const bars = r.assignments.map(a => {
                const l = toLeft(a.start).toFixed(2);
                const w = toWidth(a.start, a.end).toFixed(2);
                const opacity = a.isDone ? 0.3 : 1;
                const stripe  = a.isDone ? 'repeating-linear-gradient(45deg,transparent,transparent 3px,rgba(255,255,255,0.12) 3px,rgba(255,255,255,0.12) 6px)' : 'none';
                const border  = a.isOverdue ? '1.5px solid #ef4444' : a.critical ? '1.5px solid #f59e0b' : 'none';
                const title   = `${a.taskName} · ${a.projName} · ${a.pct}% · ${a.remaining}d remaining`;
                return `<div class="pf-res-tl-bar"
                    style="left:${l}%;width:${w}%;background:${a.projColor};opacity:${opacity};background-image:${stripe};border:${border}"
                    title="${escapeHTML(title)}"></div>`;
            }).join('');

            const conflictZones = r.conflicts.map(c => {
                const l = toLeft(c.oStart).toFixed(2);
                const w = toWidth(c.oStart, c.oEnd).toFixed(2);
                return `<div class="pf-res-tl-conflict-zone" style="left:${l}%;width:${w}%"
                    title="Conflict: ${escapeHTML(c.proj1)} ↔ ${escapeHTML(c.proj2)} (${c.overlapDays}d)"></div>`;
            }).join('');

            const monthGridLines = months.map(m =>
                `<div class="pf-res-tl-month-grid-line" style="left:${m.left}%"></div>`
            ).join('');

            return `
            <div class="pf-res-tl-row">
                <div class="pf-res-tl-row-label">
                    <div class="pf-res-avatar"
                        style="background:${statusColor}18;color:${statusColor};border-color:${statusColor}50">${initials}</div>
                    <div>
                        <div class="pf-res-name" title="${escapeHTML(r.name)}">${escapeHTML(r.name)}</div>
                        <div class="pf-res-meta">${r.projects.length} proj · ${r.activeTasks} active · ${r.remainingDays}d left</div>
                    </div>
                </div>
                <div class="pf-res-tl-bars">
                    ${monthGridLines}${bars}${conflictZones}
                    <div class="pf-res-tl-today-line" style="left:${todayPct}%"></div>
                </div>
            </div>`;
        }).join('');

        // Project colour legend (max 8)
        const legendHTML = Object.entries(projMap).slice(0, 8).map(([n, c]) =>
            `<span class="pf-res-legend-item"><span class="pf-res-leg-dot" style="background:${c}"></span>${escapeHTML(n)}</span>`
        ).join('');

        wrap.innerHTML = `
        <div class="pf-res-tl-header">
            <div class="pf-res-tl-lcol"></div>
            <div class="pf-res-tl-track" style="padding-top:16px;padding-bottom:4px">
                ${months.map(m => `<span class="pf-res-tl-month" style="left:${m.left}%">${m.label}</span>`).join('')}
            </div>
        </div>
        ${rowsHTML}
        <div style="margin-top:10px;display:flex;flex-wrap:wrap;gap:10px;padding-left:168px">
            ${legendHTML}
            <span class="pf-res-legend-item" style="color:#9ca3af">
                <span class="pf-res-leg-dot" style="background:repeating-linear-gradient(45deg,#6366f1,#6366f1 3px,transparent 3px,transparent 6px)"></span>Completed
            </span>
            <span class="pf-res-legend-item" style="color:#ef4444">
                <span class="pf-res-leg-dot" style="background:#ef4444;opacity:0.5"></span>Conflict zone
            </span>
        </div>`;
    }

    /** Grid of resource utilisation cards */
    function _pfRenderResCards(resMap) {
        const wrap = $('pfResCards');
        if (!wrap) return;
        const { sortedResources, today } = resMap;

        if (sortedResources.length === 0) {
            wrap.innerHTML = '<div class="pf-empty">No resource assignments found in tasks</div>';
            return;
        }

        wrap.innerHTML = sortedResources.map(r => {
            const pct      = r.utilizationPct;
            const barColor = r.conflicts.length > 0 ? '#ef4444'
                           : pct > 90             ? '#f59e0b'
                           : pct > 50             ? '#6366f1'
                           : '#22c55e';
            const badgeClass = r.conflicts.length > 0 ? 'overloaded'
                             : r.activeTasks > 2      ? 'busy'
                             : r.activeTasks > 0      ? 'active'
                             : 'free';
            const badgeLabel = r.conflicts.length > 0 ? '⚠ Conflict'
                             : r.activeTasks > 2      ? 'Busy'
                             : r.activeTasks > 0      ? 'Active'
                             : 'Free';
            const initials   = r.name.split(' ').map(w => w[0] || '').join('').substring(0, 2).toUpperCase() || '?';
            const cardClass  = r.conflicts.length > 0 ? 'pf-res-card is-overloaded' : 'pf-res-card';

            // Project colour tags
            const projTags = r.projects.slice(0, 4).map(projName => {
                const a = r.assignments.find(a => a.projName === projName);
                const c = a ? a.projColor : '#6366f1';
                return `<span class="pf-res-proj-tag" style="color:${c};border-color:${c}30;background:${c}10">
                    <span style="width:5px;height:5px;border-radius:50%;background:${c}"></span>${escapeHTML(projName.length > 16 ? projName.substring(0, 14) + '…' : projName)}
                </span>`;
            }).join('');

            // Active task list (max 4 shown)
            const activeTasks = r.assignments.filter(a => a.isActive || a.isOverdue).slice(0, 4);
            const taskListHTML = activeTasks.map(a => {
                const overdueTxt = a.isOverdue ? `<span class="pf-res-task-late">⚠ ${Math.round((today - a.end) / 86400000)}d late</span>` : `<span class="pf-res-task-pct">${a.pct}% · ${a.remaining}d left</span>`;
                return `<div class="pf-res-task-row">
                    <span class="pf-res-task-dot" style="background:${a.projColor}"></span>
                    <span class="pf-res-task-name" title="${escapeHTML(a.taskName)}">${escapeHTML(a.taskName.length > 30 ? a.taskName.substring(0, 28) + '…' : a.taskName)}</span>
                    ${overdueTxt}
                </div>`;
            }).join('');

            const moreTasksNote = r.assignments.filter(a => a.isActive || a.isOverdue).length > 4
                ? `<div style="font-size:0.6rem;color:var(--text-muted);padding-top:4px">+ ${r.assignments.filter(a => a.isActive || a.isOverdue).length - 4} more active tasks</div>` : '';

            return `
            <div class="${cardClass}">
                <div class="pf-res-card-header">
                    <div class="pf-res-avatar" style="width:36px;height:36px;font-size:0.78rem;background:${barColor}18;color:${barColor};border-color:${barColor}40">${initials}</div>
                    <div class="pf-res-card-info">
                        <div class="pf-res-card-title">${escapeHTML(r.name)}</div>
                        <div class="pf-res-card-sub">${r.projects.length} project${r.projects.length !== 1 ? 's' : ''} · ${r.totalTasks} task${r.totalTasks !== 1 ? 's' : ''}</div>
                    </div>
                    <span class="pf-res-badge ${badgeClass}">${badgeLabel}</span>
                </div>

                <div class="pf-res-util-row">
                    <div class="pf-res-util-bar">
                        <div class="pf-res-util-fill" style="width:${Math.min(100, pct)}%;background:${barColor}"></div>
                    </div>
                    <span class="pf-res-util-pct" style="color:${barColor}">${pct}%</span>
                </div>

                <div class="pf-res-proj-tags">${projTags}</div>

                ${activeTasks.length > 0 ? `<div style="border-top:1px solid rgba(255,255,255,0.05);padding-top:8px">${taskListHTML}${moreTasksNote}</div>` : '<div style="font-size:0.68rem;color:var(--text-muted);padding-top:4px">No active tasks right now</div>'}

                <div style="display:flex;gap:12px;margin-top:8px;padding-top:8px;border-top:1px solid rgba(255,255,255,0.05)">
                    <span style="font-size:0.63rem;color:var(--text-muted)">⏳ <b style="color:var(--text-secondary)">${r.remainingDays}d</b> remaining</span>
                    <span style="font-size:0.63rem;color:var(--text-muted)">✅ <b style="color:var(--text-secondary)">${r.completedTasks}/${r.totalTasks}</b> done</span>
                    ${r.conflicts.length > 0 ? `<span style="font-size:0.63rem;color:#ef4444">⚠ ${r.conflicts.length} conflict${r.conflicts.length > 1 ? 's' : ''}</span>` : ''}
                </div>
            </div>`;
        }).join('');
    }

    /** Conflict detail table */
    function _pfRenderConflicts(resMap) {
        const { conflicts } = resMap;
        const card = $('pfConflictCard');
        const list = $('pfConflictList');
        const badge = $('pfConflictBadge');

        if (!card || !list) return;

        if (conflicts.length === 0) {
            card.style.display = 'none';
            return;
        }

        card.style.display = '';
        if (badge) badge.textContent = conflicts.length + ' conflict' + (conflicts.length > 1 ? 's' : '');

        const fmtDate = d => d instanceof Date ? d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' }) : '—';

        list.innerHTML = conflicts.map(c => {
            const icon = c.severity === 'high' ? '🔴' : c.severity === 'medium' ? '🟡' : '🔵';
            return `
            <div class="pf-conflict-item ${c.severity}">
                <span class="pf-conflict-severity ${c.severity}">${icon} ${c.severity.toUpperCase()}</span>
                <div>
                    <div class="pf-conflict-body">
                        <b>${escapeHTML(c.resource)}</b> is double-allocated:
                        <span style="display:inline-flex;align-items:center;gap:4px;margin:0 4px">
                            <span style="width:7px;height:7px;border-radius:50%;background:${c.proj1Color};display:inline-block"></span>${escapeHTML(c.proj1)}
                        </span>
                        <b>↔</b>
                        <span style="display:inline-flex;align-items:center;gap:4px;margin:0 4px">
                            <span style="width:7px;height:7px;border-radius:50%;background:${c.proj2Color};display:inline-block"></span>${escapeHTML(c.proj2)}
                        </span>
                    </div>
                    <div class="pf-conflict-meta">
                        Tasks: "<em>${escapeHTML(c.task1)}</em>" &amp; "<em>${escapeHTML(c.task2)}</em>"
                        &nbsp;·&nbsp;Overlap: ${fmtDate(c.oStart)} – ${fmtDate(c.oEnd)}
                    </div>
                </div>
                <span class="pf-conflict-days">${c.overlapDays}d overlap</span>
            </div>`;
        }).join('');
    }

    /** Populate the 6 KPI cards */
    function _pfRenderKpis(index) {
        const n = index.length;
        const totalTasks   = index.reduce((s, p) => s + (p.taskCount || 0), 0);
        const avgProg      = n > 0 ? Math.round(index.reduce((s, p) => s + (p.progress || 0), 0) / n) : 0;
        const atRisk       = index.filter(p => { const h = estimateHealth(p); return h.label === 'At Risk'; }).length;
        const critical     = index.filter(p => { const h = estimateHealth(p); return h.label === 'Critical'; }).length;
        const overdue      = index.filter(p => p.finishDate && new Date(p.finishDate) < new Date() && (p.progress || 0) < 100).length;

        // SPI estimate: ratio completed tasks / expected completed tasks (simplified)
        const completedTasks = index.reduce((s, p) => s + Math.round((p.taskCount || 0) * (p.progress || 0) / 100), 0);
        const expectedTasks  = totalTasks > 0 ? Math.round(totalTasks * avgProg / 100) : 0;
        const spi            = expectedTasks > 0 ? (completedTasks / expectedTasks).toFixed(2) : '—';

        // Update project count label
        if ($('pfProjectCount')) $('pfProjectCount').textContent = n + (n === 1 ? ' Project' : ' Projects');

        const set = (id, val, trendId, trendTxt, trendClass) => {
            if ($(id)) $(id).textContent = val;
            if (trendId && $(trendId)) {
                $(trendId).textContent = trendTxt;
                $(trendId).className = 'pf-kpi-trend' + (trendClass ? ' ' + trendClass : '');
            }
        };

        set('pfKpiProjects',  n,           'pfKpiProjectsTrend', n === 1 ? '1 active' : n + ' active', '');
        set('pfKpiProgress',  avgProg + '%','pfKpiProgressTrend', avgProg >= 70 ? '▲ On Track' : avgProg >= 40 ? '▶ In Progress' : '▼ Needs attention',
            avgProg >= 70 ? 'up' : avgProg >= 40 ? 'warn' : 'down');
        set('pfKpiTasks',     totalTasks,  'pfKpiTasksTrend',    'across ' + n + ' project' + (n === 1 ? '' : 's'), '');
        set('pfKpiSpi',       spi,         'pfKpiSpiTrend',      spi === '—' ? '—' : spi >= 1 ? '▲ Ahead/On schedule' : '▼ Behind schedule',
            spi === '—' ? '' : spi >= 1 ? 'up' : 'down');
        set('pfKpiRisk',      atRisk + critical, 'pfKpiRiskTrend',
            (atRisk + critical) === 0 ? '✔ All healthy' : atRisk + ' at risk, ' + critical + ' critical',
            (atRisk + critical) === 0 ? 'up' : 'down');
        set('pfKpiOverdue',   overdue,     'pfKpiOverdueTrend',  overdue === 0 ? '✔ None overdue' : overdue + ' past deadline',
            overdue === 0 ? 'up' : 'down');
    }

    /** Render the coloured health-distribution bar */
    function _pfRenderHealthBar(index) {
        const n = index.length;
        if (!n) return;
        const counts = { complete: 0, healthy: 0, atRisk: 0, critical: 0 };
        index.forEach(p => {
            const h = estimateHealth(p);
            if ((p.progress || 0) >= 100)        counts.complete++;
            else if (h.label === 'Healthy')       counts.healthy++;
            else if (h.label === 'At Risk')       counts.atRisk++;
            else                                   counts.critical++;
        });
        const pct = v => (v / n * 100).toFixed(1) + '%';
        if ($('pfHealthComplete')) $('pfHealthComplete').style.width = pct(counts.complete);
        if ($('pfHealthHealthy'))  $('pfHealthHealthy').style.width  = pct(counts.healthy);
        if ($('pfHealthAtRisk'))   $('pfHealthAtRisk').style.width   = pct(counts.atRisk);
        if ($('pfHealthCritical')) $('pfHealthCritical').style.width = pct(counts.critical);

        // Legend
        const legend = $('pfHealthLegend');
        if (legend) {
            legend.innerHTML = [
                ['#6366f1', 'Complete',  counts.complete],
                ['#22c55e', 'Healthy',   counts.healthy],
                ['#f59e0b', 'At Risk',   counts.atRisk],
                ['#ef4444', 'Critical',  counts.critical],
            ].map(([c, l, v]) =>
                `<span style="display:inline-flex;align-items:center;gap:4px;margin-right:12px;font-size:0.72rem">
                    <span style="width:8px;height:8px;border-radius:50%;background:${c};flex-shrink:0"></span>${l} (${v})
                </span>`
            ).join('');
        }
    }

    /** HTML progress-bar comparison list */
    function _pfRenderProgressBars(index) {
        const wrap = $('pfProgressBars');
        if (!wrap) return;
        if (index.length === 0) { wrap.innerHTML = '<div class="pf-empty">No projects yet</div>'; return; }

        if ($('pfProgressBadge')) $('pfProgressBadge').textContent = index.length + ' projects';

        wrap.innerHTML = index.map(p => {
            const pct   = p.progress || 0;
            const color = p.color || '#6366f1';
            const h     = estimateHealth(p);
            const name  = escapeHTML(p.name.length > 22 ? p.name.substring(0, 20) + '…' : p.name);
            return `
            <div class="pf-progress-row" title="${escapeHTML(p.name)} — ${pct}%">
                <div class="pf-progress-label">
                    <span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${color};margin-right:5px;flex-shrink:0"></span>
                    <span>${name}</span>
                </div>
                <div class="pf-progress-track">
                    <div class="pf-progress-fill" style="width:${pct}%;background:${color}"></div>
                </div>
                <span class="pf-progress-pct">${pct}%</span>
                <span class="pf-health-badge ${h.label === 'Healthy' ? 'healthy' : h.label === 'At Risk' ? 'at-risk' : (p.progress||0) >= 100 ? 'complete' : 'critical'}">${h.label}</span>
            </div>`;
        }).join('');
    }

    /** Mini donut chart on canvas */
    function _pfRenderDonut(index) {
        const canvas = $('pfDonutCanvas');
        if (!canvas) return;

        const totalTasks = index.reduce((s, p) => s + (p.taskCount || 0), 0);
        if ($('pfTasksBadge')) $('pfTasksBadge').textContent = totalTasks + ' task' + (totalTasks !== 1 ? 's' : '');

        const completed  = index.reduce((s, p) => s + Math.round((p.taskCount || 0) * (p.progress || 0) / 100), 0);
        const inProgress = index.reduce((s, p) => {
            const done = Math.round((p.taskCount || 0) * (p.progress || 0) / 100);
            return s + Math.max(0, (p.taskCount || 0) - done);
        }, 0);
        const notStarted = Math.max(0, totalTasks - completed - inProgress);

        const slices = [
            { label: 'Done',        value: completed,  color: '#22c55e' },
            { label: 'In Progress', value: inProgress, color: '#6366f1' },
            { label: 'Not Started', value: notStarted, color: '#374151' },
        ].filter(s => s.value > 0);

        const size = 160;
        const dpr  = window.devicePixelRatio || 1;
        canvas.width  = size * dpr; canvas.height = size * dpr;
        canvas.style.width = size + 'px'; canvas.style.height = size + 'px';
        const ctx = canvas.getContext('2d');
        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        ctx.clearRect(0, 0, size, size);

        if (totalTasks === 0) {
            ctx.fillStyle = 'rgba(255,255,255,0.08)';
            ctx.beginPath(); ctx.arc(size/2, size/2, 60, 0, Math.PI*2); ctx.fill();
            return;
        }

        let angle = -Math.PI / 2;
        const cx = size / 2, cy = size / 2, r = 62, inner = 38;
        slices.forEach(s => {
            const sweep = (s.value / totalTasks) * Math.PI * 2;
            ctx.beginPath();
            ctx.moveTo(cx, cy);
            ctx.arc(cx, cy, r, angle, angle + sweep);
            ctx.closePath();
            ctx.fillStyle = s.color;
            ctx.fill();
            angle += sweep;
        });

        // Donut hole
        ctx.beginPath(); ctx.arc(cx, cy, inner, 0, Math.PI*2);
        ctx.fillStyle = getComputedStyle(document.documentElement).getPropertyValue('--bg-secondary').trim() || '#1e2030';
        ctx.fill();

        // Centre text
        const isDark = document.documentElement.getAttribute('data-theme') !== 'light';
        ctx.fillStyle = isDark ? '#e2e8f0' : '#1e293b';
        ctx.font = 'bold 20px Inter';
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.fillText(Math.round(completed / totalTasks * 100) + '%', cx, cy - 6);
        ctx.font = '10px Inter';
        ctx.fillStyle = isDark ? '#9aa0b4' : '#64748b';
        ctx.fillText('complete', cx, cy + 10);

        // Legend
        const legend = $('pfDonutLegend');
        if (legend) {
            legend.innerHTML = slices.map(s =>
                `<div style="display:flex;align-items:center;gap:6px;font-size:0.72rem;margin-bottom:4px">
                    <span style="width:10px;height:10px;border-radius:2px;background:${s.color};flex-shrink:0"></span>
                    <span>${escapeHTML(s.label)}</span>
                    <span style="margin-left:auto;font-weight:600">${s.value}</span>
                </div>`
            ).join('');
        }
    }

    /** HTML master timeline (inside dashboard card) */
    function _pfRenderTimeline(index) {
        const wrap = $('pfTimelineHtml');
        if (!wrap) return;
        if (index.length === 0) { wrap.innerHTML = '<div class="pf-empty">No projects with dates</div>'; return; }

        const dated = index.filter(p => p.startDate && p.finishDate);
        if (dated.length === 0) { wrap.innerHTML = '<div class="pf-empty">Add start/finish dates to see timeline</div>'; return; }

        let gStart = Infinity, gEnd = -Infinity;
        dated.forEach(p => {
            gStart = Math.min(gStart, new Date(p.startDate).getTime());
            gEnd   = Math.max(gEnd,   new Date(p.finishDate).getTime());
        });
        const range   = gEnd - gStart || 1;
        const today   = Date.now();
        const todayPct = Math.min(100, Math.max(0, ((today - gStart) / range) * 100));

        const fmtShort = d => new Date(d).toLocaleDateString(undefined, { month:'short', day:'numeric' });

        wrap.innerHTML = `
        <div class="pf-tl-header">
            <span style="font-size:0.7rem;color:var(--text-muted)">${fmtShort(gStart)}</span>
            <span style="font-size:0.7rem;color:var(--text-muted)">${fmtShort(gEnd)}</span>
        </div>
        <div style="position:relative">
            ${dated.map(p => {
                const s    = new Date(p.startDate).getTime();
                const e    = new Date(p.finishDate).getTime();
                const left = ((s - gStart) / range * 100).toFixed(1);
                const wid  = Math.max(1, ((e - s) / range * 100)).toFixed(1);
                const fill = ((p.progress || 0) / 100 * parseFloat(wid)).toFixed(1);
                const color = p.color || '#6366f1';
                const name  = escapeHTML(p.name.length > 18 ? p.name.substring(0, 16) + '…' : p.name);
                return `
                <div class="pf-tl-row">
                    <div class="pf-tl-label" title="${escapeHTML(p.name)}">${name}</div>
                    <div class="pf-tl-bar-bg">
                        <div style="position:absolute;left:${left}%;width:${wid}%;top:0;bottom:0;background:rgba(255,255,255,0.06);border-radius:3px"></div>
                        <div style="position:absolute;left:${left}%;width:${fill}%;top:0;bottom:0;background:${color};border-radius:3px;opacity:0.9"></div>
                    </div>
                </div>`;
            }).join('')}
            <div class="pf-tl-today" style="left:${todayPct.toFixed(1)}%" title="Today"></div>
        </div>`;
    }

    /** Canvas timeline used for PDF export (exported as public function) */
    function renderPortfolioTimeline(index) {
        const canvas = $('portfolioTimeline');
        if (!canvas) return;
        const dated = (index || []).filter(p => p.startDate && p.finishDate);
        if (dated.length === 0) return;

        const dpr = window.devicePixelRatio || 1;
        const w   = (canvas.parentElement ? canvas.parentElement.clientWidth : 0) || 700;
        const ROW = 28, PAD = { left: 130, right: 24, top: 28, bottom: 16 };
        const h   = PAD.top + dated.length * ROW + PAD.bottom;
        canvas.width  = w * dpr; canvas.height = h * dpr;
        canvas.style.width = w + 'px'; canvas.style.height = h + 'px';
        const ctx = canvas.getContext('2d');
        ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
        ctx.clearRect(0, 0, w, h);

        const isDark   = document.documentElement.getAttribute('data-theme') !== 'light';
        const textColor = isDark ? '#9aa0b4' : '#4a5068';
        const gridColor = isDark ? 'rgba(255,255,255,0.06)' : 'rgba(0,0,0,0.06)';

        let gStart = Infinity, gEnd = -Infinity;
        dated.forEach(p => {
            gStart = Math.min(gStart, new Date(p.startDate).getTime());
            gEnd   = Math.max(gEnd,   new Date(p.finishDate).getTime());
        });
        const range = gEnd - gStart || 1;
        const timeW = w - PAD.left - PAD.right;

        // Grid lines (monthly)
        const d0 = new Date(gStart); d0.setDate(1);
        for (let d = new Date(d0); d.getTime() <= gEnd; d.setMonth(d.getMonth() + 1)) {
            const x = PAD.left + ((d.getTime() - gStart) / range) * timeW;
            ctx.strokeStyle = gridColor; ctx.lineWidth = 1;
            ctx.beginPath(); ctx.moveTo(x, PAD.top - 10); ctx.lineTo(x, h - PAD.bottom); ctx.stroke();
            ctx.fillStyle = textColor; ctx.font = '9px Inter'; ctx.textAlign = 'center';
            ctx.fillText(d.toLocaleDateString(undefined, { month: 'short' }), x, PAD.top - 14);
        }

        // Today line
        const today = Date.now();
        if (today >= gStart && today <= gEnd) {
            const tx = PAD.left + ((today - gStart) / range) * timeW;
            ctx.strokeStyle = '#ef4444'; ctx.lineWidth = 1.5; ctx.setLineDash([4, 3]);
            ctx.beginPath(); ctx.moveTo(tx, PAD.top - 8); ctx.lineTo(tx, h - PAD.bottom); ctx.stroke();
            ctx.setLineDash([]);
            ctx.fillStyle = '#ef4444'; ctx.font = 'bold 9px Inter'; ctx.textAlign = 'center';
            ctx.fillText('Today', tx, PAD.top - 16);
        }

        // Rows
        dated.forEach((p, i) => {
            const y   = PAD.top + i * ROW;
            const barH = 16;
            const barY = y + (ROW - barH) / 2;
            const s   = new Date(p.startDate).getTime();
            const e   = new Date(p.finishDate).getTime();
            const x1  = PAD.left + ((s - gStart) / range) * timeW;
            const x2  = PAD.left + ((e - gStart) / range) * timeW;
            const bw  = Math.max(x2 - x1, 4);

            // Row stripe
            if (i % 2 === 0) {
                ctx.fillStyle = isDark ? 'rgba(255,255,255,0.02)' : 'rgba(0,0,0,0.02)';
                ctx.fillRect(0, y, w, ROW);
            }

            // Label
            ctx.fillStyle = textColor; ctx.font = '11px Inter';
            ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
            const label = p.name.length > 17 ? p.name.substring(0, 15) + '…' : p.name;
            ctx.fillText(label, PAD.left - 8, y + ROW / 2);

            // Bar background
            ctx.fillStyle = isDark ? 'rgba(255,255,255,0.07)' : 'rgba(0,0,0,0.06)';
            ctx.beginPath(); ctx.roundRect(x1, barY, bw, barH, 3); ctx.fill();

            // Progress fill
            const fillW = bw * (p.progress || 0) / 100;
            if (fillW > 0) {
                const grad = ctx.createLinearGradient(x1, 0, x1 + fillW, 0);
                grad.addColorStop(0, p.color || '#6366f1');
                grad.addColorStop(1, (p.color || '#6366f1') + 'bb');
                ctx.fillStyle = grad;
                ctx.beginPath(); ctx.roundRect(x1, barY, Math.max(fillW, 3), barH, 3); ctx.fill();
            }

            // Percentage label
            ctx.fillStyle = textColor; ctx.font = 'bold 9px Inter';
            ctx.textAlign = 'left'; ctx.textBaseline = 'middle';
            ctx.fillText((p.progress || 0) + '%', x1 + bw + 4, y + ROW / 2);
        });
    }

    /** Attention panel — overdue / critical projects */
    function _pfRenderAttention(index) {
        const wrap = $('pfAttentionList');
        if (!wrap) return;

        const today = new Date();
        const items = index.map(p => {
            const h        = estimateHealth(p);
            const isOverdue = p.finishDate && new Date(p.finishDate) < today && (p.progress || 0) < 100;
            const daysLeft  = p.finishDate ? Math.ceil((new Date(p.finishDate) - today) / 86400000) : null;
            return { p, h, isOverdue, daysLeft };
        }).filter(it => it.h.label === 'Critical' || it.h.label === 'At Risk' || it.isOverdue)
          .sort((a, b) => (a.daysLeft || 9999) - (b.daysLeft || 9999));

        if (items.length === 0) {
            wrap.innerHTML = '<div class="pf-empty" style="color:#22c55e">✔ All projects healthy</div>';
            return;
        }

        wrap.innerHTML = items.map(({ p, h, isOverdue, daysLeft }) => {
            const color  = p.color || '#6366f1';
            const name   = escapeHTML(p.name.length > 24 ? p.name.substring(0, 22) + '…' : p.name);
            const badge  = isOverdue ? 'late' : (daysLeft !== null && daysLeft <= 7 ? 'soon' : '');
            const detail = isOverdue
                ? `${Math.abs(daysLeft)} day${Math.abs(daysLeft) !== 1 ? 's' : ''} overdue`
                : daysLeft !== null
                    ? `${daysLeft} day${daysLeft !== 1 ? 's' : ''} left`
                    : h.label;
            return `
            <div class="pf-att-item">
                <span style="display:inline-block;width:8px;height:8px;border-radius:50%;background:${color};margin-right:6px;flex-shrink:0"></span>
                <span style="flex:1;min-width:0">${name}</span>
                <div class="pf-att-bar">
                    <div style="width:${p.progress||0}%;height:100%;background:${color};border-radius:2px"></div>
                </div>
                <span style="font-size:0.7rem;color:var(--text-muted);white-space:nowrap">${p.progress||0}%</span>
                ${badge ? `<span class="pf-att-badge ${badge}">${detail}</span>` : `<span style="font-size:0.7rem;color:var(--text-muted)">${detail}</span>`}
            </div>`;
        }).join('');
    }

    /** Resource workload heatmap (next 4 weeks, per-project row) */
    function _pfRenderHeatmap(index) {
        const wrap = $('pfHeatmap');
        if (!wrap) return;
        if (index.length === 0) { wrap.innerHTML = '<div class="pf-empty">No projects</div>'; return; }

        // Build a 4-week grid (Mon → Sun each week, 4 weeks)
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const monday = new Date(today);
        monday.setDate(today.getDate() - ((today.getDay() + 6) % 7)); // this Monday

        const WEEKS = 4, DAYS = 7;
        const days = [];
        for (let w = 0; w < WEEKS; w++) {
            for (let d = 0; d < DAYS; d++) {
                const dt = new Date(monday);
                dt.setDate(monday.getDate() + w * 7 + d);
                days.push(dt);
            }
        }

        // Score each (project × day): 0 = not active, 1-3 = low/med/high (based on progress pace needed)
        const dayLabels = ['Mo','Tu','We','Th','Fr','Sa','Su'];
        const weekLabels = days.filter((_, i) => i % 7 === 0).map(d =>
            d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' }));

        const rows = index.map(p => {
            const s = p.startDate ? new Date(p.startDate) : null;
            const e = p.finishDate ? new Date(p.finishDate) : null;
            const cells = days.map(day => {
                if (!s || !e) return 0;
                if (day < s || day > e) return 0;
                const remaining = (100 - (p.progress || 0)) / 100;
                if (remaining < 0.1) return 1;
                if (remaining < 0.4) return 2;
                return 3;
            });
            return { p, cells };
        });

        const intensityColor = (v) => {
            if (v === 0) return 'rgba(255,255,255,0.04)';
            if (v === 1) return '#22c55e44';
            if (v === 2) return '#f59e0b66';
            return '#ef444488';
        };

        wrap.innerHTML = `
        <div class="pf-hm-row" style="padding-left:110px;margin-bottom:2px">
            ${weekLabels.map(l => `<span style="flex:7;text-align:left;font-size:0.65rem;color:var(--text-muted)">${l}</span>`).join('')}
        </div>
        ${rows.map(({ p, cells }) => {
            const name = escapeHTML(p.name.length > 14 ? p.name.substring(0, 12) + '…' : p.name);
            return `
            <div class="pf-hm-row">
                <span class="pf-hm-label" title="${escapeHTML(p.name)}">${name}</span>
                ${cells.map(v => `<div class="pf-hm-cell" style="background:${intensityColor(v)}" title="${v === 0 ? 'Inactive' : v === 1 ? 'Low load' : v === 2 ? 'Medium load' : 'High load'}"></div>`).join('')}
            </div>`;
        }).join('')}
        <div style="display:flex;gap:6px;align-items:center;margin-top:8px;font-size:0.65rem;color:var(--text-muted)">
            <span>Load:</span>
            <div style="width:10px;height:10px;border-radius:2px;background:rgba(255,255,255,0.04);border:1px solid rgba(255,255,255,0.1)"></div><span>None</span>
            <div style="width:10px;height:10px;border-radius:2px;background:#22c55e44"></div><span>Low</span>
            <div style="width:10px;height:10px;border-radius:2px;background:#f59e0b66"></div><span>Med</span>
            <div style="width:10px;height:10px;border-radius:2px;background:#ef444488"></div><span>High</span>
        </div>`;
    }

    /** Table view — rich project table */
    function _pfRenderTable(index) {
        const tbody = $('portfolioTableBody');
        if (!tbody) return;
        tbody.innerHTML = '';

        index.forEach(p => {
            const pct      = p.progress || 0;
            const h        = estimateHealth(p);
            const startStr = p.startDate  ? new Date(p.startDate).toLocaleDateString()  : '—';
            const endStr   = p.finishDate ? new Date(p.finishDate).toLocaleDateString() : '—';
            const budget   = (p.tasks || []).reduce((s, t) => s + (t.cost || 0), 0);
            const budgetStr = budget > 0 ? (settings.currency || '$') + budget.toLocaleString(undefined, { maximumFractionDigits: 0 }) : '—';
            const healthCls = h.label === 'Healthy' ? 'healthy' : h.label === 'At Risk' ? 'at-risk' : pct >= 100 ? 'complete' : 'critical';
            // Simple SPI per project index (no full task load here)
            const spiRaw    = p.spi || null;
            const spiTxt    = spiRaw ? spiRaw.toFixed(2) : '—';
            const spiCls    = !spiRaw ? '' : spiRaw >= 1 ? 'pf-spi-good' : spiRaw >= 0.8 ? 'pf-spi-warn' : 'pf-spi-bad';

            const tr = document.createElement('tr');
            tr.style.cursor = 'pointer';
            tr.innerHTML = `
                <td><input type="checkbox" class="pf-compare-cb" data-id="${p.id}"></td>
                <td class="pf-tbl-proj-cell">
                    <span class="pf-tbl-dot" style="background:${p.color||'#6366f1'}"></span>
                    <span class="pf-tbl-name">${escapeHTML(p.name)}</span>
                </td>
                <td>${p.taskCount || 0}</td>
                <td>
                    <div style="display:flex;align-items:center;gap:6px">
                        <div style="width:56px;height:4px;background:var(--bg-active,#2a2d3e);border-radius:2px;overflow:hidden;flex-shrink:0">
                            <div style="width:${pct}%;height:100%;background:${p.color||'#6366f1'};border-radius:2px"></div>
                        </div>
                        <span style="font-size:0.72rem;white-space:nowrap">${pct}%</span>
                    </div>
                </td>
                <td class="${spiCls}">${spiTxt}</td>
                <td><span class="pf-health-badge ${healthCls}">${h.label}</span></td>
                <td style="font-size:0.74rem;white-space:nowrap">${startStr}</td>
                <td style="font-size:0.74rem;white-space:nowrap">${endStr}</td>
                <td style="font-size:0.74rem">${budgetStr}</td>
            `;
            tr.querySelector('.pf-compare-cb').addEventListener('click', e => e.stopPropagation());
            tr.addEventListener('dblclick', () => switchProject(p.id));
            tbody.appendChild(tr);
        });
    }

    // ══════════════════════════════════════════════
    // COMPARE MODE (Tier 2)
    // ══════════════════════════════════════════════
    async function handleCompareMode() {
        const selectedIds = [];
        document.querySelectorAll('.pf-compare-cb:checked').forEach(cb => selectedIds.push(cb.dataset.id));

        if (selectedIds.length < 2) {
            showToast('warning', 'Select at least 2 projects to compare');
            return;
        }
        if (selectedIds.length > 5) {
            showToast('warning', 'Select up to 5 projects for comparison');
            return;
        }

        const container = $('compareResults');
        const content = $('compareContent');
        if (!container || !content) return;

        container.classList.remove('hidden');
        content.innerHTML = '<div style="padding:20px;text-align:center;color:var(--text-muted)">Loading comparison…</div>';

        const projects = [];
        for (const id of selectedIds) {
            const meta = ProjectStore.getIndex().find(m => m.id === id);
            const data = await ProjectStore.load(id);
            if (meta && data) projects.push({ meta, data });
        }

        // Build comparison table
        let html = '<table class="data-table compare-table"><thead><tr><th>Metric</th>';
        projects.forEach(p => {
            html += `<th style="border-bottom:3px solid ${p.meta.color||'#6366f1'}">${escapeHTML(p.meta.name)}</th>`;
        });
        html += '</tr></thead><tbody>';

        const metrics = [
            { label: 'Total Tasks', fn: p => (p.data.tasks || []).length },
            { label: 'Progress', fn: p => (p.meta.progress || 0) + '%' },
            { label: 'Leaf Tasks', fn: p => (p.data.tasks || []).filter(t => !t.summary).length },
            { label: 'Completed', fn: p => (p.data.tasks || []).filter(t => t.percentComplete >= 100).length },
            { label: 'Critical Tasks', fn: p => (p.data.tasks || []).filter(t => t.critical).length },
            { label: 'Milestones', fn: p => (p.data.tasks || []).filter(t => t.milestone).length },
            { label: 'Resources', fn: p => (p.data.resources || []).length },
            { label: 'Start Date', fn: p => p.data.startDate ? new Date(p.data.startDate).toLocaleDateString() : '—' },
            { label: 'Finish Date', fn: p => p.data.finishDate ? new Date(p.data.finishDate).toLocaleDateString() : '—' },
            { label: 'Duration (days)', fn: p => {
                if (!p.data.startDate || !p.data.finishDate) return '—';
                return Math.round((new Date(p.data.finishDate) - new Date(p.data.startDate)) / 86400000);
            }},
            { label: 'Total Cost', fn: p => settings.currency + ((p.data.tasks || []).reduce((s,t) => s + (t.cost || 0), 0)).toFixed(0) },
            { label: 'Health Score', fn: p => estimateHealth(p.meta).score + '/100' },
        ];

        for (const m of metrics) {
            html += '<tr>';
            html += `<td style="font-weight:600">${m.label}</td>`;
            projects.forEach(p => {
                html += `<td>${m.fn(p)}</td>`;
            });
            html += '</tr>';
        }
        html += '</tbody></table>';

        content.innerHTML = html;
        content.scrollIntoView({ behavior: 'smooth' });
        showToast('success', `Comparing ${projects.length} projects`);
    }

    // ══════════════════════════════════════════════
    // PORTFOLIO EXPORT PDF (Tier 2)
    // ══════════════════════════════════════════════
    async function handlePortfolioExport() {
        if (typeof Reports === 'undefined' || !Reports.generatePortfolioPDF) {
            showToast('error', 'Reports module not loaded');
            return;
        }

        const index = ProjectStore.getIndex().filter(p => !p.archived);
        if (index.length === 0) {
            showToast('warning', 'No projects to export');
            return;
        }

        setStatus('Generating portfolio report…');
        showToast('info', 'Generating Portfolio PDF…');

        try {
            // Load full project data for each project
            const projects = [];
            for (const meta of index) {
                const proj = await ProjectStore.load(meta.id);
                if (proj) projects.push(proj);
            }

            await Reports.generatePortfolioPDF(projects, settings);
            setStatus('Ready');
            showToast('success', 'Portfolio PDF generated!');
        } catch (e) {
            console.error('Portfolio export failed:', e);
            showToast('error', 'Export failed: ' + e.message);
            setStatus('Ready');
        }
    }

    // ══════════════════════════════════════════════════════════
    // SPRINT C & D — NEW FEATURES
    // ══════════════════════════════════════════════════════════

    // ── C.2: Advanced Filter State ────────────────────────────
    let _advFilter = {
        tags: [],
        resource: '',
        startAfter: null,
        finishBefore: null,
        pctMin: null,
        pctMax: null
    };
    let _filterPresets = [];

    function _advFilterActive() {
        return _advFilter.tags.length > 0 || _advFilter.resource ||
            _advFilter.startAfter || _advFilter.finishBefore ||
            _advFilter.pctMin !== null || _advFilter.pctMax !== null;
    }

    function _applyAdvFilter(tasks) {
        if (!_advFilterActive()) return tasks;
        return tasks.filter(t => {
            if (_advFilter.tags.length && !_advFilter.tags.some(tag => (t.tags || []).includes(tag))) return false;
            if (_advFilter.resource && !(t.resourceNames || []).some(r => r.toLowerCase().includes(_advFilter.resource.toLowerCase()))) return false;
            if (_advFilter.startAfter && new Date(t.start) < _advFilter.startAfter) return false;
            if (_advFilter.finishBefore && new Date(t.finish) > _advFilter.finishBefore) return false;
            if (_advFilter.pctMin !== null && (t.percentComplete || 0) < _advFilter.pctMin) return false;
            if (_advFilter.pctMax !== null && (t.percentComplete || 0) > _advFilter.pctMax) return false;
            return true;
        });
    }

    function _loadFilterPresets() {
        try { _filterPresets = JSON.parse(localStorage.getItem('pf_filter_presets') || '[]'); } catch(_) { _filterPresets = []; }
    }

    function _saveFilterPresets() {
        try { localStorage.setItem('pf_filter_presets', JSON.stringify(_filterPresets)); } catch(_) {}
    }

    function _updateAdvFilterBadge() {
        const badge = $('advFilterBadge');
        if (!badge) return;
        const count = _advFilter.tags.length +
            (_advFilter.resource ? 1 : 0) +
            (_advFilter.startAfter ? 1 : 0) +
            (_advFilter.finishBefore ? 1 : 0) +
            (_advFilter.pctMin !== null || _advFilter.pctMax !== null ? 1 : 0);
        badge.textContent = count;
        badge.classList.toggle('hidden', count === 0);
        const btn = $('btnAdvFilter');
        if (btn) btn.classList.toggle('active', count > 0);
    }

    function initAdvancedFilter() {
        _loadFilterPresets();

        const btn = $('btnAdvFilter');
        const panel = $('advFilterPanel');
        if (!btn || !panel) return;

        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            const rect = btn.getBoundingClientRect();
            panel.style.top  = (rect.bottom + 6) + 'px';
            panel.style.left = Math.min(rect.left, window.innerWidth - 320) + 'px';
            panel.classList.toggle('hidden');
            if (!panel.classList.contains('hidden')) _populateAdvFilterPanel();
        });

        document.addEventListener('click', (e) => {
            if (!panel.contains(e.target) && e.target !== btn) panel.classList.add('hidden');
        });

        $('btnCloseAdvFilter').addEventListener('click', () => panel.classList.add('hidden'));

        $('btnClearFilters').addEventListener('click', () => {
            _advFilter = { tags: [], resource: '', startAfter: null, finishBefore: null, pctMin: null, pctMax: null };
            _populateAdvFilterPanel();
            _updateAdvFilterBadge();
            renderTable();
        });

        $('btnApplyAdvFilter').addEventListener('click', () => {
            // Read panel values
            _advFilter.resource    = ($('advFilterResource').value || '').trim();
            const sa = $('advFilterStartAfter').value;   _advFilter.startAfter   = sa ? new Date(sa) : null;
            const fb = $('advFilterFinishBefore').value; _advFilter.finishBefore = fb ? new Date(fb) : null;
            const mn = $('advFilterPctMin').value; _advFilter.pctMin = mn !== '' ? parseInt(mn) : null;
            const mx = $('advFilterPctMax').value; _advFilter.pctMax = mx !== '' ? parseInt(mx) : null;
            _updateAdvFilterBadge();
            panel.classList.add('hidden');
            renderTable();
        });

        $('btnSaveFilterPreset').addEventListener('click', () => {
            const name = prompt('Preset name:');
            if (!name) return;
            _filterPresets.push({ name, filter: { ..._advFilter, tags: [..._advFilter.tags] } });
            _saveFilterPresets();
            _populateAdvFilterPanel();
            showToast('success', `Preset "${name}" saved`);
        });
    }

    function _populateAdvFilterPanel() {
        if (!project) return;
        const tagContainer = $('advFilterTags');
        const resSelect    = $('advFilterResource');
        if (!tagContainer || !resSelect) return;

        // Tags
        const allTags = [...new Set(project.tasks.flatMap(t => t.tags || []))];
        tagContainer.innerHTML = '';
        allTags.forEach(tag => {
            const pill = document.createElement('button');
            pill.className = 'adv-tag-pill' + (_advFilter.tags.includes(tag) ? ' active' : '');
            pill.textContent = tag;
            pill.addEventListener('click', () => {
                if (_advFilter.tags.includes(tag)) _advFilter.tags = _advFilter.tags.filter(x => x !== tag);
                else _advFilter.tags.push(tag);
                pill.classList.toggle('active', _advFilter.tags.includes(tag));
            });
            tagContainer.appendChild(pill);
        });
        if (!allTags.length) tagContainer.innerHTML = '<span style="color:var(--text-muted);font-size:0.7rem">No tags defined</span>';

        // Resources
        const allRes = [...new Set(project.tasks.flatMap(t => t.resourceNames || []))].filter(Boolean);
        resSelect.innerHTML = '<option value="">All Resources</option>';
        allRes.forEach(r => {
            const opt = document.createElement('option');
            opt.value = r; opt.textContent = r;
            if (_advFilter.resource === r) opt.selected = true;
            resSelect.appendChild(opt);
        });

        // Restore date / pct values
        if ($('advFilterStartAfter'))  $('advFilterStartAfter').value  = _advFilter.startAfter  ? _advFilter.startAfter.toISOString().split('T')[0]  : '';
        if ($('advFilterFinishBefore')) $('advFilterFinishBefore').value = _advFilter.finishBefore ? _advFilter.finishBefore.toISOString().split('T')[0] : '';
        if ($('advFilterPctMin')) $('advFilterPctMin').value = _advFilter.pctMin  !== null ? _advFilter.pctMin  : '';
        if ($('advFilterPctMax')) $('advFilterPctMax').value = _advFilter.pctMax  !== null ? _advFilter.pctMax  : '';

        // Presets
        const presetsWrap = $('advFilterPresets');
        if (presetsWrap) {
            presetsWrap.innerHTML = '';
            _filterPresets.forEach((p, i) => {
                const pill = document.createElement('button');
                pill.className = 'adv-tag-pill';
                pill.textContent = p.name;
                pill.title = 'Click to apply · Right-click to delete';
                pill.addEventListener('click', () => {
                    _advFilter = { ...p.filter, tags: [...(p.filter.tags || [])] };
                    _populateAdvFilterPanel();
                    _updateAdvFilterBadge();
                    $('advFilterPanel').classList.add('hidden');
                    renderTable();
                });
                pill.addEventListener('contextmenu', (e) => {
                    e.preventDefault();
                    if (confirm(`Delete preset "${p.name}"?`)) {
                        _filterPresets.splice(i, 1); _saveFilterPresets(); _populateAdvFilterPanel();
                    }
                });
                presetsWrap.appendChild(pill);
            });
        }
    }

    // ── C.3: Table Row Drag & Drop ────────────────────────────
    let _dragUid = null;

    function initTableDragDrop() {
        const tbody = $('taskTableBody');
        if (!tbody) return;

        tbody.addEventListener('dragstart', e => {
            const tr = e.target.closest('tr[data-uid]');
            if (!tr) return;
            _dragUid = tr.dataset.uid;
            tr.classList.add('row-dragging');
            e.dataTransfer.effectAllowed = 'move';
        });

        tbody.addEventListener('dragend', () => {
            tbody.querySelectorAll('.row-dragging, .row-drag-over').forEach(el => el.classList.remove('row-dragging','row-drag-over'));
            _dragUid = null;
        });

        tbody.addEventListener('dragover', e => {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            const tr = e.target.closest('tr[data-uid]');
            tbody.querySelectorAll('.row-drag-over').forEach(el => el.classList.remove('row-drag-over'));
            if (tr && tr.dataset.uid !== _dragUid) tr.classList.add('row-drag-over');
        });

        tbody.addEventListener('drop', e => {
            e.preventDefault();
            const tr = e.target.closest('tr[data-uid]');
            if (!tr || !_dragUid || !project) return;
            const targetUid = parseInt(tr.dataset.uid);
            const srcUid    = parseInt(_dragUid);
            if (srcUid === targetUid) return;

            saveUndoState();
            const srcIdx    = project.tasks.findIndex(t => t.uid === srcUid);
            const targetIdx = project.tasks.findIndex(t => t.uid === targetUid);
            if (srcIdx < 0 || targetIdx < 0) return;

            const [moved] = project.tasks.splice(srcIdx, 1);
            const newIdx  = project.tasks.findIndex(t => t.uid === targetUid);
            project.tasks.splice(newIdx + (srcIdx < targetIdx ? 1 : 0), 0, moved);

            reindexTasks(); ProjectAnalytics.invalidate(); recalculate(); renderAll(); autoSave();
        });
    }

    // ── C.4: Batch Operations ─────────────────────────────────
    function initBatchOps() {
        const toolbar = $('batchToolbar');
        if (!toolbar) return;

        $('btnBatchClear').addEventListener('click', () => { selectedTaskIds.clear(); _updateBatchToolbar(); renderTable(); });
        $('btnBatchDelete').addEventListener('click', () => {
            if (!project || !selectedTaskIds.size) return;
            if (!confirm(`Delete ${selectedTaskIds.size} task(s)?`)) return;
            saveUndoState();
            project.tasks = project.tasks.filter(t => !selectedTaskIds.has(t.uid));
            selectedTaskIds.clear(); reindexTasks(); recalculate(); renderAll(); autoSave();
        });
        $('btnBatchStatus').addEventListener('click', () => {
            if (!project || !selectedTaskIds.size) return;
            const val = prompt('Set % Complete for selected tasks (0–100):');
            if (val === null) return;
            const pct = Math.min(100, Math.max(0, parseInt(val) || 0));
            saveUndoState();
            project.tasks.forEach(t => { if (selectedTaskIds.has(t.uid)) t.percentComplete = pct; });
            recalculate(); renderAll(); autoSave();
        });
        $('btnBatchResource').addEventListener('click', () => {
            if (!project || !selectedTaskIds.size) return;
            const val = prompt('Set Resource (comma separated) for selected tasks:');
            if (val === null) return;
            const names = val.split(',').map(s => s.trim()).filter(Boolean);
            saveUndoState();
            project.tasks.forEach(t => { if (selectedTaskIds.has(t.uid)) t.resourceNames = names; });
            recalculate(); renderAll(); autoSave();
        });
        $('btnBatchIndent').addEventListener('click', () => {
            if (!project || !selectedTaskIds.size) return;
            saveUndoState();
            project.tasks.forEach(t => { if (selectedTaskIds.has(t.uid)) { t.outlineLevel = (t.outlineLevel || 1) + 1; } });
            reindexTasks(); recalculate(); renderAll(); autoSave();
        });
        $('btnBatchOutdent').addEventListener('click', () => {
            if (!project || !selectedTaskIds.size) return;
            saveUndoState();
            project.tasks.forEach(t => { if (selectedTaskIds.has(t.uid)) { t.outlineLevel = Math.max(1, (t.outlineLevel || 1) - 1); } });
            reindexTasks(); recalculate(); renderAll(); autoSave();
        });
    }

    function _updateBatchToolbar() {
        const toolbar = $('batchToolbar');
        const countEl = $('batchCount');
        if (!toolbar || !countEl) return;
        const n = selectedTaskIds.size;
        toolbar.classList.toggle('hidden', n < 2);
        countEl.textContent = n + ' selected';
    }

    // ── D.1: Board View ───────────────────────────────────────
    function initBoardView() {
        BoardView.setCallbacks({
            onSave: (task) => { saveUndoState(); recalculate(); autoSave(); }
        });
        const groupSel  = $('boardGroupBy');
        const searchInp = $('boardSearch');
        if (groupSel)  groupSel.addEventListener('change',  () => { if (project) BoardView.render(project, groupSel.value, searchInp ? searchInp.value : ''); });
        if (searchInp) searchInp.addEventListener('input', debounce(() => { if (project) BoardView.render(project, groupSel ? groupSel.value : 'status', searchInp.value); }, 200));
    }

    function renderBoardView() {
        if (!project) return;
        const groupSel  = $('boardGroupBy');
        const searchInp = $('boardSearch');
        BoardView.render(project, groupSel ? groupSel.value : 'status', searchInp ? searchInp.value : '');
    }

    // ── D.3: Baseline Comparison (enhanced in Dashboard) ──────
    // Baseline bars already rendered in GanttChart via showBaseline flag.
    // We surface the toggle in the status bar tooltip text here.

    // ── D.4: Custom Fields ────────────────────────────────────
    function initCustomFields() {
        const btn = $('btnCustomFields');
        if (!btn) return;
        btn.addEventListener('click', () => {
            if (!project) return;
            _renderCustomFieldsList();
            toggleModal('modalCustomFields', true);
        });
        $('btnCloseCustomFields').addEventListener('click', () => toggleModal('modalCustomFields', false));

        const typeSelect = $('cfNewType');
        const optWrap    = $('cfDropdownOptions');
        if (typeSelect && optWrap) {
            typeSelect.addEventListener('change', () => { optWrap.classList.toggle('hidden', typeSelect.value !== 'dropdown'); });
        }

        $('btnAddCustomField').addEventListener('click', () => {
            const name = ($('cfNewName').value || '').trim();
            if (!name) { showToast('warning', 'Field name required'); return; }
            const type = $('cfNewType').value;
            const opts = type === 'dropdown'
                ? ($('cfDropdownOptionsList').value || '').split(',').map(s => s.trim()).filter(Boolean)
                : [];
            if (!project.customFields) project.customFields = [];
            project.customFields.push({ id: 'cf_' + Date.now(), name, type, options: opts });
            saveUndoState(); autoSave();
            _renderCustomFieldsList();
            $('cfNewName').value = '';
            renderTable(); // Re-render to show new column
            showToast('success', `Field "${name}" added`);
        });
    }

    function _renderCustomFieldsList() {
        const list = $('customFieldsList');
        if (!list || !project) return;
        list.innerHTML = '';
        (project.customFields || []).forEach((cf, i) => {
            const row = document.createElement('div');
            row.className = 'cf-list-row';
            row.innerHTML = `<span class="cf-name">${cf.name}</span><span class="cf-type">${cf.type}</span>`;
            const del = document.createElement('button');
            del.className = 'btn btn-ghost btn-xs cf-del-btn';
            del.textContent = '🗑';
            del.addEventListener('click', () => {
                if (!confirm(`Delete field "${cf.name}"?`)) return;
                project.customFields.splice(i, 1);
                saveUndoState(); autoSave(); _renderCustomFieldsList(); renderTable();
            });
            row.appendChild(del);
            list.appendChild(row);
        });
        if (!(project.customFields || []).length) {
            list.innerHTML = '<div style="color:var(--text-muted);font-size:0.75rem;padding:8px 0">No custom fields yet.</div>';
        }
    }

    // ── D.5: CSV Import ───────────────────────────────────────
    function initCSVImport() {
        const btn   = $('btnImportCSV');
        const input = $('csvImportInput');
        if (!btn || !input) return;

        btn.addEventListener('click', () => input.click());
        input.addEventListener('change', (e) => {
            const file = e.target.files[0]; if (!file) return;
            const reader = new FileReader();
            reader.onload = (ev) => {
                try {
                    const proj = _parseCSV(ev.target.result, file.name);
                    if (!proj) { showToast('error', 'Could not parse CSV'); return; }
                    project = proj;
                    activeProjectId = ProjectStore.generateId();
                    reindexTasks();
                    onProjectLoaded();
                    showToast('success', `Imported "${proj.name}" — ${proj.tasks.length} tasks`);
                } catch(err) {
                    showToast('error', 'CSV parse error: ' + err.message);
                }
                input.value = '';
            };
            reader.readAsText(file);
        });
    }

    /**
     * Parse a CSV file into a project object.
     * Supports standard columns: Name/Task, Start, Finish/End, Duration,
     * % Complete/Progress, Resource, Predecessor, Cost, Notes.
     * Also handles JIRA CSV exports.
     */
    function _parseCSV(text, filename) {
        const lines = text.split(/\r?\n/).filter(l => l.trim());
        if (lines.length < 2) throw new Error('CSV too short');

        // Parse header
        const sep = lines[0].includes('\t') ? '\t' : ',';
        const headers = _csvLine(lines[0], sep).map(h => h.toLowerCase().replace(/[^a-z0-9%]/g,''));

        // Column index helpers
        const col = (names) => { for (const n of names) { const i = headers.findIndex(h => h.includes(n)); if (i >= 0) return i; } return -1; };
        const iName     = col(['name','task','summary','issuetype','summary']);
        const iStart    = col(['start','startdate','created']);
        const iFinish   = col(['finish','end','enddate','duedate','due']);
        const iDur      = col(['duration','dur']);
        const iPct      = col(['complete','progress','pct','done','status']);
        const iRes      = col(['resource','assignee','owner','assigned']);
        const iPred     = col(['predecessor','depends','blockedby']);
        const iCost     = col(['cost','budget','estimate','storypoints']);
        const iNotes    = col(['notes','description','comment','body']);
        const iTag      = col(['tag','label','component','type','priority']);

        if (iName < 0) throw new Error('No Name/Task column found');

        const projectName = filename.replace(/\.(csv|txt)$/i, '') || 'Imported Project';
        const today = new Date(); today.setHours(0,0,0,0);
        const tasks = [];
        let uid = 1;

        lines.slice(1).forEach(line => {
            const cells = _csvLine(line, sep);
            const name = (cells[iName] || '').trim();
            if (!name) return;

            let start  = iStart  >= 0 ? new Date(cells[iStart]  || '') : new Date(today);
            let finish = iFinish >= 0 ? new Date(cells[iFinish] || '') : null;
            let dur    = iDur    >= 0 ? parseInt(cells[iDur] || '') || 0 : 0;

            if (isNaN(start.getTime()))  start  = new Date(today);
            if (finish && isNaN(finish.getTime())) finish = null;

            if (!finish && dur > 0) finish = addDays(start, dur);
            if (!finish) finish = addDays(start, 1);
            if (finish <= start) finish = addDays(start, 1);
            dur = Math.max(1, Math.round((finish - start) / 86400000));

            const rawPct = iPct >= 0 ? cells[iPct] : '';
            let pct = 0;
            // Handle JIRA-style status text
            if (/done|complete|closed|resolved/i.test(rawPct)) pct = 100;
            else if (/progress|active|open/i.test(rawPct)) pct = 50;
            else pct = Math.min(100, Math.max(0, parseInt(rawPct) || 0));

            const resourceNames = iRes >= 0 && cells[iRes] ? cells[iRes].split(/[,;]/).map(s=>s.trim()).filter(Boolean) : [];
            const cost          = iCost >= 0 ? (parseFloat(cells[iCost]) || 0) : 0;
            const notes         = iNotes >= 0 ? (cells[iNotes] || '') : '';
            const tags          = iTag >= 0 && cells[iTag] ? cells[iTag].split(/[,;|]/).map(s=>s.trim()).filter(Boolean) : [];
            const predRaw       = iPred >= 0 ? (cells[iPred] || '') : '';
            const predecessors  = predRaw ? predRaw.split(/[,;]/).map(s => {
                const n = parseInt(s); return isNaN(n) ? null : { predecessorUID: n, type: 1, typeName: 'FS', lag: 0 };
            }).filter(Boolean) : [];

            tasks.push({
                uid, id: uid, outlineLevel: 1, outlineNumber: String(uid),
                name, start, finish, durationDays: dur,
                percentComplete: pct, resourceNames, cost, notes, tags,
                predecessors, summary: false, milestone: false,
                isExpanded: true, isVisible: true,
                wbs: String(uid), baselineStart: null, baselineFinish: null
            });
            uid++;
        });

        if (!tasks.length) throw new Error('No tasks found in CSV');

        const projectStart  = new Date(Math.min(...tasks.map(t => t.start.getTime())));
        const projectFinish = new Date(Math.max(...tasks.map(t => t.finish.getTime())));

        return {
            name: projectName, startDate: projectStart, finishDate: projectFinish,
            tasks, resources: [], assignments: [], customFields: [],
            hoursPerDay: 8, minutesPerDay: 480
        };
    }

    /** Parse a single CSV line respecting quoted fields */
    function _csvLine(line, sep) {
        const result = []; let cur = ''; let inQ = false;
        for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (ch === '"') { inQ = !inQ; }
            else if (ch === sep && !inQ) { result.push(cur.trim()); cur = ''; }
            else cur += ch;
        }
        result.push(cur.trim());
        return result;
    }

    // ══════════════════════════════════════════════════════════
    // SPRINT E & F + PHASE 6b
    // ══════════════════════════════════════════════════════════

    // ── 6b.1: Share via Link ──────────────────────────────────
    function initShareLink() {
        const btn = $('btnShareLink');
        if (btn) btn.addEventListener('click', handleShareLink);
        // Read share link from URL hash on startup
        if (location.hash && location.hash.startsWith('#share=')) {
            _loadFromShareHash(location.hash.slice(7));
        }
    }

    function handleShareLink() {
        if (!project) return;
        try {
            // Strip attachments for size, then encode
            const slim = JSON.parse(JSON.stringify(project));
            (slim.tasks || []).forEach(t => { delete t.attachments; });
            const json = JSON.stringify(slim);
            const encoded = btoa(encodeURIComponent(json));
            if (encoded.length > 200000) {
                showToast('warning', 'Project too large for share link (>200KB). Export XML instead.');
                return;
            }
            const url = location.origin + location.pathname + '#share=' + encoded;
            navigator.clipboard.writeText(url).then(() => {
                showToast('success', 'Share link copied! Anyone with the link can view this project.');
            }).catch(() => {
                prompt('Copy this share link:', url);
            });
        } catch(e) {
            showToast('error', 'Could not create share link: ' + e.message);
        }
    }

    function _loadFromShareHash(encoded) {
        try {
            const json = decodeURIComponent(atob(encoded));
            const proj = JSON.parse(json);
            if (!proj || !proj.tasks) return;
            // Restore date objects
            (proj.tasks || []).forEach(t => {
                if (t.start)  t.start  = new Date(t.start);
                if (t.finish) t.finish = new Date(t.finish);
                t.isExpanded = true; t.isVisible = true;
            });
            if (proj.startDate)  proj.startDate  = new Date(proj.startDate);
            if (proj.finishDate) proj.finishDate = new Date(proj.finishDate);
            project = proj;
            activeProjectId = ProjectStore.generateId();
            reindexTasks();
            onProjectLoaded();
            showToast('info', `Loaded shared project: "${proj.name}" (read from link)`);
            history.replaceState(null, '', location.pathname); // clean URL
        } catch(e) {
            console.warn('[ShareLink] Failed to load from hash:', e);
        }
    }

    // ── 6b.2 + E.4: What-If Scenarios ────────────────────────
    function initScenarios() {
        const btnSave    = $('btnSaveScenario');
        const btnManage  = $('btnManageScenarios');
        const btnClose   = $('btnCloseScenariosModal');
        if (btnSave)   btnSave.addEventListener('click',   handleSaveScenario);
        if (btnManage) btnManage.addEventListener('click', handleOpenScenarios);
        if (btnClose)  btnClose.addEventListener('click',  () => toggleModal('modalScenarios', false));
    }

    function handleSaveScenario() {
        if (!project) return;
        const name = prompt('Scenario name:', `${project.name} — ${new Date().toLocaleDateString()}`);
        if (!name) return;
        const snap = ScenariosManager.save(project, name);
        showToast('success', `Scenario "${snap.name}" saved (${snap.id})`);
    }

    function handleOpenScenarios() {
        if (!project) { showToast('info', 'Open a project first'); return; }
        const container = $('scenariosPanelBody');
        if (!container) return;
        ScenariosManager.renderScenariosPanel(container, project, {
            onLoad: (snapshot) => {
                if (!confirm('Load this scenario? Unsaved changes will be lost.')) return;
                const loaded = JSON.parse(JSON.stringify(snapshot));
                (loaded.tasks || []).forEach(t => { t.start = new Date(t.start); t.finish = new Date(t.finish); t.isExpanded = true; t.isVisible = true; });
                if (loaded.startDate)  loaded.startDate  = new Date(loaded.startDate);
                if (loaded.finishDate) loaded.finishDate = new Date(loaded.finishDate);
                project = loaded;
                ProjectAnalytics.reset();
                reindexTasks(); recalculate(); renderAll();
                toggleModal('modalScenarios', false);
                showToast('success', 'Scenario loaded');
            },
            onCompare: (snapA, snapB, nameA, nameB) => {
                const container = $('scenariosCompareBody');
                if (!container) return;
                ScenariosManager.renderCompareUI(container, snapA, snapB, nameA, nameB);
                $('scenariosComparePanelTitle').textContent = `Compare: ${nameA} vs ${nameB}`;
                $('scenariosComparePanel').classList.remove('hidden');
            }
        });
        toggleModal('modalScenarios', true);
    }

    // ── E.1: PWA Registration ─────────────────────────────────
    function initPWA() {
        if ('serviceWorker' in navigator) {
            navigator.serviceWorker.register('/sw.js', { scope: '/' })
                .then(reg => {
                    console.log('[PWA] Service Worker registered:', reg.scope);
                    // Check for update
                    reg.addEventListener('updatefound', () => {
                        const newWorker = reg.installing;
                        newWorker.addEventListener('statechange', () => {
                            if (newWorker.state === 'installed' && navigator.serviceWorker.controller) {
                                showToast('info', '🔄 App update available — refresh to apply');
                            }
                        });
                    });
                })
                .catch(err => console.warn('[PWA] SW registration failed:', err));
        }
        // Handle PWA install prompt
        let _deferredInstall = null;
        window.addEventListener('beforeinstallprompt', (e) => {
            e.preventDefault();
            _deferredInstall = e;
            const btn = $('btnInstallPWA');
            if (btn) { btn.classList.remove('hidden'); btn.addEventListener('click', () => { _deferredInstall.prompt(); }); }
        });
        // Handle URL shortcuts from manifest
        const params = new URLSearchParams(location.search);
        if (params.get('action') === 'new') setTimeout(() => toggleModal('modalNewProject', true), 500);
        if (params.get('view') === 'portfolio') setTimeout(() => setView('portfolio'), 500);
    }

    // ── E.3: Portfolio Bubble Chart ───────────────────────────
    function renderPortfolioBubbleChart(index) {
        const canvas = $('portfolioBubbleCanvas');
        if (!canvas || !index || !index.length) return;
        const ctx = canvas.getContext('2d');
        const W = canvas.width  = Math.floor(canvas.parentElement.offsetWidth  || 600);
        const H = canvas.height = Math.floor(canvas.parentElement.offsetHeight || 320);
        ctx.clearRect(0, 0, W, H);

        // Axes: X = health score (0-100), Y = progress (0-100), size = task count
        const pad = 48;
        const plotW = W - pad * 2, plotH = H - pad * 2;

        // Grid
        ctx.strokeStyle = 'rgba(255,255,255,0.06)'; ctx.lineWidth = 1;
        for (let i = 0; i <= 4; i++) {
            const x = pad + (i / 4) * plotW;
            const y = pad + (i / 4) * plotH;
            ctx.beginPath(); ctx.moveTo(x, pad); ctx.lineTo(x, pad + plotH); ctx.stroke();
            ctx.beginPath(); ctx.moveTo(pad, y); ctx.lineTo(pad + plotW, y); ctx.stroke();
        }

        // Labels
        ctx.fillStyle = 'rgba(255,255,255,0.35)'; ctx.font = '10px Inter,sans-serif'; ctx.textAlign = 'center';
        ctx.fillText('Health Score →', W / 2, H - 6);
        ctx.save(); ctx.translate(12, H / 2); ctx.rotate(-Math.PI / 2);
        ctx.fillText('Progress % →', 0, 0); ctx.restore();

        // Bubbles
        const maxTasks = Math.max(...index.map(p => p.taskCount || 1), 1);
        index.forEach(p => {
            const health   = p.health ? p.health.score : (p.progress || 0);
            const progress = p.progress || 0;
            const bx = pad + (health   / 100) * plotW;
            const by = pad + ((100 - progress) / 100) * plotH;
            const r  = 8 + 18 * Math.sqrt((p.taskCount || 1) / maxTasks);
            const color = health >= 75 ? 'rgba(34,197,94,0.5)' : health >= 45 ? 'rgba(245,158,11,0.5)' : 'rgba(239,68,68,0.5)';
            const border = health >= 75 ? '#22c55e' : health >= 45 ? '#f59e0b' : '#ef4444';

            ctx.beginPath(); ctx.arc(Math.floor(bx), Math.floor(by), Math.floor(r), 0, Math.PI * 2);
            ctx.fillStyle = color; ctx.fill();
            ctx.strokeStyle = border; ctx.lineWidth = 1.5; ctx.stroke();

            // Label
            const name = p.name && p.name.length > 12 ? p.name.slice(0, 12) + '…' : (p.name || '?');
            ctx.fillStyle = 'rgba(255,255,255,0.85)'; ctx.font = 'bold 9px Inter,sans-serif'; ctx.textAlign = 'center';
            ctx.fillText(name, Math.floor(bx), Math.floor(by) + 3);
        });
    }

    // ── E.5: Plugin System Integration ───────────────────────
    function initPluginSystem() {
        if (typeof PluginSystem === 'undefined') return;
        // Wire EventBus → PluginSystem
        EventBus.on('project:changed', (data) => PluginSystem.emit('project:changed', data));
        EventBus.on('project:loaded',  (data) => PluginSystem.emit('project:loaded',  data));

        // Plugin manager modal
        const btnPlugins = $('btnPluginManager');
        const btnClose   = $('btnClosePluginManager');
        if (btnPlugins) btnPlugins.addEventListener('click', () => {
            const body = $('pluginManagerBody');
            if (body && typeof PluginSystem !== 'undefined') PluginSystem.renderPluginManager(body);
            toggleModal('modalPluginManager', true);
        });
        if (btnClose) btnClose.addEventListener('click', () => toggleModal('modalPluginManager', false));

        // Auto-activate saved plugins from localStorage
        try {
            const active = JSON.parse(localStorage.getItem('pf_active_plugins') || '[]');
            active.forEach(id => {
                const builtin = PluginSystem.builtins[id];
                if (builtin) PluginSystem.register(builtin);
            });
        } catch(_) {}
    }

    // ── F.2: Project Roles ────────────────────────────────────
    const _PROJECT_ROLES = ['Owner', 'Manager', 'Member', 'Viewer'];

    function initRoles() {
        const btnRoles = $('btnProjectRoles');
        if (btnRoles) btnRoles.addEventListener('click', () => {
            if (!project) return;
            if (!project.roles) project.roles = { members: [] };
            _renderRolesPanel();
            toggleModal('modalRoles', true);
        });
        if ($('btnCloseRoles')) $('btnCloseRoles').addEventListener('click', () => toggleModal('modalRoles', false));
        if ($('btnAddMember')) $('btnAddMember').addEventListener('click', _handleAddMember);
    }

    function _renderRolesPanel() {
        const list = $('membersList');
        if (!list || !project) return;
        list.innerHTML = '';
        const members = (project.roles || {}).members || [];
        if (!members.length) {
            list.innerHTML = '<div style="color:var(--text-muted);font-size:0.75rem">No members yet. Add email addresses below.</div>';
            return;
        }
        members.forEach((m, i) => {
            const row = document.createElement('div'); row.className = 'member-row';
            const em = document.createElement('span'); em.className = 'member-email'; em.textContent = m.email;
            const roleSelect = document.createElement('select'); roleSelect.className = 'filter-select';
            _PROJECT_ROLES.forEach(r => { const opt = document.createElement('option'); opt.value = r; opt.textContent = r; if (m.role === r) opt.selected = true; roleSelect.appendChild(opt); });
            roleSelect.addEventListener('change', () => { m.role = roleSelect.value; autoSave(); });
            const del = document.createElement('button'); del.className = 'btn btn-ghost btn-xs'; del.textContent = '✕';
            del.addEventListener('click', () => { members.splice(i, 1); autoSave(); _renderRolesPanel(); });
            row.appendChild(em); row.appendChild(roleSelect); row.appendChild(del);
            list.appendChild(row);
        });
    }

    function _handleAddMember() {
        const emailInput = $('newMemberEmail');
        const roleSelect = $('newMemberRole');
        if (!emailInput || !project) return;
        const email = (emailInput.value || '').trim().toLowerCase();
        if (!email || !email.includes('@')) { showToast('warning', 'Enter a valid email'); return; }
        if (!project.roles) project.roles = { members: [] };
        if (project.roles.members.some(m => m.email === email)) { showToast('warning', 'Already added'); return; }
        project.roles.members.push({ email, role: roleSelect ? roleSelect.value : 'Member' });
        emailInput.value = '';
        autoSave(); _renderRolesPanel();
        showToast('success', `${email} added as ${roleSelect ? roleSelect.value : 'Member'}`);
    }

    // ── F.3: Webhook Framework ────────────────────────────────
    let _webhooks = [];

    function initWebhooks() {
        _loadWebhooks();
        // Wire EventBus to fire webhooks
        ['project:loaded','project:changed','project:saved'].forEach(evt => {
            EventBus.on(evt, (data) => _fireWebhooks(evt, data));
        });
        const btn = $('btnWebhooks');
        if (btn) btn.addEventListener('click', () => { _renderWebhooksPanel(); toggleModal('modalWebhooks', true); });
        if ($('btnCloseWebhooks')) $('btnCloseWebhooks').addEventListener('click', () => toggleModal('modalWebhooks', false));
        if ($('btnAddWebhook'))    $('btnAddWebhook').addEventListener('click',    _handleAddWebhook);
    }

    function _loadWebhooks()  { try { _webhooks = JSON.parse(localStorage.getItem('pf_webhooks') || '[]'); } catch(_) { _webhooks = []; } }
    function _saveWebhooks()  { try { localStorage.setItem('pf_webhooks', JSON.stringify(_webhooks)); } catch(_) {} }

    function _fireWebhooks(event, data) {
        const matching = _webhooks.filter(w => w.enabled && (w.event === event || w.event === '*'));
        if (!matching.length) return;
        const payload = { event, timestamp: new Date().toISOString(), projectName: project ? project.name : null, data: { taskCount: project ? project.tasks.length : 0 } };
        matching.forEach(w => {
            fetch(w.url, { method: 'POST', headers: { 'Content-Type': 'application/json', 'X-ProjectFlow-Event': event }, body: JSON.stringify(payload) })
                .catch(err => console.warn(`[Webhook] ${w.url} failed:`, err.message));
        });
    }

    function _renderWebhooksPanel() {
        const list = $('webhooksList');
        if (!list) return;
        list.innerHTML = '';
        if (!_webhooks.length) { list.innerHTML = '<div style="color:var(--text-muted);font-size:0.75rem">No webhooks configured.</div>'; return; }
        _webhooks.forEach((w, i) => {
            const row = document.createElement('div'); row.className = 'member-row';
            const url = document.createElement('span'); url.className = 'member-email'; url.textContent = w.url;
            url.style.cssText = 'max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap';
            const badge = document.createElement('span'); badge.className = 'cf-type'; badge.textContent = w.event;
            const toggle = document.createElement('button'); toggle.className = 'btn btn-ghost btn-xs'; toggle.textContent = w.enabled ? '🟢' : '🔴';
            toggle.addEventListener('click', () => { w.enabled = !w.enabled; _saveWebhooks(); _renderWebhooksPanel(); });
            const del = document.createElement('button'); del.className = 'btn btn-ghost btn-xs'; del.textContent = '✕';
            del.addEventListener('click', () => { _webhooks.splice(i, 1); _saveWebhooks(); _renderWebhooksPanel(); });
            row.appendChild(url); row.appendChild(badge); row.appendChild(toggle); row.appendChild(del);
            list.appendChild(row);
        });
    }

    function _handleAddWebhook() {
        const urlEl   = $('webhookURL');
        const evtEl   = $('webhookEvent');
        if (!urlEl) return;
        const url = (urlEl.value || '').trim();
        if (!url.startsWith('http')) { showToast('warning', 'Enter a valid URL'); return; }
        _webhooks.push({ url, event: evtEl ? evtEl.value : '*', enabled: true });
        _saveWebhooks(); urlEl.value = ''; _renderWebhooksPanel();
        showToast('success', 'Webhook added');
    }

    // ── E.3: Extend renderPortfolioView with Bubble Chart ─────
    EventBus.on('project:loaded', () => {
        // Bubble chart is rendered when portfolio view is opened
    });

    // ══════════════════════════════════════════════════════════
    // MS INTEGRATIONS — Planner Live Sync + D365
    // ══════════════════════════════════════════════════════════

    // ── MS Planner Live Sync ──────────────────────────────────
    let _plannerConnectedPlanId = null;

    function initPlannerSync() {
        const btn      = $('btnPlannerSync');
        const btnClose = $('btnClosePlannerSync');
        if (btn)      btn.addEventListener('click', openPlannerSyncModal);
        if (btnClose) btnClose.addEventListener('click', () => toggleModal('modalPlannerSync', false));
    }

    function openPlannerSyncModal() {
        const body = $('plannerSyncBody');
        if (!body) return;
        body.innerHTML = '';
        toggleModal('modalPlannerSync', true);

        if (typeof MSGraphClient === 'undefined') {
            body.textContent = '⚠ MS Graph library not loaded. Check internet connection.';
            return;
        }

        if (MSGraphClient.isAuthenticated() && _plannerConnectedPlanId && project) {
            // Already connected with active project — show sync panel
            MSGraphClient.renderSyncPanel(body, project, _plannerConnectedPlanId);
        } else {
            // Show setup wizard — always allow connecting and importing
            MSGraphClient.renderSetupWizard(body, async ({ planId, planTitle }) => {
                _plannerConnectedPlanId = planId;
                showToast('info', `Connected to Planner: "${planTitle}". Importing...`);
                // Always import the plan as a new project
                try {
                    setStatus('Importing from MS Planner…');
                    const imported = await MSGraphClient.importPlan(planId);
                    if (!imported) { showToast('error', 'Import returned empty project'); return; }
                    project = imported;
                    // Pre-process project data
                    project.tasks.forEach(t => {
                        if (t.start && !(t.start instanceof Date)) t.start = new Date(t.start);
                        if (t.finish && !(t.finish instanceof Date)) t.finish = new Date(t.finish);
                        t.isExpanded = true; t.isVisible = true;
                        if (!t.predecessors) t.predecessors = [];
                        if (!t.resourceNames) t.resourceNames = [];
                    });
                    if (project.startDate && !(project.startDate instanceof Date)) project.startDate = new Date(project.startDate);
                    if (project.finishDate && !(project.finishDate instanceof Date)) project.finishDate = new Date(project.finishDate);
                    reindexTasks();
                    activeProjectId = ProjectStore.generateId();
                    onProjectLoaded();
                    showToast('success', `Imported from Planner: "${project.name}" — ${project.tasks.length} items`);
                    toggleModal('modalPlannerSync', false);
                } catch(e) {
                    showToast('error', 'Planner import failed: ' + e.message);
                } finally { setStatus('Ready'); }
            });
        }
    }

    /** Import a Planner plan as a new ProjectFlow project */
    async function handlePlannerImport(planId) {
        if (typeof MSGraphClient === 'undefined') return;
        try {
            setStatus('Importing from MS Planner…');
            const imported = await MSGraphClient.importPlan(planId);
            if (!imported) { showToast('error', 'Import returned empty project'); return; }
            project = imported;
            activeProjectId = ProjectStore.generateId();
            reindexTasks();
            onProjectLoaded();
            showToast('success', `Imported from Planner: "${project.name}" — ${project.tasks.length} items`);
            toggleModal('modalPlannerSync', false);
        } catch(e) {
            showToast('error', 'Planner import failed: ' + e.message);
        } finally { setStatus('Ready'); }
    }

    /**
     * Auto-import a single Planner plan on startup (no UI needed)
     */
    async function _autoImportPlan(planId, planTitle) {
        try {
            setStatus('Loading Planner project…');
            const imported = await MSGraphClient.importPlan(planId);
            if (!imported) return;
            project = imported;
            _plannerConnectedPlanId = planId;
            project.tasks.forEach(t => {
                if (t.start  && !(t.start  instanceof Date)) t.start  = new Date(t.start);
                if (t.finish && !(t.finish instanceof Date)) t.finish = new Date(t.finish);
                t.isExpanded = true; t.isVisible = true;
                if (!t.predecessors)  t.predecessors  = [];
                if (!t.resourceNames) t.resourceNames = [];
            });
            if (project.startDate  && !(project.startDate  instanceof Date)) project.startDate  = new Date(project.startDate);
            if (project.finishDate && !(project.finishDate instanceof Date)) project.finishDate = new Date(project.finishDate);
            reindexTasks();
            activeProjectId = ProjectStore.generateId();
            onProjectLoaded();
            TeamsBridge.saveLastPlan(planId);
            showToast('success', `✅ Loaded from Planner: "${planTitle}" — ${project.tasks.length} tasks`);
        } catch(e) {
            showToast('error', 'Auto-import failed: ' + e.message);
        } finally { setStatus('Ready'); }
    }

    /**
     * Show a modal for the user to pick which Planner plan to open.
     * Highlights the last-used plan.
     */
    function _showPlanPickerModal(plans, lastPlanId) {
        // Build a simple inline modal overlay
        const overlay = document.createElement('div');
        overlay.id = 'planPickerOverlay';
        overlay.style.cssText = `
            position:fixed;inset:0;background:rgba(0,0,0,0.65);
            display:flex;align-items:center;justify-content:center;z-index:9999;
        `;

        const card = document.createElement('div');
        card.style.cssText = `
            background:var(--bg-card,#1e1e2e);border:1px solid rgba(255,255,255,0.12);
            border-radius:14px;padding:28px;width:420px;max-width:92vw;
            font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
            box-shadow:0 20px 60px rgba(0,0,0,0.5);
        `;

        card.innerHTML = `
            <div style="text-align:center;margin-bottom:20px;">
                <div style="font-size:2rem;">📋</div>
                <h3 style="margin:8px 0 4px;font-size:1.1rem;color:var(--text-primary,#e2e8f0);">
                    Select a Planner Project
                </h3>
                <p style="margin:0;font-size:0.8rem;color:var(--text-muted,#888);">
                    ${plans.length} plan${plans.length > 1 ? 's' : ''} found in your account
                </p>
            </div>
            <div id="planPickerList" style="display:flex;flex-direction:column;gap:8px;max-height:300px;overflow-y:auto;"></div>
        `;

        const list = card.querySelector('#planPickerList');
        plans.forEach(plan => {
            const isLast = plan.id === lastPlanId;
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.style.cssText = `
                width:100%;padding:12px 16px;text-align:left;border-radius:8px;cursor:pointer;
                border:1px solid ${isLast ? '#6366f1' : 'rgba(255,255,255,0.1)'};
                background:${isLast ? 'rgba(99,102,241,0.15)' : 'rgba(255,255,255,0.04)'};
                color:var(--text-primary,#e2e8f0);font-size:0.9rem;font-weight:500;
                transition:all 0.15s;
            `;
            btn.innerHTML = `${isLast ? '⭐ ' : ''}${plan.title}`;
            btn.addEventListener('mouseenter', () => { btn.style.borderColor = '#6366f1'; btn.style.background = 'rgba(99,102,241,0.12)'; });
            btn.addEventListener('mouseleave', () => { btn.style.borderColor = isLast ? '#6366f1' : 'rgba(255,255,255,0.1)'; btn.style.background = isLast ? 'rgba(99,102,241,0.15)' : 'rgba(255,255,255,0.04)'; });
            btn.addEventListener('click', () => {
                overlay.remove();
                _autoImportPlan(plan.id, plan.title);
            });
            list.appendChild(btn);
        });

        overlay.appendChild(card);
        document.body.appendChild(overlay);
    }

    // ── D365 Project Accounting ───────────────────────────────
    let _d365ConnectedProjectId = null;

    function initD365Sync() {
        const btn      = $('btnD365Sync');
        const btnClose = $('btnCloseD365Sync');
        if (btn)      btn.addEventListener('click', openD365Modal);
        if (btnClose) btnClose.addEventListener('click', () => toggleModal('modalD365Sync', false));
    }

    function openD365Modal() {
        const body = $('d365SyncBody');
        if (!body) return;
        body.innerHTML = '';
        toggleModal('modalD365Sync', true);

        if (typeof D365Client === 'undefined') {
            body.textContent = '⚠ D365 library not loaded. Check internet connection.';
            return;
        }

        if (D365Client.isAuthenticated() && _d365ConnectedProjectId) {
            D365Client.renderSyncPanel(body, project, _d365ConnectedProjectId);
        } else {
            D365Client.renderSetupWizard(body, async ({ projectId, projectName, mode }) => {
                _d365ConnectedProjectId = projectId;
                showToast('success', `Connected to D365: "${projectName}" (${mode})`);
                try {
                    setStatus('Importing from D365…');
                    const imported = await D365Client.importProject(projectId);
                    if (imported) {
                        project = imported;
                        activeProjectId = ProjectStore.generateId();
                        reindexTasks(); onProjectLoaded();
                        showToast('success', `Imported D365: "${project.name}"`);
                    }
                } catch(e) { showToast('error', 'D365 import failed: ' + e.message); }
                finally { setStatus('Ready'); }
                body.innerHTML = '';
                D365Client.renderSyncPanel(body, project, projectId);
            });
        }
    }

    // ══════════════════════════════════════════════════════════
    // SPRINT B.3 — window.PF Public Namespace
    // Exposes internal state + actions so external modules
    // (StateManager, TaskEditor, ProjectIO, UIHelpers, plugins)
    // can work without touching the closure directly.
    // ══════════════════════════════════════════════════════════
    function _exposePFNamespace() {
        window.PF = Object.freeze({
            // ── Live state (via getters) ──────────────────────
            get project()           { return project; },
            set project(p)          { project = p; },
            get settings()          { return settings; },
            get undoStack()         { return undoStack; },
            get selectedTaskIds()   { return selectedTaskIds; },
            get activeProjectId()   { return activeProjectId; },
            set activeProjectId(id) { activeProjectId = id; },
            get activeView()        { return activeView; },

            // ── Core actions ──────────────────────────────────
            showToast,
            autoSave,
            recalculate,
            renderAll,
            renderTable,
            renderGantt,
            saveUndoState,
            mutation,
            reindexTasks,
            openDetailPanel,
            setView,
            setStatus,
            toggleModal,
            formatDate,
            parseInputDate,
            addDays,
            debounce,
            downloadFile,
            onProjectLoaded,

            // ── Module references ─────────────────────────────
            ProjectStore,
            ProjectAnalytics,
            EventBus,
        });
    }

    // ══════════════════════════════════════════════════════════
    // HEADER DROPDOWNS — Settings & Scenarios
    // ══════════════════════════════════════════════════════════
    function initHeaderDropdowns() {
        const drops = [
            { triggerId: 'hdrSettingsBtn',  menuId: 'hdrSettingsDrop'  },
            { triggerId: 'hdrScenariosBtn', menuId: 'hdrScenariosDrop' },
        ];

        drops.forEach(({ triggerId, menuId }) => {
            const trigger = $(triggerId);
            const wrap    = $(menuId);
            if (!trigger || !wrap) return;
            const menu = wrap.querySelector('.hdr-drop-menu');
            if (!menu) return;

            trigger.setAttribute('aria-expanded', 'false');

            trigger.addEventListener('click', (e) => {
                e.stopPropagation();
                const isOpen = menu.classList.toggle('open');
                trigger.setAttribute('aria-expanded', isOpen);
                // Close sibling dropdowns
                drops.forEach(d => {
                    if (d.menuId === menuId) return;
                    const otherWrap = $(d.menuId);
                    if (!otherWrap) return;
                    otherWrap.querySelector('.hdr-drop-menu')?.classList.remove('open');
                    $(d.triggerId)?.setAttribute('aria-expanded', 'false');
                });
                // Also close the nav More menu
                document.getElementById('navMoreMenu')?.classList.remove('open');
            });
        });

        // Close all header dropdowns on outside click
        document.addEventListener('click', () => {
            drops.forEach(({ triggerId, menuId }) => {
                $(menuId)?.querySelector('.hdr-drop-menu')?.classList.remove('open');
                $(triggerId)?.setAttribute('aria-expanded', 'false');
            });
        });

        // Close on Escape
        document.addEventListener('keydown', (e) => {
            if (e.key !== 'Escape') return;
            drops.forEach(({ triggerId, menuId }) => {
                $(menuId)?.querySelector('.hdr-drop-menu')?.classList.remove('open');
                $(triggerId)?.setAttribute('aria-expanded', 'false');
            });
        });
    }

    document.addEventListener('DOMContentLoaded', async () => {
        await init();
        _exposePFNamespace();
        // Notify B.3 modules that PF is ready
        EventBus.emit('pf:ready', { version: '1.0' });
    });
