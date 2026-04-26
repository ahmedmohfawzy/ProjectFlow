/**
 * ProjectFlow™ © 2026 Ahmed M. Fawzy. All Rights Reserved.
 */
/**
 * ═══════════════════════════════════════════════════════════════
 * Network Diagram v4  —  PERT / CPM  +  Decision Intelligence
 *
 *  Key fix over v3:
 *  ─ Summary predecessor resolution: when a task's predecessor
 *    is a summary task (filtered from the visible set), we
 *    transparently replace it with the last non-summary
 *    descendant of that summary.  This ensures edges are drawn
 *    and the Sugiyama layer assignment is correct.
 *
 *  Layout    : Sugiyama + gravity-based Y positioning
 *  Nodes     : compact PERT boxes  (3 sizes: normal/compact/micro)
 *  Edges     : orthogonal elbow routing
 *  Highlight : Critical path glow, hover ripple, impact chain
 *  Filters   : Critical-only | Late | At-Risk | All
 *  Search    : Find & highlight task by name
 *  Risk Halo : Per-node risk score (0–100) via colour border
 *  Impact    : Click → dim all non-downstream tasks
 *  Bottleneck: ⚡ badge on nodes with ≥3 successors
 *  Stats Bar : Critical path length, bottlenecks, avg float
 *  Minimap   : Overview with viewport rect
 *  Keyboard  : +/- zoom, F fit, C critical-only, Esc deselect
 *  Export    : Tight-crop PNG at up to 4K resolution
 * ═══════════════════════════════════════════════════════════════
 */


    // ── Node dimensions (3 modes) ──────────────────────────────
    const DIMS = {
        normal:  { w: 180, h: 74,  fs: 10.5, gapX: 56, gapY: 18, pad: 36 },
        compact: { w: 136, h: 42,  fs: 9.5,  gapX: 40, gapY: 12, pad: 26 },
        micro:   { w: 100, h: 28,  fs: 8.5,  gapX: 28, gapY:  8, pad: 20 },
    };
    let _mode = 'normal';

    // ── State ──────────────────────────────────────────────────
    let _canvas, _ctx, _wrap;
    let _allTasks   = [];   // full non-summary visible list (for re-filter)
    let _allTasksRaw = [];  // ALL tasks including summaries (for pred resolution)
    let _tasks = [], _nodes = [], _edges = [];
    let _nodeMap  = new Map();  // uid → node
    let _succMap  = new Map();  // uid → [uid]  (resolved successors)
    let _predMap  = new Map();  // uid → [uid]  (resolved predecessors)
    let _impactSet = new Set();
    let _statsCache = null;

    let _panX = 0, _panY = 0, _scale = 1;
    let _isPanning = false, _pStart = { x:0, y:0 };
    let _hovUid = null, _selUid = null;
    let _filterMode = 'all';   // 'all' | 'critical' | 'late' | 'atrisk'
    let _searchQuery = '';
    let _highlightUids = new Set();
    let _abortCtrl = null;
    let _dpr = 1;
    let _touches = [], _touchDist = 0;
    let _tipEl = null;

    // RAF-based draw scheduling — coalesces rapid redraws into one per frame
    let _rafId = null;
    function _schedDraw() {
        if (_rafId) return;
        _rafId = requestAnimationFrame(() => { _rafId = null; _draw(); });
    }

    // Dot-grid pattern cached as an offscreen canvas — avoids O(W×H/784) arc() calls
    let _gridPattern = null, _gridBg = '';
    function _getGridPattern() {
        if (_gridPattern && _gridBg === C.bg) return _gridPattern;
        const off = document.createElement('canvas'); off.width = 28; off.height = 28;
        const oc  = off.getContext('2d');
        oc.fillStyle = C.bg; oc.fillRect(0, 0, 28, 28);
        oc.fillStyle = 'rgba(255,255,255,0.025)';
        oc.beginPath(); oc.arc(0, 0, 1, 0, Math.PI * 2); oc.fill();
        _gridPattern = _ctx.createPattern(off, 'repeat');
        _gridBg = C.bg;
        return _gridPattern;
    }
    // Whether the source project had dependency data available.
    // false  → show "no links" banner over the diagram.
    // null   → unknown / not yet set (no banner shown).
    let _dependenciesAvailable = null;
    let _projectStartDate = null; // Actual project start date for date conversion

    // Minimap
    const MM = { w: 148, h: 88, pad: 8 };

    // Cached offscreen minimap — rebuilt only when layout/theme changes
    let _mmCanvas = null, _mmDirty = true;
    function _invalidateMinimap() { _mmDirty = true; }

    // ── Colors ─────────────────────────────────────────────────
    let C = {};
    function _clr() {
        _gridPattern = null; _gridBg = ''; // invalidate cached grid on theme change
        _invalidateMinimap();               // theme colours changed → rebuild minimap
        const s = getComputedStyle(document.documentElement);
        const g = (v,fb) => s.getPropertyValue(v).trim() || fb;
        C = {
            bg:      g('--bg-primary',    '#0f1117'),
            surf:    g('--bg-secondary',  '#1c1f2e'),
            surf2:   g('--bg-tertiary',   '#252839'),
            bord:    g('--border-color',  '#3b3f54'),
            txt:     g('--text-primary',  '#e8eaed'),
            sub:     g('--text-secondary','#9aa0b4'),
            dim:     g('--text-muted',    '#5c6378'),
            acc:     '#6366f1',
            crit:    '#ef4444', critFill: 'rgba(239,68,68,0.10)',
            critGlow:'rgba(239,68,68,0.38)',
            done:    '#22c55e', late: '#f59e0b',
            mile:    '#f59e0b', prog: '#6366f1',
            risk0:   '#22c55e', risk50: '#f59e0b', risk100: '#ef4444',
            bottle:  '#f97316',
            eNorm:   'rgba(148,163,184,0.28)',
            eCrit:   '#ef4444', eHov: '#818cf8',
            eDim:    'rgba(148,163,184,0.08)',
        };
    }

    // ═══════════════════════════════════════════════════════════
    // SUMMARY PREDECESSOR RESOLUTION
    // ═══════════════════════════════════════════════════════════

    /**
     * Build a map: summaryUid → lastLeafDescendantUid
     * "Last leaf" = the final non-summary task inside the summary's
     * WBS subtree, in task-list order.
     */
    function _buildSummaryLastLeaf(allTasks) {
        const map = new Map();
        allTasks.forEach((t, idx) => {
            if (!t.summary) return;
            const level = t.outlineLevel || 1;
            let lastLeaf = null;
            for (let j = idx + 1; j < allTasks.length; j++) {
                if ((allTasks[j].outlineLevel || 1) <= level) break;
                if (!allTasks[j].summary) lastLeaf = allTasks[j];
            }
            if (lastLeaf) map.set(t.uid, lastLeaf.uid);
        });
        return map;
    }

    /**
     * For a given task, return the list of effective predecessor UIDs
     * that are actually present in the VISIBLE set.
     *
     * If a predecessor UID is a summary task (not visible), we
     * transparently replace it with the summary's last leaf descendant.
     */
    function _resolveEffectivePreds(task, visibleUids, summaryLastLeaf) {
        const resolved = [];
        const seen = new Set();
        (task.predecessors || []).forEach(pred => {
            let uid = pred.predecessorUID;
            if (!visibleUids.has(uid)) {
                // Try resolving through summary map
                uid = summaryLastLeaf.get(uid);
            }
            if (!uid || !visibleUids.has(uid)) return;
            if (uid === task.uid) return;
            if (seen.has(uid)) return;
            seen.add(uid);
            resolved.push({ ...pred, predecessorUID: uid });
        });
        return resolved;
    }

    // ═══════════════════════════════════════════════════════════
    // PUBLIC API
    // ═══════════════════════════════════════════════════════════
    function init(canvasEl) {
        cleanup();
        _canvas = canvasEl; _ctx = canvasEl.getContext('2d');
        _wrap   = canvasEl.parentElement;
        _dpr    = window.devicePixelRatio || 1;
        // Canvas is fixed to the wrapper viewport — panning is via mouse drag, not scroll
        if (_wrap) { _wrap.style.overflow = 'hidden'; _wrap.style.position = 'relative'; }
        _abortCtrl = new AbortController();
        const sig = _abortCtrl.signal;
        _canvas.addEventListener('wheel',      _onWheel,  { passive:false, signal:sig });
        _canvas.addEventListener('mousedown',  _onDown,   { signal:sig });
        _canvas.addEventListener('mousemove',  _onMove,   { signal:sig });
        _canvas.addEventListener('mouseup',    _onUp,     { signal:sig });
        _canvas.addEventListener('mouseleave', _onLeave,  { signal:sig });
        _canvas.addEventListener('dblclick',   _onDbl,    { signal:sig });
        _canvas.addEventListener('touchstart', _onTS,     { passive:false, signal:sig });
        _canvas.addEventListener('touchmove',  _onTM,     { passive:false, signal:sig });
        _canvas.addEventListener('touchend',   _onTE,     { signal:sig });
        document.addEventListener('keydown',   _onKey,    { signal:sig });
        window.addEventListener('resize', () => { _resize(); _draw(); }, { signal:sig });
        _clr();
    }

    /**
     * Main entry point — receives the FULL task list (incl. summaries).
     * Summaries are stored for predecessor resolution but excluded from layout.
     *
     * @param {Task[]} taskList - Full task list (may include summaries)
     * @param {object} [options]
     * @param {boolean|null} [options.dependenciesAvailable]
     *   Pass `false` when the import source could not provide dependency data
     *   (e.g. Dataverse unavailable).  The diagram will show an explanatory
     *   banner so users understand why all tasks appear unlinked.
     *   Pass `true` or omit to show the diagram normally.
     */
    function update(taskList, options = {}) {
        // Store the raw full list (including summaries) for predecessor resolution
        _allTasksRaw = taskList || [];

        // Update flag — prefer explicit option, fallback to array property set by CPM
        if (typeof options.dependenciesAvailable === 'boolean') {
            _dependenciesAvailable = options.dependenciesAvailable;
        } else if (typeof taskList._cpmDepsAvailable === 'boolean') {
            _dependenciesAvailable = taskList._cpmDepsAvailable;
        } else {
            _dependenciesAvailable = null; // unknown — no banner
        }

        // Store the earliest task start date as the project reference date
        _projectStartDate = null;
        if (_allTasksRaw.length > 0) {
            let minTs = Infinity;
            _allTasksRaw.forEach(t => {
                if (t.start) { const ts = new Date(t.start).getTime(); if (ts < minTs) minTs = ts; }
            });
            if (isFinite(minTs)) _projectStartDate = new Date(minTs);
        }

        // Visible set = non-summary + visible
        _allTasks = _allTasksRaw.filter(t => !t.summary && t.isVisible !== false);

        // Apply current filter
        _tasks = _applyFilter(_allTasks);

        _clr(); _buildMaps(); _layout(); _buildStats(); _resize(); _draw();
    }

    function _applyFilter(list) {
        if (_filterMode === 'critical') return list.filter(t => t.critical);
        if (_filterMode === 'late')     return list.filter(t => t.status === 'late');
        if (_filterMode === 'atrisk')   return list.filter(t => t.status === 'late' || t.status === 'at-risk' || t.critical);
        return list.slice();
    }

    function cleanup() {
        if (_abortCtrl) { _abortCtrl.abort(); _abortCtrl = null; }
        if (_rafId) { cancelAnimationFrame(_rafId); _rafId = null; }
        _gridPattern = null; _gridBg = '';
        _mmCanvas = null; _mmDirty = true;
        _clearTip();
    }

    function setMode(m)     { _mode = m; if (_tasks.length) { _layout(); _resize(); _draw(); setTimeout(() => _fit(), 60); } }
    function setFilter(f)   {
        _filterMode = f;
        _tasks = _applyFilter(_allTasks);
        _clr(); _buildMaps(); _layout(); _buildStats(); _resize(); _draw();
        setTimeout(() => _fit(), 60);
    }
    function setSearch(q)   { _searchQuery = (q||'').toLowerCase(); _buildHighlight(); _draw(); }
    function zoomIn()       { _zoom(_scale * 1.2); }
    function zoomOut()      { _zoom(_scale / 1.2); }
    function fitToScreen()  { _fit(); }
    function render()       { _draw(); }
    function getStats()     { return _statsCache; }

    // ═══════════════════════════════════════════════════════════
    // MAPS + STATS
    // ═══════════════════════════════════════════════════════════
    function _buildMaps() {
        _invalidateMinimap();
        _succMap.clear(); _predMap.clear();

        // Build summary-to-lastLeaf lookup using the full raw list
        const summaryLastLeaf = _buildSummaryLastLeaf(_allTasksRaw);

        // Full UID set (all non-summary visible tasks — before filter)
        const allUids     = new Set(_allTasks.map(t => t.uid));
        // Filtered visible set (e.g. only critical tasks when filter='critical')
        const visibleUids = new Set(_tasks.map(t => t.uid));

        // ── Pre-compute resolved preds for ALL tasks against the full set ──
        // These are cached as _allResolvedPreds and used for transitive edge resolution.
        _allTasks.forEach(t => {
            t._allResolvedPreds = _resolveEffectivePreds(t, allUids, summaryLastLeaf);
        });

        // Full pred map (all tasks) for transitive ancestor search
        const fullPredMap = new Map(_allTasks.map(t => [t.uid, t._allResolvedPreds]));

        if (_filterMode === 'critical' && visibleUids.size < allUids.size) {
            // ── Critical filter: trace transitive edges through non-critical tasks ──
            // If the chain is  A(crit) → B(non-crit) → C(crit), draw A→C directly.
            // This ensures the critical chain is connected even when intermediate tasks
            // are not on the critical path.
            _tasks.forEach(t => {
                const critPreds = [];
                const visited   = new Set();
                const findCritAncestors = uid => {
                    (fullPredMap.get(uid) || []).forEach(p => {
                        const pid = p.predecessorUID;
                        if (visited.has(pid)) return;
                        visited.add(pid);
                        if (visibleUids.has(pid)) {
                            critPreds.push({ ...p, predecessorUID: pid });
                        } else {
                            findCritAncestors(pid); // walk through non-critical
                        }
                    });
                };
                findCritAncestors(t.uid);
                t._resolvedPreds = critPreds;
                _succMap.set(t.uid, []);
                _predMap.set(t.uid, []);
            });
        } else {
            // ── Normal filter: resolve preds only within visible set ──
            _tasks.forEach(t => {
                t._resolvedPreds = _resolveEffectivePreds(t, visibleUids, summaryLastLeaf);
                _succMap.set(t.uid, []);
                _predMap.set(t.uid, []);
            });
        }

        _tasks.forEach(t => {
            t._resolvedPreds.forEach(p => {
                const pid = p.predecessorUID;
                if (_succMap.has(pid)) _succMap.get(pid).push(t.uid);
                if (_predMap.has(t.uid)) _predMap.get(t.uid).push(pid);
            });
        });
    }

    function _buildStats() {
        const n = _tasks.length;
        if (!n) { _statsCache = null; return; }

        const critTasks   = _tasks.filter(t => t.critical);
        const lateTasks   = _tasks.filter(t => t.status === 'late');
        const zeroFloat   = _tasks.filter(t => (t.totalFloat || 0) === 0);
        const bottlenecks = _tasks.filter(t => (_succMap.get(t.uid)||[]).length >= 3);
        const floatVals   = _tasks.map(t => t.totalFloat).filter(v => v != null && isFinite(v));
        const avgFloat    = floatVals.length ? Math.round(floatVals.reduce((a,b)=>a+b,0)/floatVals.length) : 0;

        // Critical path length: compute actual longest path through critical tasks
        // Sum durations of tasks with 0 float that form a connected chain.
        // Use the max EF of critical tasks minus min ES of critical tasks.
        let critPathLen = 0;
        if (critTasks.length > 0) {
            const critEFs = critTasks.map(t => t._ef || 0).filter(isFinite);
            const critESs = critTasks.map(t => t._es || 0).filter(isFinite);
            if (critEFs.length > 0 && critESs.length > 0) {
                critPathLen = Math.max(...critEFs) - Math.min(...critESs);
            }
        }

        _statsCache = {
            n, critCount: critTasks.length, lateCount: lateTasks.length,
            zeroFloatCount: zeroFloat.length, bottleneckCount: bottlenecks.length,
            avgFloat, critPathLen,
        };
        _renderStatsBar();
    }

    function _renderStatsBar() {
        const el = document.getElementById('ndStatsBar');
        if (!el || !_statsCache) return;
        const s = _statsCache;
        el.innerHTML = '';
        const items = [
            { label: 'Critical path',  val: s.critPathLen + 'd', color: s.critCount > 0 ? '#ef4444' : '' },
            { label: 'Critical tasks', val: s.critCount,         color: s.critCount > 0 ? '#ef4444' : '' },
            { label: 'Late tasks',     val: s.lateCount,         color: s.lateCount > 0 ? '#f59e0b' : '' },
            { label: 'Zero float',     val: s.zeroFloatCount,    color: s.zeroFloatCount > 0 ? '#f59e0b' : '' },
            { label: 'Bottlenecks',    val: s.bottleneckCount,   color: s.bottleneckCount > 0 ? '#f97316' : '' },
            { label: 'Avg float',      val: s.avgFloat + 'd',    color: '' },
        ];
        items.forEach(item => {
            const span = document.createElement('span'); span.className = 'nd-stat-item';
            const k = document.createElement('span'); k.className = 'nd-stat-key'; k.textContent = item.label;
            const v = document.createElement('span'); v.className = 'nd-stat-val'; v.textContent = item.val;
            if (item.color) v.style.color = item.color;
            span.appendChild(k); span.appendChild(v); el.appendChild(span);
        });

        // "Trace Critical Path" button — fits view to critical nodes only
        if (s.critCount > 0) {
            const btn = document.createElement('button');
            btn.textContent = '🔴 Trace Critical Path';
            btn.title = 'Fit view to critical tasks only (keyboard: C)';
            btn.style.cssText = [
                'margin-left:auto', 'padding:3px 10px', 'border-radius:5px', 'cursor:pointer',
                'font-size:0.7rem', 'font-weight:700', 'border:1.5px solid #ef4444',
                'background:rgba(239,68,68,0.13)', 'color:#ef4444',
                'transition:background 0.15s',
            ].join(';');
            btn.addEventListener('mouseenter', () => btn.style.background = 'rgba(239,68,68,0.25)');
            btn.addEventListener('mouseleave', () => btn.style.background = 'rgba(239,68,68,0.13)');
            btn.addEventListener('click', () => {
                // Switch to critical-only filter and fit to those nodes
                const wasFilter = _filterMode;
                if (_filterMode !== 'critical') {
                    _filterMode = 'critical';
                    _tasks = _applyFilter(_allTasks);
                    _clr(); _buildMaps(); _layout(); _buildStats(); _resize(); _draw();
                    btn.textContent = '◀ Show All';
                    btn.style.borderColor = '#818cf8';
                    btn.style.color = '#818cf8';
                    btn.style.background = 'rgba(129,140,248,0.13)';
                } else {
                    _filterMode = 'all';
                    _tasks = _applyFilter(_allTasks);
                    _clr(); _buildMaps(); _layout(); _buildStats(); _resize(); _draw();
                    btn.textContent = '🔴 Trace Critical Path';
                    btn.style.borderColor = '#ef4444';
                    btn.style.color = '#ef4444';
                    btn.style.background = 'rgba(239,68,68,0.13)';
                }
                setTimeout(() => _fit(), 80);
            });
            el.appendChild(btn);
        }
    }

    function _riskScore(task) {
        let r = 0;
        if (task.critical)                              r += 50;
        if (task.status === 'late')                     r += 30;
        if ((_succMap.get(task.uid)||[]).length >= 3)   r += 12;
        if ((task.percentComplete||0) < 20 && new Date(task.start) < new Date()) r += 8;
        return Math.min(100, r);
    }

    function _riskColor(score) {
        if (score >= 70) return C.risk100;
        if (score >= 40) return C.risk50;
        return score > 0 ? '#84cc16' : C.bord;
    }

    function _buildHighlight() {
        _highlightUids.clear();
        if (!_searchQuery) return;
        _tasks.forEach(t => {
            if (t.name.toLowerCase().includes(_searchQuery) ||
                (t.resourceNames||[]).join(' ').toLowerCase().includes(_searchQuery) ||
                String(t.uid).includes(_searchQuery))
                _highlightUids.add(t.uid);
        });
    }

    function _buildImpactChain(uid) {
        _impactSet.clear();
        const visited = new Set();
        const dfs = u => {
            if (visited.has(u)) return;
            visited.add(u); _impactSet.add(u);
            (_succMap.get(u) || []).forEach(dfs);
        };
        dfs(uid);
    }

    // ═══════════════════════════════════════════════════════════
    // LAYOUT  —  Sugiyama + gravity Y
    // ═══════════════════════════════════════════════════════════
    function _layout() {
        _nodeMap.clear(); _nodes = []; _edges = [];
        _invalidateMinimap();
        if (!_tasks.length) return;
        const D = DIMS[_mode];

        const uidIdx = new Map();
        _tasks.forEach((t,i) => uidIdx.set(t.uid, i));
        const n = _tasks.length;
        const succ = Array.from({length:n}, () => []);
        const pred = Array.from({length:n}, () => []);

        // Use RESOLVED predecessors (summary → leaf)
        _tasks.forEach((t,i) => {
            (t._resolvedPreds || []).forEach(p => {
                const pi = uidIdx.get(p.predecessorUID);
                if (pi != null && pi !== i) { succ[pi].push(i); pred[i].push(pi); }
            });
        });

        // Layer assignment (longest path from sources)
        const layer = new Array(n).fill(0);
        const vis   = new Array(n).fill(false);
        const hasAnyEdges = _tasks.some(t => (t._resolvedPreds||[]).length > 0);

        if (hasAnyEdges) {
            const dfsL = u => {
                if (vis[u]) return layer[u];
                vis[u] = true;
                pred[u].forEach(p => { layer[u] = Math.max(layer[u], dfsL(p) + 1); });
                return layer[u];
            };
            for (let i = 0; i < n; i++) dfsL(i);
        } else {
            // No dependencies → time-based columns
            const starts = _tasks.map(t => t.start ? new Date(t.start).getTime() : 0).filter(v => v > 0);
            if (starts.length > 0) {
                const minS = Math.min(...starts), maxS = Math.max(...starts);
                const range = maxS - minS || 1;
                const maxCols = Math.min(12, Math.max(4, Math.ceil(_tasks.length / 5)));
                _tasks.forEach((t, i) => {
                    const ts = t.start ? new Date(t.start).getTime() : minS;
                    layer[i] = Math.min(maxCols - 1, Math.floor(((ts - minS) / range) * (maxCols - 1)));
                });
            }
        }

        const maxL = Math.max(...layer, 0);
        const grps = Array.from({length: maxL + 1}, () => []);
        layer.forEach((lv, i) => grps[lv].push(i));

        // Barycenter crossing minimisation (6 passes)
        for (let pass = 0; pass < 6; pass++) {
            const fwd = pass % 2 === 0;
            const seq = fwd ? grps.slice(1) : [...grps].slice(0,-1).reverse();
            seq.forEach(grp => {
                const li = grps.indexOf(grp);
                const ref = fwd ? grps[li-1] : grps[li+1];
                if (!ref) return;
                const rpos = new Map(ref.map((x,j) => [x,j]));
                const bary = grp.map(i => {
                    const nb = (fwd ? pred[i] : succ[i]).filter(x => rpos.has(x));
                    return { i, score: nb.length ? nb.reduce((s,x) => s+rpos.get(x), 0)/nb.length : Infinity };
                });
                bary.sort((a,b) => a.score - b.score);
                bary.forEach(({i}, pos) => grp[pos] = i);
            });
        }

        // Gravity-based Y positioning
        const posY = new Float64Array(n).fill(-1);
        grps.forEach((grp, lv) => {
            const count = grp.length;
            const totalH = count * (D.h + D.gapY) - D.gapY;
            grp.forEach((ti, row) => { posY[ti] = row * (D.h + D.gapY) - totalH / 2; });

            if (lv > 0) {
                grp.forEach(ti => {
                    const preds = pred[ti].filter(pi => posY[pi] >= 0);
                    if (!preds.length) return;
                    const avgY = preds.reduce((s,pi) => s + posY[pi], 0) / preds.length;
                    posY[ti] = posY[ti] * 0.4 + avgY * 0.6;
                });
                grp.sort((a,b) => posY[a] - posY[b]);

                // Forward pass: push nodes down to maintain minimum vertical gap
                for (let i = 1; i < grp.length; i++) {
                    const minY = posY[grp[i-1]] + D.h + D.gapY;
                    if (posY[grp[i]] < minY) posY[grp[i]] = minY;
                }

                // Backward pass: push nodes up to maintain minimum vertical gap.
                // Without this, a forward-only push causes the whole column to drift
                // downward in dense layers, creating overlap when nodes are re-sorted
                // in subsequent iterations.
                for (let i = grp.length - 2; i >= 0; i--) {
                    const maxY = posY[grp[i+1]] - D.h - D.gapY;
                    if (posY[grp[i]] > maxY) posY[grp[i]] = maxY;
                }
            }
        });

        // Normalise to positive + padding
        const minY = Math.min(...Array.from(posY));
        grps.forEach((grp, lv) => {
            grp.forEach(ti => {
                const t  = _tasks[ti];
                const nd = {
                    task: t,
                    x: D.pad + lv * (D.w + D.gapX),
                    y: D.pad + (posY[ti] - minY),
                    w: D.w, h: D.h,
                    risk: _riskScore(t),
                    isBotl: (_succMap.get(t.uid)||[]).length >= 3,
                };
                _nodes.push(nd); _nodeMap.set(t.uid, nd);
            });
        });

        // Build edges from resolved preds
        _nodes.forEach(nd => {
            (nd.task._resolvedPreds || []).forEach(p => {
                const src = _nodeMap.get(p.predecessorUID);
                if (src) _edges.push({
                    from: src, to: nd,
                    isCrit: src.task.critical && nd.task.critical,
                    type: p.typeName || 'FS',
                });
            });
        });
    }

    // ═══════════════════════════════════════════════════════════
    // CANVAS  —  fixed to wrapper size; content scrolled via pan/zoom
    // ═══════════════════════════════════════════════════════════
    function _resize() {
        if (!_canvas || !_wrap) return;
        const pw = _wrap.clientWidth  || 800;
        const ph = _wrap.clientHeight || 600;
        const nw = Math.floor(pw * _dpr);
        const nh = Math.floor(ph * _dpr);
        // Only resize when dimensions actually change (avoids invalidating GPU texture)
        if (_canvas.width !== nw || _canvas.height !== nh) {
            _canvas.width  = nw; _canvas.height = nh;
            _canvas.style.width  = pw + 'px';
            _canvas.style.height = ph + 'px';
            _ctx.setTransform(_dpr, 0, 0, _dpr, 0, 0);
            _gridPattern = null; // invalidate grid pattern (new ctx backing)
            _mmCanvas    = null; // invalidate minimap cache
        }
    }

    // ═══════════════════════════════════════════════════════════
    // DRAW
    // ═══════════════════════════════════════════════════════════
    function _draw() {
        if (!_canvas || !_ctx) return;
        const cw = _canvas.width/_dpr, ch = _canvas.height/_dpr;
        _ctx.save();
        _ctx.clearRect(0,0,cw,ch);
        _ctx.fillStyle = C.bg; _ctx.fillRect(0,0,cw,ch);

        // Dot grid — use cached pattern for performance (avoids thousands of arc() calls)
        const gp = _getGridPattern();
        if (gp) { _ctx.fillStyle = gp; _ctx.fillRect(0, 0, cw, ch); }

        _ctx.translate(_panX, _panY); _ctx.scale(_scale, _scale);

        // ── Viewport culling ──────────────────────────────────
        // World-space visible bounds (with a shadow/glow margin)
        const margin = 40 / _scale;
        const vpX0 = (-_panX / _scale) - margin,  vpY0 = (-_panY / _scale) - margin;
        const vpX1 = vpX0 + cw / _scale + margin*2, vpY1 = vpY0 + ch / _scale + margin*2;

        const hasImpact = _impactSet.size > 0;
        const hasSearch = _highlightUids.size > 0;

        _edges.forEach(e => {
            // Skip if both endpoints are entirely outside viewport
            const ex0 = Math.min(e.from.x + e.from.w, e.to.x);
            const ex1 = Math.max(e.from.x + e.from.w, e.to.x + e.to.w);
            const ey0 = Math.min(e.from.y, e.to.y);
            const ey1 = Math.max(e.from.y + e.from.h, e.to.y + e.to.h);
            if (ex1 < vpX0 || ex0 > vpX1 || ey1 < vpY0 || ey0 > vpY1) return;
            const dimmed = (hasImpact && !(_impactSet.has(e.from.task.uid) && _impactSet.has(e.to.task.uid)))
                        || (hasSearch && !(_highlightUids.has(e.from.task.uid) || _highlightUids.has(e.to.task.uid)));
            _drawEdge(e, dimmed);
        });

        _nodes.forEach(nd => {
            // Skip nodes entirely outside viewport
            if (nd.x + nd.w < vpX0 || nd.x > vpX1 || nd.y + nd.h < vpY0 || nd.y > vpY1) return;
            const dimmed = (hasImpact && !_impactSet.has(nd.task.uid))
                        || (hasSearch && !_highlightUids.has(nd.task.uid));
            _drawNode(nd, dimmed);
        });

        _ctx.restore();
        if (_nodes.length > 0) _drawMinimap(cw, ch);
        else _drawEmpty(cw, ch);
        _drawLegend(cw, ch);

        // "No dependency data" banner — shown when the import source could not
        // provide predecessor links (e.g. Dataverse unavailable, Scenario B).
        // Only shown when there ARE nodes so it doesn't compete with _drawEmpty.
        if (_dependenciesAvailable === false && _nodes.length > 0) {
            _drawNoDepsBanner(cw);
        }

        // Critical path filter with no edges → show explanation overlay
        if (_filterMode === 'critical' && _edges.length === 0 && _nodes.length > 0) {
            _drawCritFilterNoDepsMsg(cw);
        }
    }

    // ── Edge ──────────────────────────────────────────────────
    function _drawEdge(e, dimmed) {
        const {from:s, to:t, isCrit, type} = e;
        const isHov = _hovUid && (s.task.uid===_hovUid || t.task.uid===_hovUid);
        const color = dimmed ? C.eDim : isHov ? C.eHov : isCrit ? C.eCrit : C.eNorm;
        // Critical edges are 3.5× thicker so the path is visible even when zoomed out
        const lw    = isHov ? 2.5 : isCrit ? 3.5 : 1.5;

        _ctx.save();
        _ctx.strokeStyle = color; _ctx.lineWidth = lw; _ctx.lineJoin = 'round';
        if (isCrit && !dimmed && _scale > 0.22) {
            // Strong double-pass glow for critical edges (skip at very low zoom)
            _ctx.shadowColor = C.critGlow; _ctx.shadowBlur = 18;
        }
        if (type !== 'FS') _ctx.setLineDash([5,4]);

        const x1=s.x+s.w, y1=Math.floor(s.y+s.h/2);
        const x2=t.x,     y2=Math.floor(t.y+t.h/2);
        const mx=Math.floor((x1+x2)/2);

        _ctx.beginPath();
        _ctx.moveTo(x1,y1); _ctx.lineTo(mx,y1); _ctx.lineTo(mx,y2); _ctx.lineTo(x2,y2);
        _ctx.stroke();

        _ctx.setLineDash([]); _ctx.shadowBlur=0; _ctx.fillStyle=color;
        // Larger arrowhead on critical edges
        const ah = isCrit && !dimmed ? 11 : 8;
        const av = isCrit && !dimmed ?  5 : 4;
        _ctx.beginPath(); _ctx.moveTo(x2,y2); _ctx.lineTo(x2-ah,y2-av); _ctx.lineTo(x2-ah,y2+av); _ctx.closePath(); _ctx.fill();

        if (type !== 'FS' && !dimmed) {
            _ctx.fillStyle = isHov ? C.eHov : C.dim;
            _ctx.font='bold 8px Inter,sans-serif'; _ctx.textAlign='center';
            _ctx.fillText(type, mx, Math.min(y1,y2)-4);
        }
        _ctx.restore();
    }

    // ── Node ──────────────────────────────────────────────────
    function _drawNode(nd, dimmed) {
        const {task:t, x, y, w, h, risk, isBotl} = nd;
        const D = DIMS[_mode];
        const pct    = t.percentComplete || 0;
        const isCrit = t.critical, isDone = pct >= 100;
        const isLate = t.status === 'late';
        const isMile = t.milestone;
        const isHov  = _hovUid === t.uid;
        const isSel  = _selUid === t.uid;
        const alpha  = dimmed ? 0.22 : 1;

        _ctx.save();
        _ctx.globalAlpha = alpha;

        // Skip expensive shadowBlur at very low zoom — nodes are too small to show it
        const _canGlow = _scale > 0.22;
        if (_canGlow && (isHov || isSel) && !dimmed) {
            _ctx.shadowColor = isCrit ? C.critGlow : 'rgba(99,102,241,0.55)';
            _ctx.shadowBlur  = isSel ? 24 : 16;
        } else if (_canGlow && isCrit && !dimmed) {
            // Stronger glow so critical nodes stand out even when zoomed far out
            _ctx.shadowColor = C.critGlow; _ctx.shadowBlur = 22;
        }

        if (isMile) {
            _ctx.fillStyle = 'rgba(245,158,11,0.12)';
            _diamond(x+w/2, y+h/2, w*0.42, h*0.42); _ctx.fill();
            _ctx.strokeStyle = C.mile; _ctx.lineWidth = 1.5;
            _diamond(x+w/2, y+h/2, w*0.42, h*0.42); _ctx.stroke();
            _ctx.shadowBlur=0;
            _ctx.fillStyle = C.mile; _ctx.font=`700 ${D.fs}px Inter,sans-serif`;
            _ctx.textAlign='center'; _ctx.textBaseline='middle';
            const mn = t.name.length>16 ? t.name.slice(0,14)+'…' : t.name;
            _ctx.fillText('⭐ '+mn, x+w/2, y+h/2);
            _ctx.restore(); return;
        }

        // Node body fill
        _ctx.fillStyle = isCrit ? 'rgba(239,68,68,0.18)' : isDone ? 'rgba(34,197,94,0.07)' : isLate ? 'rgba(245,158,11,0.07)' : C.surf;
        _rr(x, y, w, h, 8); _ctx.fill();

        // Left risk-colour strip — simple 4px rect (was a complex arcTo path)
        const riskCol = _riskColor(risk);
        _ctx.fillStyle = riskCol;
        _ctx.fillRect(x, y + 8, 4, h - 16);

        // Border — shadow is cleared before stroke so it doesn't double-glow
        _ctx.shadowBlur = 0;
        _ctx.strokeStyle = (isHov||isSel) ? (isCrit?C.crit:C.acc) : isCrit ? C.crit : C.bord;
        _ctx.lineWidth   = (isHov||isSel) ? 2.5 : isCrit ? 2.5 : 1;
        _rr(x, y, w, h, 8); _ctx.stroke();

        if (isBotl) {
            _ctx.fillStyle = C.bottle;
            _ctx.font = `bold 9px Inter,sans-serif`; _ctx.textAlign='right'; _ctx.textBaseline='top';
            _ctx.fillText('⚡', x+w-5, y+3);
        }

        if (_mode === 'micro')        _drawMicro(nd, D);
        else if (_mode === 'compact') _drawCompact(nd, D);
        else                          _drawNormal(nd, D);

        _ctx.restore();
    }

    /**
     * Convert a CPM day-offset to a short date string (e.g. "Apr 26").
     * Falls back to showing the raw number if no project start date is known.
     * NOTE: kept for backward-compat; prefer _fmtDate for direct Date values.
     */
    function _dayToDate(days) {
        if (!isFinite(days) || days == null) return '—';
        if (_projectStartDate) {
            const d = new Date(_projectStartDate.getTime() + Math.round(days) * 86400000);
            return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
        }
        return 'd' + Math.round(days);
    }

    /**
     * Format an actual Date (or date-like value) to "Apr 26" style.
     * Used in PERT boxes and tooltip to display real task start/finish dates.
     */
    function _fmtDate(d) {
        if (!d) return '—';
        try {
            const dt = new Date(d);
            if (isNaN(dt.getTime())) return '—';
            return dt.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
        } catch(e) { return '—'; }
    }

    /**
     * Add `floatDays` calendar days to a Date and format the result.
     * Used to compute LS = start + float  and  LF = finish + float.
     */
    function _addDaysAndFmt(d, floatDays) {
        if (!d || !isFinite(floatDays)) return '—';
        try {
            const dt = new Date(new Date(d).getTime() + Math.round(floatDays) * 86400000);
            return _fmtDate(dt);
        } catch(e) { return '—'; }
    }

    function _drawNormal(nd, D) {
        const {task:t, x, y, w, h} = nd;
        const pct = t.percentComplete || 0;
        const rH  = Math.floor(h/3);
        const tx  = x+12;

        _ctx.fillStyle='rgba(255,255,255,0.04)';
        _ctx.fillRect(x+4,y+rH,w-8,1); _ctx.fillRect(x+4,y+rH*2,w-8,1);

        // Row 1: ES (actual start) | Task name | EF (actual finish)
        // Use real task dates to avoid off-by-one from calendar-day offset conversion
        const esV = _fmtDate(t.start);
        const efV = _fmtDate(t.finish);
        _ctx.font=`400 7.5px Inter,sans-serif`; _ctx.fillStyle=C.dim;
        _ctx.textAlign='left';  _ctx.textBaseline='top'; _ctx.fillText('ES '+esV, tx, y+5);
        _ctx.textAlign='right'; _ctx.fillText('EF '+efV, x+w-6, y+5);
        _ctx.fillStyle = t.critical ? C.crit : C.txt;
        _ctx.font=`600 ${D.fs}px Inter,sans-serif`;
        _ctx.textAlign='center'; _ctx.textBaseline='middle';
        const nm = t.name.length>19 ? t.name.slice(0,17)+'…' : t.name;
        _ctx.fillText(nm, x+w/2, y+rH/2);

        // Row 2: progress bar + % + duration
        const bY=y+rH+7, bW=Math.floor(w*0.54), bH=5;
        _ctx.fillStyle='rgba(255,255,255,0.06)'; _rrp(tx,bY,bW,bH,3); _ctx.fill();
        if (pct>0) {
            _ctx.fillStyle = pct>=100?C.done:t.critical?C.crit:C.prog;
            _rrp(tx,bY,Math.max(2,bW*pct/100),bH,3); _ctx.fill();
        }
        _ctx.fillStyle=C.sub; _ctx.font=`600 ${D.fs-1}px Inter,sans-serif`;
        _ctx.textAlign='left'; _ctx.textBaseline='middle';
        _ctx.fillText(pct+'%', tx+bW+4, bY+2.5);
        _ctx.fillStyle=C.dim; _ctx.textAlign='right';
        _ctx.fillText((t.durationDays||0)+'d', x+w-6, bY+2.5);

        // Row 3: LS | TF | LF
        // LS = task.start + float days,  LF = task.finish + float days
        const r3=y+rH*2+5;
        const tf = t.totalFloat!=null&&isFinite(t.totalFloat)?Math.round(t.totalFloat):null;
        const lsV = tf!==null ? _addDaysAndFmt(t.start, tf)  : '—';
        const lfV = tf!==null ? _addDaysAndFmt(t.finish, tf) : '—';
        _ctx.font=`400 7.5px Inter,sans-serif`; _ctx.fillStyle=C.dim;
        _ctx.textAlign='left';  _ctx.textBaseline='top'; _ctx.fillText('LS '+lsV, tx, r3);
        _ctx.textAlign='right'; _ctx.fillText('LF '+lfV, x+w-6, r3);
        if (tf!==null) {
            _ctx.fillStyle = tf===0?C.crit:tf<=2?C.late:C.dim;
            _ctx.textAlign='center'; _ctx.fillText('TF '+(tf===0?'0 ⚠':tf+'d'), x+w/2, r3);
        }
        if (tf !== null && tf > 0) {
            const maxF=20, fW=Math.min(tf/maxF,1)*(w-20);
            _ctx.fillStyle='rgba(99,102,241,0.18)';
            _ctx.fillRect(x+10, r3+11, fW, 3);
        }
    }

    function _drawCompact(nd, D) {
        const {task:t, x, y, w, h} = nd;
        const pct = t.percentComplete||0;
        const nm  = t.name.length>16 ? t.name.slice(0,14)+'…' : t.name;
        _ctx.fillStyle = t.critical?C.crit:C.txt;
        _ctx.font=`600 ${D.fs}px Inter,sans-serif`;
        _ctx.textAlign='left'; _ctx.textBaseline='top';
        _ctx.fillText(nm, x+11, y+6);
        _ctx.fillStyle=C.sub; _ctx.font=`400 ${D.fs-1}px Inter,sans-serif`;
        _ctx.textBaseline='bottom';
        _ctx.fillText((t.durationDays||0)+'d · '+pct+'%', x+11, y+h-5);
        _ctx.fillStyle='rgba(255,255,255,0.06)'; _ctx.fillRect(x+4, y+h-3, w-8, 3);
        if (pct>0) {
            _ctx.fillStyle=pct>=100?C.done:t.critical?C.crit:C.prog;
            _ctx.fillRect(x+4, y+h-3, Math.max(2,(w-8)*pct/100), 3);
        }
    }

    function _drawMicro(nd, D) {
        const {task:t, x, y, w, h} = nd;
        const pct = t.percentComplete||0;
        const nm  = t.name.length>12 ? t.name.slice(0,11)+'…' : t.name;
        _ctx.fillStyle = t.critical?C.crit:t.status==='late'?C.late:C.txt;
        _ctx.font=`500 ${D.fs}px Inter,sans-serif`;
        _ctx.textAlign='left'; _ctx.textBaseline='middle';
        _ctx.fillText(nm, x+8, y+h/2);
        _ctx.fillStyle='rgba(255,255,255,0.06)'; _ctx.fillRect(x+4, y+h-2, w-8, 2);
        if (pct>0) { _ctx.fillStyle=pct>=100?C.done:C.prog; _ctx.fillRect(x+4, y+h-2, Math.max(1,(w-8)*pct/100), 2); }
    }

    // ── Minimap ───────────────────────────────────────────────
    // Node dots are cached on an offscreen canvas (_mmCanvas).
    // Only the viewport rectangle is drawn live (changes every pan/zoom).
    let _mmBounds = { x0:0, x1:1, y0:0, y1:1, sc:1 }; // cached world bounds

    function _rebuildMinimapCache() {
        if (!_nodes.length) return;
        let x0=Infinity,x1=-Infinity,y0=Infinity,y1=-Infinity;
        _nodes.forEach(n => { x0=Math.min(x0,n.x); x1=Math.max(x1,n.x+n.w); y0=Math.min(y0,n.y); y1=Math.max(y1,n.y+n.h); });
        const dw=x1-x0||1, dh=y1-y0||1;
        const sc=Math.min((MM.w-6)/dw,(MM.h-6)/dh, 0.25);
        _mmBounds = { x0, x1, y0, y1, sc };

        if (!_mmCanvas) { _mmCanvas = document.createElement('canvas'); }
        _mmCanvas.width = MM.w; _mmCanvas.height = MM.h;
        const mc = _mmCanvas.getContext('2d');

        // Background pill
        mc.clearRect(0, 0, MM.w, MM.h);
        mc.fillStyle = 'rgba(8,9,16,0.92)';
        mc.beginPath(); mc.roundRect(0, 0, MM.w, MM.h, 7); mc.fill();
        mc.strokeStyle = 'rgba(255,255,255,0.07)'; mc.lineWidth = 1;
        mc.beginPath(); mc.roundRect(0, 0, MM.w, MM.h, 7); mc.stroke();

        // Node dots
        mc.save(); mc.beginPath(); mc.rect(1, 1, MM.w-2, MM.h-2); mc.clip();
        _nodes.forEach(n => {
            const nw=Math.max(3, n.w*sc), nh=Math.max(2, n.h*sc);
            const nx=3+(n.x-x0)*sc,      ny=3+(n.y-y0)*sc;
            mc.globalAlpha = 0.7;
            mc.fillStyle = n.task.critical?C.crit:n.task.percentComplete>=100?C.done:n.task.status==='late'?C.late:C.acc;
            mc.fillRect(nx, ny, nw, nh);
        });
        mc.restore();
        _mmDirty = false;
    }

    function _drawMinimap(cw, ch) {
        if (!_nodes.length) return;
        if (_mmDirty || !_mmCanvas) _rebuildMinimapCache();

        const mx = cw - MM.w - MM.pad, my = ch - MM.h - MM.pad;
        _ctx.drawImage(_mmCanvas, mx, my);

        // Viewport indicator — drawn live since it changes every pan/zoom
        const {x0, y0, sc} = _mmBounds;
        const vw = (_wrap.clientWidth  || 800) / _scale;
        const vh = (_wrap.clientHeight || 600) / _scale;
        const vx = mx+3+(-_panX/_scale-x0)*sc;
        const vy = my+3+(-_panY/_scale-y0)*sc;
        _ctx.strokeStyle = 'rgba(255,255,255,0.65)'; _ctx.lineWidth = 1.5;
        _ctx.strokeRect(Math.round(vx), Math.round(vy), Math.round(vw*sc), Math.round(vh*sc));
    }

    // ── Legend ────────────────────────────────────────────────
    function _drawLegend(cw, ch) {
        const items = [
            [C.crit,'Critical'],[C.done,'Complete'],[C.late,'Late/At-Risk'],
            [C.acc,'Normal'],[C.bottle,'Bottleneck ⚡'],
        ];
        let lx = Math.floor(cw/2 - items.length*38);
        const ly = ch - 18;
        items.forEach(([col, lbl]) => {
            _ctx.fillStyle=col; _ctx.beginPath(); _ctx.arc(lx+5,ly-1,4,0,Math.PI*2); _ctx.fill();
            _ctx.fillStyle=C.dim; _ctx.font='9px Inter,sans-serif';
            _ctx.textAlign='left'; _ctx.textBaseline='middle';
            _ctx.fillText(lbl, lx+13, ly);
            lx += lbl.length*5.3+18;
        });
    }

    function _drawEmpty(cw, ch) {
        _ctx.fillStyle=C.dim; _ctx.font='14px Inter,sans-serif';
        _ctx.textAlign='center'; _ctx.textBaseline='middle';
        _ctx.fillText('No tasks available for Network view', cw/2, ch/2-16);
        _ctx.font='11px Inter,sans-serif';
        _ctx.fillText('Import a project with tasks to see the Network / PERT diagram', cw/2, ch/2+10);
    }

    /**
     * Draw an overlay when Critical Path filter is active but no dependency edges
     * exist (either because deps unavailable or this filter has no linked tasks).
     * Shown at top of canvas in fixed position.
     */
    function _drawCritFilterNoDepsMsg(cw) {
        _ctx.save();
        const bH = 52, pad = 14, top = (_dependenciesAvailable === false) ? 62 : 14;
        _ctx.fillStyle = 'rgba(239,68,68,0.10)';
        _ctx.strokeStyle = 'rgba(239,68,68,0.45)';
        _ctx.lineWidth = 1;
        _rrp(pad, top, cw - pad * 2, bH, 8); _ctx.fill();
        _rrp(pad, top, cw - pad * 2, bH, 8); _ctx.stroke();

        _ctx.fillStyle = '#ef4444';
        _ctx.font = `600 11.5px Inter,sans-serif`;
        _ctx.textAlign = 'center';
        _ctx.textBaseline = 'middle';
        _ctx.fillText(
            '🔴 Critical Path — showing tasks with zero float (no predecessor links found)',
            cw / 2, top + 16
        );
        _ctx.fillStyle = 'rgba(233,233,233,0.55)';
        _ctx.font = `400 10px Inter,sans-serif`;
        _ctx.fillText(
            'Import an MS Project .xml file with predecessor data to see the connected critical path',
            cw / 2, top + 34
        );
        _ctx.restore();
    }

    /**
     * Draw a "no dependency data" warning banner anchored to the top of the canvas.
     * Called by _draw() when _dependenciesAvailable === false and nodes are present.
     *
     * The banner is drawn AFTER the pan/scale transform is restored so it stays
     * in a fixed position regardless of zoom level.
     */
    function _drawNoDepsBanner(cw) {
        const bH = 38, pad = 14;
        // Background pill
        _ctx.save();
        _ctx.fillStyle = 'rgba(245,158,11,0.13)';
        _ctx.strokeStyle = 'rgba(245,158,11,0.55)';
        _ctx.lineWidth = 1;
        _rrp(pad, pad, cw - pad * 2, bH, 8); _ctx.fill();
        _rrp(pad, pad, cw - pad * 2, bH, 8); _ctx.stroke();

        // Icon + text
        _ctx.fillStyle = '#f59e0b';
        _ctx.font = `600 11.5px Inter,sans-serif`;
        _ctx.textAlign = 'center';
        _ctx.textBaseline = 'middle';
        _ctx.fillText(
            '⚠  الشبكة تُظهر مهام معزولة — بيانات الاعتماديات غير متاحة من المصدر (Dataverse / Project Ops)',
            cw / 2, pad + bH / 2
        );
        _ctx.restore();
    }

    // ── PNG Export (tight-crop, up to 4K) ────────────────────
    function exportPNG() {
        if (!_nodes.length) return;

        let x0=Infinity, y0=Infinity, x1=-Infinity, y1=-Infinity;
        _nodes.forEach(n => {
            x0=Math.min(x0,n.x); y0=Math.min(y0,n.y);
            x1=Math.max(x1,n.x+n.w); y1=Math.max(y1,n.y+n.h);
        });

        const PAD  = 40;
        const bW   = x1 - x0 + PAD * 2;
        const bH   = y1 - y0 + PAD * 2;
        const scale = Math.min(1, 3840 / bW, 2160 / bH);
        const outW = Math.ceil(bW * scale);
        const outH = Math.ceil(bH * scale);

        const off  = document.createElement('canvas');
        off.width  = outW; off.height = outH;
        const ctx  = off.getContext('2d');

        ctx.fillStyle = C.bg || '#0f1117';
        ctx.fillRect(0, 0, outW, outH);

        // Dot grid for export
        const exportOff = document.createElement('canvas'); exportOff.width = 28; exportOff.height = 28;
        const eo = exportOff.getContext('2d');
        eo.fillStyle = C.bg || '#0f1117'; eo.fillRect(0,0,28,28);
        eo.fillStyle = 'rgba(255,255,255,0.025)'; eo.beginPath(); eo.arc(0,0,1,0,Math.PI*2); eo.fill();
        const ep = ctx.createPattern(exportOff, 'repeat');
        if (ep) { ctx.fillStyle = ep; ctx.fillRect(0, 0, outW, outH); }

        ctx.save();
        ctx.translate((-x0 + PAD) * scale, (-y0 + PAD) * scale);
        ctx.scale(scale, scale);

        const realCtx = _ctx; _ctx = ctx;
        _edges.forEach(e => _drawEdge(e, false));
        _nodes.forEach(n => _drawNode(n, false));
        _ctx = realCtx;

        ctx.restore();

        const a = document.createElement('a');
        a.download = 'network-diagram.png';
        a.href = off.toDataURL('image/png');
        a.click();
    }

    // ═══════════════════════════════════════════════════════════
    // INTERACTION
    // ═══════════════════════════════════════════════════════════
    function _hit(wx, wy) {
        for (let i=_nodes.length-1; i>=0; i--) {
            const n=_nodes[i];
            if (wx>=n.x && wx<=n.x+n.w && wy>=n.y && wy<=n.y+n.h) return n;
        }
        return null;
    }
    function _xy(e)  { const r=_canvas.getBoundingClientRect(); return { x:(e.clientX-r.left-_panX)/_scale, y:(e.clientY-r.top-_panY)/_scale }; }
    function _cxy(e) { const r=_canvas.getBoundingClientRect(); return { cx:e.clientX-r.left, cy:e.clientY-r.top }; }

    function _onWheel(e) {
        e.preventDefault();
        const {cx,cy}=_cxy(e);
        const ns=Math.min(3,Math.max(0.15,_scale*(e.deltaY>0?0.88:1.14)));
        _panX=cx-(cx-_panX)*(ns/_scale); _panY=cy-(cy-_panY)*(ns/_scale);
        _scale=ns; _schedDraw();
    }
    function _onDown(e) {
        _clearTip();
        const {x,y}=_xy(e); const hit=_hit(x,y);
        if (hit) {
            if (_selUid===hit.task.uid) { _selUid=null; _impactSet.clear(); }
            else { _selUid=hit.task.uid; _buildImpactChain(hit.task.uid); }
            _draw(); return;
        }
        _isPanning=true; _pStart={x:e.clientX-_panX, y:e.clientY-_panY};
        _canvas.style.cursor='grabbing';
    }
    function _onMove(e) {
        const {x,y}=_xy(e); const hit=_hit(x,y);
        const newHov=hit?hit.task.uid:null;
        if (newHov!==_hovUid) {
            _hovUid=newHov;
            _canvas.style.cursor=hit?'pointer':(_isPanning?'grabbing':'grab');
            _schedDraw();
            if (hit) _showTip(e.clientX-_canvas.getBoundingClientRect().left, e.clientY-_canvas.getBoundingClientRect().top, hit);
            else _clearTip();
        }
        if (_isPanning) { _panX=e.clientX-_pStart.x; _panY=e.clientY-_pStart.y; _schedDraw(); }
    }
    function _onUp()    { _isPanning=false; _canvas.style.cursor=_hovUid?'pointer':'grab'; }
    function _onLeave() { _isPanning=false; _hovUid=null; _clearTip(); _schedDraw(); }
    function _onDbl(e) {
        const {x,y}=_xy(e); const hit=_hit(x,y);
        if (hit) _canvas.dispatchEvent(new CustomEvent('nodeDoubleClick',{bubbles:true,detail:{task:hit.task}}));
    }
    function _onKey(e) {
        if (e.target.tagName==='INPUT') return;
        if (e.key==='+' || e.key==='=') zoomIn();
        else if (e.key==='-') zoomOut();
        else if (e.key==='f' || e.key==='F') fitToScreen();
        else if (e.key==='c' || e.key==='C') { _filterMode=_filterMode==='critical'?'all':'critical'; setFilter(_filterMode); }
        else if (e.key==='Escape') { _selUid=null; _impactSet.clear(); _highlightUids.clear(); _searchQuery=''; const si=document.getElementById('ndSearch'); if(si) si.value=''; _draw(); }
    }

    // ── Tooltip ───────────────────────────────────────────────
    function _showTip(cx, cy, nd) {
        _clearTip();
        const {task:t, risk, isBotl} = nd;
        const el = document.createElement('div'); el.className='nd-tooltip';

        const title=document.createElement('div'); title.className='nd-tip-title';
        title.textContent=(t.critical?'🔴 ':t.status==='late'?'⏰ ':'')+t.name;
        el.appendChild(title);

        const riskBar=document.createElement('div');
        riskBar.style.cssText=`height:3px;border-radius:2px;background:${_riskColor(risk)};width:${risk}%;margin-bottom:8px`;
        el.appendChild(riskBar);

        const riskLabel=document.createElement('div');
        riskLabel.textContent=`Risk Score: ${risk}/100 ${risk>=70?'🔴 High':risk>=40?'🟡 Medium':'🟢 Low'}`;
        riskLabel.style.cssText='font-size:0.7rem;margin-bottom:6px;font-weight:600;color:'+_riskColor(risk);
        el.appendChild(riskLabel);

        const tf = t.totalFloat!=null&&isFinite(t.totalFloat)?Math.round(t.totalFloat):null;
        const sucCount = (_succMap.get(t.uid)||[]).length;
        const preCount = (_predMap.get(t.uid)||[]).length;
        // Use actual task dates for ES/EF; compute LS/LF as start/finish + float
        const esStr = _fmtDate(t.start);
        const efStr = _fmtDate(t.finish);
        const lsStr = tf!==null ? _addDaysAndFmt(t.start, tf)  : '—';
        const lfStr = tf!==null ? _addDaysAndFmt(t.finish, tf) : '—';
        const rows = [
            ['Duration',    (t.durationDays||0)+'d'],
            ['Progress',    (t.percentComplete||0)+'%'],
            ['Float',       tf!==null?(tf+'d'+(tf===0?' ⚠ Critical path':'')):'—'],
            ['ES → EF',     esStr+' → '+efStr],
            ['LS → LF',     lsStr+' → '+lfStr],
            ['Successors',  sucCount+(isBotl?' ⚡ Bottleneck':'')],
            ['Predecessors',preCount],
            ['Resource',    (t.resourceNames||[]).join(', ')||'—'],
            ['Status',      (t.statusIcon||'')+' '+(t.status||'normal')],
        ];
        rows.forEach(([k,v]) => {
            const row=document.createElement('div'); row.className='nd-tip-row';
            const ke=document.createElement('span'); ke.className='nd-tip-key'; ke.textContent=k;
            const ve=document.createElement('span'); ve.className='nd-tip-val'; ve.textContent=v;
            row.appendChild(ke); row.appendChild(ve); el.appendChild(row);
        });

        if (sucCount>0) {
            const hint=document.createElement('div');
            hint.textContent='💡 Click to highlight downstream impact chain';
            hint.style.cssText='font-size:0.65rem;color:var(--text-muted);margin-top:8px;border-top:1px solid var(--border-color);padding-top:6px';
            el.appendChild(hint);
        }

        const parent=_canvas.closest('.network-canvas-wrap')||document.body;
        parent.style.position='relative';
        el.style.left=Math.min(cx+16,(parent.clientWidth||600)-250)+'px';
        el.style.top=Math.max(cy-30,8)+'px';
        parent.appendChild(el); _tipEl=el;
    }
    function _clearTip() { if (_tipEl) { _tipEl.remove(); _tipEl=null; } }

    // ── Touch ─────────────────────────────────────────────────
    function _onTS(e) { e.preventDefault(); _touches=[...e.touches]; if (_touches.length===2) _touchDist=Math.hypot(_touches[0].clientX-_touches[1].clientX,_touches[0].clientY-_touches[1].clientY); else { _isPanning=true; _pStart={x:_touches[0].clientX-_panX,y:_touches[0].clientY-_panY}; } }
    function _onTM(e) { e.preventDefault(); _touches=[...e.touches]; if (_touches.length===2) { const d=Math.hypot(_touches[0].clientX-_touches[1].clientX,_touches[0].clientY-_touches[1].clientY); _zoom(_scale*(d/(_touchDist||1))); _touchDist=d; } else if (_isPanning) { _panX=_touches[0].clientX-_pStart.x; _panY=_touches[0].clientY-_pStart.y; _schedDraw(); } }
    function _onTE()  { _isPanning=false; _touches=[]; }

    // ── Zoom / Fit ────────────────────────────────────────────
    function _zoom(ns) { _scale=Math.min(3,Math.max(0.12,ns)); _schedDraw(); }
    function _fit() {
        if (!_nodes.length||!_wrap) return;
        let x0=Infinity,x1=-Infinity,y0=Infinity,y1=-Infinity;
        _nodes.forEach(n => { x0=Math.min(x0,n.x); x1=Math.max(x1,n.x+n.w); y0=Math.min(y0,n.y); y1=Math.max(y1,n.y+n.h); });
        const dw=x1-x0+60, dh=y1-y0+60;
        const ww=_wrap.clientWidth||800, wh=_wrap.clientHeight||600;
        _scale=Math.min(ww/dw, wh/dh, 1.5);
        _panX=Math.floor((ww-dw*_scale)/2)-x0*_scale+30*_scale;
        _panY=Math.floor((wh-dh*_scale)/2)-y0*_scale+30*_scale;
        _draw();
    }

    // ── Shape helpers ─────────────────────────────────────────
    function _rr(x,y,w,h,r)  { _ctx.beginPath(); _ctx.moveTo(x+r,y); _ctx.arcTo(x+w,y,x+w,y+h,r); _ctx.arcTo(x+w,y+h,x,y+h,r); _ctx.arcTo(x,y+h,x,y,r); _ctx.arcTo(x,y,x+w,y,r); _ctx.closePath(); }
    function _rrp(x,y,w,h,r) { _ctx.beginPath(); _ctx.moveTo(x+r,y); _ctx.arcTo(x+w,y,x+w,y+r,r); _ctx.arcTo(x+w,y+h,x+w-r,y+h,r); _ctx.arcTo(x,y+h,x,y+h-r,r); _ctx.arcTo(x,y,x+r,y,r); _ctx.closePath(); }
    function _diamond(cx,cy,rw,rh) { _ctx.beginPath(); _ctx.moveTo(cx,cy-rh); _ctx.lineTo(cx+rw,cy); _ctx.lineTo(cx,cy+rh); _ctx.lineTo(cx-rw,cy); _ctx.closePath(); }

    export const NetworkDiagram = { init, update, render, cleanup, setMode, setFilter, setSearch, zoomIn, zoomOut, fitToScreen, exportPNG, getStats };
