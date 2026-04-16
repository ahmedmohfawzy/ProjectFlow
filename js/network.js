/**
 * ProjectFlow™ © 2026 Ahmed M. Fawzy. All Rights Reserved.
 */
/**
 * ═══════════════════════════════════════════════════════════════
 * Network Diagram v3  —  PERT / CPM  +  Decision Intelligence
 *
 *  Layout    : Sugiyama + gravity-based Y positioning
 *  Nodes     : compact PERT boxes  (3 sizes: normal/compact/micro)
 *  Edges     : orthogonal elbow routing, no node overlap
 *  Highlight : Critical path glow, hover ripple, impact chain
 *  Filters   : Critical-only | Late | At-Risk | All
 *  Search    : Find & highlight task by name
 *  Risk Halo : Per-node risk score (0–100) via colour border
 *  Impact    : Click → dim all non-downstream tasks
 *  Bottleneck: ⚡ badge on nodes with ≥3 successors
 *  Stats Bar : Critical path length, bottlenecks, avg float
 *  Minimap   : Overview with viewport rect
 *  Keyboard  : +/- zoom, F fit, C critical-only, Esc deselect
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
    let _tasks = [], _nodes = [], _edges = [];
    let _nodeMap = new Map();           // uid → node
    let _succMap = new Map();           // uid → [uid]  successors
    let _predMap = new Map();           // uid → [uid]  predecessors
    let _impactSet = new Set();         // uids in impact chain of selected
    let _statsCache = null;

    let _panX = 0, _panY = 0, _scale = 1;
    let _isPanning = false, _pStart = { x:0, y:0 };
    let _hovUid = null, _selUid = null;
    let _filterMode = 'all';           // 'all' | 'critical' | 'late' | 'atrisk'
    let _searchQuery = '';
    let _highlightUids = new Set();    // from search
    let _abortCtrl = null;
    let _dpr = 1;
    let _touches = [], _touchDist = 0;
    let _tipEl = null;

    // Minimap
    const MM = { w: 148, h: 88, pad: 8 };

    // ── Colors ─────────────────────────────────────────────────
    let C = {};
    function _clr() {
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
            crit:    '#ef4444',   critFill: 'rgba(239,68,68,0.10)',
            critGlow:'rgba(239,68,68,0.38)',
            done:    '#22c55e',   late: '#f59e0b',
            mile:    '#f59e0b',   prog: '#6366f1',
            risk0:   '#22c55e',   // low risk
            risk50:  '#f59e0b',   // medium
            risk100: '#ef4444',   // high
            bottle:  '#f97316',   // bottleneck orange
            eNorm:   'rgba(148,163,184,0.28)',
            eCrit:   '#ef4444',   eHov: '#818cf8',
            eDim:    'rgba(148,163,184,0.08)',
        };
    }

    // ═══════════════════════════════════════════════════════════
    // PUBLIC API
    // ═══════════════════════════════════════════════════════════
    function init(canvasEl) {
        cleanup();
        _canvas = canvasEl; _ctx = canvasEl.getContext('2d');
        _wrap   = canvasEl.parentElement;
        _dpr    = window.devicePixelRatio || 1;
        _abortCtrl = new AbortController();
        const sig = _abortCtrl.signal;
        _canvas.addEventListener('wheel',       _onWheel,     { passive:false, signal:sig });
        _canvas.addEventListener('mousedown',   _onDown,      { signal:sig });
        _canvas.addEventListener('mousemove',   _onMove,      { signal:sig });
        _canvas.addEventListener('mouseup',     _onUp,        { signal:sig });
        _canvas.addEventListener('mouseleave',  _onLeave,     { signal:sig });
        _canvas.addEventListener('dblclick',    _onDbl,       { signal:sig });
        _canvas.addEventListener('touchstart',  _onTS,        { passive:false, signal:sig });
        _canvas.addEventListener('touchmove',   _onTM,        { passive:false, signal:sig });
        _canvas.addEventListener('touchend',    _onTE,        { signal:sig });
        document.addEventListener('keydown',    _onKey,       { signal:sig });
        window.addEventListener('resize', () => { _resize(); _draw(); }, { signal:sig });
        _clr();
    }

    function update(taskList) {
        _tasks = taskList.filter(t => !t.summary && t.isVisible !== false);
        if (_filterMode === 'critical') _tasks = _tasks.filter(t => t.critical);
        else if (_filterMode === 'late') _tasks = _tasks.filter(t => t.status === 'late');
        else if (_filterMode === 'atrisk') _tasks = _tasks.filter(t => t.status === 'late' || t.status === 'at-risk' || t.critical);
        _clr(); _buildMaps(); _layout(); _buildStats(); _resize(); _draw();
    }

    function cleanup() {
        if (_abortCtrl) { _abortCtrl.abort(); _abortCtrl = null; }
        _clearTip();
    }

    function setMode(m)          { _mode = m; if (_tasks.length) { _layout(); _resize(); _draw(); } }
    function setFilter(f)        { _filterMode = f; update(_tasks); }
    function setSearch(q)        { _searchQuery = (q||'').toLowerCase(); _buildHighlight(); _draw(); }
    function zoomIn()            { _zoom(_scale * 1.2); }
    function zoomOut()           { _zoom(_scale / 1.2); }
    function fitToScreen()       { _fit(); }
    function exportPNG()         { const a = document.createElement('a'); a.download = 'network.png'; a.href = _canvas.toDataURL('image/png'); a.click(); }
    function render()            { _draw(); }
    function getStats()          { return _statsCache; }

    // ═══════════════════════════════════════════════════════════
    // MAPS + STATS
    // ═══════════════════════════════════════════════════════════
    function _buildMaps() {
        _succMap.clear(); _predMap.clear();
        _tasks.forEach(t => { _succMap.set(t.uid, []); _predMap.set(t.uid, []); });
        _tasks.forEach(t => {
            (t.predecessors || []).forEach(p => {
                const pid = p.predecessorUID;
                if (_succMap.has(pid)) _succMap.get(pid).push(t.uid);
                if (_predMap.has(t.uid)) _predMap.get(t.uid).push(pid);
            });
        });
    }

    function _buildStats() {
        const n = _tasks.length;
        if (!n) { _statsCache = null; return; }
        const critTasks  = _tasks.filter(t => t.critical);
        const lateTasks  = _tasks.filter(t => t.status === 'late');
        const zeroFloat  = _tasks.filter(t => t.totalFloat === 0);
        const bottlenecks = _tasks.filter(t => (_succMap.get(t.uid)||[]).length >= 3);
        const floatValues = _tasks.map(t => t.totalFloat).filter(v => v != null && isFinite(v));
        const avgFloat   = floatValues.length ? Math.round(floatValues.reduce((a,b)=>a+b,0)/floatValues.length) : 0;
        const critPathLen = critTasks.reduce((s,t) => s + (t.durationDays||0), 0);
        _statsCache = { n, critCount: critTasks.length, lateCount: lateTasks.length, zeroFloatCount: zeroFloat.length, bottleneckCount: bottlenecks.length, avgFloat, critPathLen };
        // Update stats bar in toolbar
        _renderStatsBar();
    }

    function _renderStatsBar() {
        const el = document.getElementById('ndStatsBar');
        if (!el || !_statsCache) return;
        const s = _statsCache;
        el.innerHTML = '';
        const items = [
            { label: 'Critical path', val: s.critPathLen + 'd', color: s.critCount > 0 ? '#ef4444' : '' },
            { label: 'Critical tasks', val: s.critCount, color: s.critCount > 0 ? '#ef4444' : '' },
            { label: 'Late tasks', val: s.lateCount, color: s.lateCount > 0 ? '#f59e0b' : '' },
            { label: 'Zero float', val: s.zeroFloatCount, color: s.zeroFloatCount > 0 ? '#f59e0b' : '' },
            { label: 'Bottlenecks', val: s.bottleneckCount, color: s.bottleneckCount > 0 ? '#f97316' : '' },
            { label: 'Avg float', val: s.avgFloat + 'd', color: '' },
        ];
        items.forEach(item => {
            const span = document.createElement('span');
            span.className = 'nd-stat-item';
            const k = document.createElement('span'); k.className = 'nd-stat-key'; k.textContent = item.label;
            const v = document.createElement('span'); v.className = 'nd-stat-val';
            v.textContent = item.val;
            if (item.color) v.style.color = item.color;
            span.appendChild(k); span.appendChild(v);
            el.appendChild(span);
        });
    }

    function _riskScore(task) {
        let r = 0;
        if (task.totalFloat === 0)                     r += 38;
        if (task.critical)                              r += 24;
        if (task.status === 'late')                     r += 20;
        if ((_succMap.get(task.uid)||[]).length >= 3)   r += 10; // bottleneck
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
                String(t.uid).includes(_searchQuery)) {
                _highlightUids.add(t.uid);
            }
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
        if (!_tasks.length) return;
        const D = DIMS[_mode];

        const uidIdx = new Map();
        _tasks.forEach((t,i) => uidIdx.set(t.uid, i));
        const n = _tasks.length;
        const succ = Array.from({length:n}, () => []);
        const pred = Array.from({length:n}, () => []);
        _tasks.forEach((t,i) => {
            (t.predecessors||[]).forEach(p => {
                const pi = uidIdx.get(p.predecessorUID);
                if (pi != null && pi !== i) { succ[pi].push(i); pred[i].push(pi); }
            });
        });

        // Layer assignment (longest path from sources)
        const layer = new Array(n).fill(0);
        const vis   = new Array(n).fill(false);
        const dfsL  = u => {
            if (vis[u]) return layer[u];
            vis[u] = true;
            pred[u].forEach(p => { layer[u] = Math.max(layer[u], dfsL(p) + 1); });
            return layer[u];
        };
        for (let i = 0; i < n; i++) dfsL(i);

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

        // ── Gravity-based Y: pull each node toward avg Y of its predecessors
        const posY = new Float64Array(n).fill(-1);

        grps.forEach((grp, lv) => {
            const count = grp.length;
            // Initial Y: spread evenly
            const totalH = count * (D.h + D.gapY) - D.gapY;
            grp.forEach((ti, row) => {
                posY[ti] = row * (D.h + D.gapY) - totalH / 2;
            });

            if (lv > 0) {
                // Adjust each node toward the Y-average of its predecessors
                grp.forEach(ti => {
                    const preds = pred[ti].filter(pi => posY[pi] >= 0);
                    if (!preds.length) return;
                    const avgY = preds.reduce((s,pi) => s + posY[pi], 0) / preds.length;
                    // Blend: 60% gravity, 40% original position
                    posY[ti] = posY[ti] * 0.4 + avgY * 0.6;
                });

                // Re-sort within column by adjusted Y
                grp.sort((a,b) => posY[a] - posY[b]);

                // Fix overlaps (push apart if too close)
                for (let i = 1; i < grp.length; i++) {
                    const minY = posY[grp[i-1]] + D.h + D.gapY;
                    if (posY[grp[i]] < minY) posY[grp[i]] = minY;
                }
            }
        });

        // Normalize: shift all to positive + padding
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

        // Edges
        _nodes.forEach(nd => {
            (nd.task.predecessors||[]).forEach(p => {
                const src = _nodeMap.get(p.predecessorUID);
                if (src) _edges.push({ from:src, to:nd, isCrit: src.task.critical && nd.task.critical, type: p.typeName||'FS' });
            });
        });
    }

    // ═══════════════════════════════════════════════════════════
    // CANVAS
    // ═══════════════════════════════════════════════════════════
    function _resize() {
        if (!_canvas) return;
        const D = DIMS[_mode];
        let mx=0, my=0;
        _nodes.forEach(nd => { mx=Math.max(mx,nd.x+nd.w); my=Math.max(my,nd.y+nd.h); });
        const pw = _wrap?.clientWidth  || 800;
        const ph = _wrap?.clientHeight || 600;
        const cw = Math.max(mx + D.pad*2, pw);
        const ch = Math.max(my + D.pad*2, ph);
        _canvas.width  = Math.floor(cw * _dpr);
        _canvas.height = Math.floor(ch * _dpr);
        _canvas.style.width  = cw + 'px';
        _canvas.style.height = ch + 'px';
        _ctx.setTransform(_dpr, 0, 0, _dpr, 0, 0);
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

        // Dot grid
        _ctx.fillStyle = 'rgba(255,255,255,0.025)';
        for (let x=28; x<cw; x+=28) for (let y=28; y<ch; y+=28) {
            _ctx.beginPath(); _ctx.arc(x,y,1,0,Math.PI*2); _ctx.fill();
        }

        _ctx.translate(_panX, _panY); _ctx.scale(_scale, _scale);

        const hasImpact = _impactSet.size > 0;
        const hasSearch = _highlightUids.size > 0;

        // Edges
        _edges.forEach(e => {
            const dimmed = (hasImpact && !(_impactSet.has(e.from.task.uid) && _impactSet.has(e.to.task.uid)))
                        || (hasSearch && !(_highlightUids.has(e.from.task.uid) || _highlightUids.has(e.to.task.uid)));
            _drawEdge(e, dimmed);
        });

        // Nodes
        _nodes.forEach(nd => {
            const dimmed = (hasImpact && !_impactSet.has(nd.task.uid))
                        || (hasSearch && !_highlightUids.has(nd.task.uid));
            _drawNode(nd, dimmed);
        });

        _ctx.restore();

        if (_nodes.length > 0) _drawMinimap(cw, ch);
        else _drawEmpty(cw, ch);
        _drawLegend(cw, ch);
    }

    // ── Edge ──────────────────────────────────────────────────
    function _drawEdge(e, dimmed) {
        const {from:s, to:t, isCrit, type} = e;
        const isHov = _hovUid && (s.task.uid===_hovUid || t.task.uid===_hovUid);
        const color = dimmed ? C.eDim : isHov ? C.eHov : isCrit ? C.eCrit : C.eNorm;
        const lw    = isHov ? 2.5 : isCrit ? 2 : 1.5;

        _ctx.save();
        _ctx.strokeStyle = color; _ctx.lineWidth = lw; _ctx.lineJoin = 'round';
        if (isCrit && !dimmed) { _ctx.shadowColor = C.critGlow; _ctx.shadowBlur = 6; }
        if (type !== 'FS') _ctx.setLineDash([5,4]);

        const x1=s.x+s.w, y1=Math.floor(s.y+s.h/2);
        const x2=t.x,     y2=Math.floor(t.y+t.h/2);
        const mx=Math.floor((x1+x2)/2);

        _ctx.beginPath();
        _ctx.moveTo(x1,y1); _ctx.lineTo(mx,y1); _ctx.lineTo(mx,y2); _ctx.lineTo(x2,y2);
        _ctx.stroke();

        // Arrow
        _ctx.setLineDash([]); _ctx.shadowBlur=0; _ctx.fillStyle=color;
        _ctx.beginPath(); _ctx.moveTo(x2,y2); _ctx.lineTo(x2-8,y2-4); _ctx.lineTo(x2-8,y2+4); _ctx.closePath(); _ctx.fill();

        if (type!=='FS' && !dimmed) {
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
        const pct  = t.percentComplete || 0;
        const isCrit = t.critical, isDone = pct >= 100;
        const isLate = t.status === 'late';
        const isMile = t.milestone;
        const isHov  = _hovUid === t.uid;
        const isSel  = _selUid === t.uid;
        const alpha  = dimmed ? 0.22 : 1;

        _ctx.save();
        _ctx.globalAlpha = alpha;

        // Glow for selected/hovered
        if ((isHov || isSel) && !dimmed) {
            _ctx.shadowColor = isCrit ? C.critGlow : 'rgba(99,102,241,0.55)';
            _ctx.shadowBlur  = isSel ? 20 : 14;
        } else if (isCrit && !dimmed) {
            _ctx.shadowColor = C.critGlow; _ctx.shadowBlur = 5;
        }

        // Milestone diamond
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

        // Fill
        _ctx.fillStyle = isCrit ? C.critFill : isDone ? 'rgba(34,197,94,0.07)' : isLate ? 'rgba(245,158,11,0.07)' : C.surf;
        _rr(x, y, w, h, 8); _ctx.fill();

        // Risk border (left stripe — colour from risk score)
        const riskCol = _riskColor(risk);
        _ctx.fillStyle = riskCol;
        _ctx.beginPath();
        _ctx.moveTo(x+8,y); _ctx.arcTo(x,y,x,y+8,8);
        _ctx.lineTo(x,y+h-8); _ctx.arcTo(x,y+h,x+8,y+h,8);
        _ctx.lineTo(x+4,y+h); _ctx.lineTo(x+4,y); _ctx.closePath(); _ctx.fill();

        // Outline
        _ctx.shadowBlur=0;
        _ctx.strokeStyle = (isHov||isSel) ? (isCrit ? C.crit : C.acc) : isCrit ? C.crit : C.bord;
        _ctx.lineWidth   = (isHov||isSel) ? 2 : isCrit ? 1.5 : 1;
        _rr(x, y, w, h, 8); _ctx.stroke();

        // Bottleneck badge ⚡
        if (isBotl) {
            _ctx.fillStyle = C.bottle;
            _ctx.font = `bold 9px Inter,sans-serif`; _ctx.textAlign='right'; _ctx.textBaseline='top';
            _ctx.fillText('⚡', x+w-5, y+3);
        }

        // Content
        if (_mode === 'micro')   _drawMicro(nd, D);
        else if (_mode === 'compact') _drawCompact(nd, D);
        else                          _drawNormal(nd, D);

        _ctx.restore();
    }

    function _drawNormal(nd, D) {
        const {task:t, x, y, w, h} = nd;
        const pct = t.percentComplete || 0;
        const rH  = Math.floor(h/3);
        const tx  = x+12;

        // Dividers
        _ctx.fillStyle='rgba(255,255,255,0.04)';
        _ctx.fillRect(x+4,y+rH,w-8,1); _ctx.fillRect(x+4,y+rH*2,w-8,1);

        // Row 1: ES | name | EF
        const esV = t._es!=null?Math.round(t._es):'—';
        const efV = t._ef!=null?Math.round(t._ef):'—';
        _ctx.font=`400 8px Inter,sans-serif`; _ctx.fillStyle=C.dim;
        _ctx.textAlign='left';  _ctx.textBaseline='top'; _ctx.fillText('ES '+esV, tx, y+5);
        _ctx.textAlign='right'; _ctx.fillText('EF '+efV, x+w-6, y+5);
        _ctx.fillStyle = t.critical ? C.crit : C.txt;
        _ctx.font=`600 ${D.fs}px Inter,sans-serif`;
        _ctx.textAlign='center'; _ctx.textBaseline='middle';
        const nm = t.name.length>19 ? t.name.slice(0,17)+'…' : t.name;
        _ctx.fillText(nm, x+w/2, y+rH/2);

        // Row 2: progress bar + % + dur
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
        const r3=y+rH*2+5;
        const lsV = t._ls!=null&&isFinite(t._ls)?Math.round(t._ls):'—';
        const lfV = t._lf!=null&&isFinite(t._lf)?Math.round(t._lf):'—';
        const tf  = t.totalFloat!=null&&isFinite(t.totalFloat)?Math.round(t.totalFloat):null;
        _ctx.font=`400 8px Inter,sans-serif`; _ctx.fillStyle=C.dim;
        _ctx.textAlign='left';  _ctx.textBaseline='top'; _ctx.fillText('LS '+lsV, tx, r3);
        _ctx.textAlign='right'; _ctx.fillText('LF '+lfV, x+w-6, r3);
        if (tf!==null) {
            _ctx.fillStyle = tf===0?C.crit:tf<=2?C.late:C.dim;
            _ctx.textAlign='center'; _ctx.fillText('TF '+tf, x+w/2, r3);
        }
        // Float bar (visual)
        if (tf !== null && tf > 0) {
            const maxF = 20, fW = Math.min(tf/maxF, 1) * (w-20);
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
        // progress strip
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
    function _drawMinimap(cw, ch) {
        const mx=cw-MM.w-MM.pad, my=ch-MM.h-MM.pad;
        let x0=Infinity,x1=-Infinity,y0=Infinity,y1=-Infinity;
        _nodes.forEach(n => { x0=Math.min(x0,n.x); x1=Math.max(x1,n.x+n.w); y0=Math.min(y0,n.y); y1=Math.max(y1,n.y+n.h); });
        const dw=x1-x0||1, dh=y1-y0||1;
        const sc=Math.min((MM.w-6)/dw,(MM.h-6)/dh, 0.25);
        _ctx.fillStyle='rgba(8,9,16,0.92)'; _rrp(mx,my,MM.w,MM.h,7); _ctx.fill();
        _ctx.strokeStyle='rgba(255,255,255,0.07)'; _ctx.lineWidth=1; _rrp(mx,my,MM.w,MM.h,7); _ctx.stroke();
        _ctx.save(); _ctx.beginPath(); _ctx.rect(mx+1,my+1,MM.w-2,MM.h-2); _ctx.clip();
        _nodes.forEach(n => {
            const nw=Math.max(3,n.w*sc), nh=Math.max(2,n.h*sc);
            const nx=mx+3+(n.x-x0)*sc, ny=my+3+(n.y-y0)*sc;
            _ctx.globalAlpha = _impactSet.size>0 ? (_impactSet.has(n.task.uid)?0.9:0.18) : 0.7;
            _ctx.fillStyle=n.task.critical?C.crit:n.task.percentComplete>=100?C.done:n.task.status==='late'?C.late:C.acc;
            _ctx.fillRect(nx,ny,nw,nh);
        });
        _ctx.globalAlpha=1;
        const vw=(_wrap?.clientWidth||800)/_scale, vh=(_wrap?.clientHeight||600)/_scale;
        const vx=mx+3+(-_panX/_scale-x0)*sc, vy=my+3+(-_panY/_scale-y0)*sc;
        _ctx.strokeStyle='rgba(255,255,255,0.6)'; _ctx.lineWidth=1.5;
        _ctx.strokeRect(vx,vy,vw*sc,vh*sc);
        _ctx.restore();
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
        _ctx.fillStyle=C.dim; _ctx.font='14px Inter,sans-serif'; _ctx.textAlign='center'; _ctx.textBaseline='middle';
        _ctx.fillText('No tasks with predecessors found', cw/2, ch/2-16);
        _ctx.font='11px Inter,sans-serif';
        _ctx.fillText('Add predecessor relationships to tasks — then switch to Network view', cw/2, ch/2+10);
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
    function _xy(e) { const r=_canvas.getBoundingClientRect(); return { x:(e.clientX-r.left-_panX)/_scale, y:(e.clientY-r.top-_panY)/_scale }; }
    function _cxy(e) { const r=_canvas.getBoundingClientRect(); return { cx:e.clientX-r.left, cy:e.clientY-r.top }; }

    function _onWheel(e) {
        e.preventDefault();
        const {cx,cy}=_cxy(e);
        const ns=Math.min(3,Math.max(0.15,_scale*(e.deltaY>0?0.88:1.14)));
        _panX=cx-(cx-_panX)*(ns/_scale); _panY=cy-(cy-_panY)*(ns/_scale);
        _scale=ns; _draw();
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
            _draw();
            if (hit) _showTip(e.clientX-_canvas.getBoundingClientRect().left, e.clientY-_canvas.getBoundingClientRect().top, hit);
            else _clearTip();
        }
        if (_isPanning) { _panX=e.clientX-_pStart.x; _panY=e.clientY-_pStart.y; _draw(); }
    }
    function _onUp()    { _isPanning=false; _canvas.style.cursor=_hovUid?'pointer':'grab'; }
    function _onLeave() { _isPanning=false; _hovUid=null; _clearTip(); _draw(); }
    function _onDbl(e) {
        const {x,y}=_xy(e); const hit=_hit(x,y);
        if (hit) _canvas.dispatchEvent(new CustomEvent('nodeDoubleClick',{bubbles:true,detail:{task:hit.task}}));
    }
    function _onKey(e) {
        if (e.target.tagName==='INPUT') return;
        if (e.key==='+' || e.key==='=') zoomIn();
        else if (e.key==='-') zoomOut();
        else if (e.key==='f' || e.key==='F') fitToScreen();
        else if (e.key==='c' || e.key==='C') { _filterMode = _filterMode==='critical'?'all':'critical'; update(_tasks); }
        else if (e.key==='Escape') { _selUid=null; _impactSet.clear(); _highlightUids.clear(); _searchQuery=''; const si=$nd('ndSearch'); if(si) si.value=''; _draw(); }
    }
    function $nd(id) { return document.getElementById(id); }

    // ── Tooltip ───────────────────────────────────────────────
    function _showTip(cx, cy, nd) {
        _clearTip();
        const {task:t, risk, isBotl} = nd;
        const el = document.createElement('div'); el.className='nd-tooltip';

        const title=document.createElement('div'); title.className='nd-tip-title';
        title.textContent=(t.critical?'🔴 ':t.status==='late'?'⏰ ':'')+t.name;
        el.appendChild(title);

        const riskBar = document.createElement('div');
        riskBar.style.cssText = `height:3px;border-radius:2px;background:${_riskColor(risk)};width:${risk}%;margin-bottom:8px;transition:width .3s`;
        el.appendChild(riskBar);

        const riskLabel = document.createElement('div');
        riskLabel.textContent = `Risk Score: ${risk}/100 ${risk>=70?'🔴 High':risk>=40?'🟡 Medium':'🟢 Low'}`;
        riskLabel.style.cssText='font-size:0.7rem;margin-bottom:6px;font-weight:600;color:'+_riskColor(risk);
        el.appendChild(riskLabel);

        const tf = t.totalFloat!=null&&isFinite(t.totalFloat)?Math.round(t.totalFloat):null;
        const sucCount = (_succMap.get(t.uid)||[]).length;
        const preCount = (_predMap.get(t.uid)||[]).length;
        const rows = [
            ['Duration',  (t.durationDays||0)+'d'],
            ['Progress',  (t.percentComplete||0)+'%'],
            ['Float',     tf!==null?(tf+'d'+(tf===0?' ⚠ Critical path':'')):'—'],
            ['ES→EF',     (t._es!=null?Math.round(t._es):'?')+' → '+(t._ef!=null?Math.round(t._ef):'?')],
            ['LS→LF',     (t._ls!=null&&isFinite(t._ls)?Math.round(t._ls):'?')+' → '+(t._lf!=null&&isFinite(t._lf)?Math.round(t._lf):'?')],
            ['Successors', sucCount+(isBotl?' ⚡ Bottleneck':'')],
            ['Predecessors', preCount],
            ['Resource',  (t.resourceNames||[]).join(', ')||'—'],
            ['Status',    (t.statusIcon||'')+' '+(t.status||'normal')],
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
    function _onTS(e) { e.preventDefault(); _touches=[...e.touches]; if (_touches.length===2) { _touchDist=Math.hypot(_touches[0].clientX-_touches[1].clientX,_touches[0].clientY-_touches[1].clientY); } else { _isPanning=true; _pStart={x:_touches[0].clientX-_panX,y:_touches[0].clientY-_panY}; } }
    function _onTM(e) { e.preventDefault(); _touches=[...e.touches]; if (_touches.length===2) { const d=Math.hypot(_touches[0].clientX-_touches[1].clientX,_touches[0].clientY-_touches[1].clientY); _zoom(_scale*(d/(_touchDist||1))); _touchDist=d; } else if (_isPanning) { _panX=_touches[0].clientX-_pStart.x; _panY=_touches[0].clientY-_pStart.y; _draw(); } }
    function _onTE() { _isPanning=false; _touches=[]; }

    // ── Zoom/Fit ──────────────────────────────────────────────
    function _zoom(ns) { _scale=Math.min(3,Math.max(0.12,ns)); _draw(); }
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

    // ── Helpers ───────────────────────────────────────────────
    function _rr(x,y,w,h,r) { _ctx.beginPath(); _ctx.moveTo(x+r,y); _ctx.arcTo(x+w,y,x+w,y+h,r); _ctx.arcTo(x+w,y+h,x,y+h,r); _ctx.arcTo(x,y+h,x,y,r); _ctx.arcTo(x,y,x+w,y,r); _ctx.closePath(); }
    function _rrp(x,y,w,h,r) { _ctx.beginPath(); _ctx.moveTo(x+r,y); _ctx.arcTo(x+w,y,x+w,y+r,r); _ctx.arcTo(x+w,y+h,x+w-r,y+h,r); _ctx.arcTo(x,y+h,x,y+h-r,r); _ctx.arcTo(x,y,x+r,y,r); _ctx.closePath(); }
    function _diamond(cx,cy,rw,rh) { _ctx.beginPath(); _ctx.moveTo(cx,cy-rh); _ctx.lineTo(cx+rw,cy); _ctx.lineTo(cx,cy+rh); _ctx.lineTo(cx-rw,cy); _ctx.closePath(); }

    export const NetworkDiagram = { init, update, render, cleanup, setMode, setFilter, setSearch, zoomIn, zoomOut, fitToScreen, exportPNG, getStats };
