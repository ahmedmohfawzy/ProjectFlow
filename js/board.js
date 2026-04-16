/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 */
/**
 * ═══════════════════════════════════════════════════════
 * Sprint D.1 — Board / Kanban View
 * Sprint D.2 — Resource Heatmap
 * ═══════════════════════════════════════════════════════
 */

// ══════════════════════════════════════════════════════════════
// D.1  KANBAN BOARD
// ══════════════════════════════════════════════════════════════


    const COLUMNS = [
        { id: 'notstarted', label: 'Not Started', color: '#64748b', test: t => (t.percentComplete || 0) === 0 && !t.milestone },
        { id: 'inprogress', label: 'In Progress',  color: '#6366f1', test: t => (t.percentComplete || 0) > 0 && (t.percentComplete || 0) < 100 },
        { id: 'complete',   label: 'Complete',     color: '#22c55e', test: t => (t.percentComplete || 0) >= 100 },
        { id: 'milestone',  label: 'Milestones',   color: '#f59e0b', test: t => !!t.milestone }
    ];

    let _project   = null;
    let _groupBy   = 'status';
    let _query     = '';
    let _onSave    = null;

    /** Attach external save callback */
    function setCallbacks({ onSave }) { _onSave = onSave; }

    /** Full re-render */
    function render(project, groupBy, query) {
        _project = project;
        if (groupBy !== undefined) _groupBy = groupBy;
        if (query   !== undefined) _query   = (query || '').toLowerCase();

        const wrap = document.getElementById('boardColumns');
        if (!wrap || !_project) return;
        wrap.innerHTML = '';

        const columns = _buildColumns();
        columns.forEach(col => wrap.appendChild(_renderColumn(col)));
        _initDragDrop(wrap);
    }

    /** Build column definitions based on groupBy */
    function _buildColumns() {
        const tasks = (_project.tasks || []).filter(t => !t.summary);
        const filtered = _query
            ? tasks.filter(t => t.name.toLowerCase().includes(_query) || (t.resourceNames || []).join(' ').toLowerCase().includes(_query))
            : tasks;

        if (_groupBy === 'resource') {
            const resMap = new Map();
            resMap.set('__unassigned__', { id: '__unassigned__', label: 'Unassigned', color: '#64748b', cards: [] });
            filtered.forEach(t => {
                const names = t.resourceNames && t.resourceNames.length ? t.resourceNames : ['__unassigned__'];
                names.forEach(n => {
                    if (!resMap.has(n)) resMap.set(n, { id: n, label: n, color: '#6366f1', cards: [] });
                    resMap.get(n).cards.push(t);
                });
            });
            return [...resMap.values()].filter(c => c.cards.length > 0);
        }

        if (_groupBy === 'tag') {
            const tagMap = new Map();
            tagMap.set('__none__', { id: '__none__', label: 'No Tag', color: '#64748b', cards: [] });
            filtered.forEach(t => {
                const tags = t.tags && t.tags.length ? t.tags : ['__none__'];
                tags.forEach(tg => {
                    if (!tagMap.has(tg)) tagMap.set(tg, { id: tg, label: tg, color: '#8b5cf6', cards: [] });
                    tagMap.get(tg).cards.push(t);
                });
            });
            return [...tagMap.values()].filter(c => c.cards.length > 0);
        }

        // Default: by status
        return COLUMNS.map(col => ({
            ...col,
            cards: filtered.filter(col.test)
        }));
    }

    /** Render a single Kanban column */
    function _renderColumn(col) {
        const div = document.createElement('div');
        div.className = 'board-col';
        div.dataset.colId = col.id;

        const header = document.createElement('div');
        header.className = 'board-col-header';
        header.innerHTML = `
            <span class="board-col-dot" style="background:${col.color}"></span>
            <span class="board-col-title">${col.label}</span>
            <span class="board-col-count">${col.cards.length}</span>`;
        div.appendChild(header);

        const list = document.createElement('div');
        list.className = 'board-card-list';
        list.dataset.colId = col.id;
        col.cards.forEach(t => list.appendChild(_renderCard(t, col)));
        div.appendChild(list);

        return div;
    }

    /** Render a single task card */
    function _renderCard(task, col) {
        const card = document.createElement('div');
        card.className = 'board-card';
        card.draggable = true;
        card.dataset.uid = task.uid;
        if (task.critical) card.classList.add('board-card-critical');
        if (task.status === 'late') card.classList.add('board-card-late');

        const pct = task.percentComplete || 0;
        const tags = (task.tags || []).map(t => `<span class="board-tag">${t}</span>`).join('');
        const res  = (task.resourceNames || []).join(', ');
        const due  = task.finish ? _fmtDate(task.finish) : '';
        const isLate = task.status === 'late';

        card.innerHTML = `
            <div class="board-card-name">${_esc(task.name)}</div>
            ${tags ? `<div class="board-card-tags">${tags}</div>` : ''}
            <div class="board-card-meta">
                ${res ? `<span class="board-card-res">👤 ${_esc(res)}</span>` : ''}
                ${due ? `<span class="board-card-due${isLate ? ' late' : ''}">📅 ${due}</span>` : ''}
            </div>
            <div class="board-card-progress">
                <div class="board-card-bar" style="width:${pct}%;background:${pct>=100?'#22c55e':col.color}"></div>
            </div>
            <div class="board-card-pct">${pct}%</div>`;

        return card;
    }

    /** Initialise HTML5 drag-and-drop for status-based grouping */
    function _initDragDrop(wrap) {
        if (_groupBy !== 'status') return; // DnD only for status columns

        let dragUid = null;

        wrap.addEventListener('dragstart', e => {
            const card = e.target.closest('.board-card');
            if (!card) return;
            dragUid = card.dataset.uid;
            card.classList.add('board-card-dragging');
            e.dataTransfer.effectAllowed = 'move';
        });

        wrap.addEventListener('dragend', e => {
            wrap.querySelectorAll('.board-card-dragging').forEach(c => c.classList.remove('board-card-dragging'));
            wrap.querySelectorAll('.board-col-list-over').forEach(c => c.classList.remove('board-col-list-over'));
        });

        wrap.addEventListener('dragover', e => {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            const list = e.target.closest('.board-card-list');
            if (list) {
                wrap.querySelectorAll('.board-col-list-over').forEach(c => c.classList.remove('board-col-list-over'));
                list.classList.add('board-col-list-over');
            }
        });

        wrap.addEventListener('dragleave', e => {
            const list = e.target.closest('.board-card-list');
            if (list && !list.contains(e.relatedTarget)) list.classList.remove('board-col-list-over');
        });

        wrap.addEventListener('drop', e => {
            e.preventDefault();
            const list = e.target.closest('.board-card-list');
            if (!list || !dragUid || !_project) return;
            const colId = list.dataset.colId;
            const task  = _project.tasks.find(t => String(t.uid) === String(dragUid));
            if (!task) return;

            // Map column → percentComplete
            const pctMap = { notstarted: 0, inprogress: task.percentComplete > 0 && task.percentComplete < 100 ? task.percentComplete : 50, complete: 100, milestone: task.percentComplete };
            if (pctMap[colId] !== undefined) {
                task.percentComplete = pctMap[colId];
                if (_onSave) _onSave(task);
                render(_project);
            }
        });
    }

    // ── Utilities ──────────────────────────────────────────────
    function _esc(s) { const d = document.createElement('span'); d.textContent = String(s || ''); return d.innerHTML; }
    function _fmtDate(d) {
        const dt = new Date(d); if (isNaN(dt)) return '';
        return `${dt.getDate()} ${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][dt.getMonth()]}`;
    }

    export const BoardView = { render, setCallbacks };


// ══════════════════════════════════════════════════════════════
// D.2  RESOURCE HEATMAP
// ══════════════════════════════════════════════════════════════


    /**
     * Render a weekly resource load heatmap inside `container`.
     * @param {HTMLElement} container
     * @param {Object} project
     * @param {number} [weeks=12] — number of weeks to show
     */
    function renderHeatmap(container, project, weeks) {
        weeks = weeks || 12;
        if (!container || !project) return;
        container.innerHTML = '';

        const resources = (project.resources || []).filter(r => r.name);
        if (!resources.length) {
            container.innerHTML = '<div class="pf-empty">No resources found. Assign resources to tasks first.</div>';
            return;
        }

        // Determine week starts
        const startDate = project.startDate ? new Date(project.startDate) : new Date();
        startDate.setHours(0,0,0,0);
        // Align to Monday
        const dayOfWeek = startDate.getDay();
        const daysToMon = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
        const weekStart = new Date(startDate);
        weekStart.setDate(weekStart.getDate() + daysToMon);

        const weekStarts = [];
        for (let i = 0; i < weeks; i++) {
            const d = new Date(weekStart);
            d.setDate(d.getDate() + i * 7);
            weekStarts.push(d);
        }

        // Calculate load per resource per week (hours)
        const hoursPerDay = (project.hoursPerDay || 8);
        const loads = new Map(); // resUID → { weekIdx: hours }

        (project.assignments || []).forEach(a => {
            const res = resources.find(r => r.uid === a.resourceUID);
            if (!res) return;
            const task = (project.tasks || []).find(t => t.uid === a.taskUID);
            if (!task) return;
            const tStart = new Date(task.start); const tEnd = new Date(task.finish);
            if (!loads.has(res.uid)) loads.set(res.uid, new Array(weeks).fill(0));
            const loadArr = loads.get(res.uid);

            weekStarts.forEach((ws, wi) => {
                const we = new Date(ws); we.setDate(we.getDate() + 5); // Mon–Fri
                const overlapStart = tStart > ws ? tStart : ws;
                const overlapEnd   = tEnd < we ? tEnd : we;
                if (overlapStart < overlapEnd) {
                    const days = Math.ceil((overlapEnd - overlapStart) / 86400000);
                    loadArr[wi] += days * hoursPerDay * (a.units || 1);
                }
            });
        });

        const maxHrsPerWeek = 5 * hoursPerDay; // normal 5-day week

        // Build HTML table
        const wrap = document.createElement('div');
        wrap.className = 'heatmap-wrap';

        // Header row
        const table = document.createElement('table');
        table.className = 'heatmap-table';
        const thead = document.createElement('thead');
        const headerRow = document.createElement('tr');
        const thName = document.createElement('th'); thName.textContent = 'Resource'; thName.className = 'hm-col-name';
        headerRow.appendChild(thName);
        weekStarts.forEach(ws => {
            const th = document.createElement('th'); th.className = 'hm-col-week';
            th.innerHTML = `<span>${_fmtWeek(ws)}</span>`;
            headerRow.appendChild(th);
        });
        thead.appendChild(headerRow); table.appendChild(thead);

        const tbody = document.createElement('tbody');
        resources.forEach(res => {
            const tr = document.createElement('tr');
            const tdName = document.createElement('td'); tdName.className = 'hm-res-name'; tdName.textContent = res.name;
            tr.appendChild(tdName);

            const loadArr = loads.get(res.uid) || new Array(weeks).fill(0);
            loadArr.forEach((hrs, wi) => {
                const td = document.createElement('td'); td.className = 'hm-cell';
                const ratio = hrs / maxHrsPerWeek;
                const color = _heatColor(ratio);
                td.style.background = color;
                td.title = `${res.name} · ${_fmtWeek(weekStarts[wi])}: ${Math.round(hrs)}h (${Math.round(ratio * 100)}%)`;
                if (ratio > 1) {
                    td.classList.add('hm-overalloc');
                    td.innerHTML = `<span class="hm-label">⚠ ${Math.round(ratio * 100)}%</span>`;
                } else if (hrs > 0) {
                    td.innerHTML = `<span class="hm-label">${Math.round(hrs)}h</span>`;
                }
                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);
        wrap.appendChild(table);

        // Legend
        const legend = document.createElement('div');
        legend.className = 'hm-legend';
        legend.innerHTML = `
            <span class="hm-legend-item"><span class="hm-dot" style="background:#1e293b"></span>0%</span>
            <span class="hm-legend-item"><span class="hm-dot" style="background:#166534"></span>1–50%</span>
            <span class="hm-legend-item"><span class="hm-dot" style="background:#854d0e"></span>51–100%</span>
            <span class="hm-legend-item"><span class="hm-dot" style="background:#ef4444"></span>&gt;100% (overloaded)</span>`;
        wrap.appendChild(legend);

        container.appendChild(wrap);
    }

    function _fmtWeek(d) {
        return `${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getMonth()]} ${d.getDate()}`;
    }

    function _heatColor(ratio) {
        if (ratio <= 0)    return 'rgba(255,255,255,0.03)';
        if (ratio <= 0.25) return '#14532d';
        if (ratio <= 0.50) return '#166534';
        if (ratio <= 0.75) return '#15803d';
        if (ratio <= 1.00) return '#854d0e';
        if (ratio <= 1.25) return '#b91c1c';
        return '#ef4444';
    }

    export const ResourceHeatmap = { render: renderHeatmap };
