/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 * Proprietary Software — Unauthorized use prohibited.
 * https://www.linkedin.com/in/ahmed-m-fawzy
 */
/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — Gantt Chart Renderer v2
 * Canvas-based with drag & drop, baseline bars,
 * critical path highlighting, dependency types
 * ═══════════════════════════════════════════════════════
 */
    const CONFIG = {
        rowHeight: 32, barHeight: 12, barRadius: 3, summaryHeight: 7,
        milestoneSize: 9, baselineHeight: 4, baselineOffset: 3,
        fontFamily: 'Inter, -apple-system, sans-serif', fontSize: 10,
        padding: { right: 40, bottom: 40 },
        colors: {
            bg: '#0f1117', gridLine: 'rgba(255,255,255,0.04)', gridLineMajor: 'rgba(255,255,255,0.08)',
            weekend: 'rgba(255,255,255,0.015)', today: 'rgba(239,68,68,0.12)', todayLine: '#ef4444',
            bar: '#6366f1', barHover: '#818cf8', barBorder: '#4f46e5',
            progress: '#22c55e', progressBg: 'rgba(34,197,94,0.25)',
            summary: '#8b5cf6', summaryBorder: '#7c3aed',
            milestone: '#f59e0b', milestoneBorder: '#d97706',
            critical: '#ef4444', criticalBorder: '#dc2626',
            baseline: 'rgba(148,163,184,0.35)', baselineBorder: 'rgba(148,163,184,0.5)',
            link: 'rgba(100,116,139,0.5)', linkArrow: 'rgba(100,116,139,0.7)',
            text: '#9aa0b4', textBright: '#e8eaed', textMuted: '#5c6378',
            selectedRow: 'rgba(99,102,241,0.1)', dragPreview: 'rgba(99,102,241,0.3)'
        }
    };

    const ZOOM_LEVELS = [
        { name: 'Hours', dayWidth: 80 }, { name: 'Days', dayWidth: 32 },
        { name: 'Weeks', dayWidth: 12 }, { name: 'Months', dayWidth: 4 }
    ];

    let currentZoom = 1, canvas, ctx, headerEl;
    let tasks = [], projectStart, projectEnd, dayWidth;
    let scrollX = 0, scrollY = 0;
    let hoveredTaskIndex = -1, selectedTaskIndex = -1;
    let onTaskSelect = null, onTaskUpdate = null;
    let showBaseline = true, showCritical = true, showLinks = true;

    // Drag state
    let isDragging = false, dragType = null; // 'move', 'resize-right', 'link'
    let dragTaskIndex = -1, dragStartX = 0, dragStartDate = null, dragOrigDuration = 0;
    let linkFromIndex = -1;
    let _abortController = null; // For grouped listener cleanup

    function init(canvasEl, headerElement, options = {}) {
        // Cleanup previous listeners to prevent memory leak
        cleanup();

        canvas = canvasEl; headerEl = headerElement;
        ctx = canvas.getContext('2d');
        onTaskSelect = options.onTaskSelect || null;
        onTaskUpdate = options.onTaskUpdate || null;

        _abortController = new AbortController();
        const signal = _abortController.signal;

        canvas.addEventListener('mousemove', handleMouseMove, { signal });
        canvas.addEventListener('mousedown', handleMouseDown, { signal });
        canvas.addEventListener('mouseup', handleMouseUp, { signal });
        canvas.addEventListener('click', handleClick, { signal });
        canvas.addEventListener('mouseleave', () => { hoveredTaskIndex = -1; canvas.style.cursor = 'default'; render(); }, { signal });

        const ganttBody = canvas.parentElement;
        ganttBody.addEventListener('scroll', (e) => { scrollX = e.target.scrollLeft; scrollY = e.target.scrollTop; renderHeader(); }, { signal });
    }

    function cleanup() {
        if (_abortController) {
            _abortController.abort();
            _abortController = null;
        }
        hoveredTaskIndex = -1;
        selectedTaskIndex = -1;
        isDragging = false;
        dragType = null;
        dragTaskIndex = -1;
    }

    function update(taskList, options = {}) {
        tasks = taskList.filter(t => t.isVisible !== false);
        selectedTaskIndex = options.selectedIndex >= 0 ? options.selectedIndex : -1;
        showBaseline = options.showBaseline !== false;
        showCritical = options.showCritical !== false;
        showLinks = options.showLinks !== false;
        calculateProjectBounds();
        autoZoomForDuration();
        resize(); render(); renderHeader();
    }

    function setOptions(opts) {
        if (opts.showBaseline !== undefined) showBaseline = opts.showBaseline;
        if (opts.showCritical !== undefined) showCritical = opts.showCritical;
        if (opts.showLinks !== undefined) showLinks = opts.showLinks;
        render();
    }

    function calculateProjectBounds() {
        if (tasks.length === 0) { projectStart = new Date(); projectEnd = new Date(); projectEnd.setDate(projectEnd.getDate() + 30); return; }
        let minDate = Infinity, maxDate = -Infinity;
        for (const t of tasks) {
            const s = new Date(t.start).getTime(); const f = new Date(t.finish).getTime();
            if (s < minDate) minDate = s; if (f > maxDate) maxDate = f;
            if (t.baselineStart) { const bs = new Date(t.baselineStart).getTime(); if (bs < minDate) minDate = bs; }
            if (t.baselineFinish) { const bf = new Date(t.baselineFinish).getTime(); if (bf > maxDate) maxDate = bf; }
        }
        projectStart = new Date(minDate); projectStart.setDate(projectStart.getDate() - 3);
        projectEnd = new Date(maxDate); projectEnd.setDate(projectEnd.getDate() + 7);
    }

    /** Auto-set zoom level based on project duration to prevent canvas overflow */
    function autoZoomForDuration() {
        const totalDays = daysBetween(projectStart, projectEnd);
        if (totalDays > 365) currentZoom = 3;      // Months
        else if (totalDays > 180) currentZoom = 2;  // Weeks
        else if (totalDays > 60) currentZoom = 1;   // Days
        // else keep current (Hours/Days)
    }

    function resize() {
        dayWidth = ZOOM_LEVELS[currentZoom].dayWidth;
        const totalDays = daysBetween(projectStart, projectEnd);
        let w = Math.max(totalDays * dayWidth + CONFIG.padding.right, canvas.parentElement.clientWidth);
        let h = Math.max(tasks.length * CONFIG.rowHeight + CONFIG.padding.bottom, canvas.parentElement.clientHeight);
        const dpr = window.devicePixelRatio || 1;
        const MAX_CANVAS_DIM = 16000; // Safe limit for all browsers
        // Cap logical size to keep pixel size under limit
        if (w * dpr > MAX_CANVAS_DIM) w = Math.floor(MAX_CANVAS_DIM / dpr);
        if (h * dpr > MAX_CANVAS_DIM) h = Math.floor(MAX_CANVAS_DIM / dpr);
        // Fix: Use Math.floor to prevent non-integer canvas dimensions (TD-10)
        w = Math.floor(w);
        h = Math.floor(h);
        canvas.width = Math.floor(w * dpr); canvas.height = Math.floor(h * dpr);
        canvas.style.width = w + 'px'; canvas.style.height = h + 'px';
        ctx.scale(dpr, dpr);
    }

    function render() {
        const w = canvas.width / window.devicePixelRatio; const h = canvas.height / window.devicePixelRatio;
        ctx.clearRect(0, 0, w, h); ctx.fillStyle = CONFIG.colors.bg; ctx.fillRect(0, 0, w, h);
        drawGrid(w, h); drawTodayLine(h);
        if (showBaseline) drawBaselineBars();
        if (showLinks) drawDependencyLinks();
        drawBars();
    }

    function drawGrid(w, h) {
        const totalDays = daysBetween(projectStart, projectEnd);
        for (let d = 0; d <= totalDays; d++) {
            const x = d * dayWidth; const date = new Date(projectStart); date.setDate(date.getDate() + d);
            const dow = date.getDay();
            if (dow === 0 || dow === 6) { ctx.fillStyle = CONFIG.colors.weekend; ctx.fillRect(x, 0, dayWidth, h); }
            if (dayWidth >= 12) { ctx.strokeStyle = dow === 1 ? CONFIG.colors.gridLineMajor : CONFIG.colors.gridLine; ctx.lineWidth = dow === 1 ? 1 : 0.5; ctx.beginPath(); ctx.moveTo(x + 0.5, 0); ctx.lineTo(x + 0.5, h); ctx.stroke(); }
            else if (dow === 1) { ctx.strokeStyle = CONFIG.colors.gridLine; ctx.lineWidth = 0.5; ctx.beginPath(); ctx.moveTo(x + 0.5, 0); ctx.lineTo(x + 0.5, h); ctx.stroke(); }
        }
        for (let i = 0; i <= tasks.length; i++) { const y = i * CONFIG.rowHeight; ctx.strokeStyle = CONFIG.colors.gridLine; ctx.lineWidth = 0.5; ctx.beginPath(); ctx.moveTo(0, y + 0.5); ctx.lineTo(w, y + 0.5); ctx.stroke(); }
    }

    function drawTodayLine(h) {
        const today = new Date(); today.setHours(0, 0, 0, 0);
        const d = daysBetween(projectStart, today); if (d < 0) return;
        const x = d * dayWidth;
        ctx.fillStyle = CONFIG.colors.today; ctx.fillRect(x, 0, dayWidth, h);
        ctx.strokeStyle = CONFIG.colors.todayLine; ctx.lineWidth = 2; ctx.setLineDash([4, 4]);
        ctx.beginPath(); ctx.moveTo(x + 0.5, 0); ctx.lineTo(x + 0.5, h); ctx.stroke(); ctx.setLineDash([]);
    }

    function drawBaselineBars() {
        for (let i = 0; i < tasks.length; i++) {
            const task = tasks[i];
            if (!task.baselineStart || !task.baselineFinish || task.summary || task.milestone) continue;
            const startDay = daysBetween(projectStart, new Date(task.baselineStart));
            const endDay = daysBetween(projectStart, new Date(task.baselineFinish));
            const x = startDay * dayWidth; const w = Math.max((endDay - startDay) * dayWidth, dayWidth);
            const y = i * CONFIG.rowHeight + (CONFIG.rowHeight - CONFIG.barHeight) / 2 + CONFIG.barHeight + CONFIG.baselineOffset;
            ctx.fillStyle = CONFIG.colors.baseline; roundRect(ctx, x, y, w, CONFIG.baselineHeight, 2); ctx.fill();
            ctx.strokeStyle = CONFIG.colors.baselineBorder; ctx.lineWidth = 0.5; roundRect(ctx, x, y, w, CONFIG.baselineHeight, 2); ctx.stroke();
        }
    }

    function drawBars() {
        for (let i = 0; i < tasks.length; i++) {
            const task = tasks[i]; const y = i * CONFIG.rowHeight;
            if (i === selectedTaskIndex) { ctx.fillStyle = CONFIG.colors.selectedRow; ctx.fillRect(0, y, canvas.width / window.devicePixelRatio, CONFIG.rowHeight); }
            if (i === hoveredTaskIndex && i !== selectedTaskIndex) { ctx.fillStyle = 'rgba(255,255,255,0.03)'; ctx.fillRect(0, y, canvas.width / window.devicePixelRatio, CONFIG.rowHeight); }

            const startDay = daysBetween(projectStart, new Date(task.start));
            const endDay = daysBetween(projectStart, new Date(task.finish));
            const barX = startDay * dayWidth; const barW = Math.max((endDay - startDay) * dayWidth, dayWidth);

            if (task.milestone) drawMilestone(barX, y, task);
            else if (task.summary) drawSummaryBar(barX, y, barW, task);
            else drawTaskBar(barX, y, barW, task, i === hoveredTaskIndex);
        }
    }

    function drawTaskBar(x, y, w, task, isHovered) {
        const barY = y + (CONFIG.rowHeight - CONFIG.barHeight) / 2;
        const isCrit = showCritical && task.critical;
        const customColor = task.color || null;
        const color = isCrit ? CONFIG.colors.critical : (customColor || (isHovered ? CONFIG.colors.barHover : CONFIG.colors.bar));

        ctx.fillStyle = color; roundRect(ctx, x, barY, w, CONFIG.barHeight, CONFIG.barRadius); ctx.fill();

        if (task.percentComplete > 0) {
            const pw = w * (task.percentComplete / 100);
            ctx.fillStyle = CONFIG.colors.progressBg; ctx.globalAlpha = 0.7;
            roundRect(ctx, x, barY, pw, CONFIG.barHeight, CONFIG.barRadius); ctx.fill();
            ctx.globalAlpha = 1;
            // Progress line
            ctx.fillStyle = CONFIG.colors.progress;
            ctx.fillRect(x, barY + CONFIG.barHeight - 2, pw, 2);
        }

        ctx.strokeStyle = isCrit ? CONFIG.colors.criticalBorder : (customColor ? 'rgba(255,255,255,0.2)' : CONFIG.colors.barBorder);
        ctx.lineWidth = 1; roundRect(ctx, x, barY, w, CONFIG.barHeight, CONFIG.barRadius); ctx.stroke();

        // Label
        if (w > 55 && dayWidth >= 12) {
            ctx.fillStyle = 'rgba(255,255,255,0.9)'; ctx.font = `500 ${CONFIG.fontSize - 1}px ${CONFIG.fontFamily}`; ctx.textBaseline = 'middle';
            const label = task.name.length > Math.floor(w / 6) ? task.name.substring(0, Math.floor(w / 6) - 1) + '…' : task.name;
            ctx.fillText(label, x + 5, barY + CONFIG.barHeight / 2);
        } else if (dayWidth >= 16 && w <= 55) {
            ctx.fillStyle = CONFIG.colors.text; ctx.font = `400 ${CONFIG.fontSize - 1}px ${CONFIG.fontFamily}`; ctx.textBaseline = 'middle';
            ctx.fillText(task.name.substring(0, 22), x + w + 5, y + CONFIG.rowHeight / 2);
        }

        // Drag handles (on hover)
        if (isHovered && dayWidth >= 12) {
            // Right resize handle
            ctx.fillStyle = 'rgba(255,255,255,0.5)';
            ctx.fillRect(x + w - 4, barY + 2, 3, CONFIG.barHeight - 4);
        }
    }

    function drawSummaryBar(x, y, w, task) {
        const barY = y + (CONFIG.rowHeight - CONFIG.summaryHeight) / 2;
        ctx.fillStyle = CONFIG.colors.summary; ctx.fillRect(x, barY, w, CONFIG.summaryHeight);
        ctx.beginPath(); ctx.moveTo(x, barY + CONFIG.summaryHeight); ctx.lineTo(x + 5, barY + CONFIG.summaryHeight); ctx.lineTo(x, barY + CONFIG.summaryHeight + 4); ctx.closePath(); ctx.fill();
        ctx.beginPath(); ctx.moveTo(x + w, barY + CONFIG.summaryHeight); ctx.lineTo(x + w - 5, barY + CONFIG.summaryHeight); ctx.lineTo(x + w, barY + CONFIG.summaryHeight + 4); ctx.closePath(); ctx.fill();
    }

    function drawMilestone(x, y, task) {
        const cy = y + CONFIG.rowHeight / 2; const s = CONFIG.milestoneSize;
        ctx.fillStyle = CONFIG.colors.milestone;
        ctx.beginPath(); ctx.moveTo(x, cy - s / 2); ctx.lineTo(x + s / 2, cy); ctx.lineTo(x, cy + s / 2); ctx.lineTo(x - s / 2, cy); ctx.closePath(); ctx.fill();
        ctx.strokeStyle = CONFIG.colors.milestoneBorder; ctx.lineWidth = 1; ctx.stroke();
        if (dayWidth >= 12) { ctx.fillStyle = CONFIG.colors.text; ctx.font = `500 ${CONFIG.fontSize - 1}px ${CONFIG.fontFamily}`; ctx.textBaseline = 'middle'; ctx.fillText(task.name, x + s + 2, cy); }
    }

    function drawDependencyLinks() {
        ctx.lineWidth = 1.5;
        for (let i = 0; i < tasks.length; i++) {
            const task = tasks[i]; if (!task.predecessors) continue;
            for (const pred of task.predecessors) {
                const pi = tasks.findIndex(t => t.uid === pred.predecessorUID); if (pi === -1) continue;
                const pt = tasks[pi]; const isCritLink = showCritical && task.critical && pt.critical;
                ctx.strokeStyle = isCritLink ? 'rgba(239,68,68,0.4)' : CONFIG.colors.link;

                const typeName = pred.typeName || getTypeName(pred.type);
                let fromX, fromY, toX, toY;
                const predStartDay = daysBetween(projectStart, new Date(pt.start));
                const predEndDay = daysBetween(projectStart, new Date(pt.finish));
                const taskStartDay = daysBetween(projectStart, new Date(task.start));
                const taskEndDay = daysBetween(projectStart, new Date(task.finish));

                switch (typeName) {
                    case 'SS': fromX = predStartDay * dayWidth; toX = taskStartDay * dayWidth; break;
                    case 'FF': fromX = predEndDay * dayWidth; toX = taskEndDay * dayWidth; break;
                    case 'SF': fromX = predStartDay * dayWidth; toX = taskEndDay * dayWidth; break;
                    default: fromX = predEndDay * dayWidth; toX = taskStartDay * dayWidth; // FS
                }
                fromY = pi * CONFIG.rowHeight + CONFIG.rowHeight / 2;
                toY = i * CONFIG.rowHeight + CONFIG.rowHeight / 2;

                ctx.beginPath(); ctx.moveTo(fromX, fromY);
                const midX = fromX + 8;
                ctx.lineTo(midX, fromY); ctx.lineTo(midX, toY); ctx.lineTo(toX, toY); ctx.stroke();

                ctx.fillStyle = isCritLink ? 'rgba(239,68,68,0.6)' : CONFIG.colors.linkArrow;
                ctx.beginPath(); ctx.moveTo(toX, toY); ctx.lineTo(toX - 5, toY - 3); ctx.lineTo(toX - 5, toY + 3); ctx.closePath(); ctx.fill();
            }
        }
    }

    function renderHeader() {
        if (!headerEl) return;
        dayWidth = ZOOM_LEVELS[currentZoom].dayWidth;
        const totalDays = daysBetween(projectStart, projectEnd);
        let html = '<div class="gantt-time-header"><div class="gantt-header-row">';
        let currentMonth = -1, monthStartX = 0, monthLabel = '';
        for (let d = 0; d <= totalDays; d++) {
            const date = new Date(projectStart); date.setDate(date.getDate() + d);
            if (date.getMonth() !== currentMonth) {
                if (currentMonth !== -1) { const w = d * dayWidth - monthStartX; html += `<div class="gantt-header-cell month" style="width:${w}px;min-width:${w}px">${monthLabel}</div>`; }
                currentMonth = date.getMonth(); monthStartX = d * dayWidth;
                monthLabel = date.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
            }
        }
        const lastW = totalDays * dayWidth - monthStartX;
        if (lastW > 0) html += `<div class="gantt-header-cell month" style="width:${lastW}px;min-width:${lastW}px">${monthLabel}</div>`;
        html += '</div><div class="gantt-header-row">';
        if (dayWidth >= 20) {
            for (let d = 0; d < totalDays; d++) { const date = new Date(projectStart); date.setDate(date.getDate() + d); const isWe = date.getDay() === 0 || date.getDay() === 6; html += `<div class="gantt-header-cell" style="width:${dayWidth}px;min-width:${dayWidth}px;${isWe ? 'opacity:0.4' : ''}">${date.getDate()}</div>`; }
        } else if (dayWidth >= 8) {
            let ws = 0;
            for (let d = 0; d < totalDays; d++) { const date = new Date(projectStart); date.setDate(date.getDate() + d); if (date.getDay() === 1 || d === 0) { if (d > 0) { const w = (d - ws) * dayWidth; html += `<div class="gantt-header-cell" style="width:${w}px;min-width:${w}px">${new Date(projectStart.getTime() + ws * 86400000).getDate()}</div>`; } ws = d; } }
            const w = (totalDays - ws) * dayWidth; html += `<div class="gantt-header-cell" style="width:${w}px;min-width:${w}px">${new Date(projectStart.getTime() + ws * 86400000).getDate()}</div>`;
        } else { html += `<div class="gantt-header-cell" style="flex:1"></div>`; }
        html += '</div></div>';
        headerEl.innerHTML = html; headerEl.scrollLeft = scrollX;
    }

    // ─── Drag & Drop ───
    function handleMouseDown(e) {
        const rect = canvas.getBoundingClientRect();
        const mx = e.clientX - rect.left; const my = e.clientY - rect.top;
        const idx = Math.floor(my / CONFIG.rowHeight);
        if (idx < 0 || idx >= tasks.length) return;
        const task = tasks[idx];
        if (task.summary || task.milestone) return;

        const startDay = daysBetween(projectStart, new Date(task.start));
        const endDay = daysBetween(projectStart, new Date(task.finish));
        const barX = startDay * dayWidth;
        const barW = Math.max((endDay - startDay) * dayWidth, dayWidth);
        const barY = idx * CONFIG.rowHeight + (CONFIG.rowHeight - CONFIG.barHeight) / 2;

        // Check if click is on bar
        if (mx >= barX && mx <= barX + barW && my >= barY && my <= barY + CONFIG.barHeight) {
            isDragging = true;
            dragTaskIndex = idx;
            dragStartX = mx;
            dragStartDate = new Date(task.start);
            dragOrigDuration = task.durationDays;

            if (mx >= barX + barW - 4) {
                dragType = 'resize-right';
                canvas.style.cursor = 'ew-resize';
            } else {
                dragType = 'move';
                canvas.style.cursor = 'grabbing';
            }
            e.preventDefault();
        }
    }

    function handleMouseMove(e) {
        const rect = canvas.getBoundingClientRect();
        const mx = e.clientX - rect.left; const my = e.clientY - rect.top;

        if (isDragging && dragTaskIndex >= 0) {
            const task = tasks[dragTaskIndex];
            const dx = mx - dragStartX;
            const daysDelta = Math.round(dx / dayWidth);

            if (dragType === 'move') {
                const newStart = new Date(dragStartDate);
                newStart.setDate(newStart.getDate() + daysDelta);
                const newFinish = new Date(newStart);
                newFinish.setDate(newFinish.getDate() + dragOrigDuration);
                task.start = newStart;
                task.finish = newFinish;
            } else if (dragType === 'resize-right') {
                const newDuration = Math.max(1, dragOrigDuration + daysDelta);
                task.durationDays = newDuration;
                const newFinish = new Date(task.start);
                newFinish.setDate(newFinish.getDate() + newDuration);
                task.finish = newFinish;
            }
            render();
            return;
        }

        // Hover detection — only re-render when hover state changes
        const prevHover = hoveredTaskIndex;
        const idx = Math.floor(my / CONFIG.rowHeight);
        if (idx >= 0 && idx < tasks.length) {
            hoveredTaskIndex = idx;
            const task = tasks[idx];
            if (!task.summary && !task.milestone) {
                const startDay = daysBetween(projectStart, new Date(task.start));
                const endDay = daysBetween(projectStart, new Date(task.finish));
                const barX = startDay * dayWidth;
                const barW = Math.max((endDay - startDay) * dayWidth, dayWidth);
                const barY = idx * CONFIG.rowHeight + (CONFIG.rowHeight - CONFIG.barHeight) / 2;
                if (mx >= barX && mx <= barX + barW && my >= barY && my <= barY + CONFIG.barHeight) {
                    canvas.style.cursor = mx >= barX + barW - 4 ? 'ew-resize' : 'grab';
                } else { canvas.style.cursor = 'pointer'; }
            } else { canvas.style.cursor = 'pointer'; }
        } else { hoveredTaskIndex = -1; canvas.style.cursor = 'default'; }
        if (prevHover !== hoveredTaskIndex) requestAnimationFrame(() => render());
    }

    function handleMouseUp(e) {
        if (isDragging && dragTaskIndex >= 0 && onTaskUpdate) {
            onTaskUpdate(tasks[dragTaskIndex], dragTaskIndex);
        }
        isDragging = false; dragType = null; dragTaskIndex = -1;
        canvas.style.cursor = 'default';
    }

    function handleClick(e) {
        if (isDragging) return;
        const rect = canvas.getBoundingClientRect();
        const y = e.clientY - rect.top;
        const idx = Math.floor(y / CONFIG.rowHeight);
        if (idx >= 0 && idx < tasks.length) {
            selectedTaskIndex = idx; render();
            if (onTaskSelect) onTaskSelect(tasks[idx], idx);
        }
    }

    function zoomIn() { if (currentZoom > 0) { currentZoom--; resize(); render(); renderHeader(); } return ZOOM_LEVELS[currentZoom].name; }
    function zoomOut() { if (currentZoom < ZOOM_LEVELS.length - 1) { currentZoom++; resize(); render(); renderHeader(); } return ZOOM_LEVELS[currentZoom].name; }
    function getZoomLevel() { return ZOOM_LEVELS[currentZoom].name; }

    function daysBetween(d1, d2) { const t1 = new Date(d1); t1.setHours(0,0,0,0); const t2 = new Date(d2); t2.setHours(0,0,0,0); return Math.round((t2 - t1) / 86400000); }
    function getTypeName(t) { return ['FF','FS','SF','SS'][t] || 'FS'; }
    function roundRect(ctx, x, y, w, h, r) { if (w < 2*r) r=w/2; if (h < 2*r) r=h/2; ctx.beginPath(); ctx.moveTo(x+r,y); ctx.arcTo(x+w,y,x+w,y+h,r); ctx.arcTo(x+w,y+h,x,y+h,r); ctx.arcTo(x,y+h,x,y,r); ctx.arcTo(x,y,x+w,y,r); ctx.closePath(); }

    export const GanttChart = { init, update, render, resize, zoomIn, zoomOut, getZoomLevel, setOptions, cleanup };
