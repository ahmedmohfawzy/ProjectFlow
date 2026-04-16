/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 * Proprietary Software — Unauthorized use prohibited.
 * https://www.linkedin.com/in/ahmed-m-fawzy
 */
/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — Dashboard Engine
 * KPI Cards, Donut/Bar/S-Curve charts (pure Canvas)
 * Notifications system
 * ═══════════════════════════════════════════════════════
 */



    // ─── Colors ───
    const COLORS = {
        bg: '#161822', card: '#1c1f2e', border: 'rgba(255,255,255,0.06)',
        text: '#e8eaed', textMuted: '#9aa0b4', textDim: '#5c6378',
        accent: '#6366f1', accentLight: '#818cf8', purple: '#8b5cf6',
        success: '#22c55e', warning: '#f59e0b', danger: '#ef4444', info: '#3b82f6',
        notStarted: '#64748b', inProgress: '#3b82f6', complete: '#22c55e', late: '#ef4444',
        baseline: '#8b5cf6', actual: '#22c55e',
    };

    let _notifications = [];

    // ═══════════════════════════════
    // KPI COMPUTATION
    // ═══════════════════════════════
    function computeKPIs(project) {
        if (!project || !project.tasks) return null;
        const tasks = project.tasks.filter(t => !t.summary);
        const total = tasks.length;
        if (total === 0) return null;

        const complete = tasks.filter(t => t.percentComplete >= 100).length;
        const inProgress = tasks.filter(t => t.percentComplete > 0 && t.percentComplete < 100).length;
        const notStarted = tasks.filter(t => t.percentComplete === 0).length;
        const late = tasks.filter(t => t.status === 'late').length;
        const atRisk = tasks.filter(t => t.status === 'at-risk').length;
        const critical = tasks.filter(t => t.critical).length;
        const milestones = tasks.filter(t => t.milestone).length;
        const milestonesComplete = tasks.filter(t => t.milestone && t.percentComplete >= 100).length;

        const totalPct = tasks.reduce((s, t) => s + (t.percentComplete || 0), 0);
        const overallProgress = Math.round(totalPct / total);

        const totalCost = tasks.reduce((s, t) => s + (t.cost || 0), 0);
        const earnedCost = tasks.reduce((s, t) => s + ((t.cost || 0) * (t.percentComplete || 0) / 100), 0);

        const startDate = new Date(project.startDate);
        const finishDate = new Date(project.finishDate);
        const totalDays = Math.max(1, Math.round((finishDate - startDate) / 86400000));
        const elapsed = Math.round((Date.now() - startDate.getTime()) / 86400000);
        const timeProgress = Math.min(100, Math.max(0, Math.round((elapsed / totalDays) * 100)));

        const daysRemaining = Math.max(0, Math.round((finishDate.getTime() - Date.now()) / 86400000));

        // Phase progress (summary tasks)
        const phases = project.tasks.filter(t => t.summary && t.outlineLevel === 1);
        const phaseData = phases.map(p => {
            const children = project.tasks.filter(t => !t.summary && t.outlineLevel > p.outlineLevel);
            // Simple: get tasks between this summary and next
            const idx = project.tasks.indexOf(p);
            const nextSummaryIdx = project.tasks.findIndex((t, i) => i > idx && t.summary && t.outlineLevel <= p.outlineLevel);
            const end = nextSummaryIdx === -1 ? project.tasks.length : nextSummaryIdx;
            const phaseTasks = project.tasks.slice(idx + 1, end).filter(t => !t.summary);
            const avg = phaseTasks.length > 0 ? Math.round(phaseTasks.reduce((s, t) => s + (t.percentComplete || 0), 0) / phaseTasks.length) : 0;
            return { name: p.name, progress: avg, taskCount: phaseTasks.length, critical: phaseTasks.filter(t => t.critical).length };
        });

        return {
            total, complete, inProgress, notStarted, late, atRisk, critical,
            milestones, milestonesComplete, overallProgress, totalCost, earnedCost,
            startDate, finishDate, totalDays, elapsed, timeProgress, daysRemaining,
            phases: phaseData
        };
    }

    // ═══════════════════════════════
    // NOTIFICATIONS
    // ═══════════════════════════════
    function generateNotifications(project) {
        if (!project) return [];
        _notifications = [];
        const today = new Date(); today.setHours(0, 0, 0, 0);
        const tomorrow = new Date(today); tomorrow.setDate(tomorrow.getDate() + 1);
        const nextWeek = new Date(today); nextWeek.setDate(nextWeek.getDate() + 7);

        const tasks = project.tasks.filter(t => !t.summary);

        // Late tasks
        tasks.filter(t => t.status === 'late').forEach(t => {
            const days = Math.round((today - new Date(t.finish)) / 86400000);
            _notifications.push({ type: 'danger', icon: '🔴', title: `"${t.name}" is ${days}d late`, subtitle: `Due: ${fmtDate(t.finish)}`, taskUid: t.uid });
        });

        // Due today
        tasks.filter(t => {
            const f = new Date(t.finish); f.setHours(0, 0, 0, 0);
            return f.getTime() === today.getTime() && t.percentComplete < 100;
        }).forEach(t => {
            _notifications.push({ type: 'warning', icon: '⏰', title: `"${t.name}" is due today`, subtitle: `${t.percentComplete}% complete`, taskUid: t.uid });
        });

        // Due this week
        tasks.filter(t => {
            const f = new Date(t.finish); f.setHours(0, 0, 0, 0);
            return f > today && f <= nextWeek && t.percentComplete < 100;
        }).forEach(t => {
            const days = Math.round((new Date(t.finish) - today) / 86400000);
            _notifications.push({ type: 'info', icon: '📅', title: `"${t.name}" due in ${days}d`, subtitle: `${t.percentComplete}% complete`, taskUid: t.uid });
        });

        // Over-allocated resources
        if (project.resources) {
            // Simple check
        }

        // Milestones approaching
        tasks.filter(t => t.milestone && t.percentComplete < 100).forEach(t => {
            const days = Math.round((new Date(t.finish) - today) / 86400000);
            if (days >= 0 && days <= 7) {
                _notifications.push({ type: 'warning', icon: '⭐', title: `Milestone "${t.name}" in ${days}d`, subtitle: fmtDate(t.finish), taskUid: t.uid });
            }
        });

        // Critical tasks not started
        tasks.filter(t => t.critical && t.percentComplete === 0 && new Date(t.start) <= today).forEach(t => {
            _notifications.push({ type: 'danger', icon: '⚠️', title: `Critical task "${t.name}" not started`, subtitle: `Start was: ${fmtDate(t.start)}`, taskUid: t.uid });
        });

        // Upcoming holidays
        if (typeof WorkCalendar !== 'undefined') {
            const cfg = WorkCalendar.getConfig();
            const year = today.getFullYear();
            const holidays = cfg.holidays?.[year] || {};
            Object.values(holidays).forEach(h => {
                const hDate = new Date(h.date); hDate.setHours(0, 0, 0, 0);
                const diff = Math.round((hDate - today) / 86400000);
                if (diff >= 0 && diff <= 7) {
                    _notifications.push({ type: 'info', icon: h.flag || '📅', title: `Holiday in ${diff}d: ${h.localName || h.name}`, subtitle: h.date });
                }
            });
        }

        return _notifications;
    }

    function getNotifications() { return _notifications; }
    function getNotificationCount() { return _notifications.length; }
    function dismissNotification(index) { _notifications.splice(index, 1); }

    // ═══════════════════════════════
    // CHART DRAWING (Pure Canvas)
    // ═══════════════════════════════

    /**
     * Draw Donut Chart
     */
    function drawDonut(canvas, data, options = {}) {
        const ctx = canvas.getContext('2d');
        const dpr = window.devicePixelRatio || 1;
        const w = canvas.clientWidth; const h = canvas.clientHeight;
        canvas.width = w * dpr; canvas.height = h * dpr;
        ctx.scale(dpr, dpr);

        const cx = w / 2, cy = h / 2;
        const outerR = Math.min(cx, cy) - 8;
        const innerR = outerR * 0.62;
        const total = data.reduce((s, d) => s + d.value, 0);

        if (total === 0) { drawEmptyCircle(ctx, cx, cy, outerR, innerR); return; }

        let startAngle = -Math.PI / 2;
        for (const d of data) {
            const sweep = (d.value / total) * Math.PI * 2;
            ctx.beginPath();
            ctx.arc(cx, cy, outerR, startAngle, startAngle + sweep);
            ctx.arc(cx, cy, innerR, startAngle + sweep, startAngle, true);
            ctx.closePath();
            ctx.fillStyle = d.color;
            ctx.fill();
            startAngle += sweep;
        }

        // Center text
        ctx.fillStyle = COLORS.text;
        ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
        ctx.font = `700 ${outerR * 0.38}px Inter, sans-serif`;
        ctx.fillText(options.centerText || '', cx, cy - 4);
        ctx.font = `400 ${outerR * 0.16}px Inter, sans-serif`;
        ctx.fillStyle = COLORS.textMuted;
        ctx.fillText(options.centerSub || '', cx, cy + outerR * 0.22);
    }

    /**
     * Draw Horizontal Bar Chart (Phase progress)
     */
    function drawBars(canvas, data) {
        const ctx = canvas.getContext('2d');
        const dpr = window.devicePixelRatio || 1;
        const w = canvas.clientWidth; const h = canvas.clientHeight;
        canvas.width = w * dpr; canvas.height = h * dpr;
        ctx.scale(dpr, dpr);

        if (!data || data.length === 0) {
            ctx.fillStyle = COLORS.textDim;
            ctx.font = '12px Inter, sans-serif';
            ctx.textAlign = 'center';
            ctx.fillText('No phases to show', w / 2, h / 2);
            return;
        }

        const padding = { top: 8, right: 50, bottom: 8, left: 10 };
        const barH = Math.min(28, (h - padding.top - padding.bottom) / data.length - 6);
        const maxNameW = 120;
        const barAreaX = padding.left + maxNameW + 8;
        const barAreaW = w - barAreaX - padding.right;

        data.forEach((d, i) => {
            const y = padding.top + i * (barH + 6);

            // Name
            ctx.fillStyle = COLORS.textMuted;
            ctx.font = '11px Inter, sans-serif';
            ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
            const name = d.name.length > 18 ? d.name.substring(0, 16) + '…' : d.name;
            ctx.fillText(name, barAreaX - 8, y + barH / 2);

            // Background bar
            ctx.fillStyle = 'rgba(255,255,255,0.04)';
            roundRect(ctx, barAreaX, y, barAreaW, barH, 4);
            ctx.fill();

            // Progress bar
            const fillW = Math.max(0, (d.progress / 100) * barAreaW);
            if (fillW > 0) {
                const grad = ctx.createLinearGradient(barAreaX, 0, barAreaX + fillW, 0);
                grad.addColorStop(0, d.progress >= 100 ? COLORS.success : COLORS.accent);
                grad.addColorStop(1, d.progress >= 100 ? '#16a34a' : COLORS.purple);
                ctx.fillStyle = grad;
                roundRect(ctx, barAreaX, y, fillW, barH, 4);
                ctx.fill();
            }

            // Percentage
            ctx.fillStyle = COLORS.text;
            ctx.font = '600 11px Inter, sans-serif';
            ctx.textAlign = 'left';
            ctx.fillText(d.progress + '%', barAreaX + barAreaW + 6, y + barH / 2);
        });
    }

    /**
     * Draw S-Curve (Planned vs Actual)
     */
    function drawSCurve(canvas, project) {
        const ctx = canvas.getContext('2d');
        const dpr = window.devicePixelRatio || 1;
        const w = canvas.clientWidth; const h = canvas.clientHeight;
        canvas.width = w * dpr; canvas.height = h * dpr;
        ctx.scale(dpr, dpr);

        if (!project || !project.tasks) return;

        const tasks = project.tasks.filter(t => !t.summary);
        if (tasks.length === 0) return;

        const start = new Date(project.startDate);
        const finish = new Date(project.finishDate);
        const totalDays = Math.max(1, Math.round((finish - start) / 86400000));
        const today = new Date();

        const pad = { top: 30, right: 20, bottom: 40, left: 45 };
        const chartW = w - pad.left - pad.right;
        const chartH = h - pad.top - pad.bottom;

        // Axes
        ctx.strokeStyle = COLORS.border;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(pad.left, pad.top);
        ctx.lineTo(pad.left, pad.top + chartH);
        ctx.lineTo(pad.left + chartW, pad.top + chartH);
        ctx.stroke();

        // Y labels
        ctx.fillStyle = COLORS.textDim;
        ctx.font = '10px Inter, sans-serif';
        ctx.textAlign = 'right';
        for (let pct = 0; pct <= 100; pct += 25) {
            const y = pad.top + chartH - (pct / 100) * chartH;
            ctx.fillText(pct + '%', pad.left - 6, y + 3);
            if (pct > 0 && pct < 100) {
                ctx.strokeStyle = 'rgba(255,255,255,0.04)';
                ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(pad.left + chartW, y); ctx.stroke();
            }
        }

        // X labels
        ctx.textAlign = 'center';
        const xSteps = Math.min(8, totalDays);
        for (let i = 0; i <= xSteps; i++) {
            const day = Math.round((i / xSteps) * totalDays);
            const x = pad.left + (day / totalDays) * chartW;
            const d = new Date(start); d.setDate(d.getDate() + day);
            ctx.fillText(fmtShort(d), x, pad.top + chartH + 16);
        }

        // Compute planned S-curve (cumulative)
        const planned = [];
        const actual = [];
        const numPoints = Math.min(totalDays, 60);

        for (let i = 0; i <= numPoints; i++) {
            const day = Math.round((i / numPoints) * totalDays);
            const d = new Date(start); d.setDate(d.getDate() + day);

            // Planned: tasks that should be done by this date (based on finish date)
            let plannedPct = 0;
            tasks.forEach(t => {
                const tf = new Date(t.baselineFinish || t.finish);
                if (tf <= d) plannedPct += 100;
                else {
                    const ts = new Date(t.baselineStart || t.start);
                    const taskDur = Math.max(1, (tf - ts) / 86400000);
                    const elapsed = Math.max(0, (d - ts) / 86400000);
                    plannedPct += Math.min(100, (elapsed / taskDur) * 100);
                }
            });
            planned.push({ day, pct: plannedPct / tasks.length });

            // Actual: real progress up to today
            if (d <= today) {
                let actualPct = 0;
                tasks.forEach(t => {
                    const ts = new Date(t.start);
                    const tf = new Date(t.finish);
                    const taskDur = Math.max(1, (tf - ts) / 86400000);
                    const elapsed = Math.max(0, (d - ts) / 86400000);
                    const timeRatio = Math.min(1, elapsed / taskDur);
                    actualPct += Math.min(t.percentComplete, timeRatio * 100);
                });
                actual.push({ day, pct: actualPct / tasks.length });
            }
        }

        // Draw planned line
        drawCurveLine(ctx, planned, totalDays, chartW, chartH, pad, COLORS.baseline, 'Planned', true);

        // Draw actual line
        if (actual.length > 1) {
            drawCurveLine(ctx, actual, totalDays, chartW, chartH, pad, COLORS.actual, 'Actual', false);
        }

        // Today line
        const todayDay = Math.round((today - start) / 86400000);
        if (todayDay > 0 && todayDay < totalDays) {
            const tx = pad.left + (todayDay / totalDays) * chartW;
            ctx.strokeStyle = COLORS.danger;
            ctx.lineWidth = 1;
            ctx.setLineDash([4, 3]);
            ctx.beginPath(); ctx.moveTo(tx, pad.top); ctx.lineTo(tx, pad.top + chartH); ctx.stroke();
            ctx.setLineDash([]);
            ctx.fillStyle = COLORS.danger;
            ctx.font = '600 9px Inter, sans-serif';
            ctx.textAlign = 'center';
            ctx.fillText('Today', tx, pad.top - 6);
        }

        // Legend
        const legY = pad.top - 14;
        ctx.font = '10px Inter, sans-serif';
        // Planned
        ctx.fillStyle = COLORS.baseline;
        ctx.fillRect(pad.left + chartW - 130, legY - 4, 12, 3);
        ctx.fillStyle = COLORS.textMuted;
        ctx.textAlign = 'left';
        ctx.fillText('Planned', pad.left + chartW - 114, legY);
        // Actual
        ctx.fillStyle = COLORS.actual;
        ctx.fillRect(pad.left + chartW - 56, legY - 4, 12, 3);
        ctx.fillStyle = COLORS.textMuted;
        ctx.fillText('Actual', pad.left + chartW - 40, legY);
    }

    function drawCurveLine(ctx, points, totalDays, chartW, chartH, pad, color, label, dashed) {
        if (points.length < 2) return;
        ctx.strokeStyle = color;
        ctx.lineWidth = 2.5;
        ctx.lineJoin = 'round';
        if (dashed) ctx.setLineDash([6, 3]); else ctx.setLineDash([]);

        ctx.beginPath();
        points.forEach((p, i) => {
            const x = pad.left + (p.day / totalDays) * chartW;
            const y = pad.top + chartH - (p.pct / 100) * chartH;
            if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
        });
        ctx.stroke();
        ctx.setLineDash([]);

        // Fill under
        ctx.globalAlpha = 0.06;
        ctx.fillStyle = color;
        ctx.beginPath();
        points.forEach((p, i) => {
            const x = pad.left + (p.day / totalDays) * chartW;
            const y = pad.top + chartH - (p.pct / 100) * chartH;
            if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
        });
        const lastX = pad.left + (points[points.length - 1].day / totalDays) * chartW;
        ctx.lineTo(lastX, pad.top + chartH);
        ctx.lineTo(pad.left + (points[0].day / totalDays) * chartW, pad.top + chartH);
        ctx.closePath();
        ctx.fill();
        ctx.globalAlpha = 1;
    }

    // ─── Mini Timeline ───
    function drawTimeline(canvas, project) {
        const ctx = canvas.getContext('2d');
        const dpr = window.devicePixelRatio || 1;
        const w = canvas.clientWidth; const h = canvas.clientHeight;
        canvas.width = w * dpr; canvas.height = h * dpr;
        ctx.scale(dpr, dpr);

        if (!project) return;
        const start = new Date(project.startDate);
        const finish = new Date(project.finishDate);
        const totalMs = Math.max(1, finish - start);
        const today = new Date();
        const pad = 16;
        const barY = h / 2 - 4; const barH = 8;

        // Track
        ctx.fillStyle = 'rgba(255,255,255,0.06)';
        roundRect(ctx, pad, barY, w - pad * 2, barH, 4);
        ctx.fill();

        // Elapsed
        const elapsed = Math.min(1, Math.max(0, (today - start) / totalMs));
        const grad = ctx.createLinearGradient(pad, 0, pad + (w - pad * 2) * elapsed, 0);
        grad.addColorStop(0, COLORS.accent);
        grad.addColorStop(1, COLORS.purple);
        ctx.fillStyle = grad;
        roundRect(ctx, pad, barY, (w - pad * 2) * elapsed, barH, 4);
        ctx.fill();

        // Milestones
        const milestones = project.tasks.filter(t => t.milestone);
        milestones.forEach(m => {
            const ms = new Date(m.finish);
            const x = pad + ((ms - start) / totalMs) * (w - pad * 2);
            ctx.fillStyle = m.percentComplete >= 100 ? COLORS.success : COLORS.warning;
            ctx.beginPath();
            ctx.moveTo(x, barY - 4); ctx.lineTo(x + 5, barY + barH / 2);
            ctx.lineTo(x, barY + barH + 4); ctx.lineTo(x - 5, barY + barH / 2);
            ctx.closePath(); ctx.fill();
        });

        // Labels
        ctx.fillStyle = COLORS.textDim;
        ctx.font = '10px Inter, sans-serif';
        ctx.textAlign = 'left';
        ctx.fillText(fmtShort(start), pad, barY + barH + 16);
        ctx.textAlign = 'right';
        ctx.fillText(fmtShort(finish), w - pad, barY + barH + 16);
        ctx.textAlign = 'center';
        ctx.fillStyle = COLORS.danger;
        ctx.fillText('▼ Today', pad + (w - pad * 2) * elapsed, barY - 8);
    }

    // ─── Helpers ───
    function roundRect(ctx, x, y, w, h, r) {
        ctx.beginPath();
        ctx.moveTo(x + r, y);
        ctx.lineTo(x + w - r, y);
        ctx.quadraticCurveTo(x + w, y, x + w, y + r);
        ctx.lineTo(x + w, y + h - r);
        ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
        ctx.lineTo(x + r, y + h);
        ctx.quadraticCurveTo(x, y + h, x, y + h - r);
        ctx.lineTo(x, y + r);
        ctx.quadraticCurveTo(x, y, x + r, y);
        ctx.closePath();
    }

    function drawEmptyCircle(ctx, cx, cy, outerR, innerR) {
        ctx.beginPath();
        ctx.arc(cx, cy, outerR, 0, Math.PI * 2);
        ctx.arc(cx, cy, innerR, 0, Math.PI * 2, true);
        ctx.closePath();
        ctx.fillStyle = 'rgba(255,255,255,0.04)';
        ctx.fill();
    }

    function fmtDate(d) { const dt = new Date(d); return dt.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' }); }
    function fmtShort(d) { const dt = new Date(d); return dt.toLocaleDateString('en-US', { month: 'short', day: 'numeric' }); }

    export const Dashboard = {
        computeKPIs, generateNotifications, getNotifications, getNotificationCount,
        dismissNotification, drawDonut, drawBars, drawSCurve, drawTimeline, COLORS
    };
