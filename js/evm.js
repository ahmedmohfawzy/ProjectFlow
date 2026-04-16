/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 * Proprietary Software — Unauthorized use prohibited.
 * https://www.linkedin.com/in/ahmed-m-fawzy
 */
/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — Earned Value Management (EVM) Engine
 * PV, EV, AC, SPI, CPI, SV, CV, EAC, ETC, VAC
 * ═══════════════════════════════════════════════════════
 */


    const MS_PER_DAY = 86400000;

    /**
     * Compute EVM metrics for a project
     * @param {Object} project - The project data
     * @returns {Object} EVM metrics
     */
    function compute(project) {
        if (!project || !project.tasks) return null;
        const tasks = project.tasks.filter(t => !t.summary);
        if (tasks.length === 0) return null;

        const today = new Date(); today.setHours(0, 0, 0, 0);
        const pStart = new Date(project.startDate);
        const pFinish = new Date(project.finishDate);

        // BAC: Budget at Completion (total planned cost)
        const BAC = tasks.reduce((s, t) => s + (t.cost || 0), 0);

        // If no costs assigned, use task-count weighting (equal weight per task)
        const useCostWeight = BAC > 0;
        const totalWeight = useCostWeight ? BAC : tasks.length;

        // PV: Planned Value — how much work SHOULD be done by today
        let PV = 0;
        tasks.forEach(t => {
            const ts = new Date(t.baselineStart || t.start);
            const tf = new Date(t.baselineFinish || t.finish);
            const dur = Math.max(1, (tf - ts) / MS_PER_DAY);
            const elapsed = Math.max(0, (today - ts) / MS_PER_DAY);
            const plannedPct = Math.min(1, elapsed / dur);
            const weight = useCostWeight ? (t.cost || 0) : 1;
            PV += weight * plannedPct;
        });

        // EV: Earned Value — how much work IS actually done
        let EV = 0;
        tasks.forEach(t => {
            const weight = useCostWeight ? (t.cost || 0) : 1;
            EV += weight * ((t.percentComplete || 0) / 100);
        });

        // AC: Actual Cost — how much has been spent
        // Use task.actualCost if available, otherwise estimate from % × cost
        let AC = 0;
        tasks.forEach(t => {
            if (t.actualCost != null) {
                AC += t.actualCost;
            } else {
                AC += (t.cost || 0) * ((t.percentComplete || 0) / 100);
            }
        });

        // Normalize for non-cost mode
        if (!useCostWeight) {
            PV = (PV / totalWeight) * 100;
            EV = (EV / totalWeight) * 100;
            // AC based on elapsed time (not equal to EV) — gives meaningful CPI
            let timeBasedAC = 0;
            const today = new Date(); today.setHours(0, 0, 0, 0);
            tasks.forEach(t => {
                const ts = new Date(t.start);
                const tf = new Date(t.finish);
                const dur = Math.max(1, (tf - ts) / 86400000);
                const elapsed = Math.max(0, Math.min(dur, (today - ts) / 86400000));
                timeBasedAC += elapsed / dur;
            });
            AC = (timeBasedAC / totalWeight) * 100;
        }

        // ── Indices ──
        const SV = EV - PV;                              // Schedule Variance
        const CV = EV - AC;                              // Cost Variance
        const SPI = PV > 0 ? EV / PV : 1;               // Schedule Performance Index
        const CPI = AC > 0 ? EV / AC : 1;               // Cost Performance Index

        // ── Forecasts ──
        const effectiveBAC = useCostWeight ? BAC : 100;
        const EAC = CPI > 0 ? effectiveBAC / CPI : effectiveBAC;   // Estimate at Completion
        const ETC = Math.max(0, EAC - AC);                          // Estimate to Complete
        const VAC = effectiveBAC - EAC;                              // Variance at Completion
        const TCPI = (effectiveBAC - AC) === 0 
            ? ((effectiveBAC - EV) > 0 ? Infinity : 1) 
            : ((effectiveBAC - EV) > 0 ? (effectiveBAC - EV) / (effectiveBAC - AC) : 1); // To-Complete Performance Index

        // ── Time-phased data for chart ──
        const timeData = computeTimePhasedEVM(project, tasks, useCostWeight);

        // ── Health indicators ──
        const scheduleHealth = SPI >= 1.0 ? 'good' : SPI >= 0.9 ? 'warning' : 'danger';
        const costHealth = CPI >= 1.0 ? 'good' : CPI >= 0.9 ? 'warning' : 'danger';

        return {
            BAC: effectiveBAC, PV, EV, AC, SV, CV, SPI, CPI,
            EAC, ETC, VAC, TCPI,
            scheduleHealth, costHealth,
            timeData,
            hasCosts: useCostWeight
        };
    }

    /**
     * Compute time-phased EVM data for charting
     */
    function computeTimePhasedEVM(project, tasks, useCostWeight) {
        const pStart = new Date(project.startDate);
        const pFinish = new Date(project.finishDate);
        const today = new Date(); today.setHours(0, 0, 0, 0);
        const totalDays = Math.max(1, Math.round((pFinish - pStart) / MS_PER_DAY));
        const numPoints = Math.min(totalDays, 50);

        const pvData = [], evData = [], acData = [];
        const totalWeight = useCostWeight
            ? tasks.reduce((s, t) => s + (t.cost || 0), 0) || 1
            : tasks.length || 1;

        for (let i = 0; i <= numPoints; i++) {
            const day = Math.round((i / numPoints) * totalDays);
            const d = new Date(pStart); d.setDate(d.getDate() + day);

            // PV at this date
            let pv = 0;
            tasks.forEach(t => {
                const ts = new Date(t.baselineStart || t.start);
                const tf = new Date(t.baselineFinish || t.finish);
                const dur = Math.max(1, (tf - ts) / MS_PER_DAY);
                const elapsed = Math.max(0, (d - ts) / MS_PER_DAY);
                const w = useCostWeight ? (t.cost || 0) : 1;
                pv += w * Math.min(1, elapsed / dur);
            });
            pvData.push({ day, value: (pv / totalWeight) * 100 });

            // EV and AC only up to today
            if (d <= today) {
                let ev = 0, ac = 0;
                tasks.forEach(t => {
                    const ts = new Date(t.start);
                    const tf = new Date(t.finish);
                    const dur = Math.max(1, (tf - ts) / MS_PER_DAY);
                    const elapsed = Math.max(0, (d - ts) / MS_PER_DAY);
                    const timeRatio = Math.min(1, elapsed / dur);
                    const w = useCostWeight ? (t.cost || 0) : 1;
                    // Scale actual progress based on time position
                    const actualPct = Math.min((t.percentComplete || 0) / 100, timeRatio);
                    ev += w * actualPct;
                    ac += w * actualPct; // Simplified
                });
                evData.push({ day, value: (ev / totalWeight) * 100 });
                acData.push({ day, value: (ac / totalWeight) * 100 });
            }
        }

        return { pvData, evData, acData, totalDays };
    }

    /**
     * Draw EVM Chart on canvas
     */
    function drawEVMChart(canvas, timeData) {
        if (!canvas || !timeData) return;
        const ctx = canvas.getContext('2d');
        const dpr = window.devicePixelRatio || 1;
        const w = canvas.clientWidth; const h = canvas.clientHeight;
        canvas.width = w * dpr; canvas.height = h * dpr;
        ctx.scale(dpr, dpr);

        const pad = { top: 28, right: 16, bottom: 30, left: 40 };
        const cW = w - pad.left - pad.right;
        const cH = h - pad.top - pad.bottom;

        // Read theme colors
        const cs = getComputedStyle(document.documentElement);
        const bgColor = cs.getPropertyValue('--bg-secondary').trim() || '#161822';
        const gridColor = cs.getPropertyValue('--border-subtle').trim() || 'rgba(255,255,255,0.08)';
        const labelColor = cs.getPropertyValue('--text-muted').trim() || '#5c6378';
        const legendColor = cs.getPropertyValue('--text-secondary').trim() || '#9aa0b4';

        // Background
        ctx.fillStyle = bgColor;
        ctx.fillRect(0, 0, w, h);

        // Axes
        ctx.strokeStyle = gridColor;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(pad.left, pad.top);
        ctx.lineTo(pad.left, pad.top + cH);
        ctx.lineTo(pad.left + cW, pad.top + cH);
        ctx.stroke();

        // Y labels
        ctx.fillStyle = labelColor;
        ctx.font = '10px Inter, sans-serif';
        ctx.textAlign = 'right';
        for (let pct = 0; pct <= 100; pct += 25) {
            const y = pad.top + cH - (pct / 100) * cH;
            ctx.fillText(pct + '%', pad.left - 6, y + 3);
            ctx.strokeStyle = gridColor;
            ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(pad.left + cW, y); ctx.stroke();
        }

        const { pvData, evData, acData, totalDays } = timeData;

        // Draw lines
        const drawLine = (data, color, dash) => {
            if (data.length < 2) return;
            ctx.strokeStyle = color;
            ctx.lineWidth = 2.5;
            ctx.setLineDash(dash || []);
            ctx.beginPath();
            data.forEach((p, i) => {
                const x = pad.left + (p.day / totalDays) * cW;
                const y = pad.top + cH - (p.value / 100) * cH;
                if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
            });
            ctx.stroke();
            ctx.setLineDash([]);
        };

        drawLine(pvData, '#8b5cf6', [6, 3]); // PV — purple dashed
        drawLine(evData, '#22c55e', []);      // EV — green solid
        drawLine(acData, '#ef4444', [3, 3]);  // AC — red dotted

        // Legend
        const legY = 12;
        ctx.font = '10px Inter, sans-serif';
        ctx.textAlign = 'left';
        const items = [
            { label: 'PV (Planned)', color: '#8b5cf6', x: pad.left },
            { label: 'EV (Earned)', color: '#22c55e', x: pad.left + 90 },
            { label: 'AC (Actual)', color: '#ef4444', x: pad.left + 170 }
        ];
        items.forEach(it => {
            ctx.fillStyle = it.color;
            ctx.fillRect(it.x, legY - 3, 12, 3);
            ctx.fillStyle = legendColor;
            ctx.fillText(it.label, it.x + 16, legY + 1);
        });
    }

    /**
     * Format EVM metric for display
     */
    function fmt(value, type = 'number') {
        if (value == null || !isFinite(value)) return '—';
        switch (type) {
            case 'currency': return '$' + Math.round(value).toLocaleString();
            case 'index': return value.toFixed(2);
            case 'percent': return Math.round(value) + '%';
            default: return Math.round(value).toLocaleString();
        }
    }

    export const EVMEngine = { compute, drawEVMChart, fmt };
