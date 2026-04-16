/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 * Proprietary Software — Unauthorized use prohibited.
 * https://www.linkedin.com/in/ahmed-m-fawzy
 */
/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — Reports Engine
 * PDF, Excel, Email Summary, Gantt PNG, Print
 * ═══════════════════════════════════════════════════════
 */



    // ═══════════════════════════════
    // EMAIL / CLIPBOARD SUMMARY
    // ═══════════════════════════════
    // ══════════════════════════════════════════════════════════
    //  PROFESSIONAL EXECUTIVE SUMMARY GENERATOR
    // ══════════════════════════════════════════════════════════

    /** Build all data needed for any summary format in one pass */
    function _buildSummaryData(project, settings) {
        const tasks    = project.tasks.filter(t => !t.summary);
        const total    = tasks.length;
        const cur      = settings.currency || '$';
        const today    = new Date(); today.setHours(0,0,0,0);

        const complete   = tasks.filter(t => t.percentComplete >= 100).length;
        const inProg     = tasks.filter(t => t.percentComplete > 0 && t.percentComplete < 100).length;
        const notStarted = tasks.filter(t => t.percentComplete === 0).length;
        const lateTasks  = tasks.filter(t => new Date(t.finish) < today && t.percentComplete < 100);
        const critical   = tasks.filter(t => t.critical).length;
        const milestones = tasks.filter(t => t.milestone && t.percentComplete < 100)
            .sort((a,b) => new Date(a.finish) - new Date(b.finish)).slice(0, 4);

        const progress = total > 0 ? Math.round(tasks.reduce((s,t) => s+(t.percentComplete||0),0)/total) : 0;
        const totalCost = tasks.reduce((s,t) => s+(t.cost||0), 0);
        const earned    = tasks.reduce((s,t) => s+((t.cost||0)*(t.percentComplete||0)/100), 0);
        const remaining = totalCost - earned;

        // SPI / CPI
        const ps  = project.startDate  ? new Date(project.startDate)  : null;
        const pf  = project.finishDate ? new Date(project.finishDate) : null;
        const dur = ps && pf ? Math.max(1, (pf-ps)/864e5) : 0;
        const elapsed = ps ? Math.max(0,(today-ps)/864e5) : 0;
        const daysLeft = pf ? Math.max(0, Math.round((pf-today)/864e5)) : null;
        const elapsedPct = dur > 0 ? Math.min(100, Math.round(elapsed/dur*100)) : 0;
        const spi = elapsedPct > 0 && totalCost > 0 ? +(progress/elapsedPct).toFixed(2) : null;
        const cpi = earned > 0 && totalCost > 0 ? +((earned/totalCost)*(100/Math.max(progress,1))).toFixed(2) : null;

        // Overall health
        let health, healthIcon;
        if (lateTasks.length === 0 && (spi === null || spi >= 0.95)) { health = 'ON TRACK';  healthIcon = '🟢'; }
        else if (lateTasks.length <= 2 || (spi && spi >= 0.80))       { health = 'AT RISK';   healthIcon = '🟡'; }
        else                                                            { health = 'OFF TRACK'; healthIcon = '🔴'; }

        // Phase breakdown (summary tasks → outline level 1)
        const phases = project.tasks.filter(t => t.summary && (t.outlineLevel||1) === 1).map(p => {
            const pidx = project.tasks.indexOf(p);
            const nextSum = project.tasks.findIndex((t,i) => i > pidx && t.summary && (t.outlineLevel||1) <= (p.outlineLevel||1));
            const end = nextSum < 0 ? project.tasks.length : nextSum;
            const pts = project.tasks.slice(pidx+1, end).filter(t => !t.summary);
            const avg = pts.length ? Math.round(pts.reduce((s,t)=>s+(t.percentComplete||0),0)/pts.length) : 0;
            const isLate = pts.some(t => new Date(t.finish)<today && t.percentComplete<100);
            return { name: p.name, pct: avg, count: pts.length, isLate };
        });

        const fmt = n => n.toLocaleString('en-US');
        const fmtC = n => cur + fmt(Math.round(n));
        const fmtDate = d => { if (!d) return '—'; const dt=new Date(d); return dt.toLocaleDateString('en-US',{month:'short',day:'numeric',year:'numeric'}); };
        const fmtDateLong = d => { if(!d) return '—'; return new Date(d).toLocaleDateString('en-US',{weekday:'long',year:'numeric',month:'long',day:'numeric'}); };
        const diffDays = d => { const dt=new Date(d); return Math.round((today-dt)/864e5); };

        return { tasks, total, cur, today, complete, inProg, notStarted, lateTasks, critical,
                 milestones, progress, totalCost, earned, remaining, spi, cpi, health, healthIcon,
                 phases, daysLeft, elapsedPct, ps, pf, fmt, fmtC, fmtDate, fmtDateLong, diffDays };
    }

    function generateSummary(project, settings, format) {
        if (!project) return '';
        format = format || 'text';
        const d = _buildSummaryData(project, settings);
        if (format === 'markdown') return _summaryMarkdown(d, project, settings);
        if (format === 'html')     return _summaryHTML(d, project, settings);
        if (format === 'teams')    return _summaryTeams(d, project, settings);
        return _summaryPlainText(d, project, settings);
    }

    // ══════════════════════════════════════════════════════════
    //  FORMAT RENDERERS  —  Corporate / No Emojis
    // ══════════════════════════════════════════════════════════

    /** Build an ASCII progress bar:  ■■■■■□□□□□  */
    function _bar(pct, width) {
        const w = width || 20;
        const f = Math.round(pct / 100 * w);
        return '■'.repeat(f) + '□'.repeat(w - f);
    }

    /** Map health/SPI/CPI to a word-only status label */
    function _label(val, threshHigh, threshLow) {
        if (val >= threshHigh) return 'ON TRACK';
        if (val >= threshLow)  return 'AT RISK';
        return 'OFF TRACK';
    }

    // ──────────────────────────────────────────────────────────
    //  PLAIN TEXT  — memo / email body style
    // ──────────────────────────────────────────────────────────
    function _summaryPlainText(d, project) {
        const W  = 68;
        const HR = '─'.repeat(W);
        const HR2= '═'.repeat(W);
        const pad = (s, n) => String(s).slice(0, n).padEnd(n);
        const rpad = (s, n) => String(s).slice(0, n).padStart(n);
        const L = [];

        // ── Header ──
        L.push(HR2);
        L.push(pad('  PROJECT STATUS REPORT', W));
        L.push(pad('  Generated by ProjectFlow™  |  Confidential', W));
        L.push(HR2);
        L.push('');
        L.push(`  Project  :  ${project.name}`);
        L.push(`  Date     :  ${d.fmtDateLong(d.today)}`);
        L.push(`  Manager  :  ${project.manager || 'N/A'}`);
        if (d.ps && d.pf) {
            L.push(`  Period   :  ${d.fmtDate(d.ps)}  to  ${d.fmtDate(d.pf)}  (${d.daysLeft ?? '?'} days remaining)`);
        }
        L.push(`  Status   :  ${d.health}`);
        L.push('');

        // ── Executive Summary ──
        L.push(HR);
        L.push('  1.  EXECUTIVE SUMMARY');
        L.push(HR);
        L.push('');
        L.push(`  Overall progress stands at ${d.progress}% with ${d.complete} of ${d.total} tasks`);
        L.push(`  completed as of the reporting date. The project is currently assessed`);
        L.push(`  as ${d.health}.`);
        if (d.lateTasks.length > 0) {
            L.push(`  There are ${d.lateTasks.length} overdue task(s) requiring immediate attention.`);
        } else {
            L.push(`  There are no overdue tasks at this time.`);
        }
        L.push('');

        // ── Key Metrics ──
        L.push(HR);
        L.push('  2.  KEY PERFORMANCE METRICS');
        L.push(HR);
        L.push('');
        L.push(`  ${'METRIC'.padEnd(24)} ${'VALUE'.padEnd(14)} INDICATOR`);
        L.push(`  ${'─'.repeat(22)} ${'─'.repeat(12)} ${'─'.repeat(22)}`);
        L.push(`  ${'Overall Progress'.padEnd(24)} ${(d.progress + '%').padEnd(14)} ${_bar(d.progress, 18)}`);
        L.push(`  ${'Tasks — Complete'.padEnd(24)} ${String(d.complete).padEnd(14)} ${Math.round(d.complete/Math.max(d.total,1)*100)}% of total`);
        L.push(`  ${'Tasks — In Progress'.padEnd(24)} ${String(d.inProg).padEnd(14)}`);
        L.push(`  ${'Tasks — Not Started'.padEnd(24)} ${String(d.notStarted).padEnd(14)}`);
        L.push(`  ${'Tasks — Overdue'.padEnd(24)} ${String(d.lateTasks.length).padEnd(14)} ${d.lateTasks.length > 0 ? 'ACTION REQUIRED' : 'None'}`);
        L.push(`  ${'Critical Path Tasks'.padEnd(24)} ${String(d.critical).padEnd(14)}`);
        if (d.spi !== null) {
            L.push(`  ${'Schedule Perf. (SPI)'.padEnd(24)} ${d.spi.toFixed(2).padEnd(14)} ${_label(d.spi, 1.0, 0.85)}`);
        }
        if (d.cpi) {
            L.push(`  ${'Cost Perf. (CPI)'.padEnd(24)} ${d.cpi.toFixed(2).padEnd(14)} ${_label(d.cpi, 1.0, 0.85)}`);
        }
        L.push('');
        if (d.ps && d.pf) {
            L.push(`  Timeline`);
            L.push(`  ${d.fmtDate(d.ps)}  ${_bar(d.elapsedPct, 36)}  ${d.fmtDate(d.pf)}`);
            L.push(`  ${String(d.elapsedPct + '% elapsed').padEnd(42)} ${d.daysLeft ?? '?'} days remaining`);
            L.push('');
        }

        // ── Budget ──
        if (d.totalCost > 0) {
            L.push(HR);
            L.push('  3.  BUDGET & COST');
            L.push(HR);
            L.push('');
            L.push(`  ${'Total Budget'.padEnd(24)} ${d.fmtC(d.totalCost)}`);
            L.push(`  ${'Earned Value (BCWP)'.padEnd(24)} ${d.fmtC(d.earned)}`);
            L.push(`  ${'Remaining'.padEnd(24)} ${d.fmtC(d.remaining)}`);
            const pctSpent = d.totalCost > 0 ? Math.round(d.earned / d.totalCost * 100) : 0;
            L.push(`  ${'Budget Consumed'.padEnd(24)} ${pctSpent}%`);
            L.push('');
        }

        // ── Phase Breakdown ──
        if (d.phases.length > 0) {
            const sectionNum = d.totalCost > 0 ? 4 : 3;
            L.push(HR);
            L.push(`  ${sectionNum}.  PHASE / WORKSTREAM STATUS`);
            L.push(HR);
            L.push('');
            L.push(`  ${'PHASE / WORKSTREAM'.padEnd(32)} ${'PROGRESS'.padEnd(8)} ${'BAR'.padEnd(18)} STATUS`);
            L.push(`  ${'─'.repeat(30)} ${'─'.repeat(6)} ${'─'.repeat(16)} ${'─'.repeat(12)}`);
            d.phases.forEach(p => {
                const status = p.pct >= 100 ? 'COMPLETE' : p.isLate ? 'BEHIND' : p.pct > 0 ? 'IN PROGRESS' : 'NOT STARTED';
                L.push(`  ${pad(p.name, 32)} ${(p.pct + '%').padEnd(8)} ${_bar(p.pct, 16)}  ${status}`);
            });
            L.push('');
        }

        // ── Risks ──
        if (d.lateTasks.length > 0) {
            const sn = d.phases.length > 0 ? (d.totalCost > 0 ? 5 : 4) : (d.totalCost > 0 ? 4 : 3);
            L.push(HR);
            L.push(`  ${sn}.  RISKS & ISSUES  (${d.lateTasks.length} overdue task(s))`);
            L.push(HR);
            L.push('');
            L.push(`  ${'TASK'.padEnd(34)} ${'OVERDUE'.padEnd(12)} OWNER`);
            L.push(`  ${'─'.repeat(32)} ${'─'.repeat(10)} ${'─'.repeat(18)}`);
            d.lateTasks.slice(0, 6).forEach(t => {
                const days = d.diffDays(t.finish);
                const res  = (t.resourceNames || []).join(', ').slice(0, 18) || 'Unassigned';
                L.push(`  ${pad(t.name, 34)} ${(days + 'd overdue').padEnd(12)} ${res}`);
            });
            if (d.lateTasks.length > 6) L.push(`  ... and ${d.lateTasks.length - 6} additional overdue task(s)`);
            L.push('');
        }

        // ── Milestones ──
        if (d.milestones.length > 0) {
            const lastSn = d.lateTasks.length > 0
                ? (d.phases.length > 0 ? (d.totalCost > 0 ? 6 : 5) : (d.totalCost > 0 ? 5 : 4))
                : (d.phases.length > 0 ? (d.totalCost > 0 ? 5 : 4) : (d.totalCost > 0 ? 4 : 3));
            L.push(HR);
            L.push(`  ${lastSn}.  UPCOMING MILESTONES`);
            L.push(HR);
            L.push('');
            L.push(`  ${'MILESTONE'.padEnd(34)} ${'TARGET DATE'.padEnd(16)} DAYS`);
            L.push(`  ${'─'.repeat(32)} ${'─'.repeat(14)} ${'─'.repeat(8)}`);
            d.milestones.forEach(m => {
                const daysTo = -d.diffDays(m.finish);
                L.push(`  ${pad(m.name, 34)} ${d.fmtDate(m.finish).padEnd(16)} In ${daysTo} day(s)`);
            });
            L.push('');
        }

        // ── Footer ──
        L.push(HR2);
        L.push(`  CONFIDENTIAL  |  ProjectFlow™  |  ${new Date().toLocaleString('en-US')}`);
        L.push(HR2);

        return L.join('\n');
    }

    // ──────────────────────────────────────────────────────────
    //  MARKDOWN  — Confluence / Notion / GitHub
    // ──────────────────────────────────────────────────────────
    function _summaryMarkdown(d, project) {
        const L = [];
        L.push(`# Project Status Report`);
        L.push(`**${project.name}**`);
        L.push('');
        L.push(`| Field | Value |`);
        L.push(`|:---|:---|`);
        L.push(`| Report Date | ${d.fmtDateLong(d.today)} |`);
        L.push(`| Project Manager | ${project.manager || 'N/A'} |`);
        if (d.ps && d.pf) L.push(`| Project Period | ${d.fmtDate(d.ps)} — ${d.fmtDate(d.pf)} |`);
        if (d.daysLeft !== null) L.push(`| Days Remaining | ${d.daysLeft} |`);
        L.push(`| Overall Status | **${d.health}** |`);
        L.push('');
        L.push('---');
        L.push('');
        L.push('## 1. Key Performance Metrics');
        L.push('');
        L.push('| Metric | Value | Assessment |');
        L.push('|:---|:---:|:---|');
        L.push(`| Overall Progress | **${d.progress}%** | ${_bar(d.progress, 14)} |`);
        L.push(`| Tasks Complete | ${d.complete} / ${d.total} | ${Math.round(d.complete/Math.max(d.total,1)*100)}% |`);
        L.push(`| Tasks In Progress | ${d.inProg} | — |`);
        L.push(`| Tasks Overdue | ${d.lateTasks.length} | ${d.lateTasks.length > 0 ? 'Action Required' : 'None'} |`);
        L.push(`| Critical Path Tasks | ${d.critical} | — |`);
        if (d.spi !== null) L.push(`| Schedule Performance (SPI) | ${d.spi.toFixed(2)} | ${_label(d.spi, 1.0, 0.85)} |`);
        if (d.cpi)          L.push(`| Cost Performance (CPI) | ${d.cpi.toFixed(2)} | ${_label(d.cpi, 1.0, 0.85)} |`);
        if (d.totalCost > 0) {
            L.push(`| Total Budget | ${d.fmtC(d.totalCost)} | — |`);
            L.push(`| Earned Value | ${d.fmtC(d.earned)} | ${d.progress}% delivered |`);
            L.push(`| Remaining | ${d.fmtC(d.remaining)} | — |`);
        }
        L.push('');

        if (d.phases.length > 0) {
            L.push('---');
            L.push('');
            L.push('## 2. Phase / Workstream Status');
            L.push('');
            L.push('| Phase | Progress | Status |');
            L.push('|:---|:---:|:---|');
            d.phases.forEach(p => {
                const s = p.pct >= 100 ? 'Complete' : p.isLate ? 'Behind Schedule' : p.pct > 0 ? 'In Progress' : 'Not Started';
                L.push(`| ${p.name} | ${p.pct}% | ${s} |`);
            });
            L.push('');
        }

        if (d.lateTasks.length > 0) {
            L.push('---');
            L.push('');
            L.push('## 3. Risks & Issues');
            L.push('');
            L.push('| Task | Days Overdue | Owner |');
            L.push('|:---|:---:|:---|');
            d.lateTasks.slice(0, 8).forEach(t => {
                const days = d.diffDays(t.finish);
                const res  = (t.resourceNames || []).join(', ') || 'Unassigned';
                L.push(`| ${t.name} | ${days} | ${res} |`);
            });
            L.push('');
        }

        if (d.milestones.length > 0) {
            L.push('---');
            L.push('');
            L.push('## 4. Upcoming Milestones');
            L.push('');
            L.push('| Milestone | Target Date | Days Remaining |');
            L.push('|:---|:---|:---:|');
            d.milestones.forEach(m => {
                const daysTo = -d.diffDays(m.finish);
                L.push(`| ${m.name} | ${d.fmtDate(m.finish)} | ${daysTo} |`);
            });
            L.push('');
        }

        L.push('---');
        L.push('');
        L.push(`*Confidential — Generated by ProjectFlow™ — ${new Date().toLocaleString('en-US')}*`);
        return L.join('\n');
    }

    // ──────────────────────────────────────────────────────────
    //  MS TEAMS  — plain adaptive card text
    // ──────────────────────────────────────────────────────────
    function _summaryTeams(d, project) {
        const L = [];
        L.push(`**PROJECT STATUS UPDATE  |  ${project.name}**`);
        L.push(`${d.fmtDateLong(d.today)}  |  Manager: ${project.manager || 'N/A'}`);
        L.push('');
        L.push(`**Overall Status: ${d.health}**`);
        L.push('');
        L.push('**METRICS**');
        L.push(`Progress:      ${d.progress}%   ${_bar(d.progress, 16)}`);
        L.push(`Complete:      ${d.complete} tasks (${Math.round(d.complete/Math.max(d.total,1)*100)}%)`);
        L.push(`Overdue:       ${d.lateTasks.length} task(s)${d.lateTasks.length > 0 ? ' — ACTION REQUIRED' : ''}`);
        L.push(`Critical:      ${d.critical} tasks`);
        if (d.spi !== null) L.push(`SPI:           ${d.spi.toFixed(2)}  (${_label(d.spi, 1.0, 0.85)})`);
        if (d.totalCost > 0) {
            L.push(`Budget:        ${d.fmtC(d.totalCost)}  |  Earned: ${d.fmtC(d.earned)}  |  Remaining: ${d.fmtC(d.remaining)}`);
        }
        if (d.daysLeft !== null) L.push(`Days Remaining: ${d.daysLeft}`);
        L.push('');

        if (d.lateTasks.length > 0) {
            L.push('**OVERDUE TASKS — ACTION REQUIRED**');
            d.lateTasks.slice(0, 5).forEach(t => {
                const days = d.diffDays(t.finish);
                const res  = (t.resourceNames || []).slice(0, 1).join('') || 'Unassigned';
                L.push(`  • ${t.name}  (${days}d overdue — ${res})`);
            });
            L.push('');
        }

        if (d.milestones.length > 0) {
            L.push('**UPCOMING MILESTONES**');
            d.milestones.slice(0, 3).forEach(m => {
                const daysTo = -d.diffDays(m.finish);
                L.push(`  • ${m.name}  —  ${d.fmtDate(m.finish)}  (${daysTo} days)`);
            });
            L.push('');
        }

        L.push('---');
        L.push(`Confidential  |  ProjectFlow™  |  ${new Date().toLocaleString('en-US')}`);
        return L.join('\n');
    }

    // ──────────────────────────────────────────────────────────
    //  RICH HTML  — Outlook / Gmail professional email
    // ──────────────────────────────────────────────────────────
    function _summaryHTML(d, project) {
        const statusColor  = d.health === 'ON TRACK'  ? '#166534' : d.health === 'AT RISK' ? '#92400e' : '#991b1b';
        const statusBg     = d.health === 'ON TRACK'  ? '#dcfce7' : d.health === 'AT RISK' ? '#fef3c7' : '#fee2e2';
        const statusBorder = d.health === 'ON TRACK'  ? '#16a34a' : d.health === 'AT RISK' ? '#d97706' : '#dc2626';
        const progColor    = d.progress >= 75 ? '#1d4ed8' : d.progress >= 40 ? '#1d4ed8' : '#b91c1c';

        const kpiCard = (label, value, note, borderColor) =>
            `<td style="width:25%;padding:0 6px 0 0">
               <div style="border:1px solid #e2e8f0;border-top:3px solid ${borderColor||'#64748b'};border-radius:4px;padding:14px 12px;text-align:center">
                 <div style="font-size:24px;font-weight:700;color:#0f172a;line-height:1">${value}</div>
                 <div style="font-size:10px;font-weight:600;color:#64748b;text-transform:uppercase;letter-spacing:.06em;margin-top:5px">${label}</div>
                 ${note ? `<div style="font-size:10px;color:#94a3b8;margin-top:3px">${note}</div>` : ''}
               </div>
             </td>`;

        const phaseRows = d.phases.map(p => {
            const bar_filled = Math.round(p.pct / 100 * 120);
            const statusText = p.pct >= 100 ? 'Complete' : p.isLate ? 'Behind Schedule' : p.pct > 0 ? 'In Progress' : 'Not Started';
            const statusC    = p.pct >= 100 ? '#166534' : p.isLate ? '#991b1b' : p.pct > 0 ? '#1e40af' : '#64748b';
            return `<tr style="border-bottom:1px solid #f1f5f9">
              <td style="padding:8px 12px;font-size:12px;color:#1e293b;font-weight:500">${p.name}</td>
              <td style="padding:8px 12px">
                <div style="background:#f1f5f9;border-radius:2px;height:8px;width:120px">
                  <div style="background:#1d4ed8;border-radius:2px;height:8px;width:${bar_filled}px"></div>
                </div>
              </td>
              <td style="padding:8px 12px;font-size:12px;font-weight:700;color:#0f172a;text-align:right">${p.pct}%</td>
              <td style="padding:8px 12px;font-size:11px;color:${statusC};font-weight:600;text-align:right">${statusText}</td>
            </tr>`;
        }).join('');

        const riskRows = d.lateTasks.slice(0, 6).map((t, i) => {
            const days = d.diffDays(t.finish);
            const res  = (t.resourceNames || []).join(', ') || 'Unassigned';
            const bg   = i % 2 === 0 ? '#fff5f5' : '#ffffff';
            return `<tr style="background:${bg}">
              <td style="padding:8px 12px;font-size:12px;color:#1e293b">${t.name}</td>
              <td style="padding:8px 12px;font-size:12px;color:#dc2626;font-weight:700;white-space:nowrap">${days} day(s)</td>
              <td style="padding:8px 12px;font-size:12px;color:#64748b">${res}</td>
            </tr>`;
        }).join('');

        const msRows = d.milestones.map((m, i) => {
            const daysTo = -d.diffDays(m.finish);
            const bg     = i % 2 === 0 ? '#f8fafc' : '#ffffff';
            return `<tr style="background:${bg}">
              <td style="padding:8px 12px;font-size:12px;color:#1e293b;font-weight:500">${m.name}</td>
              <td style="padding:8px 12px;font-size:12px;color:#1e293b;white-space:nowrap">${d.fmtDate(m.finish)}</td>
              <td style="padding:8px 12px;font-size:12px;color:#64748b;white-space:nowrap">${daysTo} day(s)</td>
            </tr>`;
        }).join('');

        const progFilled = Math.round(d.progress * 5.4); // 0–540px

        return `<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:24px;background:#f8fafc;font-family:'Segoe UI',Arial,sans-serif;color:#1e293b">
<table width="640" cellpadding="0" cellspacing="0" align="center"
       style="background:#ffffff;border:1px solid #e2e8f0;border-radius:4px">

  <\!-- LETTERHEAD -->
  <tr>
    <td style="padding:28px 32px 20px;border-bottom:3px solid #1e3a8a">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td>
            <div style="font-size:9px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:#64748b;margin-bottom:6px">Project Status Report — Confidential</div>
            <div style="font-size:20px;font-weight:700;color:#0f172a;line-height:1.2">${project.name}</div>
          </td>
          <td style="text-align:right;vertical-align:top">
            <div style="display:inline-block;background:${statusBg};border:1px solid ${statusBorder};border-radius:3px;padding:5px 14px">
              <div style="font-size:10px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:${statusColor}">${d.health}</div>
            </div>
          </td>
        </tr>
      </table>
      <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:12px">
        <tr>
          <td style="font-size:11px;color:#64748b">Report Date: <strong style="color:#1e293b">${d.fmtDateLong(d.today)}</strong></td>
          <td style="font-size:11px;color:#64748b;text-align:center">Manager: <strong style="color:#1e293b">${project.manager || 'N/A'}</strong></td>
          <td style="font-size:11px;color:#64748b;text-align:right">${d.ps && d.pf ? `Period: <strong style="color:#1e293b">${d.fmtDate(d.ps)} – ${d.fmtDate(d.pf)}</strong>` : ''}</td>
        </tr>
      </table>
    </td>
  </tr>

  <\!-- KPI CARDS -->
  <tr>
    <td style="padding:20px 32px">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          ${kpiCard('Progress', d.progress + '%', d.daysLeft !== null ? d.daysLeft + ' days remaining' : null, '#1d4ed8')}
          ${kpiCard('Complete', d.complete + ' / ' + d.total, Math.round(d.complete/Math.max(d.total,1)*100) + '% of tasks', '#16a34a')}
          ${kpiCard('Overdue', String(d.lateTasks.length), d.lateTasks.length > 0 ? 'Action required' : 'No overdue tasks', d.lateTasks.length > 0 ? '#dc2626' : '#16a34a')}
          ${kpiCard('Critical', String(d.critical), 'critical path tasks', '#dc2626')}
        </tr>
      </table>
    </td>
  </tr>

  <\!-- PROGRESS BAR -->
  <tr>
    <td style="padding:0 32px 20px">
      <div style="font-size:10px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:#64748b;margin-bottom:8px">Overall Progress</div>
      <div style="background:#f1f5f9;border-radius:3px;height:14px;overflow:hidden">
        <div style="background:#1d4ed8;height:14px;width:${d.progress}%;border-radius:3px"></div>
      </div>
      <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:6px">
        <tr>
          <td style="font-size:10px;color:#94a3b8">${d.fmtDate(d.ps) || ''}</td>
          <td style="font-size:11px;font-weight:700;color:#1d4ed8;text-align:center">${d.progress}%  —  ${d.complete} of ${d.total} tasks complete</td>
          <td style="font-size:10px;color:#94a3b8;text-align:right">${d.fmtDate(d.pf) || ''}</td>
        </tr>
      </table>
    </td>
  </tr>

  <\!-- SPI / CPI / BUDGET -->
  ${(d.spi !== null || d.totalCost > 0) ? `<tr>
    <td style="padding:0 32px 20px">
      <div style="font-size:10px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:#64748b;margin-bottom:8px">Performance Indices</div>
      <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e2e8f0;border-radius:4px">
        <tr style="background:#f8fafc">
          ${d.spi !== null ? `<td style="padding:10px 16px;font-size:12px;border-right:1px solid #e2e8f0"><span style="color:#64748b">SPI (Schedule)</span><br><strong style="color:${d.spi>=1?'#166534':d.spi>=.85?'#92400e':'#991b1b'}">${d.spi.toFixed(2)}</strong>  <span style="font-size:10px;color:#94a3b8">${_label(d.spi,1.0,.85)}</span></td>` : ''}
          ${d.cpi ? `<td style="padding:10px 16px;font-size:12px;border-right:1px solid #e2e8f0"><span style="color:#64748b">CPI (Cost)</span><br><strong style="color:${d.cpi>=1?'#166534':d.cpi>=.85?'#92400e':'#991b1b'}">${d.cpi.toFixed(2)}</strong>  <span style="font-size:10px;color:#94a3b8">${_label(d.cpi,1.0,.85)}</span></td>` : ''}
          ${d.totalCost > 0 ? `<td style="padding:10px 16px;font-size:12px;border-right:1px solid #e2e8f0"><span style="color:#64748b">Total Budget</span><br><strong>${d.fmtC(d.totalCost)}</strong></td>
          <td style="padding:10px 16px;font-size:12px;border-right:1px solid #e2e8f0"><span style="color:#64748b">Earned Value</span><br><strong style="color:#166534">${d.fmtC(d.earned)}</strong></td>
          <td style="padding:10px 16px;font-size:12px"><span style="color:#64748b">Remaining</span><br><strong>${d.fmtC(d.remaining)}</strong></td>` : ''}
        </tr>
      </table>
    </td>
  </tr>` : ''}

  <\!-- PHASE STATUS -->
  ${d.phases.length > 0 ? `<tr>
    <td style="padding:0 32px 20px">
      <div style="font-size:10px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:#64748b;margin-bottom:8px">Phase / Workstream Status</div>
      <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e2e8f0;border-radius:4px;overflow:hidden">
        <tr style="background:#f8fafc;border-bottom:1px solid #e2e8f0">
          <th style="padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#64748b;text-align:left">Phase</th>
          <th style="padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#64748b;text-align:left">Progress</th>
          <th style="padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#64748b;text-align:right">%</th>
          <th style="padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#64748b;text-align:right">Status</th>
        </tr>
        ${phaseRows}
      </table>
    </td>
  </tr>` : ''}

  <\!-- RISKS -->
  ${d.lateTasks.length > 0 ? `<tr>
    <td style="padding:0 32px 20px">
      <div style="font-size:10px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:#dc2626;margin-bottom:8px">Risks &amp; Issues — Action Required</div>
      <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #fecaca;border-radius:4px;overflow:hidden">
        <tr style="background:#fef2f2;border-bottom:1px solid #fecaca">
          <th style="padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#dc2626;text-align:left">Task</th>
          <th style="padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#dc2626;text-align:center">Days Overdue</th>
          <th style="padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#dc2626;text-align:left">Owner</th>
        </tr>
        ${riskRows}
      </table>
    </td>
  </tr>` : ''}

  <\!-- MILESTONES -->
  ${d.milestones.length > 0 ? `<tr>
    <td style="padding:0 32px 20px">
      <div style="font-size:10px;font-weight:600;letter-spacing:.06em;text-transform:uppercase;color:#64748b;margin-bottom:8px">Upcoming Milestones</div>
      <table width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e2e8f0;border-radius:4px;overflow:hidden">
        <tr style="background:#f8fafc;border-bottom:1px solid #e2e8f0">
          <th style="padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#64748b;text-align:left">Milestone</th>
          <th style="padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#64748b;text-align:left">Target Date</th>
          <th style="padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.05em;color:#64748b;text-align:right">Days</th>
        </tr>
        ${msRows}
      </table>
    </td>
  </tr>` : ''}

  <\!-- FOOTER -->
  <tr>
    <td style="padding:16px 32px;background:#f8fafc;border-top:1px solid #e2e8f0">
      <table width="100%" cellpadding="0" cellspacing="0">
        <tr>
          <td style="font-size:10px;color:#94a3b8">Confidential — For internal distribution only</td>
          <td style="font-size:10px;color:#94a3b8;text-align:right">ProjectFlow™ — ${new Date().toLocaleString('en-US')}</td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body></html>`;
    }


    async function copySummary(project, settings, format) {
        const content = generateSummary(project, settings, format);

        // HTML format: copy as rich HTML + plain text fallback
        if (format === 'html') {
            try {
                const htmlBlob  = new Blob([content], { type: 'text/html' });
                const plainBlob = new Blob([_buildSummaryData(project, settings) ? generateSummary(project, settings, 'text') : content], { type: 'text/plain' });
                await navigator.clipboard.write([
                    new ClipboardItem({ 'text/html': htmlBlob, 'text/plain': plainBlob })
                ]);
                return true;
            } catch(e) {
                // Fallback: copy HTML as text
                return _copyText(content);
            }
        }

        return _copyText(content);
    }

    function _copyText(text) {
        try {
            navigator.clipboard.writeText(text);
            return true;
        } catch(e) {
            const ta = document.createElement('textarea');
            ta.value = text; document.body.appendChild(ta);
            ta.select(); document.execCommand('copy');
            document.body.removeChild(ta);
            return true;
        }
    }

    // ═══════════════════════════════
    // CSV / EXCEL EXPORT
    // ═══════════════════════════════
    function exportCSV(project, settings) {
        if (!project) return;
        const rows = [['ID', 'WBS', 'Task Name', 'Duration', 'Start', 'Finish', '% Complete', 'Float', 'Predecessors', 'Resources', 'Cost', 'Critical', 'Outline Level']];
        const cur = settings.currency || '$';

        project.tasks.forEach(t => {
            rows.push([
                t.id, t.wbs, t.name, t.duration + 'd',
                fmtDate(t.start), fmtDate(t.finish),
                t.percentComplete + '%', (t.totalFloat || 0) + 'd',
                t.predecessors || '', (t.resourceNames || ''),
                cur + (t.cost || 0), t.critical ? 'Yes' : 'No',
                t.outlineLevel || 1
            ]);
        });

        const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g, '""')}"`).join(',')).join('\n');
        return csv;
    }

    function exportExcel(project, settings) {
        if (!project || typeof XLSX === 'undefined') return null;
        const wb = XLSX.utils.book_new();
        const cur = settings.currency || '$';

        // Sheet 1: Tasks
        const taskData = [['ID', 'WBS', 'Task Name', 'Duration', 'Start', 'Finish', '% Complete', 'Float', 'Predecessors', 'Resources', 'Cost', 'Critical']];
        project.tasks.forEach(t => {
            taskData.push([t.id, t.wbs, t.name, t.duration + 'd', fmtDate(t.start), fmtDate(t.finish),
                t.percentComplete + '%', (t.totalFloat || 0) + 'd', t.predecessors || '', t.resourceNames || '',
                (t.cost || 0), t.critical ? 'Yes' : 'No']);
        });
        const ws1 = XLSX.utils.aoa_to_sheet(taskData);
        ws1['!cols'] = [{wch:4},{wch:6},{wch:30},{wch:8},{wch:12},{wch:12},{wch:8},{wch:6},{wch:12},{wch:15},{wch:10},{wch:8}];
        XLSX.utils.book_append_sheet(wb, ws1, 'Tasks');

        // Sheet 2: Resources
        if (project.resources && project.resources.length > 0) {
            const resData = [['ID', 'Name', 'Type', 'Max Units', 'Cost/Hour']];
            project.resources.forEach(r => {
                resData.push([r.id, r.name, r.type || 'Work', (r.maxUnits || 100) + '%', cur + (r.costPerHour || 0)]);
            });
            const ws2 = XLSX.utils.aoa_to_sheet(resData);
            ws2['!cols'] = [{wch:4},{wch:20},{wch:8},{wch:10},{wch:10}];
            XLSX.utils.book_append_sheet(wb, ws2, 'Resources');
        }

        // Sheet 3: Summary
        const tasks = project.tasks.filter(t => !t.summary);
        const total = tasks.length;
        const complete = tasks.filter(t => t.percentComplete >= 100).length;
        const critical = tasks.filter(t => t.critical).length;
        const totalCost = tasks.reduce((s, t) => s + (t.cost || 0), 0);
        const progress = total > 0 ? Math.round(tasks.reduce((s, t) => s + (t.percentComplete || 0), 0) / total) : 0;

        const summaryData = [
            ['Project Summary'],
            [''], ['Name', project.name], ['Manager', project.manager || '—'],
            ['Start Date', fmtDate(project.startDate)], ['Finish Date', fmtDate(project.finishDate)],
            [''], ['Metric', 'Value'],
            ['Total Tasks', total], ['Complete', complete], ['Progress', progress + '%'],
            ['Critical Tasks', critical], ['Total Cost', cur + totalCost.toLocaleString()],
        ];
        const ws3 = XLSX.utils.aoa_to_sheet(summaryData);
        ws3['!cols'] = [{wch:15},{wch:25}];
        XLSX.utils.book_append_sheet(wb, ws3, 'Summary');

        // Sheet 4: Critical Path
        const critTasks = project.tasks.filter(t => t.critical && !t.summary);
        if (critTasks.length > 0) {
            const critData = [['ID', 'Task Name', 'Duration', 'Start', 'Finish', 'Float', 'Resources']];
            critTasks.forEach(t => {
                critData.push([t.id, t.name, t.duration + 'd', fmtDate(t.start), fmtDate(t.finish), (t.totalFloat || 0) + 'd', t.resourceNames || '']);
            });
            const ws4 = XLSX.utils.aoa_to_sheet(critData);
            ws4['!cols'] = [{wch:4},{wch:30},{wch:8},{wch:12},{wch:12},{wch:6},{wch:15}];
            XLSX.utils.book_append_sheet(wb, ws4, 'Critical Path');
        }

        return wb;
    }

    function downloadExcel(wb, filename) {
        if (!wb || typeof XLSX === 'undefined') return;
        XLSX.writeFile(wb, filename);
    }

    // ═══════════════════════════════
    // GANTT IMAGE EXPORT
    // ═══════════════════════════════
    function exportGanttPNG(ganttCanvas) {
        if (!ganttCanvas) return;
        const link = document.createElement('a');
        link.download = 'gantt_chart.png';
        link.href = ganttCanvas.toDataURL('image/png');
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    // ═══════════════════════════════
    // PDF REPORT
    // ═══════════════════════════════
    async function generatePDF(project, settings, ganttCanvas) {
        if (!project) return;
        if (typeof jspdf === 'undefined' && typeof window.jspdf === 'undefined') {
            throw new Error('jsPDF not loaded');
        }

        var jsPDF = window.jspdf.jsPDF;
        var pdf = new jsPDF('landscape', 'mm', 'a4');
        var W = pdf.internal.pageSize.getWidth();
        var H = pdf.internal.pageSize.getHeight();
        var margin = 14;

        var tasks = project.tasks.filter(function(t) { return !t.summary; });
        var total = tasks.length;
        var complete = tasks.filter(function(t) { return t.percentComplete >= 100; }).length;
        var inProgress = tasks.filter(function(t) { return t.percentComplete > 0 && t.percentComplete < 100; }).length;
        var critical = tasks.filter(function(t) { return t.critical; }).length;
        var progress = total > 0 ? Math.round(tasks.reduce(function(s, t) { return s + (t.percentComplete || 0); }, 0) / total) : 0;
        var today = new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });

        // ── PAGE 1: Cover ──
        // White background
        pdf.setFillColor(255, 255, 255);
        pdf.rect(0, 0, W, H, 'F');

        // Top accent bar
        pdf.setFillColor(79, 70, 229); // indigo-600
        pdf.rect(0, 0, W, 6, 'F');

        // Company logo (with aspect ratio preservation)
        var logoSrc = localStorage.getItem('pf_report_logo') || (typeof PROART_LOGO !== 'undefined' ? PROART_LOGO : null);
        if (logoSrc) {
            try {
                await new Promise(function(resolve) {
                    var im = new Image();
                    im.onload = function() {
                        var lw = 60, lh = 18;
                        if (im.naturalWidth && im.naturalHeight) {
                            var ra = im.naturalWidth / im.naturalHeight;
                            if (ra > 1) { lh = lw / ra; } else { lw = lh * ra; }
                        }
                        pdf.addImage(logoSrc, 'PNG', (W - lw) / 2, 14, lw, lh);
                        resolve();
                    };
                    im.onerror = resolve;
                    im.src = logoSrc;
                });
            } catch (e) { /* logo failed */ }
        }

        // Project name
        pdf.setTextColor(30, 27, 75); // indigo-950
        pdf.setFontSize(28);
        pdf.text(project.name || 'Project Report', W / 2, 55, { align: 'center' });

        // Subtitle
        pdf.setFontSize(13);
        pdf.setTextColor(100, 116, 139); // slate-500
        pdf.text('Project Status Report', W / 2, 67, { align: 'center' });

        // Meta info
        pdf.setFontSize(10);
        pdf.setTextColor(71, 85, 105); // slate-600
        pdf.text('Project Manager: ' + (project.manager || '--'), W / 2, 85, { align: 'center' });
        pdf.text('Report Date: ' + today, W / 2, 92, { align: 'center' });
        pdf.text('Period: ' + fmtDate(project.startDate) + ' -- ' + fmtDate(project.finishDate), W / 2, 99, { align: 'center' });

        // KPI Cards
        var kpiY = 115;
        var kpiW = 55;
        var kpiH = 30;
        var kpis = [
            { label: 'OVERALL PROGRESS', value: progress + '%', color: [79, 70, 229] },
            { label: 'TOTAL TASKS', value: String(total), color: [59, 130, 246] },
            { label: 'COMPLETED', value: String(complete), color: [34, 197, 94] },
            { label: 'IN PROGRESS', value: String(inProgress), color: [245, 158, 11] },
            { label: 'CRITICAL', value: String(critical), color: critical > 0 ? [239, 68, 68] : [100, 116, 139] },
        ];
        var totalKpiW = kpis.length * kpiW + (kpis.length - 1) * 8;
        var kpiX = (W - totalKpiW) / 2;
        kpis.forEach(function(k) {
            // Card bg
            pdf.setFillColor(248, 250, 252); // slate-50
            pdf.roundedRect(kpiX, kpiY, kpiW, kpiH, 3, 3, 'F');
            // Left accent
            pdf.setFillColor(k.color[0], k.color[1], k.color[2]);
            pdf.rect(kpiX, kpiY + 3, 3, kpiH - 6, 'F');
            // Label
            pdf.setFontSize(6.5);
            pdf.setTextColor(100, 116, 139);
            pdf.text(k.label, kpiX + kpiW / 2, kpiY + 10, { align: 'center' });
            // Value
            pdf.setFontSize(16);
            pdf.setTextColor(k.color[0], k.color[1], k.color[2]);
            pdf.text(k.value, kpiX + kpiW / 2, kpiY + 23, { align: 'center' });
            kpiX += kpiW + 8;
        });

        // Progress bar visual
        var barY = kpiY + kpiH + 14;
        var barW = W - margin * 4;
        var barX = (W - barW) / 2;
        pdf.setFillColor(226, 232, 240); // slate-200
        pdf.roundedRect(barX, barY, barW, 5, 2.5, 2.5, 'F');
        pdf.setFillColor(79, 70, 229);
        var fillW = Math.max(2, barW * progress / 100);
        pdf.roundedRect(barX, barY, fillW, 5, 2.5, 2.5, 'F');
        pdf.setFontSize(8);
        pdf.setTextColor(71, 85, 105);
        pdf.text(progress + '% Complete', W / 2, barY + 12, { align: 'center' });

        // Footer
        pdf.setFontSize(7);
        pdf.setTextColor(148, 163, 184);
        pdf.text('Generated by ProjectFlow™', W / 2, H - 8, { align: 'center' });

        // ── PAGE 2+: Task Table ──
        pdf.addPage();
        addPageHeader(pdf, W, 'Task Schedule', project.name);

        var y = 26;
        var cols = [
            { label: 'ID', w: 10 },
            { label: 'WBS', w: 18 },
            { label: 'Task Name', w: 100 },
            { label: 'Duration', w: 18 },
            { label: 'Start', w: 28 },
            { label: 'Finish', w: 28 },
            { label: '% Done', w: 16 },
            { label: 'Critical', w: 14 },
        ];

        // Table header
        var x = margin;
        pdf.setFillColor(30, 27, 75); // indigo-950
        pdf.rect(margin, y, W - margin * 2, 7, 'F');
        pdf.setFontSize(7);
        pdf.setTextColor(255, 255, 255);
        cols.forEach(function(c) { pdf.text(c.label, x + 2, y + 5); x += c.w; });
        y += 8;

        // Data rows
        project.tasks.forEach(function(t, i) {
            if (y > H - 14) {
                pdf.addPage();
                addPageHeader(pdf, W, 'Task Schedule (cont.)', project.name);
                y = 26;
                // Repeat header
                x = margin;
                pdf.setFillColor(30, 27, 75);
                pdf.rect(margin, y, W - margin * 2, 7, 'F');
                pdf.setFontSize(7);
                pdf.setTextColor(255, 255, 255);
                cols.forEach(function(c) { pdf.text(c.label, x + 2, y + 5); x += c.w; });
                y += 8;
            }

            // Zebra stripe
            if (i % 2 === 0) { pdf.setFillColor(248, 250, 252); pdf.rect(margin, y - 1, W - margin * 2, 6, 'F'); }

            x = margin;
            var isCrit = t.critical && !t.summary;

            // Summary rows: bold indigo
            if (t.summary) {
                pdf.setFillColor(238, 242, 255); // indigo-50
                pdf.rect(margin, y - 1, W - margin * 2, 6, 'F');
                pdf.setTextColor(30, 27, 75);
            } else if (isCrit) {
                pdf.setTextColor(185, 28, 28); // red-700
            } else {
                pdf.setTextColor(30, 41, 59); // slate-800
            }

            pdf.setFontSize(7);
            var indent = '';
            for (var lv = 1; lv < (t.outlineLevel || 1); lv++) indent += '  ';
            var taskName = indent + (t.name || '');
            if (taskName.length > 55) taskName = taskName.substring(0, 52) + '...';

            var rowData = [
                String(t.id),
                t.wbs || t.outlineNumber || '',
                taskName,
                (t.durationDays || t.duration || 0) + 'd',
                fmtDate(t.start),
                fmtDate(t.finish),
                (t.percentComplete || 0) + '%',
                isCrit ? '\u25cf' : ''
            ];
            cols.forEach(function(c, ci) { pdf.text(String(rowData[ci] || ''), x + 2, y + 3); x += c.w; });
            y += 6;
        });

        // ══════════════════════════════════════════
        // PAGE: CRITICAL PATH
        // ══════════════════════════════════════════
        var critTasks = project.tasks.filter(function(t) { return t.critical && !t.summary; });
        if (critTasks.length > 0) {
            pdf.addPage();
            addPageHeader(pdf, W, 'Critical Path Analysis', project.name);
            y = 26;

            // Summary box
            pdf.setFillColor(254, 242, 242); // red-50
            pdf.roundedRect(margin, y, W - margin * 2, 14, 2, 2, 'F');
            pdf.setFontSize(9);
            pdf.setTextColor(153, 27, 27); // red-800
            pdf.text('[!] ' + critTasks.length + ' critical tasks on the longest path  |  Total Float: 0 days  |  Any delay impacts project finish date', margin + 6, y + 9);
            y += 20;

            // Critical path table
            var critCols = [
                { label: '#', w: 8 }, { label: 'WBS', w: 16 }, { label: 'Task Name', w: 110 },
                { label: 'Duration', w: 18 }, { label: 'Start', w: 30 }, { label: 'Finish', w: 30 },
                { label: '% Done', w: 16 }, { label: 'Status', w: 20 }
            ];

            x = margin;
            pdf.setFillColor(127, 29, 29); // red-900
            pdf.rect(margin, y, W - margin * 2, 7, 'F');
            pdf.setFontSize(6.5);
            pdf.setTextColor(255, 255, 255);
            critCols.forEach(function(c) { pdf.text(c.label, x + 2, y + 5); x += c.w; });
            y += 8;

            var todayMs = new Date().setHours(0,0,0,0);
            critTasks.forEach(function(t, i) {
                if (y > H - 14) {
                    pdf.addPage();
                    addPageHeader(pdf, W, 'Critical Path (cont.)', project.name);
                    y = 26;
                    x = margin;
                    pdf.setFillColor(127, 29, 29);
                    pdf.rect(margin, y, W - margin * 2, 7, 'F');
                    pdf.setFontSize(6.5);
                    pdf.setTextColor(255, 255, 255);
                    critCols.forEach(function(c) { pdf.text(c.label, x + 2, y + 5); x += c.w; });
                    y += 8;
                }
                if (i % 2 === 0) { pdf.setFillColor(254, 249, 249); pdf.rect(margin, y - 1, W - margin * 2, 6, 'F'); }

                var pct = t.percentComplete || 0;
                var isLate = t.finish && new Date(t.finish).getTime() < todayMs && pct < 100;
                var status = pct >= 100 ? 'Complete' : isLate ? 'LATE' : pct > 0 ? 'In Progress' : 'Not Started';

                pdf.setFontSize(6.5);
                pdf.setTextColor(isLate ? 185 : 30, isLate ? 28 : 41, isLate ? 28 : 59);
                x = margin;
                var critRow = [String(i+1), t.wbs || '', (t.name || '').substring(0, 60),
                    (t.durationDays || 0) + 'd', fmtDate(t.start), fmtDate(t.finish), pct + '%', status];
                critCols.forEach(function(c, ci) { pdf.text(critRow[ci], x + 2, y + 3); x += c.w; });
                y += 6;
            });

            // ── Critical Path Timeline Chart ──
            pdf.addPage();
            addPageHeader(pdf, W, 'Critical Path -- Timeline View', project.name);

            var cpChartY = 26;
            var cpNameW = 70; // width for task names
            var cpChartLeft = margin + cpNameW;
            var cpChartRight = W - margin;
            var cpChartW = cpChartRight - cpChartLeft;

            // Find earliest start and latest finish of critical tasks
            var cpStartMs = Infinity, cpEndMs = -Infinity;
            critTasks.forEach(function(t) {
                if (t.start) { var s = new Date(t.start).getTime(); if (s < cpStartMs) cpStartMs = s; }
                if (t.finish) { var f = new Date(t.finish).getTime(); if (f > cpEndMs) cpEndMs = f; }
            });
            var cpSpan = Math.max(1, cpEndMs - cpStartMs);

            // Dynamic row height to fit one page
            var cpAvailH = H - cpChartY - 20;
            var cpRowH = Math.min(8, Math.max(3, cpAvailH / (critTasks.length + 1)));

            // Date axis
            pdf.setFillColor(127, 29, 29);
            pdf.rect(cpChartLeft, cpChartY, cpChartW, 7, 'F');
            pdf.setFontSize(5.5);
            pdf.setTextColor(255, 255, 255);

            var cpAxisDate = new Date(cpStartMs);
            cpAxisDate.setDate(1);
            cpAxisDate.setMonth(cpAxisDate.getMonth() + 1);
            while (cpAxisDate.getTime() < cpEndMs) {
                var cpAx = cpChartLeft + ((cpAxisDate.getTime() - cpStartMs) / cpSpan) * cpChartW;
                if (cpAx > cpChartLeft + 5 && cpAx < cpChartRight - 10) {
                    pdf.text(cpAxisDate.toLocaleDateString('en-US', { month: 'short', year: '2-digit' }), cpAx, cpChartY + 5);
                    pdf.setDrawColor(245, 230, 230);
                    pdf.setLineWidth(0.1);
                    pdf.line(cpAx, cpChartY + 7, cpAx, cpChartY + 7 + critTasks.length * cpRowH);
                }
                cpAxisDate.setMonth(cpAxisDate.getMonth() + 1);
            }

            // Today line
            var cpTodayX = cpChartLeft + ((todayMs - cpStartMs) / cpSpan) * cpChartW;
            if (cpTodayX > cpChartLeft && cpTodayX < cpChartRight) {
                pdf.setDrawColor(239, 68, 68);
                pdf.setLineWidth(0.5);
                pdf.setLineDashPattern([1.5, 1.5]);
                pdf.line(cpTodayX, cpChartY, cpTodayX, cpChartY + 7 + critTasks.length * cpRowH + 2);
                pdf.setLineDashPattern([]);
                pdf.setFontSize(5);
                pdf.setTextColor(239, 68, 68);
                pdf.text('Today', cpTodayX, cpChartY - 1, { align: 'center' });
            }

            // Critical task bars
            var cpBarY = cpChartY + 8;
            critTasks.forEach(function(t, i) {
                var sMs = t.start ? new Date(t.start).getTime() : cpStartMs;
                var fMs = t.finish ? new Date(t.finish).getTime() : sMs;
                var bx = cpChartLeft + ((sMs - cpStartMs) / cpSpan) * cpChartW;
                var bw = Math.max(2, ((fMs - sMs) / cpSpan) * cpChartW);
                var bh = cpRowH - 1.5;

                // Zebra
                if (i % 2 === 0) {
                    pdf.setFillColor(254, 249, 249);
                    pdf.rect(margin, cpBarY - 0.3, W - margin * 2, cpRowH, 'F');
                }

                // Task name
                var cpName = (t.name || '').substring(0, 35);
                if (cpName.length >= 35) cpName += '...';
                pdf.setFontSize(Math.min(6, cpRowH * 0.65));
                var cpPct = t.percentComplete || 0;
                var cpLate = fMs < todayMs && cpPct < 100;
                pdf.setTextColor(cpLate ? 185 : 71, cpLate ? 28 : 29, cpLate ? 28 : 29);
                pdf.text(cpName, margin, cpBarY + bh * 0.7);

                // Bar bg (light red)
                pdf.setFillColor(254, 226, 226);
                pdf.roundedRect(bx, cpBarY + 0.2, bw, bh - 0.4, 0.6, 0.6, 'F');

                // Progress fill (red gradient)
                if (cpPct > 0) {
                    var cpFillW = bw * cpPct / 100;
                    pdf.setFillColor(cpPct >= 100 ? 34 : 239, cpPct >= 100 ? 197 : 68, cpPct >= 100 ? 94 : 68);
                    pdf.roundedRect(bx, cpBarY + 0.2, cpFillW, bh - 0.4, 0.6, 0.6, 'F');
                }

                // Percentage text on bar
                if (bw > 10) {
                    pdf.setFontSize(Math.min(5, bh * 0.6));
                    pdf.setTextColor(255, 255, 255);
                    pdf.text(cpPct + '%', bx + bw / 2, cpBarY + bh * 0.7, { align: 'center' });
                }

                cpBarY += cpRowH;
            });

            // Legend
            pdf.setFontSize(5.5);
            pdf.setTextColor(100, 116, 139);
            pdf.text(critTasks.length + ' critical tasks  |  Red = In Progress/Late  |  Green = Complete', margin, H - 6);
        }

        // ══════════════════════════════════════════
        // PAGE: NEEDS ATTENTION
        // ══════════════════════════════════════════
        pdf.addPage();
        addPageHeader(pdf, W, 'Needs Attention & Upcoming Tasks', project.name);
        y = 26;

        var todayDate = new Date(); todayDate.setHours(0,0,0,0);
        var nextWeek = new Date(todayDate); nextWeek.setDate(nextWeek.getDate() + 7);

        // Late tasks
        var lateTasks = tasks.filter(function(t) {
            var f = new Date(t.finish); f.setHours(0,0,0,0);
            return f < todayDate && t.percentComplete < 100;
        });
        var upcomingTasks = tasks.filter(function(t) {
            var f = new Date(t.finish); f.setHours(0,0,0,0);
            return f >= todayDate && f <= nextWeek && t.percentComplete < 100;
        });

        // Section: Late Tasks
        pdf.setFillColor(254, 242, 242);
        pdf.roundedRect(margin, y, W - margin * 2, 10, 2, 2, 'F');
        pdf.setFontSize(10);
        pdf.setTextColor(153, 27, 27);
        pdf.text('[!] Overdue Tasks (' + lateTasks.length + ')', margin + 5, y + 7);
        y += 14;

        if (lateTasks.length > 0) {
            var attTableW = W - margin * 2;
            var attCols = [{ label: 'Task Name', w: attTableW * 0.38 }, { label: 'Assigned To', w: attTableW * 0.20 }, { label: 'Due Date', w: attTableW * 0.14 }, { label: 'Days Late', w: attTableW * 0.14 }, { label: '% Done', w: attTableW * 0.14 }];
            x = margin;
            pdf.setFillColor(239, 68, 68);
            pdf.rect(margin, y, attTableW, 6, 'F');
            pdf.setFontSize(6.5);
            pdf.setTextColor(255, 255, 255);
            attCols.forEach(function(c) { pdf.text(c.label, x + 2, y + 4); x += c.w; });
            y += 7;
            lateTasks.slice(0, 15).forEach(function(t, i) {
                var daysLate = Math.round((todayDate - new Date(t.finish)) / 86400000);
                var rn = t.resourceNames || []; if (typeof rn === 'string') rn = rn.split(',').map(function(s){return s.trim();}).filter(Boolean);
                var resStr = rn.length > 0 ? rn.join(', ') : 'Unassigned';
                pdf.setFontSize(6.5);
                pdf.setTextColor(30, 41, 59);
                if (i % 2 === 0) { pdf.setFillColor(254, 249, 249); pdf.rect(margin, y - 1, attTableW, 5.5, 'F'); }
                // Severity Stripe
                if (daysLate > 30) { pdf.setFillColor(239,68,68); pdf.rect(margin, y-1, 2, 5.5, 'F'); }
                else if (daysLate > 14) { pdf.setFillColor(245,158,11); pdf.rect(margin, y-1, 2, 5.5, 'F'); }
                else { pdf.setFillColor(34,197,94); pdf.rect(margin, y-1, 2, 5.5, 'F'); }
                x = margin;
                pdf.setTextColor(30, 41, 59);
                pdf.text((t.name||'').substring(0,50), x + 4, y + 3); x += attCols[0].w;
                pdf.setTextColor(79, 70, 229);
                pdf.text(resStr.substring(0,28), x + 2, y + 3); x += attCols[1].w;
                pdf.setTextColor(30, 41, 59);
                pdf.text(fmtDate(t.finish), x + 2, y + 3); x += attCols[2].w;
                pdf.setTextColor(daysLate > 30 ? 239 : daysLate > 14 ? 245 : 30, daysLate > 30 ? 68 : daysLate > 14 ? 158 : 41, daysLate > 30 ? 68 : daysLate > 14 ? 11 : 59);
                pdf.text(daysLate + 'd', x + 2, y + 3); x += attCols[3].w;
                pdf.setTextColor(30, 41, 59);
                pdf.text((t.percentComplete||0) + '%', x + 2, y + 3);
                y += 5.5;
            });
        } else {
            pdf.setFontSize(9);
            pdf.setTextColor(34, 197, 94);
            pdf.text('[OK] No overdue tasks!', margin + 5, y + 6);
            y += 12;
        }
        y += 8;

        // Section: Upcoming Tasks
        pdf.setFillColor(254, 252, 232); // yellow-50
        pdf.roundedRect(margin, y, W - margin * 2, 10, 2, 2, 'F');
        pdf.setFontSize(10);
        pdf.setTextColor(133, 77, 14); // yellow-800
        pdf.text('[>] Due This Week (' + upcomingTasks.length + ')', margin + 5, y + 7);
        y += 14;

        if (upcomingTasks.length > 0) {
            var upTableW = W - margin * 2;
            var upCols = [{ label: 'Task Name', w: upTableW * 0.50 }, { label: 'Due Date', w: upTableW * 0.18 }, { label: 'Days Left', w: upTableW * 0.14 }, { label: '% Done', w: upTableW * 0.18 }];
            x = margin;
            pdf.setFillColor(245, 158, 11);
            pdf.rect(margin, y, upTableW, 6, 'F');
            pdf.setFontSize(6.5);
            pdf.setTextColor(255, 255, 255);
            upCols.forEach(function(c) { pdf.text(c.label, x + 2, y + 4); x += c.w; });
            y += 7;
            upcomingTasks.slice(0, 15).forEach(function(t, i) {
                var daysLeft = Math.round((new Date(t.finish) - todayDate) / 86400000);
                pdf.setFontSize(6.5);
                pdf.setTextColor(30, 41, 59);
                if (i % 2 === 0) { pdf.setFillColor(255, 251, 235); pdf.rect(margin, y - 1, upTableW, 5.5, 'F'); }
                x = margin;
                var row = [(t.name||'').substring(0,60), fmtDate(t.finish), daysLeft + 'd', (t.percentComplete||0) + '%'];
                upCols.forEach(function(c, ci) { pdf.text(row[ci], x + 2, y + 3); x += c.w; });
                y += 5.5;
            });
        } else {
            pdf.setFontSize(9);
            pdf.setTextColor(100, 116, 139);
            pdf.text('No tasks due this week.', margin + 5, y + 6);
            y += 12;
        }

        // ══════════════════════════════════════════
        // PAGE: S-CURVE (PLANNED vs ACTUAL)
        // ══════════════════════════════════════════
        pdf.addPage();
        addPageHeader(pdf, W, 'S-Curve \u2014 Planned vs Actual Progress', project.name);

        var chartX = margin + 10;
        var chartY = 30;
        var chartW = W - margin * 2 - 20;
        var chartH = H - 70;

        // Build time series data (weekly intervals)
        var projStart = project.startDate ? new Date(project.startDate) : new Date();
        var projEnd = project.finishDate ? new Date(project.finishDate) : new Date();
        var totalDays = Math.max(1, Math.ceil((projEnd - projStart) / 86400000));
        var numPoints = Math.min(40, Math.max(10, Math.ceil(totalDays / 7)));
        var interval = totalDays / numPoints;

        var planned = [];
        var actual = [];
        for (var p = 0; p <= numPoints; p++) {
            var dateAt = new Date(projStart.getTime() + p * interval * 86400000);
            // Planned: tasks that SHOULD be complete by this date (based on finish date)
            var plannedDone = 0;
            var actualDone = 0;
            tasks.forEach(function(t) {
                if (t.finish && new Date(t.finish) <= dateAt) plannedDone++;
                // Actual: use % complete weighted by whether task finish <= dateAt
                if (t.finish && new Date(t.finish) <= dateAt) {
                    actualDone += (t.percentComplete || 0) / 100;
                } else if (t.start && new Date(t.start) <= dateAt && t.percentComplete > 0) {
                    actualDone += (t.percentComplete || 0) / 100;
                }
            });
            planned.push(total > 0 ? (plannedDone / total) * 100 : 0);
            actual.push(total > 0 ? (actualDone / total) * 100 : 0);
        }

        // Draw chart background
        pdf.setFillColor(248, 250, 252);
        pdf.roundedRect(chartX - 8, chartY - 5, chartW + 16, chartH + 20, 3, 3, 'F');

        // Grid lines
        pdf.setDrawColor(226, 232, 240);
        pdf.setLineWidth(0.2);
        for (var g = 0; g <= 4; g++) {
            var gy = chartY + chartH - (g / 4) * chartH;
            pdf.line(chartX, gy, chartX + chartW, gy);
            pdf.setFontSize(7);
            pdf.setTextColor(100, 116, 139);
            pdf.text((g * 25) + '%', chartX - 8, gy + 2);
        }

        // X-axis labels
        for (var lbl = 0; lbl <= numPoints; lbl += Math.max(1, Math.floor(numPoints / 8))) {
            var lblDate = new Date(projStart.getTime() + lbl * interval * 86400000);
            var lblX = chartX + (lbl / numPoints) * chartW;
            pdf.setFontSize(5.5);
            pdf.setTextColor(100, 116, 139);
            pdf.text(lblDate.toLocaleDateString('en-US', { month: 'short', year: '2-digit' }), lblX, chartY + chartH + 8, { align: 'center' });
        }

        // Draw Planned line (blue)
        pdf.setDrawColor(59, 130, 246);
        pdf.setLineWidth(0.8);
        for (var i = 1; i <= numPoints; i++) {
            var x1 = chartX + ((i-1) / numPoints) * chartW;
            var y1 = chartY + chartH - (planned[i-1] / 100) * chartH;
            var x2 = chartX + (i / numPoints) * chartW;
            var y2 = chartY + chartH - (planned[i] / 100) * chartH;
            pdf.line(x1, y1, x2, y2);
        }

        // Draw Actual line (green)
        // Only draw up to "today" position
        var todayPos = Math.min(numPoints, Math.ceil((todayDate - projStart) / (interval * 86400000)));
        pdf.setDrawColor(34, 197, 94);
        pdf.setLineWidth(0.8);
        for (var i = 1; i <= todayPos && i <= numPoints; i++) {
            var x1 = chartX + ((i-1) / numPoints) * chartW;
            var y1 = chartY + chartH - (actual[i-1] / 100) * chartH;
            var x2 = chartX + (i / numPoints) * chartW;
            var y2 = chartY + chartH - (actual[i] / 100) * chartH;
            pdf.line(x1, y1, x2, y2);
        }

        // Today marker
        if (todayPos > 0 && todayPos < numPoints) {
            var todayX = chartX + (todayPos / numPoints) * chartW;
            pdf.setDrawColor(239, 68, 68);
            pdf.setLineWidth(0.4);
            pdf.setLineDashPattern([2, 2]);
            pdf.line(todayX, chartY, todayX, chartY + chartH);
            pdf.setLineDashPattern([]);
            pdf.setFontSize(6);
            pdf.setTextColor(239, 68, 68);
            pdf.text('Today', todayX, chartY - 2, { align: 'center' });
        }

        // Legend
        var legY = chartY + chartH + 15;
        pdf.setFillColor(59, 130, 246);
        pdf.rect(W / 2 - 50, legY, 8, 3, 'F');
        pdf.setFontSize(7);
        pdf.setTextColor(30, 41, 59);
        pdf.text('Planned', W / 2 - 40, legY + 3);
        pdf.setFillColor(34, 197, 94);
        pdf.rect(W / 2 + 10, legY, 8, 3, 'F');
        pdf.text('Actual', W / 2 + 20, legY + 3);
        pdf.setFillColor(239, 68, 68);
        pdf.rect(W / 2 + 50, legY, 8, 3, 'F');
        pdf.text('Today', W / 2 + 60, legY + 3);

        // ══════════════════════════════════════════
        // PAGE: EVM (EARNED VALUE MANAGEMENT)
        // ══════════════════════════════════════════
        pdf.addPage();
        addPageHeader(pdf, W, 'Earned Value Management (EVM)', project.name);
        y = 28;

        // Calculate EVM metrics
        var BAC = total; // Using task count as budget units
        var PV = 0; // Planned Value
        var EV = 0; // Earned Value
        tasks.forEach(function(t) {
            if (t.finish && new Date(t.finish) <= todayDate) PV++;
            EV += (t.percentComplete || 0) / 100;
        });
        var pvPct = BAC > 0 ? Math.round((PV / BAC) * 100) : 0;
        var evPct = BAC > 0 ? Math.round((EV / BAC) * 100) : 0;
        var SV = EV - PV;
        var SPI = PV > 0 ? (EV / PV) : 0;
        var svPct = BAC > 0 ? Math.round((SV / BAC) * 100) : 0;
        var EAC = SPI > 0 ? Math.round(BAC / SPI) : BAC;
        var TCPI = (BAC - EV) > 0 && (BAC - PV) !== 0 ? ((BAC - EV) / (BAC - PV)) : 1;

        // EVM KPI Cards (2 rows of 4)
        var evmCards = [
            { label: 'BAC\n(Budget At Completion)', value: BAC + ' tasks', color: [79, 70, 229], sub: 'Total scope' },
            { label: 'PV\n(Planned Value)', value: pvPct + '%', color: [59, 130, 246], sub: PV + ' tasks planned done' },
            { label: 'EV\n(Earned Value)', value: evPct + '%', color: [34, 197, 94], sub: Math.round(EV) + ' tasks earned' },
            { label: 'SV\n(Schedule Variance)', value: svPct + '%', color: SV >= 0 ? [34, 197, 94] : [239, 68, 68], sub: SV >= 0 ? 'Ahead' : 'Behind' },
            { label: 'SPI\n(Schedule Perf Index)', value: SPI.toFixed(2), color: SPI >= 1 ? [34, 197, 94] : SPI >= 0.9 ? [245, 158, 11] : [239, 68, 68], sub: SPI >= 1 ? 'On/Ahead' : 'Behind' },
            { label: 'EAC\n(Est. At Completion)', value: EAC + ' tasks', color: EAC <= BAC ? [34, 197, 94] : [239, 68, 68], sub: EAC <= BAC ? 'Under budget' : 'Over budget' },
            { label: 'TCPI\n(To Complete Perf)', value: TCPI.toFixed(2), color: TCPI <= 1 ? [34, 197, 94] : [239, 68, 68], sub: TCPI <= 1 ? 'Achievable' : 'Challenging' },
            { label: 'Progress\n(Overall)', value: progress + '%', color: [79, 70, 229], sub: complete + '/' + total + ' complete' },
        ];

        var cardW = 60;
        var cardH = 35;
        var gap = 6;
        var cardsPerRow = 4;
        var rowStartX = (W - (cardsPerRow * cardW + (cardsPerRow - 1) * gap)) / 2;

        evmCards.forEach(function(card, i) {
            var row = Math.floor(i / cardsPerRow);
            var col = i % cardsPerRow;
            var cx = rowStartX + col * (cardW + gap);
            var cy = y + row * (cardH + gap + 4);

            // Card background
            pdf.setFillColor(248, 250, 252);
            pdf.roundedRect(cx, cy, cardW, cardH, 3, 3, 'F');
            // Color accent
            pdf.setFillColor(card.color[0], card.color[1], card.color[2]);
            pdf.rect(cx, cy + 4, 3, cardH - 8, 'F');

            // Label
            pdf.setFontSize(6);
            pdf.setTextColor(100, 116, 139);
            var labelLines = card.label.split('\n');
            pdf.text(labelLines[0], cx + cardW / 2, cy + 8, { align: 'center' });
            if (labelLines[1]) {
                pdf.setFontSize(5);
                pdf.text(labelLines[1], cx + cardW / 2, cy + 12, { align: 'center' });
            }

            // Value
            pdf.setFontSize(16);
            pdf.setTextColor(card.color[0], card.color[1], card.color[2]);
            pdf.text(card.value, cx + cardW / 2, cy + 24, { align: 'center' });

            // Sub text
            pdf.setFontSize(5.5);
            pdf.setTextColor(148, 163, 184);
            pdf.text(card.sub, cx + cardW / 2, cy + 30, { align: 'center' });
        });

        // EVM Health Summary
        var healthY = y + 2 * (cardH + gap + 4) + 10;
        var healthW = W - margin * 4;
        var healthX = (W - healthW) / 2;
        pdf.setFillColor(SPI >= 0.95 ? 240 : SPI >= 0.85 ? 254 : 254, SPI >= 0.95 ? 253 : SPI >= 0.85 ? 252 : 242, SPI >= 0.95 ? 244 : SPI >= 0.85 ? 232 : 242);
        pdf.roundedRect(healthX, healthY, healthW, 20, 3, 3, 'F');
        pdf.setFontSize(11);
        var healthColor = SPI >= 0.95 ? [22, 163, 74] : SPI >= 0.85 ? [202, 138, 4] : [220, 38, 38];
        pdf.setTextColor(healthColor[0], healthColor[1], healthColor[2]);
        var healthIcon = SPI >= 0.95 ? '[OK]' : SPI >= 0.85 ? '[!!]' : '[!!!]';
        var healthText = SPI >= 0.95 ? 'Project Health: GOOD -- Schedule is on track (SPI = ' + SPI.toFixed(2) + ')' :
            SPI >= 0.85 ? 'Project Health: WARNING -- Slight schedule delay (SPI = ' + SPI.toFixed(2) + ')' :
            'Project Health: AT RISK -- Significant schedule variance (SPI = ' + SPI.toFixed(2) + ')';
        pdf.text(healthIcon + '  ' + healthText, W / 2, healthY + 13, { align: 'center' });

        // ── PAGE: Gantt Chart (single page) ──
        pdf.addPage();
        addPageHeader(pdf, W, 'Gantt Chart -- Summary View', project.name);

        var ganttY = 24;
        var ganttLeft = margin + 80; // space for task names
        var ganttRight = W - margin;
        var ganttW = ganttRight - ganttLeft;
        var availH = H - ganttY - 14; // available height

        var projStartMs = projStart.getTime();
        var projEndMs = projEnd.getTime();
        var projSpan = Math.max(1, projEndMs - projStartMs);

        // Filter tasks to fit one page: show level 1, 2, 3 only
        var ganttTasks = project.tasks.filter(function(t) {
            return (t.outlineLevel || 1) <= 3;
        });
        // If still too many, show level 1-2 only
        if (ganttTasks.length > 60) {
            ganttTasks = project.tasks.filter(function(t) {
                return (t.outlineLevel || 1) <= 2;
            });
        }
        // If STILL too many, show summaries only
        if (ganttTasks.length > 60) {
            ganttTasks = project.tasks.filter(function(t) {
                return t.summary || (t.outlineLevel || 1) <= 1;
            });
        }

        // Dynamic row height
        var rowH = Math.min(6, Math.max(2.5, (availH - 8) / ganttTasks.length));
        var totalGanttH = ganttTasks.length * rowH;

        // Date axis header
        pdf.setFillColor(30, 27, 75);
        pdf.rect(ganttLeft, ganttY, ganttW, 7, 'F');
        pdf.setFontSize(5.5);
        pdf.setTextColor(255, 255, 255);

        // Monthly markers
        var axisDate = new Date(projStart);
        axisDate.setDate(1);
        axisDate.setMonth(axisDate.getMonth() + 1);
        while (axisDate < projEnd) {
            var axisX = ganttLeft + ((axisDate.getTime() - projStartMs) / projSpan) * ganttW;
            if (axisX > ganttLeft && axisX < ganttRight - 10) {
                pdf.text(axisDate.toLocaleDateString('en-US', { month: 'short', year: '2-digit' }), axisX, ganttY + 5);
                // Grid line
                pdf.setDrawColor(235, 237, 243);
                pdf.setLineWidth(0.1);
                pdf.line(axisX, ganttY + 7, axisX, ganttY + 7 + totalGanttH);
            }
            axisDate.setMonth(axisDate.getMonth() + 1);
        }

        // Today vertical line
        var todayGanttX = ganttLeft + ((todayDate.getTime() - projStartMs) / projSpan) * ganttW;
        if (todayGanttX > ganttLeft && todayGanttX < ganttRight) {
            pdf.setDrawColor(239, 68, 68);
            pdf.setLineWidth(0.4);
            pdf.setLineDashPattern([1.5, 1.5]);
            pdf.line(todayGanttX, ganttY, todayGanttX, ganttY + 7 + totalGanttH);
            pdf.setLineDashPattern([]);
            pdf.setFontSize(5);
            pdf.setTextColor(239, 68, 68);
            pdf.text('Today', todayGanttX, ganttY - 1, { align: 'center' });
        }

        // Draw task bars
        var barY = ganttY + 8;
        var fontSize = Math.min(5, Math.max(3.5, rowH * 0.7));
        ganttTasks.forEach(function(t, idx) {
            var sMs = t.start ? new Date(t.start).getTime() : projStartMs;
            var fMs = t.finish ? new Date(t.finish).getTime() : sMs;
            var bx = ganttLeft + ((sMs - projStartMs) / projSpan) * ganttW;
            var bw = Math.max(1, ((fMs - sMs) / projSpan) * ganttW);
            var bh = rowH - 1;

            // Zebra stripe
            if (idx % 2 === 0) {
                pdf.setFillColor(248, 250, 252);
                pdf.rect(margin, barY - 0.3, W - margin * 2, rowH, 'F');
            }

            // Task name
            pdf.setFontSize(fontSize);
            var indent = '';
            for (var li = 1; li < (t.outlineLevel || 1); li++) indent += '  ';
            var tName = indent + (t.name || '');
            if (tName.length > 42) tName = tName.substring(0, 39) + '...';

            if (t.summary) {
                pdf.setTextColor(30, 27, 75);
                pdf.text(tName, margin, barY + bh * 0.75);
                // Summary bar - thin dark line
                pdf.setFillColor(30, 27, 75);
                pdf.rect(bx, barY + bh * 0.35, bw, bh * 0.3, 'F');
                // End diamonds
                pdf.triangle(bx, barY + bh * 0.2, bx + 1, barY + bh * 0.5, bx, barY + bh * 0.8, 'F');
                pdf.triangle(bx + bw, barY + bh * 0.2, bx + bw - 1, barY + bh * 0.5, bx + bw, barY + bh * 0.8, 'F');
            } else {
                var isCritical = t.critical;
                pdf.setTextColor(isCritical ? 185 : 71, isCritical ? 28 : 85, isCritical ? 28 : 105);
                pdf.text(tName, margin, barY + bh * 0.75);

                // Bar background
                pdf.setFillColor(isCritical ? 254 : 224, isCritical ? 226 : 231, isCritical ? 226 : 255);
                pdf.roundedRect(bx, barY + 0.2, bw, bh - 0.4, 0.5, 0.5, 'F');

                // Progress fill
                var pct = t.percentComplete || 0;
                if (pct > 0) {
                    var fillW = bw * pct / 100;
                    pdf.setFillColor(isCritical ? 239 : 79, isCritical ? 68 : 70, isCritical ? 68 : 229);
                    pdf.roundedRect(bx, barY + 0.2, fillW, bh - 0.4, 0.5, 0.5, 'F');
                }

                // Milestone
                if (t.milestone || (t.duration === 0 && !t.summary)) {
                    pdf.setFillColor(79, 70, 229);
                    var mx = bx;
                    var my = barY + bh / 2;
                    pdf.triangle(mx, my - 1.2, mx + 1.2, my, mx, my + 1.2, 'F');
                    pdf.triangle(mx, my - 1.2, mx - 1.2, my, mx, my + 1.2, 'F');
                }
            }

            barY += rowH;
        });

        // Legend at bottom
        pdf.setFontSize(5.5);
        pdf.setTextColor(100, 116, 139);
        pdf.text('Showing ' + ganttTasks.length + ' of ' + project.tasks.length + ' tasks (summary view, levels 1-3)', margin, H - 6);

        // Save
        var filename = sanitize(project.name) + '_report.pdf';
        pdf.save(filename);
        return filename;
    }

    function addPageHeader(pdf, W, title, projectName) {
        // Top accent
        pdf.setFillColor(79, 70, 229);
        pdf.rect(0, 0, W, 3, 'F');
        // Title
        pdf.setFontSize(11);
        pdf.setTextColor(30, 27, 75);
        pdf.text(title, 14, 13);
        // Right side
        pdf.setFontSize(7.5);
        pdf.setTextColor(100, 116, 139);
        pdf.text(projectName + '  |  ' + new Date().toLocaleDateString(), W - 14, 13, { align: 'right' });
        // Line
        pdf.setDrawColor(203, 213, 225); // slate-300
        pdf.line(14, 17, W - 14, 17);
    }

    // ─── Print View ───
    function printProject() {
        window.print();
    }

    // ─── Helpers ───
    function fmtDate(d) {
        if (!d) return '';
        const dt = new Date(d);
        return dt.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
    }
    function sanitize(n) { return (n || 'project').replace(/[^a-zA-Z0-9_-]/g, '_').substring(0, 50); }

    // ══════════════════════════════════════════════════════
    //  PORTFOLIO PDF REPORT — Enterprise Multi-Page
    // ══════════════════════════════════════════════════════
    async function generatePortfolioPDF(projects, settings) {
        if (typeof jspdf === 'undefined') throw new Error('jsPDF not loaded');
        var pdf = new jspdf.jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
        var W = 297, H = 210, M = 14;
        var cur = (settings && settings.currency) || '$';
        var today = new Date(); today.setHours(0,0,0,0);
        var pageNum = 0;
        var C = {
            indigo:[79,70,229], deepIndigo:[30,27,75], purple:[139,92,246],
            blue:[59,130,246], green:[34,197,94], yellow:[245,158,11],
            red:[239,68,68], slate50:[248,250,252], slate100:[241,245,249],
            slate200:[226,232,240], slate300:[203,213,225], slate500:[100,116,139],
            slate700:[51,65,85], slate900:[30,41,59], white:[255,255,255]
        };
        function sc(c){pdf.setFillColor(c[0],c[1],c[2]);}
        function tc(c){pdf.setTextColor(c[0],c[1],c[2]);}
        function addFt(){
            pageNum++;
            // Footer bar
            sc(C.slate100);pdf.rect(0,H-10,W,10,'F');
            pdf.setFontSize(5.5);tc(C.slate500);
            pdf.text('Generated by ProjectFlow\u2122  \u2014  Confidential',M,H-4.5);
            pdf.text(new Date().toLocaleDateString('en-US',{year:'numeric',month:'long',day:'numeric'}),W/2,H-4.5,{align:'center'});
            pdf.text('Page '+pageNum,W-M,H-4.5,{align:'right'});
        }
        function addPH(t){
            // Top accent: indigo + purple gradient simulation (3 bands)
            sc(C.indigo);pdf.rect(0,0,W*0.6,4,'F');
            sc(C.purple);pdf.rect(W*0.6,0,W*0.25,4,'F');
            sc([236,72,153]);pdf.rect(W*0.85,0,W*0.15,4,'F');
            // Header row
            sc(C.slate50);pdf.rect(0,4,W,14,'F');
            pdf.setFontSize(12);tc(C.deepIndigo);
            pdf.setFont('helvetica','bold');pdf.text(t,M,14.5);pdf.setFont('helvetica','normal');
            pdf.setFontSize(7);tc(C.slate500);
            pdf.text('Portfolio Report  \u00B7  '+fmtDate(today),W-M,14.5,{align:'right'});
            // Subtle divider
            sc(C.indigo);pdf.rect(M,18,40,0.6,'F');
            sc(C.slate200);pdf.rect(M+40,18,W-M*2-40,0.6,'F');
        }
        function calculateHealthScore(pData) {
            var tasks = pData.tasks||[]; if(!tasks.length) return {s:0, l:'Not Started', c:C.slate500};
            var total = tasks.length, completePct = tasks.filter(function(t){return (t.percentComplete||0)>=100;}).length/total;
            var score = 100;
            var lateTasks = tasks.filter(function(t){return (t.percentComplete||0)<100 && new Date(t.finish).getTime()<today.getTime();});
            score -= Math.min(30, lateTasks.length*5);
            var critStalled = tasks.filter(function(t){return t.critical && (t.percentComplete||0)<50 && new Date(t.start).getTime()<today.getTime();});
            score -= Math.min(20, critStalled.length*3);
            if(pData.sd && pData.fd) {
                var ps = new Date(pData.sd).getTime(), pf = new Date(pData.fd).getTime(), dur = pf-ps;
                if(dur>0 && today.getTime()>ps) {
                    var elapsed = Math.min(1, (today.getTime()-ps)/dur);
                    var avgPct = tasks.reduce(function(s,t){return s+(t.percentComplete||0);},0)/total/100;
                    var spi = elapsed>0 ? avgPct/elapsed : 1;
                    if(spi<0.7) score -= 20; else if(spi<0.85) score -= 10; else if(spi<0.95) score -= 5;
                }
            }
            if(completePct>=0.9) score = Math.max(score, 90);
            if(completePct>=1) score = 100;
            score = Math.max(0, Math.min(100, Math.round(score)));
            if(completePct>=1||score===100) return {s:100, l:'Complete', c:C.green};
            if(score>=75) return {s:score, l:'Healthy', c:[52,211,153]};
            if(score>=45) return {s:score, l:'At Risk', c:C.yellow};
            return {s:score, l:'Critical', c:C.red};
        }

        // ── DATA ANALYSIS ──
        var tTasks=0,cTasks=0,crCount=0,tCost=0,hCounts={Complete:0,Healthy:0,'At Risk':0,Critical:0,'Not Started':0};
        var allLate=[],allRes=[],pData=[],pUP=[],pCrit=[],evmTPV=0,evmTEV=0;
        allRes={}; // fix from array
        var gStart = Infinity, gEnd = -Infinity;
        projects.forEach(function(p){
            if(p.startDate){var d=new Date(p.startDate).getTime(); if(d<gStart)gStart=d;}
            if(p.finishDate){var fd=new Date(p.finishDate).getTime(); if(fd>gEnd)gEnd=fd;}
            
            var lf=p.tasks||[];var tasks=lf.filter(function(t){return !t.summary;});
            var pT=tasks.length,pC=tasks.filter(function(t){return(t.percentComplete||0)>=100;}).length;
            var pCr=tasks.filter(function(t){return t.critical;}).length;
            var pCo=tasks.reduce(function(s,t){return s+(t.cost||0);},0);
            var pP=pT>0?Math.round(tasks.reduce(function(s,t){return s+(t.percentComplete||0);},0)/pT):0;
            var pLate=[];
            tasks.forEach(function(t){
                var rn=t.resourceNames||[];if(typeof rn==='string')rn=rn.split(',').map(function(s){return s.trim();}).filter(Boolean);
                var resStr=rn.length>0?rn.join(', '):'Unassigned';
                var tsd=new Date(t.start); tsd.setHours(0,0,0,0); var ts=tsd.getTime();
                var tfd=new Date(t.finish); tfd.setHours(0,0,0,0); var tf=tfd.getTime();
                
                if(tf<today.getTime()&&(t.percentComplete||0)<100){pLate.push({name:t.name,proj:p.name,days:Math.round((today.getTime()-tf)/864e5),pct:t.percentComplete||0,crit:t.critical,res:resStr});}
                else if(ts>today.getTime()&&ts<today.getTime()+14*864e5&&(t.percentComplete||0)<100){pUP.push({name:t.name,proj:p.name,start:t.start,res:resStr});}
                if(t.critical&&(t.percentComplete||0)<100){pCrit.push({name:t.name,proj:p.name,s:ts,f:tf,pct:t.percentComplete||0});}
                
                var dur = tf - ts || 1;
                var planP = today.getTime() < ts ? 0 : (today.getTime() > tf ? 100 : ((today.getTime()-ts)/dur)*100);
                evmTPV += planP; evmTEV += (t.percentComplete||0);
                
                rn.forEach(function(r){if(!allRes[r])allRes[r]={name:r,projs:[],tc:0,work:0};if(allRes[r].projs.indexOf(p.name)===-1)allRes[r].projs.push(p.name);allRes[r].tc++;allRes[r].work+=(t.durationDays||0)*((100-(t.percentComplete||0))/100);});
            });
            tTasks+=pT;cTasks+=pC;crCount+=pCr;tCost+=pCo;allLate=allLate.concat(pLate);
            var h=calculateHealthScore({tasks: tasks, sd: p.startDate, fd: p.finishDate});hCounts[h.l]=(hCounts[h.l]||0)+1;
            var ptpv=0, ptev=0;
            tasks.forEach(function(t){
                var dur = new Date(t.finish).getTime() - new Date(t.start).getTime() || 1;
                var planP = today.getTime() < new Date(t.start).getTime() ? 0 : (today.getTime() > new Date(t.finish).getTime() ? 100 : ((today.getTime()-new Date(t.start).getTime())/dur)*100);
                ptpv += planP; ptev += (t.percentComplete||0);
            });
            var pspi = ptpv>0 ? (ptev/ptpv).toFixed(2) : 1.0;
            var pSCurve = [];
            var psd = p.startDate?new Date(p.startDate).getTime():gStart, pfd = p.finishDate?new Date(p.finishDate).getTime():gEnd;
            var prng = Math.max(864e5, pfd - psd);
            for(var cpt=0; cpt<=20; cpt++){
                var ctDate = psd + (cpt/20)*prng;
                var ctotP=0, ctotE=0, ccount=0;
                tasks.forEach(function(t){
                    var ts = new Date(t.start).getTime(), tf = new Date(t.finish).getTime(), dur = tf - ts || 1;
                    var planP = ctDate < ts ? 0 : (ctDate > tf ? 100 : ((ctDate-ts)/dur)*100);
                    var actP = 0;
                    if(ctDate <= today.getTime()){
                         actP = planP * ((t.percentComplete||0)/(planP||1));
                         if(actP > t.percentComplete) actP = t.percentComplete;
                    }
                    ctotP += planP; ctotE += actP; ccount++;
                });
                pSCurve.push({ x: ctDate, pv: ccount?ctotP/ccount:0, ev: ccount?(ctDate>today.getTime()?null:ctotE/ccount):null });
            }
            pData.push({name:p.name||'Untitled',mgr:p.manager||'\u2014',sd:p.startDate,fd:p.finishDate,total:pT,done:pC,crit:pCr,cost:pCo,prog:pP,late:pLate.length,h:h,allT:lf,leafT:tasks, ptpv:ptpv, ptev:ptev, pspi:pspi, pSCurve:pSCurve});
        });
        if(!isFinite(gStart)||!isFinite(gEnd)){ gStart=today.getTime()-864e5*30; gEnd=today.getTime()+864e5*60; }
        var gRange = gEnd - gStart;
        var spi = evmTPV>0 ? (evmTEV/evmTPV).toFixed(2) : 1.0;

        var evmPoints = 20, sCurve = [];
        for(var pt=0; pt<=evmPoints; pt++){
            var tDate = gStart + (pt/evmPoints)*gRange;
            var totP=0, totE=0, count=0;
            pData.forEach(function(p){
                p.leafT.forEach(function(t){
                    var ts = new Date(t.start).getTime(), tf = new Date(t.finish).getTime();
                    var dur = tf - ts || 1;
                    var planP = tDate < ts ? 0 : (tDate > tf ? 100 : ((tDate-ts)/dur)*100);
                    var actP = 0;
                    if(tDate <= today.getTime()){
                         actP = planP * ((t.percentComplete||0)/(planP||1));
                         if(actP > t.percentComplete) actP = t.percentComplete;
                    }
                    totP += planP; totE += actP; count++;
                });
            });
            sCurve.push({ x: tDate, pv: count?totP/count:0, ev: count?(tDate>today.getTime()?null:totE/count):null });
        }
        
        var avgP=pData.length>0?Math.round(pData.reduce(function(s,p){return s+p.prog;},0)/pData.length):0;
        allLate.sort(function(a,b){return b.days-a.days;});
        var sharedR=[],overR=[];
        Object.keys(allRes).forEach(function(k){var r=allRes[k];if(r.projs.length>1)sharedR.push(r);if(r.work>60)overR.push(r);});
        sharedR.sort(function(a,b){return b.projs.length-a.projs.length;});

        // ══════ PAGE 1: COVER ══════
        // Background: deep navy with angled accent
        sc(C.deepIndigo);pdf.rect(0,0,W,62,'F');
        // Angled accent stripe (simulated with triangle shapes)
        sc([49,46,129]);
        pdf.triangle(W*0.55,0, W,0, W,62,'F');
        sc([79,70,229]);
        pdf.triangle(W*0.72,0, W,0, W,38,'F');
        sc([139,92,246]);
        pdf.triangle(W*0.86,0, W,0, W,18,'F');

        // Logo area
        try{var logo=localStorage.getItem('pf_report_logo')||(typeof PROART_LOGO!=='undefined'?PROART_LOGO:null);
            if(logo){
                await new Promise(function(resolve){
                    var im=new Image();
                    im.onload=function(){
                        var lw=40,lh=24;
                        if(im.naturalWidth&&im.naturalHeight){var ra=im.naturalWidth/im.naturalHeight;if(ra>1){lh=lw/ra;}else{lw=lh*ra;}}
                        pdf.addImage(logo,'PNG',M,7,lw,lh);
                        resolve();
                    }; im.onerror=resolve; im.src=logo;
                });
            }
        }catch(e){}

        // Title block
        pdf.setFont('helvetica','bold');
        tc(C.white);pdf.setFontSize(24);pdf.text('Portfolio Executive Report',M+50,22);
        pdf.setFont('helvetica','normal');
        pdf.setFontSize(8.5);tc([196,181,253]);
        pdf.text(new Date().toLocaleDateString('en-US',{year:'numeric',month:'long',day:'numeric'}),M+50,30);
        // Separator line
        sc([196,181,253]);pdf.rect(M+50,33,100,0.4,'F');
        pdf.setFontSize(7.5);tc([196,181,253]);
        pdf.text(pData.length+' Projects  \u00B7  '+tTasks+' Tasks  \u00B7  SPI: '+spi,M+50,38.5);

        // Big progress arc (right side)
        var pcx=W-M-22, pcy=31, prad=18;
        // Background circle
        sc([49,46,129]);pdf.circle(pcx,pcy,prad,'F');
        // Progress arc (drawn as triangle fans)
        var pgPct=avgP/100, pgSegs=Math.max(4,Math.round(pgPct*30));
        var pgCol=avgP>=75?C.green:avgP>=50?C.blue:C.yellow;
        sc(pgCol);
        var pgStart=-Math.PI/2, pgSweep=pgPct*Math.PI*2;
        for(var si2=0;si2<pgSegs;si2++){
            var pa1=pgStart+(si2/pgSegs)*pgSweep, pa2=pgStart+((si2+1)/pgSegs)*pgSweep;
            pdf.triangle(pcx,pcy,pcx+Math.cos(pa1)*prad,pcy+Math.sin(pa1)*prad,pcx+Math.cos(pa2)*prad,pcy+Math.sin(pa2)*prad,'F');
        }
        // Inner circle (donut hole)
        sc(C.deepIndigo);pdf.circle(pcx,pcy,prad*0.62,'F');
        pdf.setFont('helvetica','bold');
        tc(C.white);pdf.setFontSize(14);pdf.text(avgP+'%',pcx,pcy+1.5,{align:'center'});
        pdf.setFont('helvetica','normal');
        pdf.setFontSize(4.5);tc([196,181,253]);pdf.text('PROGRESS',pcx,pcy+7,{align:'center'});

        // ── Executive Summary Strip
        var arC=(hCounts['At Risk']||0)+(hCounts.Critical||0);
        sc([37,34,100]);pdf.roundedRect(M,50,W-M*2,10,2,2,'F');
        sc(arC>0?C.red:C.green);pdf.roundedRect(M,50,2.5,10,1,1,'F');
        pdf.setFontSize(7.5);tc(C.white);
        var summTxt2='SPI: '+spi+'   \u00B7   '+(arC>0?arC+' projects need attention':'All projects healthy')+'   \u00B7   '+allLate.length+' overdue tasks   \u00B7   '+crCount+' critical tasks';
        pdf.text(summTxt2,W/2,56.5,{align:'center'});

        // ── Health Distribution Strip ──
        var hdY=65;
        var hdTotal=pData.length||1;
        var hdSegs=[
            {l:'Complete',n:hCounts.Complete||0,c:[99,102,241]},
            {l:'Healthy',n:hCounts.Healthy||0,c:[34,197,94]},
            {l:'At Risk',n:hCounts['At Risk']||0,c:[245,158,11]},
            {l:'Critical',n:hCounts.Critical||0,c:[239,68,68]},
            {l:'Not Started',n:hCounts['Not Started']||0,c:[100,116,139]},
        ].filter(function(s){return s.n>0;});
        pdf.setFontSize(5.5);tc(C.slate500);pdf.text('HEALTH DISTRIBUTION',M,hdY);
        var hdBarY=hdY+2, hdBarH=4, hdBarX=M, hdBarW=W-M*2;
        var hdCurX=hdBarX;
        hdSegs.forEach(function(seg,si){
            var sw=(seg.n/hdTotal)*hdBarW;
            sc(seg.c);
            if(si===0&&hdSegs.length>1)pdf.roundedRect(hdCurX,hdBarY,sw,hdBarH,1.5,1.5,'F');
            else if(si===hdSegs.length-1)pdf.roundedRect(hdCurX,hdBarY,sw,hdBarH,1.5,1.5,'F');
            else pdf.rect(hdCurX,hdBarY,sw,hdBarH,'F');
            hdCurX+=sw;
        });
        // Legend
        var lgX=M; hdSegs.forEach(function(seg){
            sc(seg.c);pdf.circle(lgX+1.5,hdBarY+hdBarH+4,1.5,'F');
            pdf.setFontSize(5);tc(C.slate700);pdf.text(seg.l+' ('+seg.n+')',lgX+5,hdBarY+hdBarH+5.5);
            lgX+=34;
        });

        // ── KPI Cards ──
        var ky=hdY+hdBarH+13;
        var kCards=[
            {l:'PROJECTS',     v:String(pData.length),              c:C.indigo,  icon:'\u25A3'},
            {l:'TOTAL TASKS',  v:String(tTasks),                    c:C.blue,    icon:'\u2714'},
            {l:'COMPLETED',    v:String(cTasks),                    c:C.green,   icon:'\u2605'},
            {l:'CRITICAL',     v:String(crCount),                   c:crCount>0?C.red:[100,116,139],  icon:'\u26A0'},
            {l:'LATE TASKS',   v:String(allLate.length),            c:allLate.length>0?C.red:C.green, icon:'\u23F0'},
            {l:'BUDGET',       v:cur+tCost.toLocaleString(),        c:C.purple,  icon:'\u2219'},
        ];
        var kw=(W-M*2-5*5)/6, kh=26;
        kCards.forEach(function(k,i){
            var kx=M+i*(kw+5);
            // Card background with subtle shadow effect
            sc(C.slate50);pdf.roundedRect(kx,ky,kw,kh,2.5,2.5,'F');
            // Top accent bar
            sc(k.c);pdf.roundedRect(kx,ky,kw,3,2.5,2.5,'F');pdf.rect(kx,ky+1.5,kw,1.5,'F');
            // Label
            pdf.setFontSize(5);tc(C.slate500);
            pdf.text(k.l,kx+kw/2,ky+9,{align:'center'});
            // Value
            pdf.setFont('helvetica','bold');pdf.setFontSize(12);tc(k.c);
            pdf.text(k.v.length>8?k.v.substring(0,8):k.v, kx+kw/2, ky+20,{align:'center'});
            pdf.setFont('helvetica','normal');
        });

        // ── BAR CHART: Progress Comparison ──
        var cy=ky+kh+10;
        pdf.setFontSize(8);tc(C.deepIndigo);pdf.setFont('helvetica','bold');
        pdf.text('Progress Comparison',M,cy);pdf.setFont('helvetica','normal');cy+=5;
        var cH=40,cW=(W-M*2)*0.55;
        var bW=Math.min(30,(cW-10)/pData.length-4),bG=4;
        var tbW=pData.length*(bW+bG)-bG,bx0=M+(cW-tbW)/2;
        // Chart background
        sc(C.slate50);pdf.roundedRect(M-2,cy-2,cW+4,cH+12,2,2,'F');
        for(var g=0;g<=4;g++){var gy=cy+cH-(g/4)*cH;pdf.setDrawColor(C.slate200[0],C.slate200[1],C.slate200[2]);pdf.line(M,gy,M+cW,gy);pdf.setFontSize(5);tc(C.slate500);pdf.text((g*25)+'%',M-2,gy+1.5,{align:'right'});}
        pData.forEach(function(p,i){
            var bx=bx0+i*(bW+bG);var bh=Math.max(1,(p.prog/100)*cH);var by=cy+cH-bh;
            // Bar shadow
            sc([200,200,210]);pdf.roundedRect(bx+0.5,by+0.5,bW,bh,1.5,1.5,'F');
            // Bar fill
            sc(p.h.c);pdf.roundedRect(bx,by,bW,bh,1.5,1.5,'F');
            // % label
            pdf.setFont('helvetica','bold');pdf.setFontSize(6);tc(p.h.c);
            pdf.text(p.prog+'%',bx+bW/2,by-2,{align:'center'});pdf.setFont('helvetica','normal');
            pdf.setFontSize(4.5);tc(C.slate500);
            pdf.text(p.name.length>12?p.name.substring(0,10)+'\u2026':p.name,bx+bW/2,cy+cH+5,{align:'center'});
        });

        // ── DONUT CHART: Health ──
        var dx=M+cW+25,dcy=cy+cH/2+2,dr=18;
        pdf.setFontSize(9);tc(C.deepIndigo);pdf.text('Health Distribution',dx-10,cy);
        var hd=[{l:'Complete',n:hCounts.Complete||0,c:C.green},{l:'Healthy',n:hCounts.Healthy||0,c:[52,211,153]},{l:'At Risk',n:hCounts['At Risk']||0,c:C.yellow},{l:'Critical',n:hCounts.Critical||0,c:C.red},{l:'Not Started',n:hCounts['Not Started']||0,c:C.slate500}].filter(function(x){return x.n>0;});
        if(hd.length>0){
            var tH=hd.reduce(function(s,x){return s+x.n;},0);var sa=-Math.PI/2;
            hd.forEach(function(x){var sw=(x.n/tH)*Math.PI*2;var ea=sa+sw;sc(x.c);var segs=Math.max(8,Math.round(sw*15));
                for(var si=0;si<segs;si++){var a1=sa+(si/segs)*sw;var a2=sa+((si+1)/segs)*sw;pdf.triangle(dx,dcy,dx+Math.cos(a1)*dr,dcy+Math.sin(a1)*dr,dx+Math.cos(a2)*dr,dcy+Math.sin(a2)*dr,'F');}
                sa=ea;});
            sc(C.white);pdf.circle(dx,dcy,dr*0.55,'F');
            pdf.setFontSize(10);tc(C.deepIndigo);pdf.text(tH.toString(),dx,dcy+1,{align:'center'});
            pdf.setFontSize(4.5);tc(C.slate500);pdf.text('Projects',dx,dcy+5,{align:'center'});
            var lgy=dcy-dr;hd.forEach(function(x){sc(x.c);pdf.circle(dx+dr+8,lgy+1.5,1.5,'F');pdf.setFontSize(6);tc(C.slate700);pdf.text(x.l+' ('+x.n+')',dx+dr+12,lgy+3);lgy+=8;});
        }
        addFt();

        // ══════ PAGE 2: PROJECT TABLE ══════
        pdf.addPage();addPH('Project Comparison Dashboard');var ty=22;
        // Mini health distribution bar under header
        var mhdSegs=[
            {n:hCounts.Complete||0,c:[99,102,241]},
            {n:hCounts.Healthy||0,c:[34,197,94]},
            {n:hCounts['At Risk']||0,c:[245,158,11]},
            {n:hCounts.Critical||0,c:[239,68,68]},
        ].filter(function(s){return s.n>0;});
        var mhdTotal=pData.length||1, mhdX=M, mhdW=W-M*2, mhBarH=2.5;
        mhdSegs.forEach(function(s){ sc(s.c);pdf.rect(mhdX,ty,Math.max(2,(s.n/mhdTotal)*mhdW),mhBarH,'F');mhdX+=(s.n/mhdTotal)*mhdW;});
        ty+=mhBarH+4;
        // Summary line
        pdf.setFontSize(6);tc(C.slate500);
        pdf.text(pData.length+' projects  \u00B7  avg progress '+avgP+'%  \u00B7  SPI '+spi+'  \u00B7  '+allLate.length+' overdue',M,ty);ty+=5;
        var cs=[{l:'#',w:7},{l:'Project Name',w:55},{l:'Tasks',w:14},{l:'Done',w:16},{l:'Progress',w:42},{l:'Critical',w:14},{l:'Late',w:12},{l:'Budget',w:24},{l:'Health',w:18},{l:'Score',w:12},{l:'Start',w:22},{l:'Finish',w:22}];
        var rh=8;sc(C.deepIndigo);pdf.roundedRect(M,ty,W-M*2,rh,1,1,'F');pdf.setFontSize(5.5);tc(C.white);var hx=M+2;cs.forEach(function(c){pdf.text(c.l,hx,ty+5.5);hx+=c.w;});ty+=rh;
        pData.forEach(function(p,idx){
            if(ty+rh>H-16){addFt();pdf.addPage();addPH('Project Comparison (cont.)');ty=22;}
            if(idx%2===0){sc(C.slate50);pdf.rect(M,ty,W-M*2,rh,'F');}
            // RAG Traffic Light
            var ragC=p.h.l==='Complete'||p.h.l==='Healthy'?C.green:p.h.l==='At Risk'?C.yellow:p.h.l==='Critical'?C.red:C.slate500;
            sc(ragC);pdf.circle(M+3,ty+rh/2,2,'F');
            sc(p.h.c);pdf.rect(M,ty,1.5,rh,'F');
            pdf.setFontSize(5.5);tc(C.slate900);var rx=M+2;
            pdf.text(String(idx+1),rx+1,ty+5.5);rx+=cs[0].w;
            pdf.text(p.name.substring(0,30),rx,ty+5.5);rx+=cs[1].w;
            pdf.text(String(p.total),rx+2,ty+5.5);rx+=cs[2].w;
            pdf.text(p.done+'/'+p.total,rx,ty+5.5);rx+=cs[3].w;
            var pbW2=cs[4].w-14;sc(C.slate200);pdf.roundedRect(rx,ty+2.5,pbW2,3,1,1,'F');
            var fc2=p.prog>=80?C.green:p.prog>=50?C.blue:p.prog>=25?C.yellow:C.red;
            sc(fc2);pdf.roundedRect(rx,ty+2.5,Math.max(0.5,pbW2*p.prog/100),3,1,1,'F');
            pdf.setFontSize(5);tc(C.slate700);pdf.text(p.prog+'%',rx+pbW2+2,ty+5.5);rx+=cs[4].w;
            tc(p.crit>0?C.red:C.slate500);pdf.setFontSize(5.5);pdf.text(String(p.crit),rx+3,ty+5.5);rx+=cs[5].w;
            tc(p.late>0?C.red:C.green);pdf.text(String(p.late),rx+2,ty+5.5);rx+=cs[6].w;
            tc(C.slate900);pdf.text(cur+p.cost.toLocaleString(),rx,ty+5.5);rx+=cs[7].w;
            sc(p.h.c);pdf.roundedRect(rx,ty+1.5,15,5,2,2,'F');pdf.setFontSize(4.5);tc(C.white);pdf.text(p.h.l,rx+7.5,ty+5,{align:'center'});rx+=cs[8].w;
            tc(p.h.c);pdf.setFontSize(6);pdf.text(String(p.h.s),rx+3,ty+5.5);rx+=cs[9].w;
            tc(C.slate500);pdf.setFontSize(5);pdf.text(fmtDate(p.sd),rx,ty+5.5);rx+=cs[10].w;
            pdf.text(fmtDate(p.fd),rx,ty+5.5);
            ty+=rh;
        });
        addFt();

        // ══════ PAGE 2: MASTER GANTT / SUMMARY VIEW ══════
        pdf.addPage(); addPH('Master Gantt / Summary View');
        var mgy=22;
        pdf.setFontSize(9);tc(C.deepIndigo);pdf.text('Portfolio Master Timeline',M,mgy);mgy+=6;
        var mgX = M+55, mgW = W-M-mgX, mgH = 6;
        sc(C.slate200);pdf.rect(mgX,mgy,mgW,mgH,'F');
        pdf.setFontSize(5);tc(C.slate700);
        for(var q=0;q<=4;q++){var qx=mgX+(q/4)*mgW; pdf.setDrawColor(C.slate300[0],C.slate300[1],C.slate300[2]);pdf.line(qx,mgy,qx,H-M); var dt=new Date(gStart+(q/4)*gRange); pdf.text(fmtDate(dt),qx,mgy+4,{align:'center'});}
        // Today Line on Master Gantt
        var mgTodayX = mgX + ((today.getTime()-gStart)/gRange)*mgW;
        if(mgTodayX>mgX&&mgTodayX<W-M){
            pdf.setDrawColor(C.red[0],C.red[1],C.red[2]);pdf.setLineWidth(0.5);
            pdf.setLineDashPattern([2,1.5]);pdf.line(mgTodayX,mgy,mgTodayX,H-M);
            pdf.setLineDashPattern([]);pdf.setFontSize(5);tc(C.red);pdf.text('Today',mgTodayX,mgy-1,{align:'center'});
        }
        pdf.setLineWidth(0.2);
        mgy+=mgH+2;
        pData.forEach(function(p,idx){
            if(mgy+10>H-16){addFt();pdf.addPage();addPH('Master Gantt (cont.)');mgy=22;}
            if(idx%2===0){sc(C.slate50);pdf.rect(M,mgy,W-M*2,8,'F');}
            // RAG dot + project name
            var mgRag=p.h.l==='Complete'||p.h.l==='Healthy'?C.green:p.h.l==='At Risk'?C.yellow:p.h.l==='Critical'?C.red:C.slate500;
            sc(mgRag);pdf.circle(M+4,mgy+4,1.5,'F');
            pdf.setFontSize(6);tc(C.slate900);pdf.text(p.name.substring(0,32),M+8,mgy+5);
            var ps = p.sd?new Date(p.sd).getTime():gStart, pf = p.fd?new Date(p.fd).getTime():gEnd;
            var bx = mgX + ((ps-gStart)/gRange)*mgW, bw = Math.max(2, ((pf-ps)/gRange)*mgW);
            if(bx<mgX){bw-=(mgX-bx);bx=mgX;} if(bx+bw>W-M)bw=(W-M)-bx;
            if(bw>0){
                sc(p.h.c); pdf.roundedRect(bx,mgy+1.5,bw,5,1,1,'F');
                if(p.prog>0&&p.prog<100){sc(C.slate200); pdf.roundedRect(bx,mgy+1.5,Math.max(1,bw*(100-p.prog)/100),5,1,1,'F'); sc(C.deepIndigo); pdf.roundedRect(bx,mgy+1.5,Math.max(1,bw*p.prog/100),5,1,1,'F');}
                pdf.setFontSize(4);tc(C.white);pdf.text(p.prog+'%',bx+bw/2,mgy+5,{align:'center'});
            }
            mgy+=8;
        });
        // Legend for Master Gantt
        var mgLY=mgy+2;
        sc(C.green);pdf.circle(M,mgLY+1,1.5,'F');pdf.setFontSize(5);tc(C.slate700);pdf.text('Healthy',M+4,mgLY+2);
        sc(C.yellow);pdf.circle(M+25,mgLY+1,1.5,'F');pdf.text('At Risk',M+29,mgLY+2);
        sc(C.red);pdf.circle(M+50,mgLY+1,1.5,'F');pdf.text('Critical',M+54,mgLY+2);
        sc(C.indigo);pdf.rect(M+80,mgLY,6,2,'F');pdf.text('Progress',M+88,mgLY+2);
        sc(C.slate200);pdf.rect(M+108,mgLY,6,2,'F');pdf.text('Remaining',M+116,mgLY+2);
        addFt();

        // ══════ PAGE 3: PORTFOLIO & PROJECT EVM & S-CURVES ══════
        pdf.addPage(); addPH('Earned Value Management (EVM) Dashboard');
        var evY = 22;
        pdf.setFontSize(9);tc(C.deepIndigo);pdf.text('1. Portfolio Overall Performance',M,evY);evY+=6;
        var pvd = tTasks ? (evmTPV/tTasks).toFixed(1) : 0, evd = tTasks ? (evmTEV/tTasks).toFixed(1) : 0;
        var kpX = M;
        [{l:'Planned Value',v:pvd+'%',c:C.indigo},{l:'Earned Value',v:evd+'%',c:C.blue},{l:'SPI',v:spi,c:spi>=1?C.green:C.red}].forEach(function(k){
            sc(C.slate50);pdf.roundedRect(kpX,evY,28,12,1.5,1.5,'F');sc(k.c);pdf.rect(kpX,evY,28,1.5,'F');
            pdf.setFontSize(4.5);tc(C.slate500);pdf.text(k.l,kpX+14,evY+5,{align:'center'});pdf.setFontSize(8);tc(k.c);pdf.text(k.v,kpX+14,evY+10,{align:'center'}); kpX+=30;
        });
        var pScX = kpX + 10, pScW = W-M-pScX, pScH = 30;
        sc(C.slate50); pdf.rect(pScX,evY,pScW,pScH,'F');
        pdf.setFontSize(4);tc(C.slate500);pdf.setDrawColor(C.slate200[0],C.slate200[1],C.slate200[2]);
        for(var ly=0;ly<=4;ly++){var lyy=evY+pScH-(ly/4)*pScH; pdf.line(pScX,lyy,pScX+pScW,lyy); pdf.text((ly*25)+'%',pScX+2,lyy-1.5);}
        for(var lx=0;lx<=4;lx++){var lxx=pScX+(lx/4)*pScW; pdf.line(lxx,evY,lxx,evY+pScH); var ddt=new Date(gStart+(lx/4)*gRange); pdf.text(fmtDate(ddt),lxx+1,evY+pScH-1.5);}
        var pPrev=null, ePrev=null;
        sCurve.forEach(function(pt){
            var cx = pScX + ((pt.x-gStart)/gRange)*pScW, py = evY + pScH - (pt.pv/100)*pScH;
            if(pPrev){ pdf.setDrawColor(C.indigo[0],C.indigo[1],C.indigo[2]); pdf.setLineWidth(0.6); pdf.line(pPrev.x,pPrev.y,cx,py); }
            pPrev = {x:cx,y:py};
            if(pt.ev !== null){
                var ey = evY + pScH - (pt.ev/100)*pScH;
                if(ePrev){ pdf.setDrawColor(C.green[0],C.green[1],C.green[2]); pdf.setLineWidth(0.9); pdf.line(ePrev.x,ePrev.y,cx,ey); }
                ePrev = {x:cx,y:ey};
            }
        });
        pdf.setLineWidth(0.2);
        // S-Curve Legend
        var scLY = evY + pScH + 4;
        sc(C.indigo);pdf.rect(pScX,scLY,6,2,'F');pdf.setFontSize(5);tc(C.slate700);pdf.text('Planned Value (PV)',pScX+8,scLY+2);
        sc(C.green);pdf.rect(pScX+52,scLY,6,2,'F');pdf.text('Earned Value (EV)',pScX+60,scLY+2);
        // SPI Interpretation
        var spiVal = parseFloat(spi);
        var spiIcon = spiVal>=1?'\u2705':spiVal>=0.9?'\u26A0':'\u274C';
        var spiMsg = spiVal>=1?'Portfolio is on or ahead of schedule.':spiVal>=0.9?'Portfolio has slight schedule slippage.':'Portfolio has significant schedule delay — corrective action needed.';
        sc(spiVal>=1?[240,253,244]:spiVal>=0.9?[254,252,232]:[254,242,242]);
        pdf.roundedRect(M,scLY+5,W/3,8,2,2,'F');
        pdf.setFontSize(6);tc(spiVal>=1?C.green:spiVal>=0.9?C.yellow:C.red);
        pdf.text(spiIcon+' '+spiMsg,M+3,scLY+10);
        evY = scLY + 18;
        pdf.setDrawColor(C.slate200[0],C.slate200[1],C.slate200[2]); pdf.line(M,evY,W-M,evY); evY+=6;

        pdf.setFontSize(9);tc(C.deepIndigo);pdf.text('2. EVM by Project',M,evY);evY+=6;
        pData.forEach(function(p, i){
            if(evY+25>H-15){addFt();pdf.addPage();addPH('EVM by Project (cont.)');evY=22;}
            pdf.setFontSize(6.5);tc(C.slate900);pdf.text((i+1)+'. '+p.name.substring(0,30), M, evY+4);
            var pvR = (p.total?p.ptpv/p.total:0).toFixed(1), evR = (p.total?p.ptev/p.total:0).toFixed(1);
            var kpX2 = M;
            [{l:'PV',v:pvR+'%',c:C.indigo},{l:'EV',v:evR+'%',c:C.blue},{l:'SPI',v:p.pspi,c:p.pspi>=1?C.green:C.red}].forEach(function(k){
                sc(C.slate50);pdf.roundedRect(kpX2,evY+6,18,9,1,1,'F');
                pdf.setFontSize(4);tc(C.slate500);pdf.text(k.l,kpX2+9,evY+9.5,{align:'center'});pdf.setFontSize(6);tc(k.c);pdf.text(k.v,kpX2+9,evY+13.5,{align:'center'}); kpX2+=20;
            });
            var scX2 = M+70, scW2 = W-M-scX2, scH2 = 18;
            sc(C.slate50); pdf.rect(scX2,evY,scW2,scH2,'F');
            var pgS = p.sd?new Date(p.sd).getTime():gStart, pgE = p.fd?new Date(p.fd).getTime():gEnd, pRng = Math.max(864e5, pgE-pgS);
            pdf.setFontSize(4);tc(C.slate500);pdf.setDrawColor(C.slate200[0],C.slate200[1],C.slate200[2]);
            for(var ly=0;ly<=2;ly++){var lyy=evY+scH2-(ly/2)*scH2; pdf.line(scX2,lyy,scX2+scW2,lyy); pdf.text((ly*50)+'%',scX2+1,lyy-1);}
            for(var lx=0;lx<=4;lx++){var lxx=scX2+(lx/4)*scW2; pdf.line(lxx,evY,lxx,evY+scH2);}
            
            var ppPrev=null, pePrev=null;
            if(p.pSCurve){
                p.pSCurve.forEach(function(pt){
                    var cx = scX2 + ((pt.x-pgS)/pRng)*scW2, py = evY + scH2 - (pt.pv/100)*scH2;
                    if(ppPrev){ pdf.setDrawColor(C.indigo[0],C.indigo[1],C.indigo[2]); pdf.setLineWidth(0.4); pdf.line(ppPrev.x,ppPrev.y,cx,py); }
                    ppPrev = {x:cx,y:py};
                    if(pt.ev !== null){
                        var ey = evY + scH2 - (pt.ev/100)*scH2;
                        if(pePrev){ pdf.setDrawColor(C.green[0],C.green[1],C.green[2]); pdf.setLineWidth(0.6); pdf.line(pePrev.x,pePrev.y,cx,ey); }
                        pePrev = {x:cx,y:ey};
                    }
                });
            }
            pdf.setLineWidth(0.2);
            evY += 22;
        });
        addFt();

        // ══════ PAGE 5: NEEDS ATTENTION & UPCOMING ══════
        pdf.addPage(); addPH('Needs Attention & Upcoming Tasks');
        var naY = 23;
        // Alert banner
        if(allLate.length>0){
            sc([254,226,226]);pdf.roundedRect(M,naY,W-M*2,8,2,2,'F');
            sc(C.red);pdf.roundedRect(M,naY,3,8,1,1,'F');
            pdf.setFont('helvetica','bold');pdf.setFontSize(7);tc(C.red);
            pdf.text('\u26A0  '+allLate.length+' overdue tasks detected  \u00B7  '+crCount+' critical tasks  \u00B7  Immediate attention required',M+6,naY+5.5);
            pdf.setFont('helvetica','normal');
        } else {
            sc([240,253,244]);pdf.roundedRect(M,naY,W-M*2,8,2,2,'F');
            sc(C.green);pdf.roundedRect(M,naY,3,8,1,1,'F');
            pdf.setFontSize(7);tc(C.green);pdf.text('\u2705  All tasks are on schedule. No overdue items.',M+6,naY+5.5);
        }
        naY+=12;
        pdf.setFont('helvetica','bold');pdf.setFontSize(9);tc(C.red);pdf.text('Late & Overdue Tasks',M,naY);pdf.setFont('helvetica','normal');naY+=6;
        if(allLate.length===0){
            pdf.setFontSize(7);tc(C.slate500);pdf.text('No late tasks across the portfolio.',M,naY);naY+=8;
        } else {
            pdf.setFontSize(7);tc(C.slate700);pdf.text(allLate.length+' tasks overdue across the portfolio',M+50,naY-1);
            sc(C.red);pdf.rect(M,naY+2,W-M*2,7,'F');pdf.setFontSize(5.5);tc(C.white);
            pdf.text('Task Name',M+5,naY+7);pdf.text('Project',M+82,naY+7);pdf.text('Assigned To',M+142,naY+7);
            pdf.text('Days Late',M+212,naY+7);pdf.text('% Done',M+242,naY+7);naY+=9;
            allLate.forEach(function(lt,li){
                if(naY+6>H-15){addFt();pdf.addPage();addPH('Late (cont.)');naY=22;}
                if(li%2===0){sc(C.slate50);pdf.rect(M,naY,W-M*2,6,'F');}
                // Severity Stripe
                if(lt.days>30){sc(C.red);pdf.rect(M,naY,2.5,6,'F');}
                else if(lt.days>14){sc(C.yellow);pdf.rect(M,naY,2.5,6,'F');}
                else{sc(C.green);pdf.rect(M,naY,2.5,6,'F');}
                pdf.setFontSize(5.5);tc(C.slate900);pdf.text((lt.name||'').substring(0,45),M+5,naY+4);
                pdf.text((lt.proj||'').substring(0,32),M+82,naY+4);
                tc(C.indigo);pdf.text((lt.res||'Unassigned').substring(0,38),M+142,naY+4);
                tc(lt.days>30?C.red:lt.days>14?C.yellow:C.slate700);pdf.text(lt.days+' days',M+212,naY+4);
                tc(C.slate700);pdf.text(lt.pct+'%',M+242,naY+4);naY+=6;
            });
        }
        naY+=6;
        if(naY+30>H-20){addFt();pdf.addPage();addPH('Upcoming Tasks (Next 14 Days)');naY=22;}
        pdf.setFontSize(9);tc(C.blue);pdf.text('Upcoming Tasks (Next 14 Days)',M,naY);naY+=6;
        if(pUP.length===0){
            pdf.setFontSize(7);tc(C.slate500);pdf.text('No upcoming tasks identified.',M,naY);naY+=8;
        } else {
            sc(C.blue);pdf.rect(M,naY,W-M*2,7,'F');pdf.setFontSize(5.5);tc(C.white);
            pdf.text('Task Name',M+2,naY+5);pdf.text('Project',M+80,naY+5);pdf.text('Assigned To',M+140,naY+5);
            pdf.text('Start Date',M+210,naY+5);naY+=7;
            pUP.sort(function(a,b){return new Date(a.start)-new Date(b.start);}).slice(0,18).forEach(function(ut,ui){
                if(naY+6>H-15){addFt();pdf.addPage();addPH('Upcoming (cont.)');naY=22;}
                if(ui%2===0){sc(C.slate50);pdf.rect(M,naY,W-M*2,6,'F');}
                pdf.setFontSize(5.5);tc(C.slate900);pdf.text((ut.name||'').substring(0,45),M+2,naY+4);
                pdf.text((ut.proj||'').substring(0,35),M+80,naY+4);
                tc(C.indigo);pdf.text((ut.res||'').substring(0,40),M+140,naY+4);
                tc(C.slate700);pdf.text(fmtDate(ut.start),M+210,naY+4);naY+=6;
            });
        }
        addFt();

        // ══════ PAGE 4: RESOURCE ANALYSIS ══════
        if(sharedR.length>0||overR.length>0){
            pdf.addPage();addPH('Resource Analysis \u2014 Cross-Project');var ry=22;
            if(sharedR.length>0){
                pdf.setFontSize(9);tc(C.deepIndigo);pdf.text('Shared Resources ('+sharedR.length+' across multiple projects)',M,ry);ry+=5;
                sc(C.purple);pdf.rect(M,ry,W-M*2,7,'F');pdf.setFontSize(5.5);tc(C.white);
                pdf.text('Resource Name',M+4,ry+5);pdf.text('Assigned Projects',M+60,ry+5);pdf.text('Tasks',M+200,ry+5);pdf.text('Remaining Work',M+225,ry+5);ry+=7;
                sharedR.forEach(function(sr,si){
                    if(ry+6>H-40){addFt();pdf.addPage();addPH('Resource Analysis (cont.)');ry=22;}
                    if(si%2===0){sc(C.slate50);pdf.rect(M,ry,W-M*2,6,'F');}
                    if(sr.projs.length>=3){sc(C.red);pdf.rect(M,ry,2,6,'F');}else{sc(C.yellow);pdf.rect(M,ry,2,6,'F');}
                    pdf.setFontSize(5.5);tc(C.slate900);pdf.text(sr.name.substring(0,28),M+4,ry+4);
                    tc(C.slate700);pdf.text(sr.projs.map(function(n){return n.substring(0,18);}).join(', ').substring(0,68),M+60,ry+4);
                    tc(C.slate900);pdf.text(String(sr.tc),M+205,ry+4);
                    tc(sr.work>60?C.red:C.slate700);pdf.text(Math.round(sr.work)+'d',M+232,ry+4);ry+=6;
                });
                ry+=8;
            }
            if(overR.length>0){
                if(ry+30>H-16){addFt();pdf.addPage();addPH('Resource Analysis (cont.)');ry=22;}
                pdf.setFontSize(9);tc(C.red);pdf.text('Overloaded Resources ('+overR.length+' with >60 remaining workdays)',M,ry);ry+=5;
                overR.sort(function(a,b){return b.work-a.work;});var mxW=overR[0].work;var obW2=W-M*2-60;
                overR.slice(0,15).forEach(function(or){if(ry+8>H-16)return;pdf.setFontSize(5.5);tc(C.slate900);pdf.text(or.name.substring(0,25),M,ry+4);
                    var bw2=(or.work/mxW)*obW2;sc(or.work>120?C.red:C.yellow);pdf.roundedRect(M+55,ry+1,Math.max(2,bw2),4,1.5,1.5,'F');
                    pdf.setFontSize(5);tc(C.slate700);pdf.text(Math.round(or.work)+'d  |  '+or.tc+' tasks  |  '+or.projs.length+' proj',M+55+bw2+3,ry+4);ry+=7;});
            }
            // Recommendations & Variance Commentary
            ry+=8;if(ry+60>H-16){addFt();pdf.addPage();addPH('Recommendations & Variance Analysis');ry=22;}
            pdf.setFont('helvetica','bold');pdf.setFontSize(9);tc(C.deepIndigo);
            pdf.text('Recommendations & Variance Analysis',M,ry);pdf.setFont('helvetica','normal');ry+=6;
            var rcs=[];
            if(overR.length>0)rcs.push({t:'warn',m:''+overR.length+' resources overloaded (>60 workdays remaining). Consider reallocating or adding capacity.'});
            if(sharedR.length>0)rcs.push({t:'info',m:''+sharedR.length+' resources shared across projects. Coordinate scheduling to prevent conflicts.'});
            if(allLate.length>0)rcs.push({t:'crit',m:''+allLate.length+' tasks overdue. Prioritize critical-path late tasks for immediate recovery.'});
            if(crCount>0)rcs.push({t:'warn',m:''+crCount+' critical-path tasks. Any delay directly extends project completion dates.'});
            var arC2=(hCounts['At Risk']||0)+(hCounts.Critical||0);
            if(arC2>0)rcs.push({t:'crit',m:''+arC2+'/'+pData.length+' projects at risk or critical. Executive attention required.'});
            pData.forEach(function(pd){
                var pSpi = parseFloat(pd.pspi);
                if(pSpi<0.85&&pd.prog<90){
                    var topLate = allLate.filter(function(l){return l.proj===pd.name;}).sort(function(a,b){return b.days-a.days;})[0];
                    var varMsg=pd.name+': SPI='+pd.pspi+' (behind schedule). ';
                    if(topLate) varMsg+='Worst: "'+topLate.name.substring(0,28)+'" is '+topLate.days+'d late.';
                    rcs.push({t:'proj',m:varMsg});
                }
            });
            if(rcs.length===0)rcs.push({t:'ok',m:'All projects and resources within normal parameters. No action needed.'});
            // Draw recommendation cards
            rcs.forEach(function(r){
                if(ry+9>H-16){addFt();pdf.addPage();addPH('Recommendations (cont.)');ry=22;}
                var rBgC = r.t==='crit'?[254,226,226]:r.t==='warn'?[254,243,199]:r.t==='proj'?[238,242,255]:r.t==='ok'?[240,253,244]:[241,245,249];
                var rTxC = r.t==='crit'?C.red:r.t==='warn'?C.yellow:r.t==='proj'?C.indigo:r.t==='ok'?C.green:C.slate500;
                var rAccC = r.t==='crit'?C.red:r.t==='warn'?[180,120,0]:r.t==='proj'?C.indigo:r.t==='ok'?C.green:C.slate500;
                var rIcon = r.t==='crit'?'\u26A0':r.t==='warn'?'\u26A0':r.t==='proj'?'\u25B6':r.t==='ok'?'\u2705':'\u2139';
                sc(rBgC);pdf.roundedRect(M,ry,W-M*2,8,1.5,1.5,'F');
                sc(rAccC);pdf.roundedRect(M,ry,3,8,1,1,'F');
                pdf.setFontSize(5.5);tc(rTxC);pdf.text(rIcon+' '+r.m.substring(0,130),M+5,ry+5.3);
                ry+=9.5;
            });
            addFt();
        }

        // ══════ PROJECT DETAIL PAGES (WBS Tree) ══════
        pData.forEach(function(pd,pidx){
            pdf.addPage();var h=pd.h;
            // Header bar with health color
            sc(h.c);pdf.rect(0,0,W,4,'F');
            // Lighter header background
            sc(C.slate50);pdf.rect(0,4,W,16,'F');
            // Project number badge
            sc(h.c);pdf.roundedRect(M,6,7,10,1.5,1.5,'F');
            pdf.setFont('helvetica','bold');pdf.setFontSize(7);tc(C.white);pdf.text(String(pidx+1),M+3.5,12.5,{align:'center'});pdf.setFont('helvetica','normal');
            // Title
            pdf.setFont('helvetica','bold');pdf.setFontSize(11);tc(C.deepIndigo);pdf.text(pd.name.substring(0,50),M+10,12.5);pdf.setFont('helvetica','normal');
            // Health badge
            sc(h.c);pdf.roundedRect(W-M-28,6,26,10,2,2,'F');
            pdf.setFontSize(5.5);tc(C.white);pdf.text(h.l+'  '+h.s+'/100',W-M-15,12.5,{align:'center'});
            // Meta
            pdf.setFontSize(6);tc(C.slate500);pdf.text(fmtDate(pd.sd)+' \u2014 '+fmtDate(pd.fd)+'   \u00B7   Manager: '+pd.mgr,M+10,18.5);
            // Divider
            sc(h.c);pdf.rect(M,20.5,W-M*2,0.6,'F');

            var mky2=23;var mks2=[
                {l:'Progress',  v:pd.prog+'%',               c:C.indigo},
                {l:'Tasks',     v:pd.done+'/'+pd.total,       c:C.blue},
                {l:'SPI',       v:String(pd.pspi),            c:parseFloat(pd.pspi)>=1?C.green:C.red},
                {l:'Critical',  v:String(pd.crit),            c:pd.crit>0?C.red:C.slate500},
                {l:'Late',      v:String(pd.late),            c:pd.late>0?C.red:C.green},
                {l:'Budget',    v:cur+pd.cost.toLocaleString(),c:C.purple},
            ];
            var mkw2=(W-M*2-5*4)/6;
            mks2.forEach(function(mk,mi){
                var mkx=M+mi*(mkw2+4);
                sc(C.slate50);pdf.roundedRect(mkx,mky2,mkw2,16,2,2,'F');
                sc(mk.c);pdf.roundedRect(mkx,mky2,mkw2,2.5,2,2,'F');pdf.rect(mkx,mky2+1.5,mkw2,1,'F');
                pdf.setFontSize(4.5);tc(C.slate500);pdf.text(mk.l,mkx+mkw2/2,mky2+7,{align:'center'});
                pdf.setFont('helvetica','bold');pdf.setFontSize(8.5);tc(mk.c);
                pdf.text(mk.v.length>8?mk.v.substring(0,7):mk.v,mkx+mkw2/2,mky2+13.5,{align:'center'});
                pdf.setFont('helvetica','normal');
            });

            var pby2=mky2+19;
            // Progress bar with label
            pdf.setFontSize(5);tc(C.slate500);pdf.text('Overall Progress',M,pby2-1);
            pdf.text(pd.prog+'%',W-M,pby2-1,{align:'right'});
            sc(C.slate200);pdf.roundedRect(M,pby2,W-M*2,3.5,1.5,1.5,'F');
            var pfC2=pd.prog>=80?C.green:pd.prog>=50?C.indigo:C.yellow;
            sc(pfC2);pdf.roundedRect(M,pby2,Math.max(1,(W-M*2)*pd.prog/100),3.5,1.5,1.5,'F');

            // WBS Table & Mini Gantt
            var conciseTasks = pd.allT.filter(function(t){ return (t.summary && (t.outlineLevel||1)<=2) || t.critical; });
            var tty=pby2+8;pdf.setFontSize(8);tc(C.deepIndigo);pdf.text('Work Breakdown Structure & Mini Gantt ('+conciseTasks.length+' Summary & Critical Tasks)',M,tty);tty+=4;
            function drawTH(){sc(C.deepIndigo);pdf.rect(M,tty,W-M*2,7,'F');pdf.setFontSize(5);tc(C.white);
                pdf.text('WBS',M+2,tty+5);pdf.text('Task Name',M+14,tty+5);pdf.text('Dur',M+75,tty+5);
                pdf.text('Start',M+85,tty+5);pdf.text('Finish',M+103,tty+5);
                pdf.text('Timeline (Mini Gantt)',M+120,tty+5);tty+=7;}
            drawTH();

            var ppSD=pd.sd?new Date(pd.sd).getTime():0, ppFD=pd.fd?new Date(pd.fd).getTime():0;
            if(!ppSD||!ppFD){pd.allT.forEach(function(t){var s=new Date(t.start).getTime(),f=new Date(t.finish).getTime();if(!ppSD||s<ppSD)ppSD=s;if(!ppFD||f>ppFD)ppFD=f;});}
            var gRange=Math.max(86400000, ppFD-ppSD), gW=W-M*2-120;
            
            var rH2=5.5;
            conciseTasks.forEach(function(t,ti){
                if(tty+rH2>H-14){addFt();pdf.addPage();sc(h.c);pdf.rect(0,0,W,2,'F');pdf.setFontSize(8);tc(C.slate500);pdf.text(pd.name+' \u2014 WBS (continued)',M,10);tty=14;drawTH();}
                var indent=((t.outlineLevel||1)-1);var isSm=t.summary;var pct=t.percentComplete||0;
                var isLt=new Date(t.finish)<today&&pct<100&&!isSm;
                
                if(t.critical&&!isSm){sc([254,226,226]);pdf.rect(M,tty,W-M*2,rH2,'F');} // Light red background for critical
                else if(isSm){sc(C.slate100);pdf.rect(M,tty,W-M*2,rH2,'F');}else if(ti%2===0){sc([253,253,253]);pdf.rect(M,tty,W-M*2,rH2,'F');}
                if(t.critical&&!isSm){sc(C.red);pdf.rect(M,tty,1.5,rH2,'F');}
                if(isLt){sc(C.yellow);pdf.rect(M+1.5,tty,1,rH2,'F');}
                pdf.setFontSize(isSm?5.5:5);tc(t.critical&&!isSm?C.red:(isSm?C.deepIndigo:C.slate900));
                pdf.text((t.wbs||String(ti+1)).substring(0,8),M+2,tty+4);
                var nX=M+14+indent*3;var pfx=isSm?'\u25B8 ':'';var tn=pfx+(t.name||'');
                var mNL=Math.max(10,Math.round((70-indent*3)/1.5));
                pdf.text(tn.substring(0,mNL),nX,tty+4);
                tc(C.slate700);pdf.setFontSize(5);
                pdf.text(isSm?'':((t.durationDays||t.duration||0)+'d'),M+75,tty+4);
                pdf.text(fmtDate(t.start),M+85,tty+4);pdf.text(fmtDate(t.finish),M+103,tty+4);
                
                var tS=new Date(t.start).getTime(), tF=new Date(t.finish).getTime();
                var bx=M+120+((tS-ppSD)/gRange)*gW, bw=Math.max(1.5, ((tF-tS)/gRange)*gW);
                if(bx<M+120){bw-=(M+120-bx);bx=M+120;} if(bx+bw>W-M)bw=(W-M)-bx;
                if(bw>0){
                    var mainColor = t.critical?C.red:(pct>=100?C.green:C.blue);
                    if(isSm){
                        sc(C.slate700);
                        pdf.roundedRect(bx,tty+2,bw,1.5,0.5,0.5,'F');
                        pdf.triangle(bx,tty+3.5, bx+1.5,tty+3.5, bx,tty+4.5, 'F');
                        pdf.triangle(bx+bw,tty+3.5, bx+bw-1.5,tty+3.5, bx+bw,tty+4.5, 'F');
                    } else {
                        sc(C.slate200);
                        pdf.roundedRect(bx,tty+1.5,bw,2.5,0.5,0.5,'F');
                        if(pct>0){
                            sc(mainColor);
                            pdf.roundedRect(bx,tty+1.5,Math.max(0.5,bw*pct/100),2.5,0.5,0.5,'F');
                        }
                    }
                }
                tty+=rH2;
            });
            addFt();
        });

        // Page numbering finalize
        var tp=pdf.internal.getNumberOfPages();
        for(var pp=1;pp<=tp;pp++){pdf.setPage(pp);pdf.setFontSize(5.5);tc(C.slate300);pdf.text(pp+' / '+tp,W/2,H-6,{align:'center'});}
        pdf.save('Portfolio_Report_'+new Date().toISOString().split('T')[0]+'.pdf');
    }


    export const Reports = {
        generateSummary, copySummary, exportCSV, exportExcel, downloadExcel,
        exportGanttPNG, generatePDF, printProject, generatePortfolioPDF
    };
