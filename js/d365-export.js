/**
 * ═══════════════════════════════════════════════════════════════════════
 * ProjectFlow™ — D365 Export Module
 * © 2026 Ahmed M. Fawzy
 *
 * Generates Excel (XLSX) files that match the native import schemas of:
 *   1. Dynamics 365 Project Operations (Dataverse — msdyn_ entities)
 *   2. Dynamics 365 Finance & Operations — Project Accounting (ProjTable / ProjActivity)
 *
 * Files can be uploaded directly to:
 *   • Power Platform → Data → Import (for Project Operations)
 *   • Data Management Framework → Import (for F&O / Project Accounting)
 *
 * Usage:
 *   D365Export.exportOperations(project);   // Project Operations .xlsx
 *   D365Export.exportFinanceOps(project);   // Finance & Operations .xlsx
 *   D365Export.exportBoth(project);         // both files, zipped-like sequential download
 * ═══════════════════════════════════════════════════════════════════════
 */

// ───────────────────────────────────────────────────────────────────
// Helpers
// ───────────────────────────────────────────────────────────────────

    /** Escape a value for safe inclusion in Excel (prevents formula-injection XSS) */
    function cleanCell(v) {
        if (v === null || v === undefined) return '';
        let s = String(v);
        // Prevent CSV/Excel formula injection (=, +, -, @)
        if (/^[=+\-@]/.test(s)) s = "'" + s;
        return s;
    }

    /** ISO-8601 date for D365 imports (both Dataverse and F&O) */
    function isoDate(d) {
        if (!d) return '';
        const date = (d instanceof Date) ? d : new Date(d);
        if (isNaN(date.getTime())) return '';
        return date.toISOString().split('T')[0]; // YYYY-MM-DD
    }

    /** Full ISO datetime (F&O often wants time component) */
    function isoDateTime(d) {
        if (!d) return '';
        const date = (d instanceof Date) ? d : new Date(d);
        if (isNaN(date.getTime())) return '';
        return date.toISOString(); // YYYY-MM-DDTHH:mm:ss.sssZ
    }

    /** Convert duration (days) → minutes for Dataverse msdyn_scheduleddurationminutes */
    function durationToMinutes(durationDays, minutesPerDay) {
        const mpd = Number(minutesPerDay) || 480;
        return Math.max(0, Math.round((Number(durationDays) || 0) * mpd));
    }

    /** Convert cost-hours back to effort hours (reverse of d365.js:329 import logic) */
    function costToEffortHours(cost) {
        return Math.max(0, (Number(cost) || 0) / 8);
    }

    /** Build hierarchical WBS IDs. If task already has wbs, keep it; otherwise generate (1.1.2 style) */
    function buildWBSIndex(tasks) {
        const wbsMap = new Map();
        const counters = []; // counters[level] = current index at that outline level
        (tasks || []).forEach(t => {
            if (t.wbs && /^[\d.]+$/.test(t.wbs)) {
                wbsMap.set(t.uid, t.wbs);
                return;
            }
            const level = Math.max(0, t.outlineLevel || 0);
            // Reset deeper counters when going back up the tree
            counters.length = level + 1;
            counters[level] = (counters[level] || 0) + 1;
            const wbs = counters.slice(0, level + 1).map(c => c || 1).join('.');
            wbsMap.set(t.uid, wbs);
        });
        return wbsMap;
    }

    /** Find parent WBS (drops last segment) */
    function parentWBS(wbs) {
        if (!wbs) return '';
        const parts = String(wbs).split('.');
        if (parts.length <= 1) return '';
        return parts.slice(0, -1).join('.');
    }

    /**
     * Convert predecessor spec (stored in ProjectFlow) to D365 format.
     * ProjectFlow predecessors look like: "3FS+2d,5SS" → D365 wants "WBS-ID:Type:Lag"
     * Returns comma-separated list of WBS IDs (simplified — most verticals accept this).
     */
    function formatPredecessors(predsRaw, taskById, wbsMap) {
        if (!predsRaw) return '';
        // Normalize to array of strings
        const parts = String(predsRaw)
            .split(/[,;]/)
            .map(s => s.trim())
            .filter(Boolean);
        const out = [];
        parts.forEach(p => {
            // match e.g. "3FS+2d", "5SS", "7", "3FS-1d"
            const m = p.match(/^(\d+)\s*(FS|SS|FF|SF)?\s*([+-]?\d+)?\s*([dhw]?)/i);
            if (!m) return;
            const predId = Number(m[1]);
            const type = (m[2] || 'FS').toUpperCase();
            const lag = m[3] || '';
            const unit = (m[4] || 'd').toLowerCase();
            const predTask = taskById.get(predId) || taskById.get(String(predId));
            const predWbs = predTask ? (wbsMap.get(predTask.uid) || predTask.wbs || predId) : predId;
            let piece = `${predWbs}${type}`;
            if (lag) piece += `${lag}${unit}`;
            out.push(piece);
        });
        return out.join(',');
    }

    /** Map ProjectFlow task type → F&O ProjActivity milestone flag */
    function isMilestoneNo(t) {
        if (t.milestone) return 'Yes';
        if (Number(t.durationDays) === 0) return 'Yes';
        return 'No';
    }

    // ───────────────────────────────────────────────────────────────────
    // Schema 1: Project Operations (Dataverse)
    // Upload path: Power Platform → Data Management → Import
    // ───────────────────────────────────────────────────────────────────

    /**
     * Build the workbook for Project Operations import.
     * Two sheets that match the exact Dataverse entity column names.
     */
    function buildOperationsWorkbook(project) {
        if (typeof XLSX === 'undefined') {
            throw new Error('SheetJS (XLSX) is not loaded.');
        }
        if (!project) throw new Error('No project provided.');

        const wb = XLSX.utils.book_new();
        const tasks = Array.isArray(project.tasks) ? project.tasks : [];
        const wbsMap = buildWBSIndex(tasks);
        const taskById = new Map();
        tasks.forEach(t => { if (t.id != null) taskById.set(t.id, t); });

        // ─── Sheet 1: msdyn_projects ───────────────────────────────
        const projectsSheet = [
            [
                'msdyn_subject',              // Project name (required)
                'msdyn_description',          // Description
                'msdyn_scheduledstart',       // Start (ISO 8601)
                'msdyn_scheduledend',         // Finish (ISO 8601)
                'msdyn_projectmanager',       // Manager user email (lookup)
                'msdyn_customer',             // Customer account name (lookup)
                'msdyn_totalplannedcost',     // Planned cost
                'msdyn_overallprojectstatus', // Status code (192350000=OnTrack)
                'transactioncurrency'         // Currency code (ISO 4217)
            ],
            [
                cleanCell(project.name || 'Untitled Project'),
                cleanCell(project.description || ''),
                isoDateTime(project.startDate),
                isoDateTime(project.finishDate),
                cleanCell(project.managerEmail || project.manager || ''),
                cleanCell(project.customer || project.client || ''),
                Number(project.totalCost || 0),
                192350000, // OnTrack by default
                cleanCell(project.currencyCode || 'USD')
            ]
        ];
        const ws1 = XLSX.utils.aoa_to_sheet(projectsSheet);
        ws1['!cols'] = [
            {wch:30},{wch:40},{wch:12},{wch:12},{wch:25},{wch:25},{wch:15},{wch:22},{wch:10}
        ];
        XLSX.utils.book_append_sheet(wb, ws1, 'msdyn_projects');

        // ─── Sheet 2: msdyn_projecttasks ───────────────────────────
        const tasksHeader = [
            'msdyn_subject',                   // Task name (required)
            'msdyn_project',                   // Parent project name (lookup)
            'msdyn_wbsid',                     // WBS ID (e.g. "1.2.3")
            'msdyn_parenttask_wbsid',          // Parent WBS (empty for roots)
            'msdyn_outlinelevel',              // 0-based level
            'msdyn_scheduledstart',            // Start (ISO 8601)
            'msdyn_scheduledend',              // Finish (ISO 8601)
            'msdyn_scheduleddurationminutes',  // Duration in minutes
            'msdyn_progress',                  // 0-100
            'msdyn_iscritical',                // true/false
            'msdyn_ismilestone',               // true/false
            'msdyn_effort',                    // Effort hours
            'msdyn_description',               // Description
            'msdyn_predecessors_wbsid'         // Comma list of WBS with FS/SS/FF/SF + lag
        ];
        const tasksSheet = [tasksHeader];
        const mpd = project.minutesPerDay || 480;
        const projName = cleanCell(project.name || 'Untitled Project');
        tasks.forEach(t => {
            const wbs = wbsMap.get(t.uid) || '';
            tasksSheet.push([
                cleanCell(t.name || 'Unnamed Task'),
                projName,
                cleanCell(wbs),
                cleanCell(parentWBS(wbs)),
                Number(t.outlineLevel || 0),
                isoDateTime(t.start),
                isoDateTime(t.finish),
                durationToMinutes(t.durationDays, mpd),
                Math.max(0, Math.min(100, Number(t.percentComplete || 0))),
                t.critical ? 'true' : 'false',
                (t.milestone || Number(t.durationDays) === 0) ? 'true' : 'false',
                costToEffortHours(t.cost),
                cleanCell(t.description || ''),
                cleanCell(formatPredecessors(t.predecessors, taskById, wbsMap))
            ]);
        });
        const ws2 = XLSX.utils.aoa_to_sheet(tasksSheet);
        ws2['!cols'] = [
            {wch:35},{wch:25},{wch:10},{wch:10},{wch:8},{wch:12},{wch:12},
            {wch:10},{wch:8},{wch:8},{wch:8},{wch:8},{wch:40},{wch:20}
        ];
        XLSX.utils.book_append_sheet(wb, ws2, 'msdyn_projecttasks');

        // ─── Sheet 3: README (optional help) ──────────────────────
        const readme = [
            ['ProjectFlow → D365 Project Operations Export'],
            [''],
            ['Upload instructions:'],
            ['1. Sign in to Power Platform admin center (make.powerapps.com)'],
            ['2. Open your environment → Solutions → Default Solution'],
            ['3. Tools → Import data → Excel'],
            ['4. Choose this file. The tab names match entity logical names.'],
            ['5. For lookup columns (msdyn_projectmanager, msdyn_customer) D365 will'],
            ['   match on the email / account-name text provided.'],
            [''],
            ['Notes:'],
            ['• Dates are in ISO-8601 UTC.'],
            ['• Duration is in minutes. 480 min = 1 working day (8h).'],
            ['• Status 192350000 = OnTrack, 192350001 = AtRisk, 192350002 = OffTrack.'],
            ['• Predecessors use format "1.2FS+1d,2.1SS".'],
            [''],
            ['Generated: ' + new Date().toISOString()],
            ['Source:   ProjectFlow™ © Ahmed M. Fawzy']
        ];
        const ws3 = XLSX.utils.aoa_to_sheet(readme);
        ws3['!cols'] = [{wch:70}];
        XLSX.utils.book_append_sheet(wb, ws3, 'README');

        return wb;
    }

    // ───────────────────────────────────────────────────────────────────
    // Schema 2: Finance & Operations (Project Accounting, Legacy)
    // Upload path: Data Management → Import → File + Entity selection
    // ───────────────────────────────────────────────────────────────────

    /**
     * Build the workbook for F&O Project Accounting import.
     * Sheet names match the DMF target-entity names.
     */
    function buildFinanceOpsWorkbook(project) {
        if (typeof XLSX === 'undefined') {
            throw new Error('SheetJS (XLSX) is not loaded.');
        }
        if (!project) throw new Error('No project provided.');

        const wb = XLSX.utils.book_new();
        const tasks = Array.isArray(project.tasks) ? project.tasks : [];
        const wbsMap = buildWBSIndex(tasks);
        const taskById = new Map();
        tasks.forEach(t => { if (t.id != null) taskById.set(t.id, t); });

        // Safe ProjId: F&O limits to ~20 chars, alphanumeric + dash
        const safeProjId = (project.projId || project.name || 'PROJ')
            .toUpperCase().replace(/[^A-Z0-9-]/g, '-').slice(0, 20) || 'PROJ';

        // ─── Sheet 1: ProjTable (Project Header) ───────────────────
        const projTableHeader = [
            'PROJID',                  // Primary key *
            'NAME',                    // Project name *
            'PROJECTTYPE',             // TimeMaterial / FixedPrice / Internal / Cost / Investment / Time
            'PROJECTGROUPID',          // Project group (must exist in F&O) — leave editable
            'STATUS',                  // Created / InProcess / Completed / OnHold / Canceled
            'SCHEDULEDSTARTDATE',      // Start (ISO)
            'SCHEDULEDENDDATE',        // Finish (ISO)
            'PROJECTMANAGER',          // Worker personnel number
            'CUSTOMERACCOUNT',         // Customer code
            'DESCRIPTION',             // Description
            'DEFAULTLINEPROPERTY',     // "Chargeable" / "Non-chargeable"
            'CURRENCY'                 // ISO currency
        ];
        const projTableSheet = [
            projTableHeader,
            [
                cleanCell(safeProjId),
                cleanCell(project.name || 'Untitled Project'),
                cleanCell(project.projectType || 'TimeMaterial'),
                cleanCell(project.projectGroup || 'TM'),
                cleanCell(project.status || 'InProcess'),
                isoDate(project.startDate),
                isoDate(project.finishDate),
                cleanCell(project.managerId || project.manager || ''),
                cleanCell(project.customerAccount || project.customer || ''),
                cleanCell(project.description || ''),
                'Chargeable',
                cleanCell(project.currencyCode || 'USD')
            ]
        ];
        const ws1 = XLSX.utils.aoa_to_sheet(projTableSheet);
        ws1['!cols'] = [
            {wch:15},{wch:30},{wch:15},{wch:15},{wch:12},{wch:12},{wch:12},
            {wch:15},{wch:15},{wch:40},{wch:15},{wch:10}
        ];
        XLSX.utils.book_append_sheet(wb, ws1, 'ProjTable');

        // ─── Sheet 2: ProjActivity (WBS) ───────────────────────────
        const actHeader = [
            'ACTIVITYNUMBER',          // WBS code (unique) *
            'PROJID',                  // FK to ProjTable.PROJID *
            'NAME',                    // Activity name *
            'PARENTACTIVITYNUMBER',    // Parent WBS for hierarchy
            'OUTLINELEVEL',            // 0 = root
            'SCHEDULEDSTARTDATE',      // Start date
            'SCHEDULEDENDDATE',        // End date
            'EFFORTHOURS',             // Planned hours
            'REMAININGEFFORTHOURS',    // Remaining (effort × (1 - % complete))
            'PERCENTCOMPLETE',         // 0-100
            'ISMILESTONE',             // Yes/No
            'ISCRITICAL',              // Yes/No
            'DESCRIPTION',             // Description
            'DEPENDENCIES'             // Predecessor list (WBS + type + lag)
        ];
        const actSheet = [actHeader];
        tasks.forEach(t => {
            const wbs = wbsMap.get(t.uid) || '';
            const effort = costToEffortHours(t.cost);
            const pct = Math.max(0, Math.min(100, Number(t.percentComplete || 0)));
            actSheet.push([
                cleanCell(wbs),
                cleanCell(safeProjId),
                cleanCell(t.name || 'Unnamed Task'),
                cleanCell(parentWBS(wbs)),
                Number(t.outlineLevel || 0),
                isoDate(t.start),
                isoDate(t.finish),
                Number(effort.toFixed(2)),
                Number((effort * (1 - pct / 100)).toFixed(2)),
                pct,
                isMilestoneNo(t),
                t.critical ? 'Yes' : 'No',
                cleanCell(t.description || ''),
                cleanCell(formatPredecessors(t.predecessors, taskById, wbsMap))
            ]);
        });
        const ws2 = XLSX.utils.aoa_to_sheet(actSheet);
        ws2['!cols'] = [
            {wch:12},{wch:15},{wch:35},{wch:12},{wch:8},{wch:12},{wch:12},
            {wch:10},{wch:12},{wch:8},{wch:8},{wch:8},{wch:40},{wch:20}
        ];
        XLSX.utils.book_append_sheet(wb, ws2, 'ProjActivity');

        // ─── Sheet 3: README ───────────────────────────────────────
        const readme = [
            ['ProjectFlow → D365 Finance & Operations (Project Accounting) Export'],
            [''],
            ['Upload via Data Management Framework (DMF):'],
            ['1. In F&O: Workspaces → Data Management'],
            ['2. "Import" → give job name, select Source = Excel'],
            ['3. Entity "Projects V2" → upload this file, tab "ProjTable"'],
            ['4. Add another record for entity "Project WBS activities" → tab "ProjActivity"'],
            ['5. Validate mapping then "Import".'],
            [''],
            ['Required master data that must exist in F&O BEFORE import:'],
            ['• Project group (PROJECTGROUPID) — default "TM"'],
            ['• Customer account (CUSTOMERACCOUNT) — if provided'],
            ['• Worker personnel number (PROJECTMANAGER) — if provided'],
            ['• Currency code (CURRENCY) — default "USD"'],
            [''],
            ['Column notes:'],
            ['• Dates are ISO-8601 (YYYY-MM-DD).'],
            ['• EFFORTHOURS = hours; F&O converts to lines internally.'],
            ['• DEPENDENCIES format: "1.2FS+1d,2.1SS" (WBS + type + optional lag).'],
            [''],
            ['Generated: ' + new Date().toISOString()],
            ['Source:   ProjectFlow™ © Ahmed M. Fawzy']
        ];
        const ws3 = XLSX.utils.aoa_to_sheet(readme);
        ws3['!cols'] = [{wch:70}];
        XLSX.utils.book_append_sheet(wb, ws3, 'README');

        return wb;
    }

    // ───────────────────────────────────────────────────────────────────
    // Public API
    // ───────────────────────────────────────────────────────────────────

    function _filename(project, suffix) {
        const safe = (project.name || 'project').replace(/[^\w-]+/g, '_').slice(0, 40);
        const ts = new Date().toISOString().slice(0, 10);
        return `${safe}_${suffix}_${ts}.xlsx`;
    }

    function _download(wb, filename) {
        if (typeof XLSX === 'undefined') {
            throw new Error('SheetJS (XLSX) not loaded.');
        }
        XLSX.writeFile(wb, filename);
    }

    /** Export → D365 Project Operations (Dataverse) */
    function exportOperations(project) {
        const wb = buildOperationsWorkbook(project);
        _download(wb, _filename(project, 'D365-ProjectOps'));
        return true;
    }

    /** Export → D365 Finance & Operations (Project Accounting legacy) */
    function exportFinanceOps(project) {
        const wb = buildFinanceOpsWorkbook(project);
        _download(wb, _filename(project, 'D365-FinanceOps'));
        return true;
    }

    /** Export both flavours (two sequential downloads) */
    function exportBoth(project) {
        exportOperations(project);
        // small delay so some browsers don't drop the second download
        setTimeout(() => exportFinanceOps(project), 400);
        return true;
    }

// ESM export (matches the pattern of d365.js, evm.js, etc.)
export const D365Export = {
    exportOperations,
    exportFinanceOps,
    exportBoth,
    // expose internal builders for tests
    _buildOperationsWorkbook: buildOperationsWorkbook,
    _buildFinanceOpsWorkbook: buildFinanceOpsWorkbook,
    _buildWBSIndex: buildWBSIndex,
    _formatPredecessors: formatPredecessors
};

// Also expose on window so console / plugins can reach it
if (typeof window !== 'undefined') window.D365Export = D365Export;
