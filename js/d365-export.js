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

    /** dd/mm/yyyy for the Project Operations WBS template (matches the
     *  default UK/Europe locale that D365's exported template uses). */
    function ddmmyyyy(d) {
        if (!d) return '';
        const date = (d instanceof Date) ? d : new Date(d);
        if (isNaN(date.getTime())) return '';
        const dd = String(date.getUTCDate()).padStart(2, '0');
        const mm = String(date.getUTCMonth() + 1).padStart(2, '0');
        const yyyy = date.getUTCFullYear();
        return `${dd}/${mm}/${yyyy}`;
    }

    /**
     * Return a real Date object for cell values — stored in the XLSX as a
     * proper Excel date serial (what D365 needs for Edm.DateTimeOffset).
     * Falls back to '' when the input is empty/invalid so the AoA still
     * renders cleanly for blank rows.
     */
    function toExcelDate(d) {
        if (!d) return '';
        const date = (d instanceof Date) ? d : new Date(d);
        if (isNaN(date.getTime())) return '';
        // Normalize to a UTC-midnight Date so Excel's epoch conversion is stable
        return new Date(Date.UTC(
            date.getUTCFullYear(),
            date.getUTCMonth(),
            date.getUTCDate(),
            0, 0, 0, 0
        ));
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
    // Schema 1: Project Operations — "Work breakdown structure" template
    // Matches the native Excel template that D365 Project Operations
    // exports via "Export to Excel" from the project's WBS page.
    //
    // Layout (Sheet1):
    //   A1:  "Work breakdown structure"    (title, bold)
    //   A4:  "Project ID"     B4: <id>
    //   A5:  "Project name"   B5: <name>
    //   Row 7: header row (WBS ID, Task name, Predecessors, Category,
    //                      Effort in hours, Start date, End date,
    //                      Duration, Number of resources, Role)
    //   Row 8+: one row per task
    //
    // A second sheet "data_cache" holds the allowed Category / Role
    // lookup values used for the dropdowns in D365.
    // ───────────────────────────────────────────────────────────────────

    /**
     * Duration cell for the WBS template.
     * The official D365 template stores Duration as a STRING like "186 days"
     * (shared string), e.g. sharedStrings entries "1 day", "2 days", "186 days".
     * D365's importer reads Duration as Edm.String; emitting a raw number
     * caused the IEEE754Compatible/Decimal conflict we saw earlier.
     */
    function durationDays(durationDays) {
        const d = Number(durationDays);
        if (!isFinite(d) || d < 0) return '0 days';
        const n = Math.round(d);
        return n === 1 ? '1 day' : `${n} days`;
    }

    /**
     * Resolve the role / category / count / effort fields from the task.
     * Falls back sensibly when fields aren't present on the ProjectFlow task.
     */
    function pickTaskExtras(t, project) {
        const assignments = Array.isArray(t.assignments) ? t.assignments : [];
        // Number of resources = distinct resource IDs
        let nRes = assignments.length;
        if (!nRes && t.resources) {
            nRes = Array.isArray(t.resources) ? t.resources.length
                 : String(t.resources).split(/[,;]/).filter(Boolean).length;
        }
        // Role — take first assignment's role, or explicit t.role
        let role = t.role || '';
        if (!role && assignments.length && project && Array.isArray(project.resources)) {
            const first = assignments[0];
            const r = project.resources.find(r => r.id === first.resourceId || r.id === first.resId);
            if (r) role = r.role || r.group || r.name || '';
        }
        // Category — default "Other" (matches D365 msdyn_transactioncategory default)
        const category = t.category || t.transactionCategory || 'Other';
        // Effort in hours — prefer explicit effortHours, else cost/8
        let effort = 0;
        if (isFinite(Number(t.effortHours))) effort = Number(t.effortHours);
        else if (isFinite(Number(t.work))) effort = Number(t.work);
        else effort = costToEffortHours(t.cost);
        return {
            nRes,
            role: role || '',
            category,
            effort: Number(effort.toFixed(2))
        };
    }

    /**
     * Build the D365 Project Operations WBS import workbook (matches the
     * official template the customer's Dynamics tenant exports).
     *
     * @param {object}  project   ProjectFlow project
     * @param {object}  [opts]
     * @param {string}  [opts.projId]  Optional explicit D365 Project ID (e.g. "PAOM-000361")
     */
    function buildOperationsWorkbook(project, opts) {
        if (typeof XLSX === 'undefined') {
            throw new Error('SheetJS (XLSX) is not loaded.');
        }
        if (!project) throw new Error('No project provided.');

        const wb = XLSX.utils.book_new();
        const tasks = Array.isArray(project.tasks) ? project.tasks : [];
        const wbsMap = buildWBSIndex(tasks);
        const taskById = new Map();
        tasks.forEach(t => { if (t.id != null) taskById.set(t.id, t); });

        const projId = cleanCell(
            (opts && opts.projId) || project.projId || project.d365ProjectId || ''
        );
        const projName = cleanCell(project.name || 'Untitled Project');

        // ─── Sheet 1: "Sheet1" — WBS template ──────────────────────
        // Row layout matches the official D365 Project Operations template:
        //   Row 1: blank
        //   Row 2: "Work breakdown structure" (title, merged A2:J2)
        //   Row 3-4: blank
        //   Row 5: Project ID  | <projId>
        //   Row 6: Project name| <projName>
        //   Row 7: blank
        //   Row 8: table header
        //   Row 9+: data rows
        const aoa = [];
        aoa.push([]);                                        // Row 1
        aoa.push(['Work breakdown structure']);              // Row 2
        aoa.push([]);                                        // Row 3
        aoa.push([]);                                        // Row 4
        aoa.push(['Project ID',   projId]);                  // Row 5
        aoa.push(['Project name', projName]);                // Row 6
        aoa.push([]);                                        // Row 7
        // Row 8: table header
        const tableHeader = [
            'WBS ID',
            'Task name',
            'Predecessors',
            'Category',
            'Effort in hours',
            'Start date',
            'End date',
            'Duration',
            'Number of resources',
            'Role'
        ];
        aoa.push(tableHeader);

        // Rows 8+: one row per task (preserves input order so WBS hierarchy is intact)
        tasks.forEach(t => {
            const wbs = wbsMap.get(t.uid) || '';
            const extras = pickTaskExtras(t, project);
            aoa.push([
                cleanCell(wbs),
                cleanCell(t.name || 'Unnamed Task'),
                cleanCell(formatPredecessors(t.predecessors, taskById, wbsMap)),
                cleanCell(extras.category),
                Number(extras.effort) || 0,        // Edm.Decimal
                toExcelDate(t.start),              // Edm.DateTimeOffset (see post-pass)
                toExcelDate(t.finish),             // Edm.DateTimeOffset
                durationDays(t.durationDays),      // Edm.Decimal (days)
                Number(extras.nRes) || 0,          // Edm.Int32
                cleanCell(extras.role)
            ]);
        });

        // cellDates:true tells SheetJS to serialize Date objects in our AoA as
        // real Excel dates (type 'd') rather than coercing them to strings.
        // dateNF applies the dd/mm/yyyy display format D365 expects.
        const ws1 = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true, dateNF: 'dd/mm/yyyy' });

        // Post-pass: force every cell in the Start date / End date columns
        // to be numeric-date (t:'n' with the serial code) + format dd/mm/yyyy.
        // SheetJS sometimes emits ISO strings (t:'d') which D365 rejects —
        // converting to the number representation is the portable fix.
        const START_COL = 5; // column F (0-based)
        const END_COL   = 6; // column G
        const HEADER_ROW = 6; // 0-based → Excel row 7
        for (let r = HEADER_ROW + 1; r < aoa.length; r++) {
            [START_COL, END_COL].forEach(c => {
                const addr = XLSX.utils.encode_cell({ r, c });
                const cell = ws1[addr];
                if (!cell) return;
                if (cell.v instanceof Date) {
                    // Excel date serial: days since 1899-12-30 (Lotus 1-2-3 quirk)
                    const EPOCH = Date.UTC(1899, 11, 30);
                    const MS_PER_DAY = 86400000;
                    cell.v = (cell.v.getTime() - EPOCH) / MS_PER_DAY;
                    cell.t = 'n';
                    cell.z = 'dd/mm/yyyy';
                } else if (typeof cell.v === 'string' && cell.v) {
                    // Last-resort string → date (dd/mm/yyyy parsed as UK)
                    const m = cell.v.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
                    if (m) {
                        const dt = Date.UTC(+m[3], +m[2] - 1, +m[1]);
                        const EPOCH = Date.UTC(1899, 11, 30);
                        cell.v = (dt - EPOCH) / 86400000;
                        cell.t = 'n';
                        cell.z = 'dd/mm/yyyy';
                    }
                }
            });
        }

        // Column widths tuned to mirror D365's template
        ws1['!cols'] = [
            {wch:12}, // WBS ID
            {wch:40}, // Task name
            {wch:18}, // Predecessors
            {wch:18}, // Category
            {wch:15}, // Effort in hours
            {wch:12}, // Start date
            {wch:12}, // End date
            {wch:12}, // Duration
            {wch:20}, // Number of resources
            {wch:22}  // Role
        ];

        // Style: bold title + bold header row + bold label cells
        // (SheetJS Community Edition honours these style hints in most viewers)
        const boldStyle   = { font: { bold: true, sz: 14 } };
        const labelStyle  = { font: { bold: true }, fill: { fgColor: { rgb: 'DEEBF7' } } };
        const headerStyle = { font: { bold: true }, fill: { fgColor: { rgb: 'DEEBF7' } } };

        if (ws1['A1']) ws1['A1'].s = boldStyle;
        ['A4', 'A5'].forEach(addr => { if (ws1[addr]) ws1[addr].s = labelStyle; });
        // Header row (row index 7 → Excel row number 7 → 0-based col index)
        for (let c = 0; c < tableHeader.length; c++) {
            const addr = XLSX.utils.encode_cell({ r: 6, c });
            if (ws1[addr]) ws1[addr].s = headerStyle;
        }

        // Merge the title across the full width of the table
        ws1['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: tableHeader.length - 1 } }
        ];

        // Auto-filter on the header row so it behaves like D365's template
        const lastCol = XLSX.utils.encode_col(tableHeader.length - 1);
        ws1['!autofilter'] = { ref: `A7:${lastCol}${aoa.length}` };

        XLSX.utils.book_append_sheet(wb, ws1, 'Sheet1');

        // ─── Sheet 2: data_cache (dropdown lookup values) ──────────
        // D365's template uses this hidden sheet to back data-validation
        // dropdowns. Populating it with safe defaults lets the uploaded
        // file open cleanly in Excel/D365 even when the user never touches
        // it.  Users can edit the values to match their Dynamics tenant.
        const categories = [
            'Other',
            'Consulting',
            'Development',
            'Design',
            'Testing',
            'Training',
            'Documentation',
            'Project management',
            'Support',
            'Installation',
            'Analysis'
        ];
        const roles = [
            'Project Manager',
            'Business Analyst',
            'Solution Architect',
            'Developer',
            'Senior Developer',
            'Tester',
            'QA Engineer',
            'Consultant',
            'Technical Lead',
            'Designer',
            'Trainer',
            'Support Engineer'
        ];
        // Build a 2-column sheet: Category | Role
        const cache = [['Category', 'Role']];
        const nRows = Math.max(categories.length, roles.length);
        for (let i = 0; i < nRows; i++) {
            cache.push([categories[i] || '', roles[i] || '']);
        }
        const ws2 = XLSX.utils.aoa_to_sheet(cache);
        ws2['!cols'] = [{ wch: 24 }, { wch: 24 }];
        XLSX.utils.book_append_sheet(wb, ws2, 'data_cache');

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

    /** Export → D365 Project Operations (WBS template) */
    function exportOperations(project, opts) {
        const wb = buildOperationsWorkbook(project, opts || {});
        _download(wb, _filename(project, 'D365-WBS'));
        return true;
    }

    /** Export → D365 Finance & Operations (Project Accounting legacy) */
    function exportFinanceOps(project) {
        const wb = buildFinanceOpsWorkbook(project);
        _download(wb, _filename(project, 'D365-FinanceOps'));
        return true;
    }

    /** Export both flavours (two sequential downloads) */
    function exportBoth(project, opts) {
        exportOperations(project, opts);
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
