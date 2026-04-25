/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — MS Planner Excel Parser
 * Converts Microsoft Planner Excel export to ProjectFlow JSON
 * Supports: Metadata skipping, Arabic/English, Dependencies
 * ═══════════════════════════════════════════════════════
 */


    /* ─── Column Aliases (EN + AR) ─── */
    const ALIASES = {
        taskNum:  ['Task number', 'رقم المهمة', '#'],
        outline:  ['Outline number', 'رقم المخطط'],
        name:     ['Name', 'Task Name', 'اسم المهمة', 'العنوان', 'Subject'],
        bucket:   ['Bucket', 'المجموعة', 'Category', 'الفئة'],
        progress: ['% complete', 'Progress', 'التقدم', 'Status', 'الحالة'],
        start:    ['Start', 'Start Date', 'تاريخ البدء', 'البداية'],
        finish:   ['Finish', 'Due Date', 'تاريخ الاستحقاق', 'النهاية', 'Due'],
        duration: ['Duration', 'المدة'],
        assigned: ['Assigned to', 'Assigned To', 'تم التعيين إلى', 'المسؤول'],
        depends:  ['Depends on', 'Dependencies', 'الاعتمادات', 'Predecessors'],
        milestone:['Milestone', 'مهمة حاسمة'],
        notes:    ['Notes', 'Description', 'الوصف', 'ملاحظات'],
    };

    /* ─── All known header keywords (for row detection) ─── */
    const HEADER_KEYWORDS = [];
    Object.values(ALIASES).forEach(arr => arr.forEach(a => HEADER_KEYWORDS.push(a.toLowerCase())));

    /**
     * Main entry point — parse Excel file
     */
    async function parse(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    if (typeof XLSX === 'undefined') throw new Error('Excel library (XLSX) not loaded. Refresh page.');

                    const wb  = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
                    const ws  = wb.Sheets[wb.SheetNames[0]];
                    const all = XLSX.utils.sheet_to_json(ws, { header: 1 }); // 2-D array

                    /* ── Find the REAL header row ── */
                    const hIdx = findHeaderRow(all);
                    if (hIdx === -1) throw new Error('Cannot find data headers. Expected columns like "Name", "Start", "Finish".');

                    /* ── Extract metadata from rows BEFORE the header ── */
                    const meta = extractMeta(all, hIdx);

                    /* ── Parse data rows using detected header ── */
                    const dataRows = XLSX.utils.sheet_to_json(ws, { range: hIdx, defval: '' });
                    if (!dataRows.length) throw new Error('No task rows found after headers.');

                    const project = buildProject(dataRows, meta);
                    resolve(project);
                } catch (err) {
                    reject(err);
                }
            };
            reader.onerror = () => reject(new Error('File read error.'));
            reader.readAsArrayBuffer(file);
        });
    }

    /**
     * Find header row by requiring AT LEAST 3 recognised column names in the same row.
     * This prevents false-positives on metadata rows like "Project name	TLD…"
     */
    function findHeaderRow(allRows) {
        for (let i = 0; i < Math.min(allRows.length, 30); i++) {
            const row = allRows[i];
            if (!row || !Array.isArray(row)) continue;

            let matches = 0;
            for (const cell of row) {
                if (!cell) continue;
                const lc = String(cell).trim().toLowerCase();
                if (HEADER_KEYWORDS.includes(lc)) matches++;
            }
            // Require at least 3 recognised headers to be sure
            if (matches >= 3) return i;
        }
        return -1;
    }

    /**
     * Extract project-level metadata from rows before the header
     */
    function extractMeta(allRows, headerIdx) {
        const meta = { name: '', owner: '', startDate: null, finishDate: null, pctComplete: 0 };
        for (let i = 0; i < headerIdx; i++) {
            const row = allRows[i];
            if (!row || row.length < 2) continue;
            const key = String(row[0] || '').toLowerCase().trim();
            const val = row[1];
            if (key.includes('project name'))       meta.name = String(val || '');
            else if (key.includes('plan owner'))    meta.owner = String(val || '');
            else if (key.includes('start date'))    meta.startDate = excelDate(val);
            else if (key.includes('finish date'))   meta.finishDate = excelDate(val);
            else if (key.includes('% complete'))    meta.pctComplete = typeof val === 'number' ? Math.round(val * 100) : parseInt(val) || 0;
        }
        return meta;
    }

    /* ── Alias-based cell reader ── */
    function getVal(row, key) {
        const aliases = ALIASES[key];
        if (!aliases) return '';
        const rk = Object.keys(row);
        for (const alias of aliases) {
            const found = rk.find(k => k.trim().toLowerCase() === alias.toLowerCase());
            if (found !== undefined) return row[found];
        }
        return '';
    }

    /**
     * Build the ProjectFlow data model from data rows + metadata
     */
    function buildProject(rows, meta) {
        const project = {
            name:        meta.name || 'Imported Project',
            manager:     meta.owner || '',
            startDate:   meta.startDate || new Date(),
            finishDate:  meta.finishDate || new Date(),
            minutesPerDay: 480,
            minutesPerWeek: 2400,
            daysPerMonth: 20,
            currencySymbol: '$',
            tasks:       [],
            resources:   [],
            assignments: []
        };

        const resMap      = new Map();   // name → uid
        const taskNumMap  = new Map();   // "Task number" string → uid we assign
        const bucketMap   = new Map();   // bucket name → count (for summary detection only)

        /* ──────── PASS 1: collect resources + count buckets ──────── */
        rows.forEach((row, idx) => {
            const tName = String(getVal(row, 'name') || '').trim();
            if (!tName) return;

            const bName = String(getVal(row, 'bucket') || '').trim() || 'Other';
            bucketMap.set(bName, (bucketMap.get(bName) || 0) + 1);

            const assigned = splitList(getVal(row, 'assigned'));
            assigned.forEach(n => {
                if (!resMap.has(n)) {
                    const uid = resMap.size + 1;
                    resMap.set(n, uid);
                    project.resources.push({ uid, id: uid, name: n, type: 1, maxUnits: 1, cost: 0 });
                }
            });
        });

        /* ──────── PASS 2: build tasks using OUTLINE NUMBER from Excel ──────── */
        let uidSeq = 1;

        rows.forEach((row, idx) => {
            const tName = String(getVal(row, 'name') || '').trim();
            if (!tName) return;

            const tNum = String(getVal(row, 'taskNum') || '');
            const tUID = uidSeq++;
            if (tNum) taskNumMap.set(tNum, tUID);
            
            // Allow looking up predecessors by WBS/Outline
            const outlineStr = String(getVal(row, 'outline') || '').trim();
            if (outlineStr) taskNumMap.set(outlineStr, tUID);

            // Use Outline Number directly for WBS and depth
            const dotCount = (outlineStr.match(/\./g) || []).length;
            const outlineLevel = dotCount + 1; // "1" → level 1, "1.1" → level 2, "1.1.1" → level 3

            const sDate = excelDate(getVal(row, 'start'));
            const fDate = excelDate(getVal(row, 'finish')) || sDate;
            let pct = parseProgress(getVal(row, 'progress'));
            const isMilestone = String(getVal(row, 'milestone') || '').toLowerCase() === 'yes';

            let durDays = 0;
            if (sDate && fDate) {
                durDays = Math.ceil((fDate.getTime() - sDate.getTime()) / 864e5);
                if (!isMilestone && durDays === 0) durDays = 1;
                durDays = Math.max(0, durDays);
            }

            const assigned = splitList(getVal(row, 'assigned'));

            project.tasks.push({
                uid: tUID, id: tUID,
                name: tName,
                wbs: outlineStr || '', // preserve original outline number as WBS
                outlineLevel: outlineLevel,
                outlineNumber: outlineStr || '',
                summary: false, // will be detected in Pass 3
                milestone: isMilestone,
                start: sDate || new Date(),
                finish: fDate || sDate || new Date(),
                durationDays: durDays,
                percentComplete: pct,
                notes: String(getVal(row, 'notes') || ''),
                resourceNames: assigned,
                isExpanded: true, isVisible: true,
                predecessors: [],
                _rawDeps: String(getVal(row, 'depends') || ''),
            });

            // Assignments
            assigned.forEach(n => {
                const rUID = resMap.get(n);
                if (rUID) project.assignments.push({ taskUID: tUID, resourceUID: rUID, units: 1 });
            });
        });

        /* ──────── PASS 3: detect summary tasks from outline ──────── */
        project.tasks.forEach(cur => {
            if (cur.wbs) {
                // If any task has a WBS that starts with "cur.wbs." (e.g. "1" -> "1.1")
                cur.summary = project.tasks.some(t => t.wbs && t.wbs.startsWith(cur.wbs + '.'));
            }
        });

        /* ──────── PASS 4: resolve dependencies ──────── */
        // Dependency formats accepted from MS Planner/Project Excel export:
        //   "3FS+2"       → task number 3, Finish-to-Start, lag +2 days
        //   "1.2.3FS"     → outline number 1.2.3, FS, no lag
        //   "9FS, 10FS"   → two predecessors (split on comma/semicolon)
        //   "3"  / "1.2"  → FS assumed, lag 0  (simpler export formats)
        //   "Design Phase FS+1" — name-based: fallback after regex fails
        //
        // DEP_RE accepts: plain integer OR dotted-outline (1.2.3)
        //   + optional link type (FS/FF/SS/SF) + optional lag (±N)
        const DEP_RE = /^(\d+(?:\.\d+)*)\s*(FS|FF|SS|SF)?\s*([+\-]\s*\d+)?$/i;
        const LINK_TYPE_CODE = { FS: 1, FF: 0, SS: 3, SF: 2 };

        // Build a name → uid map for the name-based fallback
        const nameToUid = new Map();
        project.tasks.forEach(t => {
            if (t.name) nameToUid.set(t.name.trim().toLowerCase(), t.uid);
        });

        project.tasks.forEach(t => {
            if (!t._rawDeps) { delete t._rawDeps; return; }

            const parts = t._rawDeps.split(/[,;]/).map(s => s.trim()).filter(Boolean);
            parts.forEach(part => {
                const m = part.match(DEP_RE);

                if (!m) {
                    // ── Fallback: try to match the leading text as a task name ──
                    // e.g. "Design Phase FS+2" → name="Design Phase", type=FS, lag=2
                    const nameFallback = part.match(/^(.+?)\s*(FS|FF|SS|SF)?\s*([+\-]\s*\d+)?$/i);
                    const candidateName = nameFallback ? nameFallback[1].trim().toLowerCase() : part.toLowerCase();
                    const nameUid = nameToUid.get(candidateName);

                    if (nameUid && nameUid !== t.uid) {
                        const linkType = ((nameFallback && nameFallback[2]) || 'FS').toUpperCase();
                        const lagDays  = (nameFallback && nameFallback[3])
                            ? parseInt(nameFallback[3].replace(/\s/g,'')) : 0;
                        t.predecessors.push({
                            predecessorUID: nameUid,
                            type: LINK_TYPE_CODE[linkType] ?? 1,
                            typeName: linkType,
                            lag: lagDays,
                        });
                        return;
                    }

                    // Completely unrecognised — log and track for audit
                    if (!t._missingPredecessors) t._missingPredecessors = [];
                    t._missingPredecessors.push(part);
                    console.warn(`[PlannerParser] Unrecognised dependency format "${part}" on task "${t.name}" — skipped`);
                    return;
                }

                const taskNumStr = m[1];                               // e.g. "3" or "1.2.3"
                const linkType   = (m[2] || 'FS').toUpperCase();       // e.g. "FS"
                const lagDays    = m[3] ? parseInt(m[3].replace(/\s/g,'')) : 0;

                // Resolve via taskNumMap (holds both "Task number" strings and outline strings)
                let pUID = taskNumMap.get(taskNumStr)
                        || taskNumMap.get(String(parseInt(taskNumStr)));

                if (!pUID) {
                    // Not found by number — try as a name last resort
                    pUID = nameToUid.get(taskNumStr.toLowerCase());
                }

                if (!pUID) {
                    if (!t._missingPredecessors) t._missingPredecessors = [];
                    t._missingPredecessors.push(part);
                    console.warn(`[PlannerParser] Predecessor "${part}" not found in project tasks (task "${t.name}") — skipped`);
                    return;
                }
                if (pUID === t.uid) return; // self-loop guard

                t.predecessors.push({
                    predecessorUID: pUID,
                    type: LINK_TYPE_CODE[linkType] ?? 1,
                    typeName: linkType,
                    lag: lagDays,
                });
            });

            // Audit summary for this task
            if (t._missingPredecessors?.length) {
                console.warn(`[PlannerParser] Task "${t.name}" has ${t._missingPredecessors.length} unresolved predecessor(s):`, t._missingPredecessors);
            }
            delete t._rawDeps;
            delete t._missingPredecessors;
        });

        /* ──────── Final: set global dates ──────── */
        let gStart = Infinity, gFinish = -Infinity;
        for (const t of project.tasks) {
            const ts = t.start.getTime();
            const tf = t.finish.getTime();
            if (ts < gStart) gStart = ts;
            if (tf > gFinish) gFinish = tf;
        }
        project.startDate  = gStart  === Infinity ? new Date() : new Date(gStart);
        project.finishDate = gFinish === -Infinity ? project.startDate : new Date(gFinish);

        return project;
    }

    /* ══════ HELPERS ══════ */

    /** Convert Excel serial number OR Date OR date-string → JS Date */
    function excelDate(val) {
        if (!val) return null;
        if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
        if (typeof val === 'number') {
            if (val < 10000) return null; // Ignore durations/small numbers erroneously passed
            // P1 #30: Excel serial date with 1900 leap-year bug correction
            // Excel treats 1900 as leap year (it wasn't), so serial > 59 is off by 1
            const adjusted = val > 59 ? val - 1 : val;
            return new Date(Math.round((adjusted - 25569) * 864e5));
        }
        const d = new Date(val);
        return isNaN(d.getTime()) ? null : d;
    }

    /** Parse progress value: handles 0.73, "73%", "73", "Completed", "In Progress" etc. */
    function parseProgress(val) {
        if (val === null || val === undefined || val === '') return 0;
        if (typeof val === 'number') {
            return val <= 1 ? Math.round(val * 100) : Math.min(100, Math.round(val));
        }
        const s = String(val).trim().toLowerCase();
        if (s.includes('complete') || s.includes('مكتمل')) return 100;
        if (s.includes('progress') || s.includes('تقدم'))  return 50;
        const n = parseFloat(s);
        if (!isNaN(n)) return n <= 1 ? Math.round(n * 100) : Math.min(100, Math.round(n));
        return 0;
    }

    /** Split a comma/semicolon/newline separated list */
    function splitList(val) {
        return String(val || '').split(/[;,\n]/).map(s => s.trim()).filter(Boolean);
    }

    export const PlannerParser = { parse };
