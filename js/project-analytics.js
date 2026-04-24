/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 * Proprietary Software — Unauthorized use prohibited.
 * https://www.linkedin.com/in/ahmed-m-fawzy
 */
/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — Unified Analytics Engine
 * Sprint B.1 + B.4 — Eliminate O(n²) + Deduplicate Calcs
 *
 * Single-pass task analysis with:
 *   • O(n) resource Map lookups (replaces nested filter+find)
 *   • Dirty flag + cache — recomputes only on data change
 *   • Single source of truth for KPIs, health, footer stats
 * ═══════════════════════════════════════════════════════
 */



    /** @type {Object|null} Cached result of last compute() */
    let _cache = null;

    /** @type {boolean} True when data has changed and cache is stale */
    let _dirty = true;

    /** @type {Object|null} Reference to the project object last computed */
    let _projectRef = null;

    // ══════════════════════════════════════════════════════════════
    // RESOURCE MAPS  (B.1 — eliminates O(n²) nested loops)
    // ══════════════════════════════════════════════════════════════

    /**
     * Build O(1) lookup Maps from project's resources + assignments arrays.
     * Replaces the pattern:
     *   assignments.filter(a => a.taskUID === t.uid)
     *              .map(a => resources.find(r => r.uid === a.resourceUID))
     * with Map.get() calls — ~100x faster for 1000+ tasks.
     *
     * @param {Object} project
     * @returns {{ resourceByUID: Map, assignmentsByTask: Map, taskResourceNames: Map }}
     */
    function buildResourceMaps(project) {
        // Map: resourceUID → resource object
        const resourceByUID = new Map();
        (project.resources || []).forEach(r => resourceByUID.set(r.uid, r));

        // Map: taskUID → [assignment]
        const assignmentsByTask = new Map();
        (project.assignments || []).forEach(a => {
            if (!assignmentsByTask.has(a.taskUID)) assignmentsByTask.set(a.taskUID, []);
            assignmentsByTask.get(a.taskUID).push(a);
        });

        // Map: taskUID → [resourceName strings]
        const taskResourceNames = new Map();
        (project.tasks || []).forEach(t => {
            const assigns = assignmentsByTask.get(t.uid) || [];
            const names = assigns
                .map(a => { const r = resourceByUID.get(a.resourceUID); return r ? r.name : null; })
                .filter(Boolean);
            taskResourceNames.set(t.uid, names);
        });

        return { resourceByUID, assignmentsByTask, taskResourceNames };
    }

    /**
     * Apply resource names from Maps onto project.tasks in-place (O(n)).
     * @param {Object} project
     * @param {Map} taskResourceNames
     */
    function applyResourceNames(project, taskResourceNames) {
        (project.tasks || []).forEach(t => {
            // Only overwrite if assignments exist — preserve manually-entered names
            // when there are no structured assignments
            const mapped = taskResourceNames.get(t.uid);
            if (mapped !== undefined) {
                t.resourceNames = mapped;
            } else if (!t.resourceNames) {
                t.resourceNames = [];
            }
        });
    }

    // ══════════════════════════════════════════════════════════════
    // SINGLE-PASS ANALYTICS  (B.4 — eliminate duplicate calcs)
    // ══════════════════════════════════════════════════════════════

    /**
     * Run full project analytics in a single O(n) pass.
     * Results are cached until invalidate() is called.
     *
     * @param {Object} project
     * @returns {Object} analytics result (also stored in cache)
     */
    function compute(project) {
        if (!project || !project.tasks) { _cache = null; return null; }

        // Return cached result if clean
        if (!_dirty && _cache && _projectRef === project) return _cache;

        _projectRef = project;
        _dirty = false;

        // ── Build resource Maps & apply names ─────────────────
        const resourceMaps = buildResourceMaps(project);
        applyResourceNames(project, resourceMaps.taskResourceNames);

        // ── Single-pass task statistics ───────────────────────
        const leafTasks   = [];
        const summaryTasks = [];

        let totalPct     = 0;
        let critCount    = 0;
        let lateCount    = 0;
        let atRiskCount  = 0;
        let completeCount = 0;
        let inProgressCount = 0;
        let notStartedCount = 0;
        let milestoneCount  = 0;
        let milestoneCompleteCount = 0;
        let totalCost    = 0;
        let earnedCost   = 0;

        const now = Date.now();

        project.tasks.forEach(t => {
            if (t.summary) { summaryTasks.push(t); return; }
            leafTasks.push(t);

            const pct = t.percentComplete || 0;
            totalPct  += pct;
            totalCost  += t.cost || 0;
            earnedCost += (t.cost || 0) * pct / 100;

            if (pct >= 100)      completeCount++;
            else if (pct > 0)    inProgressCount++;
            else                 notStartedCount++;

            if (t.critical)         critCount++;
            if (t.status === 'late')     lateCount++;
            if (t.status === 'at-risk')  atRiskCount++;

            if (t.milestone) {
                milestoneCount++;
                if (pct >= 100) milestoneCompleteCount++;
            }
        });

        const total = leafTasks.length;
        const overallProgress = total > 0 ? Math.round(totalPct / total) : 0;

        // ── Time metrics ──────────────────────────────────────
        const startDate  = project.startDate  ? new Date(project.startDate)  : null;
        const finishDate = project.finishDate ? new Date(project.finishDate) : null;
        const totalDays  = (startDate && finishDate)
            ? Math.max(1, Math.round((finishDate - startDate) / 86400000))
            : 0;
        const elapsed = startDate
            ? Math.round((now - startDate.getTime()) / 86400000)
            : 0;
        const timeProgress = totalDays > 0
            ? Math.min(100, Math.max(0, Math.round((elapsed / totalDays) * 100)))
            : 0;
        const daysRemaining = finishDate
            ? Math.max(0, Math.round((finishDate.getTime() - now) / 86400000))
            : 0;

        // ── Health Score (replaces duplicate calculateHealthScore) ─
        const health = _computeHealth(project, leafTasks, total, lateCount, now, startDate, finishDate, totalDays);

        // ── Phase data ────────────────────────────────────────
        const phaseData = _computePhases(project);

        // ── SPI (quick estimate for health + footer) ──────────
        const avgPct = total > 0 ? (totalPct / total) / 100 : 0;
        const elapsedPct = totalDays > 0 ? Math.min(1, elapsed / totalDays) : 0;
        const spi = elapsedPct > 0 ? avgPct / elapsedPct : 1;

        _cache = {
            // Counts
            total, completeCount, inProgressCount, notStartedCount,
            critCount, lateCount, atRiskCount,
            milestoneCount, milestoneCompleteCount,
            // Rates
            overallProgress, spi,
            // Financials
            totalCost, earnedCost,
            // Time
            startDate, finishDate, totalDays, elapsed, timeProgress, daysRemaining,
            // Collections
            leafTasks, summaryTasks, phases: phaseData,
            // Health
            health,
            // Resource Maps (exposed for callers that need them)
            resourceMaps,
            // Meta
            timestamp: Date.now()
        };

        return _cache;
    }

    /**
     * Compute health score in O(n) using already-filtered leafTasks.
     * @private
     */
    function _computeHealth(project, leafTasks, total, lateCount, now, startDate, finishDate, totalDays) {
        if (!total) return { score: 0, label: 'Unknown', icon: '⬜', cssClass: 'not-started' };

        let score = 100;

        // 1. Late tasks (−5 each, max −30)
        score -= Math.min(30, lateCount * 5);

        // 2. Critical tasks stalled (−3 each, max −20)
        const critStalled = leafTasks.filter(t =>
            t.critical && (t.percentComplete || 0) < 50 && new Date(t.start) < new Date(now)
        ).length;
        score -= Math.min(20, critStalled * 3);

        // 3. SPI vs expected progress
        if (startDate && finishDate && totalDays > 0 && now > startDate.getTime()) {
            const elapsed = Math.min(1, (now - startDate.getTime()) / (finishDate.getTime() - startDate.getTime()));
            const avgPct = leafTasks.reduce((s, t) => s + (t.percentComplete || 0), 0) / total / 100;
            const spi = elapsed > 0 ? avgPct / elapsed : 1;
            if      (spi < 0.70) score -= 20;
            else if (spi < 0.85) score -= 10;
            else if (spi < 0.95) score -= 5;
        }

        // 4. Budget overrun (if cost data exists)
        const totalBudget = leafTasks.reduce((s, t) => s + (t.budget || t.cost || 0), 0);
        const actualCost  = leafTasks.reduce((s, t) => s + (t.actualCost || 0), 0);
        if (totalBudget > 0 && actualCost > totalBudget) {
            const overrun = (actualCost - totalBudget) / totalBudget;
            if      (overrun > 0.20) score -= 15;
            else if (overrun > 0.10) score -= 8;
            else                     score -= 4;
        }

        score = Math.max(0, Math.min(100, Math.round(score)));

        let label, icon, cssClass;
        if      (score >= 80) { label = 'Healthy';   icon = '🟢'; cssClass = 'healthy';   }
        else if (score >= 50) { label = 'At Risk';   icon = '🟡'; cssClass = 'at-risk';   }
        else                  { label = 'Critical';  icon = '🔴'; cssClass = 'critical';  }

        if (total === 0) { label = 'Empty'; icon = '⬜'; cssClass = 'not-started'; }

        return { score, label, icon, cssClass };
    }

    /**
     * Compute per-phase progress (summary task breakdown).
     * @private
     */
    function _computePhases(project) {
        const phases = project.tasks.filter(t => t.summary && t.outlineLevel === 1);
        return phases.map(p => {
            const idx = project.tasks.indexOf(p);
            const nextSummaryIdx = project.tasks.findIndex(
                (t, i) => i > idx && t.summary && t.outlineLevel <= p.outlineLevel
            );
            const end = nextSummaryIdx === -1 ? project.tasks.length : nextSummaryIdx;
            const phaseTasks = project.tasks.slice(idx + 1, end).filter(t => !t.summary);
            const avg = phaseTasks.length > 0
                ? Math.round(phaseTasks.reduce((s, t) => s + (t.percentComplete || 0), 0) / phaseTasks.length)
                : 0;
            return {
                name: p.name,
                progress: avg,
                taskCount: phaseTasks.length,
                critical: phaseTasks.filter(t => t.critical).length
            };
        });
    }

    // ══════════════════════════════════════════════════════════════
    // CACHE CONTROL
    // ══════════════════════════════════════════════════════════════

    /** Mark cache as stale — call after any data mutation */
    function invalidate() { _dirty = true; }

    /** @returns {Object|null} Last cached result without recomputing */
    function getCache() { return _cache; }

    /** @returns {boolean} */
    function isDirty() { return _dirty; }

    /** Wipe cache completely (e.g. on project switch) */
    function reset() {
        _cache = null;
        _dirty = true;
        _projectRef = null;
    }

    export const ProjectAnalytics = {
        compute,
        buildResourceMaps,
        applyResourceNames,
        invalidate,
        getCache,
        isDirty,
        reset
    };
