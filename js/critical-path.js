/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 * Proprietary Software — Unauthorized use prohibited.
 * https://www.linkedin.com/in/ahmed-m-fawzy
 */
/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — Critical Path Method (CPM) Engine v2
 *
 * Design principles:
 *  1. Summary predecessors are resolved to their last leaf child
 *     BEFORE the forward/backward pass — so the network math is
 *     always leaf-to-leaf and never encounters _ef=0 from a
 *     skipped summary task.
 *  2. "Isolated" treatment (LF=EF) is only applied when the
 *     project actually has connected leaf tasks.  A flat task
 *     list with no dependencies falls back to the classic
 *     "all tasks share projectEnd" model so only the latest
 *     task(s) are critical.
 *  3. Calendar start dates act as a minimum floor for ES —
 *     no task can start before its scheduled date even if the
 *     network would allow it (ASAP with date constraints).
 * ═══════════════════════════════════════════════════════
 */


    /* ─── helpers ─────────────────────────────────────────── */

    function _daysBetween(d1, d2) {
        const t1 = new Date(d1); t1.setHours(0,0,0,0);
        const t2 = new Date(d2); t2.setHours(0,0,0,0);
        if (typeof WorkCalendar !== 'undefined' && WorkCalendar.getWorkingDays)
            return WorkCalendar.getWorkingDays(t1, t2);
        return Math.round((t2 - t1) / 86400000);
    }

    function _minDate(tasks) {
        let min = Infinity;
        tasks.forEach(t => { const d = new Date(t.start).getTime(); if (d < min) min = d; });
        return new Date(min);
    }

    function _typeName(type) {
        switch (type) { case 0: return 'FF'; case 1: return 'FS'; case 2: return 'SF'; case 3: return 'SS'; default: return 'FS'; }
    }

    /* ─── Pre-process: resolve summary predecessors ─────────
     *
     * Summary tasks are skipped in the forward/backward pass
     * (_ef stays 0).  Any task whose predecessor IS a summary
     * task therefore gets depEnd = 0 and floats freely.
     *
     * Fix: for each summary task build a "last leaf descendant"
     * pointer, then rewrite all predecessor lists so they point
     * to that leaf instead of the summary.
     *
     * The original task.predecessors array is NOT mutated;
     * a new task._cpmPreds array is written and used by the
     * forward/backward pass.
     */
    function _buildCpmPreds(tasks, taskMap) {
        // Step 1 — for every summary task find its last non-summary descendant
        const summaryLastLeaf = new Map(); // summaryUid → leafUid

        tasks.forEach((t, idx) => {
            if (!t.summary) return;
            const level = t.outlineLevel || 1;
            let lastLeaf = null;
            for (let j = idx + 1; j < tasks.length; j++) {
                if ((tasks[j].outlineLevel || 1) <= level) break;
                if (!tasks[j].summary) lastLeaf = tasks[j];
            }
            if (lastLeaf) summaryLastLeaf.set(t.uid, lastLeaf.uid);
        });

        // Step 2 — build _cpmPreds for every task
        tasks.forEach(t => {
            if (!t.predecessors || t.predecessors.length === 0) {
                t._cpmPreds = [];
                return;
            }
            const resolved = [];
            const seen = new Set();

            t.predecessors.forEach(pred => {
                const predTask = taskMap.get(pred.predecessorUID);
                if (!predTask) {
                    // Stale or cross-project reference — log so the user can audit
                    console.warn(`[CPM] predecessorUID=${pred.predecessorUID} not found in taskMap`
                        + ` — predecessor link on "${t.name}" (uid=${t.uid}) skipped.`
                        + ' This may be a stale reference from a deleted or renamed task.');
                    return;
                }

                let targetUid = pred.predecessorUID;
                if (predTask.summary) {
                    // Replace summary with its last leaf child
                    const leafUid = summaryLastLeaf.get(pred.predecessorUID);
                    if (!leafUid) {
                        // Summary has no non-summary descendants (fully-nested structure).
                        // Cannot resolve to a leaf — log and skip to avoid a dangling reference.
                        console.warn(`[CPM] Summary uid=${pred.predecessorUID} "${predTask.name}"`
                            + ` has no non-summary leaf descendants`
                            + ` — predecessor link from "${t.name}" (uid=${t.uid}) skipped.`
                            + ' Check that the summary contains at least one leaf task.');
                        return;
                    }
                    targetUid = leafUid;
                }

                if (targetUid === t.uid) return; // self-loop guard
                if (seen.has(targetUid)) return;  // deduplicate
                seen.add(targetUid);

                resolved.push({ ...pred, predecessorUID: targetUid });
            });

            t._cpmPreds = resolved;
        });
    }

    /* ─── Main CPM computation ───────────────────────────── */

    function compute(tasks, minutesPerDay = 480) {
        if (!tasks || tasks.length === 0) return tasks;

        /* ── 0. Initialise CPM fields ── */
        const taskMap = new Map();
        tasks.forEach(t => {
            t._es = 0; t._ef = 0; t._ls = Infinity; t._lf = Infinity;
            t._totalFloat = 0; t._freeFloat = 0;
            t._critical = false; t._isolated = false;
            t.totalFloat = 0; t.freeFloat = 0; t.critical = false;
            taskMap.set(t.uid, t);
        });

        /* ── 1. Resolve summary predecessors ── */
        _buildCpmPreds(tasks, taskMap);

        /* ── 1b. Dependency health check ──
         * After resolving preds, test whether ANY leaf task has a predecessor.
         * If not — and there are at least 2 leaf tasks — the CPM result will
         * technically be "every task is critical" which is misleading.
         * Emit a warning and attach a _cpmWarning property on the tasks array
         * so the UI layer can surface a user-visible notice.
         */
        const leafTasks = tasks.filter(t => !t.summary);
        const resolvedPredCount = leafTasks.reduce((s, t) => s + (t._cpmPreds?.length || 0), 0);

        if (leafTasks.length > 1 && resolvedPredCount === 0) {
            const msg = `[CPM] No predecessor links resolved among ${leafTasks.length} leaf tasks.`
                + ' All tasks will appear critical (TF = 0), which is technically correct'
                + ' but may be misleading if dependencies exist in the source.'
                + ' Possible causes: Dataverse unavailable (Scenario B), or unrecognised'
                + ' dependency format in the Excel import (Scenario A).';
            console.warn(msg);
            // Expose on the array so callers (UI, network.js) can surface a banner
            tasks._cpmWarning = msg;
            tasks._cpmDepsAvailable = false;
        } else {
            delete tasks._cpmWarning;
            tasks._cpmDepsAvailable = resolvedPredCount > 0 || leafTasks.length <= 1;
        }

        /* ── 2. Build successors map (from resolved preds) ── */
        const successors = new Map(); // uid → [{ task, typeName, lag }]
        tasks.forEach(t => {
            t._cpmPreds.forEach(pred => {
                if (!successors.has(pred.predecessorUID))
                    successors.set(pred.predecessorUID, []);
                successors.get(pred.predecessorUID).push({
                    task: t,
                    typeName: pred.typeName || _typeName(pred.type ?? 1),
                    lag: pred.lag || 0,
                });
            });
        });

        /* ── 3. Topological sort with cycle detection ── */
        const UNVISITED = 0, PROCESSING = 1, DONE = 2;
        const visitState = new Map();
        const sorted = [];
        tasks.forEach(t => visitState.set(t.uid, UNVISITED));

        function topoSort(task) {
            const state = visitState.get(task.uid);
            if (state === DONE) return;
            if (state === PROCESSING)
                throw new Error(`Cycle at uid=${task.uid} (${task.name})`);
            visitState.set(task.uid, PROCESSING);
            task._cpmPreds.forEach(pred => {
                const pt = taskMap.get(pred.predecessorUID);
                if (pt) topoSort(pt);
            });
            visitState.set(task.uid, DONE);
            sorted.push(task);
        }
        tasks.forEach(t => { try { topoSort(t); } catch(e) { console.warn('[CPM]', e.message); } });

        /* ── 4. Forward Pass ── */
        const projectStart = _minDate(tasks);

        sorted.forEach(task => {
            if (task.summary) return;

            // Calendar constraint: task cannot start before its scheduled date
            const calES = _daysBetween(projectStart, new Date(task.start));

            let es = calES; // calendar is the minimum floor

            task._cpmPreds.forEach(pred => {
                const pt = taskMap.get(pred.predecessorUID);
                if (!pt || pt.summary) return;

                const type = pred.typeName || _typeName(pred.type ?? 1);
                const lag  = pred.lag || 0;
                let depEnd = 0;
                switch (type) {
                    case 'FS': depEnd = pt._ef + lag; break;
                    case 'SS': depEnd = pt._es + lag; break;
                    case 'FF': depEnd = pt._ef + lag - (task.durationDays || 0); break;
                    case 'SF': depEnd = pt._es + lag - (task.durationDays || 0); break;
                    default:   depEnd = pt._ef + lag;
                }
                es = Math.max(es, depEnd);
            });

            task._es = Math.max(0, es);
            task._ef = task._es + Math.max(0, task.durationDays || 0);
        });

        /* ── 5. Determine project end ──
         *
         * "Connected" = has at least one real (non-summary) predecessor
         * or at least one real (non-summary) successor.
         *
         * If NO tasks are connected (pure flat list), we skip isolated
         * detection and let all tasks share the same projectEnd — only
         * the latest-ending tasks will have float ≈ 0 (correct).
         */
        const anyConnected = tasks.some(t =>
            !t.summary && (
                t._cpmPreds.length > 0 ||
                (successors.get(t.uid) || []).some(s => !s.task.summary)
            )
        );

        const isolatedUids = new Set();
        if (anyConnected) {
            tasks.forEach(t => {
                if (t.summary) return;
                const hasPreds = t._cpmPreds.length > 0;
                const hasSuccs = (successors.get(t.uid) || []).some(s => !s.task.summary);
                if (!hasPreds && !hasSuccs) {
                    isolatedUids.add(t.uid);
                    t._isolated = true;
                }
            });
        }

        // projectEnd = max EF of connected (non-isolated) leaf tasks
        const connectedEFs = tasks
            .filter(t => !t.summary && !isolatedUids.has(t.uid))
            .map(t => t._ef).filter(isFinite);
        const allLeafEFs = tasks
            .filter(t => !t.summary)
            .map(t => t._ef).filter(isFinite);
        const projectEnd = Math.max(...(connectedEFs.length ? connectedEFs : allLeafEFs), 0);

        /* ── 6. Backward Pass ── */
        for (let i = sorted.length - 1; i >= 0; i--) {
            const task = sorted[i];
            if (task.summary) continue;

            const realSuccs = (successors.get(task.uid) || [])
                .filter(s => !s.task.summary);

            if (isolatedUids.has(task.uid)) {
                // Isolated task: its own mini critical path
                task._lf = task._ef;
            } else if (realSuccs.length === 0) {
                // Terminal connected task
                task._lf = projectEnd;
            } else {
                let lf = Infinity;
                realSuccs.forEach(succ => {
                    if (!isFinite(succ.task._ls) || !isFinite(succ.task._lf)) return;
                    const type = succ.typeName || 'FS';
                    const lag  = succ.lag || 0;
                    const dur  = task.durationDays || 0;
                    let val;
                    switch (type) {
                        case 'FS': val = succ.task._ls - lag; break;
                        case 'SS': val = succ.task._ls - lag + dur; break;
                        case 'FF': val = succ.task._lf - lag; break;
                        case 'SF': val = succ.task._lf - lag + dur; break;
                        default:   val = succ.task._ls - lag;
                    }
                    if (isFinite(val)) lf = Math.min(lf, val);
                });
                task._lf = isFinite(lf) ? lf : projectEnd;
            }

            task._ls = task._lf - Math.max(0, task.durationDays || 0);
        }

        /* ── 7. Float & Criticality ── */
        tasks.forEach(task => {
            if (task.summary) return;
            task._totalFloat = Math.max(0, task._ls - task._es);
            task.totalFloat  = task._totalFloat;
            task._critical   = task._totalFloat < 0.001;
            task.critical    = task._critical;
        });

        /* ── 7b. Dataverse isCritical override ──
         * Planner Premium / Project for the Web uses a calendar-aware, resource-aware
         * scheduler that accounts for weekends, holidays, and constraints — our simple
         * duration-based CPM cannot replicate this exactly and will produce different
         * TF values.  Whenever Dataverse supplies msdyn_iscritical on ANY leaf task,
         * treat it as the authoritative critical path and override our calculation.
         *
         * This applies regardless of whether dep links were resolved: even with 122+
         * dependency records, the day-offset CPM may disagree with Planner's scheduler.
         */
        const dvCritCount = tasks.filter(t => !t.summary && t.isCritical === true).length;
        if (dvCritCount > 0) {
            tasks.forEach(task => {
                if (task.summary) return;
                task._critical = task.isCritical === true;
                task.critical  = task._critical;
                // Non-critical tasks get a nominal float so the filter hides them
                if (!task._critical && task._totalFloat < 0.001) {
                    task._totalFloat = task.durationDays || 1;
                    task.totalFloat  = task._totalFloat;
                }
            });
            console.log(`[CPM] Dataverse msdyn_iscritical applied: ${dvCritCount} critical tasks (overrides CPM TF)`);
        } else if (!tasks._cpmDepsAvailable) {
            // No Dataverse flags AND no dep links → all tasks wrongly appear critical.
            // Nothing useful to show; leave as-is and let the UI show the banner.
            console.log('[CPM] No Dataverse isCritical flags and no dep links — CPM result may be misleading');
        }

        /* ── 8. Free Float ── */
        tasks.forEach(task => {
            if (task.summary) return;
            const realSuccs = (successors.get(task.uid) || []).filter(s => !s.task.summary);
            if (!realSuccs.length) {
                task._freeFloat = task._totalFloat;
            } else {
                let minDriven = Infinity;
                realSuccs.forEach(s => {
                    const type = s.typeName || 'FS';
                    const lag  = s.lag || 0;
                    let driven;
                    switch (type) {
                        case 'FS': driven = s.task._es - lag; break;
                        case 'SS': driven = s.task._es - lag + (task.durationDays || 0); break;
                        case 'FF': driven = s.task._ef - lag; break;
                        case 'SF': driven = s.task._ef - lag + (task.durationDays || 0); break;
                        default:   driven = s.task._es - lag;
                    }
                    minDriven = Math.min(minDriven, driven);
                });
                task._freeFloat = Math.max(0, minDriven - task._ef);
            }
            task.freeFloat = task._freeFloat;
        });

        /* ── 9. Propagate results to summary tasks ── */
        // Walk bottom-up so nested summaries propagate correctly
        for (let i = tasks.length - 1; i >= 0; i--) {
            const task = tasks[i];
            if (!task.summary) continue;

            const level = task.outlineLevel || 1;
            let minES = Infinity, maxEF = -Infinity, isCrit = false;

            for (let j = i + 1; j < tasks.length; j++) {
                if ((tasks[j].outlineLevel || 1) <= level) break;
                if (tasks[j].summary) continue; // only leaf children
                if (isFinite(tasks[j]._es)) minES = Math.min(minES, tasks[j]._es);
                if (isFinite(tasks[j]._ef)) maxEF = Math.max(maxEF, tasks[j]._ef);
                if (tasks[j]._critical) isCrit = true;
            }

            task._es = isFinite(minES) ? minES : 0;
            task._ef = isFinite(maxEF) ? maxEF : 0;
            task._ls = task._es;
            task._lf = task._ef;
            task._totalFloat = 0; task.totalFloat = 0;
            task._freeFloat  = 0; task.freeFloat  = 0;
            task._critical   = isCrit; task.critical = isCrit;
        }

        /* ── 10. Clean up temp field ── */
        tasks.forEach(t => { delete t._cpmPreds; });

        return tasks;
    }

    /* ─── Baseline ─────────────────────────────────────── */

    function setBaseline(tasks) {
        tasks.forEach(t => {
            t.baselineStart    = new Date(t.start);
            t.baselineFinish   = new Date(t.finish);
            t.baselineDuration = t.durationDays;
        });
        return tasks;
    }

    function calculateVariance(tasks) {
        tasks.forEach(t => {
            if (!t.baselineStart || !t.baselineFinish) {
                t.startVariance = t.finishVariance = t.durationVariance = 0;
                return;
            }
            const bs = new Date(t.baselineStart), bf = new Date(t.baselineFinish);
            const as_ = new Date(t.start),        af  = new Date(t.finish);
            if ([bs,bf,as_,af].some(d => isNaN(d.getTime()))) {
                t.startVariance = t.finishVariance = t.durationVariance = 0;
                return;
            }
            t.startVariance    = _daysBetween(bs, as_);
            t.finishVariance   = _daysBetween(bf, af);
            t.durationVariance = (t.durationDays || 0) - (t.baselineDuration || 0);
        });
        return tasks;
    }

    /* ─── Status ───────────────────────────────────────── */

    function calculateStatus(tasks) {
        const today = new Date(); today.setHours(0,0,0,0);
        const PRI = { late: 4, 'at-risk': 3, 'on-track': 2, 'not-started': 1, complete: 0 };
        const META = {
            complete:     ['✅','#22c55e'],
            'not-started':['⬜','#64748b'],
            'on-track':   ['🟢','#22c55e'],
            'at-risk':    ['🟡','#f59e0b'],
            late:         ['🔴','#ef4444'],
        };

        const _set = (t, s) => {
            t.status = s;
            [t.statusIcon, t.statusColor] = META[s] || ['⬜','#64748b'];
        };

        // Leaf tasks first
        tasks.forEach(t => {
            if (t.summary) return;

            if (t.percentComplete >= 100) { _set(t,'complete'); return; }

            const start  = new Date(t.start);
            const finish = new Date(t.finish);

            if (today < start) { _set(t,'not-started'); return; }

            if (today > finish) { _set(t,'late'); return; }

            const totalDur   = Math.max(_daysBetween(start, finish), 1);
            const elapsed    = _daysBetween(start, today);
            const expectedPct = Math.min(100, Math.round((elapsed / totalDur) * 100));

            if (t.percentComplete < expectedPct - 15) _set(t,'at-risk');
            else _set(t,'on-track');
        });

        // Summary tasks: inherit worst child status (bottom-up)
        for (let i = tasks.length - 1; i >= 0; i--) {
            const task = tasks[i];
            if (!task.summary) continue;
            const level = task.outlineLevel || 1;
            let worst = 'complete', worstP = 0;
            for (let j = i + 1; j < tasks.length; j++) {
                if ((tasks[j].outlineLevel || 1) <= level) break;
                const p = PRI[tasks[j].status] || 0;
                if (p > worstP) { worstP = p; worst = tasks[j].status; }
            }
            _set(task, worst);
        }

        return tasks;
    }

    export const CPMEngine = { compute, setBaseline, calculateVariance, calculateStatus };
