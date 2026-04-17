/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 * Proprietary Software — Unauthorized use prohibited.
 * https://www.linkedin.com/in/ahmed-m-fawzy
 */
/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — Critical Path Method (CPM) Engine
 * + Baseline tracking + Progress indicators
 * ═══════════════════════════════════════════════════════
 */


    /**
     * Compute Critical Path for a list of tasks
     * Uses Forward Pass + Backward Pass
     * @param {Array} tasks — project tasks with predecessors
     * @param {number} minutesPerDay — working minutes per day (default 480)
     * @returns {Array} tasks with ES, EF, LS, LF, totalFloat, freeFloat, critical
     */
    function compute(tasks, minutesPerDay = 480) {
        if (!tasks || tasks.length === 0) return tasks;

        // Idempotency: computed props are cleared at the start of the pass
        
        // Build lookup
        const taskMap = new Map();
        tasks.forEach(t => {
            t._es = 0; t._ef = 0; t._ls = Infinity; t._lf = Infinity;
            t._totalFloat = 0; t._freeFloat = 0; t._critical = false;
            taskMap.set(t.uid, t);
        });

        // Build successors map
        const successors = new Map();
        tasks.forEach(t => {
            if (t.predecessors) {
                t.predecessors.forEach(pred => {
                    if (!successors.has(pred.predecessorUID)) {
                        successors.set(pred.predecessorUID, []);
                    }
                    successors.get(pred.predecessorUID).push({
                        task: t,
                        type: pred.type || 1, // FS default
                        typeName: pred.typeName || 'FS',
                        lag: pred.lag || 0
                    });
                });
            }
        });

        // ─── Forward Pass (compute ES, EF) ───
        // Topological sort with cycle detection
        const UNVISITED = 0, PROCESSING = 1, DONE = 2;
        const visitState = new Map();
        const sorted = [];
        const cycleWarnings = [];

        tasks.forEach(t => visitState.set(t.uid, UNVISITED));
        
        function topoSort(task) {
            const state = visitState.get(task.uid);
            if (state === DONE) return;
            if (state === PROCESSING) {
                // Cycle detected — throw to prevent silent miscalculations
                throw new Error(`Cycle detected at task uid=${task.uid} (${task.name}). Network contains a circular dependency.`);
            }
            visitState.set(task.uid, PROCESSING);
            if (task.predecessors) {
                task.predecessors.forEach(pred => {
                    const predTask = taskMap.get(pred.predecessorUID);
                    if (predTask) topoSort(predTask);
                });
            }
            visitState.set(task.uid, DONE);
            sorted.push(task);
        }

        tasks.forEach(t => topoSort(t));

        // Forward pass
        const projectStartDate = getMinDate(tasks);

        sorted.forEach(task => {
            if (task.summary) return; // Skip summary tasks

            let es = 0;
            if (task.predecessors && task.predecessors.length > 0) {
                task.predecessors.forEach(pred => {
                    const predTask = taskMap.get(pred.predecessorUID);
                    if (!predTask) return;

                    let depEnd = 0;
                    const type = pred.typeName || getTypeName(pred.type);
                    const lag = pred.lag || 0;

                    switch (type) {
                        case 'FS': depEnd = predTask._ef + lag; break;
                        case 'SS': depEnd = predTask._es + lag; break;
                        case 'FF': depEnd = predTask._ef + lag - task.durationDays; break;
                        case 'SF': depEnd = predTask._es + lag - task.durationDays; break;
                        default:   depEnd = predTask._ef + lag;
                    }
                    es = Math.max(es, depEnd);
                });
            } else {
                // Tasks without predecessors start at their own start
                const taskStart = new Date(task.start);
                es = daysBetween(projectStartDate, taskStart);
            }

            task._es = Math.max(0, es);
            task._ef = task._es + Math.max(0, task.durationDays || 0); // Guard negative duration (P0 #14)
        });

        // ─── Backward Pass (compute LS, LF) ───
        const efValues = tasks.filter(t => !t.summary).map(t => t._ef);
        const projectEnd = efValues.length > 0 ? Math.max(...efValues) : 0;

        // Reverse order
        for (let i = sorted.length - 1; i >= 0; i--) {
            const task = sorted[i];
            if (task.summary) continue;

            const succs = successors.get(task.uid);
            if (!succs || succs.length === 0) {
                task._lf = projectEnd;
                task._ls = task._lf - Math.max(0, task.durationDays || 0); // Guard negative duration
            } else {
                let lf = Infinity;
                succs.forEach(succ => {
                    // Validate successor values rigorously before use
                    if (!isFinite(succ.task._ls) || !isFinite(succ.task._lf)) return;
                    const type = succ.typeName || getTypeName(succ.type);
                    const lag = succ.lag || 0;
                    const dur = task.durationDays || 0;

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
                task._ls = task._lf - Math.max(0, task.durationDays || 0); // Guard negative duration
            }
        }

        // Calc Float
        tasks.forEach(task => {
            if (task.summary) return;
            task._totalFloat = Math.max(0, task._ls - task._es);
            task.totalFloat = task._totalFloat;
            task._critical = task._totalFloat === 0;
            task.critical = task._critical;
        });

        // Free Float with Dependency Type Formula mapping
        tasks.forEach(task => {
            if (task.summary) return;
            const succs = successors.get(task.uid);
            if (!succs || succs.length === 0) {
                task._freeFloat = task._totalFloat;
            } else {
                let minSuccES = Infinity;
                succs.forEach(s => {
                    const type = s.typeName || getTypeName(s.type);
                    const lag = s.lag || 0;
                    let drivenFinish = Infinity;
                    switch (type) {
                        case 'FS': drivenFinish = s.task._es - lag; break;
                        case 'SS': drivenFinish = s.task._es - lag + (task.durationDays||0); break;
                        case 'FF': drivenFinish = s.task._ef - lag; break;
                        case 'SF': drivenFinish = s.task._ef - lag + (task.durationDays||0); break;
                        default:   drivenFinish = s.task._es - lag;
                    }
                    minSuccES = Math.min(minSuccES, drivenFinish);
                });
                task._freeFloat = Math.max(0, minSuccES - task._ef);
            }
            task.freeFloat = task._freeFloat;
        });

        // Mark summary tasks as critical if any child is
        tasks.forEach(task => {
            if (!task.summary) return;
            const level = task.outlineLevel || 1;
            const idx = tasks.indexOf(task);
            for (let j = idx + 1; j < tasks.length; j++) {
                if ((tasks[j].outlineLevel || 1) <= level) break;
                if (tasks[j]._critical) {
                    task.critical = true;
                    task._critical = true;
                    break;
                }
            }
        });

        return tasks;
    }

    // ─── Baseline ───

    /**
     * Set baseline: save current dates as baseline
     */
    function setBaseline(tasks) {
        tasks.forEach(t => {
            t.baselineStart = new Date(t.start);
            t.baselineFinish = new Date(t.finish);
            t.baselineDuration = t.durationDays;
        });
        return tasks;
    }

    /**
     * Calculate schedule variance for each task
     */
    function calculateVariance(tasks) {
        tasks.forEach(t => {
            if (t.baselineStart && t.baselineFinish) {
                const bStart = new Date(t.baselineStart);
                const bFinish = new Date(t.baselineFinish);
                const aStart = new Date(t.start);
                const aFinish = new Date(t.finish);

                // Guard against invalid dates producing NaN
                if (isNaN(bStart.getTime()) || isNaN(bFinish.getTime()) ||
                    isNaN(aStart.getTime()) || isNaN(aFinish.getTime())) {
                    t.startVariance = 0;
                    t.finishVariance = 0;
                    t.durationVariance = 0;
                    return;
                }

                t.startVariance = daysBetween(bStart, aStart); // positive = late
                t.finishVariance = daysBetween(bFinish, aFinish); // positive = late
                t.durationVariance = (t.durationDays || 0) - (t.baselineDuration || 0);
            } else {
                t.startVariance = 0;
                t.finishVariance = 0;
                t.durationVariance = 0;
            }
        });
        return tasks;
    }

    // ─── Progress Status ───

    /**
     * Calculate progress status for each task
     * Returns: 'on-track', 'at-risk', 'late', 'complete', 'not-started'
     */
    function calculateStatus(tasks) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        tasks.forEach(t => {
            if (t.summary) {
                // Summary inherits worst child status
                return;
            }

            if (t.percentComplete >= 100) {
                t.status = 'complete';
                t.statusIcon = '✅';
                t.statusColor = '#22c55e';
                return;
            }

            const start = new Date(t.start);
            const finish = new Date(t.finish);

            if (today < start) {
                t.status = 'not-started';
                t.statusIcon = '⬜';
                t.statusColor = '#64748b';
                return;
            }

            // Calculate expected progress
            const totalDuration = Math.max(daysBetween(start, finish), 1);
            const elapsed = daysBetween(start, today);
            const expectedPct = Math.min(100, Math.round((elapsed / totalDuration) * 100));

            if (today > finish && t.percentComplete < 100) {
                t.status = 'late';
                t.statusIcon = '🔴';
                t.statusColor = '#ef4444';
            } else if (t.percentComplete < expectedPct - 15) {
                t.status = 'at-risk';
                t.statusIcon = '🟡';
                t.statusColor = '#f59e0b';
            } else {
                t.status = 'on-track';
                t.statusIcon = '🟢';
                t.statusColor = '#22c55e';
            }
        });

        // Summary tasks: inherit worst child status
        const statusPriority = { 'late': 4, 'at-risk': 3, 'on-track': 2, 'not-started': 1, 'complete': 0 };
        for (let i = tasks.length - 1; i >= 0; i--) {
            const task = tasks[i];
            if (!task.summary) continue;
            const level = task.outlineLevel || 1;
            let worstStatus = 'complete';
            let worstPriority = 0;
            for (let j = i + 1; j < tasks.length; j++) {
                if ((tasks[j].outlineLevel || 1) <= level) break;
                const p = statusPriority[tasks[j].status] || 0;
                if (p > worstPriority) {
                    worstPriority = p;
                    worstStatus = tasks[j].status;
                }
            }
            task.status = worstStatus;
            const meta = { 'complete': ['✅','#22c55e'], 'not-started': ['⬜','#64748b'], 'on-track': ['🟢','#22c55e'], 'at-risk': ['🟡','#f59e0b'], 'late': ['🔴','#ef4444'] };
            task.statusIcon = meta[worstStatus]?.[0] || '⬜';
            task.statusColor = meta[worstStatus]?.[1] || '#64748b';
        }

        return tasks;
    }

    // ─── Helpers ───

    function daysBetween(d1, d2) {
        const t1 = new Date(d1); t1.setHours(0, 0, 0, 0);
        const t2 = new Date(d2); t2.setHours(0, 0, 0, 0);
        return Math.round((t2 - t1) / 86400000);
    }

    function getMinDate(tasks) {
        let min = Infinity;
        tasks.forEach(t => {
            const d = new Date(t.start).getTime();
            if (d < min) min = d;
        });
        return new Date(min);
    }

    function getTypeName(type) {
        switch (type) {
            case 0: return 'FF';
            case 1: return 'FS';
            case 2: return 'SF';
            case 3: return 'SS';
            default: return 'FS';
        }
    }
    export const CPMEngine = { compute, setBaseline, calculateVariance, calculateStatus };
