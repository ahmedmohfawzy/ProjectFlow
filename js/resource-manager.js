/**
 * ProjectFlow™ — Professional Project Management System
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 * Proprietary Software — Unauthorized use prohibited.
 * https://www.linkedin.com/in/ahmed-m-fawzy
 */
/**
 * ═══════════════════════════════════════════════════════
 * ProjectFlow — Resource Manager
 * Resource sheet, assignment, histogram data
 * ═══════════════════════════════════════════════════════
 */


    /**
     * Calculate resource load data for histogram
     * @param {Array} tasks
     * @param {Array} resources
     * @param {Array} assignments
     * @returns {Object} { resourceLoads, overAllocations }
     */
    function calculateResourceLoad(tasks, resources, assignments) {
        if (!resources || resources.length === 0) return { resourceLoads: {}, overAllocations: [] };

        const loads = {};
        const overAllocs = [];

        resources.forEach(res => {
            loads[res.uid] = {
                resource: res,
                dailyLoad: {},  // date -> hours
                totalHours: 0,
                maxDailyHours: (res.maxUnits || 1) * 8
            };
        });

        // Build task map
        const taskMap = new Map();
        tasks.forEach(t => taskMap.set(t.uid, t));

        // Calculate daily allocation per resource
        assignments.forEach(asgn => {
            const task = taskMap.get(asgn.taskUID);
            const resLoad = loads[asgn.resourceUID];
            if (!task || !resLoad || task.summary || task.milestone) return;

            const start = new Date(task.start);
            const finish = new Date(task.finish);
            const days = Math.max(daysBetween(start, finish), 1);
            const dailyHours = ((asgn.units || 1) * 8);

            for (let d = 0; d < days; d++) {
                const date = new Date(start);
                date.setDate(date.getDate() + d);
                // Skip weekends
                if (date.getDay() === 0 || date.getDay() === 6) continue;

                const key = date.toISOString().split('T')[0];
                resLoad.dailyLoad[key] = (resLoad.dailyLoad[key] || 0) + dailyHours;
                resLoad.totalHours += dailyHours;

                // Check over-allocation
                if (resLoad.dailyLoad[key] > resLoad.maxDailyHours) {
                    overAllocs.push({
                        resourceUID: asgn.resourceUID,
                        resourceName: resLoad.resource.name,
                        date: key,
                        allocated: resLoad.dailyLoad[key],
                        capacity: resLoad.maxDailyHours
                    });
                }
            }
        });

        return { resourceLoads: loads, overAllocations: uniqueOverAllocs(overAllocs) };
    }

    function uniqueOverAllocs(arr) {
        const seen = new Set();
        return arr.filter(item => {
            const key = `${item.resourceUID}_${item.date}`;
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });
    }

    /**
     * Get resource utilization summary
     */
    function getUtilizationSummary(resourceLoads, projectStart, projectEnd) {
        const summary = [];
        const totalDays = daysBetween(new Date(projectStart), new Date(projectEnd));
        const workDays = Math.round(totalDays * 5 / 7); // Approximate

        Object.values(resourceLoads).forEach(rl => {
            const daysWorked = Object.keys(rl.dailyLoad).length;
            const utilization = workDays > 0 ? Math.round((daysWorked / workDays) * 100) : 0;
            summary.push({
                uid: rl.resource.uid,
                name: rl.resource.name,
                totalHours: Math.round(rl.totalHours),
                daysWorked,
                utilization: Math.min(utilization, 100),
                overAllocated: Object.values(rl.dailyLoad).some(h => h > rl.maxDailyHours)
            });
        });

        return summary;
    }

    /**
     * Generate histogram data for a specific resource
     */
    function getHistogramData(resourceLoad, projectStart, projectEnd) {
        const data = [];
        const start = new Date(projectStart);
        const end = new Date(projectEnd);

        // P1 #11: Use constant ms increment to avoid DST infinite loop
        for (let d = new Date(start); d.getTime() <= end.getTime(); d = new Date(d.getTime() + 86400000)) {
            if (d.getDay() === 0 || d.getDay() === 6) continue;
            const key = d.toISOString().split('T')[0];
            data.push({
                date: key,
                hours: resourceLoad.dailyLoad[key] || 0,
                capacity: resourceLoad.maxDailyHours,
                overAllocated: (resourceLoad.dailyLoad[key] || 0) > resourceLoad.maxDailyHours
            });
        }

        return data;
    }

    /**
     * Auto-assign resource to task
     */
    function assignResource(assignments, taskUID, resourceUID, units = 1) {
        // Check if already assigned
        const existing = assignments.find(a => a.taskUID === taskUID && a.resourceUID === resourceUID);
        if (existing) {
            existing.units = units;
            return assignments;
        }
        assignments.push({ taskUID, resourceUID, units });
        return assignments;
    }

    /**
     * Remove resource assignment
     */
    function unassignResource(assignments, taskUID, resourceUID) {
        return assignments.filter(a => !(a.taskUID === taskUID && a.resourceUID === resourceUID));
    }

    function daysBetween(d1, d2) {
        const t1 = new Date(d1); t1.setHours(0, 0, 0, 0);
        const t2 = new Date(d2); t2.setHours(0, 0, 0, 0);
        return Math.round((t2 - t1) / 86400000);
    }

    /**
     * Auto-Level: delay non-critical tasks to resolve over-allocations
     * Returns a list of changes made (for undo support)
     */
    function autoLevel(tasks, resources, assignments) {
        const changes = [];
        if (!resources || !assignments || !tasks) return changes;

        const taskMap = new Map();
        tasks.forEach(t => taskMap.set(t.uid, t));

        // Sort non-critical, non-summary tasks by float (highest first—most flexible)
        const moveable = tasks.filter(t => !t.summary && !t.milestone && !t.critical)
            .sort((a, b) => (b.totalFloat || 0) - (a.totalFloat || 0));

        let iterations = 0;
        const MAX_ITER = 50;

        while (iterations++ < MAX_ITER) {
            const { overAllocations } = calculateResourceLoad(tasks, resources, assignments);
            if (overAllocations.length === 0) break;

            // Find first over-allocation
            const oa = overAllocations[0];
            const oaDate = new Date(oa.date);

            // Find a moveable task on that date for that resource
            let moved = false;
            for (const task of moveable) {
                const ts = new Date(task.start); ts.setHours(0,0,0,0);
                const tf = new Date(task.finish); tf.setHours(0,0,0,0);

                // Is this task active on the over-allocated date?
                if (oaDate >= ts && oaDate <= tf) {
                    // Is this task assigned to the over-allocated resource?
                    const isAssigned = assignments.some(a => a.taskUID === task.uid && a.resourceUID === oa.resourceUID);
                    if (!isAssigned) continue;

                    // Can we delay it? Check float
                    const float = task.totalFloat || 0;
                    if (float <= 0) continue;

                    // Delay by 1 day
                    const oldStart = new Date(task.start);
                    task.start = new Date(task.start); task.start.setDate(task.start.getDate() + 1);
                    task.finish = new Date(task.finish); task.finish.setDate(task.finish.getDate() + 1);

                    changes.push({
                        taskUID: task.uid,
                        taskName: task.name,
                        oldStart: oldStart,
                        newStart: new Date(task.start),
                        delayed: 1
                    });
                    moved = true;
                    break;
                }
            }
            if (!moved) break; // No more tasks can be moved
        }

        return changes;
    }

    /**
     * Compute time tracking summary per resource
     */
    function getTimeTrackingSummary(tasks, resources, assignments) {
        const summary = [];
        if (!resources) return summary;

        const taskMap = new Map();
        tasks.forEach(t => taskMap.set(t.uid, t));

        resources.forEach(res => {
            let plannedHours = 0, actualHours = 0;
            const resAsgns = assignments.filter(a => a.resourceUID === res.uid);
            resAsgns.forEach(a => {
                const t = taskMap.get(a.taskUID);
                if (!t || t.summary) return;
                plannedHours += t.plannedHours || (t.durationDays * 8);
                actualHours += t.actualHours || 0;
            });
            summary.push({
                uid: res.uid, name: res.name,
                plannedHours, actualHours,
                variance: actualHours - plannedHours,
                efficiency: plannedHours > 0 ? Math.round((plannedHours / Math.max(actualHours, 1)) * 100) : 100
            });
        });
        return summary;
    }

    export const ResourceManager = {
        calculateResourceLoad,
        getUtilizationSummary,
        getHistogramData,
        assignResource,
        unassignResource,
        autoLevel,
        getTimeTrackingSummary
    };
