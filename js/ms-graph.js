/**
 * ProjectFlow™ © 2026 Ahmed M. Fawzy
 * Microsoft Graph API Client for Live Bi-Directional Planner Sync
 * Uses MSAL.js (PublicClientApplication) with PKCE flow
 */
import * as msal from '@azure/msal-browser';



    // ============================================================================
    // STATE & CONFIG
    // ============================================================================

    let msalApp = null;
    const CONFIG_KEY = 'pf_msgraph_config';
    const SCOPES = ['Tasks.ReadWrite', 'Group.Read.All', 'User.Read', 'User.ReadBasic.All', 'offline_access'];
    const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0';
    let autoSyncInterval = null;

    // ── ProjectFlow Commercial App — Multi-Tenant Azure AD ──
    // Registered by: Ahmed M. Fawzy | App: ProjectFlow
    // Supports any Microsoft 365 organization (multi-tenant)
    const DEFAULT_CLIENT_ID = '5c5eccbf-b7fb-4041-b969-44da0d6cf406';
    const DEFAULT_TENANT    = 'organizations'; // Any work/school Microsoft account

    // ============================================================================
    // AUTHENTICATION
    // ============================================================================

    // Detect the correct redirect URI (works on GitHub Pages, localhost, Teams)
    function _getRedirectUri() {
        const origin = window.location.origin;
        const path   = window.location.pathname.replace(/\/[^/]*$/, '/'); // strip filename
        return origin + path;
    }

    function configure(clientId, tenantId) {
        try {
            const config = {
                auth: {
                    clientId,
                    authority: `https://login.microsoftonline.com/${tenantId || DEFAULT_TENANT}`,
                    redirectUri: _getRedirectUri(),
                    navigateToLoginRequestUrl: false,
                },
                cache: {
                    cacheLocation: 'localStorage',
                    storeAuthStateInCookie: true,
                },
                system: {
                    allowNativeBroker: false,
                    loggerOptions: {
                        loggerCallback: () => {},
                        piiLoggingEnabled: false,
                    },
                },
            };

            msalApp = new msal.PublicClientApplication(config);
            localStorage.setItem(CONFIG_KEY, JSON.stringify({ clientId, tenantId: tenantId || DEFAULT_TENANT }));

            return msalApp.initialize();
        } catch (err) {
            throw new Error(`MSGraph configure failed: ${err.message}`);
        }
    }

    // Returns admin consent URL for IT admins of client organizations
    function getAdminConsentUrl(redirectUri) {
        const uri = redirectUri || _getRedirectUri();
        return `https://login.microsoftonline.com/organizations/adminconsent?client_id=${DEFAULT_CLIENT_ID}&redirect_uri=${encodeURIComponent(uri)}`;
    }

    // Use redirect instead of popup — works in Teams iframe + all browsers
    function signIn() {
        try {
            if (!msalApp) throw new Error('MSGraphClient not configured.');
            // loginRedirect navigates the page — no return value
            msalApp.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
        } catch (err) {
            throw new Error(`Sign-in failed: ${err.message}`);
        }
    }

    function signOut() {
        try {
            if (msalApp) {
                return msalApp.logoutPopup();
            }
        } catch (err) {
            throw new Error(`Sign-out failed: ${err.message}`);
        }
    }

    function isAuthenticated() {
        if (!msalApp) return false;
        const accounts = msalApp.getAllAccounts();
        return accounts && accounts.length > 0;
    }

    function getAccount() {
        try {
            if (!msalApp || !isAuthenticated()) {
                return null;
            }
            const accounts = msalApp.getAllAccounts();
            if (!accounts || accounts.length === 0) return null;

            const account = accounts[0];
            return {
                name: account.name || account.username,
                email: account.username,
                tenantId: account.tenantId,
            };
        } catch (err) {
            throw new Error(`getAccount failed: ${err.message}`);
        }
    }

    async function _getAccessToken() {
        try {
            if (!msalApp) {
                throw new Error('MSGraphClient not configured.');
            }

            const accounts = msalApp.getAllAccounts();
            if (!accounts || accounts.length === 0) {
                throw new Error('No authenticated account. Call signIn() first.');
            }

            try {
                const response = await msalApp.acquireTokenSilent({
                    scopes: SCOPES,
                    account: accounts[0],
                });
                return response.accessToken;
            } catch (silentErr) {
                // Fallback to popup
                const response = await msalApp.acquireTokenPopup({ scopes: SCOPES });
                return response.accessToken;
            }
        } catch (err) {
            throw new Error(`Failed to acquire access token: ${err.message}`);
        }
    }

    // ============================================================================
    // GRAPH API CALL WRAPPER
    // ============================================================================

    async function _call(method, path, body = null, retryCount = 0, extraHeaders = {}) {
        try {
            const token = await _getAccessToken();
            const url = `${GRAPH_ENDPOINT}${path}`;
            const headers = {
                Authorization: `Bearer ${token}`,
                'Content-Type': 'application/json',
                ...extraHeaders,
            };

            const options = { method, headers, signal: AbortSignal.timeout(30000) };
            if (body) options.body = JSON.stringify(body);

            const response = await fetch(url, options);

            // Handle 429 Rate Limit — exponential backoff
            if (response.status === 429) {
                const MAX_RETRIES = 5;
                if (retryCount >= MAX_RETRIES) {
                    throw new Error(`Rate limit exceeded after ${MAX_RETRIES} retries.`);
                }
                const retryAfter = parseInt(response.headers.get('Retry-After') || '1', 10);
                const delay = Math.max(retryAfter * 1000, Math.pow(2, retryCount) * 1000);
                console.warn(`[MSGraph] Rate limited. Retry ${retryCount + 1}/${MAX_RETRIES} in ${delay}ms`);
                await new Promise(r => setTimeout(r, delay));
                return _call(method, path, body, retryCount + 1, extraHeaders);
            }

            // Handle 409 Conflict (ETag mismatch) — refresh ETag and retry once
            if (response.status === 409 && retryCount === 0) {
                console.warn('[MSGraph] ETag conflict — will retry after refresh');
                throw new Error('ETag conflict. Task was modified remotely. Please refresh and retry.');
            }

            // Handle 204 No Content (PATCH/DELETE responses)
            if (response.status === 204) {
                // Capture new ETag if present
                const newEtag = response.headers.get('ETag');
                return newEtag ? { '@odata.etag': newEtag } : {};
            }

            if (!response.ok) {
                const errData = await response.text().catch(() => response.statusText);
                throw new Error(`Graph API error ${response.status}: ${errData}`);
            }

            const result = await response.json();
            return result;
        } catch (err) {
            if (err.name === 'TimeoutError') {
                throw new Error(`Graph API timeout (${method} ${path})`);
            }
            throw new Error(`Graph API call failed (${method} ${path}): ${err.message}`);
        }
    }

    // ============================================================================
    // PLANNER READ OPERATIONS
    // ============================================================================

    /**
     * Fetch ALL pages of a paginated Graph endpoint.
     * Follows @odata.nextLink automatically.
     */
    async function _fetchAllPages(path) {
        const items = [];
        let nextPath = path;
        while (nextPath) {
            // nextLink is a full URL; strip the base endpoint prefix so _call can add it back
            const relPath = nextPath.startsWith('http')
                ? nextPath.replace(GRAPH_ENDPOINT, '')
                : nextPath;
            const result = await _call('GET', relPath);
            (result.value || []).forEach(i => items.push(i));
            nextPath = result['@odata.nextLink']
                ? result['@odata.nextLink'].replace(GRAPH_ENDPOINT, '')
                : null;
        }
        return items;
    }

    async function getMyPlans() {
        try {
            // Fetch ALL tasks and ALL group memberships (follow pagination)
            const [myTasks, allMemberships] = await Promise.allSettled([
                _fetchAllPages('/me/planner/tasks?$select=planId&$top=100'),
                _fetchAllPages('/me/memberOf?$select=id,displayName&$top=100'),
            ]);

            const planIdSet = new Set();

            // Collect plan IDs from the user's own tasks
            if (myTasks.status === 'fulfilled') {
                myTasks.value.forEach(t => { if (t.planId) planIdSet.add(t.planId); });
            }

            // Collect plans from ALL groups (no 20-group limit)
            if (allMemberships.status === 'fulfilled') {
                // Filter to groups only — @odata.type is returned as metadata even without $select
                const groups = allMemberships.value.filter(
                    g => g['@odata.type'] === '#microsoft.graph.group'
                      || g['@odata.type'] === '#microsoft.graph.Group'
                      || (g.id && g.displayName && !g.userPrincipalName)  // fallback: groups have no UPN
                );
                const BATCH = 10; // parallel group-plan fetches
                for (let i = 0; i < groups.length; i += BATCH) {
                    const batch = groups.slice(i, i + BATCH);
                    const results = await Promise.allSettled(
                        batch.map(g => _call('GET', `/groups/${g.id}/planner/plans`))
                    );
                    results.forEach(r => {
                        if (r.status === 'fulfilled') {
                            (r.value?.value || []).forEach(p => planIdSet.add(p.id));
                        }
                    });
                }
            }

            if (planIdSet.size === 0) return [];

            // Fetch all plan details in parallel
            const BATCH = 15;
            const ids = [...planIdSet];
            const plans = [];
            for (let i = 0; i < ids.length; i += BATCH) {
                const batch = ids.slice(i, i + BATCH);
                const results = await Promise.allSettled(
                    batch.map(id => _call('GET', `/planner/plans/${id}`))
                );
                results.forEach(r => {
                    if (r.status === 'fulfilled' && r.value?.id) {
                        plans.push({
                            id:         r.value.id,
                            title:      r.value.title || '(Untitled)',
                            owner:      r.value.owner,
                            createdBy:  r.value.createdBy,
                        });
                    }
                });
            }

            // Sort alphabetically
            plans.sort((a, b) => a.title.localeCompare(b.title));
            return plans;
        } catch (err) {
            throw new Error(`getMyPlans failed: ${err.message}`);
        }
    }

    async function getGroupPlans(groupId) {
        try {
            const result = await _call('GET', `/groups/${groupId}/planner/plans`);
            return result.value || [];
        } catch (err) {
            throw new Error(`getGroupPlans failed: ${err.message}`);
        }
    }

    async function getAllMyGroups() {
        try {
            const result = await _call(
                'GET',
                "/me/memberOf?$filter=startswith(tolower(createdDateTime), '2')"
            );
            return result.value || [];
        } catch (err) {
            throw new Error(`getAllMyGroups failed: ${err.message}`);
        }
    }

    async function getPlanDetails(planId) {
        try {
            // Fetch plan metadata, ALL buckets, and ALL tasks (with pagination)
            const [planData, allBuckets, allTasks] = await Promise.all([
                _call('GET', `/planner/plans/${planId}`),
                _fetchAllPages(`/planner/plans/${planId}/buckets`),
                _fetchAllPages(`/planner/plans/${planId}/tasks`),
            ]);

            const buckets = allBuckets.map(b => ({
                id:        b.id,
                name:      b.name,
                orderHint: b.orderHint,
            }));

            return {
                plan: {
                    id:        planData.id,
                    title:     planData.title,
                    owner:     planData.owner,
                    createdBy: planData.createdBy,
                },
                buckets,
                tasks: allTasks,
            };
        } catch (err) {
            throw new Error(`getPlanDetails failed: ${err.message}`);
        }
    }

    async function getPlanTaskDetails(taskId) {
        try {
            return await _call('GET', `/planner/tasks/${taskId}/details`);
        } catch (err) {
            throw new Error(`getPlanTaskDetails failed: ${err.message}`);
        }
    }

    // ============================================================================
    // USER DISPLAY NAME CACHE
    // ============================================================================

    const _userCache = new Map(); // userId → displayName

    /**
     * Resolve an array of user IDs to display names via Graph API.
     * Results are cached to avoid duplicate requests.
     */
    async function _resolveUserDisplayNames(userIds) {
        const missing = userIds.filter(id => id && !_userCache.has(id));
        if (missing.length > 0) {
            const BATCH = 15;
            for (let i = 0; i < missing.length; i += BATCH) {
                const batch = missing.slice(i, i + BATCH);
                const results = await Promise.allSettled(
                    batch.map(id =>
                        _call('GET', `/users/${id}?$select=displayName,userPrincipalName,mail`)
                    )
                );
                results.forEach((r, idx) => {
                    const id = batch[idx];
                    if (r.status === 'fulfilled' && r.value) {
                        const v = r.value;
                        // Use best available name: displayName → first part of UPN → mail prefix
                        const name = v.displayName
                            || (v.userPrincipalName ? v.userPrincipalName.split('@')[0] : null)
                            || (v.mail ? v.mail.split('@')[0] : null)
                            || id.substring(0, 8) + '…';
                        _userCache.set(id, name);
                    } else {
                        // Permission denied or not found — keep UUID short form
                        _userCache.set(id, id.substring(0, 8) + '…');
                    }
                });
            }
        }
        return userIds.map(id => _userCache.get(id) || id);
    }

    // ============================================================================
    // MAPPING: PLANNER → PROJECTFLOW
    // ============================================================================

    /**
     * Convert Planner plan data to a ProjectFlow project with a FLAT task list.
     * Buckets become summary tasks (outlineLevel 1).
     * Tasks inside each bucket become leaf tasks (outlineLevel 2).
     *
     * @param {Object} plan
     * @param {Array}  buckets
     * @param {Array}  tasks
     * @param {Object} taskDetailsMap   taskId → details object
     * @param {Object} userIdToName     userId  → displayName (pre-resolved)
     */
    function plannerToProject(plan, buckets, tasks, taskDetailsMap, userIdToName = {}) {
        const project = {
            id: plan.id,
            name: plan.title,
            owner: plan.owner,
            createdBy: plan.createdBy,
            tasks: [],          // FLAT list — app uses outlineLevel for hierarchy
            resources: [],
            assignments: [],
            _plannerId: plan.id,
        };

        const today = new Date().toISOString().split('T')[0];
        let uidSeq = 1;

        // Sort buckets by orderHint (ascending)
        const sortedBuckets = [...buckets].sort((a, b) => {
            if (!a.orderHint && !b.orderHint) return 0;
            if (!a.orderHint) return 1;
            if (!b.orderHint) return -1;
            return a.orderHint.localeCompare(b.orderHint);
        });

        sortedBuckets.forEach(bucket => {
            // Sort tasks within bucket by orderHint (same ordering Planner uses)
            const bucketTasks = tasks
                .filter(t => t.bucketId === bucket.id)
                .sort((a, b) => {
                    if (!a.orderHint && !b.orderHint) return 0;
                    if (!a.orderHint) return 1;
                    if (!b.orderHint) return -1;
                    return a.orderHint.localeCompare(b.orderHint);
                });

            // Compute bucket date range from its children
            let bucketStart = null;
            let bucketFinish = null;

            bucketTasks.forEach(task => {
                const s = task.startDateTime ? task.startDateTime.split('T')[0] : today;
                const f = task.dueDateTime   ? task.dueDateTime.split('T')[0]   : _addDays(s, 1);
                if (!bucketStart  || s < bucketStart)  bucketStart  = s;
                if (!bucketFinish || f > bucketFinish) bucketFinish = f;
            });

            if (!bucketStart)  bucketStart  = today;
            if (!bucketFinish) bucketFinish = _addDays(bucketStart, bucketTasks.length || 1);

            const bucketDur = Math.max(1, Math.ceil(
                (new Date(bucketFinish) - new Date(bucketStart)) / 864e5
            ));

            // Compute bucket % complete from children
            let bucketPct = 0;
            if (bucketTasks.length > 0) {
                bucketPct = Math.round(
                    bucketTasks.reduce((s, t) => s + (t.percentComplete || 0), 0) / bucketTasks.length
                );
            }

            const bucketUid = uidSeq++;

            // ── Add bucket as summary task ──
            project.tasks.push({
                uid:            bucketUid,
                id:             bucketUid,
                name:           bucket.name,
                outlineLevel:   1,
                summary:        true,
                milestone:      false,
                start:          bucketStart,
                finish:         bucketFinish,
                durationDays:   bucketDur,
                percentComplete: bucketPct,
                resourceNames:  [],
                predecessors:   [],
                tags:           [],
                notes:          '',
                isExpanded:     true,
                isVisible:      true,
                _plannerId:     bucket.id,
                _plannerBucketId: bucket.id,
            });

            // ── Add each task as a flat leaf (outlineLevel 2) ──
            bucketTasks.forEach(task => {
                const details   = taskDetailsMap[task.id] || {};
                const startDate = task.startDateTime ? task.startDateTime.split('T')[0] : today;
                const finishDate = task.dueDateTime  ? task.dueDateTime.split('T')[0]  : _addDays(startDate, 1);
                const dur = Math.max(1, Math.ceil(
                    (new Date(finishDate) - new Date(startDate)) / 864e5
                ));

                // Resolve assignment user IDs → display names
                // Priority: resolved display name → cached → assignedBy name → short UUID
                const resourceNames = [];
                if (task.assignments && typeof task.assignments === 'object') {
                    Object.entries(task.assignments).forEach(([userId, assignment]) => {
                        const name = userIdToName[userId]
                            || _userCache.get(userId)
                            || assignment?.assignedBy?.user?.displayName
                            || assignment?.createdBy?.user?.displayName
                            || null;
                        if (name) resourceNames.push(name);
                    });
                }

                // Map appliedCategories → tags
                const tags = [];
                if (task.appliedCategories && typeof task.appliedCategories === 'object') {
                    Object.keys(task.appliedCategories).forEach(key => {
                        if (task.appliedCategories[key] === true) tags.push(key);
                    });
                }

                // Build notes = description + checklist
                let notes = details.description || '';
                if (details.checklist && Object.keys(details.checklist).length > 0) {
                    const checklistItems = Object.values(details.checklist)
                        .map(item => `- ${item.title}${item.isChecked ? ' ✓' : ''}`)
                        .join('\n');
                    notes = notes
                        ? `${notes}\n\nChecklist:\n${checklistItems}`
                        : `Checklist:\n${checklistItems}`;
                }

                const taskUid = uidSeq++;

                project.tasks.push({
                    uid:            taskUid,
                    id:             taskUid,
                    name:           task.title,
                    outlineLevel:   2,
                    summary:        false,
                    milestone:      false,
                    start:          startDate,
                    finish:         finishDate,
                    durationDays:   dur,
                    percentComplete: task.percentComplete || 0,
                    resourceNames,
                    tags,
                    notes,
                    predecessors:   [],
                    isExpanded:     true,
                    isVisible:      true,
                    _plannerId:       task.id,
                    _plannerEtag:     task['@odata.etag'],
                    _plannerBucketId: bucket.id,
                    // Real Azure AD user IDs — needed for correct PATCH/POST assignments
                    _plannerAssigneeIds: Object.keys(task.assignments || {}),
                });
            });
        });

        return project;
    }

    function _addDays(dateStr, days) {
        const date = new Date(dateStr);
        date.setDate(date.getDate() + days);
        return date.toISOString().split('T')[0];
    }

    // ============================================================================
    // MAPPING: PROJECTFLOW → PLANNER (PUSH)
    // ============================================================================

    /**
     * Safely convert a Date object or ISO date string to "YYYY-MM-DD"
     */
    function _toDateStr(val) {
        if (!val) return null;
        if (val instanceof Date) return val.toISOString().split('T')[0];
        if (typeof val === 'string') return val.split('T')[0]; // handle "2026-04-16T00:00:00Z"
        return null;
    }

    function projectTaskToPlanner(task, bucketId) {
        const body = {
            title: task.name || '',
        };

        if (task.percentComplete !== undefined) {
            // Planner only accepts 0, 50, or 100
            const rounded = Math.round(task.percentComplete / 50) * 50;
            body.percentComplete = Math.min(100, Math.max(0, rounded));
        }

        // Safe date handling — task.start may be a Date object or a string
        const startStr  = _toDateStr(task.start);
        const finishStr = _toDateStr(task.finish);

        if (startStr)  body.startDateTime = `${startStr}T00:00:00Z`;
        if (finishStr) body.dueDateTime   = `${finishStr}T23:59:59Z`;

        // Use REAL Azure AD user IDs stored at import time.
        // _plannerAssigneeIds is set by plannerToProject and kept up-to-date by _mergeRemoteChanges.
        // Fallback: if somehow missing, skip assignments (don't send fake IDs).
        if (task._plannerAssigneeIds && task._plannerAssigneeIds.length > 0) {
            body.assignments = {};
            task._plannerAssigneeIds.forEach(userId => {
                body.assignments[userId] = {
                    '@odata.type': '#microsoft.graph.plannerAssignment',
                    'orderHint': ' !',
                };
            });
        }

        // Categories → appliedCategories
        if (task.tags && task.tags.length > 0) {
            body.appliedCategories = {};
            task.tags.forEach(tag => {
                body.appliedCategories[tag] = true;
            });
        }

        return body;
    }

    // ============================================================================
    // IMPORT & PUSH OPERATIONS
    // ============================================================================

    async function importPlan(planId) {
        try {
            const { plan, buckets, tasks } = await getPlanDetails(planId);

            // Fetch ALL task details in parallel (much faster than sequential)
            const BATCH = 10; // max concurrent requests
            const taskDetailsMap = {};

            for (let i = 0; i < tasks.length; i += BATCH) {
                const batch = tasks.slice(i, i + BATCH);
                const results = await Promise.allSettled(
                    batch.map(t => getPlanTaskDetails(t.id))
                );
                results.forEach((r, idx) => {
                    const id = batch[idx].id;
                    taskDetailsMap[id] = r.status === 'fulfilled' ? r.value : {};
                });
            }

            // ── Strategy 1: Fetch group members (uses Group.Read.All — always works) ──
            // The plan.owner is the Microsoft 365 Group ID that owns the plan.
            // Group members always have displayName and are the source of truth for assignments.
            const userIdToName = {};
            if (plan.owner) {
                try {
                    const members = await _fetchAllPages(
                        `/groups/${plan.owner}/members?$select=id,displayName,userPrincipalName&$top=100`
                    );
                    members.forEach(m => {
                        if (m.id && (m.displayName || m.userPrincipalName)) {
                            const name = m.displayName || m.userPrincipalName.split('@')[0];
                            userIdToName[m.id] = name;
                            _userCache.set(m.id, name); // cache for merge/sync use
                        }
                    });
                } catch (e) {
                    console.warn('[MSGraph] Could not fetch group members:', e.message);
                }
            }

            // ── Strategy 2: Fallback — try /users/{id} for any still-unresolved IDs ──
            const allUserIds = new Set();
            tasks.forEach(task => {
                if (task.assignments && typeof task.assignments === 'object') {
                    Object.keys(task.assignments).forEach(uid => {
                        if (!userIdToName[uid]) allUserIds.add(uid);
                    });
                }
            });
            if (allUserIds.size > 0) {
                await _resolveUserDisplayNames([...allUserIds]);
                allUserIds.forEach(id => {
                    const name = _userCache.get(id);
                    if (name) userIdToName[id] = name;
                });
            }

            const project = plannerToProject(plan, buckets, tasks, taskDetailsMap, userIdToName);
            return project;
        } catch (err) {
            throw new Error(`importPlan failed: ${err.message}`);
        }
    }

    async function pushTaskToPlanner(task, planId) {
        try {
            if (!task._plannerBucketId && !task._plannerId) {
                throw new Error('Task missing bucket/plan IDs for push.');
            }

            const body = projectTaskToPlanner(task, task._plannerBucketId);

            let result;
            if (task._plannerId) {
                // Existing Planner task: PATCH — must send If-Match ETag header
                const etagHeaders = { 'If-Match': task._plannerEtag || '*' };
                result = await _call('PATCH', `/planner/tasks/${task._plannerId}`, body, 0, etagHeaders);
                // Update local ETag from response so next PATCH doesn't get 409
                if (result && result['@odata.etag']) {
                    task._plannerEtag = result['@odata.etag'];
                }
            } else {
                // New task: POST
                body.planId = planId;
                body.bucketId = task._plannerBucketId;
                result = await _call('POST', '/planner/tasks', body);
                // Store new Planner ID and ETag for future pushes
                if (result && result.id) {
                    task._plannerId    = result.id;
                    task._plannerEtag  = result['@odata.etag'];
                }
            }

            return result;
        } catch (err) {
            throw new Error(`pushTaskToPlanner failed: ${err.message}`);
        }
    }

    /**
     * READ-ONLY MODE: pushing to Planner is disabled.
     * ProjectFlow treats Planner as the source of truth — data flows IN only.
     * To push changes, upgrade to the bi-directional sync version.
     */
    async function syncProjectToPlanner(_project, _planId) {
        console.info('[MSGraph] Write-back disabled — this build is read-only from Planner.');
        return { updated: 0, created: 0, failed: 0, readOnly: true };
    }

    // ============================================================================
    // AUTO-SYNC
    // ============================================================================

    function startAutoSync(project, planId, intervalMs = 60000) {
        try {
            if (autoSyncInterval) {
                clearInterval(autoSyncInterval);
            }

            autoSyncInterval = setInterval(async () => {
                try {
                    // Pull-only: fetch remote changes and merge into local project
                    const { tasks: remoteTasks } = await getPlanDetails(planId);
                    _mergeRemoteChanges(project, remoteTasks);
                    console.log('[MSGraph] Auto-pull completed ✓');
                } catch (err) {
                    console.warn('[MSGraph] Auto-pull error:', err.message);
                }
            }, intervalMs);

            return autoSyncInterval;
        } catch (err) {
            throw new Error(`startAutoSync failed: ${err.message}`);
        }
    }

    function stopAutoSync() {
        if (autoSyncInterval) {
            clearInterval(autoSyncInterval);
            autoSyncInterval = null;
        }
    }

    function _mergeRemoteChanges(project, remoteTasks) {
        // project.tasks is a flat list — no recursion needed
        const remoteMap = new Map(remoteTasks.map(t => [t.id, t]));

        (project.tasks || []).forEach(localTask => {
            // Skip bucket summary rows (their _plannerId is a bucket ID, not a task ID)
            if (localTask.summary || localTask.outlineLevel === 1) return;

            const remote = remoteMap.get(localTask._plannerId);
            if (!remote) return;

            // Update only if remote was modified more recently
            const localMod  = new Date(localTask._lastModified || 0);
            const remoteMod = new Date(remote.lastModifiedDateTime || 0);

            if (remoteMod > localMod) {
                localTask.percentComplete  = remote.percentComplete || 0;
                localTask._plannerEtag     = remote['@odata.etag'];
                localTask._lastModified    = remote.lastModifiedDateTime;

                // Update start/finish if Planner has dates
                if (remote.startDateTime) {
                    localTask.start = remote.startDateTime.split('T')[0];
                }
                if (remote.dueDateTime) {
                    localTask.finish = remote.dueDateTime.split('T')[0];
                }

                // Sync assignments: keep real IDs and resolve display names
                if (remote.assignments && typeof remote.assignments === 'object') {
                    const userIds = Object.keys(remote.assignments);
                    // Keep real Azure AD IDs for future push operations
                    localTask._plannerAssigneeIds = userIds;
                    // Resolve to display names for UI
                    localTask.resourceNames = userIds.map(
                        id => _userCache.get(id) || id.substring(0, 8) + '…'
                    );
                }
            }
        });

        // Recompute bucket summary % from their leaf tasks
        const buckets = (project.tasks || []).filter(t => t.summary || t.outlineLevel === 1);
        buckets.forEach(bucket => {
            const leaves = (project.tasks || []).filter(
                t => t._plannerBucketId === bucket._plannerBucketId
                    && (t.outlineLevel !== 1 && !t.summary)
            );
            if (leaves.length > 0) {
                bucket.percentComplete = Math.round(
                    leaves.reduce((s, t) => s + (t.percentComplete || 0), 0) / leaves.length
                );
            }
        });
    }

    // ============================================================================
    // UI: SETUP WIZARD
    // ============================================================================

    /**
     * Auto-initialize MSAL using the built-in Client ID (no user input needed).
     * Always uses DEFAULT_CLIENT_ID — multi-tenant, works for any organization.
     */
    async function _autoInit() {
        if (msalApp) return true;
        try {
            await configure(DEFAULT_CLIENT_ID, DEFAULT_TENANT);
            return true;
        } catch (e) {
            console.warn('[MSGraph] _autoInit failed:', e.message);
            return false;
        }
    }

    /**
     * Try to sign in silently or catch a returning redirect result.
     * Returns true if authenticated (silently or via redirect).
     */
    async function trySilentSignIn() {
        try {
            const ready = await _autoInit();
            if (!ready) return false;

            // ── Case 1: Returning from loginRedirect ──
            let redirectResult = null;
            try { redirectResult = await msalApp.handleRedirectPromise(); } catch (_) {}
            if (redirectResult && redirectResult.account) {
                console.log('[MSGraph] Signed in via redirect ✓', redirectResult.account.username);
                return true;
            }

            // ── Case 2: Cached session exists ──
            const accounts = msalApp.getAllAccounts();
            if (!accounts || accounts.length === 0) return false;

            await msalApp.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
            return true;
        } catch (e) {
            return false;
        }
    }

    function renderSetupWizard(container, onComplete) {
        if (!container) {
            throw new Error('renderSetupWizard: container not found.');
        }

        const wizard = document.createElement('div');
        wizard.className = 'ms-graph-wizard';
        wizard.style.cssText = `
            max-width: 500px;
            margin: 20px auto;
            padding: 24px;
            border: 1px solid rgba(255,255,255,0.1);
            border-radius: 12px;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            background: var(--bg-card, #1e1e2e);
        `;

        // ── Status message helper ──
        function showStatus(msg, type) {
            let el = wizard.querySelector('.wizard-status');
            if (!el) {
                el = document.createElement('div');
                el.className = 'wizard-status';
                el.style.cssText = 'padding:10px;border-radius:8px;margin:12px 0;font-size:13px;text-align:center;';
                wizard.appendChild(el);
            }
            el.textContent = msg;
            el.style.background = type === 'error' ? 'rgba(239,68,68,0.15)' : type === 'success' ? 'rgba(34,197,94,0.15)' : 'rgba(59,130,246,0.15)';
            el.style.color = type === 'error' ? '#f87171' : type === 'success' ? '#4ade80' : '#60a5fa';
        }

        // ── Main flow: auto-init always succeeds (Client ID is baked in) ──
        async function startWizard() {
            wizard.innerHTML = '';

            // Header
            const header = document.createElement('div');
            header.style.cssText = 'text-align:center;margin-bottom:20px;';
            header.innerHTML = `
                <div style="font-size:2rem;margin-bottom:8px;">📋</div>
                <h3 style="margin:0 0 4px 0;font-size:1.1rem;">Connect to Microsoft Planner</h3>
                <p style="margin:0;font-size:0.8rem;color:var(--text-muted,#888);">Sign in with your Microsoft account to load your plans</p>
            `;
            wizard.appendChild(header);

            // Auto-init with built-in Client ID — always succeeds
            showStatus('⏳ Initializing...', 'info');
            const ready = await _autoInit();

            if (!ready) {
                showStatus('❌ Failed to initialize. Please refresh the page.', 'error');
                return;
            }

            // Already authenticated — go straight to plan selection
            if (isAuthenticated()) {
                showStatus('✅ Already signed in', 'success');
                await renderPlanSelection();
            } else {
                // First time — show sign-in button only (no Client ID prompt)
                wizard.querySelector('.wizard-status')?.remove();
                renderSignIn();
            }
        }

        // ── Step 1: Sign In (redirect flow — works in Teams iframe) ──
        function renderSignIn() {
            const msLogo = `<svg viewBox="0 0 21 21" fill="none" width="20" height="20" style="vertical-align:middle;margin-right:10px;flex-shrink:0">
                <rect x="1" y="1" width="9" height="9" fill="#f25022"/>
                <rect x="11" y="1" width="9" height="9" fill="#7fba00"/>
                <rect x="1" y="11" width="9" height="9" fill="#00a4ef"/>
                <rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
            </svg>`;

            const note = document.createElement('p');
            note.style.cssText = 'font-size:0.78rem;color:var(--text-muted,#888);text-align:center;margin:0 0 14px;line-height:1.5;';
            note.textContent = 'You will be redirected to Microsoft login and brought back automatically.';
            wizard.appendChild(note);

            const signInBtn = document.createElement('button');
            signInBtn.type = 'button';
            signInBtn.innerHTML = `${msLogo} Sign in with Microsoft`;
            signInBtn.style.cssText = `
                width:100%; padding:14px 20px;
                background:linear-gradient(135deg,#0078d4,#106ebe);
                color:white; border:none; border-radius:8px;
                font-size:15px; font-weight:600; cursor:pointer;
                transition:all 0.2s; display:flex; align-items:center; justify-content:center;
            `;
            signInBtn.addEventListener('mouseenter', () => signInBtn.style.transform = 'translateY(-2px)');
            signInBtn.addEventListener('mouseleave', () => signInBtn.style.transform = '');

            signInBtn.addEventListener('click', () => {
                // loginRedirect navigates away — no await needed
                signInBtn.disabled = true;
                signInBtn.innerHTML = '⏳ Redirecting to Microsoft…';
                try {
                    signIn(); // triggers full-page redirect
                } catch (err) {
                    showStatus('Sign-in failed: ' + err.message, 'error');
                    signInBtn.disabled = false;
                    signInBtn.innerHTML = `${msLogo} Sign in with Microsoft`;
                }
            });

            wizard.appendChild(signInBtn);
        }

        // ── Step 2: Select Plan ──
        async function renderPlanSelection() {
            const loadingMsg = document.createElement('div');
            loadingMsg.style.cssText = 'text-align:center;padding:20px;color:var(--text-muted,#888);font-size:0.85rem;';
            loadingMsg.innerHTML = '🔄 Loading your Planner plans...';
            wizard.appendChild(loadingMsg);

            try {
                const plans = await getMyPlans();
                loadingMsg.remove();

                if (!plans || plans.length === 0) {
                    showStatus('No plans found in your Planner account.', 'error');
                    return;
                }

                const label = document.createElement('label');
                label.textContent = 'Select a Plan to import:';
                label.style.cssText = 'font-weight:500;display:block;margin-bottom:10px;font-size:0.85rem;color:var(--text-secondary,#a0aec0);';
                wizard.appendChild(label);

                const select = document.createElement('select');
                select.style.cssText = 'width:100%;padding:10px;border:1px solid rgba(255,255,255,0.15);border-radius:8px;background:var(--bg-input,#2a2a3e);color:var(--text-primary,#e2e8f0);font-size:13px;margin-bottom:12px;';

                const placeholder = document.createElement('option');
                placeholder.value = '';
                placeholder.textContent = `-- ${plans.length} plan${plans.length > 1 ? 's' : ''} found --`;
                select.appendChild(placeholder);

                plans.forEach(plan => {
                    const option = document.createElement('option');
                    option.value = plan.id;
                    option.textContent = plan.title;
                    select.appendChild(option);
                });

                const importBtn = document.createElement('button');
                importBtn.type = 'button';
                importBtn.textContent = '📥 Import Plan';
                importBtn.style.cssText = `
                    width: 100%;
                    padding: 14px;
                    background: linear-gradient(135deg, #22c55e, #16a34a);
                    color: white;
                    border: none;
                    border-radius: 8px;
                    font-size: 15px;
                    font-weight: 600;
                    cursor: pointer;
                    transition: all 0.2s;
                `;
                importBtn.addEventListener('mouseenter', () => importBtn.style.transform = 'translateY(-1px)');
                importBtn.addEventListener('mouseleave', () => importBtn.style.transform = '');

                importBtn.addEventListener('click', () => {
                    const planId = select.value;
                    if (!planId) {
                        showStatus('Please select a plan', 'error');
                        return;
                    }
                    const planTitle = select.options[select.selectedIndex].textContent;
                    importBtn.disabled = true;
                    importBtn.textContent = '⏳ Importing...';
                    importBtn.style.opacity = '0.7';
                    onComplete({ planId, planTitle });
                });

                wizard.appendChild(select);
                wizard.appendChild(importBtn);

            } catch (err) {
                loadingMsg.remove();
                showStatus('Failed to load plans: ' + err.message, 'error');
            }
        }

        startWizard();
        container.appendChild(wizard);
    }

    // ============================================================================
    // UI: SYNC STATUS PANEL
    // ============================================================================

    function renderSyncPanel(container, project, planId) {
        if (!container) {
            throw new Error('renderSyncPanel: container not found.');
        }

        const panel = document.createElement('div');
        panel.className = 'ms-graph-sync-panel';
        panel.style.cssText = `
            max-width: 600px;
            padding: 20px;
            background: #f3f2f1;
            border: 1px solid #d4cfcb;
            border-radius: 8px;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
            font-size: 14px;
        `;

        const title = document.createElement('h3');
        title.textContent = 'Planner Sync Status';
        title.style.cssText = 'margin: 0 0 15px 0; font-size: 16px;';
        panel.appendChild(title);

        // Account info
        const accountInfo = document.createElement('p');
        const account = getAccount();
        accountInfo.textContent = account
            ? `Connected: ${account.email} | Plan: ${project.name}`
            : 'Not connected';
        accountInfo.style.cssText = 'margin: 0 0 10px 0; color: #3c3c3c;';
        panel.appendChild(accountInfo);

        // Last sync time
        const syncTime = document.createElement('p');
        syncTime.textContent = 'Last sync: Never';
        syncTime.style.cssText = 'margin: 0 0 15px 0; color: #605e5c; font-size: 12px;';
        panel.appendChild(syncTime);

        // Button group
        const buttonGroup = document.createElement('div');
        buttonGroup.style.cssText = 'display: flex; gap: 10px; margin-bottom: 15px; flex-wrap: wrap;';

        // ── Read-only badge ──
        const readOnlyBadge = document.createElement('div');
        readOnlyBadge.style.cssText = `
            padding: 4px 10px; border-radius: 12px; font-size: 11px; font-weight: 600;
            background: rgba(99,102,241,0.15); color: #818cf8;
            border: 1px solid rgba(99,102,241,0.3); margin-bottom: 12px;
            display: inline-block;
        `;
        readOnlyBadge.textContent = '🔒 Read-Only — Planner is the source of truth';
        panel.insertBefore(readOnlyBadge, buttonGroup);

        // ── Pull button ──
        const pullBtn = document.createElement('button');
        pullBtn.textContent = '🔄 Pull from Planner';
        pullBtn.style.cssText = `
            padding: 8px 16px; background: #0078d4; color: white;
            border: none; border-radius: 6px; cursor: pointer;
            font-size: 13px; font-weight: 600;
        `;
        pullBtn.addEventListener('click', async () => {
            try {
                pullBtn.disabled = true;
                pullBtn.textContent = '⏳ Pulling...';
                const { tasks: remoteTasks } = await getPlanDetails(planId);
                _mergeRemoteChanges(project, remoteTasks);
                syncTime.textContent = `Last pull: ${new Date().toLocaleTimeString()}`;
                pullBtn.textContent = '🔄 Pull from Planner';
                pullBtn.disabled = false;
            } catch (err) {
                alert(`Pull failed: ${err.message}`);
                pullBtn.textContent = '🔄 Pull from Planner';
                pullBtn.disabled = false;
            }
        });

        // ── Auto-pull toggle ──
        const autoSyncLabel = document.createElement('label');
        autoSyncLabel.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; font-size: 12px;';

        const autoSyncCheckbox = document.createElement('input');
        autoSyncCheckbox.type = 'checkbox';
        autoSyncCheckbox.style.cssText = 'cursor: pointer;';
        autoSyncCheckbox.addEventListener('change', (e) => {
            if (e.target.checked) {
                startAutoSync(project, planId, 60000);
                autoSyncLabel.style.color = '#4ade80';
                autoSyncLabel.querySelector('span').textContent = 'Auto-Pull: ON (every 60s)';
            } else {
                stopAutoSync();
                autoSyncLabel.style.color = '';
                autoSyncLabel.querySelector('span').textContent = 'Auto-Pull: OFF';
            }
        });

        const autoSyncText = document.createElement('span');
        autoSyncText.textContent = 'Auto-Pull: OFF';
        autoSyncLabel.appendChild(autoSyncCheckbox);
        autoSyncLabel.appendChild(autoSyncText);

        buttonGroup.appendChild(pullBtn);
        buttonGroup.appendChild(autoSyncLabel);
        panel.appendChild(buttonGroup);

        // Sync log
        const logTitle = document.createElement('p');
        logTitle.textContent = 'Sync Log:';
        logTitle.style.cssText = 'margin: 15px 0 8px 0; font-weight: 500; font-size: 12px;';
        panel.appendChild(logTitle);

        const logContainer = document.createElement('div');
        logContainer.style.cssText = `
            background: white;
            padding: 10px;
            border: 1px solid #d4cfcb;
            border-radius: 4px;
            max-height: 150px;
            overflow-y: auto;
            font-size: 11px;
            color: #605e5c;
            font-family: monospace;
        `;
        logContainer.textContent = '-- No sync events yet --';

        panel.appendChild(logContainer);

        container.appendChild(panel);
    }

    // ============================================================================
    // PUBLIC API
    // ============================================================================

    export const MSGraphClient = {
        configure,
        signIn,
        signOut,
        isAuthenticated,
        getAccount,
        trySilentSignIn,
        getAdminConsentUrl,
        getMyPlans,
        getGroupPlans,
        getAllMyGroups,
        getPlanDetails,
        getPlanTaskDetails,
        importPlan,
        pushTaskToPlanner,
        syncProjectToPlanner,
        startAutoSync,
        stopAutoSync,
        renderSetupWizard,
        renderSyncPanel,
    };
