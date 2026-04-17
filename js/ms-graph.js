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

    // Detect if running inside an iframe (e.g. Microsoft Teams)
    function _isInIframe() {
        try { return window.self !== window.top; } catch (_) { return true; }
    }

    // Use popup when in iframe (Teams), redirect otherwise
    async function signIn() {
        try {
            if (!msalApp) throw new Error('MSGraphClient not configured.');

            if (_isInIframe()) {
                // Teams iframe blocks full-page redirects — use popup instead
                const result = await msalApp.loginPopup({ scopes: SCOPES, prompt: 'select_account' });
                return result;
            } else {
                // Standalone browser — redirect flow
                msalApp.loginRedirect({ scopes: SCOPES, prompt: 'select_account' });
            }
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
            // Fetch plan metadata, plan details (labels), ALL buckets, and ALL tasks
            // Use $top=999 to minimize pagination round-trips for large projects
            const [planData, planDetailsData, allBuckets, allTasks] = await Promise.all([
                _call('GET', `/planner/plans/${planId}`),
                _call('GET', `/planner/plans/${planId}/details`).catch(() => ({})),
                _fetchAllPages(`/planner/plans/${planId}/buckets?$top=100`),
                _fetchAllPages(`/planner/plans/${planId}/tasks?$top=999`),
            ]);

            const buckets = allBuckets.map(b => ({
                id:        b.id,
                name:      b.name,
                orderHint: b.orderHint,
            }));

            // Category descriptions map: { category1: "Design", category2: "Dev", ... }
            // Empty string means no label was set for that category slot
            const categoryDescriptions = planDetailsData.categoryDescriptions || {};

            return {
                plan: {
                    id:        planData.id,
                    title:     planData.title,
                    owner:     planData.owner,
                    createdBy: planData.createdBy,
                },
                buckets,
                tasks: allTasks,
                categoryDescriptions,
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
     * Graph $batch helper — bundles up to 20 sub-requests into one HTTP call.
     * Each request: { id, method, url }
     * Returns: array of { id, status, body } in same order as requests.
     */
    async function _batchCall(requests) {
        const token = await _getAccessToken();
        const response = await fetch(`${GRAPH_ENDPOINT}/$batch`, {
            method:  'POST',
            headers: {
                Authorization:  `Bearer ${token}`,
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ requests }),
            signal: AbortSignal.timeout(30000),
        });
        if (!response.ok) {
            const errText = await response.text().catch(() => response.statusText);
            throw new Error(`Graph $batch error ${response.status}: ${errText}`);
        }
        const result = await response.json();
        // Return responses sorted by id so callers can zip with original requests
        const map = Object.fromEntries((result.responses || []).map(r => [r.id, r]));
        return requests.map(req => map[req.id] || { id: req.id, status: 500, body: null });
    }

    /**
     * Resolve an array of user IDs to display names via Graph $batch.
     * Bundles 20 /users/{id} lookups per HTTP call instead of N individual calls.
     * Results are cached to avoid duplicate requests.
     */
    async function _resolveUserDisplayNames(userIds) {
        const missing = userIds.filter(id => id && !_userCache.has(id));
        if (missing.length > 0) {
            const BATCH_SIZE = 20; // Graph $batch limit
            for (let i = 0; i < missing.length; i += BATCH_SIZE) {
                const chunk = missing.slice(i, i + BATCH_SIZE);
                const requests = chunk.map((id, idx) => ({
                    id:     String(idx),
                    method: 'GET',
                    url:    `/users/${id}?$select=displayName,userPrincipalName,mail`,
                }));
                let responses;
                try {
                    responses = await _batchCall(requests);
                } catch (batchErr) {
                    // $batch failed (e.g. no permission) — fall back to empty names
                    console.warn('[MSGraph] $batch user lookup failed:', batchErr.message);
                    chunk.forEach(id => _userCache.set(id, id.substring(0, 8) + '…'));
                    continue;
                }
                responses.forEach((resp, idx) => {
                    const id = chunk[idx];
                    if (resp.status === 200 && resp.body) {
                        const v = resp.body;
                        const name = v.displayName
                            || (v.userPrincipalName ? v.userPrincipalName.split('@')[0] : null)
                            || (v.mail ? v.mail.split('@')[0] : null)
                            || id.substring(0, 8) + '…';
                        _userCache.set(id, name);
                    } else {
                        _userCache.set(id, id.substring(0, 8) + '…');
                    }
                });
            }
        }
        return userIds.map(id => _userCache.get(id) || id);
    }

    // ============================================================================
    // PLAN-LEVEL CACHE  (avoids redundant API calls on each auto-pull)
    // ============================================================================

    /**
     * Per-plan cache: categoryDescriptions + resolved member name map.
     * Keyed by planId. Populated on first importPlan; reused by every subsequent pull.
     * TTL: cleared when importPlan is called again for the same plan.
     */
    const _planCache = new Map();
    // planId → { categoryDescriptions: {…}, userIdToName: {…} }

    // ============================================================================
    // MAPPING: PLANNER → PROJECTFLOW
    // ============================================================================

    // Planner priority numbers → human-readable labels
    const PRIORITY_LABELS = { 0: 'Urgent', 1: 'Important', 2: 'Medium', 9: 'Low' };

    /**
     * Convert Planner plan data to a ProjectFlow project with a FLAT task list.
     * Buckets become summary tasks (outlineLevel 1).
     * Tasks inside each bucket become leaf tasks (outlineLevel 2).
     *
     * @param {Object} plan
     * @param {Array}  buckets
     * @param {Array}  tasks
     * @param {Object} taskDetailsMap       taskId → details object
     * @param {Object} userIdToName         userId  → displayName (pre-resolved)
     * @param {Object} categoryDescriptions { category1: "Label Name", ... } from plan details
     */
    function plannerToProject(plan, buckets, tasks, taskDetailsMap, userIdToName = {}, categoryDescriptions = {}) {
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

                // Map appliedCategories → real label names from categoryDescriptions
                // Falls back to raw key (e.g. "category1") only if no label was defined
                const tags = [];
                if (task.appliedCategories && typeof task.appliedCategories === 'object') {
                    Object.keys(task.appliedCategories).forEach(key => {
                        if (task.appliedCategories[key] === true) {
                            const labelName = categoryDescriptions[key];
                            tags.push(labelName && labelName.trim() ? labelName.trim() : key);
                        }
                    });
                }

                // Priority: 0=Urgent 1=Important 2=Medium 9=Low
                const priorityNum   = typeof task.priority === 'number' ? task.priority : null;
                const priorityLabel = priorityNum !== null
                    ? (PRIORITY_LABELS[priorityNum] || 'Medium')
                    : null;

                // Compute % complete from checklist when available (granular 0-100)
                // Falls back to Planner's coarse 0/50/100 value
                let pct = task.percentComplete || 0;
                if (details.checklist) {
                    const checkItems = Object.values(details.checklist);
                    if (checkItems.length > 0) {
                        const checked = checkItems.filter(i => i.isChecked).length;
                        pct = Math.round((checked / checkItems.length) * 100);
                    }
                }

                // Build notes = description + checklist items
                let notes = details.description || '';
                if (details.checklist && Object.keys(details.checklist).length > 0) {
                    const checklistItems = Object.values(details.checklist)
                        .sort((a, b) => (a.orderHint || '').localeCompare(b.orderHint || ''))
                        .map(item => `- ${item.isChecked ? '✓' : '○'} ${item.title}`)
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
                    percentComplete: pct,
                    priority:       priorityLabel,   // 'Urgent' | 'Important' | 'Medium' | 'Low' | null
                    priorityNum:    priorityNum,     // raw 0/1/2/9 for sorting
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
            const { plan, buckets, tasks, categoryDescriptions } = await getPlanDetails(planId);

            // ── Fetch ALL task details via $batch (20 per HTTP call, much faster than N individual calls) ──
            // Run batch groups in parallel (5 concurrent) for large projects
            const taskDetailsMap = {};
            const BATCH_SIZE = 20;
            const PARALLEL_BATCHES = 5; // 5 × 20 = 100 tasks per wave
            const allChunks = [];
            for (let i = 0; i < tasks.length; i += BATCH_SIZE) {
                allChunks.push(tasks.slice(i, i + BATCH_SIZE));
            }

            for (let w = 0; w < allChunks.length; w += PARALLEL_BATCHES) {
                const wave = allChunks.slice(w, w + PARALLEL_BATCHES);
                const waveResults = await Promise.allSettled(
                    wave.map(chunk => {
                        const requests = chunk.map((t, idx) => ({
                            id:     String(idx),
                            method: 'GET',
                            url:    `/planner/tasks/${t.id}/details`,
                        }));
                        return _batchCall(requests).then(responses => ({ chunk, responses }));
                    })
                );
                waveResults.forEach(result => {
                    if (result.status === 'fulfilled') {
                        const { chunk, responses } = result.value;
                        responses.forEach((resp, idx) => {
                            taskDetailsMap[chunk[idx].id] = resp.status === 200 && resp.body ? resp.body : {};
                        });
                    } else {
                        // Fallback: mark as empty
                        console.warn('[MSGraph] Batch wave failed:', result.reason?.message);
                    }
                });
            }

            // ── Resolve user display names (cached per plan) ──
            // Strategy 1: group members (Group.Read.All — always works)
            // Strategy 2: $batch /users/{id} fallback for any remaining IDs
            let userIdToName = {};

            // Check plan cache first — avoids re-fetching on repeated imports
            const cached = _planCache.get(planId);
            if (cached) {
                userIdToName = { ...cached.userIdToName };
            } else {
                if (plan.owner) {
                    try {
                        const members = await _fetchAllPages(
                            `/groups/${plan.owner}/members?$select=id,displayName,userPrincipalName&$top=100`
                        );
                        members.forEach(m => {
                            if (m.id && (m.displayName || m.userPrincipalName)) {
                                const name = m.displayName || m.userPrincipalName.split('@')[0];
                                userIdToName[m.id] = name;
                                _userCache.set(m.id, name);
                            }
                        });
                    } catch (e) {
                        console.warn('[MSGraph] Could not fetch group members:', e.message);
                    }
                }

                // $batch fallback for IDs still unresolved
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

                // Populate plan cache so future pulls skip this work
                _planCache.set(planId, { categoryDescriptions, userIdToName: { ...userIdToName } });
            }

            const project = plannerToProject(plan, buckets, tasks, taskDetailsMap, userIdToName, categoryDescriptions);

            // Seed the delta token so the first auto-pull only fetches *changes* from this point
            // (runs in background — don't await; a failure here just means first pull is a full fetch)
            _fetchDeltaTasks(planId).catch(() => {});

            return project;
        } catch (err) {
            throw new Error(`importPlan failed: ${err.message}`);
        }
    }

    /**
     * Import multiple Planner plans in sequence.
     * Returns an array of { planId, success, project?, error? }
     *
     * @param {string[]} planIds     - Plan IDs to import
     * @param {Function} onProgress  - Optional callback(completed, total, planTitle)
     */
    async function importMultiplePlans(planIds, onProgress) {
        const results = [];
        for (let i = 0; i < planIds.length; i++) {
            const planId = planIds[i];
            try {
                const proj = await importPlan(planId);
                results.push({ planId, success: true, project: proj });
            } catch (err) {
                results.push({ planId, success: false, error: err.message });
            }
            if (onProgress) onProgress(i + 1, planIds.length);
        }
        return results;
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
    // DELTA SYNC  (incremental task updates — only fetch what changed)
    // ============================================================================

    /**
     * Per-plan delta tokens.
     * On first call returns ALL tasks and seeds the token.
     * Subsequent calls return ONLY tasks changed since the last token — much faster.
     * If the token expires (410 Gone) the function automatically falls back to a full fetch.
     */
    const _deltaTokens = new Map(); // planId → deltaLink URL

    /**
     * Fetch changed tasks for a plan using delta query.
     * Returns { tasks, isDelta } where isDelta=false means a full refresh was done.
     */
    async function _fetchDeltaTasks(planId) {
        const storedToken = _deltaTokens.get(planId);
        const startPath   = storedToken
            ? storedToken.replace(GRAPH_ENDPOINT, '')
            : `/planner/plans/${planId}/tasks/delta`;

        const items = [];
        let nextPath = startPath;
        let tokenExpired = false;

        while (nextPath) {
            let result;
            try {
                result = await _call('GET', nextPath);
            } catch (err) {
                // 410 Gone = delta token expired → do a full refresh
                if (err.message.includes('410') || err.message.toLowerCase().includes('gone')
                    || err.message.includes('resync')) {
                    console.warn('[MSGraph] Delta token expired — doing full refresh');
                    _deltaTokens.delete(planId);
                    tokenExpired = true;
                    break;
                }
                throw err;
            }

            (result.value || []).forEach(t => items.push(t));

            if (result['@odata.deltaLink']) {
                // End of delta page — store new token
                _deltaTokens.set(planId, result['@odata.deltaLink']);
                nextPath = null;
            } else {
                nextPath = result['@odata.nextLink']
                    ? result['@odata.nextLink'].replace(GRAPH_ENDPOINT, '')
                    : null;
            }
        }

        // If token expired, recurse once for a clean full fetch + new token
        if (tokenExpired) {
            return _fetchDeltaTasks(planId);
        }

        return { tasks: items, isDelta: !!storedToken };
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
                    // Delta pull: only fetch tasks changed since last sync
                    const { tasks: changedTasks, isDelta } = await _fetchDeltaTasks(planId);
                    if (changedTasks.length > 0 || !isDelta) {
                        _mergeRemoteChanges(project, changedTasks, isDelta);
                        console.log(`[MSGraph] Auto-pull ✓ — ${isDelta ? changedTasks.length + ' changed' : 'full refresh'}`);
                    } else {
                        console.log('[MSGraph] Auto-pull ✓ — no changes');
                    }
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

    /**
     * Merge remote (delta or full) task list into the local project.
     *
     * @param {Object}  project      - Local ProjectFlow project (flat task list)
     * @param {Array}   remoteTasks  - Array from Graph (may include @removed entries for delta)
     * @param {boolean} isDelta      - true = only changed tasks; false = full task list
     */
    function _mergeRemoteChanges(project, remoteTasks, isDelta = false) {
        // Separate normal updates from deleted tasks (delta only)
        const deleted = new Set(
            remoteTasks
                .filter(t => t['@removed'])
                .map(t => t.id)
        );
        const remoteMap = new Map(
            remoteTasks
                .filter(t => !t['@removed'])
                .map(t => [t.id, t])
        );

        (project.tasks || []).forEach(localTask => {
            // Skip bucket summary rows
            if (localTask.summary || localTask.outlineLevel === 1) return;

            const plannerId = localTask._plannerId;

            // Handle deleted tasks — hide them from the Gantt
            if (deleted.has(plannerId)) {
                localTask.isVisible  = false;
                localTask._deleted   = true;
                return;
            }

            const remote = remoteMap.get(plannerId);
            // For full refreshes, if the remote task is absent it may have moved bucket — skip
            if (!remote) return;

            // Update only if remote was modified more recently (or it's a full refresh)
            const localMod  = new Date(localTask._lastModified || 0);
            const remoteMod = new Date(remote.lastModifiedDateTime || 0);

            if (!isDelta || remoteMod > localMod) {
                localTask.percentComplete = remote.percentComplete || 0;
                localTask._plannerEtag    = remote['@odata.etag'];
                localTask._lastModified   = remote.lastModifiedDateTime;

                // Title may have been edited in Planner
                if (remote.title && remote.title !== localTask.name) {
                    localTask.name = remote.title;
                }

                // Update dates
                if (remote.startDateTime) localTask.start  = remote.startDateTime.split('T')[0];
                if (remote.dueDateTime)   localTask.finish = remote.dueDateTime.split('T')[0];

                // Recompute duration
                if (localTask.start && localTask.finish) {
                    localTask.durationDays = Math.max(1, Math.ceil(
                        (new Date(localTask.finish) - new Date(localTask.start)) / 864e5
                    ));
                }

                // Sync assignments — keep real Azure AD IDs + resolve display names
                if (remote.assignments && typeof remote.assignments === 'object') {
                    const userIds = Object.keys(remote.assignments);
                    localTask._plannerAssigneeIds = userIds;
                    localTask.resourceNames = userIds.map(
                        id => _userCache.get(id) || id.substring(0, 8) + '…'
                    );
                }

                // Priority update
                if (typeof remote.priority === 'number') {
                    localTask.priorityNum   = remote.priority;
                    localTask.priority      = PRIORITY_LABELS[remote.priority] || 'Medium';
                }
            }
        });

        // Recompute bucket summary % from their (visible) leaf tasks
        const buckets = (project.tasks || []).filter(t => t.summary || t.outlineLevel === 1);
        buckets.forEach(bucket => {
            const leaves = (project.tasks || []).filter(
                t => t._plannerBucketId === bucket._plannerBucketId
                    && !t.summary && t.outlineLevel !== 1
                    && !t._deleted
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

        // ── Step 1: Sign In (popup in Teams iframe, redirect in standalone) ──
        function renderSignIn() {
            const msLogo = `<svg viewBox="0 0 21 21" fill="none" width="20" height="20" style="vertical-align:middle;margin-right:10px;flex-shrink:0">
                <rect x="1" y="1" width="9" height="9" fill="#f25022"/>
                <rect x="11" y="1" width="9" height="9" fill="#7fba00"/>
                <rect x="1" y="11" width="9" height="9" fill="#00a4ef"/>
                <rect x="11" y="11" width="9" height="9" fill="#ffb900"/>
            </svg>`;

            const inIframe = _isInIframe();

            const note = document.createElement('p');
            note.style.cssText = 'font-size:0.78rem;color:var(--text-muted,#888);text-align:center;margin:0 0 14px;line-height:1.5;';
            note.textContent = inIframe
                ? 'A popup will open for Microsoft login. Please allow popups if prompted.'
                : 'You will be redirected to Microsoft login and brought back automatically.';
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

            signInBtn.addEventListener('click', async () => {
                signInBtn.disabled = true;
                signInBtn.innerHTML = inIframe
                    ? '⏳ Waiting for sign-in popup…'
                    : '⏳ Redirecting to Microsoft…';
                try {
                    await signIn();
                    // If popup flow (iframe/Teams): we get here after successful sign-in
                    if (isAuthenticated()) {
                        showStatus('✅ Signed in successfully!', 'success');
                        await renderPlanSelection();
                    }
                } catch (err) {
                    showStatus('Sign-in failed: ' + err.message, 'error');
                    signInBtn.disabled = false;
                    signInBtn.innerHTML = `${msLogo} Sign in with Microsoft`;
                }
            });

            wizard.appendChild(signInBtn);
        }

        // ── Step 2: Select Plan(s) ──
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

                const label = document.createElement('div');
                label.style.cssText = 'font-weight:500;margin-bottom:10px;font-size:0.85rem;color:var(--text-secondary,#a0aec0);display:flex;justify-content:space-between;align-items:center;';
                label.innerHTML = `
                    <span>Select plans to import:</span>
                    <span style="font-size:11px;opacity:0.7">${plans.length} plan${plans.length !== 1 ? 's' : ''} found</span>
                `;
                wizard.appendChild(label);

                // ── Checkbox list ──
                const listWrap = document.createElement('div');
                listWrap.style.cssText = `
                    max-height: 220px; overflow-y: auto;
                    border: 1px solid rgba(255,255,255,0.12); border-radius: 8px;
                    background: var(--bg-input, rgba(0,0,0,0.2));
                    margin-bottom: 12px;
                `;

                const checkboxes = [];
                plans.forEach((plan, idx) => {
                    const row = document.createElement('label');
                    row.style.cssText = `
                        display: flex; align-items: center; gap: 10px;
                        padding: 9px 12px; cursor: pointer; font-size: 13px;
                        border-bottom: 1px solid rgba(255,255,255,0.06);
                        color: var(--text-primary, #e2e8f0);
                        transition: background 0.15s;
                    `;
                    row.addEventListener('mouseenter', () => row.style.background = 'rgba(255,255,255,0.05)');
                    row.addEventListener('mouseleave', () => row.style.background = '');

                    const cb = document.createElement('input');
                    cb.type = 'checkbox';
                    cb.value = plan.id;
                    cb.dataset.title = plan.title;
                    cb.style.cssText = 'cursor:pointer;flex-shrink:0;accent-color:#6366f1;';
                    cb.addEventListener('change', updateImportBtn);
                    checkboxes.push(cb);

                    const nameSpan = document.createElement('span');
                    nameSpan.textContent = plan.title;
                    nameSpan.style.overflow = 'hidden';
                    nameSpan.style.textOverflow = 'ellipsis';
                    nameSpan.style.whiteSpace = 'nowrap';

                    row.appendChild(cb);
                    row.appendChild(nameSpan);
                    listWrap.appendChild(row);
                });
                wizard.appendChild(listWrap);

                // ── Select-all toggle ──
                const toggleRow = document.createElement('div');
                toggleRow.style.cssText = 'margin-bottom:12px;font-size:12px;';
                const toggleLink = document.createElement('a');
                toggleLink.href = '#';
                toggleLink.textContent = 'Select all';
                toggleLink.style.cssText = 'color:#6366f1;text-decoration:none;';
                let allSelected = false;
                toggleLink.addEventListener('click', e => {
                    e.preventDefault();
                    allSelected = !allSelected;
                    checkboxes.forEach(cb => cb.checked = allSelected);
                    toggleLink.textContent = allSelected ? 'Deselect all' : 'Select all';
                    updateImportBtn();
                });
                toggleRow.appendChild(toggleLink);
                wizard.appendChild(toggleRow);

                // ── Import button ──
                const importBtn = document.createElement('button');
                importBtn.type = 'button';
                importBtn.textContent = '📥 Import Plan';
                importBtn.disabled = true;
                importBtn.style.cssText = `
                    width: 100%; padding: 14px;
                    background: linear-gradient(135deg, #22c55e, #16a34a);
                    color: white; border: none; border-radius: 8px;
                    font-size: 15px; font-weight: 600; cursor: pointer;
                    transition: all 0.2s; opacity: 0.5;
                `;
                importBtn.addEventListener('mouseenter', () => { if (!importBtn.disabled) importBtn.style.transform = 'translateY(-1px)'; });
                importBtn.addEventListener('mouseleave', () => importBtn.style.transform = '');

                function updateImportBtn() {
                    const selected = checkboxes.filter(cb => cb.checked);
                    importBtn.disabled = selected.length === 0;
                    importBtn.style.opacity = selected.length > 0 ? '1' : '0.5';
                    importBtn.style.cursor  = selected.length > 0 ? 'pointer' : 'not-allowed';
                    if (selected.length === 0) {
                        importBtn.textContent = '📥 Import Plan';
                    } else if (selected.length === 1) {
                        importBtn.textContent = `📥 Import 1 Plan`;
                    } else {
                        importBtn.textContent = `📥 Import ${selected.length} Plans → Portfolio`;
                    }
                }

                importBtn.addEventListener('click', () => {
                    const selected = checkboxes.filter(cb => cb.checked);
                    if (selected.length === 0) { showStatus('Please select at least one plan', 'error'); return; }

                    importBtn.disabled = true;
                    importBtn.style.opacity = '0.7';
                    importBtn.textContent = `⏳ Importing ${selected.length} plan${selected.length > 1 ? 's' : ''}…`;

                    if (selected.length === 1) {
                        // Single plan: original behavior
                        onComplete({ planId: selected[0].value, planTitle: selected[0].dataset.title });
                    } else {
                        // Multiple plans: portfolio import
                        const planIds    = selected.map(cb => cb.value);
                        const planTitles = selected.map(cb => cb.dataset.title);
                        onComplete({ planIds, planTitles, isPortfolioImport: true });
                    }
                });

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
                pullBtn.textContent = '⏳ Pulling…';
                const { tasks: changedTasks, isDelta } = await _fetchDeltaTasks(planId);
                _mergeRemoteChanges(project, changedTasks, isDelta);
                const changeCount = isDelta ? `${changedTasks.length} change${changedTasks.length !== 1 ? 's' : ''}` : 'full refresh';
                syncTime.textContent = `Last pull: ${new Date().toLocaleTimeString()} — ${changeCount}`;
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
        importMultiplePlans,
        pushTaskToPlanner,
        syncProjectToPlanner,
        startAutoSync,
        stopAutoSync,
        renderSetupWizard,
        renderSyncPanel,
    };
