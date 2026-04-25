/**
 * ProjectFlow™ © 2026 Ahmed M. Fawzy
 * Microsoft Graph API Client for Live Bi-Directional Planner Sync
 * Uses MSAL.js (PublicClientApplication) with PKCE flow
 */
// MSAL is loaded globally via CDN <script> tag in index.html (UMD build v2.38.3)
// IMPORTANT: Read window.msal at CALL TIME, not at module-load time,
// because the CDN <script> may not have finished loading yet.
function _getMsal() {
    return window.msal || window.Msal || null;
}



    // ============================================================================
    // STATE & CONFIG
    // ============================================================================

    let msalApp = null;
    const CONFIG_KEY = 'pf_msgraph_config';
    const SCOPES = ['Tasks.ReadWrite', 'Group.Read.All', 'User.Read', 'User.ReadBasic.All', 'offline_access'];
    const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0';
    const GRAPH_BETA     = 'https://graph.microsoft.com/beta';
    let autoSyncInterval = null;
    const DV_CACHE_KEY = 'pf_dataverse_url';

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

    async function configure(clientId, tenantId) {
        try {
            const msalLib = _getMsal();
            if (!msalLib || !msalLib.PublicClientApplication) {
                throw new Error('MSAL library not loaded yet. Ensure the CDN script is in <head> or loaded before app modules.');
            }

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

            // Create MSAL instance — works with v2 and v3 UMD
            msalApp = new msalLib.PublicClientApplication(config);

            // MSAL v3+ requires initialize() before any other call
            if (typeof msalApp.initialize === 'function') {
                await msalApp.initialize();
            }

            // Handle any returning redirect result
            try { await msalApp.handleRedirectPromise(); } catch (_) {}

            localStorage.setItem(CONFIG_KEY, JSON.stringify({ clientId, tenantId: tenantId || DEFAULT_TENANT }));
            console.log('[MSGraph] MSAL configured ✓');
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

    // Detect if Teams SDK is available
    function _hasTeamsSDK() {
        return typeof microsoftTeams !== 'undefined' && microsoftTeams.authentication;
    }

    // Build auth.html URL relative to current page
    function _getAuthPopupUrl() {
        const origin = window.location.origin;
        const path   = window.location.pathname.replace(/\/[^/]*$/, '/');
        return origin + path + 'auth.html';
    }

    // Use Teams SDK auth popup in Teams, loginPopup for standalone (preserves page state)
    async function signIn() {
        try {
            if (!msalApp) throw new Error('MSGraphClient not configured.');

            // Build extra scopes for Dataverse (single consent prompt)
            const extraScopes = [];
            const cachedDvUrl = localStorage.getItem(DV_CACHE_KEY);
            if (cachedDvUrl) {
                extraScopes.push(`${cachedDvUrl}/user_impersonation`);
            }

            if (_isInIframe() && _hasTeamsSDK()) {
                const resultStr = await microsoftTeams.authentication.authenticate({
                    url: _getAuthPopupUrl(),
                    width: 600,
                    height: 600,
                });
                const accounts = msalApp.getAllAccounts();
                return accounts.length > 0 ? { account: accounts[0] } : JSON.parse(resultStr || '{}');
            } else {
                // Always use popup — avoids full page reload and preserves app state
                return await msalApp.loginPopup({
                    scopes: SCOPES,
                    extraScopesToConsent: extraScopes,
                    prompt: 'select_account',
                });
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
    async function _fetchAllPages(path, useBeta = false) {
        const items = [];
        let nextPath = path;
        const baseUrl = useBeta ? GRAPH_BETA : GRAPH_ENDPOINT;
        while (nextPath) {
            // nextLink is a full URL; strip the base endpoint prefix so we can add it back
            let relPath = nextPath;
            if (nextPath.startsWith('http')) {
                relPath = nextPath.replace(GRAPH_BETA, '').replace(GRAPH_ENDPOINT, '');
            }
            // Use the chosen base URL directly
            const token = await _getAccessToken();
            const url = `${baseUrl}${relPath}`;
            const response = await fetch(url, {
                method: 'GET',
                headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
                signal: AbortSignal.timeout(30000),
            });
            if (!response.ok) {
                const errData = await response.text().catch(() => response.statusText);
                throw new Error(`Graph API error ${response.status}: ${errData}`);
            }
            const result = await response.json();
            (result.value || []).forEach(i => items.push(i));
            nextPath = result['@odata.nextLink'] || null;
        }
        // Log first item keys when using beta (for hierarchy discovery)
        if (useBeta && items.length > 0) {
            console.log('[MSGraph BETA] Task keys:', Object.keys(items[0]));
            console.log('[MSGraph BETA] First task:', JSON.stringify(items[0], null, 2));
        }
        return items;
    }

    // Cache plans for 5 minutes to avoid repeated API calls
    let _plansCache = null;
    let _plansCacheTime = 0;
    const PLANS_CACHE_TTL = 300000; // 5 minutes

    async function getMyPlans(forceRefresh = false) {
        // Return cached plans if still fresh
        if (!forceRefresh && _plansCache && (Date.now() - _plansCacheTime) < PLANS_CACHE_TTL) {
            console.log(`[MSGraph] Returning ${_plansCache.length} cached plans`);
            return _plansCache;
        }

        try {
            // Step 1: Get user's group memberships (single paginated call)
            let groups = [];
            try {
                const memberships = await _fetchAllPages('/me/memberOf?$select=id,displayName&$top=100');
                groups = memberships.filter(
                    g => g['@odata.type'] === '#microsoft.graph.group'
                      || g['@odata.type'] === '#microsoft.graph.Group'
                      || (g.id && g.displayName && !g.userPrincipalName)
                );
            } catch (e) {
                console.warn('[MSGraph] Failed to fetch group memberships:', e.message);
            }

            console.log(`[MSGraph] Found ${groups.length} groups — fetching plans via $batch`);

            // Step 2: Use $batch to fetch plans from all groups (20 per batch call)
            const planIdSet = new Set();
            const planMetaMap = new Map(); // planId → { title, owner }

            if (groups.length > 0) {
                const BATCH_SIZE = 20; // Graph $batch limit
                for (let i = 0; i < groups.length; i += BATCH_SIZE) {
                    const chunk = groups.slice(i, i + BATCH_SIZE);
                    const requests = chunk.map((g, idx) => ({
                        id: String(idx),
                        method: 'GET',
                        url: `/groups/${g.id}/planner/plans`,
                    }));
                    try {
                        const responses = await _batchCall(requests);
                        responses.forEach(resp => {
                            if (resp.status === 200 && resp.body?.value) {
                                resp.body.value.forEach(p => {
                                    planIdSet.add(p.id);
                                    planMetaMap.set(p.id, {
                                        id: p.id,
                                        title: p.title || '(Untitled)',
                                        owner: p.owner,
                                        createdBy: p.createdBy,
                                    });
                                });
                            }
                        });
                    } catch (batchErr) {
                        console.warn('[MSGraph] $batch group plans failed:', batchErr.message);
                    }
                    // Small delay between batch calls
                    if (i + BATCH_SIZE < groups.length) {
                        await new Promise(r => setTimeout(r, 500));
                    }
                }
            }

            // Step 3: Also try user's own tasks for plans not in any group
            try {
                const myTasks = await _call('GET', '/me/planner/tasks?$select=planId&$top=50');
                (myTasks?.value || []).forEach(t => {
                    if (t.planId && !planMetaMap.has(t.planId)) {
                        planIdSet.add(t.planId);
                    }
                });
            } catch (e) {
                console.warn('[MSGraph] Failed to fetch user tasks:', e.message);
            }

            // Step 4: Fetch metadata for plans discovered via tasks (not already in meta map)
            const missingMeta = [...planIdSet].filter(id => !planMetaMap.has(id));
            if (missingMeta.length > 0) {
                const BATCH_SIZE = 20;
                for (let i = 0; i < missingMeta.length; i += BATCH_SIZE) {
                    const chunk = missingMeta.slice(i, i + BATCH_SIZE);
                    const requests = chunk.map((id, idx) => ({
                        id: String(idx),
                        method: 'GET',
                        url: `/planner/plans/${id}`,
                    }));
                    try {
                        const responses = await _batchCall(requests);
                        responses.forEach((resp, idx) => {
                            if (resp.status === 200 && resp.body?.id) {
                                planMetaMap.set(resp.body.id, {
                                    id: resp.body.id,
                                    title: resp.body.title || '(Untitled)',
                                    owner: resp.body.owner,
                                    createdBy: resp.body.createdBy,
                                });
                            }
                        });
                    } catch (batchErr) {
                        console.warn('[MSGraph] $batch plan details failed:', batchErr.message);
                    }
                }
            }

            const plans = [...planMetaMap.values()];

            // Sort alphabetically
            plans.sort((a, b) => a.title.localeCompare(b.title));

            // Cache result
            _plansCache = plans;
            _plansCacheTime = Date.now();
            console.log(`[MSGraph] Discovered ${plans.length} plans (cached for ${PLANS_CACHE_TTL / 1000}s)`);

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
                _fetchAllPages(`/planner/plans/${planId}/tasks?$top=999`, true),
            ]);

            const buckets = allBuckets.map(b => ({
                id:        b.id,
                name:      b.name,
                orderHint: b.orderHint,
            }));

            // Category descriptions map: { category1: "Design", category2: "Dev", ... }
            // Empty string means no label was set for that category slot
            const categoryDescriptions = planDetailsData.categoryDescriptions || {};

            // ── Auto-detect Dataverse info for hierarchy ──
            // Check first whether this is a Premium plan so we can warn if hierarchy fails
            const hasPremiumTasks = allTasks.some(t => {
                const sid = t.creationSource?.contextScenarioId || '';
                return sid === 'com.microsoft.project.plannerIntegration'
                    || sid.includes('plannerIntegration')
                    || sid.includes('projectIntegration');
            });

            let dataverseHierarchy = null;
            let dvHierarchyWarning = null;
            try {
                dataverseHierarchy = await _fetchDataverseHierarchy(allTasks);
                if (!dataverseHierarchy && hasPremiumTasks) {
                    dvHierarchyWarning = 'Planner Premium plan detected but task hierarchy could not be loaded from Dataverse. '
                        + 'Tasks will appear flat. Open browser DevTools (F12 → Console) for details. '
                        + 'An admin may need to grant Dataverse consent for this app.';
                    console.error('[MSGraph] ⚠️ Premium plan hierarchy FAILED to load. Admin consent URL:');
                    console.error(`https://login.microsoftonline.com/organizations/adminconsent?client_id=5c5eccbf-b7fb-4041-b969-44da0d6cf406&redirect_uri=${encodeURIComponent(window.location.origin + window.location.pathname)}`);
                }
            } catch (dvErr) {
                console.warn('[MSGraph] Dataverse hierarchy fetch failed (non-fatal):', dvErr.message);
                if (hasPremiumTasks) {
                    dvHierarchyWarning = `Hierarchy load failed: ${dvErr.message}. Tasks will appear flat.`;
                }
            }

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
                dataverseHierarchy,
                dvHierarchyWarning,
            };
        } catch (err) {
            throw new Error(`getPlanDetails failed: ${err.message}`);
        }
    }

    /**
     * Auto-discover Dataverse URL from task creationSource and fetch hierarchy.
     * Returns a Map<plannerTaskId, { outlineLevel, wbsId, parentTaskId }> or null.
     */
    async function _fetchDataverseHierarchy(tasks) {
        // Find a task with creationSource containing Dataverse info.
        // Match flexibly: exact value OR any ID containing 'plannerIntegration' / 'projectIntegration'
        // (guards against minor Microsoft value changes across tenants/versions)
        const _isPremiumScenario = (sid) => {
            if (!sid) return false;
            return sid === 'com.microsoft.project.plannerIntegration'
                || sid.includes('plannerIntegration')
                || sid.includes('projectIntegration');
        };
        const sampleTask = tasks.find(t =>
            _isPremiumScenario(t.creationSource?.contextScenarioId)
            && t.creationSource?.externalObjectVersion
        );
        if (!sampleTask) {
            console.log('[MSGraph] Not a Premium/Project plan — no Dataverse hierarchy available');
            return null;
        }

        // Extract Dataverse org URL from externalObjectVersion
        // Format: "msxrm_org6863b5bd.crm4.dynamics.com_projectId_version"
        const versionStr = sampleTask.creationSource.externalObjectVersion || '';
        const orgMatch = versionStr.match(/msxrm_([\w.]+\.dynamics\.com)/);
        if (!orgMatch) {
            console.warn('[MSGraph] Could not extract Dataverse URL from:', versionStr);
            return null;
        }
        const dataverseUrl = `https://${orgMatch[1]}`;
        console.log('[MSGraph] Dataverse URL discovered:', dataverseUrl);

        // Extract project ID from externalContextId
        // Format: "orgId_projectId"
        const contextId = sampleTask.creationSource.externalContextId || '';
        const projectId = contextId.split('_')[1];
        if (!projectId) {
            console.warn('[MSGraph] Could not extract project ID from:', contextId);
            return null;
        }
        console.log('[MSGraph] Project ID discovered:', projectId);

        // Build plannerTaskId → dataverseTaskId mapping
        const plannerToDataverse = new Map();
        tasks.forEach(t => {
            if (t.creationSource?.externalObjectId) {
                const parts = t.creationSource.externalObjectId.split('|');
                // The task ID might be at parts[2], or parts could be shorter. Handle gracefully:
                const dvTaskId = (parts.length >= 3 ? parts[2] : parts[parts.length - 1])?.toLowerCase();
                if (dvTaskId) {
                    plannerToDataverse.set(t.id, dvTaskId);
                }
            }
        });

        // Try to get a Dataverse token
        let dvToken;
        try {
            const accounts = msalApp.getAllAccounts();
            if (!accounts.length) return null;
            
            const account = accounts[0];
            const tenantId = account.tenantId;
            const dvScopes = [`${dataverseUrl}/user_impersonation`];
            const authority = `https://login.microsoftonline.com/${tenantId}`;
            
            console.log('[MSGraph] Requesting Dataverse token for tenant:', tenantId, 'scope:', dvScopes[0]);
            
            try {
                const resp = await msalApp.acquireTokenSilent({ 
                    scopes: dvScopes, 
                    account,
                    authority,
                });
                dvToken = resp.accessToken;
            } catch (_) {
                // Try popup consent with company tenant authority
                const resp = await msalApp.acquireTokenPopup({ 
                    scopes: dvScopes,
                    authority,
                    loginHint: account.username,
                });
                dvToken = resp.accessToken;
            }
            console.log('[MSGraph] Dataverse token acquired ✓');
        } catch (authErr) {
            console.warn('[MSGraph] Dataverse auth failed:', authErr.message);
            console.info('[MSGraph] To enable hierarchy: Ask your admin to open this URL:');
            console.info(`https://login.microsoftonline.com/organizations/adminconsent?client_id=5c5eccbf-b7fb-4041-b969-44da0d6cf406&redirect_uri=${encodeURIComponent(window.location.origin + window.location.pathname)}`);
            return null;
        }

        // Fetch ALL project data from Dataverse
        try {
            const dvHeaders = {
                Authorization: `Bearer ${dvToken}`,
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                Accept: 'application/json',
                Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
            };

            // 1. Fetch tasks, excluding fields that don't exist in all Dataverse environments (e.g. msdyn_wbsid)
            const tasksUrl = `${dataverseUrl}/api/data/v9.2/msdyn_projecttasks`
                + `?$filter=_msdyn_project_value eq '${projectId}'`
                + `&$select=msdyn_projecttaskid,msdyn_subject,msdyn_outlinelevel,msdyn_displaysequence,`
                + `_msdyn_parenttask_value,msdyn_scheduledstart,msdyn_scheduledend,`
                + `msdyn_duration,msdyn_progress,msdyn_effort,msdyn_description`
                + `&$orderby=msdyn_displaysequence asc`
                + `&$top=500`;

            // 2. Fetch resource assignments
            const assignUrl = `${dataverseUrl}/api/data/v9.2/msdyn_resourceassignments`
                + `?$filter=_msdyn_projectid_value eq '${projectId}'`
                + `&$select=msdyn_resourceassignmentid,_msdyn_taskid_value,_msdyn_bookableresourceid_value,msdyn_plannedwork`
                + `&$top=500`;

            // 3. Fetch team members (for resource names)
            const teamUrl = `${dataverseUrl}/api/data/v9.2/msdyn_projectteams`
                + `?$filter=_msdyn_project_value eq '${projectId}'`
                + `&$select=msdyn_projectteamid,msdyn_name,_msdyn_bookableresourceid_value`
                + `&$top=100`;

            // 4. Fetch project entity for manager + scheduled dates
            const projectEntityUrl = `${dataverseUrl}/api/data/v9.2/msdyn_projects`
                + `?$filter=msdyn_projectid eq '${projectId}'`
                + `&$select=msdyn_subject,_msdyn_projectmanager_value,msdyn_scheduledstart,msdyn_scheduledend`
                + `&$top=1`;

            const [tasksResp, assignResp, teamResp, projEntityResp] = await Promise.all([
                fetch(tasksUrl, { method: 'GET', headers: dvHeaders }),
                fetch(assignUrl, { method: 'GET', headers: dvHeaders }).catch(() => ({ ok: false })),
                fetch(teamUrl,   { method: 'GET', headers: dvHeaders }).catch(() => ({ ok: false })),
                fetch(projectEntityUrl, { method: 'GET', headers: dvHeaders }).catch(() => ({ ok: false })),
            ]);

            if (!tasksResp.ok) {
                const errText = await tasksResp.text().catch(() => '');
                console.warn('[MSGraph] Dataverse tasks query failed:', tasksResp.status, errText);
                return null;
            }

            const tasksData = await tasksResp.json();
            const dvTasks = tasksData.value || [];
            console.log(`[MSGraph] Dataverse returned ${dvTasks.length} tasks with full data`);

            // Parse resource assignments
            let dvAssignments = [];
            if (assignResp.ok) {
                const assignData = await assignResp.json();
                dvAssignments = assignData.value || [];
                console.log(`[MSGraph] Dataverse returned ${dvAssignments.length} resource assignments`);
            }

            // Parse team members (resource name lookup)
            const resourceNameMap = new Map(); // bookableResourceId → name
            if (teamResp.ok) {
                const teamData = await teamResp.json();
                const bookableIds = [];
                (teamData.value || []).forEach(tm => {
                    if (tm._msdyn_bookableresourceid_value) {
                        resourceNameMap.set(tm._msdyn_bookableresourceid_value, tm.msdyn_name || 'Unknown');
                        bookableIds.push(tm._msdyn_bookableresourceid_value);
                    }
                });
                console.log(`[MSGraph] Dataverse returned ${resourceNameMap.size} team members`);

                // Resolve real names from bookableresources → systemuser
                if (bookableIds.length > 0) {
                    try {
                        const brFilter = bookableIds.map(id => `bookableresourceid eq '${id}'`).join(' or ');
                        const brUrl = `${dataverseUrl}/api/data/v9.2/bookableresources`
                            + `?$filter=${brFilter}`
                            + `&$select=bookableresourceid,name,_userid_value`
                            + `&$top=100`;
                        const brResp = await fetch(brUrl, { method: 'GET', headers: dvHeaders }).catch(() => null);
                        if (brResp && brResp.ok) {
                            const brData = await brResp.json();
                            (brData.value || []).forEach(br => {
                                if (br.name && br.name.trim() && !br.name.match(/^Team Member \d+$/i)) {
                                    resourceNameMap.set(br.bookableresourceid, br.name);
                                }
                            });
                            console.log(`[MSGraph] Resolved ${(brData.value || []).length} bookable resource names`);
                        }
                    } catch (brErr) {
                        console.warn('[MSGraph] bookableresources lookup failed:', brErr.message);
                    }
                }
            }

            // Build comprehensive result
            const dvMap = new Map();
            dvTasks.forEach(dvt => {
                const taskId = (dvt.msdyn_projecttaskid || '').toLowerCase();
                dvMap.set(taskId, {
                    outlineLevel: dvt.msdyn_outlinelevel || 1,
                    wbsId: dvt.msdyn_wbsid || '', // use Dataverse WBS directly when available
                    parentTaskId: dvt._msdyn_parenttask_value ? dvt._msdyn_parenttask_value.toLowerCase() : null,
                    subject: dvt.msdyn_subject,
                    scheduledStart: dvt.msdyn_scheduledstart || null,
                    scheduledEnd: dvt.msdyn_scheduledend || null,
                    duration: dvt.msdyn_duration || 0,
                    progress: dvt.msdyn_progress || 0,
                    effort: dvt.msdyn_effort || 0,
                    priority: null,
                    description: dvt.msdyn_description || '',
                    dvPredecessors: [], // populated below from dependencies
                });
            });

            // ── Fetch task dependencies by successor task IDs (chunked) ──
            // Planner Premium does NOT support _msdyn_project_value filter on
            // msdyn_projecttaskdependencies → HTTP 400. Instead we filter by
            // successor task GUIDs we already know, in batches of 30.
            let dvDeps = [];
            const dvTaskIds = [...dvMap.keys()]; // Dataverse task GUIDs (lowercase)
            if (dvTaskIds.length > 0) {
                const DEP_CHUNK = 30;
                const depSelect = '$select=msdyn_projecttaskdependencyid,_msdyn_predecessortask_value,_msdyn_successortask_value,msdyn_linktype';
                const depChunkPromises = [];
                for (let ci = 0; ci < dvTaskIds.length; ci += DEP_CHUNK) {
                    const chunk = dvTaskIds.slice(ci, ci + DEP_CHUNK);
                    const filter = chunk.map(id => `_msdyn_successortask_value eq '${id}'`).join(' or ');
                    const url = `${dataverseUrl}/api/data/v9.2/msdyn_projecttaskdependencies?$filter=${filter}&${depSelect}&$top=500`;
                    depChunkPromises.push(
                        fetch(url, { method: 'GET', headers: dvHeaders })
                            .then(r => r.ok ? r.json() : Promise.resolve({ value: [] }))
                            .then(d => d.value || [])
                            .catch(() => [])
                    );
                }
                const chunkResults = await Promise.all(depChunkPromises);
                dvDeps = chunkResults.flat();
                console.log(`[MSGraph] Dependency chunked fetch: ${dvDeps.length} records from ${depChunkPromises.length} chunk(s)`);
            }

            if (dvDeps.length > 0) {
                const LINK_TYPES = { 192350000: 'FS', 192350001: 'FF', 192350002: 'SS', 192350003: 'SF' };
                let depStored = 0, depSkippedNoSucc = 0, depSkippedNoPred = 0;
                dvDeps.forEach(dep => {
                    // dvMap keys are lowercased — normalise GUIDs to match
                    const successorId   = (dep._msdyn_successortask_value  || '').toLowerCase();
                    const predecessorId = (dep._msdyn_predecessortask_value || '').toLowerCase();
                    if (!successorId || !predecessorId) { depSkippedNoSucc++; return; }
                    if (!dvMap.has(successorId)) {
                        // Successor not in dvMap — task may have been filtered or deleted
                        depSkippedNoSucc++;
                        return;
                    }
                    dvMap.get(successorId).dvPredecessors.push({
                        dvTaskId:   predecessorId,
                        linkType:   LINK_TYPES[dep.msdyn_linktype] || 'FS',
                        lagMinutes: 0, // msdyn_lagduration not fetched (Planner Premium doesn't support it)
                    });
                    depStored++;
                    // Note: predecessorId is stored but not validated against dvMap here.
                    // Validation happens later in the Planner-task-to-ProjectFlow-task pass.
                    if (!dvMap.has(predecessorId)) depSkippedNoPred++;
                });
                console.log(
                    `[MSGraph] Dependency parse: ${dvDeps.length} total → `
                    + `${depStored} stored, ${depSkippedNoSucc} skipped (successor not in task list)`
                    + (depSkippedNoPred > 0
                        ? `, ⚠️ ${depSkippedNoPred} predecessors not in dvMap (cross-project or deleted tasks)`
                        : '')
                );
                if (depStored === 0) {
                    console.warn('[MSGraph] ⚠️ No dependencies stored despite records existing — all successor task GUIDs are unknown.'
                        + ' This usually means Planner task IDs are not linked to Dataverse task GUIDs via creationSource.externalObjectId.');
                }
            } else {
                console.log('[MSGraph] No task dependencies found in Dataverse — plan may have no predecessors set in Project for the Web.');
            }

            // Auto-generate WBS IDs from parent-child tree
            _computeWbsIds(dvMap);

            // Build task → assigned resources mapping
            const taskResources = new Map(); // dvTaskId → [{ name, resourceId }]
            dvAssignments.forEach(a => {
                const taskId = a._msdyn_taskid_value;
                const resId = a._msdyn_bookableresourceid_value;
                if (taskId && resId) {
                    if (!taskResources.has(taskId)) taskResources.set(taskId, []);
                    taskResources.get(taskId).push({
                        name: resourceNameMap.get(resId) || resId.substring(0, 8),
                        resourceId: resId,
                    });
                }
            });

            // Map planner task IDs to full Dataverse info
            const result = new Map();
            plannerToDataverse.forEach((dvId, plannerTaskId) => {
                const info = dvMap.get(dvId);
                if (info) {
                    info.resources = taskResources.get(dvId) || [];
                    result.set(plannerTaskId, info);
                }
            });

            console.log(`[MSGraph] Mapped full Dataverse data for ${result.size} tasks`);
            return result;
        } catch (fetchErr) {
            console.warn('[MSGraph] Dataverse fetch error:', fetchErr.message);
            throw fetchErr;
        }
    }

    // ============================================================================
    // WBS ID GENERATOR (from parent-child tree)
    // ============================================================================

    /**
     * Compute WBS IDs from parent-child relationships when msdyn_wbsid is unavailable.
     * If most tasks already have wbsId from Dataverse, use them directly and only
     * compute missing ones. Mutates dvMap entries in-place.
     * @param {Map} dvMap - taskId → { outlineLevel, parentTaskId, wbsId, ... }
     */
    function _computeWbsIds(dvMap) {
        // Check how many tasks already have a WBS from Dataverse (msdyn_wbsid)
        let hasWbsCount = 0;
        dvMap.forEach(info => { if (info.wbsId) hasWbsCount++; });

        if (hasWbsCount > 0) {
            // Use the Dataverse WBS IDs directly — they are the source of truth.
            // Just derive outlineLevel from the WBS string and fill any missing ones.
            dvMap.forEach(info => {
                if (info.wbsId) {
                    info.outlineLevel = info.wbsId.split('.').length;
                }
            });

            // For any tasks without a WBS ID, fall back to parent-derived position
            const missing = [];
            dvMap.forEach((info, taskId) => { if (!info.wbsId) missing.push(taskId); });
            if (missing.length > 0) {
                missing.forEach(taskId => {
                    const info = dvMap.get(taskId);
                    // Use parent WBS + a high suffix so they sort to end of their group
                    if (info.parentTaskId) {
                        const parent = dvMap.get(info.parentTaskId);
                        if (parent?.wbsId) {
                            info.wbsId = `${parent.wbsId}.999`;
                            info.outlineLevel = parent.outlineLevel + 1;
                        }
                    }
                });
                console.warn(`[MSGraph] ${missing.length} tasks had no msdyn_wbsid — appended at end of parent`);
            }

            console.log(`[MSGraph] Using Dataverse WBS IDs for ${hasWbsCount}/${dvMap.size} tasks`);
            return;
        }

        // ── Fallback: compute WBS from parent-child tree ──
        // Group children by parent
        const childrenOf = new Map(); // parentId → [taskId, ...]
        const roots = [];
        dvMap.forEach((info, taskId) => {
            if (info.parentTaskId) {
                if (!childrenOf.has(info.parentTaskId)) childrenOf.set(info.parentTaskId, []);
                childrenOf.get(info.parentTaskId).push(taskId);
            } else {
                roots.push(taskId);
            }
        });

        // Sort children: by scheduledStart first, then outlineLevel, then subject.
        // Using outlineLevel as a secondary sort avoids alphabetical mis-ordering
        // when sibling tasks share the same start date.
        const sortChildren = (ids) => {
            return ids.sort((a, b) => {
                const ia = dvMap.get(a);
                const ib = dvMap.get(b);
                const seqA = ia?.displaySequence ?? ia?.index ?? 999999;
                const seqB = ib?.displaySequence ?? ib?.index ?? 999999;
                return seqA - seqB;
            });
        };

        // Recursive WBS assignment
        function assignWbs(ids, prefix) {
            sortChildren(ids);
            ids.forEach((id, index) => {
                const wbs = prefix ? `${prefix}.${index + 1}` : `${index + 1}`;
                const info = dvMap.get(id);
                if (info) {
                    info.wbsId = wbs;
                    info.outlineLevel = wbs.split('.').length;
                }
                const children = childrenOf.get(id);
                if (children && children.length > 0) {
                    assignWbs(children, wbs);
                }
            });
        }

        assignWbs(roots, '');
        console.log(`[MSGraph] Computed WBS IDs for ${dvMap.size} tasks (${roots.length} root tasks)`);
    }

    // ============================================================================
    // PURE DATAVERSE IMPORT (no Graph API for task data)
    // ============================================================================

    function _getCachedDataverseUrl() {
        return localStorage.getItem(DV_CACHE_KEY);
    }

    function _cacheDataverseUrl(url) {
        localStorage.setItem(DV_CACHE_KEY, url);
    }

    /**
     * Get Dataverse token for the user's company tenant.
     */
    async function _getDataverseToken(dataverseUrl) {
        const accounts = msalApp.getAllAccounts();
        if (!accounts.length) throw new Error('No authenticated account');
        
        const account = accounts[0];
        const authority = `https://login.microsoftonline.com/${account.tenantId}`;
        const dvScopes = [`${dataverseUrl}/user_impersonation`];
        
        try {
            const resp = await msalApp.acquireTokenSilent({ scopes: dvScopes, account, authority });
            return resp.accessToken;
        } catch (_) {
            const resp = await msalApp.acquireTokenPopup({ scopes: dvScopes, authority, loginHint: account.username });
            return resp.accessToken;
        }
    }

    /**
     * Discover Dataverse URL from any Premium plan's tasks (cached after first call).
     */
    async function discoverDataverseUrl() {
        // Check cache first
        const cached = _getCachedDataverseUrl();
        if (cached) return cached;

        // Fetch one plan's tasks from Beta to find creationSource
        const plans = await getMyPlans();
        if (!plans.length) return null;

        for (const plan of plans.slice(0, 5)) {
            try {
                const tasks = await _fetchAllPages(`/planner/plans/${plan.id}/tasks?$top=5`, true);
                const sample = tasks.find(t =>
                    t.creationSource?.contextScenarioId === 'com.microsoft.project.plannerIntegration'
                    && t.creationSource?.externalObjectVersion
                );
                if (sample) {
                    const match = sample.creationSource.externalObjectVersion.match(/msxrm_([\w.]+\.dynamics\.com)/);
                    if (match) {
                        const url = `https://${match[1]}`;
                        _cacheDataverseUrl(url);
                        console.log('[Dataverse] URL discovered and cached:', url);
                        return url;
                    }
                }
            } catch (_) { continue; }
        }
        return null;
    }

    /**
     * List ALL projects from Dataverse (replaces Graph plan listing).
     */
    async function listDataverseProjects(dataverseUrl) {
        const token = await _getDataverseToken(dataverseUrl);
        const url = `${dataverseUrl}/api/data/v9.2/msdyn_projects`
            + `?$select=msdyn_projectid,msdyn_subject,msdyn_description`
            + `&$orderby=msdyn_subject asc`
            + `&$top=50`;

        const response = await fetch(url, {
            method: 'GET',
            headers: {
                Authorization: `Bearer ${token}`,
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                Accept: 'application/json',
            },
        });

        if (!response.ok) {
            const err = await response.text().catch(() => '');
            throw new Error(`Dataverse projects query failed: ${response.status} ${err}`);
        }

        const data = await response.json();
        return (data.value || []).map(p => ({
            id: p.msdyn_projectid,
            title: p.msdyn_subject,
            startDate: p.msdyn_scheduledstart,
            endDate: p.msdyn_scheduledend,
            description: p.msdyn_description || '',
        }));
    }

    /**
     * Import a project entirely from Dataverse — returns a ProjectFlow project object.
     */
    async function importFromDataverse(dataverseUrl, projectId, projectTitle) {
        const token = await _getDataverseToken(dataverseUrl);
        const dvHeaders = {
            Authorization: `Bearer ${token}`,
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            Accept: 'application/json',
            Prefer: 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"',
        };

        // ── Fetch tasks, assignments, team + project entity ──
        // Schema variants: 
        //   Project Ops:   msdyn_scheduleddurationminutes, msdyn_plannedcost, etc.
        //   Planner Prem:  msdyn_duration, msdyn_effort, etc.
        // Strategy: Probe 1 task to discover available fields.
        let taskFields = ['msdyn_projecttaskid','msdyn_subject','msdyn_outlinelevel','msdyn_displaysequence','_msdyn_parenttask_value','msdyn_scheduledstart','msdyn_scheduledend','msdyn_progress','msdyn_effort','msdyn_description'];
        
        try {
            const taskProbe = await fetch(`${dataverseUrl}/api/data/v9.2/msdyn_projecttasks?$top=1`, { method: 'GET', headers: dvHeaders });
            if (taskProbe.ok) {
                const probeData = await taskProbe.json();
                const sample = (probeData.value || [])[0];
                if (sample) {
                    const keys = Object.keys(sample);
                    // Add optional fields only if they exist
                    if (keys.includes('msdyn_scheduleddurationminutes')) taskFields.push('msdyn_scheduleddurationminutes');
                    if (keys.includes('msdyn_duration'))                 taskFields.push('msdyn_duration');
                    if (keys.includes('msdyn_effortcompleted'))         taskFields.push('msdyn_effortcompleted');
                    if (keys.includes('msdyn_effortremaining'))         taskFields.push('msdyn_effortremaining');
                    if (keys.includes('msdyn_plannedcost'))             taskFields.push('msdyn_plannedcost');
                    if (keys.includes('msdyn_actualcost'))              taskFields.push('msdyn_actualcost');
                    if (keys.includes('msdyn_iscritical'))              taskFields.push('msdyn_iscritical');
                    if (keys.includes('msdyn_ismilestone'))             taskFields.push('msdyn_ismilestone');
                    if (keys.includes('msdyn_wbsid'))                   taskFields.push('msdyn_wbsid');
                }
            }
        } catch (e) { console.warn('[Dataverse] Task probe failed:', e.message); }

        const [tasksResp, assignResp, teamResp, projEntityResp] = await Promise.all([
            fetch(`${dataverseUrl}/api/data/v9.2/msdyn_projecttasks`
                + `?$filter=_msdyn_project_value eq '${projectId}'`
                + `&$select=${[...new Set(taskFields)].join(',')}`
                + `&$orderby=msdyn_displaysequence asc&$top=500`,
                { method: 'GET', headers: dvHeaders }),
            fetch(`${dataverseUrl}/api/data/v9.2/msdyn_resourceassignments`
                + `?$filter=_msdyn_projectid_value eq '${projectId}'`
                + `&$select=msdyn_resourceassignmentid,_msdyn_taskid_value,_msdyn_bookableresourceid_value,msdyn_plannedwork`
                + `&$top=500`,
                { method: 'GET', headers: dvHeaders }).catch(() => ({ ok: false })),
            fetch(`${dataverseUrl}/api/data/v9.2/msdyn_projectteams`
                + `?$filter=_msdyn_project_value eq '${projectId}'`
                + `&$select=msdyn_projectteamid,msdyn_name,_msdyn_bookableresourceid_value`
                + `&$top=100`,
                { method: 'GET', headers: dvHeaders }).catch(() => ({ ok: false })),
            fetch(`${dataverseUrl}/api/data/v9.2/msdyn_projects`
                + `?$filter=msdyn_projectid eq '${projectId}'`
                + `&$select=msdyn_projectid,msdyn_subject,_msdyn_projectmanager_value`
                + `&$top=1`,
                { method: 'GET', headers: dvHeaders }).catch(() => ({ ok: false }))
        ]);

        if (!tasksResp.ok) {
            const err = await tasksResp.text().catch(() => '');
            throw new Error(`Dataverse tasks query failed: ${tasksResp.status} ${err}`);
        }

        const dvTasks = (await tasksResp.json()).value || [];
        const dvAssignments = assignResp.ok ? ((await assignResp.json()).value || []) : [];
        const dvTeam = teamResp.ok ? ((await teamResp.json()).value || []) : [];

        // ── Fetch task dependencies (diagnostic-first, no assumed field names) ──
        // Probe confirmed correct Dataverse field names:
        //   _msdyn_predecessortask_value, _msdyn_successortask_value
        //   msdyn_projecttaskdependencylinktype  (NOT msdyn_linktype — that doesn't exist)
        //   msdyn_projecttaskdependencylinklaginseconds  (lag in seconds)
        //   _msdyn_project_value  (can filter by project ID)
        const dvDepsAll = await (async () => {
            const taskGuidSet = new Set(
                dvTasks.map(t => (t.msdyn_projecttaskid || '').toLowerCase()).filter(Boolean)
            );
            if (!taskGuidSet.size) return [];

            const baseUrl = `${dataverseUrl}/api/data/v9.2/msdyn_projecttaskdependencies`;
            const minHeaders = {
                Authorization: dvHeaders.Authorization,
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                Accept: 'application/json',
            };

            // Step 1: probe — fetch 1 record with no $select to discover field names
            let predField = '_msdyn_predecessortask_value';        // confirmed by probe
            let succField = '_msdyn_successortask_value';          // confirmed by probe
            let linkField = 'msdyn_projecttaskdependencylinktype'; // confirmed by probe (NOT msdyn_linktype)
            let lagField  = 'msdyn_projecttaskdependencylinklaginseconds'; // lag in seconds

            try {
                const probeResp = await fetch(`${baseUrl}?$top=1`, { method: 'GET', headers: minHeaders });
                if (!probeResp.ok) {
                    console.warn(`[Dataverse] Dependency table probe failed (HTTP ${probeResp.status}) — entity may not be accessible in this tenant`);
                    return [];
                }
                const probeData = await probeResp.json();
                const sample = (probeData.value || [])[0];
                if (sample) {
                    // Log real field names for debugging
                    const keys = Object.keys(sample).filter(k => !k.startsWith('@'));
                    console.log('[Dataverse] Dependency entity fields:', keys.join(', '));

                    // Auto-detect predecessor/successor/link field name variants
                    if (keys.includes('_msdyn_predecessortaskid_value'))           predField = '_msdyn_predecessortaskid_value';
                    if (keys.includes('_msdyn_predecessortask_value'))             predField = '_msdyn_predecessortask_value';
                    if (keys.includes('_msdyn_successortaskid_value'))             succField = '_msdyn_successortaskid_value';
                    if (keys.includes('_msdyn_successortask_value'))               succField = '_msdyn_successortask_value';
                    if (keys.includes('msdyn_projecttaskdependencylinktype'))      linkField = 'msdyn_projecttaskdependencylinktype';
                    else if (keys.includes('msdyn_linktype'))                      linkField = 'msdyn_linktype';
                    if (keys.includes('msdyn_projecttaskdependencylinklaginseconds')) lagField = 'msdyn_projecttaskdependencylinklaginseconds';
                    else if (keys.includes('msdyn_projecttaskdependencylinklag'))  lagField = 'msdyn_projecttaskdependencylinklag';
                    console.log(`[Dataverse] Using dep fields: pred=${predField}, succ=${succField}, link=${linkField}`);
                } else {
                    console.log('[Dataverse] Dependency table accessible but empty — no predecessors defined');
                    return [];
                }
            } catch (e) {
                console.warn('[Dataverse] Dependency probe error:', e.message);
                return [];
            }

            // Step 2: fetch this project's deps with confirmed field names + project filter
            try {
                const depUrl = `${baseUrl}`
                    + `?$filter=_msdyn_project_value eq '${projectId}'`
                    + `&$select=msdyn_projecttaskdependencyid,${predField},${succField},${linkField},${lagField}`
                    + `&$top=500`;
                const resp = await fetch(depUrl, { method: 'GET', headers: minHeaders });
                if (!resp.ok) {
                    console.warn(`[Dataverse] Dependency fetch failed (HTTP ${resp.status}) — falling back to no-filter fetch`);
                    // Fallback: fetch all without filter, filter client-side
                    const fbResp = await fetch(`${baseUrl}?$select=msdyn_projecttaskdependencyid,${predField},${succField},${linkField},${lagField}&$top=500`,
                        { method: 'GET', headers: minHeaders });
                    if (!fbResp || !fbResp.ok) return [];
                    const fbData = await fbResp.json();
                    const fbDeps = (fbData.value || []).filter(d => taskGuidSet.has((d[succField] || '').toLowerCase()));
                    console.log(`[Dataverse] Fallback: ${fbDeps.length} deps for this project`);
                    return fbDeps.map(d => normalize(d, predField, succField, linkField, lagField));
                }
                const data = await resp.json();
                const deps = data.value || [];
                console.log(`[Dataverse] Dependency fetch: ${deps.length} for this project (filtered by project ID)`);
                return deps.map(d => normalize(d, predField, succField, linkField, lagField));
            } catch (e) {
                console.warn('[Dataverse] Dependency fetch error:', e.message);
                return [];
            }

            function normalize(d, pf, sf, lf, lagf) {
                return {
                    msdyn_projecttaskdependencyid: d.msdyn_projecttaskdependencyid,
                    _msdyn_predecessortask_value:  (d[pf]   || '').toLowerCase(),
                    _msdyn_successortask_value:    (d[sf]   || '').toLowerCase(),
                    msdyn_linktype:                d[lf],
                    lagSeconds:                    Number(d[lagf]) || 0,
                };
            }
        })();
        const projEntityData = projEntityResp.ok ? ((await projEntityResp.json()).value || []) : [];
        const projEntity = projEntityData[0] || null;
        const _projectManagerResourceId = projEntity ? projEntity['_msdyn_projectmanager_value'] : null;

        console.log(`[Dataverse] Import: ${dvTasks.length} tasks, ${dvAssignments.length} assignments, ${dvTeam.length} team members`);

        // Build resource name lookup
        const resNameMap = new Map();
        const bookableIds = [];
        dvTeam.forEach(tm => {
            if (tm._msdyn_bookableresourceid_value) {
                resNameMap.set(tm._msdyn_bookableresourceid_value, tm.msdyn_name || 'Unknown');
                bookableIds.push(tm._msdyn_bookableresourceid_value);
            }
        });

        // Resolve real names from bookableresources
        if (bookableIds.length > 0) {
            try {
                const brFilter = bookableIds.map(id => `bookableresourceid eq '${id}'`).join(' or ');
                const brUrl = `${dataverseUrl}/api/data/v9.2/bookableresources`
                    + `?$filter=${brFilter}`
                    + `&$select=bookableresourceid,name,_userid_value`
                    + `&$top=100`;
                const brResp = await fetch(brUrl, { method: 'GET', headers: dvHeaders }).catch(() => null);
                if (brResp && brResp.ok) {
                    const brData = await brResp.json();
                    (brData.value || []).forEach(br => {
                        if (br.name && br.name.trim() && !br.name.match(/^Team Member \d+$/i)) {
                            resNameMap.set(br.bookableresourceid, br.name);
                        }
                    });
                    console.log(`[Dataverse] Resolved ${(brData.value || []).length} bookable resource real names`);
                }
            } catch (brErr) {
                console.warn('[Dataverse] bookableresources lookup failed:', brErr.message);
            }
        }

        // Build task → resources mapping
        const taskResMap = new Map();
        dvAssignments.forEach(a => {
            const tid = a._msdyn_taskid_value;
            const rid = a._msdyn_bookableresourceid_value;
            if (tid && rid) {
                if (!taskResMap.has(tid)) taskResMap.set(tid, []);
                taskResMap.get(tid).push(resNameMap.get(rid) || rid.substring(0, 8));
            }
        });

        // Detect parent tasks
        const parentIds = new Set();
        dvTasks.forEach(t => { if (t._msdyn_parenttask_value) parentIds.add(t._msdyn_parenttask_value); });

        // Build a dvMap — use msdyn_wbsid directly when available
        const dvMap = new Map();
        let idx = 0;
        dvTasks.forEach(t => {
            dvMap.set(t.msdyn_projecttaskid, {
                index: idx++,
                displaySequence: t.msdyn_displaysequence || 0,
                outlineLevel: t.msdyn_outlinelevel || 1,
                wbsId: t.msdyn_wbsid || '',
                parentTaskId: t._msdyn_parenttask_value || null,
                subject: t.msdyn_subject,
                scheduledStart: t.msdyn_scheduledstart || null,
                scheduledEnd: t.msdyn_scheduledend || null,
            });
        });
        _computeWbsIds(dvMap);

        // Sort tasks by computed WBS order
        const sortedDvTasks = [...dvTasks].sort((a, b) => {
            const wa = dvMap.get(a.msdyn_projecttaskid)?.wbsId || '';
            const wb = dvMap.get(b.msdyn_projecttaskid)?.wbsId || '';
            const aParts = wa.split('.').map(Number);
            const bParts = wb.split('.').map(Number);
            for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
                const av = aParts[i] || 0;
                const bv = bParts[i] || 0;
                if (av !== bv) return av - bv;
            }
            return 0;
        });

        const resourceSet = new Map();
        let resUid = 1;
        resNameMap.forEach((name, id) => {
            const u = resUid++;
            resourceSet.set(id, { uid: u, id: u, name, maxUnits: 100 });
        });

        // Resolve project manager name from resNameMap (built from bookableresources)
        let projectManagerName = null;
        if (_projectManagerResourceId) {
            projectManagerName = resNameMap.get(_projectManagerResourceId) || null;
            if (!projectManagerName) {
                // Manager may be a system user not in team — try systemusers lookup
                try {
                    const suResp = await fetch(
                        `${dataverseUrl}/api/data/v9.2/systemusers`
                            + `?$filter=systemuserid eq '${_projectManagerResourceId}'`
                            + `&$select=fullname&$top=1`,
                        { method: 'GET', headers: dvHeaders }
                    ).catch(() => null);
                    if (suResp && suResp.ok) {
                        const suData = await suResp.json();
                        const su = (suData.value || [])[0];
                        if (su && su.fullname) projectManagerName = su.fullname;
                    }
                } catch (_) {}
            }
            if (projectManagerName) console.log(`[Dataverse] Project manager: ${projectManagerName}`);
        }

        // Build project
        const today = new Date().toISOString().split('T')[0];
        const project = {
            id: projectId,
            name: projectTitle || 'Dataverse Project',
            startDate: null,
            finishDate: null,
            minutesPerDay: 480,
            minutesPerWeek: 2400,
            daysPerMonth: 20,
            tasks: [],
            resources: [...resourceSet.values()],
            assignments: [],
            projectManager: projectManagerName || '',
            _source: 'dataverse',
            _dataverseProjectId: projectId,
        };

        const toLocalYYYYMMDD = (iso) => {
            if (!iso) return null;
            const d = new Date(iso);
            return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
        };

        let uid = 1;
        sortedDvTasks.forEach(t => {
            const taskUid = uid++;
            const dvInfo = dvMap.get(t.msdyn_projecttaskid);
            const start = toLocalYYYYMMDD(t.msdyn_scheduledstart) || today;
            const finish = toLocalYYYYMMDD(t.msdyn_scheduledend) || start;
            
            // Prefer scheduleddurationminutes if available, else duration, else calculate from dates.
            // Guard: value must be a finite positive number; anything else falls through to date diff.
            let durationMin = t.msdyn_scheduleddurationminutes;
            if (durationMin === undefined || durationMin === null) durationMin = t.msdyn_duration;
            if (typeof durationMin !== 'number' || !isFinite(durationMin) || durationMin < 0) durationMin = null;
            const dur = _dvDurationDays(durationMin, start, finish);
            
            // Progress is usually 0-100 in Dataverse, not 0-1. Guard against 5000%.
            let pct = t.msdyn_progress || 0;
            if (pct > 0 && pct <= 1) pct = Math.round(pct * 100); 
            else pct = Math.round(pct);
            const isSummary = parentIds.has(t.msdyn_projecttaskid);
            const resourceNames = taskResMap.get(t.msdyn_projecttaskid) || [];

            // Build assignments
            const taskAssignments = dvAssignments.filter(a => a._msdyn_taskid_value === t.msdyn_projecttaskid);
            taskAssignments.forEach(a => {
                const res = resourceSet.get(a._msdyn_bookableresourceid_value);
                if (res) {
                    project.assignments.push({ taskUID: taskUid, resourceUID: res.uid, units: 100 });
                }
            });

            project.tasks.push({
                uid: taskUid,
                id: taskUid,
                name: t.msdyn_subject || 'Untitled',
                wbs: dvInfo?.wbsId || t.msdyn_wbsid || undefined,
                outlineLevel: dvInfo?.outlineLevel || t.msdyn_outlinelevel || 1,
                summary: isSummary,
                milestone: t.msdyn_ismilestone || dur === 0,
                start,
                finish,
                durationDays: dur,
                percentComplete: pct,
                resourceNames,
                tags: [],
                notes: t.msdyn_description || '',
                predecessors: [],
                isExpanded: true,
                isVisible: true,
                isCritical: !!t.msdyn_iscritical,
                
                // Financials & Effort
                plannedCost: t.msdyn_plannedcost || 0,
                actualCost: t.msdyn_actualcost || 0,
                plannedHours: t.msdyn_effort || 0,
                actualHours: t.msdyn_effortcompleted || 0,
                remainingHours: t.msdyn_effortremaining || 0,
                
                _dataverseTaskId: (t.msdyn_projecttaskid || '').toLowerCase(),
            });

            // Track project date range
            if (!project.startDate || start < project.startDate) project.startDate = start;
            if (!project.finishDate || finish > project.finishDate) project.finishDate = finish;
        });

        // ── Wire task dependencies into predecessors[] ──
        if (dvDepsAll.length > 0) {
            project._dependenciesAvailable = true;
            const LINK_TYPES    = { 192350000: 'FS', 192350001: 'FF', 192350002: 'SS', 192350003: 'SF' };
            const TYPE_CODES_DV = { FS: 1, FF: 0, SS: 3, SF: 2 };

            // Map: DV task GUID (lowercase) → ProjectFlow UID
            const dvIdToUid = new Map();
            project.tasks.forEach(t => { if (t._dataverseTaskId) dvIdToUid.set(t._dataverseTaskId, t.uid); });

            dvDepsAll.forEach(dep => {
                // GUIDs already lowercased by normalize()
                const successorId   = dep._msdyn_successortask_value  || '';
                const predecessorId = dep._msdyn_predecessortask_value || '';
                if (!dvIdToUid.has(successorId) || !dvIdToUid.has(predecessorId)) {
                    console.warn('[Dataverse] Dependency references unknown task(s):',
                        dep.msdyn_projecttaskdependencyid,
                        '| successor found:', dvIdToUid.has(successorId),
                        '| predecessor found:', dvIdToUid.has(predecessorId));
                    return;
                }
                const successorUid   = dvIdToUid.get(successorId);
                const predecessorUid = dvIdToUid.get(predecessorId);
                const task = project.tasks.find(t => t.uid === successorUid);
                if (task) {
                    const linkName = LINK_TYPES[dep.msdyn_linktype] || 'FS';
                    // Real lag from Dataverse (lagSeconds confirmed in schema)
                    const lagDays = dep.lagSeconds ? dep.lagSeconds / 60 / (project.minutesPerDay || 480) : 0;
                    task.predecessors.push({
                        predecessorUID: predecessorUid,
                        type: TYPE_CODES_DV[linkName] ?? 1,
                        typeName: linkName,
                        lag: Math.round(lagDays * 100) / 100,
                    });
                }
            });

            const totalDeps = project.tasks.reduce((s, t) => s + t.predecessors.length, 0);
            console.log(`[Dataverse] Resolved ${totalDeps} predecessor links from ${dvDepsAll.length} dependency records`);
            if (totalDeps === 0 && dvDepsAll.length > 0) {
                console.warn('[Dataverse] Deps fetched but none resolved — GUID mismatch between dependency records and task list');
            }
        } else {
            project._dependenciesAvailable = false;
            console.log('[Dataverse] No dependency records found — plan may have no predecessors defined in Project for the Web');
        }

        return project;
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
    async function _batchCall(requests, retryCount = 0) {
        const token = await _getAccessToken();
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 30000);

        try {
            const response = await fetch(`${GRAPH_ENDPOINT}/$batch`, {
                method:  'POST',
                headers: {
                    Authorization:  `Bearer ${token}`,
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ requests }),
                signal: controller.signal,
            });
            clearTimeout(timeoutId);

            // Handle 429 Rate Limit on $batch endpoint
            if (response.status === 429) {
            const MAX_RETRIES = 3;
            if (retryCount >= MAX_RETRIES) {
                throw new Error(`$batch rate limit exceeded after ${MAX_RETRIES} retries.`);
            }
            const retryAfter = parseInt(response.headers.get('Retry-After') || '2', 10);
            const delay = Math.max(retryAfter * 1000, Math.pow(2, retryCount + 1) * 1000);
            console.warn(`[MSGraph] $batch rate limited. Retry ${retryCount + 1}/${MAX_RETRIES} in ${delay}ms`);
            await new Promise(r => setTimeout(r, delay));
            return _batchCall(requests, retryCount + 1);
        }

        if (!response.ok) {
            const errText = await response.text().catch(() => response.statusText);
            throw new Error(`Graph $batch error ${response.status}: ${errText}`);
        }
        const result = await response.json();
        // Return responses sorted by id so callers can zip with original requests
        const map = Object.fromEntries((result.responses || []).map(r => [r.id, r]));
        return requests.map(req => map[req.id] || { id: req.id, status: 500, body: null });
        } catch (err) {
            clearTimeout(timeoutId);
            throw err;
        }
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
    function plannerToProject(plan, buckets, tasks, taskDetailsMap, userIdToName = {}, categoryDescriptions = {}, dataverseHierarchy = null) {
        const project = {
            id: plan.id,
            name: plan.title,
            owner: plan.owner,
            createdBy: plan.createdBy,
            tasks: [],
            resources: [],
            assignments: [],
            _plannerId: plan.id,
        };

        const today = new Date().toISOString().split('T')[0];
        let uidSeq = 1;

        // ── Build resources array from all unique assignees ──
        const resourceSet = new Map();
        let resIdSeq = 1;
        tasks.forEach(task => {
            if (task.assignments && typeof task.assignments === 'object') {
                Object.keys(task.assignments).forEach(userId => {
                    if (!resourceSet.has(userId)) {
                        const name = userIdToName[userId] || _userCache.get(userId) || userId.substring(0, 8) + '…';
                        resourceSet.set(userId, {
                            uid:  resIdSeq,
                            id:   resIdSeq,
                            name: name,
                            type: 'Work',
                            maxUnits: 100,
                            costPerHour: 0,
                            _plannerUserId: userId,
                        });
                        resIdSeq++;
                    }
                });
            }
        });
        project.resources = [...resourceSet.values()];

        // ── Sort tasks: use Dataverse WBS order if available, else chronological ──
        let sortedTasks;
        const hasHierarchy = dataverseHierarchy && dataverseHierarchy.size > 0;
        
        if (hasHierarchy) {
            console.log(`[MSGraph] Using Dataverse hierarchy for ${dataverseHierarchy.size} tasks`);
            // Sort by WBS ID (e.g., "1", "1.1", "1.2", "2", "2.1") 
            sortedTasks = [...tasks].sort((a, b) => {
                const extA = a.creationSource?.externalObjectId ? a.creationSource.externalObjectId.split('|') : [];
                const dvIdA = extA.length >= 3 ? extA[2].toLowerCase() : (extA.length > 0 ? extA[extA.length - 1].toLowerCase() : null);
                
                const extB = b.creationSource?.externalObjectId ? b.creationSource.externalObjectId.split('|') : [];
                const dvIdB = extB.length >= 3 ? extB[2].toLowerCase() : (extB.length > 0 ? extB[extB.length - 1].toLowerCase() : null);

                const ha = dvIdA ? dataverseHierarchy.get(dvIdA) : null;
                const hb = dvIdB ? dataverseHierarchy.get(dvIdB) : null;
                
                if (!ha && !hb) return 0;
                if (!ha) return 1;
                if (!hb) return -1;
                // Compare WBS IDs numerically: "1.2" vs "1.10" → [1,2] vs [1,10]
                const aParts = (ha.wbsId || '').split('.').map(Number);
                const bParts = (hb.wbsId || '').split('.').map(Number);
                for (let i = 0; i < Math.max(aParts.length, bParts.length); i++) {
                    const av = aParts[i] || 0;
                    const bv = bParts[i] || 0;
                    if (av !== bv) return av - bv;
                }
                return 0;
            });
        } else {
            sortedTasks = [...tasks].sort((a, b) => {
                const sa = a.startDateTime || '9999';
                const sb = b.startDateTime || '9999';
                if (sa !== sb) return sa.localeCompare(sb);
                const da = a.dueDateTime || '9999';
                const db = b.dueDateTime || '9999';
                return da.localeCompare(db);
            });
        }

        // Build bucket name lookup
        const bucketNameMap = new Map();
        buckets.forEach(b => bucketNameMap.set(b.id, b.name));

        // Detect which Dataverse tasks are parents (have children)
        const parentTaskIds = new Set();
        if (hasHierarchy) {
            dataverseHierarchy.forEach(info => {
                if (info.parentTaskId) parentTaskIds.add(info.parentTaskId);
            });
        }

        // ── Import each task ──
        sortedTasks.forEach(task => {
            const details   = taskDetailsMap[task.id] || {};
            let startDate = task.startDateTime ? task.startDateTime.split('T')[0] : today;
            let finishDate = task.dueDateTime  ? task.dueDateTime.split('T')[0]  : _addDays(startDate, 1);
            const dur = _workingDaysBetween(startDate, finishDate);

            // Resolve assignment user IDs → display names
            const resourceNames = [];
            if (task.assignments && typeof task.assignments === 'object') {
                Object.entries(task.assignments).forEach(([userId, assignment]) => {
                    const name = userIdToName[userId]
                        || _userCache.get(userId)
                        || assignment?.assignedBy?.user?.displayName
                        || assignment?.createdBy?.user?.displayName
                        || userId.substring(0, 8) + '…';
                    resourceNames.push(name);
                });
            }

            // Build assignments array entries for resource linking
            const taskUid = uidSeq++;
            if (task.assignments && typeof task.assignments === 'object') {
                Object.keys(task.assignments).forEach(userId => {
                    const res = resourceSet.get(userId);
                    if (res) {
                        project.assignments.push({
                            taskUID:     taskUid,
                            resourceUID: res.uid,
                            units:       100,
                        });
                    }
                });
            }

            // Map appliedCategories → real label names
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

            // ── Apply Dataverse data if available ──
            let outlineLevel = 1;
            let isSummary = false;
            let wbsId = '';
            const dvInfo = hasHierarchy ? dataverseHierarchy.get(task.id) : null;
            
            if (dvInfo) {
                outlineLevel = dvInfo.outlineLevel || 1;
                wbsId = dvInfo.wbsId || '';
                
                // Check if this task's Dataverse ID appears as a parent
                // NOTE: parentTaskIds stores lowercase UUIDs — must lowercase before lookup
                const dvTaskId = task.creationSource?.externalObjectId?.split('|')[2]?.toLowerCase();
                isSummary = dvTaskId ? parentTaskIds.has(dvTaskId) : false;

                // Override dates from Dataverse (more accurate for premium plans)
                if (dvInfo.scheduledStart) {
                    startDate = dvInfo.scheduledStart.split('T')[0];
                }
                if (dvInfo.scheduledEnd) {
                    finishDate = dvInfo.scheduledEnd.split('T')[0];
                }

                // Override progress from Dataverse
                if (dvInfo.progress !== undefined && dvInfo.progress !== null) {
                    pct = Math.round(dvInfo.progress * 100);
                }

                // Override resources from Dataverse
                if (dvInfo.resources && dvInfo.resources.length > 0) {
                    resourceNames.length = 0; // Clear Graph-based names
                    dvInfo.resources.forEach(r => resourceNames.push(r.name));
                }

                // Add Dataverse description if Graph didn't have one
                if (dvInfo.description && !notes) {
                    notes = dvInfo.description;
                }
            }

            // Recalculate duration — prefer Dataverse msdyn_duration (minutes → days),
            // otherwise count working days (Mon–Fri) between the final start/finish.
            const finalDur = _dvDurationDays(dvInfo?.duration, startDate, finishDate);

            const extParts = task.creationSource?.externalObjectId ? task.creationSource.externalObjectId.split('|') : [];
            const dvTaskIdStr = extParts.length >= 3 ? extParts[2] : (extParts.length > 0 ? extParts[extParts.length - 1] : null);

            project.tasks.push({
                uid:            taskUid,
                id:             taskUid,
                name:           task.title,
                wbs:            wbsId || undefined,
                outlineLevel:   outlineLevel,
                summary:        isSummary,
                milestone:      false,
                start:          startDate,
                finish:         finishDate,
                durationDays:   finalDur,
                percentComplete: pct,
                priority:       priorityLabel,
                priorityNum:    priorityNum,
                resourceNames,
                tags,
                notes,
                predecessors:   [],
                isExpanded:     true,
                isVisible:      true,
                _plannerId:       task.id,
                _plannerEtag:     task['@odata.etag'],
                _plannerBucketId: task.bucketId,
                _plannerBucketName: bucketNameMap.get(task.bucketId) || '',
                _plannerAssigneeIds: Object.keys(task.assignments || {}),
                _dvTaskId:        dvInfo && dvTaskIdStr ? dvTaskIdStr.toLowerCase() : null,
            });
        });

        // ── Resolve Dataverse predecessors into ProjectFlow predecessor UIDs ──
        if (hasHierarchy) {
            // Build DV task ID → ProjectFlow task UID mapping
            // (needed to resolve pred.dvTaskId → UID later)
            const dvIdToUid = new Map();
            project.tasks.forEach(t => {
                if (t._dvTaskId) dvIdToUid.set(t._dvTaskId, t.uid);
            });

            console.log(`[MSGraph] Predecessor resolution: ${dvIdToUid.size} tasks with DV IDs`);

            // ── KEY FIX: dataverseHierarchy is keyed by PLANNER task ID ──
            // (result.set(plannerTaskId, info) in _fetchDataverseHierarchy)
            // Previous code used t._dvTaskId as the lookup key → always undefined.
            // Must use t._plannerId to match the Map structure.
            let dvPredCount = 0;
            project.tasks.forEach(t => {
                if (!t._plannerId) return;
                const dvInfo = dataverseHierarchy.get(t._plannerId); // ← Planner ID, NOT _dvTaskId
                if (dvInfo?.dvPredecessors?.length > 0) dvPredCount += dvInfo.dvPredecessors.length;
            });
            console.log(`[MSGraph] Dataverse dvPredecessors found: ${dvPredCount}`);

            // For each task, look up its Dataverse info by Planner task ID
            project.tasks.forEach(t => {
                if (!t._plannerId) return;
                const dvInfo = dataverseHierarchy.get(t._plannerId); // ← Planner ID, NOT _dvTaskId
                if (!dvInfo || !dvInfo.dvPredecessors || dvInfo.dvPredecessors.length === 0) return;

                dvInfo.dvPredecessors.forEach(pred => {
                    // pred.dvTaskId is a Dataverse GUID → resolve to ProjectFlow UID
                    const predUid = dvIdToUid.get(pred.dvTaskId);
                    if (predUid) {
                        // CPM type codes: FF=0, FS=1, SF=2, SS=3
                        const TYPE_CODES = { FS: 1, FF: 0, SS: 3, SF: 2 };
                        // Convert lag from Dataverse minutes to working days
                        const minutesPerDay = project.minutesPerDay || 480;
                        const lagDays = pred.lagMinutes ? Math.round(pred.lagMinutes / minutesPerDay) : 0;
                        t.predecessors.push({
                            predecessorUID: predUid,
                            type: TYPE_CODES[pred.linkType] ?? 1,
                            typeName: pred.linkType || 'FS',
                            lag: lagDays,
                        });
                    } else {
                        console.warn(`[MSGraph] Predecessor GUID not found in task list: ${pred.dvTaskId}`,
                            '(cross-project dependency or task was deleted)');
                    }
                });
            });

            const totalDeps = project.tasks.reduce((sum, t) => sum + t.predecessors.length, 0);
            console.log(`[MSGraph] Final predecessor count: ${totalDeps}`);

            // Mark whether actual dependency links were resolved
            project._dependenciesAvailable = totalDeps > 0;
            if (totalDeps === 0) {
                console.warn('[MSGraph] No Dataverse task dependencies resolved.'
                    + ' Network/PERT will show tasks as unlinked nodes.'
                    + ' Verify that msdyn_projecttaskdependencies contains records for this project'
                    + ' and that the service account has read access.');
            }
        } else {
            // Dataverse hierarchy was not available — no way to fetch dependencies
            project._dependenciesAvailable = false;
            console.warn('[MSGraph] Dataverse hierarchy unavailable — Network/PERT will show unlinked tasks.'
                + ' This typically means Project Operations is not licensed or Dataverse is not configured for this tenant.');
        }

        return project;
    }

    function _addDays(dateStr, days) {
        const date = new Date(dateStr);
        date.setDate(date.getDate() + days);
        return date.toISOString().split('T')[0];
    }

    /**
     * Count working days (Mon–Fri) between two ISO date strings, inclusive of start.
     * Returns at least 1.
     */
    function _workingDaysBetween(startStr, finishStr) {
        const d1 = new Date(startStr);
        const d2 = new Date(finishStr);
        if (d2 <= d1) return 1;
        const totalDays = Math.round((d2 - d1) / 864e5);
        const fullWeeks = Math.floor(totalDays / 7);
        let workDays = fullWeeks * 5;
        const rem = totalDays % 7;
        const startDay = d1.getDay(); // 0=Sun … 6=Sat
        for (let i = 0; i < rem; i++) {
            const d = (startDay + i) % 7;
            if (d !== 0 && d !== 6) workDays++;
        }
        return Math.max(1, workDays);
    }

    /**
     * Convert a Dataverse msdyn_duration (stored in minutes, 8 h/day) to working days.
     * Falls back to _workingDaysBetween when duration is absent, zero, NaN, or non-numeric.
     *
     * Explicit typeof + isFinite guards prevent NaN propagation into the CPM engine
     * when the Dataverse field contains a null/undefined/string value.
     */
    function _dvDurationDays(durationMinutes, startStr, finishStr) {
        if (typeof durationMinutes === 'number' && isFinite(durationMinutes) && durationMinutes > 0) {
            return Math.max(1, Math.round(durationMinutes / 480)); // 480 min = 8 h
        }
        return _workingDaysBetween(startStr, finishStr);
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
            const { plan, buckets, tasks, categoryDescriptions, dataverseHierarchy, dvHierarchyWarning } = await getPlanDetails(planId);

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
                /* console.log('[MSGraph DEBUG] Assignment analysis:', {
                    tasksWithAssignments: tasks.filter(t => t.assignments && Object.keys(t.assignments).length > 0).length,
                    totalTasks: tasks.length,
                    unresolvedUserIds: allUserIds.size,
                    userIdsFromGroupMembers: Object.keys(userIdToName).length,
                    sampleAssignments: tasks.slice(0, 3).map(t => ({ title: t.title, assignments: t.assignments })),
                }); */
                if (allUserIds.size > 0) {
                    await _resolveUserDisplayNames([...allUserIds]);
                    allUserIds.forEach(id => {
                        const name = _userCache.get(id);
                        if (name) userIdToName[id] = name;
                    });
                    // console.log('[MSGraph DEBUG] After batch resolve, userIdToName:', { ...userIdToName });
                }

                // Populate plan cache so future pulls skip this work
                _planCache.set(planId, { categoryDescriptions, userIdToName: { ...userIdToName } });
            }

            // ── DEBUG: Log what we're passing to plannerToProject ──
            /* console.log('[MSGraph DEBUG] importPlan data:', {
                planTitle: plan.title,
                bucketsCount: buckets.length,
                tasksCount: tasks.length,
                userIdToName,
                sampleTask: tasks[0] ? {
                    title: tasks[0].title,
                    assignments: tasks[0].assignments,
                    bucketId: tasks[0].bucketId,
                } : 'NO TASKS',
                taskAssignmentsSummary: tasks.map(t => ({
                    title: t.title,
                    hasAssignments: !!(t.assignments && Object.keys(t.assignments).length > 0),
                    assignmentCount: t.assignments ? Object.keys(t.assignments).length : 0,
                })),
            }); */

            const project = plannerToProject(plan, buckets, tasks, taskDetailsMap, userIdToName, categoryDescriptions, dataverseHierarchy);

            // Attach hierarchy warning so app.js can show a visible toast to the user
            if (dvHierarchyWarning) {
                project._dvHierarchyWarning = dvHierarchyWarning;
            }

            // ── DEBUG: Log what plannerToProject produced ──
            /* console.log('[MSGraph DEBUG] plannerToProject result:', {
                tasksCount: project.tasks.length,
                resourcesCount: project.resources.length,
                assignmentsCount: project.assignments.length,
                sampleTaskResources: project.tasks.filter(t => !t.summary).slice(0, 3).map(t => ({
                    name: t.name,
                    resourceNames: t.resourceNames,
                })),
                resources: project.resources,
            }); */

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
    /**
     * Import multiple Dataverse projects in sequence.
     */
    async function importMultipleDataverseProjects(dataverseUrl, projectIds, projectTitles, onProgress) {
        const results = [];
        const DELAY_BETWEEN_PLANS = 1000; 

        for (let i = 0; i < projectIds.length; i++) {
            const projectId = projectIds[i];
            const title = projectTitles ? projectTitles[i] : 'Dataverse Project';
            try {
                console.log(`[Dataverse] Portfolio import: project ${i + 1}/${projectIds.length} (${projectId.substring(0, 8)}…)`);
                const proj = await importFromDataverse(dataverseUrl, projectId, title);
                results.push({ planId: projectId, success: true, project: proj });
                console.log(`[Dataverse] Project ${i + 1} imported ✓ — ${proj?.tasks?.length || 0} tasks`);
            } catch (err) {
                console.warn(`[Dataverse] Project ${i + 1} failed: ${err.message}`);
                results.push({ planId: projectId, success: false, error: err.message });
            }
            if (onProgress) onProgress(i + 1, projectIds.length);

            if (i < projectIds.length - 1) {
                await new Promise(r => setTimeout(r, DELAY_BETWEEN_PLANS));
            }
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
    // AUTO-SYNC (default: every 5 minutes)
    // ============================================================================

    let _syncFailCount = 0;
    let _lastSyncTime = null;
    const MAX_SYNC_FAILS = 3;

    // Multi-plan auto-sync: planId → { intervalId, project, failCount }
    const _syncIntervals = new Map();

    /**
     * Start auto-sync for a plan.
     * @param {Object|Function} projectOrGetter  Project object OR () => project getter.
     *   Pass a getter when the caller's `project` variable may be reassigned later,
     *   so the interval always operates on the live object.
     * @param {string} planId
     * @param {number} [intervalMs=300000]  Pull interval in ms (default 5 min)
     */
    function startAutoSync(projectOrGetter, planId, intervalMs = 300000) {
        try {
            // Accept a getter function to avoid stale-reference bugs
            const getProject = typeof projectOrGetter === 'function'
                ? projectOrGetter
                : () => projectOrGetter;

            // Stop existing sync for this specific plan (if any)
            if (_syncIntervals.has(planId)) {
                clearInterval(_syncIntervals.get(planId).intervalId);
                _syncIntervals.delete(planId);
            }

            const entry = { getProject, failCount: 0, intervalId: null };

            entry.intervalId = setInterval(async () => {
                try {
                    const { tasks: changedTasks, isDelta } = await _fetchDeltaTasks(planId);
                    if (changedTasks.length > 0 || !isDelta) {
                        _mergeRemoteChanges(entry.getProject(), changedTasks, isDelta);
                        console.log(`[MSGraph] Auto-sync ✓ [${planId.substring(0,8)}…] — ${isDelta ? changedTasks.length + ' changed' : 'full refresh'}`);
                    } else {
                        console.log(`[MSGraph] Auto-sync ✓ [${planId.substring(0,8)}…] — no changes`);
                    }
                    entry.failCount = 0;
                    _lastSyncTime = new Date();
                    _emitSyncEvent('success', changedTasks.length);
                } catch (err) {
                    entry.failCount++;
                    console.warn(`[MSGraph] Auto-sync error [${planId.substring(0,8)}…] (${entry.failCount}/${MAX_SYNC_FAILS}):`, err.message);
                    _emitSyncEvent('error', err.message);
                    if (entry.failCount >= MAX_SYNC_FAILS) {
                        stopAutoSync(planId);
                        console.error(`[MSGraph] Auto-sync stopped for plan ${planId.substring(0,8)}… after repeated failures`);
                        _emitSyncEvent('stopped', `Plan ${planId.substring(0,8)}… — too many failures`);
                    }
                }
            }, intervalMs);

            _syncIntervals.set(planId, entry);
            console.log(`[MSGraph] Auto-sync started for plan ${planId.substring(0,8)}… — every ${Math.round(intervalMs / 1000)}s (${_syncIntervals.size} plans active)`);
            _emitSyncEvent('started', intervalMs);
            return entry.intervalId;
        } catch (err) {
            throw new Error(`startAutoSync failed: ${err.message}`);
        }
    }

    function stopAutoSync(planId) {
        if (planId) {
            // Stop specific plan
            const entry = _syncIntervals.get(planId);
            if (entry) {
                clearInterval(entry.intervalId);
                _syncIntervals.delete(planId);
                console.log(`[MSGraph] Auto-sync stopped for plan ${planId.substring(0,8)}… (${_syncIntervals.size} remaining)`);
            }
        } else {
            // Stop all syncs
            _syncIntervals.forEach((entry, pid) => {
                clearInterval(entry.intervalId);
            });
            _syncIntervals.clear();
            console.log('[MSGraph] All auto-syncs stopped');
        }
        if (_syncIntervals.size === 0) {
            _emitSyncEvent('stopped', planId ? 'manual' : 'manual-all');
        }
        // Legacy compat
        if (autoSyncInterval) {
            clearInterval(autoSyncInterval);
            autoSyncInterval = null;
        }
    }

    function getSyncStatus() {
        return {
            active: _syncIntervals.size > 0,
            activePlans: _syncIntervals.size,
            lastSync: _lastSyncTime,
            failCount: _syncFailCount,
        };
    }

    function _emitSyncEvent(type, detail) {
        try {
            window.dispatchEvent(new CustomEvent('pf-sync', { detail: { type, detail, time: new Date() } }));
        } catch (_) {}
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

                // Recompute duration in working days
                if (localTask.start && localTask.finish) {
                    localTask.durationDays = _workingDaysBetween(localTask.start, localTask.finish);
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

    function renderSetupWizard(container, onComplete, options = {}) {
        // options.checkImported(planId) → { isImported, storeId, name } | { isImported: false }
        if (!container) {
            throw new Error('renderSetupWizard: container not found.');
        }

        const wizard = document.createElement('div');
        wizard.className = 'ms-graph-wizard';
        wizard.style.cssText = `
            max-width: 500px;
            margin: 0 auto;
            padding: 0;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
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
            try {
                wizard.innerHTML = '';

                // Auto-init with built-in Client ID — always succeeds
                showStatus('⏳ Initializing...', 'info');
                const ready = await _autoInit();

                if (!ready) {
                    showStatus('❌ Failed to initialize. Please refresh the page.', 'error');
                    return;
                }

                console.log('[MSGraph] Wizard: MSAL ready, authenticated=', isAuthenticated());

                // Already authenticated — go straight to plan selection
                if (isAuthenticated()) {
                    showStatus('✅ Already signed in — loading plans...', 'success');
                    await renderPlanSelection();
                } else {
                    // First time — show sign-in button only (no Client ID prompt)
                    wizard.querySelector('.wizard-status')?.remove();
                    renderSignIn();
                }
            } catch (wizardErr) {
                console.error('[MSGraph] Wizard error:', wizardErr);
                showStatus(`❌ Error: ${wizardErr.message}`, 'error');
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

        // ── Step 2: Select Plan(s) (Dataverse Premium Only) ──
        async function renderPlanSelection() {
            const loadingMsg = document.createElement('div');
            loadingMsg.style.cssText = 'text-align:center;padding:20px;color:var(--text-muted,#888);font-size:0.85rem;';
            loadingMsg.innerHTML = '🔄 Discovering Planner Premium environment...';
            wizard.appendChild(loadingMsg);

            try {
                // 1. Discover Dataverse URL (requires at least one premium plan to exist and be accessible via Graph)
                const dataverseUrl = await discoverDataverseUrl();
                if (!dataverseUrl) {
                    loadingMsg.remove();
                    showStatus('No Planner Premium (Dataverse) environment could be automatically discovered. Ensure you have at least one premium plan created in Project for the web.', 'error');
                    return;
                }

                loadingMsg.innerHTML = '🔄 Fetching Premium Projects from Dataverse...';
                const plans = await listDataverseProjects(dataverseUrl);
                loadingMsg.remove();

                if (!plans || plans.length === 0) {
                    showStatus('No projects found in the Dataverse environment.', 'error');
                    return;
                }

                const label = document.createElement('div');
                label.style.cssText = 'font-weight:500;margin-bottom:10px;font-size:0.85rem;color:var(--text-secondary,#a0aec0);display:flex;justify-content:space-between;align-items:center;';
                label.innerHTML = `
                    <span>Select Premium Projects (Dataverse):</span>
                    <span style="font-size:11px;opacity:0.7">${plans.length} project${plans.length !== 1 ? 's' : ''} found</span>
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
                    // Check if this plan is already imported
                    const importStatus = options.checkImported ? options.checkImported(plan.id) : { isImported: false };
                    const alreadyImported = importStatus.isImported;

                    const row = document.createElement('label');
                    row.style.cssText = `
                        display: flex; align-items: center; gap: 10px;
                        padding: 9px 12px; cursor: pointer; font-size: 13px;
                        border-bottom: 1px solid rgba(255,255,255,0.06);
                        color: var(--text-primary, #e2e8f0);
                        transition: background 0.15s;
                        ${alreadyImported ? 'background:rgba(99,102,241,0.07);' : ''}
                    `;
                    row.addEventListener('mouseenter', () => row.style.background = alreadyImported ? 'rgba(99,102,241,0.13)' : 'rgba(255,255,255,0.05)');
                    row.addEventListener('mouseleave', () => row.style.background = alreadyImported ? 'rgba(99,102,241,0.07)' : '');

                    const cb = document.createElement('input');
                    cb.type = 'checkbox';
                    cb.value = plan.id;
                    cb.dataset.title = plan.title;
                    cb.dataset.existingStoreId = alreadyImported ? (importStatus.storeId || '') : '';
                    cb.dataset.isUpdate = alreadyImported ? 'true' : 'false';
                    cb.style.cssText = 'cursor:pointer;flex-shrink:0;accent-color:#6366f1;';
                    cb.addEventListener('change', updateImportBtn);
                    checkboxes.push(cb);

                    const nameSpan = document.createElement('span');
                    nameSpan.textContent = plan.title;
                    nameSpan.style.cssText = 'overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;';

                    row.appendChild(cb);
                    row.appendChild(nameSpan);

                    if (alreadyImported) {
                        const badge = document.createElement('span');
                        badge.textContent = '🔄 Already imported';
                        badge.title = `Last imported as: "${importStatus.name || plan.title}". Selecting will refresh this project.`;
                        badge.style.cssText = `
                            font-size:10px;font-weight:600;padding:2px 7px;border-radius:10px;
                            background:rgba(99,102,241,0.25);color:#a5b4fc;
                            white-space:nowrap;flex-shrink:0;
                        `;
                        row.appendChild(badge);
                    }

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
                    const updateCount = selected.filter(cb => cb.dataset.isUpdate === 'true').length;
                    const newCount = selected.length - updateCount;
                    importBtn.disabled = selected.length === 0;
                    importBtn.style.opacity = selected.length > 0 ? '1' : '0.5';
                    importBtn.style.cursor  = selected.length > 0 ? 'pointer' : 'not-allowed';
                    if (selected.length === 0) {
                        importBtn.textContent = '📥 Import Project';
                    } else if (selected.length === 1 && updateCount === 1) {
                        importBtn.textContent = `🔄 Refresh Project`;
                    } else if (selected.length === 1) {
                        importBtn.textContent = `📥 Import Project`;
                    } else if (newCount === 0) {
                        importBtn.textContent = `🔄 Refresh ${selected.length} Projects`;
                    } else if (updateCount === 0) {
                        importBtn.textContent = `📥 Import ${selected.length} Projects → Portfolio`;
                    } else {
                        importBtn.textContent = `📥 Import ${newCount} + 🔄 Refresh ${updateCount}`;
                    }
                }

                importBtn.addEventListener('click', () => {
                    const selected = checkboxes.filter(cb => cb.checked);
                    if (selected.length === 0) { showStatus('Please select at least one project', 'error'); return; }

                    importBtn.disabled = true;
                    importBtn.style.opacity = '0.7';
                    importBtn.textContent = `⏳ ${selected.length > 1 ? 'Processing' : 'Loading'} ${selected.length} project${selected.length > 1 ? 's' : ''}…`;

                    if (selected.length === 1) {
                        // Single plan: Dataverse import or update
                        const cb = selected[0];
                        onComplete({
                            dataverseUrl,
                            planId: cb.value,
                            planTitle: cb.dataset.title,
                            isUpdate: cb.dataset.isUpdate === 'true',
                            existingStoreId: cb.dataset.existingStoreId || null,
                        });
                    } else {
                        // Multiple plans: portfolio import
                        const planIds         = selected.map(cb => cb.value);
                        const planTitles      = selected.map(cb => cb.dataset.title);
                        const existingStoreIds = selected.map(cb => cb.dataset.existingStoreId || null);
                        onComplete({ dataverseUrl, planIds, planTitles, existingStoreIds, isPortfolioImport: true });
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
                // Notify app to re-render Gantt/table with the merged data
                _emitSyncEvent('success', changedTasks.length);
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

        // Reflect current auto-sync state when the panel opens
        const isAlreadyRunning = _syncIntervals.has(planId);
        autoSyncCheckbox.checked = isAlreadyRunning;

        autoSyncCheckbox.addEventListener('change', (e) => {
            if (e.target.checked) {
                startAutoSync(project, planId, 300000); // 5 minutes
                autoSyncLabel.style.color = '#4ade80';
                autoSyncLabel.querySelector('span').textContent = 'Auto-Pull: ON (every 5min)';
            } else {
                stopAutoSync(planId);
                autoSyncLabel.style.color = '';
                autoSyncLabel.querySelector('span').textContent = 'Auto-Pull: OFF';
            }
        });

        const autoSyncText = document.createElement('span');
        autoSyncText.textContent = isAlreadyRunning ? 'Auto-Pull: ON (every 5min)' : 'Auto-Pull: OFF';
        if (isAlreadyRunning) autoSyncLabel.style.color = '#4ade80';
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
        importFromDataverse,
        importMultipleDataverseProjects,
        pushTaskToPlanner,
        syncProjectToPlanner,
        startAutoSync,
        stopAutoSync,
        getSyncStatus,
        renderSetupWizard,
        renderSyncPanel,
        // Pure Dataverse API
        discoverDataverseUrl,
        listDataverseProjects,
        importFromDataverse,
    };
