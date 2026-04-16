/**
 * ProjectFlow™ © 2026 Ahmed M. Fawzy
 * Microsoft Graph API Client for Live Bi-Directional Planner Sync
 * Uses MSAL.js (PublicClientApplication) with PKCE flow
 */



    // ============================================================================
    // STATE & CONFIG
    // ============================================================================

    let msalApp = null;
    const CONFIG_KEY = 'pf_msgraph_config';
    const SCOPES = ['Tasks.ReadWrite', 'Group.Read.All', 'User.Read', 'offline_access'];
    const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0';
    let autoSyncInterval = null;

    // ============================================================================
    // AUTHENTICATION
    // ============================================================================

    function configure(clientId, tenantId) {
        try {
            if (!window.msal) {
                throw new Error('MSAL.js library not loaded. Ensure msal-browser is included via CDN.');
            }

            const config = {
                auth: {
                    clientId,
                    authority: `https://login.microsoftonline.com/${tenantId}`,
                    redirectUri: window.location.origin,
                },
                cache: {
                    cacheLocation: 'localStorage',
                    storeAuthStateInCookie: false,
                },
                system: {
                    loggerOptions: {
                        loggerCallback: () => {}, // Suppress logs
                        piiLoggingEnabled: false,
                    },
                },
            };

            msalApp = new window.msal.PublicClientApplication(config);
            localStorage.setItem(CONFIG_KEY, JSON.stringify({ clientId, tenantId }));

            return msalApp.initialize();
        } catch (err) {
            throw new Error(`MSGraph configure failed: ${err.message}`);
        }
    }

    function signIn() {
        try {
            if (!msalApp) {
                throw new Error('MSGraphClient not configured. Call configure() first.');
            }
            return msalApp.loginPopup({ scopes: SCOPES });
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

    async function _call(method, path, body = null, retryCount = 0) {
        try {
            const token = await _getAccessToken();
            const url = `${GRAPH_ENDPOINT}${path}`;
            const headers = {
                Authorization: `Bearer ${token}`,
                'Content-Type': 'application/json',
            };

            const options = { method, headers };
            if (body) {
                options.body = JSON.stringify(body);
            }

            const response = await fetch(url, options);

            // Handle 429 Rate Limit
            if (response.status === 429) {
                if (retryCount >= 3) {
                    throw new Error('Rate limit exceeded after 3 retries.');
                }
                const retryAfter = parseInt(response.headers.get('Retry-After') || '2', 10);
                await new Promise(r => setTimeout(r, retryAfter * 1000));
                return _call(method, path, body, retryCount + 1);
            }

            // Handle 409 Conflict (ETag mismatch)
            if (response.status === 409) {
                throw new Error('ETag conflict. Task was modified remotely. Refresh and retry.');
            }

            if (!response.ok) {
                const errData = await response.text();
                throw new Error(`Graph API error ${response.status}: ${errData}`);
            }

            return await response.json();
        } catch (err) {
            throw new Error(`Graph API call failed (${method} ${path}): ${err.message}`);
        }
    }

    // ============================================================================
    // PLANNER READ OPERATIONS
    // ============================================================================

    async function getMyPlans() {
        try {
            const tasksData = await _call('GET', '/me/planner/tasks');
            const tasks = tasksData.value || [];

            // Extract unique planIds
            const planIdSet = new Set();
            tasks.forEach(t => {
                if (t.planId) planIdSet.add(t.planId);
            });

            // Fetch each plan details
            const plans = [];
            for (const planId of planIdSet) {
                try {
                    const plan = await _call('GET', `/planner/plans/${planId}`);
                    plans.push({
                        id: plan.id,
                        title: plan.title,
                        owner: plan.owner,
                        createdBy: plan.createdBy,
                    });
                } catch (err) {
                    // Skip plans we can't access
                }
            }

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
            const [planData, bucketsData, tasksData] = await Promise.all([
                _call('GET', `/planner/plans/${planId}`),
                _call('GET', `/planner/plans/${planId}/buckets`),
                _call('GET', `/planner/plans/${planId}/tasks`),
            ]);

            const buckets = (bucketsData.value || []).map(b => ({
                id: b.id,
                name: b.name,
                orderHint: b.orderHint,
            }));

            const tasks = tasksData.value || [];

            return {
                plan: {
                    id: planData.id,
                    title: planData.title,
                    owner: planData.owner,
                    createdBy: planData.createdBy,
                },
                buckets,
                tasks,
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
    // MAPPING: PLANNER → PROJECTFLOW
    // ============================================================================

    function plannerToProject(plan, buckets, tasks, taskDetailsMap) {
        const project = {
            id: plan.id,
            name: plan.title,
            owner: plan.owner,
            createdBy: plan.createdBy,
            tasks: [],
            _plannerId: plan.id,
        };

        // Sort buckets by orderHint
        const sortedBuckets = [...buckets].sort((a, b) => {
            if (!a.orderHint && !b.orderHint) return 0;
            if (!a.orderHint) return 1;
            if (!b.orderHint) return -1;
            return a.orderHint.localeCompare(b.orderHint);
        });

        // Create summary task for each bucket
        sortedBuckets.forEach(bucket => {
            const bucketTask = {
                id: bucket.id,
                name: bucket.name,
                outlineLevel: 1,
                resourceNames: [],
                percentComplete: 0,
                tags: [],
                notes: '',
                children: [],
                _plannerId: bucket.id,
                _plannerBucketId: bucket.id,
            };

            // Find all tasks in this bucket
            const bucketTasks = tasks.filter(t => t.bucketId === bucket.id);
            bucketTasks.forEach(task => {
                const details = taskDetailsMap[task.id] || {};
                const today = new Date().toISOString().split('T')[0];
                const startDate = task.startDateTime ? task.startDateTime.split('T')[0] : today;
                const finishDate = task.dueDateTime
                    ? task.dueDateTime.split('T')[0]
                    : _addDays(startDate, 1);

                // Parse assignments (object keyed by userId)
                const resourceNames = [];
                if (task.assignments && typeof task.assignments === 'object') {
                    Object.values(task.assignments).forEach(assignment => {
                        if (assignment.displayName) {
                            resourceNames.push(assignment.displayName);
                        } else {
                            resourceNames.push(assignment.assignedBy?.user?.displayName || 'Unknown');
                        }
                    });
                }

                // Map appliedCategories to tags
                const tags = [];
                if (task.appliedCategories && typeof task.appliedCategories === 'object') {
                    Object.keys(task.appliedCategories).forEach(key => {
                        if (task.appliedCategories[key] === true) {
                            tags.push(key);
                        }
                    });
                }

                // Build notes with description + checklist
                let notes = details.description || '';
                if (details.checklist && Object.keys(details.checklist).length > 0) {
                    const checklistItems = Object.values(details.checklist)
                        .map(item => `- ${item.title}${item.isChecked ? ' ✓' : ''}`)
                        .join('\n');
                    notes = notes ? `${notes}\n\nChecklist:\n${checklistItems}` : `Checklist:\n${checklistItems}`;
                }

                const leafTask = {
                    id: task.id,
                    name: task.title,
                    outlineLevel: 2,
                    start: startDate,
                    finish: finishDate,
                    percentComplete: task.percentComplete || 0,
                    resourceNames,
                    tags,
                    notes,
                    _plannerId: task.id,
                    _plannerEtag: task['@odata.etag'],
                };

                bucketTask.children.push(leafTask);
            });

            project.tasks.push(bucketTask);
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

    function projectTaskToPlanner(task, bucketId) {
        const body = {
            title: task.name,
        };

        if (task.percentComplete !== undefined) {
            // Round to nearest 0, 50, 100
            const rounded = Math.round(task.percentComplete / 50) * 50;
            body.percentComplete = Math.min(100, Math.max(0, rounded));
        }

        if (task.start) {
            body.startDateTime = `${task.start}T00:00:00Z`;
        }

        if (task.finish) {
            body.dueDateTime = `${task.finish}T23:59:59Z`;
        }

        // Assignments: convert resourceNames to format {userId: {displayName}}
        if (task.resourceNames && task.resourceNames.length > 0) {
            body.assignments = {};
            task.resourceNames.forEach(name => {
                const userId = _sanitizeUserId(name);
                body.assignments[userId] = { '@odata.type': '#microsoft.graph.plannerAssignment' };
            });
        }

        // Categories: convert tags to appliedCategories {category1: true, ...}
        if (task.tags && task.tags.length > 0) {
            body.appliedCategories = {};
            task.tags.forEach(tag => {
                body.appliedCategories[tag] = true;
            });
        }

        return body;
    }

    function _sanitizeUserId(name) {
        // Simple sanitization: replace spaces with dashes, lowercase
        return name.toLowerCase().replace(/\s+/g, '-');
    }

    // ============================================================================
    // IMPORT & PUSH OPERATIONS
    // ============================================================================

    async function importPlan(planId) {
        try {
            const { plan, buckets, tasks } = await getPlanDetails(planId);

            // Fetch task details for all tasks
            const taskDetailsMap = {};
            for (const task of tasks) {
                try {
                    const details = await getPlanTaskDetails(task.id);
                    taskDetailsMap[task.id] = details;
                } catch (err) {
                    // Continue even if details fail
                    taskDetailsMap[task.id] = {};
                }
            }

            const project = plannerToProject(plan, buckets, tasks, taskDetailsMap);
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
            if (task._plannerId && task._plannerId !== planId) {
                // Existing Planner task: PATCH
                const headers = {
                    'If-Match': task._plannerEtag || '*',
                };
                result = await _call('PATCH', `/planner/tasks/${task._plannerId}`, body);
            } else {
                // New task: POST
                body.planId = planId;
                body.bucketId = task._plannerBucketId;
                result = await _call('POST', '/planner/tasks', body);
            }

            return result;
        } catch (err) {
            throw new Error(`pushTaskToPlanner failed: ${err.message}`);
        }
    }

    async function syncProjectToPlanner(project, planId) {
        try {
            const summary = { updated: 0, created: 0, failed: 0 };

            const collectLeafTasks = (tasks, leaves = []) => {
                tasks.forEach(t => {
                    if (t.children && t.children.length > 0) {
                        collectLeafTasks(t.children, leaves);
                    } else {
                        leaves.push(t);
                    }
                });
                return leaves;
            };

            const leafTasks = collectLeafTasks(project.tasks);

            for (const task of leafTasks) {
                try {
                    await pushTaskToPlanner(task, planId);
                    if (task._plannerId) {
                        summary.updated++;
                    } else {
                        summary.created++;
                    }
                } catch (err) {
                    summary.failed++;
                }
            }

            return summary;
        } catch (err) {
            throw new Error(`syncProjectToPlanner failed: ${err.message}`);
        }
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
                    // Pull remote changes
                    const { tasks: remoteTasks } = await getPlanDetails(planId);
                    _mergeRemoteChanges(project, remoteTasks);

                    // Push local changes
                    await syncProjectToPlanner(project, planId);
                } catch (err) {
                    // Log but don't crash
                    console.warn('Auto-sync error:', err.message);
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
        const collectAllTasks = (tasks, all = []) => {
            tasks.forEach(t => {
                all.push(t);
                if (t.children) {
                    collectAllTasks(t.children, all);
                }
            });
            return all;
        };

        const allLocalTasks = collectAllTasks(project.tasks);
        const remoteMap = new Map(remoteTasks.map(t => [t.id, t]));

        allLocalTasks.forEach(localTask => {
            const remote = remoteMap.get(localTask._plannerId);
            if (!remote) return;

            // Update if remote is newer
            const localMod = new Date(localTask._lastModified || 0);
            const remoteMod = new Date(remote.lastModifiedDateTime || 0);

            if (remoteMod > localMod) {
                localTask.percentComplete = remote.percentComplete || 0;
                localTask._plannerEtag = remote['@odata.etag'];
                localTask._lastModified = remote.lastModifiedDateTime;

                // Update assignments if present
                if (remote.assignments && typeof remote.assignments === 'object') {
                    localTask.resourceNames = [];
                    Object.values(remote.assignments).forEach(assignment => {
                        if (assignment.displayName) {
                            localTask.resourceNames.push(assignment.displayName);
                        }
                    });
                }
            }
        });
    }

    // ============================================================================
    // UI: SETUP WIZARD
    // ============================================================================

    function renderSetupWizard(container, onComplete) {
        if (!container) {
            throw new Error('renderSetupWizard: container not found.');
        }

        let currentStep = 1;
        const steps = [
            { title: 'Azure AD App Details', id: 'step-credentials' },
            { title: 'Sign In with Microsoft', id: 'step-signin' },
            { title: 'Select Plan', id: 'step-selectplan' },
        ];

        const wizard = document.createElement('div');
        wizard.className = 'ms-graph-wizard';
        wizard.style.cssText = `
            max-width: 500px;
            margin: 20px auto;
            padding: 20px;
            border: 1px solid #ddd;
            border-radius: 8px;
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        `;

        function renderStep(step) {
            wizard.innerHTML = '';

            const header = document.createElement('div');
            header.style.cssText = 'margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 10px;';
            header.textContent = `Step ${step} of ${steps.length}: ${steps[step - 1].title}`;
            wizard.appendChild(header);

            if (step === 1) {
                renderCredentialsStep();
            } else if (step === 2) {
                renderSignInStep();
            } else if (step === 3) {
                renderSelectPlanStep();
            }
        }

        function renderCredentialsStep() {
            const form = document.createElement('form');
            form.style.cssText = 'display: flex; flex-direction: column; gap: 12px;';

            const clientIdLabel = document.createElement('label');
            clientIdLabel.textContent = 'Client ID:';
            clientIdLabel.style.cssText = 'font-weight: 500; margin-top: 10px;';

            const clientIdInput = document.createElement('input');
            clientIdInput.type = 'text';
            clientIdInput.placeholder = 'Paste your Azure AD Client ID';
            clientIdInput.style.cssText = 'padding: 8px; border: 1px solid #ccc; border-radius: 4px;';

            const tenantIdLabel = document.createElement('label');
            tenantIdLabel.textContent = 'Tenant ID:';
            tenantIdLabel.style.cssText = 'font-weight: 500; margin-top: 10px;';

            const tenantIdInput = document.createElement('input');
            tenantIdInput.type = 'text';
            tenantIdInput.placeholder = 'Paste your Azure AD Tenant ID';
            tenantIdInput.style.cssText = 'padding: 8px; border: 1px solid #ccc; border-radius: 4px;';

            const infoLink = document.createElement('a');
            infoLink.href = 'https://learn.microsoft.com/azure/active-directory/develop/quickstart-register-app';
            infoLink.target = '_blank';
            infoLink.textContent = 'How to get your app details';
            infoLink.style.cssText = 'color: #0078d4; text-decoration: none; font-size: 12px;';

            const nextBtn = document.createElement('button');
            nextBtn.type = 'button';
            nextBtn.textContent = 'Next';
            nextBtn.style.cssText = `
                padding: 10px;
                background: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                margin-top: 15px;
            `;

            nextBtn.addEventListener('click', async (e) => {
                e.preventDefault();
                const cId = clientIdInput.value.trim();
                const tId = tenantIdInput.value.trim();

                if (!cId || !tId) {
                    alert('Please enter both Client ID and Tenant ID.');
                    return;
                }

                try {
                    await configure(cId, tId);
                    currentStep = 2;
                    renderStep(currentStep);
                } catch (err) {
                    alert(`Configuration failed: ${err.message}`);
                }
            });

            form.appendChild(clientIdLabel);
            form.appendChild(clientIdInput);
            form.appendChild(tenantIdLabel);
            form.appendChild(tenantIdInput);
            form.appendChild(infoLink);
            form.appendChild(nextBtn);

            wizard.appendChild(form);
        }

        function renderSignInStep() {
            const signInBtn = document.createElement('button');
            signInBtn.type = 'button';
            signInBtn.textContent = 'Sign In with Microsoft';
            signInBtn.style.cssText = `
                padding: 12px;
                background: #0078d4;
                color: white;
                border: none;
                border-radius: 4px;
                font-size: 14px;
                cursor: pointer;
                margin: 20px 0;
            `;

            signInBtn.addEventListener('click', async () => {
                try {
                    await signIn();
                    const account = getAccount();
                    if (account) {
                        const msg = document.createElement('p');
                        msg.textContent = `Signed in as ${account.email}`;
                        msg.style.cssText = 'color: #107c10; margin: 10px 0;';
                        wizard.appendChild(msg);

                        const nextBtn = document.createElement('button');
                        nextBtn.textContent = 'Next';
                        nextBtn.style.cssText = `
                            padding: 10px;
                            background: #0078d4;
                            color: white;
                            border: none;
                            border-radius: 4px;
                            cursor: pointer;
                            margin-top: 15px;
                        `;
                        nextBtn.addEventListener('click', () => {
                            currentStep = 3;
                            renderStep(currentStep);
                        });
                        wizard.appendChild(nextBtn);
                    }
                } catch (err) {
                    alert(`Sign-in failed: ${err.message}`);
                }
            });

            wizard.appendChild(signInBtn);
        }

        function renderSelectPlanStep() {
            const loadingMsg = document.createElement('p');
            loadingMsg.textContent = 'Loading your plans...';
            loadingMsg.style.cssText = 'color: #666;';
            wizard.appendChild(loadingMsg);

            getMyPlans()
                .then(plans => {
                    wizard.innerHTML = '';
                    const header = document.createElement('div');
                    header.style.cssText = 'margin-bottom: 20px; border-bottom: 1px solid #eee; padding-bottom: 10px;';
                    header.textContent = `Step ${currentStep} of ${steps.length}: ${steps[currentStep - 1].title}`;
                    wizard.appendChild(header);

                    const label = document.createElement('label');
                    label.textContent = 'Select a Plan:';
                    label.style.cssText = 'font-weight: 500; display: block; margin-bottom: 10px;';
                    wizard.appendChild(label);

                    const select = document.createElement('select');
                    select.style.cssText = 'padding: 8px; border: 1px solid #ccc; border-radius: 4px; width: 100%;';

                    const placeholder = document.createElement('option');
                    placeholder.value = '';
                    placeholder.textContent = '-- Choose a plan --';
                    select.appendChild(placeholder);

                    plans.forEach(plan => {
                        const option = document.createElement('option');
                        option.value = plan.id;
                        option.textContent = plan.title;
                        select.appendChild(option);
                    });

                    const completeBtn = document.createElement('button');
                    completeBtn.type = 'button';
                    completeBtn.textContent = 'Connect Plan';
                    completeBtn.style.cssText = `
                        padding: 10px;
                        background: #107c10;
                        color: white;
                        border: none;
                        border-radius: 4px;
                        cursor: pointer;
                        margin-top: 15px;
                        width: 100%;
                    `;

                    completeBtn.addEventListener('click', () => {
                        const planId = select.value;
                        if (!planId) {
                            alert('Please select a plan.');
                            return;
                        }
                        const planTitle = select.options[select.selectedIndex].textContent;
                        onComplete({ planId, planTitle });
                    });

                    wizard.appendChild(select);
                    wizard.appendChild(completeBtn);
                })
                .catch(err => {
                    loadingMsg.textContent = `Error loading plans: ${err.message}`;
                    loadingMsg.style.cssText = 'color: #c50f1f;';
                });
        }

        renderStep(currentStep);
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

        const pullBtn = document.createElement('button');
        pullBtn.textContent = 'Pull from Planner';
        pullBtn.style.cssText = `
            padding: 8px 12px;
            background: #0078d4;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
        `;
        pullBtn.addEventListener('click', async () => {
            try {
                pullBtn.disabled = true;
                pullBtn.textContent = 'Pulling...';
                const { tasks: remoteTasks } = await getPlanDetails(planId);
                _mergeRemoteChanges(project, remoteTasks);
                syncTime.textContent = `Last sync: ${new Date().toLocaleTimeString()}`;
                pullBtn.textContent = 'Pull from Planner';
                pullBtn.disabled = false;
            } catch (err) {
                alert(`Pull failed: ${err.message}`);
                pullBtn.textContent = 'Pull from Planner';
                pullBtn.disabled = false;
            }
        });

        const pushBtn = document.createElement('button');
        pushBtn.textContent = 'Push to Planner';
        pushBtn.style.cssText = `
            padding: 8px 12px;
            background: #107c10;
            color: white;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 12px;
        `;
        pushBtn.addEventListener('click', async () => {
            try {
                pushBtn.disabled = true;
                pushBtn.textContent = 'Pushing...';
                const summary = await syncProjectToPlanner(project, planId);
                syncTime.textContent = `Last sync: ${new Date().toLocaleTimeString()} (${summary.updated} updated, ${summary.created} created)`;
                pushBtn.textContent = 'Push to Planner';
                pushBtn.disabled = false;
            } catch (err) {
                alert(`Push failed: ${err.message}`);
                pushBtn.textContent = 'Push to Planner';
                pushBtn.disabled = false;
            }
        });

        const autoSyncLabel = document.createElement('label');
        autoSyncLabel.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer;';

        const autoSyncCheckbox = document.createElement('input');
        autoSyncCheckbox.type = 'checkbox';
        autoSyncCheckbox.style.cssText = 'cursor: pointer;';
        autoSyncCheckbox.addEventListener('change', (e) => {
            if (e.target.checked) {
                startAutoSync(project, planId, 60000);
                autoSyncLabel.textContent = 'Auto-Sync: ON';
                autoSyncLabel.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer; color: #107c10; font-weight: 500;';
            } else {
                stopAutoSync();
                autoSyncLabel.textContent = 'Auto-Sync: OFF';
                autoSyncLabel.style.cssText = 'display: flex; align-items: center; gap: 8px; cursor: pointer;';
            }
        });

        autoSyncLabel.appendChild(autoSyncCheckbox);
        autoSyncLabel.appendChild(document.createTextNode('Auto-Sync: OFF'));

        buttonGroup.appendChild(pullBtn);
        buttonGroup.appendChild(pushBtn);
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
