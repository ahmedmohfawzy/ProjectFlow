/**
 * ProjectFlow™ © 2026 Ahmed M. Fawzy
 * Microsoft Dynamics 365 Project Accounting & Project Operations Integration
 *
 * Supports two D365 modes:
 *   1. Project Operations (Dataverse)
 *   2. Finance & Operations (OData)
 */



    // Configuration constants
    const CONFIG_KEY = 'pf_d365_config';
    const MODES = {
        OPERATIONS: 'operations',
        FINANCE: 'finance'
    };

    const DATAVERSE_VERSION = 'v9.2';
    const RATE_LIMIT_DELAY = 1000;

    // D365 Entity field mappings
    const ENTITIES = {
        PROJECTS_OPS: 'msdyn_projects',
        TASKS_OPS: 'msdyn_projecttasks',
        RESOURCES_OPS: 'msdyn_bookableresources',
        ASSIGNMENTS_OPS: 'msdyn_resourceassignments',
        BUDGET_LINES_OPS: 'msdyn_projectbudgetlines',
        JOURNALS_OPS: 'msdyn_actuals',

        PROJECTS_FO: 'ProjTable',
        ACTIVITIES_FO: 'ProjActivity',
        RESOURCES_FO: 'ResTable',
        BUDGET_FO: 'ProjBudgetLineCost',
        TRANSACTIONS_FO: 'ProjTransPosting'
    };

    const PROJECT_STATUS = {
        ON_TRACK: 192350000,
        AT_RISK: 192350001,
        OFF_TRACK: 192350002
    };

    const TRANSACTION_CLASS = {
        TIME: 192350000,
        EXPENSE: 192350001,
        MATERIAL: 192350002
    };

    // Internal state
    let config = null;
    let msalInstance = null;

    /**
     * Initialize MSAL instance reference (shared with MSGraphClient)
     */
    function _getMsalInstance() {
        if (msalInstance) return msalInstance;
        msalInstance = window.msal;
        if (!msalInstance) {
            throw new Error('MSAL instance not available. Ensure MSGraphClient is initialized first.');
        }
        return msalInstance;
    }

    /**
     * Load configuration from localStorage
     */
    function _loadConfig() {
        const stored = localStorage.getItem(CONFIG_KEY);
        if (stored) {
            try {
                config = JSON.parse(stored);
            } catch (e) {
                console.error('Failed to parse D365 config:', e);
                config = null;
            }
        }
        return config;
    }

    /**
     * Save configuration to localStorage
     */
    function _saveConfig(cfg) {
        config = cfg;
        localStorage.setItem(CONFIG_KEY, JSON.stringify(cfg));
    }

    /**
     * Get access token via MSAL (silent first, popup fallback)
     */
    async function _getAccessToken() {
        if (!config) {
            throw new Error('D365 not configured. Call configure() first.');
        }

        const msal = _getMsalInstance();
        const scopes = [`${config.environmentUrl}/.default`];
        const account = msal.getAllAccounts()[0];

        try {
            // Try silent token acquisition
            const response = await msal.acquireTokenSilent({
                scopes,
                account
            });
            return response.accessToken;
        } catch (error) {
            // Fall back to popup
            try {
                const response = await msal.acquireTokenPopup({
                    scopes,
                    account
                });
                return response.accessToken;
            } catch (popupError) {
                console.error('Token acquisition failed:', popupError);
                throw new Error('Failed to acquire access token for D365');
            }
        }
    }

    /**
     * Make HTTP request to Dataverse API
     */
    async function _callDataverse(method, entity, query = '', body = null) {
        if (!config || config.mode !== MODES.OPERATIONS) {
            throw new Error('Dataverse API called in non-Operations mode');
        }

        const token = await _getAccessToken();
        let url = `${config.environmentUrl}/api/data/${DATAVERSE_VERSION}/${entity}`;
        if (query) url += `?${query}`;

        const options = {
            method,
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json'
            }
        };

        if (body) {
            options.body = JSON.stringify(body);
        }

        try {
            const response = await fetch(url, options);

            // Handle rate limiting
            if (response.status === 429) {
                await new Promise(resolve => setTimeout(resolve, RATE_LIMIT_DELAY));
                return _callDataverse(method, entity, query, body);
            }

            if (!response.ok) {
                const error = await response.text();
                console.error(`Dataverse API error (${response.status}):`, error);
                throw new Error(`Dataverse API error: ${response.status}`);
            }

            return await response.json();
        } catch (error) {
            console.error('Dataverse API call failed:', error);
            throw error;
        }
    }

    /**
     * Make HTTP request to Finance & Operations OData API
     */
    async function _callFO(method, entity, query = '', body = null) {
        if (!config || config.mode !== MODES.FINANCE) {
            throw new Error('F&O API called in non-Finance mode');
        }

        const token = await _getAccessToken();
        let url = `${config.environmentUrl}/data/${entity}`;
        if (query) url += `?${query}`;

        const options = {
            method,
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'Accept': 'application/json'
            }
        };

        if (body) {
            options.body = JSON.stringify(body);
        }

        try {
            const response = await fetch(url, options);

            // Handle rate limiting
            if (response.status === 429) {
                await new Promise(resolve => setTimeout(resolve, RATE_LIMIT_DELAY));
                return _callFO(method, entity, query, body);
            }

            if (!response.ok) {
                const error = await response.text();
                console.error(`F&O API error (${response.status}):`, error);
                throw new Error(`F&O API error: ${response.status}`);
            }

            return await response.json();
        } catch (error) {
            console.error('F&O API call failed:', error);
            throw error;
        }
    }

    /**
     * Map Project Operations project to ProjectFlow format
     */
    function _mapProjectOps(proj) {
        let status = 'On Track';
        if (proj.msdyn_overallprojectstatus === PROJECT_STATUS.AT_RISK) {
            status = 'At Risk';
        } else if (proj.msdyn_overallprojectstatus === PROJECT_STATUS.OFF_TRACK) {
            status = 'Off Track';
        }

        return {
            id: proj.msdyn_projectid,
            name: proj.msdyn_subject || 'Unnamed Project',
            start: proj.msdyn_scheduledstart ? new Date(proj.msdyn_scheduledstart) : null,
            finish: proj.msdyn_scheduledend ? new Date(proj.msdyn_scheduledend) : null,
            status,
            manager: proj.msdyn_projectmanager_displayname || 'Unassigned',
            budget: proj.msdyn_totalplannedcost || 0,
            actualRevenue: proj.msdyn_totalactualsales || 0,
            description: proj.msdyn_description || '',
            _d365Mode: MODES.OPERATIONS
        };
    }

    /**
     * Map Finance & Operations project to ProjectFlow format
     */
    function _mapProjectFO(proj) {
        return {
            id: proj.ProjId,
            name: proj.Name || 'Unnamed Project',
            start: proj.StartDate ? new Date(proj.StartDate) : null,
            finish: proj.EndDate ? new Date(proj.EndDate) : null,
            status: proj.Status || 'InProcess',
            manager: 'Unassigned',
            budget: proj.SalesPrice || 0,
            customer: proj.CustAccount || '',
            _d365Mode: MODES.FINANCE
        };
    }

    /**
     * Map Project Task (Operations) to ProjectFlow task
     */
    function _mapTaskOps(task) {
        return {
            _d365Id: task.msdyn_projecttaskid,
            name: task.msdyn_subject || 'Unnamed Task',
            start: task.msdyn_scheduledstart ? new Date(task.msdyn_scheduledstart) : null,
            finish: task.msdyn_scheduledend ? new Date(task.msdyn_scheduledend) : null,
            durationDays: task.msdyn_scheduleddurationminutes ? (task.msdyn_scheduleddurationminutes / 480) : 0,
            percentComplete: task.msdyn_progress || 0,
            critical: task.msdyn_iscritical || false,
            outlineLevel: task.msdyn_outlinelevel || 0,
            milestone: task.msdyn_ismilestone || false,
            cost: task.msdyn_effort ? (task.msdyn_effort * 8) : 0, // Convert hours to cost-hours
            effortCompleted: task.msdyn_effortcompleted || 0,
            description: task.msdyn_description || '',
            wbsId: task.msdyn_wbsid || ''
        };
    }

    /**
     * Map Activity (Finance) to ProjectFlow task
     */
    function _mapActivityFO(activity) {
        return {
            _d365Id: activity.ActivityNumber,
            name: activity.Description || 'Unnamed Activity',
            start: activity.FromDate ? new Date(activity.FromDate) : null,
            finish: activity.ToDate ? new Date(activity.ToDate) : null,
            status: activity.Status || 'Open',
            cost: activity.CostPrice || 0,
            revenue: activity.SalesPrice || 0
        };
    }

    /**
     * Public: Configure D365 connection
     */
    function configure(cfg) {
        if (!cfg.environmentUrl || !cfg.mode || !cfg.clientId || !cfg.tenantId) {
            throw new Error('Invalid D365 configuration. Required: environmentUrl, mode, clientId, tenantId');
        }
        _saveConfig(cfg);
    }

    /**
     * Public: Sign in to D365
     */
    async function signIn() {
        if (!config) {
            throw new Error('D365 not configured');
        }

        const msal = _getMsalInstance();
        const scopes = [`${config.environmentUrl}/.default`];

        try {
            const response = await msal.loginPopup({
                scopes,
                prompt: 'select_account'
            });
            console.log('D365 sign-in successful');
            return response;
        } catch (error) {
            console.error('D365 sign-in failed:', error);
            throw error;
        }
    }

    /**
     * Public: Sign out from D365
     */
    async function signOut() {
        const msal = _getMsalInstance();
        const account = msal.getAllAccounts()[0];

        try {
            await msal.logoutPopup({
                account,
                mainWindowRedirectUri: '/'
            });
            config = null;
            localStorage.removeItem(CONFIG_KEY);
            console.log('D365 sign-out successful');
        } catch (error) {
            console.error('D365 sign-out failed:', error);
            throw error;
        }
    }

    /**
     * Public: Check if authenticated
     */
    function isAuthenticated() {
        const msal = _getMsalInstance();
        return msal && msal.getAllAccounts().length > 0 && config !== null;
    }

    /**
     * Public: Get all projects
     */
    async function getProjects() {
        if (!config) throw new Error('D365 not configured');

        try {
            if (config.mode === MODES.OPERATIONS) {
                const query = '$select=msdyn_projectid,msdyn_subject,msdyn_scheduledstart,msdyn_scheduledend,msdyn_description,msdyn_projectmanager_displayname,msdyn_overallprojectstatus,msdyn_totalplannedcost,msdyn_totalactualsales&$orderby=msdyn_subject';
                const result = await _callDataverse('GET', ENTITIES.PROJECTS_OPS, query);
                return result.value.map(_mapProjectOps);
            } else {
                const query = '$select=ProjId,Name,ProjType,Status,StartDate,EndDate,CustAccount,SalesPrice&$orderby=Name';
                const result = await _callFO('GET', ENTITIES.PROJECTS_FO, query);
                return result.value.map(_mapProjectFO);
            }
        } catch (error) {
            console.error('Failed to fetch projects:', error);
            throw error;
        }
    }

    /**
     * Public: Get tasks for a project
     */
    async function getProjectTasks(projectId) {
        if (!config) throw new Error('D365 not configured');

        try {
            if (config.mode === MODES.OPERATIONS) {
                const query = `$filter=_msdyn_project_value eq (${projectId})&$select=msdyn_projecttaskid,msdyn_subject,msdyn_scheduledstart,msdyn_scheduledend,msdyn_scheduleddurationminutes,msdyn_progress,msdyn_effort,msdyn_effortcompleted,msdyn_iscritical,msdyn_outlinelevel,msdyn_ismilestone,msdyn_wbsid,msdyn_description&$orderby=msdyn_outlinelevel,msdyn_wbsid`;
                const result = await _callDataverse('GET', ENTITIES.TASKS_OPS, query);
                return result.value.map(_mapTaskOps).sort((a, b) => {
                    if (a.outlineLevel !== b.outlineLevel) return a.outlineLevel - b.outlineLevel;
                    return (a.wbsId || '').localeCompare(b.wbsId || '');
                });
            } else {
                const query = `$filter=ProjId eq '${projectId}'&$select=ActivityNumber,Description,FromDate,ToDate,Status,CostPrice,SalesPrice`;
                const result = await _callFO('GET', ENTITIES.ACTIVITIES_FO, query);
                return result.value.map(_mapActivityFO);
            }
        } catch (error) {
            console.error('Failed to fetch project tasks:', error);
            throw error;
        }
    }

    /**
     * Public: Get available resources
     */
    async function getResources() {
        if (!config) throw new Error('D365 not configured');

        try {
            if (config.mode === MODES.OPERATIONS) {
                const query = '$select=bookableresourceid,name,resourcetype';
                const result = await _callDataverse('GET', ENTITIES.RESOURCES_OPS, query);
                return result.value.map(r => ({
                    id: r.bookableresourceid,
                    name: r.name,
                    type: r.resourcetype || 'User'
                }));
            } else {
                const query = '$select=ResourceId,Name,ResourceType';
                const result = await _callFO('GET', ENTITIES.RESOURCES_FO, query);
                return result.value.map(r => ({
                    id: r.ResourceId,
                    name: r.Name,
                    type: r.ResourceType || 'User'
                }));
            }
        } catch (error) {
            console.error('Failed to fetch resources:', error);
            throw error;
        }
    }

    /**
     * Public: Get resource assignments for a project
     */
    async function getAssignments(projectId) {
        if (!config) throw new Error('D365 not configured');

        try {
            if (config.mode === MODES.OPERATIONS) {
                const query = `$filter=_msdyn_taskid_value eq (${projectId})&$select=msdyn_resourceassignmentid,_msdyn_taskid_value,_msdyn_bookableresourceid_value,msdyn_plannedwork,msdyn_actualwork&$expand=msdyn_BookableResource_ResourceAssignment($select=name)`;
                const result = await _callDataverse('GET', ENTITIES.ASSIGNMENTS_OPS, query);
                return result.value.map(a => ({
                    id: a.msdyn_resourceassignmentid,
                    taskId: a._msdyn_taskid_value,
                    resourceId: a._msdyn_bookableresourceid_value,
                    resourceName: a.msdyn_BookableResource_ResourceAssignment ? a.msdyn_BookableResource_ResourceAssignment.name : 'Unknown',
                    plannedWork: a.msdyn_plannedwork || 0,
                    actualWork: a.msdyn_actualwork || 0
                }));
            } else {
                // F&O doesn't have direct assignment entity; return empty
                return [];
            }
        } catch (error) {
            console.error('Failed to fetch assignments:', error);
            throw error;
        }
    }

    /**
     * Public: Get budget lines for a project
     */
    async function getBudgetLines(projectId) {
        if (!config) throw new Error('D365 not configured');

        try {
            if (config.mode === MODES.OPERATIONS) {
                const query = `$filter=_msdyn_project_value eq (${projectId})&$select=msdyn_projectbudgetlineid,msdyn_description,msdyn_amount,msdyn_transactionclass`;
                const result = await _callDataverse('GET', ENTITIES.BUDGET_LINES_OPS, query);

                let totalBudget = 0, timeBudget = 0, expenseBudget = 0, materialBudget = 0;
                result.value.forEach(line => {
                    const amount = line.msdyn_amount || 0;
                    totalBudget += amount;
                    switch (line.msdyn_transactionclass) {
                        case TRANSACTION_CLASS.TIME:
                            timeBudget += amount;
                            break;
                        case TRANSACTION_CLASS.EXPENSE:
                            expenseBudget += amount;
                            break;
                        case TRANSACTION_CLASS.MATERIAL:
                            materialBudget += amount;
                            break;
                    }
                });

                return { totalBudget, timeBudget, expenseBudget, materialBudget };
            } else {
                const query = `$filter=ProjId eq '${projectId}'&$select=Amount,TransType`;
                const result = await _callFO('GET', ENTITIES.BUDGET_FO, query);
                const totalBudget = result.value.reduce((sum, line) => sum + (line.Amount || 0), 0);
                return { totalBudget, timeBudget: 0, expenseBudget: 0, materialBudget: 0 };
            }
        } catch (error) {
            console.error('Failed to fetch budget lines:', error);
            throw error;
        }
    }

    /**
     * Public: Get project transactions (for EVM)
     */
    async function getTransactions(projectId, fromDate, toDate) {
        if (!config) throw new Error('D365 not configured');

        try {
            if (config.mode === MODES.OPERATIONS) {
                const from = fromDate.toISOString().split('T')[0];
                const to = toDate.toISOString().split('T')[0];
                const query = `$filter=_msdyn_project_value eq (${projectId}) and msdyn_transactiondate ge ${from} and msdyn_transactiondate le ${to}&$select=msdyn_actualid,msdyn_description,msdyn_transactiondate,msdyn_quantity,msdyn_amount`;
                const result = await _callDataverse('GET', ENTITIES.JOURNALS_OPS, query);
                return result.value.map(t => ({
                    id: t.msdyn_actualid,
                    description: t.msdyn_description || '',
                    date: new Date(t.msdyn_transactiondate),
                    quantity: t.msdyn_quantity || 0,
                    amount: t.msdyn_amount || 0
                }));
            } else {
                const from = fromDate.toISOString().split('T')[0];
                const to = toDate.toISOString().split('T')[0];
                const query = `$filter=ProjId eq '${projectId}' and TransDate ge ${from} and TransDate le ${to}&$select=TransDate,Qty,CostAmount,SalesAmount,Description`;
                const result = await _callFO('GET', ENTITIES.TRANSACTIONS_FO, query);
                return result.value.map(t => ({
                    date: new Date(t.TransDate),
                    quantity: t.Qty || 0,
                    costAmount: t.CostAmount || 0,
                    salesAmount: t.SalesAmount || 0,
                    description: t.Description || ''
                }));
            }
        } catch (error) {
            console.error('Failed to fetch transactions:', error);
            throw error;
        }
    }

    /**
     * Public: Get accounting KPIs
     */
    async function getAccountingKPIs(projectId) {
        if (!config) throw new Error('D365 not configured');

        try {
            const budget = await getBudgetLines(projectId);
            const fromDate = new Date();
            fromDate.setMonth(fromDate.getMonth() - 3);
            const toDate = new Date();
            const transactions = await getTransactions(projectId, fromDate, toDate);

            const budgetCost = budget.totalBudget;
            const actualCost = transactions.reduce((sum, t) => {
                if (config.mode === MODES.OPERATIONS) {
                    return sum + (t.amount || 0);
                } else {
                    return sum + (t.costAmount || 0);
                }
            }, 0);
            const budgetRevenue = budget.totalBudget;
            const actualRevenue = transactions.reduce((sum, t) => {
                if (config.mode === MODES.OPERATIONS) {
                    return sum + (t.amount || 0);
                } else {
                    return sum + (t.salesAmount || 0);
                }
            }, 0);

            const forecastCost = actualCost * 1.15; // Simple projection
            const cpi = actualCost > 0 ? actualRevenue / actualCost : 0;
            const spi = budgetRevenue > 0 ? actualRevenue / budgetRevenue : 0;

            return {
                budgetCost,
                actualCost,
                forecastCost,
                budgetRevenue,
                actualRevenue,
                costVariance: budgetCost - actualCost,
                revenueVariance: budgetRevenue - actualRevenue,
                cpi,
                spi
            };
        } catch (error) {
            console.error('Failed to fetch accounting KPIs:', error);
            throw error;
        }
    }

    /**
     * Public: Import full project from D365
     */
    async function importProject(projectId) {
        if (!config) throw new Error('D365 not configured');

        try {
            const projects = await getProjects();
            const project = projects.find(p => p.id === projectId);
            if (!project) throw new Error(`Project ${projectId} not found`);

            const tasks = await getProjectTasks(projectId);
            const resources = await getResources();
            const assignments = await getAssignments(projectId);
            const budget = await getBudgetLines(projectId);

            return {
                ...project,
                tasks,
                resources,
                assignments,
                budget,
                importedAt: new Date()
            };
        } catch (error) {
            console.error('Failed to import project:', error);
            throw error;
        }
    }

    /**
     * Public: Push task progress back to D365
     */
    async function pushTaskToD365(task, projectId) {
        if (!config || !task._d365Id) {
            throw new Error('Cannot push task: D365 not configured or task missing _d365Id');
        }

        try {
            if (config.mode === MODES.OPERATIONS) {
                const update = {
                    msdyn_progress: task.percentComplete || 0,
                    msdyn_effort: task.cost ? (task.cost / 8) : 0
                };
                if (task.start) update.msdyn_scheduledstart = task.start.toISOString();
                if (task.finish) update.msdyn_scheduledend = task.finish.toISOString();

                await _callDataverse('PATCH', `${ENTITIES.TASKS_OPS}(${task._d365Id})`, '', update);
                console.log(`Task ${task._d365Id} pushed successfully`);
                return { success: true, taskId: task._d365Id };
            } else {
                // F&O OData PATCH
                const update = {
                    Status: task.status || 'Open'
                };
                await _callFO('PATCH', `${ENTITIES.ACTIVITIES_FO}('${task._d365Id}')`, '', update);
                console.log(`Activity ${task._d365Id} pushed successfully`);
                return { success: true, taskId: task._d365Id };
            }
        } catch (error) {
            console.error(`Failed to push task ${task._d365Id}:`, error);
            return { success: false, taskId: task._d365Id, error: error.message };
        }
    }

    /**
     * Public: Push entire project back to D365
     */
    async function pushProjectToD365(project, projectId) {
        if (!config) throw new Error('D365 not configured');

        const results = {
            updated: [],
            failed: []
        };

        try {
            if (!project.tasks || !Array.isArray(project.tasks)) {
                throw new Error('Project must contain tasks array');
            }

            for (const task of project.tasks) {
                if (task._d365Id) {
                    const result = await pushTaskToD365(task, projectId);
                    if (result.success) {
                        results.updated.push(result.taskId);
                    } else {
                        results.failed.push(result);
                    }
                }
            }

            console.log(`Project sync complete: ${results.updated.length} updated, ${results.failed.length} failed`);
            return results;
        } catch (error) {
            console.error('Failed to push project:', error);
            throw error;
        }
    }

    /**
     * Public: Render setup wizard
     */
    // ── UI Helpers (dark-theme aware) ─────────────────────────
    function _el(tag, cls, text) {
        const e = document.createElement(tag);
        if (cls)  e.className   = cls;
        if (text) e.textContent = text;
        return e;
    }
    function _btn(label, primary, onClick) {
        const b = _el('button', primary ? 'btn btn-primary' : 'btn btn-secondary');
        b.textContent = label;
        b.style.marginTop = '6px';
        if (onClick) b.addEventListener('click', onClick);
        return b;
    }
    function _input(placeholder, value, type) {
        const i = _el('input', 'ms-input');
        i.type        = type || 'text';
        i.placeholder = placeholder;
        if (value) i.value = value;
        return i;
    }
    function _label(text) {
        const l = _el('label');
        l.textContent = text;
        l.style.cssText = 'font-size:0.72rem;font-weight:600;color:var(--text-muted);text-transform:uppercase;letter-spacing:.04em;display:block;margin:10px 0 4px';
        return l;
    }
    function _row(label, value, color) {
        const r = _el('div', 'nd-tip-row');
        const k = _el('span', 'nd-tip-key'); k.textContent = label;
        const v = _el('span', 'nd-tip-val'); v.textContent = value;
        if (color) v.style.color = color;
        r.appendChild(k); r.appendChild(v);
        return r;
    }

    // ══════════════════════════════════════════════════════════
    // UI RENDER FUNCTIONS  — dark-theme aware (uses CSS variables)
    // ══════════════════════════════════════════════════════════

    function renderSetupWizard(container, onComplete) {
        if (!container) return;
        container.innerHTML = '';

        // Step indicator
        let _step = 1;
        const cfg_local = _loadConfig() || {};

        const wrap = document.createElement('div');
        wrap.className = 'ms-wizard-step';

        // Step progress bar
        const steps = document.createElement('div');
        steps.style.cssText = 'display:flex;gap:6px;margin-bottom:18px;align-items:center';
        ['1. Mode','2. Credentials','3. Connect'].forEach((lbl, i) => {
            const pill = document.createElement('span');
            pill.textContent = lbl;
            pill.style.cssText = 'font-size:0.7rem;padding:3px 10px;border-radius:20px;border:1px solid var(--border-color);color:var(--text-muted)';
            pill.dataset.stepPill = i + 1;
            steps.appendChild(pill);
            if (i < 2) {
                const sep = document.createElement('span');
                sep.textContent = '›';
                sep.style.color = 'var(--text-muted)';
                steps.appendChild(sep);
            }
        });
        wrap.appendChild(steps);

        const body = document.createElement('div');
        body.id = 'd365WizardBody';
        wrap.appendChild(body);

        container.appendChild(wrap);

        function _highlightStep(n) {
            wrap.querySelectorAll('[data-step-pill]').forEach(p => {
                const active = parseInt(p.dataset.stepPill) === n;
                p.style.background     = active ? 'var(--accent-primary)'   : 'transparent';
                p.style.borderColor    = active ? 'var(--accent-primary)'   : 'var(--border-color)';
                p.style.color          = active ? '#fff'                    : 'var(--text-muted)';
            });
        }

        function _renderStep(n) {
            _step = n;
            _highlightStep(n);
            body.innerHTML = '';

            // ── Step 1: Choose mode ───────────────────────────
            if (n === 1) {
                const h = _el('h4', null, '🏢 Choose Dynamics 365 Mode');
                h.style.cssText = 'font-size:0.95rem;font-weight:600;color:var(--text-primary);margin:0 0 6px';
                body.appendChild(h);

                const p = _el('p', null, 'Which environment are you connecting to?');
                p.style.cssText = 'font-size:0.78rem;color:var(--text-muted);margin:0 0 14px';
                body.appendChild(p);

                const grid = document.createElement('div');
                grid.style.cssText = 'display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:14px';

                const makeCard = (icon, title, sub, mode) => {
                    const card = document.createElement('button');
                    card.style.cssText = 'background:var(--bg-tertiary);border:1px solid var(--border-color);border-radius:10px;padding:14px 12px;cursor:pointer;text-align:left;transition:border-color .15s';
                    card.onmouseover = () => { card.style.borderColor = 'var(--accent-primary)'; };
                    card.onmouseout  = () => { card.style.borderColor = 'var(--border-color)'; };
                    const ic = _el('div', null, icon);
                    ic.style.cssText = 'font-size:1.4rem;margin-bottom:6px';
                    const tt = _el('div', null, title);
                    tt.style.cssText = 'font-size:0.78rem;font-weight:600;color:var(--text-primary);margin-bottom:3px';
                    const st = _el('div', null, sub);
                    st.style.cssText = 'font-size:0.66rem;color:var(--text-muted)';
                    card.appendChild(ic); card.appendChild(tt); card.appendChild(st);
                    card.addEventListener('click', () => {
                        cfg_local.mode = mode; _saveConfig(cfg_local); _renderStep(2);
                    });
                    return card;
                };

                grid.appendChild(makeCard('⚙️', 'Project Operations', 'Dataverse API · org.crm.dynamics.com', MODES.OPERATIONS));
                grid.appendChild(makeCard('💼', 'Finance & Operations', 'OData API · org.operations.dynamics.com', MODES.FINANCE));
                body.appendChild(grid);

                const note = _el('p', null, '💡 Not sure? Most new D365 implementations use Project Operations.');
                note.style.cssText = 'font-size:0.7rem;color:var(--text-muted);padding:8px 10px;background:var(--bg-tertiary);border-radius:6px;margin:0';
                body.appendChild(note);

            // ── Step 2: Credentials ───────────────────────────
            } else if (n === 2) {
                const h = _el('h4', null, '🔑 Azure AD Configuration');
                h.style.cssText = 'font-size:0.95rem;font-weight:600;color:var(--text-primary);margin:0 0 4px';
                body.appendChild(h);

                const modeName = cfg_local.mode === MODES.OPERATIONS ? 'Project Operations' : 'Finance & Operations';
                const badge = _el('span', null, modeName);
                badge.style.cssText = 'font-size:0.65rem;padding:2px 8px;border-radius:10px;font-weight:600;background:rgba(99,102,241,0.15);color:#818cf8;margin-bottom:14px;display:inline-block';
                body.appendChild(badge);

                const urlPh = cfg_local.mode === MODES.OPERATIONS
                    ? 'https://yourorg.crm.dynamics.com'
                    : 'https://yourorg.operations.dynamics.com';

                body.appendChild(_label('Environment URL'));
                const urlInput = _input(urlPh, cfg_local.environmentUrl || '');
                body.appendChild(urlInput);

                body.appendChild(_label('Azure App Client ID'));
                const clientInput = _input('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx', cfg_local.clientId || '');
                body.appendChild(clientInput);

                body.appendChild(_label('Azure Tenant ID'));
                const tenantInput = _input('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx', cfg_local.tenantId || '');
                body.appendChild(tenantInput);

                const helpLink = document.createElement('a');
                helpLink.textContent = '📖 How to create an Azure App Registration →';
                helpLink.href = 'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/walkthrough-register-app-azure-active-directory';
                helpLink.target = '_blank';
                helpLink.style.cssText = 'font-size:0.7rem;color:var(--accent-primary);display:block;margin:10px 0 14px';
                body.appendChild(helpLink);

                const btnRow = document.createElement('div');
                btnRow.style.cssText = 'display:flex;gap:8px;margin-top:4px';
                btnRow.appendChild(_btn('← Back', false, () => _renderStep(1)));

                const nextBtn = _btn('Next →', true, () => {
                    const url = urlInput.value.trim();
                    const cid = clientInput.value.trim();
                    const tid = tenantInput.value.trim();
                    if (!url || !cid || !tid) {
                        const err = _el('p', null, '⚠ All fields are required');
                        err.style.cssText = 'color:#ef4444;font-size:0.75rem;margin:6px 0 0';
                        btnRow.after(err);
                        setTimeout(() => err.remove(), 3000);
                        return;
                    }
                    cfg_local.environmentUrl = url;
                    cfg_local.clientId       = cid;
                    cfg_local.tenantId       = tid;
                    _saveConfig(cfg_local);
                    _renderStep(3);
                });
                btnRow.appendChild(nextBtn);
                body.appendChild(btnRow);

            // ── Step 3: Sign in + Select project ─────────────
            } else if (n === 3) {
                const h = _el('h4', null, '🔐 Connect & Select Project');
                h.style.cssText = 'font-size:0.95rem;font-weight:600;color:var(--text-primary);margin:0 0 14px';
                body.appendChild(h);

                const status = _el('div', null, '');
                status.style.cssText = 'font-size:0.78rem;color:var(--text-muted);margin-bottom:12px;min-height:20px';
                body.appendChild(status);

                const signBtn = _btn('🔑  Sign in with Microsoft', true);
                signBtn.style.cssText += ';width:100%;justify-content:center';
                body.appendChild(signBtn);

                const backBtn = _btn('← Back', false, () => _renderStep(2));
                backBtn.style.marginTop = '8px';
                body.appendChild(backBtn);

                signBtn.addEventListener('click', async () => {
                    signBtn.disabled = true;
                    signBtn.textContent = '⏳  Signing in…';
                    status.textContent  = '';

                    try {
                        await signIn();
                        status.textContent = '✅  Signed in. Fetching projects…';
                        const projects = await getProjects();

                        body.innerHTML = '';
                        if (!projects.length) {
                            body.appendChild(_el('p', null, '⚠ No projects found in your D365 environment.'));
                            body.appendChild(_btn('← Back', false, () => _renderStep(2)));
                            return;
                        }

                        body.appendChild(_el('h4', null, '📁 Select a Project')).style.cssText =
                            'font-size:0.88rem;font-weight:600;color:var(--text-primary);margin:0 0 10px';

                        const ul = document.createElement('div');
                        ul.style.cssText = 'display:flex;flex-direction:column;gap:5px;max-height:260px;overflow-y:auto;margin-bottom:12px';

                        projects.forEach(proj => {
                            const item = document.createElement('button');
                            item.style.cssText = 'background:var(--bg-tertiary);border:1px solid var(--border-color);border-radius:7px;padding:10px 12px;cursor:pointer;text-align:left;transition:border-color .15s';
                            item.onmouseover = () => { item.style.borderColor = 'var(--accent-primary)'; };
                            item.onmouseout  = () => { item.style.borderColor = 'var(--border-color)'; };

                            const name = _el('div', null, proj.name);
                            name.style.cssText = 'font-size:0.78rem;font-weight:600;color:var(--text-primary)';
                            const meta = _el('div', null, proj.id);
                            meta.style.cssText = 'font-size:0.65rem;color:var(--text-muted);margin-top:2px';
                            item.appendChild(name); item.appendChild(meta);

                            item.addEventListener('click', () => {
                                onComplete({ projectId: proj.id, projectName: proj.name, mode: cfg_local.mode });
                            });
                            ul.appendChild(item);
                        });

                        body.appendChild(ul);
                        body.appendChild(_btn('← Back', false, () => _renderStep(2)));

                    } catch (err) {
                        signBtn.disabled = false;
                        signBtn.textContent = '🔑  Sign in with Microsoft';
                        status.style.color = '#ef4444';
                        status.textContent = '✗ ' + err.message;
                    }
                });
            }
        }

        _renderStep(1);
    }

    // ─────────────────────────────────────────────────────────
    function renderSyncPanel(container, pfProject, d365ProjectId) {
        if (!container) return;
        container.innerHTML = '';

        const panel = document.createElement('div');
        panel.className = 'ms-sync-panel';

        // Header info
        const hdr = document.createElement('div');
        hdr.className = 'ms-sync-header';

        const acc = document.createElement('div');
        acc.className = 'ms-sync-account';

        const envLine = _el('div', 'ms-sync-name', config ? config.environmentUrl : 'Not configured');
        const projLine = _el('div', 'ms-sync-email', pfProject ? pfProject.name : 'No project');
        acc.appendChild(envLine); acc.appendChild(projLine);

        const modeBadge = _el('span', null, config && config.mode === MODES.OPERATIONS ? 'Operations' : 'Finance');
        modeBadge.className = config && config.mode === MODES.OPERATIONS ? 'ms-badge ms-badge-operations' : 'ms-badge ms-badge-finance';

        hdr.appendChild(acc); hdr.appendChild(modeBadge);
        panel.appendChild(hdr);

        // Last sync
        const lastSync = _el('div', null,
            '🕐 Last sync: ' + (pfProject && pfProject._lastSync
                ? new Date(pfProject._lastSync).toLocaleString() : 'Never'));
        lastSync.style.cssText = 'font-size:0.7rem;color:var(--text-muted);padding:4px 2px';
        panel.appendChild(lastSync);

        // Action buttons
        const btns = document.createElement('div');
        btns.className = 'ms-sync-actions';

        const log = document.createElement('div');
        log.className = 'ms-sync-log';

        const addLog = (msg, ok) => {
            const e = _el('div', 'ms-sync-log-entry ' + (ok ? 'ok' : 'err'),
                new Date().toLocaleTimeString() + '  ' + msg);
            log.prepend(e);
        };

        const pullBtn = _btn('⬇ Pull from D365', true);
        pullBtn.addEventListener('click', async () => {
            pullBtn.disabled = true; pullBtn.textContent = '⏳ Pulling…';
            try {
                await importProject(d365ProjectId);
                if (pfProject) pfProject._lastSync = new Date();
                addLog('Pull complete', true);
            } catch(e) { addLog('Pull failed: ' + e.message, false); }
            finally { pullBtn.disabled = false; pullBtn.textContent = '⬇ Pull from D365'; }
        });

        const pushBtn = _btn('⬆ Push to D365', false);
        pushBtn.addEventListener('click', async () => {
            pushBtn.disabled = true; pushBtn.textContent = '⏳ Pushing…';
            try {
                const r = await pushProjectToD365(pfProject, d365ProjectId);
                if (pfProject) pfProject._lastSync = new Date();
                addLog(`Pushed ${(r.updated||[]).length} tasks`, true);
            } catch(e) { addLog('Push failed: ' + e.message, false); }
            finally { pushBtn.disabled = false; pushBtn.textContent = '⬆ Push to D365'; }
        });

        const kpiBtn = _btn('📊 Financial KPIs', false);
        kpiBtn.addEventListener('click', async () => {
            kpiBtn.disabled = true; kpiBtn.textContent = '⏳ Loading…';
            try {
                const kpis = await getAccountingKPIs(d365ProjectId);
                renderAccountingPanel(panel, d365ProjectId, kpis);
                addLog('KPIs loaded', true);
            } catch(e) { addLog('KPIs failed: ' + e.message, false); }
            finally { kpiBtn.disabled = false; kpiBtn.textContent = '📊 Financial KPIs'; }
        });

        const signOutBtn = _btn('Sign Out', false, async () => {
            await signOut();
            container.innerHTML = '';
            renderSetupWizard(container, () => {});
        });
        signOutBtn.style.marginLeft = 'auto';

        [pullBtn, pushBtn, kpiBtn, signOutBtn].forEach(b => btns.appendChild(b));
        panel.appendChild(btns);
        panel.appendChild(log);
        container.appendChild(panel);
    }

    // ─────────────────────────────────────────────────────────
    function renderAccountingPanel(container, projectId, kpis) {
        if (!container || !kpis) return;

        const section = document.createElement('div');
        section.style.cssText = 'margin-top:16px;border-top:1px solid var(--border-color);padding-top:14px';

        const h = _el('div', null, '📊 Financial Dashboard');
        h.style.cssText = 'font-size:0.82rem;font-weight:600;color:var(--text-primary);margin-bottom:10px';
        section.appendChild(h);

        // KPI grid
        const grid = document.createElement('div');
        grid.className = 'd365-kpi-grid';

        const fmt = v => isNaN(v) ? '—' : '$' + Math.round(v).toLocaleString();
        const fmtR = v => isNaN(v) ? '—' : v.toFixed(2);

        const kpiItems = [
            { label: 'Budget Cost',     val: fmt(kpis.budgetCost),     color: '' },
            { label: 'Actual Cost',     val: fmt(kpis.actualCost),     color: '' },
            { label: 'Cost Variance',   val: fmt(kpis.costVariance),   color: kpis.costVariance >= 0 ? 'd365-good' : 'd365-bad' },
            { label: 'CPI',             val: fmtR(kpis.cpi),           color: kpis.cpi >= 1 ? 'd365-good' : kpis.cpi >= 0.8 ? 'd365-warn' : 'd365-bad' },
            { label: 'Budget Revenue',  val: fmt(kpis.budgetRevenue),  color: '' },
            { label: 'SPI',             val: fmtR(kpis.spi),           color: kpis.spi >= 1 ? 'd365-good' : kpis.spi >= 0.8 ? 'd365-warn' : 'd365-bad' },
        ];

        kpiItems.forEach(item => {
            const card = _el('div', 'd365-kpi');
            const val  = _el('div', 'd365-kpi-val ' + item.color, item.val);
            const lbl  = _el('div', 'd365-kpi-label', item.label);
            card.appendChild(val); card.appendChild(lbl);
            grid.appendChild(card);
        });
        section.appendChild(grid);

        // Budget vs Actual bar chart (canvas)
        const chartWrap = document.createElement('div');
        chartWrap.style.cssText = 'background:var(--bg-tertiary);border-radius:8px;padding:12px;margin-top:10px';
        const chartTitle = _el('div', null, 'Budget vs Actual Cost');
        chartTitle.style.cssText = 'font-size:0.72rem;font-weight:600;color:var(--text-muted);margin-bottom:8px';
        chartWrap.appendChild(chartTitle);

        const canvas = document.createElement('canvas');
        canvas.height = 70;
        canvas.style.cssText = 'width:100%;border-radius:4px';
        chartWrap.appendChild(canvas);
        section.appendChild(chartWrap);
        container.appendChild(section);

        // Draw chart after appended to DOM
        requestAnimationFrame(() => {
            canvas.width = Math.floor(canvas.parentElement.offsetWidth - 24);
            const ctx = canvas.getContext('2d');
            const W = canvas.width, H = 70;
            ctx.clearRect(0, 0, W, H);

            const max = Math.max(kpis.budgetCost || 1, kpis.actualCost || 1, 1);
            const bw  = Math.max(0, Math.floor((kpis.budgetCost / max) * (W - 80)));
            const aw  = Math.max(0, Math.floor((kpis.actualCost  / max) * (W - 80)));
            const r   = 4;

            // Budget bar
            ctx.fillStyle = 'rgba(99,102,241,0.7)';
            ctx.beginPath(); ctx.roundRect(0, 8, bw, 22, r); ctx.fill();
            ctx.fillStyle = 'rgba(255,255,255,0.5)'; ctx.font = '10px Inter,sans-serif';
            ctx.textBaseline = 'middle';
            ctx.fillText('Budget  ' + fmt(kpis.budgetCost), bw + 6, 19);

            // Actual bar
            const actColor = kpis.actualCost <= kpis.budgetCost ? 'rgba(34,197,94,0.7)' : 'rgba(239,68,68,0.7)';
            ctx.fillStyle = actColor;
            ctx.beginPath(); ctx.roundRect(0, 42, aw, 22, r); ctx.fill();
            ctx.fillStyle = 'rgba(255,255,255,0.5)';
            ctx.fillText('Actual  ' + fmt(kpis.actualCost), aw + 6, 53);
        });
    }

    // Public API
    export const D365Client = {
        MODES,
        configure,
        signIn,
        signOut,
        isAuthenticated,
        getProjects,
        getProjectTasks,
        getResources,
        getAssignments,
        getBudgetLines,
        getTransactions,
        getAccountingKPIs,
        importProject,
        pushTaskToD365,
        pushProjectToD365,
        renderSetupWizard,
        renderSyncPanel,
        renderAccountingPanel
    };
