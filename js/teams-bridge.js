/**
 * ProjectFlow™ — Teams Bridge
 * © 2026 Ahmed M. Fawzy
 *
 * Auto-connects to Microsoft Planner & D365 when running inside Teams.
 * Also works in standalone browser with manual sign-in.
 */

import { MSGraphClient } from './ms-graph.js';
import { D365Client } from './d365.js';
import { EventBus } from './event-bus.js';

// ============================================================================
// CONFIGURATION
// ============================================================================

const STORAGE_KEY = 'pf_teams_config';
let _isInTeams = false;
let _teamsContext = null;

// ============================================================================
// TEAMS DETECTION
// ============================================================================

/**
 * Detect if we're running inside Microsoft Teams
 */
async function detectTeams() {
    try {
        if (typeof microsoftTeams === 'undefined') return false;

        await microsoftTeams.app.initialize();
        const ctx = await microsoftTeams.app.getContext();
        _teamsContext = ctx;
        _isInTeams = true;
        console.log('[TeamsBridge] Running inside Teams ✓');
        return true;
    } catch (e) {
        console.log('[TeamsBridge] Not in Teams — standalone mode');
        _isInTeams = false;
        return false;
    }
}

function isInTeams() { return _isInTeams; }
function getTeamsContext() { return _teamsContext; }

// ============================================================================
// SAVED CONFIG
// ============================================================================

function getSavedConfig() {
    try {
        return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
    } catch { return {}; }
}

function saveConfig(cfg) {
    const current = getSavedConfig();
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ ...current, ...cfg }));
}

// ============================================================================
// AUTO-CONNECT: PLANNER
// ============================================================================

/**
 * Auto-connect to Microsoft Planner using silent auth (no popup, no manual config).
 * Uses the built-in multi-tenant Client ID — works for any Microsoft 365 org.
 * @returns {Object|null} - { plans, authenticated } or null if user must sign in
 */
async function autoConnectPlanner() {
    try {
        // Try silent sign-in (uses cached MSAL session from localStorage)
        const silentOk = await MSGraphClient.trySilentSignIn();

        if (!silentOk) {
            console.log('[TeamsBridge] No cached session — user must sign in manually');
            EventBus.emit('planner:needs-signin', {});
            return null;
        }

        console.log('[TeamsBridge] Planner: Silent auth OK ✓');

        // Fetch all plans the user has access to
        const plans = await MSGraphClient.getMyPlans();
        console.log(`[TeamsBridge] Planner: Found ${plans.length} plans ✓`);

        // Restore last-used plan if available
        const cfg = getSavedConfig();
        EventBus.emit('planner:connected', { plans, lastPlanId: cfg.lastPlanId || null });
        return { plans, authenticated: true };

    } catch (err) {
        console.error('[TeamsBridge] Planner auto-connect failed:', err);
        EventBus.emit('planner:needs-signin', { error: err.message });
        return null;
    }
}

/**
 * Save the last used plan ID so it auto-loads on next startup
 */
function saveLastPlan(planId) {
    saveConfig({ lastPlanId: planId });
}

// ============================================================================
// AUTO-CONNECT: D365
// ============================================================================

/**
 * Auto-connect to Dynamics 365 if config exists
 * @returns {Object|null} - imported projects info or null
 */
async function autoConnectD365() {
    const cfg = getSavedConfig();
    if (!cfg.d365Url || !cfg.d365ClientId) {
        console.log('[TeamsBridge] No D365 config saved — skipping auto-connect');
        return null;
    }

    try {
        D365Client.configure({
            environmentUrl: cfg.d365Url,
            clientId: cfg.d365ClientId,
            tenantId: cfg.tenantId,
            mode: cfg.d365Mode || D365Client.MODES.OPERATIONS
        });

        if (D365Client.isAuthenticated()) {
            console.log('[TeamsBridge] D365: Already authenticated ✓');
        } else {
            console.log('[TeamsBridge] D365: Not authenticated — will need manual sign-in');
            return null;
        }

        const projects = await D365Client.getProjects();
        console.log(`[TeamsBridge] D365: Found ${projects.length} projects ✓`);

        EventBus.emit('d365:connected', { projects });
        return { projects, authenticated: true };

    } catch (err) {
        console.error('[TeamsBridge] D365 auto-connect failed:', err);
        return null;
    }
}

// ============================================================================
// SETUP WIZARD HELPER
// ============================================================================

/**
 * Render the initial setup form for Teams integration
 */
function renderConnectionSetup(container) {
    if (!container) return;

    const cfg = getSavedConfig();

    container.innerHTML = `
        <div style="padding:24px; max-width:520px; margin:0 auto; font-family:Inter,sans-serif;">
            <h2 style="margin:0 0 8px; color:#e2e8f0;">🔗 Connect to Microsoft Services</h2>
            <p style="color:#94a3b8; margin:0 0 24px; font-size:14px;">
                Enter your Azure AD App Registration details to connect Planner & D365.
            </p>

            <div style="margin-bottom:20px;">
                <h3 style="color:#c4b5fd; margin:0 0 12px; font-size:15px;">Azure AD Settings</h3>
                <label style="display:block; color:#94a3b8; font-size:12px; margin-bottom:4px;">Client ID (Application ID)</label>
                <input id="tbClientId" type="text" value="${cfg.clientId || ''}" 
                    placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                    style="width:100%; padding:10px; background:#1e293b; border:1px solid #334155; border-radius:6px; color:#e2e8f0; font-size:14px; margin-bottom:12px; box-sizing:border-box;">
                
                <label style="display:block; color:#94a3b8; font-size:12px; margin-bottom:4px;">Tenant ID</label>
                <input id="tbTenantId" type="text" value="${cfg.tenantId || ''}" 
                    placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                    style="width:100%; padding:10px; background:#1e293b; border:1px solid #334155; border-radius:6px; color:#e2e8f0; font-size:14px; margin-bottom:12px; box-sizing:border-box;">
            </div>

            <div style="margin-bottom:20px;">
                <h3 style="color:#c4b5fd; margin:0 0 12px; font-size:15px;">D365 Settings (Optional)</h3>
                <label style="display:block; color:#94a3b8; font-size:12px; margin-bottom:4px;">D365 Environment URL</label>
                <input id="tbD365Url" type="text" value="${cfg.d365Url || ''}" 
                    placeholder="https://yourorg.crm.dynamics.com"
                    style="width:100%; padding:10px; background:#1e293b; border:1px solid #334155; border-radius:6px; color:#e2e8f0; font-size:14px; margin-bottom:12px; box-sizing:border-box;">

                <label style="display:block; color:#94a3b8; font-size:12px; margin-bottom:4px;">D365 Client ID (if different)</label>
                <input id="tbD365ClientId" type="text" value="${cfg.d365ClientId || ''}" 
                    placeholder="Same as above if not specified"
                    style="width:100%; padding:10px; background:#1e293b; border:1px solid #334155; border-radius:6px; color:#e2e8f0; font-size:14px; box-sizing:border-box;">
            </div>

            <div style="display:flex; gap:12px;">
                <button id="btnSaveTeamsConfig" style="flex:1; padding:12px; background:linear-gradient(135deg,#6366f1,#8b5cf6); color:white; border:none; border-radius:8px; font-size:14px; font-weight:600; cursor:pointer; transition:all 0.2s;">
                    💾 Save & Connect
                </button>
                <button id="btnConnectPlanner" style="flex:1; padding:12px; background:linear-gradient(135deg,#0ea5e9,#06b6d4); color:white; border:none; border-radius:8px; font-size:14px; font-weight:600; cursor:pointer; transition:all 0.2s;">
                    📋 Connect Planner
                </button>
            </div>

            <div id="connectionStatus" style="margin-top:16px; padding:12px; background:#1e293b; border-radius:6px; color:#94a3b8; font-size:13px; display:none;"></div>
        </div>
    `;

    // Save config
    document.getElementById('btnSaveTeamsConfig')?.addEventListener('click', () => {
        const clientId = document.getElementById('tbClientId')?.value.trim();
        const tenantId = document.getElementById('tbTenantId')?.value.trim();
        const d365Url = document.getElementById('tbD365Url')?.value.trim();
        const d365ClientId = document.getElementById('tbD365ClientId')?.value.trim() || clientId;

        if (!clientId || !tenantId) {
            _showStatus('⚠️ Client ID and Tenant ID are required', 'warning');
            return;
        }

        saveConfig({ clientId, tenantId, d365Url, d365ClientId });
        _showStatus('✅ Configuration saved! Click "Connect Planner" to sign in.', 'success');
    });

    // Connect Planner
    document.getElementById('btnConnectPlanner')?.addEventListener('click', async () => {
        const clientId = document.getElementById('tbClientId')?.value.trim();
        const tenantId = document.getElementById('tbTenantId')?.value.trim();

        if (!clientId || !tenantId) {
            _showStatus('⚠️ Enter Client ID and Tenant ID first', 'warning');
            return;
        }

        saveConfig({ clientId, tenantId });
        _showStatus('🔄 Connecting to Microsoft Planner...', 'info');

        try {
            await MSGraphClient.configure(clientId, tenantId);
            await MSGraphClient.signIn();
            const plans = await MSGraphClient.getMyPlans();
            _showStatus(`✅ Connected! Found ${plans.length} plans. Reload the app to see them.`, 'success');
            EventBus.emit('planner:connected', { plans });
        } catch (err) {
            _showStatus(`❌ Connection failed: ${err.message}`, 'error');
        }
    });

    function _showStatus(msg, type) {
        const el = document.getElementById('connectionStatus');
        if (!el) return;
        el.style.display = 'block';
        el.textContent = msg;
        el.style.borderLeft = `3px solid ${
            type === 'success' ? '#22c55e' :
            type === 'error' ? '#ef4444' :
            type === 'warning' ? '#f59e0b' : '#6366f1'
        }`;
    }
}

// ============================================================================
// MAIN INIT
// ============================================================================

/**
 * Initialize Teams Bridge — call this on app start
 */
async function init() {
    await detectTeams();

    // Auto-connect if config exists
    const [plannerResult, d365Result] = await Promise.allSettled([
        autoConnectPlanner(),
        autoConnectD365()
    ]);

    return {
        isInTeams: _isInTeams,
        planner: plannerResult.status === 'fulfilled' ? plannerResult.value : null,
        d365: d365Result.status === 'fulfilled' ? d365Result.value : null
    };
}

// ============================================================================
// EXPORT
// ============================================================================

export const TeamsBridge = {
    init,
    isInTeams,
    getTeamsContext,
    getSavedConfig,
    saveConfig,
    saveLastPlan,
    autoConnectPlanner,
    autoConnectD365,
    renderConnectionSetup
};
