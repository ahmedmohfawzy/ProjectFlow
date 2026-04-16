/**
 * ProjectFlow — Local SQLite Backend Server
 * Usage: node server.js [--port 3456] [--data /path/to/data]
 */
const express = require('express');
const cors = require('cors');
const Database = require('better-sqlite3');
const path = require('path');
const fs = require('fs');

// ── Config ──
const CONFIG_PATH = path.join(__dirname, 'server-config.json');
let config = { dataDir: '', port: 3456, activeDatabase: 'default.db' };

function loadConfig() {
    try {
        if (fs.existsSync(CONFIG_PATH)) {
            config = { ...config, ...JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8')) };
        }
    } catch (e) { console.warn('Config load failed:', e.message); }
    // Default dataDir = ./data
    if (!config.dataDir) config.dataDir = path.join(__dirname, 'data');
    // CLI overrides
    const args = process.argv.slice(2);
    const portIdx = args.indexOf('--port');
    if (portIdx > -1 && args[portIdx + 1]) config.port = parseInt(args[portIdx + 1]);
    const dataIdx = args.indexOf('--data');
    if (dataIdx > -1 && args[dataIdx + 1]) config.dataDir = path.resolve(args[dataIdx + 1]);
}

function saveConfig() {
    try { fs.writeFileSync(CONFIG_PATH, JSON.stringify(config, null, 2)); }
    catch (e) { console.warn('Config save failed:', e.message); }
}

loadConfig();

// Ensure data directory exists
if (!fs.existsSync(config.dataDir)) {
    fs.mkdirSync(config.dataDir, { recursive: true });
    console.log('📁 Created data directory:', config.dataDir);
}

// ── Database helpers ──
const openDBs = new Map(); // Cache open database connections

function getDB(name) {
    if (!name || name.includes('..') || name.includes('/')) throw new Error('Invalid database name');
    if (!name.endsWith('.db')) name += '.db';

    if (openDBs.has(name)) return openDBs.get(name);

    const dbPath = path.join(config.dataDir, name);
    const db = new Database(dbPath);
    db.pragma('journal_mode = WAL');
    db.pragma('foreign_keys = ON');

    // Create tables if they don't exist
    db.exec(`
        CREATE TABLE IF NOT EXISTS projects_meta (
            id TEXT PRIMARY KEY,
            name TEXT NOT NULL DEFAULT 'Untitled',
            color TEXT DEFAULT '#6366f1',
            pinned INTEGER DEFAULT 0,
            archived INTEGER DEFAULT 0,
            description TEXT DEFAULT '',
            task_count INTEGER DEFAULT 0,
            progress INTEGER DEFAULT 0,
            start_date TEXT,
            finish_date TEXT,
            last_modified TEXT,
            created_at TEXT DEFAULT (datetime('now'))
        );
        CREATE TABLE IF NOT EXISTS projects_data (
            id TEXT PRIMARY KEY,
            data TEXT NOT NULL,
            FOREIGN KEY (id) REFERENCES projects_meta(id) ON DELETE CASCADE
        );
    `);

    openDBs.set(name, db);
    return db;
}

function closeAllDBs() {
    for (const [name, db] of openDBs) {
        try { db.close(); } catch (e) {}
    }
    openDBs.clear();
}

// ── Express App ──
// Rate Limiter logic
const requestCounts = new Map();
setInterval(() => requestCounts.clear(), 60000); // clear every minute
const rateLimiter = (req, res, next) => {
    const ip = req.ip || req.connection.remoteAddress || 'unknown';
    const current = requestCounts.get(ip) || 0;
    if (current > 100) return res.status(429).json({ error: 'Too many requests. Please try again later.' });
    requestCounts.set(ip, current + 1);
    next();
};

const app = express();
app.use(cors({
    origin: function(origin, callback) {
        if (!origin) return callback(null, true);
        if (/^https?:\/\/(localhost|127\.0\.0\.1|0\.0\.0\.0|192\.168\.\d+\.\d+)(:\d+)?$/.test(origin)) {
            return callback(null, true);
        }
        // Fallback for dev: allow but log
        console.warn('CORS allowed origin dynamically:', origin);
        callback(null, true); 
    }
}));
app.use(express.json({ limit: '50mb' }));
app.use('/api', rateLimiter);

// Security headers
app.use((req, res, next) => {
    res.setHeader('X-Frame-Options', 'DENY');
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('X-XSS-Protection', '1; mode=block');
    res.setHeader('Strict-Transport-Security', 'max-age=31536000; includeSubDomains');
    res.setHeader('Referrer-Policy', 'strict-origin-when-cross-origin');
    res.setHeader('Content-Security-Policy', "default-src 'self'; script-src 'self' 'unsafe-inline' 'unsafe-eval' https://alcdn.msauth.net https://res.cdn.office.net; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com data:; img-src 'self' data: https:; connect-src 'self' https://date.nager.at https://fonts.googleapis.com https://fonts.gstatic.com https://alcdn.msauth.net https://login.microsoftonline.com https://graph.microsoft.com https://*.dynamics.com https://*.crm.dynamics.com; frame-ancestors 'self' https://teams.microsoft.com https://*.teams.microsoft.com https://*.office.com https://*.microsoft.com");
    next();
});

// index.html — always fresh (no-cache)
const NO_CACHE = { 'Cache-Control': 'no-store, no-cache, must-revalidate', 'Pragma': 'no-cache', 'Expires': '0' };
app.get('/',           (req, res) => { res.set(NO_CACHE); res.sendFile(path.join(__dirname, 'index.html')); });
app.get('/index.html', (req, res) => { res.set(NO_CACHE); res.sendFile(path.join(__dirname, 'index.html')); });
// sw.js — never cached by browser (SW has its own update mechanism)
app.get('/sw.js',      (req, res) => { res.set({ 'Cache-Control': 'no-store' }); res.sendFile(path.join(__dirname, 'sw.js')); });

// JS / CSS — no cache in development to avoid stale bugs
app.use('/js',  express.static(path.join(__dirname, 'js'),  { maxAge: '0' }));
app.use('/css', express.static(path.join(__dirname, 'css'), { maxAge: '0' }));

// Public assets (favicon, icons, manifest) — served at root like Vite
app.use(express.static(path.join(__dirname, 'public'), { maxAge: '0' }));

// Everything else (icons, manifest, etc.)
app.use(express.static(__dirname, { index: false }));

// ── API: Server Config ──
app.get('/api/config', (req, res) => {
    res.json({
        dataDir: config.dataDir,
        port: config.port,
        activeDatabase: config.activeDatabase,
        serverVersion: '1.0.0'
    });
});

app.put('/api/config', (req, res) => {
    try {
        const { dataDir, activeDatabase } = req.body;
        if (dataDir && typeof dataDir === 'string') {
            const resolved = path.resolve(dataDir);
            if (!fs.existsSync(resolved)) {
                fs.mkdirSync(resolved, { recursive: true });
            }
            // Close all open databases before switching
            closeAllDBs();
            config.dataDir = resolved;
        }
        if (activeDatabase) config.activeDatabase = activeDatabase;
        saveConfig();
        res.json({ success: true, config });
    } catch (e) {
        res.status(400).json({ error: e.message });
    }
});

// ── API: Database Files ──
app.get('/api/databases', (req, res) => {
    try {
        const files = fs.readdirSync(config.dataDir)
            .filter(f => f.endsWith('.db'))
            .map(f => {
                const stat = fs.statSync(path.join(config.dataDir, f));
                return {
                    name: f,
                    size: stat.size,
                    modified: stat.mtime.toISOString(),
                    active: f === config.activeDatabase
                };
            });
        res.json({ databases: files, dataDir: config.dataDir });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

app.post('/api/databases', (req, res) => {
    try {
        let { name } = req.body;
        if (!name) return res.status(400).json({ error: 'Name required' });
        name = name.replace(/[^a-zA-Z0-9_\-\. ]/g, '').trim();
        if (!name.endsWith('.db')) name += '.db';
        getDB(name); // Creates the database
        config.activeDatabase = name;
        saveConfig();
        res.json({ success: true, name });
    } catch (e) {
        res.status(400).json({ error: e.message });
    }
});

app.delete('/api/databases/:name', (req, res) => {
    try {
        let name = req.params.name;
        if (!name.endsWith('.db')) name += '.db';
        // Close if open
        if (openDBs.has(name)) {
            openDBs.get(name).close();
            openDBs.delete(name);
        }
        const dbPath = path.join(config.dataDir, name);
        if (fs.existsSync(dbPath)) fs.unlinkSync(dbPath);
        // Also remove WAL and SHM files
        if (fs.existsSync(dbPath + '-wal')) fs.unlinkSync(dbPath + '-wal');
        if (fs.existsSync(dbPath + '-shm')) fs.unlinkSync(dbPath + '-shm');
        if (config.activeDatabase === name) {
            config.activeDatabase = 'default.db';
            saveConfig();
        }
        res.json({ success: true });
    } catch (e) {
        res.status(400).json({ error: e.message });
    }
});

// ── API: Projects ──
app.get('/api/db/:dbname/projects', (req, res) => {
    try {
        const db = getDB(req.params.dbname);
        const projects = db.prepare('SELECT * FROM projects_meta ORDER BY pinned DESC, last_modified DESC').all();
        res.json(projects);
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

app.get('/api/db/:dbname/projects/:id', (req, res) => {
    try {
        const db = getDB(req.params.dbname);
        const meta = db.prepare('SELECT * FROM projects_meta WHERE id = ?').get(req.params.id);
        const data = db.prepare('SELECT data FROM projects_data WHERE id = ?').get(req.params.id);
        if (!meta) return res.status(404).json({ error: 'Project not found' });
        res.json({ meta, data: data ? JSON.parse(data.data) : null });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

app.post('/api/db/:dbname/projects/:id', (req, res) => {
    try {
        const db = getDB(req.params.dbname);
        const { meta, data } = req.body;
        const id = req.params.id;

        if (meta) {
            const stmt = db.prepare(`
                INSERT OR REPLACE INTO projects_meta 
                (id, name, color, pinned, archived, description, task_count, progress, start_date, finish_date, last_modified)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            `);
            stmt.run(
                id, meta.name || 'Untitled', meta.color || '#6366f1',
                meta.pinned ? 1 : 0, meta.archived ? 1 : 0, meta.description || '',
                meta.task_count || meta.taskCount || 0, meta.progress || 0,
                meta.start_date || meta.startDate || null,
                meta.finish_date || meta.finishDate || null,
                new Date().toISOString()
            );
        }

        if (data) {
            db.prepare('INSERT OR REPLACE INTO projects_data (id, data) VALUES (?, ?)')
                .run(id, JSON.stringify(data));
        }

        res.json({ success: true, id });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// Update meta only (for pin, color, archive, rename)
app.patch('/api/db/:dbname/projects/:id', (req, res) => {
    try {
        const db = getDB(req.params.dbname);
        const id = req.params.id;
        const updates = req.body;

        const existing = db.prepare('SELECT * FROM projects_meta WHERE id = ?').get(id);
        if (!existing) return res.status(404).json({ error: 'Not found' });

        const fields = ['name', 'color', 'pinned', 'archived', 'description', 'task_count', 'progress', 'start_date', 'finish_date'];
        const sets = [];
        const values = [];
        for (const f of fields) {
            if (updates[f] !== undefined) {
                sets.push(`${f} = ?`);
                values.push(f === 'pinned' || f === 'archived' ? (updates[f] ? 1 : 0) : updates[f]);
            }
        }
        if (sets.length > 0) {
            sets.push('last_modified = ?');
            values.push(new Date().toISOString());
            values.push(id);
            db.prepare(`UPDATE projects_meta SET ${sets.join(', ')} WHERE id = ?`).run(...values);
        }
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

app.delete('/api/db/:dbname/projects/:id', (req, res) => {
    try {
        const db = getDB(req.params.dbname);
        db.prepare('DELETE FROM projects_data WHERE id = ?').run(req.params.id);
        db.prepare('DELETE FROM projects_meta WHERE id = ?').run(req.params.id);
        res.json({ success: true });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// ── API: Native Folder Picker (macOS) ──
app.get('/api/pick-folder', (req, res) => {
    const { execSync } = require('child_process');
    try {
        const script = `osascript -e 'POSIX path of (choose folder with prompt "Choose Data Directory for ProjectFlow")'`;
        const result = execSync(script, { timeout: 60000, encoding: 'utf8' }).trim();
        if (result) {
            // Remove trailing slash
            const folderPath = result.endsWith('/') ? result.slice(0, -1) : result;
            res.json({ success: true, path: folderPath });
        } else {
            res.json({ success: false, error: 'No folder selected' });
        }
    } catch (e) {
        // User cancelled the dialog
        res.json({ success: false, error: 'Cancelled' });
    }
});

// ── API: Directory Browser ──
app.get('/api/browse', (req, res) => {
    try {
        let targetPath = req.query.path || require('os').homedir();
        targetPath = path.resolve(targetPath);

        // Security check for directory traversal
        const resolvedDataDir = path.resolve(config.dataDir);
        if (!targetPath.startsWith(resolvedDataDir)) {
            return res.status(403).json({ error: 'Access denied: Directory outside data directory' });
        }

        if (!fs.existsSync(targetPath)) {
            return res.status(404).json({ error: 'Path not found' });
        }

        const stat = fs.statSync(targetPath);
        if (!stat.isDirectory()) {
            targetPath = path.dirname(targetPath);
        }

        const entries = fs.readdirSync(targetPath, { withFileTypes: true })
            .filter(e => e.isDirectory() && !e.name.startsWith('.'))
            .map(e => ({
                name: e.name,
                path: path.join(targetPath, e.name)
            }))
            .sort((a, b) => a.name.localeCompare(b.name));

        res.json({
            current: targetPath,
            parent: path.dirname(targetPath),
            folders: entries
        });
    } catch (e) {
        res.status(500).json({ error: e.message });
    }
});

// ── API: Shutdown ──
app.post('/api/shutdown', (req, res) => {
    res.json({ success: true, message: 'Server shutting down...' });
    console.log('\n🛑 Shutdown requested from browser');
    closeAllDBs();
    setTimeout(() => process.exit(0), 500);
});

// ── API: Health check ──
app.get('/api/ping', (req, res) => res.json({ ok: true, time: new Date().toISOString() }));

// ── Start ──
const server = app.listen(config.port, () => {
    console.log('');
    console.log('  ╔═══════════════════════════════════════╗');
    console.log('  ║     🚀 ProjectFlow Server Running     ║');
    console.log('  ╠═══════════════════════════════════════╣');
    console.log(`  ║  URL:  http://localhost:${config.port}          ║`);
    console.log(`  ║  Data: ${config.dataDir}`);
    console.log(`  ║  DB:   ${config.activeDatabase}`);
    console.log('  ╚═══════════════════════════════════════╝');
    console.log('');
    // Ensure default database exists
    getDB(config.activeDatabase);
});

// Graceful shutdown
process.on('SIGINT', () => {
    console.log('\n🛑 Shutting down...');
    closeAllDBs();
    server.close();
    process.exit(0);
});
