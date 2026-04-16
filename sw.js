/**
 * ProjectFlow™ — Service Worker
 * Sprint E.1 — PWA + Offline Support
 * © 2026 Ahmed M. Fawzy. All Rights Reserved.
 */
'use strict';

const CACHE_NAME    = 'projectflow-v5';   // ← bumped: force old SW to expire
const CACHE_VERSION = '5.0.0';

// Assets to cache on install (App Shell)
const SHELL_ASSETS = [
    '/',
    '/index.html',
    '/css/style.css',
    '/js/main.js',
    '/js/app.js',
    '/js/event-bus.js',
    '/js/state-manager.js',
    '/js/project-io.js',
    '/js/project-analytics.js',
    '/js/board.js',
    '/js/scenarios.js',
    '/js/plugins.js',
    '/js/gantt.js',
    '/js/dashboard.js',
    '/js/critical-path.js',
    '/js/resource-manager.js',
    '/js/calendar.js',
    '/js/evm.js',
    '/js/network.js',
    '/js/reports.js',
    '/js/xml-parser.js',
    '/js/planner-parser.js',
    '/js/ui-helpers.js',
    '/js/d365.js',
    '/js/ms-graph.js',
    '/js/task-editor.js',
    '/js/lib/dexie.min.js',
    '/js/lib/jspdf.umd.min.js',
    '/js/lib/xlsx.full.min.js',
    '/favicon.ico',
    '/icons/icon-192.png',
    '/manifest.json'
];

// ── Install: cache app shell ──────────────────────────────────
self.addEventListener('install', (event) => {
    console.log(`[SW] Installing v${CACHE_VERSION}`);
    event.waitUntil(
        caches.open(CACHE_NAME)
            .then(cache => cache.addAll(SHELL_ASSETS).catch(err => {
                console.warn('[SW] Some assets failed to cache:', err);
            }))
            .then(() => self.skipWaiting())
    );
});

// ── Activate: clean old caches ───────────────────────────────
self.addEventListener('activate', (event) => {
    console.log('[SW] Activating');
    event.waitUntil(
        caches.keys().then(keys =>
            Promise.all(keys
                .filter(k => k !== CACHE_NAME)
                .map(k => { console.log('[SW] Deleting old cache:', k); return caches.delete(k); })
            )
        ).then(() => self.clients.claim())
    );
});

// ── Fetch strategy ────────────────────────────────────────────
self.addEventListener('fetch', (event) => {
    const { request } = event;
    const url = new URL(request.url);

    if (request.method !== 'GET') return;
    if (!url.protocol.startsWith('http')) return;

    // 1. HTML navigation → ALWAYS network-first (never serve stale HTML)
    if (request.mode === 'navigate' || url.pathname.endsWith('.html') || url.pathname === '/') {
        event.respondWith(
            fetch(request, { cache: 'no-store' })
                .catch(() => caches.match('/index.html'))
        );
        return;
    }

    // 2. API calls → network-first
    if (url.pathname.startsWith('/api/')) {
        event.respondWith(fetch(request).catch(() => caches.match(request)));
        return;
    }

    // 3. JS/CSS assets → cache-first (they change with cache version bump)
    event.respondWith(
        caches.match(request).then(cached => {
            if (cached) return cached;
            return fetch(request).then(response => {
                if (response.ok) {
                    const cloned = response.clone();
                    caches.open(CACHE_NAME).then(c => c.put(request, cloned));
                }
                return response;
            }).catch(() => new Response('Offline', { status: 503 }));
        })
    );
});

// ── Push Notifications (E.1) ──────────────────────────────────
self.addEventListener('push', (event) => {
    if (!event.data) return;
    const data = event.data.json().catch(() => ({ title: 'ProjectFlow', body: event.data.text() }));
    event.waitUntil(
        data.then(d => self.registration.showNotification(d.title || 'ProjectFlow', {
            body:    d.body || '',
            icon:    '/icons/icon-192.png',
            badge:   '/icons/icon-192.png',
            tag:     d.tag || 'pf-notification',
            vibrate: [100, 50, 100],
            data:    { url: d.url || '/' }
        }))
    );
});

self.addEventListener('notificationclick', (event) => {
    event.notification.close();
    const targetUrl = (event.notification.data || {}).url || '/';
    event.waitUntil(
        clients.matchAll({ type: 'window' }).then(clientList => {
            for (const client of clientList) {
                if (client.url === targetUrl && 'focus' in client) return client.focus();
            }
            if (clients.openWindow) return clients.openWindow(targetUrl);
        })
    );
});
