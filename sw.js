const CACHE_NAME = 'tracker-ui-v28';
const ASSETS_TO_CACHE = [
  './',
  './index.html',
  './USER_MANUAL.md',
  './css/styles.css',
  './js/mockMasterAccounts.js',
  './js/core-ui.js',
  './js/workbook-io.js',
  './js/tracker-table.js',
  './js/request-editors.js',
  './js/scratch-sheets.js',
  './js/rollover.js',
  './manifest.json',
  './assets/icons/icon.svg',
  './assets/icons/icon-192.png',
  './assets/icons/icon-512.png',
  './assets/images/manual/01-main-workspace-overview.png',
  './assets/images/manual/02-tools-and-setup-menu.png',
  './assets/images/manual/03-custom-department-and-layout-editor.png',
  './assets/images/manual/04-column-mapping-modal.png',
  './assets/images/manual/05-ticket-details-and-timeline.png',
  './assets/images/manual/06-smart-rollover-modal.png',
  './assets/images/manual/07-scratch-sheet-and-context-menu.png',
  './assets/images/manual/08-data-to-columns-modal.png',
  './assets/images/manual/09-supervisor-master-rollup.png',
  './assets/images/manual/10-light-mode-and-batch-actions.png',
  'https://cdn.jsdelivr.net/npm/exceljs@4.4.0/dist/exceljs.min.js'
];

// Install: fetch fresh application shell bypassing browser HTTP disk cache
self.addEventListener('install', (event) => {
  self.skipWaiting();
  event.waitUntil(
    caches.open(CACHE_NAME).then(async (cache) => {
      await Promise.all(
        ASSETS_TO_CACHE.map(async (url) => {
          const isCrossOrigin = /^https?:\/\//i.test(url) && !url.startsWith(self.location.origin);
          const req = isCrossOrigin
            ? new Request(url)
            : new Request(url, { cache: 'reload' });
          const res = await fetch(req);
          if (res && res.ok) {
            await cache.put(url, res);
          }
        })
      );
    })
  );
});

// Activate: cleanup all older cache buckets and immediately claim open clients
self.addEventListener('activate', (event) => {
  event.waitUntil(
    caches.keys().then((keys) => {
      return Promise.all(
        keys.map((key) => {
          if (key !== CACHE_NAME) {
            return caches.delete(key);
          }
        })
      );
    }).then(() => self.clients.claim())
  );
});

// Listen for explicit update / cache-clear commands from the client
self.addEventListener('message', (event) => {
  if (!event.data) return;
  if (event.data.type === 'SKIP_WAITING') {
    self.skipWaiting();
  }
});

// Fetch: Network-first with HTTP cache revalidation for same-origin app shell
// so HTML, CSS, and JS never get out of sync; fallback to CacheStorage when offline.
self.addEventListener('fetch', (event) => {
  if (event.request.method !== 'GET') return;

  const reqUrl = new URL(event.request.url);
  const isSameOrigin = reqUrl.origin === self.location.origin;

  event.respondWith(
    (async () => {
      try {
        const fetchOptions = isSameOrigin
          ? new Request(event.request, { cache: 'no-cache' })
          : event.request;
        const networkResponse = await fetch(fetchOptions);

        if (networkResponse && networkResponse.status === 200) {
          const responseToCache = networkResponse.clone();
          const cache = await caches.open(CACHE_NAME);
          await cache.put(event.request, responseToCache);
          // Also keep canonical path (without ?v= query string) updated in cache for offline fallback
          if (isSameOrigin && reqUrl.search) {
            const canonicalUrl = reqUrl.origin + reqUrl.pathname;
            await cache.put(canonicalUrl, networkResponse.clone());
          }
        }
        return networkResponse;
      } catch (err) {
        const exactMatch = await caches.match(event.request);
        if (exactMatch) return exactMatch;

        const ignoreSearchMatch = await caches.match(event.request, { ignoreSearch: true });
        if (ignoreSearchMatch) return ignoreSearchMatch;

        if (event.request.mode === 'navigate') {
          return caches.match('./index.html', { ignoreSearch: true });
        }
        throw err;
      }
    })()
  );
});
