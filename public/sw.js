// ============================================================
// public/sw.js — Service Worker
// Controle do Semestre
//
// Estratégias (corrigidas para auto-atualização):
//   HTML / navegação        → Network First  (sempre busca a versão nova)
//   JS/CSS com hash do Vite  → Cache First    (nome muda quando conteúdo muda)
//   Ícones / manifest / SVG  → Cache First
//   Fontes Google            → Cache First
//   Supabase API             → Network First, sem cache (dados frescos)
//
// IMPORTANTE: ao publicar uma nova versão, troque o número em CACHE_VERSION.
// Isso apaga os caches antigos e força o app a carregar o código novo.
// ============================================================

const CACHE_VERSION = "v7";  // ← incremente a cada release (v2, v3, ...)

const SHELL_CACHE = `controle-semestre-shell-${CACHE_VERSION}`;
const FONT_CACHE  = `controle-semestre-fonts-${CACHE_VERSION}`;
const DATA_CACHE  = `controle-semestre-data-${CACHE_VERSION}`;

// Apenas recursos estáveis no pré-cache (NÃO inclui o HTML, que é sempre buscado da rede)
const SHELL_ASSETS = [
  "/manifest.json",
  "/favicon.svg",
];

// ─── INSTALL ─────────────────────────────────────────────────
self.addEventListener("install", (event) => {
  event.waitUntil(
    caches.open(SHELL_CACHE).then((cache) =>
      cache.addAll(SHELL_ASSETS).catch((err) => {
        console.warn("[SW] Shell cache parcial:", err);
      })
    ).then(() => self.skipWaiting())   // ativa a versão nova imediatamente
  );
});

// ─── ACTIVATE ────────────────────────────────────────────────
// Remove TODOS os caches que não sejam da versão atual
self.addEventListener("activate", (event) => {
  const validCaches = [SHELL_CACHE, FONT_CACHE, DATA_CACHE];
  event.waitUntil(
    caches.keys().then((keys) =>
      Promise.all(
        keys
          .filter((key) => !validCaches.includes(key))
          .map((key) => {
            console.log("[SW] Removendo cache antigo:", key);
            return caches.delete(key);
          })
      )
    ).then(() => self.clients.claim())   // assume controle das abas abertas
  );
});

// ─── FETCH ───────────────────────────────────────────────────
self.addEventListener("fetch", (event) => {
  const { request } = event;
  const url = new URL(request.url);

  if (request.method !== "GET") return;
  if (url.protocol !== "http:" && url.protocol !== "https:") return;

  // ── 1. Supabase API → Network First, sem cache ──────────────
  if (url.hostname.includes("supabase.co")) {
    event.respondWith(networkFirst(request, DATA_CACHE, false));
    return;
  }

  // ── 2. Google Fonts → Cache First ───────────────────────────
  if (url.hostname === "fonts.googleapis.com" || url.hostname === "fonts.gstatic.com") {
    event.respondWith(cacheFirst(request, FONT_CACHE));
    return;
  }

  // ── 3. Navegação / HTML → Network First ─────────────────────
  // Garante que o HTML novo (com referência aos JS novos) seja sempre buscado.
  if (request.mode === "navigate" || url.pathname === "/" || url.pathname.endsWith(".html")) {
    event.respondWith(networkFirst(request, SHELL_CACHE, true));
    return;
  }

  // ── 4. Assets com hash (JS/CSS/SVG/ícones) → Cache First ────
  // O Vite muda o nome do arquivo quando o conteúdo muda, então cache é seguro.
  if (
    url.origin === self.location.origin &&
    (
      url.pathname.endsWith(".js") ||
      url.pathname.endsWith(".css") ||
      url.pathname.endsWith(".svg") ||
      url.pathname.startsWith("/icons/") ||
      url.pathname.startsWith("/assets/") ||
      url.pathname === "/manifest.json"
    )
  ) {
    event.respondWith(cacheFirst(request, SHELL_CACHE));
    return;
  }

  // ── 5. Tudo mais → Network First com fallback ───────────────
  event.respondWith(networkFirst(request, DATA_CACHE, true));
});

// ─── MENSAGENS ───────────────────────────────────────────────
// Permite que a página peça ativação imediata da nova versão
self.addEventListener("message", (event) => {
  if (event.data === "SKIP_WAITING") self.skipWaiting();
});

// ─── ESTRATÉGIAS ─────────────────────────────────────────────

async function cacheFirst(request, cacheName) {
  const cached = await caches.match(request);
  if (cached) return cached;
  try {
    const response = await fetch(request);
    if (response.ok) {
      const cache = await caches.open(cacheName);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    return new Response("Offline — recurso não disponível", {
      status: 503,
      headers: { "Content-Type": "text/plain" },
    });
  }
}

async function networkFirst(request, cacheName, useCache) {
  try {
    const response = await fetch(request);
    if (useCache && response.ok) {
      const cache = await caches.open(cacheName);
      cache.put(request, response.clone());
    }
    return response;
  } catch {
    if (useCache) {
      const cached = await caches.match(request);
      if (cached) return cached;
    }
    const fallback = await caches.match("/index.html");
    if (fallback) return fallback;
    return new Response("Você está offline.", {
      status: 503,
      headers: { "Content-Type": "text/plain" },
    });
  }
}
