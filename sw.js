const CACHE_VERSION = "catalogo-eldorado-v4-optimized";
const STATIC_CACHE = `${CACHE_VERSION}-static`;
const IMAGE_CACHE = `${CACHE_VERSION}-images`;
const DATA_CACHE = `${CACHE_VERSION}-data`;
const EXTERNAL_CACHE = `${CACHE_VERSION}-external`;
const APP_VERSION = "20261103-opt1";

const CORE_ASSETS = [
  "./",
  "./index.html",
  `./style.css?v=${APP_VERSION}`,
  `./script.js?v=${APP_VERSION}`,
  "./manifest.json",
  "./Logos/Favicon.png",
  "./Logos/Eldorado.png",
  "./Logos/Produtos-ternura.png",
  "./Logos/SoFruta.png"
];

const MAX_IMAGE_CACHE_ITEMS = 700;

self.addEventListener("install", event => {
  event.waitUntil(
    caches.open(STATIC_CACHE)
      .then(cache => cache.addAll(CORE_ASSETS))
      .catch(() => null)
  );
  self.skipWaiting();
});

self.addEventListener("activate", event => {
  const permitidos = new Set([STATIC_CACHE, IMAGE_CACHE, DATA_CACHE, EXTERNAL_CACHE]);

  event.waitUntil(
    caches.keys()
      .then(keys => Promise.all(
        keys.map(key => permitidos.has(key) ? null : caches.delete(key))
      ))
      .then(() => self.clients.claim())
  );
});

self.addEventListener("fetch", event => {
  const request = event.request;
  if (request.method !== "GET") return;

  const url = new URL(request.url);

  // HTML: rede primeiro para receber deploys novos.
  if (request.mode === "navigate") {
    event.respondWith(networkFirst(request, STATIC_CACHE, "./index.html"));
    return;
  }

  // Planilha: atualiza online e usa cache se estiver offline.
  if (url.origin === self.location.origin && url.pathname.endsWith("/Produtos.xlsx")) {
    event.respondWith(networkFirst(request, DATA_CACHE));
    return;
  }

  // Imagens e logos: cache-first para reduzir solicitações de CDN.
  if (
    url.origin === self.location.origin &&
    (
      request.destination === "image" ||
      url.pathname.includes("/Imagens/") ||
      url.pathname.includes("/Logos/")
    )
  ) {
    event.respondWith(cacheFirst(request, IMAGE_CACHE, true));
    return;
  }

  // CSS, JS e manifest locais: cache-first.
  if (
    url.origin === self.location.origin &&
    (
      request.destination === "style" ||
      request.destination === "script" ||
      url.pathname.endsWith("/manifest.json") ||
      url.pathname.endsWith(".css") ||
      url.pathname.endsWith(".js")
    )
  ) {
    event.respondWith(cacheFirst(request, STATIC_CACHE));
    return;
  }

  // Bibliotecas externas ficam em cache após a primeira carga.
  if (url.origin !== self.location.origin) {
    event.respondWith(cacheFirst(request, EXTERNAL_CACHE));
    return;
  }

  event.respondWith(cacheFirst(request, STATIC_CACHE));
});

async function cacheFirst(request, cacheName, limitarImagens = false) {
  const cache = await caches.open(cacheName);
  const cached = await cache.match(request);

  if (cached) return cached;

  try {
    const response = await fetch(request);

    if (response && (response.ok || response.type === "opaque")) {
      cache.put(request, response.clone()).catch(() => null);

      if (limitarImagens) {
        limitarCache(cacheName, MAX_IMAGE_CACHE_ITEMS).catch(() => null);
      }
    }

    return response;
  } catch {
    return cached || Response.error();
  }
}

async function networkFirst(request, cacheName, fallbackUrl = null) {
  const cache = await caches.open(cacheName);

  try {
    const response = await fetch(request);

    if (response && (response.ok || response.type === "opaque")) {
      cache.put(request, response.clone()).catch(() => null);
    }

    return response;
  } catch {
    const cached = await cache.match(request);

    if (cached) return cached;

    if (fallbackUrl) {
      const fallback = await cache.match(fallbackUrl) || await caches.match(fallbackUrl);
      if (fallback) return fallback;
    }

    return Response.error();
  }
}

async function limitarCache(cacheName, limite) {
  const cache = await caches.open(cacheName);
  const chaves = await cache.keys();

  if (chaves.length <= limite) return;

  const excedente = chaves.length - limite;
  await Promise.all(chaves.slice(0, excedente).map(chave => cache.delete(chave)));
}
