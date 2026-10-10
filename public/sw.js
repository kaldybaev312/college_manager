// Сервис-воркер ПЛ №3: установка на главный экран и заставка, пока сервер просыпается.
// Данные (API) никогда не кэшируются — всегда свежие с сети.
const VERSION = "pl3-v1";
const SHELL = ["/wake.html", "/manifest.webmanifest", "/icons/icon-192.png", "/icons/icon-512.png", "/icons/favicon-32.png", "/icons/apple-touch-icon.png", "/pwa.js"];
const WAKE_AFTER_MS = 4000; // если страница не пришла за это время — показываем заставку

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(
    caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim()),
  );
});

const isApi = (url) => url.pathname.startsWith("/api/") || url.pathname === "/login" || url.pathname === "/logout";
const isStatic = (url) => /\.(png|svg|ico|webmanifest|js|css|woff2?)$/.test(url.pathname) && url.pathname !== "/sw.js";

self.addEventListener("fetch", (e) => {
  const req = e.request;
  if (req.method !== "GET") return;
  const url = new URL(req.url);
  if (url.origin !== location.origin || isApi(url)) return; // сеть как есть

  // Страницы: всегда с сети (там проверка входа). Если сервер спит дольше 4 с или нет сети —
  // отдаём заставку, она сама перезагрузит страницу, когда сервер ответит.
  if (req.mode === "navigate") {
    e.respondWith(
      new Promise((resolve) => {
        let done = false;
        const wake = () => caches.match("/wake.html").then((r) => r || new Response("Загрузка…", { headers: { "Content-Type": "text/plain; charset=utf-8" } }));
        const timer = setTimeout(() => { if (!done) { done = true; resolve(wake()); } }, WAKE_AFTER_MS);
        fetch(req)
          .then((res) => { if (!done) { done = true; clearTimeout(timer); resolve(res); } })
          .catch(() => { if (!done) { done = true; clearTimeout(timer); resolve(wake()); } });
      }),
    );
    return;
  }

  // Статика своего сайта: отдаём из кэша сразу и тихо обновляем
  if (isStatic(url)) {
    e.respondWith(
      caches.open(VERSION).then(async (c) => {
        const hit = await c.match(req);
        const net = fetch(req).then((res) => { if (res.ok) c.put(req, res.clone()); return res; }).catch(() => hit);
        return hit || net;
      }),
    );
  }
});
