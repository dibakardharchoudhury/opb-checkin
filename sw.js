// Minimal offline shell for the OPB Check-in PWA. Scanning + registration require
// the network (camera + backend), so we only cache the static shell here.
const VERSION = "v8";
const SHELL = ["./", "./index.html", "./manifest.webmanifest", "./mandala.svg", "./version.json", "./icon.svg", "./icon-192.png", "./icon-512.png"];

self.addEventListener("install", (e) => {
  e.waitUntil(caches.open(VERSION).then((c) => c.addAll(SHELL)).then(() => self.skipWaiting()));
});
self.addEventListener("activate", (e) => {
  e.waitUntil(caches.keys().then((keys) => Promise.all(keys.filter((k) => k !== VERSION).map((k) => caches.delete(k)))).then(() => self.clients.claim()));
});
self.addEventListener("fetch", (e) => {
  const url = new URL(e.request.url);
  if (e.request.method !== "GET" || url.pathname.includes("/api/")) return; // never cache API calls
  e.respondWith(caches.match(e.request).then((hit) => hit || fetch(e.request)));
});
