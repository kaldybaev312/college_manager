// ПЛ №3 как приложение: сервис-воркер, заставка при загрузке, кнопка «Установить».
// Подключается на всех основных страницах вместе с разметкой заставки #pl3-splash.
(function () {
  // ── сервис-воркер ──
  if ("serviceWorker" in navigator && location.protocol !== "file:") {
    window.addEventListener("load", () => navigator.serviceWorker.register("/sw.js").catch(() => {}));
  }

  // ── заставка ──
  // Страница с атрибутом <body data-splash-wait> прячет заставку сама: window.pl3Ready()
  // (когда загружены данные). Остальные — по событию load. В любом случае не дольше 15 с.
  let hidden = false;
  function hide() {
    if (hidden) return;
    hidden = true;
    const s = document.getElementById("pl3-splash");
    if (!s) return;
    s.classList.add("pl3-out");
    setTimeout(() => s.remove(), 450);
  }
  window.pl3Ready = hide;
  const waitData = () => document.body && document.body.hasAttribute("data-splash-wait");
  window.addEventListener("load", () => { if (!waitData()) setTimeout(hide, 120); });
  setTimeout(hide, 15000);

  // ── установка на главный экран ──
  const KEY = "pl3-install-dismissed";
  const dismissedRecently = () => {
    try { return Date.now() - Number(localStorage.getItem(KEY) || 0) < 30 * 864e5; } catch (e) { return false; }
  };
  const standalone = () => window.matchMedia("(display-mode: standalone)").matches || navigator.standalone === true;
  const inTelegram = () => !!(window.Telegram && window.Telegram.WebApp && window.Telegram.WebApp.initData);

  function pill(html, onClick) {
    if (document.getElementById("pl3-install")) return;
    const b = document.createElement("div");
    b.id = "pl3-install";
    b.setAttribute("role", "region");
    b.setAttribute("aria-label", "Установка приложения");
    b.innerHTML = `<button type="button" class="pl3-go">${html}</button><button type="button" class="pl3-x" aria-label="Скрыть">×</button>`;
    const st = document.createElement("style");
    st.textContent = `#pl3-install{position:fixed;right:14px;bottom:14px;z-index:9998;display:flex;align-items:center;gap:2px;background:#13285c;color:#fff;border-radius:999px;box-shadow:0 10px 30px rgba(0,0,0,.25);font:600 14px system-ui,-apple-system,'Segoe UI',sans-serif;max-width:calc(100% - 28px)}
#pl3-install button{background:none;border:0;color:inherit;font:inherit;cursor:pointer}
#pl3-install .pl3-go{padding:11px 6px 11px 16px;text-align:left}
#pl3-install .pl3-x{padding:8px 14px 8px 8px;font-size:20px;line-height:1;opacity:.7}
#pl3-install .pl3-x:hover{opacity:1}`;
    document.head.appendChild(st);
    document.body.appendChild(b);
    b.querySelector(".pl3-go").onclick = onClick;
    b.querySelector(".pl3-x").onclick = () => { try { localStorage.setItem(KEY, String(Date.now())); } catch (e) {} b.remove(); };
  }

  let deferred = null;
  window.addEventListener("beforeinstallprompt", (e) => {
    e.preventDefault();
    deferred = e;
    if (dismissedRecently() || standalone() || inTelegram()) return;
    pill("📲 Установить приложение", async () => {
      document.getElementById("pl3-install")?.remove();
      deferred.prompt();
      try { await deferred.userChoice; } catch (e) {}
      deferred = null;
    });
  });
  window.addEventListener("appinstalled", () => document.getElementById("pl3-install")?.remove());

  // iPhone/iPad: Safari не показывает системное предложение — подсказываем вручную
  window.addEventListener("load", () => {
    const ios = /iphone|ipad|ipod/i.test(navigator.userAgent) && !/crios|fxios/i.test(navigator.userAgent);
    if (!ios || standalone() || inTelegram() || dismissedRecently()) return;
    setTimeout(() => pill("📲 На экран «Домой»: «Поделиться» → «На экран Домой»", () => {}), 2500);
  });
})();
