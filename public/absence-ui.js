// Причины отсутствия и отчёты в веб-панели (index.html, summary.html, journal.html).
// AbsenceUI.cell(row, day)        — ячейка журнала для дня без посещения
// AbsenceUI.bind(tbody, data, cb) — клик по ячейке открывает окно причины (если data.canEdit)
// ReportsUI.mount(el, opts)       — кнопки недельного/месячного отчёта
(function () {
  const TYPES = {
    sick: { label: "Болезнь", excused: true },
    family: { label: "Семейные обстоятельства", excused: true },
    official: { label: "По направлению лицея", excused: true },
    other: { label: "Другая причина", excused: false },
  };
  const esc = (v) =>
    String(v ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c]);
  const fmt = (dk) => (dk ? `${dk.slice(8, 10)}.${dk.slice(5, 7)}.${dk.slice(0, 4)}` : "");

  const css = `
.ab-cell{cursor:default;font-weight:800}
.ab-edit .ab-cell{cursor:pointer}
.ab-n{color:#dc2626}
.ab-r{background:#e0e7ff;color:#4338ca}
.ab-pending{background:#fef3c7;color:#b45309}
.ab-edit .ab-cell:hover{outline:2px solid #6366f1;outline-offset:-2px}
.ab-modal{position:fixed;inset:0;background:rgba(15,23,42,.5);z-index:10000;display:none;align-items:center;justify-content:center;padding:16px}
.ab-modal.open{display:flex}
.ab-box{background:#fff;color:#0f172a;border-radius:16px;width:100%;max-width:420px;padding:20px;box-shadow:0 20px 40px rgba(0,0,0,.2);font-family:inherit}
.ab-box h4{margin:0;font-size:18px;font-weight:800}
.ab-sub{color:#64748b;font-size:13px;margin:2px 0 10px}
.ab-box label{display:block;font-size:12px;font-weight:700;color:#475569;margin:10px 0 4px}
.ab-box select,.ab-box input[type=date],.ab-box textarea{width:100%;box-sizing:border-box;border:2px solid #e2e8f0;border-radius:10px;padding:8px 10px;font-size:14px;font-family:inherit}
.ab-box textarea{min-height:60px;resize:vertical}
.ab-row2{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.ab-check{display:flex;align-items:center;gap:8px;margin-top:10px;font-size:14px;font-weight:600}
.ab-btns{display:flex;gap:8px;margin-top:16px;flex-wrap:wrap}
.ab-btns button{flex:1;border:none;border-radius:10px;padding:10px;font-weight:800;font-size:14px;cursor:pointer}
.ab-save{background:#4f46e5;color:#fff}.ab-ok{background:#16a34a;color:#fff}
.ab-del{background:#fee2e2;color:#dc2626}.ab-cancel{background:#f1f5f9;color:#334155}
.ab-note{font-size:12px;color:#b45309;margin-top:8px}
.rp-wrap{display:flex;flex-wrap:wrap;gap:8px;align-items:center}
.rp-wrap select{width:auto;min-width:170px}
.rp-wrap button{white-space:nowrap}
@media (prefers-color-scheme: dark){:root:not([data-theme=light]) .ab-box{background:#1e293b;color:#f1f5f9}}`;
  const st = document.createElement("style");
  st.textContent = css;
  document.head.appendChild(st);

  let modal, ctx, onSaved;
  function ensureModal() {
    if (modal) return;
    modal = document.createElement("div");
    modal.className = "ab-modal";
    modal.innerHTML = `<div class="ab-box" role="dialog" aria-modal="true" aria-labelledby="ab-title">
      <h4 id="ab-title">Причина отсутствия</h4>
      <div class="ab-sub" id="ab-fio"></div>
      <label for="ab-type">Причина</label>
      <select id="ab-type">${Object.entries(TYPES).map(([k, v]) => `<option value="${k}">${v.label}</option>`).join("")}</select>
      <div class="ab-row2">
        <div><label for="ab-from">С</label><input type="date" id="ab-from"></div>
        <div><label for="ab-to">По</label><input type="date" id="ab-to"></div>
      </div>
      <label class="ab-check"><input type="checkbox" id="ab-excused"> Уважительная причина</label>
      <label for="ab-comment">Комментарий</label>
      <textarea id="ab-comment" maxlength="300" placeholder="Например: справка № 15 от поликлиники"></textarea>
      <div class="ab-note" id="ab-note"></div>
      <div class="ab-btns">
        <button class="ab-save" id="ab-save">Сохранить</button>
        <button class="ab-ok" id="ab-confirm">✓ Подтвердить</button>
        <button class="ab-del" id="ab-delete">Удалить</button>
        <button class="ab-cancel" id="ab-close">Отмена</button>
      </div></div>`;
    document.body.appendChild(modal);
    const $ = (id) => modal.querySelector("#" + id);
    modal.addEventListener("click", (e) => e.target === modal && close());
    $("ab-close").onclick = close;
    $("ab-type").onchange = () => ($("ab-excused").checked = TYPES[$("ab-type").value]?.excused);
    $("ab-save").onclick = () => save(false);
    $("ab-confirm").onclick = () => save(true);
    $("ab-delete").onclick = remove;
    document.addEventListener("keydown", (e) => e.key === "Escape" && close());
  }
  const $m = (id) => modal.querySelector("#" + id);
  function close() {
    modal?.classList.remove("open");
  }
  function open(row, day) {
    ensureModal();
    const r = day.reason;
    ctx = { studentId: row.id, reasonId: r?.id || null };
    $m("ab-title").textContent = r ? "Причина отсутствия" : "Указать причину";
    $m("ab-fio").textContent = `${row.fio} · ${fmt(day.date)}`;
    $m("ab-type").value = r?.type || "sick";
    $m("ab-from").value = r?.dateFrom || day.date;
    $m("ab-to").value = r?.dateTo || day.date;
    $m("ab-excused").checked = r ? !!r.excused : TYPES[$m("ab-type").value].excused;
    $m("ab-comment").value = r?.comment || "";
    $m("ab-note").textContent = r && !r.confirmed ? "Причину сообщил родитель — проверьте и подтвердите." : "";
    $m("ab-confirm").style.display = r && !r.confirmed ? "" : "none";
    $m("ab-delete").style.display = r ? "" : "none";
    modal.classList.add("open");
    $m("ab-type").focus();
  }
  async function call(url, method, body) {
    const res = await fetch(url, { method, headers: { "Content-Type": "application/json" }, body: body ? JSON.stringify(body) : undefined });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || "Ошибка сохранения");
    return data;
  }
  const notify = (msg, type) => (typeof window.toast === "function" ? window.toast(msg, type) : alert(msg));
  async function save(confirm) {
    const body = {
      type: $m("ab-type").value,
      dateFrom: $m("ab-from").value,
      dateTo: $m("ab-to").value,
      excused: $m("ab-excused").checked,
      comment: $m("ab-comment").value.trim(),
    };
    if (ctx.reasonId) body.confirmed = true;
    try {
      if (ctx.reasonId) await call(`/api/admin/absences/${ctx.reasonId}`, "PUT", body);
      else await call("/api/admin/absences", "POST", { ...body, studentId: ctx.studentId });
      close();
      notify(confirm ? "Причина подтверждена" : "Причина сохранена", "success");
      onSaved && onSaved();
    } catch (e) {
      notify(e.message, "error");
    }
  }
  async function remove() {
    if (!ctx.reasonId || !confirm("Удалить причину отсутствия?")) return;
    try {
      await call(`/api/admin/absences/${ctx.reasonId}`, "DELETE");
      close();
      notify("Причина удалена", "success");
      onSaved && onSaved();
    } catch (e) {
      notify(e.message, "error");
    }
  }

  function cell(row, day, style = "") {
    const r = day.reason;
    const data = `data-sid="${esc(row.id)}" data-date="${esc(day.date)}"`;
    if (r) {
      const title = `${r.label}${r.comment ? ": " + r.comment : ""}${r.confirmed ? "" : " (ждёт подтверждения)"}`;
      return `<td class="ab-cell ${r.confirmed ? "ab-r" : "ab-pending"}" ${data} title="${esc(title)}" style="${style}">${esc(r.code)}${r.confirmed ? "" : "?"}</td>`;
    }
    return `<td class="ab-cell ab-n" ${data} title="Без причины — нажмите, чтобы указать" style="${style}">Н</td>`;
  }

  function bind(tbody, data, cb) {
    onSaved = cb;
    const table = tbody.closest("table") || tbody;
    table.classList.toggle("ab-edit", !!data.canEdit);
    tbody.onclick = (e) => {
      if (!data.canEdit) return;
      const td = e.target.closest("td.ab-cell");
      if (!td) return;
      const row = data.matrix.find((r) => r.id === td.dataset.sid);
      const day = row?.days.find((d) => d.date === td.dataset.date);
      if (row && day) open(row, day);
    };
  }

  window.AbsenceUI = { cell, bind, open, TYPES };

  // ── Отчёты ──
  const pad = (n) => String(n).padStart(2, "0");
  const keyOf = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
  function mount(el, { getGid, getMonth, admin }) {
    if (!el) return;
    el.innerHTML = `<div class="rp-wrap">
      <select class="form-select" id="rp-period" aria-label="Период отчёта">
        <option value="week:0">Эта неделя</option>
        <option value="week:1">Прошлая неделя</option>
        <option value="month:sel">Месяц из фильтра</option>
        <option value="month:prev">Прошлый месяц</option>
      </select>
      <button class="btn btn-outline-primary fw-bold" id="rp-group">📊 Отчёт по группе</button>
      ${admin ? `<button class="btn btn-outline-success fw-bold" id="rp-all">🏫 Все группы</button>
      <button class="btn btn-outline-secondary fw-bold" id="rp-send" title="Отправить кураторам и админам в Telegram за прошлую неделю">✈️ Разослать в Telegram</button>` : ""}
    </div>`;
    const q = () => {
      const [period, which] = el.querySelector("#rp-period").value.split(":");
      const now = new Date();
      let date;
      if (period === "week") {
        now.setDate(now.getDate() - 7 * Number(which));
        date = keyOf(now);
      } else if (which === "sel") date = (getMonth() || keyOf(now).slice(0, 7)) + "-01";
      else date = keyOf(new Date(now.getFullYear(), now.getMonth() - 1, 1));
      return `period=${period}&date=${date}`;
    };
    el.querySelector("#rp-group").onclick = () => {
      const gid = getGid();
      if (!gid) return notify("Выберите группу", "error");
      window.open(`/api/admin/report/group/${gid}?${q()}`, "_blank");
    };
    if (admin) {
      el.querySelector("#rp-all").onclick = () => window.open(`/api/admin/report/all?${q()}`, "_blank");
      el.querySelector("#rp-send").onclick = async () => {
        const period = el.querySelector("#rp-period").value.startsWith("month") ? "month" : "week";
        if (!confirm(`Разослать ${period === "month" ? "месячный (за прошлый месяц)" : "недельный (за прошлую неделю)"} отчёт кураторам и админам в Telegram?`)) return;
        try {
          const r = await call("/api/admin/report/send-now", "POST", { period });
          notify(`Отправлено сообщений: ${r.sent}`, "success");
        } catch (e) {
          notify(e.message, "error");
        }
      };
    }
  }
  window.ReportsUI = { mount };
})();
