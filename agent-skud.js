// СКУД: часть локального агента (запускается на компьютере в лицее рядом с agent.js).
// Забирает задания с сайта (записать человека на терминалы / удалить), раз в минуту сообщает
// состояние терминалов и по запросу (и раз в несколько часов) сверяет список людей на терминалах.
//
// Подключение в agent.js:
//   const skud = require("./agent-skud");
//   skud.start({ serverUrl: SERVER_URL, apiKey: AGENT_API_KEY, devices: DEVICES, user: HIK_USER, pass: HIK_PASS,
//                onChanged: (employeeNo) => { for (const d of DEVICES) empCache.delete(`${d.ip}:${employeeNo}`); } });
const crypto = require("crypto");

const PLAN = { doorNo: 1, planTemplateNo: "1" };
const TZ_SHIFT_MS = 6 * 3600 * 1000; // терминалы живут по времени UTC+6

const withTimeout = (p, ms = 20000) =>
  Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error("таймаут устройства")), ms))]);

// Время для терминала без пояса: YYYY-MM-DDTHH:mm:ss по UTC+6
const deviceTime = (iso) => new Date(new Date(iso).getTime() + TZ_SHIFT_MS).toISOString().slice(0, 19);

// multipart вручную в Buffer: digest-авторизация шлёт запрос дважды, поток второй раз не отправить
function multipart(json, jpeg) {
  const b = "----skud" + crypto.randomBytes(8).toString("hex");
  const body = Buffer.concat([
    Buffer.from(`--${b}\r\nContent-Disposition: form-data; name="FaceDataRecord"\r\nContent-Type: application/json\r\n\r\n${JSON.stringify(json)}\r\n`),
    Buffer.from(`--${b}\r\nContent-Disposition: form-data; name="FaceImage"; filename="face.jpg"\r\nContent-Type: image/jpeg\r\n\r\n`),
    jpeg,
    Buffer.from(`\r\n--${b}--\r\n`),
  ]);
  return { body, type: `multipart/form-data; boundary=${b}` };
}

function isOk(r, txt) {
  if (!r.ok) return false;
  try { const j = JSON.parse(txt); return j.statusCode === undefined || j.statusCode === 1; } catch { return true; }
}

function create(cfg) {
  const DigestClient = cfg.DigestClient || require("digest-fetch").default;
  const client = () => new DigestClient(cfg.user, cfg.pass, { algorithm: "MD5" });
  const base = (ip) => `http://${ip}/ISAPI`;
  const J = { "Content-Type": "application/json" };
  const call = async (c, url, method, body, headers = J) => {
    const r = await withTimeout(c.fetch(url, { method, headers, body }));
    const t = await r.text();
    return { r, t };
  };

  /* ---------------- записать человека ---------------- */
  async function register(ip, job, jpeg) {
    const c = client(), b = base(ip);
    const until = job.validUntil ? deviceTime(job.validUntil) : "2035-12-31T23:59:59";
    const userBody = JSON.stringify({ UserInfo: {
      employeeNo: job.employeeNo, name: job.name, userType: "normal",
      Valid: { enable: true, beginTime: "2026-01-01T00:00:00", endTime: until },
      doorRight: "1", RightPlan: [PLAN],
    } });

    let { r, t } = await call(c, `${b}/AccessControl/UserInfo/Record?format=json`, "POST", userBody);
    if (!isOk(r, t) && /alreadyexist/i.test(t)) ({ r, t } = await call(c, `${b}/AccessControl/UserInfo/Modify?format=json`, "PUT", userBody));
    if (!isOk(r, t)) throw new Error(`пользователь: ${r.status} ${t.slice(0, 160)}`);

    const mp = multipart({ faceLibType: "blackFD", FDID: "1", FPID: job.employeeNo }, jpeg);
    const h = { "Content-Type": mp.type };
    ({ r, t } = await call(c, `${b}/Intelligent/FDLib/FaceDataRecord?format=json`, "POST", mp.body, h));
    if (!isOk(r, t) && /alreadyexist/i.test(t)) ({ r, t } = await call(c, `${b}/Intelligent/FDLib/FDModify?format=json`, "PUT", mp.body, h));
    if (!isOk(r, t))
      throw new Error(/SubpicAnalysisModelingError/.test(t)
        ? "лицо не распознано на фото: нужно чёткое фото анфас, лицо крупно, без очков и головного убора"
        : `лицо: ${r.status} ${t.slice(0, 160)}`);
  }

  /* ---------------- удалить человека ---------------- */
  async function remove(ip, job) {
    const c = client(), b = base(ip);
    const body = JSON.stringify({ UserInfoDelCond: { EmployeeNoList: [{ employeeNo: String(job.employeeNo) }] } });
    let { r, t } = await call(c, `${b}/AccessControl/UserInfo/Delete?format=json`, "PUT", body);
    if (!isOk(r, t) && /notexist|not exist|nodata/i.test(t)) return; // уже нет на терминале
    if (!isOk(r, t)) {
      // старые прошивки: другой формат запроса
      const alt = JSON.stringify({ UserInfoDetail: { mode: "byEmployeeNo", EmployeeNoList: [{ employeeNo: String(job.employeeNo) }] } });
      ({ r, t } = await call(c, `${b}/AccessControl/UserInfoDetail/Delete?format=json`, "PUT", alt));
      if (!isOk(r, t) && /notexist|not exist|nodata/i.test(t)) return;
      if (!isOk(r, t)) throw new Error(`удаление: ${r.status} ${t.slice(0, 160)}`);
    }
  }

  /* ---------------- опрос терминала ---------------- */
  async function listUsers(ip) {
    const c = client(), out = [];
    for (let pos = 0; pos < 20000; ) {
      const body = JSON.stringify({ UserInfoSearchCond: { searchID: "1", searchResultPosition: pos, maxResults: 30 } });
      const { r, t } = await call(c, `${base(ip)}/AccessControl/UserInfo/Search?format=json`, "POST", body);
      if (!isOk(r, t)) throw new Error(`список: ${r.status} ${t.slice(0, 120)}`);
      const s = JSON.parse(t).UserInfoSearch || {};
      const list = s.UserInfo || [];
      for (const u of list) out.push({ employeeNo: String(u.employeeNo), name: String(u.name || "") });
      pos += list.length;
      if (!list.length || s.responseStatusStrg !== "MORE") break;
    }
    return out;
  }

  async function ping(ip) {
    const c = client(), t0 = Date.now();
    const { r } = await call(c, `${base(ip)}/System/deviceInfo`, "GET");
    if (!r.ok) throw new Error(`HTTP ${r.status}`);
    const ms = Date.now() - t0;
    let users;
    try {
      const { r: r2, t } = await call(c, `${base(ip)}/AccessControl/UserInfo/Count?format=json`, "GET");
      if (r2.ok) users = JSON.parse(t).UserInfoCount?.userNumber;
    } catch { /* число людей не критично */ }
    return { ms, users };
  }

  /* ---------------- общение с сайтом ---------------- */
  const root = String(cfg.serverUrl || "").replace(/\/api\/agent\/sync\/?$/, "").replace(/\/$/, "");
  async function api(path, body) {
    const r = await fetch(root + path, { method: "POST", headers: J, body: JSON.stringify({ apiKey: cfg.apiKey, ...body }) });
    if (!r.ok) throw new Error(`${path}: HTTP ${r.status}`);
    return r.json();
  }

  async function heartbeat() {
    const devices = await Promise.all(cfg.devices.map(async (d) => {
      try { return { ip: d.ip, online: true, ...(await ping(d.ip)) }; }
      catch (e) { return { ip: d.ip, online: false, error: e.message }; }
    }));
    await api("/api/agent/heartbeat", { devices });
    return devices;
  }

  async function reconcile() {
    const devices = await Promise.all(cfg.devices.map(async (d) => {
      try { return { ip: d.ip, ok: true, people: await withTimeout(listUsers(d.ip), 120000) }; }
      catch (e) { return { ip: d.ip, ok: false, error: e.message }; }
    }));
    return api("/api/agent/reconcile", { devices });
  }


  /* ---------------- восстановление истории посещаемости ---------------- */
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const evMs = (t) => (!t ? NaN : /[zZ]|[+-]\d\d:?\d\d$/.test(t) ? new Date(t).getTime() : new Date(t + "+06:00").getTime());
  const dayKey = (ms) => new Date(ms + TZ_SHIFT_MS).toISOString().slice(0, 10);
  const nameCache = new Map();

  async function nameByNo(c, ip, no) {
    const k = ip + ":" + no;
    if (nameCache.has(k)) return nameCache.get(k);
    let name = "";
    try {
      const { r, t } = await call(c, `${base(ip)}/AccessControl/UserInfo/Search?format=json`, "POST", JSON.stringify({
        UserInfoSearchCond: { searchID: "1", searchResultPosition: 0, maxResults: 1, EmployeeNoList: [{ employeeNo: String(no) }] },
      }));
      if (r.ok) name = JSON.parse(t)?.UserInfoSearch?.UserInfo?.[0]?.name || "";
    } catch { /* имя не нашли: пропустим событие */ }
    nameCache.set(k, name);
    return name;
  }

  // Все события терминала за один день (с постраничным чтением)
  async function eventsOfDay(c, ip, day) {
    const all = [];
    let pos = 0;
    for (let page = 0; page < 300; page++) {
      const { r, t } = await call(c, `${base(ip)}/AccessControl/AcsEvent?format=json`, "POST", JSON.stringify({
        AcsEventCond: { searchID: "h" + Date.now(), searchResultPosition: pos, maxResults: 100, major: 5, minor: 75, startTime: `${day}T00:00:00`, endTime: `${day}T23:59:59` },
      }));
      if (!r.ok) throw new Error(`HTTP ${r.status}`);
      const j = JSON.parse(t)?.AcsEvent || {};
      const list = j.InfoList || [];
      all.push(...list);
      pos += list.length;
      if (!list.length || j.responseStatusStrg === "OK" || (j.totalMatches && pos >= j.totalMatches)) break;
      await sleep(150);
    }
    return all;
  }

  async function sendDays(days) {
    for (let i = 0; i < days.length; i += 200) {
      const chunk = days.slice(i, i + 200);
      let lastErr;
      for (let a = 0; a < 3; a++) {
        try { await api("/api/agent/history-sync", { days: chunk }); lastErr = null; break; }
        catch (e) { lastErr = e; await sleep(2000 * (a + 1)); }
      }
      if (lastErr) throw lastErr;
    }
  }

  const iterDays = (from, to) => {
    const out = [];
    for (let t = new Date(from + "T00:00:00Z").getTime(); t <= new Date(to + "T00:00:00Z").getTime(); t += 86400000) out.push(new Date(t).toISOString().slice(0, 10));
    return out;
  };

  let histBusy = false;
  async function runHistory({ from, to }) {
    if (histBusy) return;
    histBusy = true;
    const prog = (status, progress, sent) => api("/api/agent/history/progress", { status, progress, sent }).catch(() => {});
    const errs = new Map(); // ip -> число дней с ошибкой
    let sent = 0, events = 0;
    try {
      console.log(`⏪ История: ${from} — ${to}`);
      const c = client();
      for (const day of iterDays(from, to)) {
        const agg = new Map(); // empNo|date -> данные дня
        for (const dev of cfg.devices) {
          let list;
          try { list = await eventsOfDay(c, dev.ip, day); }
          catch (e) { errs.set(dev.ip, (errs.get(dev.ip) || 0) + 1); console.log(`   ${day} ${dev.ip}: ❌ ${e.message}`); continue; }
          for (const ev of list) {
            const no = String(ev.employeeNoString || ev.employeeNo || "").trim();
            const ms = evMs(ev.time);
            if (!no || isNaN(ms)) continue;
            const date = dayKey(ms), key = no + "|" + date;
            let a = agg.get(key);
            if (!a) agg.set(key, (a = { id: no, date, rawName: String(ev.name || ev.personName || "").trim(), ip: dev.ip, in: null, out: null }));
            if (!a.rawName) a.rawName = await nameByNo(c, dev.ip, no);
            const side = dev.direction === "out" ? "out" : "in";
            const s2 = (a[side] ||= { first: ev.time, firstMs: ms, last: ev.time, lastMs: ms, n: 0 });
            if (ms < s2.firstMs) { s2.first = ev.time; s2.firstMs = ms; }
            if (ms > s2.lastMs) { s2.last = ev.time; s2.lastMs = ms; }
            s2.n++;
            events++;
          }
          await sleep(200);
        }
        const days = [...agg.values()].filter((a) => a.rawName).map((a) => ({
          id: a.id, rawName: a.rawName, date: a.date, deviceIp: a.ip,
          firstIn: a.in?.first, lastIn: a.in?.last, inCount: a.in?.n || 0,
          firstOut: a.out?.first, lastOut: a.out?.last, outCount: a.out?.n || 0,
        }));
        await sendDays(days);
        sent += days.length;
        console.log(`   ${day}: людей ${days.length}`);
        await prog("running", `обработано до ${day}, событий прочитано: ${events}`, sent);
      }
      const bad = [...errs].map(([ip, n]) => `${ip} (${n} дн.)`).join(", ");
      const allBad = errs.size >= cfg.devices.length && [...errs.values()].every((n) => n >= iterDays(from, to).length);
      await prog(allBad ? "failed" : "done", `${allBad ? "терминалы не ответили" : "готово"}: ${from} — ${to}, событий ${events}, записей ${sent}${bad ? `. Ошибки терминалов: ${bad}` : ""}`, sent);
      console.log(`⏪ История готова: событий ${events}, записей ${sent}`);
    } catch (e) {
      console.error("история:", e.message);
      await prog("failed", `ошибка: ${e.message}`.slice(0, 250), sent);
    } finally { histBusy = false; }
  }

  let busy = false, reconBusy = false;
  async function tick() {
    if (busy) return;
    busy = true;
    try {
      const { job, reconcile: wantRec, history } = await api("/api/agent/jobs/claim");
      if (wantRec) runReconcile();
      if (history) runHistory(history);
      if (!job) return;
      const jpeg = job.photo ? Buffer.from(job.photo, "base64") : null;
      console.log(`${job.type === "delete" ? "🗑 Удаление" : "📷 Регистрация"}: ${job.name || ""} (№${job.employeeNo})`);
      const settled = await Promise.allSettled(cfg.devices.map((d) => (job.type === "delete" ? remove(d.ip, job) : register(d.ip, job, jpeg))));
      const results = settled.map((s, i) => ({ ip: cfg.devices[i].ip, ok: s.status === "fulfilled", error: s.reason?.message }));
      results.forEach((x) => console.log(`   ${x.ip}: ${x.ok ? "✅" : "❌ " + x.error}`));
      await api(`/api/agent/jobs/${job.id}/result`, { results });
      cfg.onChanged && cfg.onChanged(job.employeeNo);
    } catch (e) {
      if (!/fetch failed|HTTP 5\d\d/.test(e.message)) console.error("СКУД:", e.message);
    } finally { busy = false; }
  }

  async function runReconcile() {
    if (reconBusy) return;
    reconBusy = true;
    try {
      const r = await reconcile();
      console.log(`🔎 Сверка: лишних ${r.extra}, не хватает ${r.missing}`);
    } catch (e) { console.error("сверка:", e.message); }
    finally { reconBusy = false; }
  }

  function start() {
    const timers = [
      setInterval(tick, cfg.pollMs || 5000),
      setInterval(() => heartbeat().catch(() => {}), cfg.heartbeatMs || 60000),
      setInterval(runReconcile, (cfg.reconcileHours || 6) * 3600 * 1000),
    ];
    tick();
    heartbeat().catch(() => {});
    setTimeout(runReconcile, 90 * 1000);
    return () => timers.forEach(clearInterval);
  }

  return { start, tick, runHistory, heartbeat, reconcile, register, remove, listUsers, ping };
}

module.exports = { create, start: (cfg) => create(cfg).start(), deviceTime };
