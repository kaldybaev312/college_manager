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

  let busy = false, reconBusy = false;
  async function tick() {
    if (busy) return;
    busy = true;
    try {
      const { job, reconcile: wantRec } = await api("/api/agent/jobs/claim");
      if (wantRec) runReconcile();
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

  return { start, tick, heartbeat, reconcile, register, remove, listUsers, ping };
}

module.exports = { create, start: (cfg) => create(cfg).start(), deviceTime };
