// Восстановление истории: сайт → запрос → агент читает терминалы (постранично) → history-sync. БД в памяти.
const Module = require("module");
const path = require("path");
const assert = require("assert/strict");
const fake = require("./fake-mongoose");
const origResolve = Module._resolveFilename;
Module._resolveFilename = function (req, ...rest) {
  return req === "mongoose" ? path.join(__dirname, "fake-mongoose.js") : origResolve.call(this, req, ...rest);
};
const express = require("express");
const startMock = require("./mock-hik");
const agentSkud = require("../agent-skud");
const { Schema } = fake;
const KEY = "k";
const Group = fake.model("Group", new Schema({ name: String, shiftStart: { type: String, default: "08:00" }, kind: { type: String, default: "students" } }));
const Student = fake.model("Student", new Schema({ fio: String, inn: String, lyceumId: String, group: Schema.Types.ObjectId, attendance: [], skudDaily: [] }));
const sessionOf = (req) => { try { return JSON.parse(req.headers["x-test-session"]); } catch { return null; } };
const requireAuth = (req, res, next) => (sessionOf(req)?.role ? next() : res.status(401).json({ error: "auth" }));
const requireAdmin = (req, res, next) => (sessionOf(req)?.role === "admin" ? next() : res.status(403).json({ error: "Admin only" }));
const normalizeINN = (s) => { const d = String(s || "").replace(/\D/g, ""); return d.length >= 12 ? d : null; };
const normalizeFIO = (s) => String(s || "").replace(/[.,]/g, " ").replace(/\s+/g, " ").trim().toLowerCase();
const extractINN = (s) => { const m = String(s || "").match(/\b(\d{12,16})\b/); return m ? m[1] : null; };
const extractFIO = (s) => { const inn = extractINN(s); let f = String(s || ""); if (inn) f = f.replace(inn, ""); return f.replace(/,/g, " ").replace(/\s+/g, " ").trim(); };
const safeEq = (a, b) => a === b;

const notes = [];
const app = express();
app.use(express.json());
app.use("/api/admin", requireAuth);
const skud = require("../skud")(app, { Student, Group, requireAuth, requireAdmin, normalizeINN, AGENT_API_KEY: KEY, getSession: sessionOf, canAccessGroup: () => true, getAccessibleGroupIds: () => null, notify: async (m) => notes.push(m) });
require("../history-sync")(app, { Student, AGENT_API_KEY: KEY, safeEq, normalizeINN, extractINN, normalizeFIO, extractFIO });

const ADMIN = { role: "admin", username: "boss" };
let base;
async function api(who, method, url, json) {
  const headers = { "content-type": "application/json" };
  if (who) headers["x-test-session"] = JSON.stringify(who);
  const r = await fetch(base + url, { method, headers, body: method === "GET" ? undefined : JSON.stringify(json || {}) });
  return { status: r.status, data: await r.json().catch(() => ({})) };
}
let passed = 0, failed = 0;
const t = async (n, f) => { try { await f(); passed++; console.log("  ✓", n); } catch (e) { failed++; console.log("  ✗", n, "\n    ", e.stack.split("\n").slice(0, 3).join("\n     ")); } };

(async () => {
  const srv = app.listen(0); await new Promise((r) => srv.once("listening", r));
  base = `http://127.0.0.1:${srv.address().port}`;
  const hikIn = await startMock(), hikOut = await startMock();
  const g = await Group.create({ name: "ИТ-1/26" });
  const ali = await Student.create({ fio: "Алиев Али Алиевич", inn: "12345678901234", group: g._id });
  const bek = await Student.create({ fio: "Бекова Бека Бековна", group: g._id });
  const ev = (no, name, time) => ({ employeeNoString: no, name, time });
  const D1 = "2026-10-12", D2 = "2026-10-13";
  // Алиев: 3 входа и 1 выход в D1 (больше страницы в 2 события), Бекова — вход в D2 без имени в событии
  hikIn.events.push(ev("1", "Алиев Али Алиевич 12345678901234", `${D1}T07:50:00+06:00`), ev("1", "Алиев Али Алиевич 12345678901234", `${D1}T08:40:00+06:00`), ev("1", "Алиев Али Алиевич 12345678901234", `${D1}T13:05:00+06:00`), ev("2", "", `${D2}T09:10:00+06:00`));
  hikOut.events.push(ev("1", "Алиев Али Алиевич 12345678901234", `${D1}T15:00:00+06:00`));
  hikIn.users.set("2", { employeeNo: "2", name: "Бекова Бека Бековна" });

  const ag = agentSkud.create({ serverUrl: base, apiKey: KEY, devices: [{ ip: hikIn.ip, direction: "in" }, { ip: hikOut.ip, direction: "out" }], user: "admin", pass: "secret" });
  const dayOf = async (s, d) => (await Student.findById(s._id).lean()).skudDaily.find((x) => x.date === d);

  await t("запрос истории: только админ, проверка дат", async () => {
    assert.equal((await api(null, "POST", "/api/admin/skud-reg/history/run", { from: D1, to: D2 })).status, 401);
    assert.equal((await api(ADMIN, "POST", "/api/admin/skud-reg/history/run", { from: "bad", to: D2 })).status, 400);
    assert.equal((await api(ADMIN, "POST", "/api/admin/skud-reg/history/run", { from: D2, to: D1 })).status, 400);
    assert.equal((await api(ADMIN, "POST", "/api/admin/skud-reg/history/run", { from: "2026-01-01", to: "2026-12-31" })).status, 400);
  });
  await t("запрос → агент получает его один раз", async () => {
    assert.equal((await api(ADMIN, "POST", "/api/admin/skud-reg/history/run", { from: D1, to: D2 })).status, 200);
    assert.equal((await api(ADMIN, "POST", "/api/admin/skud-reg/history/run", { from: D1, to: D2 })).status, 409);
    assert.equal((await api(ADMIN, "GET", "/api/admin/skud-reg/history")).data.status, "queued");
    const c1 = (await api(null, "POST", "/api/agent/jobs/claim", { apiKey: KEY })).data;
    assert.deepEqual(c1.history, { from: D1, to: D2 });
    assert.equal((await api(null, "POST", "/api/agent/jobs/claim", { apiKey: KEY })).data.history, null);
  });
  await t("агент читает терминалы постранично и пишет историю без дублей", async () => {
    await ag.runHistory({ from: D1, to: D2 });
    const d = await dayOf(ali, D1);
    assert.equal(d.inCount, 3); assert.equal(d.outCount, 1); assert.equal(d.present, true);
    assert.ok(d.firstIn.includes("07:50")); assert.ok(d.lastIn.includes("13:05")); assert.ok(d.firstOut.includes("15:00"));
    const b = await dayOf(bek, D2);
    assert.equal(b.present, true); assert.ok(b.firstIn.includes("09:10")); // имя подтянуто с терминала
    assert.equal((await Student.findById(ali._id).lean()).attendance.find((x) => x.date === D1).present, true);
    const st = (await api(ADMIN, "GET", "/api/admin/skud-reg/history")).data;
    assert.equal(st.status, "done"); assert.equal(st.sent, 2);
  });
  await t("повторный запуск идемпотентен; родителям ничего не уходит", async () => {
    const before = JSON.stringify((await Student.findById(ali._id).lean()).skudDaily);
    await ag.runHistory({ from: D1, to: D2 });
    assert.equal(JSON.stringify((await Student.findById(ali._id).lean()).skudDaily), before);
    assert.equal((await Student.findById(ali._id).lean()).skudDaily.length, 1);
  });
  await t("неизвестный человек пропускается, остальные пишутся", async () => {
    const r = await api(null, "POST", "/api/agent/history-sync", { apiKey: KEY, days: [{ rawName: "Неизвестный Иван", date: D1, firstIn: `${D1}T08:00:00+06:00`, lastIn: `${D1}T08:00:00+06:00`, inCount: 1 }, { rawName: "Бекова Бека Бековна", date: D1, firstIn: `${D1}T08:05:00+06:00`, lastIn: `${D1}T08:05:00+06:00`, inCount: 1 }] });
    assert.equal(r.data.okCount, 1); assert.equal(r.data.failCount, 1);
    assert.equal((await api(null, "POST", "/api/agent/history-sync", { apiKey: "bad", days: [] })).status, 403);
  });
  await t("терминал недоступен → ошибка в статусе, не падаем", async () => {
    await hikIn.close(); await hikOut.close();
    await ag.runHistory({ from: D1, to: D1 });
    const st = (await api(ADMIN, "GET", "/api/admin/skud-reg/history")).data;
    assert.equal(st.status, "failed");
  });
  console.log(`\nИтог: ${passed} прошло, ${failed} упало`);
  srv.close();
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
