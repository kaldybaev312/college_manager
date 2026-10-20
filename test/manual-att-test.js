// Ручная правка посещаемости (только админ). БД в памяти.
const Module = require("module");
const path = require("path");
const assert = require("assert/strict");
const fake = require("./fake-mongoose");
const origResolve = Module._resolveFilename;
Module._resolveFilename = function (req, ...rest) {
  return req === "mongoose" ? path.join(__dirname, "fake-mongoose.js") : origResolve.call(this, req, ...rest);
};
const express = require("express");
const { Schema } = fake;
const att = require("../attendance-utils");
const Student = fake.model("Student", new Schema({ fio: String, attendance: [], skudDaily: [] }));
const sessionOf = (req) => { try { return JSON.parse(req.headers["x-test-session"]); } catch { return null; } };
const requireAdmin = (req, res, next) => (sessionOf(req)?.role === "admin" ? next() : res.status(403).json({ error: "Admin only" }));
const log = [];
const app = express();
app.use(express.json());
require("../attendance-manual")(app, { Student, requireAdmin, audit: async (_r, a, t, d) => log.push([a, t, d]), att, getTodayKey: () => "2026-10-20", getSession: sessionOf });

let base, passed = 0, failed = 0;
const post = async (who, body) => {
  const r = await fetch(`${base}/api/admin/attendance-manual`, { method: "POST", headers: { "content-type": "application/json", ...(who ? { "x-test-session": JSON.stringify(who) } : {}) }, body: JSON.stringify(body) });
  return { status: r.status, data: await r.json().catch(() => ({})) };
};
const t = async (n, f) => { try { await f(); passed++; console.log("  ✓", n); } catch (e) { failed++; console.log("  ✗", n, "\n    ", e.stack.split("\n").slice(0, 3).join("\n     ")); } };
const ADMIN = { role: "admin", username: "boss" };

(async () => {
  const srv = app.listen(0); await new Promise((r) => srv.once("listening", r));
  base = `http://127.0.0.1:${srv.address().port}`;
  const s = await Student.create({ fio: "Алиев Али" });
  const D = "2026-10-19";
  const day = async () => { const x = await Student.findById(s._id).lean(); return { row: x.skudDaily.find((r) => r.date === D), a: x.attendance.find((r) => r.date === D), st: att.dayStatus(x, D, att.shiftMinutes("08:00")) }; };

  await t("только админ", async () => assert.equal((await post({ role: "curator" }, { studentId: s._id, date: D, status: "present", time: "08:05" })).status, 403));
  await t("проверки: дата, будущее, время, статус", async () => {
    assert.equal((await post(ADMIN, { studentId: s._id, date: "bad", status: "present", time: "08:00" })).status, 400);
    assert.equal((await post(ADMIN, { studentId: s._id, date: "2026-10-21", status: "present", time: "08:00" })).status, 400);
    assert.equal((await post(ADMIN, { studentId: s._id, date: D, status: "present", time: "25:00" })).status, 400);
    assert.equal((await post(ADMIN, { studentId: s._id, date: D, status: "maybe" })).status, 400);
    assert.equal((await post(ADMIN, { studentId: "000000000000000000000000", date: D, status: "absent" })).status, 404);
  });
  await t("поставить приход: был, опоздал по времени, помечено вручную", async () => {
    const r = await post(ADMIN, { studentId: s._id, date: D, status: "present", time: "08:25", outTime: "15:00" });
    assert.equal(r.status, 200); assert.equal(r.data.before, "не был");
    const { row, a, st } = await day();
    assert.equal(row.present, true); assert.equal(row.manual, true); assert.equal(row.manualBy, "boss");
    assert.ok(row.firstIn.startsWith(`${D}T08:25`)); assert.ok(row.lastOut.startsWith(`${D}T15:00`));
    assert.equal(a.present, true); assert.equal(st.present, true); assert.equal(st.late, true); assert.equal(st.time, "08:25");
  });
  await t("повторная правка менять время можно, дублей нет", async () => {
    await post(ADMIN, { studentId: s._id, date: D, status: "present", time: "07:55" });
    const x = await Student.findById(s._id).lean();
    assert.equal(x.skudDaily.filter((r) => r.date === D).length, 1);
    assert.equal((await day()).st.late, false);
  });
  await t("«не был»: присутствие снято и в СКУД-записи, и в отметке", async () => {
    const r = await post(ADMIN, { studentId: s._id, date: D, status: "absent" });
    assert.match(r.data.before, /^был 07:55/);
    const { row, a, st } = await day();
    assert.equal(st.present, false); assert.equal(row.firstIn, undefined); assert.equal(a.present, false);
  });
  await t("правки пишутся в журнал с прежним значением", async () => {
    assert.ok(log.length >= 3);
    assert.deepEqual(log.at(-1), ["attendance_manual", "Алиев Али", `${D}: был 07:55 → не был`]);
  });
  console.log(`\nИтог: ${passed} прошло, ${failed} упало`);
  srv.close(); process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
