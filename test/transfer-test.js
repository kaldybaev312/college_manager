// Перевод студента между группами. БД в памяти.
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
const Group = fake.model("Group", new Schema({ name: String, kind: { type: String, default: "students" } }));
const Student = fake.model("Student", new Schema({ fio: String, group: Schema.Types.ObjectId }));
const { AbsenceReason } = require("../absence");
const sessionOf = (req) => { try { return JSON.parse(req.headers["x-test-session"]); } catch { return null; } };
const requireAdmin = (req, res, next) => (sessionOf(req)?.role === "admin" ? next() : res.status(403).json({ error: "Admin only" }));
const log = [];
const app = express();
app.use(express.json());
require("../student-transfer")(app, { Student, Group, AbsenceReason, requireAdmin, audit: async (_r, a, t, d) => log.push([a, t, d]) });

let base, passed = 0, failed = 0;
const post = async (who, id, body) => {
  const r = await fetch(`${base}/api/admin/student/${id}/transfer`, { method: "POST", headers: { "content-type": "application/json", ...(who ? { "x-test-session": JSON.stringify(who) } : {}) }, body: JSON.stringify(body) });
  return { status: r.status, data: await r.json().catch(() => ({})) };
};
const t = async (n, f) => { try { await f(); passed++; console.log("  ✓", n); } catch (e) { failed++; console.log("  ✗", n, "\n    ", e.stack.split("\n").slice(0, 3).join("\n     ")); } };

(async () => {
  const srv = app.listen(0); await new Promise((r) => srv.once("listening", r));
  base = `http://127.0.0.1:${srv.address().port}`;
  const A = await Group.create({ name: "ИТ-1/26" }), B = await Group.create({ name: "ЭГС-1/26" }), T = await Group.create({ name: "Преподаватели", kind: "teachers" });
  const s = await Student.create({ fio: "Алиев Али", group: A._id });
  await AbsenceReason.create({ student: s._id, group: A._id, dateFrom: "2026-10-05", dateTo: "2026-10-05", code: "Б", type: "sick" }).catch(() => AbsenceReason.create({ student: s._id, group: A._id }));
  const ADMIN = { role: "admin" };
  await t("только админ", async () => assert.equal((await post({ role: "curator" }, s._id, { groupId: B._id })).status, 403));
  await t("проверки: нет группы, та же группа, группа сотрудников, неизвестный студент", async () => {
    assert.equal((await post(ADMIN, s._id, {})).status, 400);
    assert.equal((await post(ADMIN, s._id, { groupId: A._id })).status, 400);
    assert.equal((await post(ADMIN, s._id, { groupId: T._id })).status, 400);
    assert.equal((await post(ADMIN, "000000000000000000000000", { groupId: B._id })).status, 404);
  });
  await t("перевод: студент и его причины переезжают, пишется в журнал", async () => {
    const r = await post(ADMIN, s._id, { groupId: B._id });
    assert.equal(r.status, 200); assert.equal(r.data.from, "ИТ-1/26"); assert.equal(r.data.to, "ЭГС-1/26");
    assert.equal(String((await Student.findById(s._id).lean()).group), String(B._id));
    const rs = await AbsenceReason.find({ student: s._id }).lean();
    assert.ok(rs.length && rs.every((x) => String(x.group) === String(B._id)));
    assert.deepEqual(log.at(-1), ["student_transfer", "Алиев Али", "ИТ-1/26 → ЭГС-1/26"]);
  });
  console.log(`\nИтог: ${passed} прошло, ${failed} упало`);
  srv.close(); process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
