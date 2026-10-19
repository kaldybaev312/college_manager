// Тест общего журнала (journal-all.js). БД — в памяти.
const Module = require("module");
const path = require("path");
const assert = require("assert/strict");
const fake = require("./fake-mongoose");
const origResolve = Module._resolveFilename;
Module._resolveFilename = function (req, ...rest) {
  return req === "mongoose" ? path.join(__dirname, "fake-mongoose.js") : origResolve.call(this, req, ...rest);
};
const ExcelJS = require("exceljs");
const { Schema } = fake;
const Group = fake.model("Group", new Schema({ name: String, duration: String, shiftStart: { type: String, default: "08:00" }, kind: { type: String, default: "students" }, saturday: { type: Boolean, default: false } }));
const Student = fake.model("Student", new Schema({ fio: String, group: Schema.Types.ObjectId, attendance: [], skudDaily: [] }));
const absence = require("../absence");
const att = require("../attendance-utils");
const { create, courseOf } = require("../journal-all");

(async () => {
  assert.equal(courseOf("ЭГС-1/26", "2026-10"), 1);
  assert.equal(courseOf("ИТ-2/25", "2026-10"), 2);
  assert.equal(courseOf("ИТ-2/25", "2026-03"), 1);
  const g1 = await Group.create({ name: "ЭГС-1/26", duration: "2 года" });
  await Group.create({ name: "ИТ-2/25", duration: "2 года" });
  await Group.create({ name: "Швея-1/26", duration: "10 месяцев" });
  await Group.create({ name: "Мастера", kind: "masters" });
  await Student.create({ fio: "Алиев Али", group: g1._id, skudDaily: [{ date: "2026-10-12", present: true, firstIn: new Date("2026-10-12T07:50:00+06:00").toISOString() }] });
  await Student.create({ fio: "Бекова Бек", group: g1._id, skudDaily: [] });
  const j = create({ Group, Student, absence, att, getTodayKey: () => "2026-10-14" });
  const buf = await j.build("2026-10");
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  assert.deepEqual(wb.worksheets.map((w) => w.name), ["2 курс", "1 курс", "10 месяцев"]);
  const ws = wb.getWorksheet("1 курс");
  assert.equal(ws.getCell(1, 1).value, "ЭГС-1/26");
  assert.equal(ws.getCell(2, 2).value, "ФИО Студента");
  assert.deepEqual([3, 4, 5].map((c) => ws.getCell(2, c).value), [1, 2, 5]); // 3–4 окт — выходные пропущены
  assert.equal(ws.getCell(3, 2).value, "Алиев Али");
  assert.equal(ws.getCell(3, 3).value, "Н"); // 1 окт нет данных
  assert.equal(ws.getCell(2, 10).value, 12);
  assert.equal(ws.getCell(3, 10).value, "П"); // 12 окт пришёл в 07:50
  assert.equal(ws.getCell(4, 10).value, "Н");
  console.log("ok");
})().catch((e) => { console.error("FAIL", e.message); process.exit(1); });
