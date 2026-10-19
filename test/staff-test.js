// Тесты учёта сотрудников (staff-attendance.js). БД — в памяти.  Запуск: npm test
const Module = require("module");
const path = require("path");
const crypto = require("crypto");
const assert = require("assert/strict");

const fake = require("./fake-mongoose");
const origResolve = Module._resolveFilename;
Module._resolveFilename = function (req, ...rest) {
  return req === "mongoose" ? path.join(__dirname, "fake-mongoose.js") : origResolve.call(this, req, ...rest);
};
process.env.SCHEDULE_API_KEY = "sched-key-123";
const express = require("express");
const xlsx = require("xlsx");
const { Schema } = fake;
const Group = fake.model("Group", new Schema({ name: String, shiftStart: { type: String, default: "08:00" }, kind: { type: String, default: "students" }, saturday: { type: Boolean, default: false } }));
const Student = fake.model("Student", new Schema({ fio: String, inn: String, lyceumId: String, group: Schema.Types.ObjectId, attendance: [], skudDaily: [] }));
const TgSub = fake.model("TgSub", new Schema({ chatId: String, role: String }));
const absence = require("../absence");
const { createStaffAttendance, nameKey, parseBell, splitStaff } = require("../staff-attendance");

const MON = "2026-10-12", TUE = "2026-10-13", WED = "2026-10-14";
const iso = (d, t) => new Date(`${d}T${t}:00+06:00`).toISOString();
const day = (d, tin, tout) => ({ date: d, present: true, firstIn: iso(d, tin), lastIn: iso(d, tin), lastOut: tout ? iso(d, tout) : undefined, inCount: 1, outCount: tout ? 1 : 0 });
const safeEq = (a, b) => crypto.timingSafeEqual(crypto.createHash("sha256").update(String(a ?? "")).digest(), crypto.createHash("sha256").update(String(b ?? "")).digest());

let passed = 0, failed = 0;
async function t(name, fn) {
  try { await fn(); passed++; console.log("  ✓", name); } catch (e) { failed++; console.log("  ✗", name, "\n    ", e.message); }
}

(async () => {
  const gT = await Group.create({ name: "Преподаватели", kind: "teachers" });
  const gM = await Group.create({ name: "Мастера", kind: "masters" });
  await Group.create({ name: "ИТ-21" });
  // Иванова: во вторник опоздала на 15 мин и ушла на 20 мин раньше; Петров: во вторник не пришёл; Сидорова — по причине
  const iv = await Student.create({ fio: "Иванова Айгуль Асановна", group: gT._id, skudDaily: [day(MON, "07:55", "14:30"), day(TUE, "08:15", "14:00")] });
  const pe = await Student.create({ fio: "Петров Бакыт Сапарович", group: gT._id, skudDaily: [day(MON, "09:20", "12:45")] });
  const si = await Student.create({ fio: "Сидорова Нуриля Кадыровна", group: gT._id, skudDaily: [] });
  await Student.create({ fio: "Мастеров Эмир", group: gM._id, skudDaily: [day(MON, "08:00", "17:00")] });
  await absence.createReason({ student: si, body: { type: "sick", dateFrom: MON, dateTo: TUE }, source: "admin" });
  await TgSub.create({ chatId: "A1", role: "admin" });

  let today = TUE, nowMin = 10 * 60;
  const staff = createStaffAttendance({ Group, Student, TgSub, getTodayKey: () => today, getBishkekMinutes: () => nowMin });
  const app = express();
  app.use(express.json());
  const isAdmin = (req) => req.headers["x-role"] === "admin";
  staff.registerRoutes(app, {
    requireAdmin: (req, res, next) => (isAdmin(req) ? next() : res.status(403).json({ error: "Admin only" })),
    requireAuth: (req, res, next) => next(),
    safeEq,
    contentDisposition: (n) => `attachment; filename="${encodeURIComponent(n)}"`,
  });
  const srv = app.listen(0);
  const base = `http://127.0.0.1:${srv.address().port}`;
  const call = async (method, url, body, headers = { "x-role": "admin" }) => {
    const r = await fetch(base + url, { method, headers: { "Content-Type": "application/json", ...headers }, body: body ? JSON.stringify(body) : undefined });
    return { status: r.status, data: await r.json().catch(() => null), r };
  };

  console.log("Расписание:");
  await t("разбор звонков, ФИО и подгрупп", () => {
    assert.deepEqual(parseBell("08:00-09:20"), { start: "08:00", end: "09:20" });
    assert.deepEqual(parseBell("8.00 – 9.20"), { start: "08:00", end: "09:20" });
    assert.equal(nameKey("Иванова Айгуль Асановна"), nameKey("Иванова А.А."));
    assert.equal(nameKey("Королёва А.Б."), nameKey("Королева Анна Борисовна"));
    assert.deepEqual(splitStaff("Иванова А.А. / Петров Б.С."), ["Иванова А.А.", "Петров Б.С."]);
  });
  const payload = {
    groups: [
      { gradeName: "ИТ-21", weekStart: MON, lessons: [
        { weekday: 1, lesson: 1, bell: "08:00-09:20", subjectName: "Математика", staffName: "Иванова А.А.", roomName: "12" },
        { weekday: 1, lesson: 4, bell: "13:00-14:20", subjectName: "Физика", staffName: "Иванова А.А.", roomName: "12" },
        { weekday: 2, lesson: 1, bell: "08:00-09:20", subjectName: "Математика", staffName: "Иванова А.А. / Петров Б.С.", roomName: "12" },
        { weekday: 2, lesson: 4, bell: "13:00-14:20", subjectName: "Физика", staffName: "Иванова А.А.", roomName: "12" },
        { weekday: 2, lesson: 2, bell: "09:30-10:50", subjectName: "Химия", staffName: "Сидорова Н.К.", roomName: "5" },
        { weekday: 1, lesson: 2, bell: "09:30-10:50", subjectName: "Право", staffName: "Неизвестный Т.Т.", roomName: "3" },
      ] },
      { gradeName: "СВ-11", weekStart: MON, lessons: [{ weekday: 1, lesson: 3, bell: "11:20-12:40", subjectName: "Черчение", staffName: "Петров Бакыт", roomName: "7" }] },
    ],
  };
  await t("приём из редактора только по ключу SCHEDULE_API_KEY", async () => {
    assert.equal((await call("POST", "/api/agent/teacher-schedule", payload, { "x-api-key": "wrong" })).status, 403);
    const r = await call("POST", "/api/agent/teacher-schedule", payload, { "x-api-key": "sched-key-123" });
    assert.equal(r.status, 200, JSON.stringify(r.data));
    assert.equal(r.data.lessons, 8);
    assert.deepEqual(r.data.unmatched, ["Неизвестный Т.Т."]);
  });
  await t("повторная отправка заменяет неделю группы, а не дублирует", async () => {
    await call("POST", "/api/agent/teacher-schedule", payload, { "x-api-key": "sched-key-123" });
    const d = await staff.day(MON);
    assert.equal(d.teachers.find((x) => x.fio.startsWith("Иванова")).lessons.length, 2);
  });

  console.log("Преподаватели:");
  await t("вовремя, опоздание, ранний уход, часы", async () => {
    const mon = (await staff.day(MON)).teachers.find((x) => x.id === String(iv._id));
    assert.deepEqual([mon.status, mon.first, mon.last, mon.late, mon.early, mon.minutes], ["ok", "08:00", "14:20", 0, 0, 395]);
    const tue = (await staff.day(TUE)).teachers.find((x) => x.id === String(iv._id));
    assert.deepEqual([tue.status, tue.in, tue.late, tue.out, tue.early], ["late", "08:15", 15, "14:00", 20]);
  });
  await t("ФИО без отчества тоже находится (Петров Бакыт), к уроку 11:20 пришёл в 09:20", async () => {
    const mon = (await staff.day(MON)).teachers.find((x) => x.id === String(pe._id));
    assert.deepEqual([mon.status, mon.first, mon.late], ["ok", "11:20", 0]);
  });
  await t("не пришёл в день с уроками; по причине — не нарушение; без уроков — не считается", async () => {
    const tue = await staff.day(TUE);
    assert.equal(tue.teachers.find((x) => x.id === String(pe._id)).status, "absent");
    assert.equal(tue.teachers.find((x) => x.id === String(si._id)).status, "excused");
    assert.equal((await staff.day(WED)).teachers.find((x) => x.id === String(iv._id)).status, "free");
  });
  await t("ручное сопоставление ФИО из расписания", async () => {
    let st = await call("GET", "/api/admin/staff/schedule/status");
    assert.deepEqual(st.data.unmatched.map((u) => u.name), ["Неизвестный Т.Т."]);
    assert.equal((await call("POST", "/api/admin/staff/alias", { name: "Неизвестный Т.Т.", staffId: String(si._id) })).status, 200);
    st = await call("GET", "/api/admin/staff/schedule/status");
    assert.equal(st.data.unmatched.length, 0);
    const mon = (await staff.day(MON)).teachers.find((x) => x.id === String(si._id));
    assert.equal(mon.lessons.length, 1);
    assert.equal(mon.status, "excused");
  });

  console.log("Мастера:");
  await t("был / не был и часы, без опозданий", async () => {
    const m = (await staff.day(MON)).masters[0];
    assert.deepEqual([m.status, m.in, m.out, m.minutes], ["ok", "08:00", "17:00", 540]);
    assert.equal((await staff.day(TUE)).masters[0].status, "absent");
    assert.equal((await staff.day("2026-10-17")).masters[0].status, "free");
  });

  console.log("Уведомления и табель:");
  await t("админу: опоздание и неявка — по одному разу", async () => {
    const sent = [];
    const bot = { sendMessage: async (chatId, text) => sent.push({ chatId, text }) };
    today = TUE; nowMin = 9 * 60;
    await staff.alertsTick(bot);
    assert.ok(sent.some((m) => /Иванова.*опоздал на 15 мин/.test(m.text)));
    assert.ok(sent.some((m) => /Петров.*не пришёл к первому уроку 08:00/.test(m.text)));
    assert.ok(!sent.some((m) => /Сидорова/.test(m.text)));
    const n = sent.length;
    await staff.alertsTick(bot);
    assert.equal(sent.length, n);
    await staff.eveningSummary(bot);
    assert.match(sent.at(-1).text, /Не пришли:\n• Петров/);
    assert.match(sent.at(-1).text, /Ушли раньше:\n• Иванова/);
    assert.match(sent.at(-1).text, /Мастера, которых не было:\n• Мастеров/);
  });
  await t("табель за месяц и Excel", async () => {
    today = WED;
    const m = await staff.month("2026-10");
    const ivr = m.teachers.find((r) => r.id === String(iv._id));
    assert.deepEqual([ivr.totals.lessonDays, ivr.totals.attended, ivr.totals.late, ivr.totals.lateMin, ivr.totals.early], [2, 2, 1, 15, 1]);
    const r = await fetch(base + "/api/admin/staff/month?month=2026-10&format=xlsx", { headers: { "x-role": "admin" } });
    const wb = xlsx.read(Buffer.from(await r.arrayBuffer()));
    assert.deepEqual(wb.SheetNames, ["Преподаватели", "Мастера"]);
    const rows = xlsx.utils.sheet_to_json(wb.Sheets["Преподаватели"], { header: 1 });
    assert.ok(rows.some((row) => row[1] === "Петров Бакыт Сапарович" && row.includes("Н")));
  });
  await t("только администратор", async () => {
    assert.equal((await call("GET", "/api/admin/staff/day", null, { "x-role": "curator" })).status, 403);
    assert.equal((await call("GET", "/api/admin/staff/month", null, {})).status, 403);
  });

  console.log("Excel из «Отчётов» редактора:");
  await t("загрузка файла отчёта", async () => {
    const aoa = [
      ["Расписание групп 1 курс 2026-2027 учебный год"],
      ["с 19.10.2026 по 24.10.2026"],
      [], ["Группа"], ["ИТ-21"], ["Дни", "Урок", "Звонок", "Название", "Преподаватель", "Кабинет"],
      ["Понедельник", 1, "08:00-09:20", "Математика", "Иванова Айгуль Асановна", "12"],
      ["Понедельник", 2, "09:30-10:50", "Физика", "Петров Б.С.", "12"],
      ["Вторник", 1, "08:00-09:20", "Математика", "Иванова Айгуль Асановна", "12"],
      [], ["Зам. Директора по УПМР"],
    ];
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(aoa), "Лист1");
    const buf = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
    const groups = staff.parseReportWorkbook(buf);
    assert.equal(groups.length, 1);
    assert.equal(groups[0].weekStart, "2026-10-19");
    assert.equal(groups[0].lessons.length, 3);
    const fd = new FormData();
    fd.append("files", new Blob([buf]), "report.xlsx");
    const r = await fetch(base + "/api/admin/staff/schedule/upload", { method: "POST", headers: { "x-role": "admin" }, body: fd });
    const j = await r.json();
    assert.equal(r.status, 200, JSON.stringify(j));
    assert.equal(j.lessons, 3);
    const d = await staff.day("2026-10-19");
    assert.equal(d.teachers.find((x) => x.id === String(pe._id)).first, "09:30");
  });

  await t("администрация: был/не был и часы, табель и лист Excel", async () => {
    const gA = await Group.create({ name: "Администрация", kind: "admin" });
    await Student.create({ fio: "Директоров Талант Сапарович", group: gA._id, skudDaily: [day(MON, "08:00", "17:00")] });
    await Student.create({ fio: "Завучева Зарина Усеновна", group: gA._id, skudDaily: [] });
    const d = (await call("GET", `/api/admin/staff/day?date=${MON}`)).data;
    assert.equal(d.admins.length, 2);
    assert.equal(d.admins.find((x) => x.fio.startsWith("Директоров")).status, "ok");
    assert.equal(d.admins.find((x) => x.fio.startsWith("Директоров")).minutes, 540);
    assert.equal(d.admins.find((x) => x.fio.startsWith("Завучева")).status, "absent");
    assert.equal(d.teachers.length >= 3 && d.masters.length, 1); // остальные категории не смешиваются
    const m = (await call("GET", "/api/admin/staff/month?month=2026-10")).data;
    assert.equal(m.admins.length, 2);
    assert.equal(m.admins.find((x) => x.fio.startsWith("Директоров")).totals.present, 1);
    const x = await fetch(`${base}/api/admin/staff/month?month=2026-10&format=xlsx`, { headers: { "x-role": "admin" } });
    const wb = xlsx.read(Buffer.from(await x.arrayBuffer()));
    assert.deepEqual(wb.SheetNames, ["Преподаватели", "Мастера", "Администрация"]);
  });

  console.log(`\nИтог: ${passed} прошло, ${failed} упало`);
  srv.close();
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
