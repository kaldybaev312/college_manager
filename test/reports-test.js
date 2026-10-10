// Тесты: подпись Telegram initData, мини-приложение, причины отсутствия, отчёты.
// БД имитируется в памяти (fake-mongoose).  Запуск: npm test
const Module = require("module");
const path = require("path");
const crypto = require("crypto");
const assert = require("assert/strict");

const fake = require("./fake-mongoose");
const origResolve = Module._resolveFilename;
Module._resolveFilename = function (req, ...rest) {
  return req === "mongoose" ? path.join(__dirname, "fake-mongoose.js") : origResolve.call(this, req, ...rest);
};

process.env.TG_TOKEN = "123456:TEST-TOKEN";
const express = require("express");
const xlsx = require("xlsx");
const att = require("../attendance-utils");
const absence = require("../absence");
const { createReports, resolvePeriod, previousPeriod } = require("../reports");
const registerWebapp = require("../webapp");
const { createWeeklyCards } = require("../weekly-card");
const { createAssistant } = require("../ai-assistant");
const { verifyInitData } = registerWebapp;

const { Schema } = fake;
const Group = fake.model("Group", new Schema({ name: String, shiftStart: { type: String, default: "08:00" }, kind: { type: String, default: "students" }, saturday: { type: Boolean, default: false } }));
const Student = fake.model("Student", new Schema({ fio: String, inn: String, group: Schema.Types.ObjectId, attendance: [], skudDaily: [] }));
const TgSub = fake.model("TgSub", new Schema({ chatId: String, name: String, role: String, linkedInns: [], curatorId: String }));
const Curator = fake.model("Curator", new Schema({ username: String, groups: [] }));
const SummaryRecord = fake.model("SummaryRecord", new Schema({ group: String, student: String, month: String, grades: Object }));

function initData(userId, { token = process.env.TG_TOKEN, authDate = Math.floor(Date.now() / 1000), tamper = false } = {}) {
  const p = new URLSearchParams({ auth_date: String(authDate), query_id: "AAE", user: JSON.stringify({ id: userId, first_name: "Тест" }) });
  const dcs = [...p.entries()].sort(([a], [b]) => (a < b ? -1 : 1)).map(([k, v]) => `${k}=${v}`).join("\n");
  const secret = crypto.createHmac("sha256", "WebAppData").update(token).digest();
  p.set("hash", crypto.createHmac("sha256", secret).update(dcs).digest("hex"));
  if (tamper) p.set("user", JSON.stringify({ id: 1, first_name: "Админ" }));
  return p.toString();
}

// "2026-10-08T08:15:00+06:00" → ISO UTC, как хранит агент
const isoAt = (dk, hhmm) => new Date(`${dk}T${hhmm}:00+06:00`).toISOString();
const day = (dk, hhmm) => ({ date: dk, present: true, firstIn: isoAt(dk, hhmm), inCount: 1, outCount: 0 });

let passed = 0,
  failed = 0;
async function t(name, fn) {
  try {
    await fn();
    passed++;
    console.log("  ✓", name);
  } catch (e) {
    failed++;
    console.log("  ✗", name, "\n    ", e.message);
  }
}

(async () => {
  const TODAY = "2026-10-10"; // суббота
  const getTodayKey = () => TODAY;

  const gA = await Group.create({ name: "ИТ-21", shiftStart: "08:30" });
  const gB = await Group.create({ name: "СВ-11", shiftStart: "08:30" });
  await Group.create({ name: "Учителя", kind: "teachers" });
  const s1 = await Student.create({ fio: "Асанов Бакыт", inn: "111111111111", group: gA._id, skudDaily: [day("2026-10-05", "08:10"), day("2026-10-06", "08:45"), day("2026-10-07", "08:00"), day("2026-10-08", "08:05"), day("2026-10-09", "08:20")] });
  const s2 = await Student.create({ fio: "Бекова Айгуль", inn: "222222222222", group: gA._id, skudDaily: [day("2026-10-05", "08:00")] });
  const s3 = await Student.create({ fio: "Чужой Студент", inn: "333333333333", group: gB._id, skudDaily: [] });
  const cur = await Curator.create({ username: "cur", groups: [String(gA._id)] });
  await TgSub.create({ chatId: "500", role: "curator", curatorId: String(cur._id) });
  await TgSub.create({ chatId: "900", role: "admin" });

  console.log("Подпись Telegram:");
  await t("верная подпись принимается", () => assert.equal(verifyInitData(initData(500), process.env.TG_TOKEN).user.id, 500));
  await t("подмена пользователя отклоняется", () => assert.equal(verifyInitData(initData(500, { tamper: true }), process.env.TG_TOKEN), null));
  await t("чужой токен отклоняется", () => assert.equal(verifyInitData(initData(500, { token: "1:OTHER" }), process.env.TG_TOKEN), null));
  await t("устаревшие данные отклоняются", () =>
    assert.equal(verifyInitData(initData(500, { authDate: Math.floor(Date.now() / 1000) - 8 * 86400 }), process.env.TG_TOKEN), null));

  console.log("Расчёты:");
  await t("время по Бишкеку и опоздание", () => {
    assert.equal(att.localTime(isoAt("2026-10-06", "08:45")).str, "08:45");
    const st = att.dayStatus(s1, "2026-10-06", att.shiftMinutes("08:30"));
    assert.deepEqual(st, { present: true, late: true, time: "08:45" });
  });
  await t("неделя и прошлые периоды", () => {
    assert.deepEqual(att.weekRange("2026-10-10"), { from: "2026-10-05", to: "2026-10-11" });
    const w = previousPeriod("week", "2026-10-12");
    assert.equal(w.from, "2026-10-05");
    assert.equal(w.to, "2026-10-11");
    const m = previousPeriod("month", "2026-01-01");
    assert.equal(m.from, "2025-12-01");
    assert.equal(m.to, "2025-12-31");
    assert.throws(() => resolvePeriod("week", "2026-13-40"));
  });
  await t("причина: проверка ввода", () => {
    assert.throws(() => absence.cleanInput({ type: "sick", dateFrom: "2026-10-09", dateTo: "2026-10-01" }), /раньше/);
    assert.throws(() => absence.cleanInput({ type: "x", dateFrom: "2026-10-09" }), /тип/);
    assert.throws(() => absence.cleanInput({ type: "sick", dateFrom: "2026-01-01", dateTo: "2026-03-01" }), /31/);
    assert.equal(absence.cleanInput({ type: "other", dateFrom: "2026-10-09" }).excused, false);
  });

  // ── Мини-приложение по HTTP ──
  const subjectsStub = async () => [{ key: "math", label: "Математика", section: "theory" }];
  const weeklyCards = createWeeklyCards({ Group, Student, TgSub, SummaryRecord, ensureGroupSummarySubjects: subjectsStub, getTodayKey });
  const app = express();
  app.use(express.json());
  const sent = [];
  const fakeBot = {
    sendMessage: async (chatId, text) => sent.push({ chatId, text }),
    sendDocument: async (chatId, buf, opts, file) => sent.push({ chatId, doc: file.filename, buf, caption: opts.caption }),
  };
  registerWebapp(app, {
    Group,
    Student,
    TgSub,
    Curator,
    SummaryRecord,
    ensureGroupSummarySubjects: async () => [{ key: "math", label: "Математика", section: "theory" }],
    formatMonthLabel: (m) => m,
    notifyParents: async () => {},
    getTodayKey,
    getBot: () => fakeBot,
    mdv2: (v) => v,
    weeklyCards,
  });
  const reports = createReports({ Group, Student, Curator, TgSub, getTodayKey });
  const srv = app.listen(0);
  const base = `http://127.0.0.1:${srv.address().port}`;
  const call = async (uid, method, url, body) => {
    const headers = { "Content-Type": "application/json" };
    if (uid) headers["x-tg-init-data"] = initData(uid);
    const r = await fetch(base + url, { method, headers, body: body ? JSON.stringify(body) : undefined });
    return { status: r.status, data: await r.json().catch(() => null) };
  };

  console.log("Мини-приложение:");
  await t("без подписи — 401", async () => {
    assert.equal((await call(null, "GET", "/api/webapp/me")).status, 401);
    const r = await fetch(base + "/api/webapp/students?gid=" + gA._id, { headers: { "x-tg-chat-id": "900" } });
    assert.equal(r.status, 401);
  });
  await t("куратор видит свою группу и не видит чужую", async () => {
    assert.equal((await call(500, "GET", `/api/webapp/attendance/month?gid=${gA._id}&month=2026-10`)).status, 200);
    assert.equal((await call(500, "GET", `/api/webapp/students?gid=${gB._id}`)).status, 403);
    assert.equal((await call(500, "GET", `/api/webapp/grades?gid=${gB._id}&month=2026-10`)).status, 403);
  });
  await t("чужой пользователь не видит данные по ИНН", async () => {
    assert.equal((await call(777, "GET", `/api/webapp/parent/attendance?inn=111111111111&month=2026-10`)).status, 403);
    assert.equal((await call(777, "GET", `/api/webapp/parent/grades?inn=111111111111&month=2026-10`)).status, 403);
  });
  await t("оценку нельзя поставить студенту из чужой группы", async () => {
    const r = await call(500, "POST", "/api/webapp/grades/set", { gid: String(gA._id), studentId: String(s3._id), month: "2026-10", grades: { math: "5" } });
    assert.equal(r.status, 404);
  });
  await t("куратор не может стать родителем, перебор ИНН ограничен", async () => {
    assert.equal((await call(500, "POST", "/api/webapp/parent/link", { inn: "222222222222" })).status, 400);
    for (let i = 0; i < 5; i++) assert.equal((await call(778, "POST", "/api/webapp/parent/link", { inn: "999999999999" })).status, 404);
    assert.equal((await call(778, "POST", "/api/webapp/parent/link", { inn: "222222222222" })).status, 429);
  });
  await t("родитель привязывает ребёнка и видит только его", async () => {
    const r = await call(600, "POST", "/api/webapp/parent/link", { inn: "222222222222" });
    assert.equal(r.status, 200);
    assert.equal(r.data.fio, "Бекова Айгуль");
    assert.equal((await call(600, "GET", `/api/webapp/parent/attendance?inn=222222222222&month=2026-10`)).status, 200);
    assert.equal((await call(600, "GET", `/api/webapp/parent/attendance?inn=111111111111&month=2026-10`)).status, 403);
  });

  console.log("Причины отсутствия:");
  let parentReasonId;
  await t("родитель сообщает о болезни — куратор получает уведомление", async () => {
    sent.length = 0;
    const r = await call(600, "POST", "/api/webapp/absences", { inn: "222222222222", type: "sick", dateFrom: "2026-10-06", dateTo: "2026-10-07", comment: "температура" });
    assert.equal(r.status, 200, JSON.stringify(r.data));
    assert.equal(r.data.confirmed, false);
    parentReasonId = r.data.id;
    await new Promise((res) => setTimeout(res, 20));
    assert.ok(sent.some((m) => m.chatId === "500" && /Бекова/.test(m.text)));
  });
  await t("родитель не может указать причину чужому ребёнку", async () => {
    assert.equal((await call(600, "POST", "/api/webapp/absences", { inn: "111111111111", type: "sick", dateFrom: "2026-10-06" })).status, 403);
  });
  await t("куратор видит и подтверждает причину", async () => {
    const m = await call(500, "GET", `/api/webapp/attendance/month?gid=${gA._id}&month=2026-10`);
    const bek = m.data.students.find((s) => s.fio === "Бекова Айгуль");
    assert.equal(bek.statuses["2026-10-06"], "e");
    assert.equal(bek.reasons["2026-10-06"].code, "Б?");
    const c = await call(500, "PUT", `/api/webapp/absences/${parentReasonId}`, {});
    assert.equal(c.data.confirmed, true);
  });
  await t("куратор не трогает причины чужой группы", async () => {
    assert.equal((await call(500, "POST", "/api/webapp/absences", { studentId: String(s3._id), type: "sick", dateFrom: "2026-10-06" })).status, 403);
  });
  await t("куратор ставит «другую причину» (не уважительная)", async () => {
    const r = await call(500, "POST", "/api/webapp/absences", { studentId: String(s2._id), type: "other", dateFrom: "2026-10-08", comment: "без справки" });
    assert.equal(r.data.excused, false);
    assert.equal(r.data.confirmed, true);
  });

  console.log("Отчёты:");
  await t("недельный отчёт по группе считает посещения, опоздания и причины", async () => {
    const rep = await reports.buildGroupReport(await Group.findById(gA._id).lean(), resolvePeriod("week", "2026-10-07"));
    assert.equal(rep.workdays, 5);
    const a = rep.rows.find((r) => r.fio.startsWith("Асанов"));
    assert.deepEqual([a.present, a.late, a.unexcused, a.pct], [5, 1, 0, 100]);
    const b = rep.rows.find((r) => r.fio.startsWith("Бекова"));
    // пн был, вт-ср болезнь (подтверждена), чт другая причина, пт без причины
    assert.deepEqual([b.present, b.excused, b.unexcused, b.pct], [1, 2, 2, 20]);
    assert.deepEqual(b.cells.slice(0, 5), ["П", "Б", "Б", "Д", "Н"]);
    assert.equal(rep.totals.pct, 60);
  });
  await t("учебная суббота: суббота входит в учебные дни только у своей группы", async () => {
    assert.equal(att.isWeekend("2026-10-10"), true);
    assert.equal(att.isWeekend("2026-10-10", true), false);
    assert.equal(att.isWeekend("2026-10-11", true), true);
    const gS = await Group.create({ name: "ПК-31", shiftStart: "08:30", saturday: true });
    await Student.create({ fio: "Суббота Студент", inn: "444444444444", group: gS._id, skudDaily: [day("2026-10-05", "08:00")] });
    const range = resolvePeriod("week", "2026-10-07");
    const rs = await reports.buildGroupReport(await Group.findById(gS._id).lean(), range);
    assert.equal(rs.workdays, 6);
    assert.equal(rs.rows[0].cells[5], "Н");
    const ra = await reports.buildGroupReport(await Group.findById(gA._id).lean(), range);
    assert.equal(ra.workdays, 5);
    assert.equal(ra.rows[0].cells[5], "");
  });
  await t("Excel по группе открывается и содержит студентов", async () => {
    const rep = await reports.buildGroupReport(await Group.findById(gA._id).lean(), resolvePeriod("month", "2026-10"));
    const wb = xlsx.read(reports.groupWorkbook(rep));
    const rows = xlsx.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 });
    assert.ok(rows.some((r) => r[1] === "Асанов Бакыт"));
    assert.ok(String(rows[0][0]).includes("Месячный"));
  });
  await t("авторассылка: куратору — его группа, админу — все группы без сотрудников", async () => {
    sent.length = 0;
    const r = await reports.sendReports(fakeBot, "week");
    assert.equal(r.range.from, "2026-09-28");
    const toCur = sent.filter((m) => m.chatId === "500");
    assert.equal(toCur.length, 1);
    assert.ok(toCur[0].doc.includes("ИТ-21"));
    const toAdm = sent.find((m) => m.chatId === "900");
    const wb = xlsx.read(toAdm.buf);
    assert.ok(wb.SheetNames.includes("Итог") && wb.SheetNames.includes("ИТ-21"));
    assert.ok(!wb.SheetNames.includes("Учителя"));
  });

  console.log("Итог недели:");
  await t("данные карточки: дни, опоздания, причины, оценки", async () => {
    await SummaryRecord.create({ group: String(gA._id), student: String(s2._id), month: "2026-10", grades: { math: "5" } });
    const c = await weeklyCards.cardFor({ inn: "222222222222" }, "2026-10-07");
    assert.deepEqual(c.data.cells.map((x) => x.st), ["p", "e", "e", "e", "a"]);
    assert.deepEqual([c.data.present, c.data.excused, c.data.absent], [1, 2, 2]);
    assert.deepEqual(c.data.grades, [{ label: "Математика", value: "5" }]);
    assert.equal(c.png.slice(1, 4).toString(), "PNG");
  });
  await t("родитель получает картинку только своего ребёнка", async () => {
    const ok = await fetch(base + "/api/webapp/parent/week-card?inn=222222222222", { headers: { "x-tg-init-data": initData(600) } });
    assert.equal(ok.status, 200);
    assert.equal(ok.headers.get("content-type"), "image/png");
    assert.equal((await call(600, "GET", "/api/webapp/parent/week-card?inn=111111111111")).status, 403);
  });
  await t("рассылка родителям (пятница)", async () => {
    sent.length = 0;
    fakeBot.sendPhoto = async (chatId, buf, opts) => sent.push({ chatId, photo: true, caption: opts.caption });
    const r = await weeklyCards.sendAll(fakeBot, { onlySaturday: false });
    assert.equal(r.sent, 1);
    assert.ok(sent[0].chatId === "600" && /Бекова/.test(sent[0].caption));
  });

  console.log("Геолокация практики:");
  await t("расстояние, координаты из ссылок, решение по радиусу", () => {
    const g = require("../practice")._geo;
    assert.deepEqual(g.readPoint({ text: "https://2gis.kg/bishkek/geo/74.5698,42.8746" }), { lat: 42.8746, lng: 74.5698 });
    assert.deepEqual(g.readPoint({ text: "https://maps.google.com/?q=42.8746,74.5698" }), { lat: 42.8746, lng: 74.5698 });
    assert.equal(g.readPoint({ text: "нет координат" }), null);
    const org = { lat: 42.8746, lng: 74.5698, radius: 300 };
    assert.equal(g.judgeCheckIn(org, { lat: 42.876, lng: 74.5698 }, 20).ok, true);
    assert.equal(g.judgeCheckIn(org, { lat: 42.8846, lng: 74.5698 }, 20).ok, false);
    assert.equal(g.judgeCheckIn(org, { lat: 42.8780, lng: 74.5698 }, 100).ok, true); // 378 м, но погрешность 100 м
  });

  console.log("ИИ-помощник:");
  // Модель-заглушка: сначала просит инструменты, потом отвечает по их результату
  const scripted = (steps) => {
    const seen = [];
    let i = 0;
    const fn = async (body) => {
      seen.push(body);
      const st = steps[i++];
      if (typeof st === "function") return st(body);
      return st;
    };
    fn.seen = seen;
    return fn;
  };
  const toolUse = (id, name, input) => ({ stop_reason: "tool_use", content: [{ type: "tool_use", id, name, input }] });
  await t("куратор: ответ строится по данным своей группы", async () => {
    const model = scripted([
      toolUse("t1", "group_attendance", { group: "ит21", from: "2026-10-05", to: "2026-10-09" }),
      (body) => {
        const res = JSON.parse(body.messages.at(-1).content[0].content);
        const worst = res.students.sort((a, b) => b.unexcused - a.unexcused)[0];
        return { stop_reason: "end_turn", content: [{ type: "text", text: `Больше всех пропустил: ${worst.fio} (${worst.unexcused})` }] };
      },
    ]);
    const ai = createAssistant({ Group, Student, Curator, SummaryRecord, ensureGroupSummarySubjects: subjectsStub, getTodayKey, reports, callModel: model });
    const r = await ai.ask("500", { role: "curator", groupIds: [String(gA._id)] }, "кто больше всех пропустил на этой неделе?");
    assert.equal(r.text, "Больше всех пропустил: Бекова Айгуль (2)");
    assert.ok(model.seen[0].system.includes("куратор"));
    assert.ok(model.seen[0].tools.some((x) => x.name === "group_attendance"));
  });
  await t("куратор не видит чужую группу через ИИ", async () => {
    const ai = createAssistant({ Group, Student, Curator, SummaryRecord, ensureGroupSummarySubjects: subjectsStub, getTodayKey, reports, callModel: async () => ({}) });
    const ctx = { role: "curator", groupIds: [String(gA._id)] };
    const r = await ai.tools.group_attendance(ctx, { group: "СВ-11", from: "2026-10-05", to: "2026-10-09" });
    assert.ok(r.error);
    assert.deepEqual((await ai.tools.list_groups(ctx)).map((g) => g.name), ["ИТ-21"]);
    assert.equal((await ai.tools.find_student(ctx, { query: "Чужой" })).length, 0);
    assert.ok((await ai.tools.student_attendance(ctx, { student_id: String(s3._id), from: "2026-10-05", to: "2026-10-09" })).error);
    assert.ok((await ai.tools.send_excel_report(ctx, { group: "all", period: "week", date: "2026-10-07" }, [])).error);
  });
  await t("ИИ присылает Excel по запросу", async () => {
    const model = scripted([toolUse("t1", "send_excel_report", { group: "ИТ-21", period: "week", date: "2026-10-07" }), { stop_reason: "end_turn", content: [{ type: "text", text: "Файл ниже." }] }]);
    const ai = createAssistant({ Group, Student, Curator, SummaryRecord, ensureGroupSummarySubjects: subjectsStub, getTodayKey, reports, callModel: model });
    const r = await ai.ask("900", { role: "admin" }, "пришли excel по ит-21 за эту неделю");
    assert.equal(r.files.length, 1);
    assert.ok(r.files[0].filename.includes("ИТ-21"));
  });
  await t("без ключа ИИ честно говорит, что не настроен; родителю недоступен", async () => {
    delete process.env.ANTHROPIC_API_KEY;
    const off = createAssistant({ Group, Student, Curator, SummaryRecord, ensureGroupSummarySubjects: subjectsStub, getTodayKey, reports });
    assert.equal(off.enabled(), false);
    assert.match((await off.ask("1", { role: "admin" }, "?")).text, /ANTHROPIC_API_KEY/);
    const on = createAssistant({ Group, Student, Curator, SummaryRecord, ensureGroupSummarySubjects: subjectsStub, getTodayKey, reports, callModel: async () => ({}) });
    assert.match((await on.ask("1", { role: "parent" }, "?")).text, /администратору и кураторам/);
  });

  console.log(`\nИтог: ${passed} прошло, ${failed} упало`);
  srv.close();
  process.exit(failed ? 1 : 0);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
