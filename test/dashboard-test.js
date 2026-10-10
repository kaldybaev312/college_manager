// Тесты дашборда и «зоны риска» (dashboard.js). БД — в памяти.  Запуск: npm test
const Module = require("module");
const path = require("path");
const assert = require("assert/strict");

const fake = require("./fake-mongoose");
const origResolve = Module._resolveFilename;
Module._resolveFilename = function (req, ...rest) {
  return req === "mongoose" ? path.join(__dirname, "fake-mongoose.js") : origResolve.call(this, req, ...rest);
};
const express = require("express");
const { createReports } = require("../reports");
const { registerDashboard, riskLevel } = require("../dashboard");
const { datesBetween, isWeekend } = require("../attendance-utils");

const { Schema } = fake;
const Group = fake.model("Group", new Schema({ name: String, shiftStart: { type: String, default: "08:30" }, kind: { type: String, default: "students" }, saturday: { type: Boolean, default: false } }));
const Student = fake.model("Student", new Schema({ fio: String, group: Schema.Types.ObjectId, attendance: [], skudDaily: [] }));
fake.model("AbsenceReason", new Schema({ student: String, group: String, dateFrom: String, dateTo: String, type: String, excused: Boolean, confirmed: Boolean }, { timestamps: true }));

const TODAY = "2026-10-14"; // среда
const at = (dk, hhmm) => new Date(`${dk}T${hhmm}:00+06:00`).toISOString();
const daysOf = (list, hhmm = "08:10") => list.map((dk) => ({ date: dk, present: true, firstIn: at(dk, hhmm), inCount: 1 }));
const work = (from, to) => datesBetween(from, to).filter((d) => !isWeekend(d) && d <= TODAY);

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
  const gA = await Group.create({ name: "ИТ-21" });
  const gB = await Group.create({ name: "СВ-11" });
  await Group.create({ name: "Учителя", kind: "teachers" });
  // ИТ-21: Асанов ходит всегда (в первый понедельник опоздал), Бекова не ходит совсем
  const w1 = work("2026-10-05", "2026-10-11"), w2 = work("2026-10-12", "2026-10-18");
  await Student.create({ fio: "Асанов Бакыт", group: gA._id, skudDaily: [...daysOf(w1.slice(1)), ...daysOf(["2026-10-05"], "09:00"), ...daysOf(w2)] });
  await Student.create({ fio: "Бекова Айгуль", group: gA._id, skudDaily: [] });
  // СВ-11: в первую неделю был пн–ср, во вторую — каждый день
  await Student.create({ fio: "Сыдыков Эмир", group: gB._id, skudDaily: daysOf([...w1.slice(0, 3), ...w2]) });

  const reports = createReports({ Group, Student, Curator: null, TgSub: null, getTodayKey: () => TODAY });
  const app = express();
  const session = (req) => { try { return JSON.parse(req.headers["x-test-session"] || "null"); } catch { return null; } };
  registerDashboard(app, {
    Group,
    reports,
    getTodayKey: () => TODAY,
    getAccessibleGroupIds: (req) => { const s = session(req); return !s ? [] : s.role === "admin" ? null : s.groupIds; },
    requireAuth: (req, res, next) => (session(req) ? next() : res.status(401).end()),
  });
  const srv = app.listen(0);
  const base = `http://127.0.0.1:${srv.address().port}`;
  const ADMIN = { role: "admin" }, CUR = { role: "curator", groupIds: [String(gA._id)] };
  const get = async (s, url) => {
    const r = await fetch(base + url, { headers: s ? { "x-test-session": JSON.stringify(s) } : {} });
    return { status: r.status, data: await r.json().catch(() => null) };
  };

  console.log("Дашборд:");
  let d;
  await t("админ видит все учебные группы, без сотрудников", async () => {
    const r = await get(ADMIN, "/api/admin/dashboard?weeks=2");
    assert.equal(r.status, 200, JSON.stringify(r.data));
    d = r.data;
    assert.deepEqual(d.groups.map((g) => g.name), ["ИТ-21", "СВ-11"]);
    assert.equal(d.from, "2026-10-05");
    assert.equal(d.to, "2026-10-18");
  });
  await t("недельная посещаемость и показатели", async () => {
    // неделя 1: ИТ-21 5 из 10, СВ-11 3 из 5 → 8/15; неделя 2: ИТ-21 3 из 6, СВ-11 3 из 3 → 6/9
    assert.deepEqual(d.trend.map((w) => w.pct), [53, 67]);
    assert.equal(d.trend[0].late, 1);
    assert.equal(d.kpi.weekPct, 67);
    assert.equal(d.kpi.weekDelta, 14);
    assert.deepEqual([d.kpi.todayPresent, d.kpi.todayTotal], [2, 3]);
    assert.equal(d.kpi.students, 3);
  });
  await t("рейтинг групп от лучшей к худшей", async () => {
    assert.deepEqual(d.ranking.map((g) => [g.name, g.pct]), [["СВ-11", 75], ["ИТ-21", 50]]);
    assert.equal(d.ranking.find((g) => g.name === "ИТ-21").atRisk, 1);
  });
  await t("тепловая карта по дням недели", async () => {
    const it = d.heatmap.find((h) => h.name === "ИТ-21");
    assert.equal(it.weekdays[0].wd, "Пн");
    assert.equal(it.weekdays[0].pct, 50);
    const sv = d.heatmap.find((h) => h.name === "СВ-11");
    assert.equal(sv.weekdays[3].pct, 0); // чт первой недели пропущен, второй ещё не наступил
    assert.equal(d.showSaturday, false);
  });
  await t("зона риска: критичные сверху, с динамикой по неделям", async () => {
    assert.deepEqual(d.risk.map((r) => [r.fio, r.level]), [["Бекова Айгуль", "critical"], ["Сыдыков Эмир", "serious"]]);
    const b = d.risk[0];
    assert.deepEqual(b.weeks, [5, 3]);
    assert.equal(b.recent, 8);
    assert.equal(b.group, "ИТ-21");
    assert.equal(d.kpi.riskCount, 2);
    assert.equal(riskLevel(1, 3), "warning");
  });
  await t("куратор видит только свои группы", async () => {
    const r = await get(CUR, "/api/admin/dashboard?weeks=2");
    assert.deepEqual(r.data.groups.map((g) => g.name), ["ИТ-21"]);
    assert.ok(r.data.risk.every((x) => x.group === "ИТ-21"));
    assert.equal((await get(CUR, `/api/admin/dashboard?gid=${gB._id}`)).status, 403);
    assert.equal((await get({ role: "curator", groupIds: [] }, "/api/admin/dashboard")).status, 403);
  });
  await t("фильтр по группе и страница только после входа", async () => {
    const r = await get(ADMIN, `/api/admin/dashboard?weeks=2&gid=${gB._id}`);
    assert.deepEqual(r.data.ranking.map((g) => g.name), ["СВ-11"]);
    assert.equal(r.data.scope, String(gB._id));
    assert.equal((await fetch(base + "/dashboard")).status, 401);
  });

  console.log("Приложение (PWA):");
  await t("манифест, иконки и сервис-воркер", () => {
    const fs = require("fs");
    const pub = path.join(__dirname, "..", "public");
    const m = JSON.parse(fs.readFileSync(path.join(pub, "manifest.webmanifest"), "utf8"));
    assert.equal(m.display, "standalone");
    assert.equal(m.short_name, "ПЛ №3");
    for (const size of ["192x192", "512x512"]) assert.ok(m.icons.some((i) => i.sizes === size), size);
    for (const i of m.icons) assert.ok(fs.existsSync(path.join(pub, i.src)), i.src);
    const sw = fs.readFileSync(path.join(pub, "sw.js"), "utf8");
    assert.match(sw, /\/api\//); // API идёт мимо кэша
    for (const f of ["wake.html", "pwa.js", "icons/apple-touch-icon.png"]) assert.ok(fs.existsSync(path.join(pub, f)), f);
  });
  await t("манифест и заставка подключены на основных страницах", () => {
    const fs = require("fs");
    for (const p of ["index", "summary", "p", "practice", "journal", "miniapp", "dashboard"]) {
      const h = fs.readFileSync(path.join(__dirname, "..", "public", p + ".html"), "utf8");
      assert.ok(h.includes('rel="manifest"') && h.includes("apple-touch-icon") && h.includes('src="/pwa.js"') && h.includes('id="pl3-splash"'), p);
    }
  });

  console.log(`\nИтог: ${passed} прошло, ${failed} упало`);
  srv.close();
  process.exit(failed ? 1 : 0);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
