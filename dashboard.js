// Дашборд посещаемости и «зона риска».
// Все цифры берутся из reports.buildGroupReport (поле kinds по дням) — расчёты не дублируются.
// Доступ: админ — все группы, куратор — только свои (проверяется на сервере).
const path = require("path");
const { weekRange, addDays, parseKey, isValidKey } = require("./attendance-utils");

const WD_SHORT = ["Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"];
const MAX_WEEKS = 16;
const pct = (num, den) => (den ? Math.round((num / den) * 100) : null);
const daysDiff = (a, b) => Math.round((parseKey(b) - parseKey(a)) / 86400000);
const isStudentGroup = (g) => !g.kind || g.kind === "students";

// Уровень риска по пропускам без уважительной причины
function riskLevel(recent, total) {
  if (recent >= 4 || total >= 8) return "critical";
  if (recent >= 2 || total >= 5) return "serious";
  return "warning";
}

// Сборка дашборда из отчётов групп. reps — результаты buildGroupReport за период from..to.
function aggregate(reps, { from, weeks, today }) {
  const weekStarts = Array.from({ length: weeks }, (_, i) => addDays(from, i * 7));
  const blank = () => ({ slots: 0, present: 0, late: 0, excused: 0, unexcused: 0 });
  const add = (acc, k) => {
    if (!k) return;
    acc.slots++;
    if (k === "p" || k === "l") acc.present++;
    if (k === "l") acc.late++;
    if (k === "x") acc.excused++;
    if (k === "u") acc.unexcused++;
  };
  const overallWeeks = weekStarts.map(blank);
  const todayAcc = blank();
  const groups = [];
  const risk = [];

  for (const rep of reps) {
    const gWeeks = weekStarts.map(blank);
    const gWd = Array.from({ length: 7 }, blank);
    const dayWeek = rep.days.map((dk) => Math.floor(daysDiff(from, dk) / 7));
    const dayWd = rep.days.map((dk) => parseKey(dk).getUTCDay());
    const todayIdx = rep.days.indexOf(today);
    for (const row of rep.rows) {
      const sWeeks = new Array(weeks).fill(0);
      row.kinds.forEach((k, i) => {
        const w = dayWeek[i];
        if (w < 0 || w >= weeks) return;
        add(gWeeks[w], k);
        add(overallWeeks[w], k);
        add(gWd[dayWd[i]], k);
        if (k === "u") sWeeks[w]++;
      });
      if (todayIdx >= 0) add(todayAcc, row.kinds[todayIdx]);
      const total = sWeeks.reduce((a, b) => a + b, 0);
      const recent = sWeeks.slice(-2).reduce((a, b) => a + b, 0);
      const prior = sWeeks.slice(-4, -2).reduce((a, b) => a + b, 0);
      if (recent >= 2 || total >= 3)
        risk.push({ id: row.id, fio: row.fio, group: rep.group.name, groupId: rep.group.id, recent, prior, trend: recent - prior, total, weeks: sWeeks, pct: row.pct, level: riskLevel(recent, total) });
    }
    const last2 = gWeeks.slice(-2).reduce((a, w) => ({ slots: a.slots + w.slots, present: a.present + w.present }), { slots: 0, present: 0 });
    const prev2 = gWeeks.slice(-4, -2).reduce((a, w) => ({ slots: a.slots + w.slots, present: a.present + w.present }), { slots: 0, present: 0 });
    const all = gWeeks.reduce((a, w) => ({ slots: a.slots + w.slots, present: a.present + w.present, late: a.late + w.late, unexcused: a.unexcused + w.unexcused }), { slots: 0, present: 0, late: 0, unexcused: 0 });
    groups.push({
      id: rep.group.id,
      name: rep.group.name,
      students: rep.rows.length,
      pct: pct(all.present, all.slots),
      recentPct: pct(last2.present, last2.slots),
      delta: last2.slots && prev2.slots ? pct(last2.present, last2.slots) - pct(prev2.present, prev2.slots) : null,
      late: all.late,
      unexcused: all.unexcused,
      weeks: gWeeks.map((w) => pct(w.present, w.slots)),
      weekdays: [1, 2, 3, 4, 5, 6].map((d) => ({ wd: WD_SHORT[d], pct: pct(gWd[d].present, gWd[d].slots), slots: gWd[d].slots })),
      atRisk: 0,
    });
  }
  const byGroup = Object.fromEntries(groups.map((g) => [g.id, g]));
  for (const r of risk) byGroup[r.groupId].atRisk++;
  const order = { critical: 0, serious: 1, warning: 2 };
  risk.sort((a, b) => order[a.level] - order[b.level] || b.recent - a.recent || b.total - a.total || a.fio.localeCompare(b.fio, "ru"));

  const trend = overallWeeks.map((w, i) => ({ from: weekStarts[i], to: addDays(weekStarts[i], 6), pct: pct(w.present, w.slots), late: w.late, unexcused: w.unexcused, excused: w.excused, slots: w.slots }));
  const cur = trend[trend.length - 1], prev = trend[trend.length - 2];
  return {
    kpi: {
      weekPct: cur?.pct ?? null,
      weekDelta: cur?.pct != null && prev?.pct != null ? cur.pct - prev.pct : null,
      todayPresent: todayAcc.present,
      todayTotal: todayAcc.slots,
      weekLate: cur?.late || 0,
      riskCount: risk.length,
      riskCritical: risk.filter((r) => r.level === "critical").length,
      students: groups.reduce((a, g) => a + g.students, 0),
    },
    trend,
    ranking: [...groups].filter((g) => g.pct != null).sort((a, b) => b.pct - a.pct || a.name.localeCompare(b.name, "ru")),
    heatmap: groups.filter((g) => g.students).map((g) => ({ id: g.id, name: g.name, weekdays: g.weekdays })),
    showSaturday: groups.some((g) => g.weekdays[5].slots > 0),
    risk: risk.slice(0, 60),
  };
}

function registerDashboard(app, { Group, reports, getTodayKey, getAccessibleGroupIds, requireAuth }) {
  app.get("/dashboard", requireAuth, (_req, res) => res.sendFile(path.join(__dirname, "public", "dashboard.html")));

  app.get("/api/admin/dashboard", async (req, res) => {
    try {
      const ids = getAccessibleGroupIds(req); // null — все группы (админ)
      if (ids !== null && !ids.length) return res.status(403).json({ error: "Нет доступных групп" });
      let groups = (await Group.find().sort({ name: 1 }).lean()).filter(isStudentGroup);
      if (ids !== null) groups = groups.filter((g) => ids.includes(String(g._id)));
      const all = groups.map((g) => ({ id: String(g._id), name: g.name }));
      const gid = req.query.gid && req.query.gid !== "all" ? String(req.query.gid) : null;
      if (gid) {
        groups = groups.filter((g) => String(g._id) === gid);
        if (!groups.length) return res.status(403).json({ error: "Нет доступа к группе" });
      }
      const weeks = Math.min(Math.max(parseInt(req.query.weeks) || 8, 2), MAX_WEEKS);
      const today = isValidKey(req.query.today) && req.query.today <= getTodayKey() ? req.query.today : getTodayKey();
      const to = weekRange(today).to;
      const from = addDays(weekRange(today).from, -7 * (weeks - 1));
      const reps = [];
      for (const g of groups) reps.push(await reports.buildGroupReport(g, { period: "custom", from, to, title: "" }));
      res.json({ today, from, to, weeks, groups: all, scope: gid || "all", ...aggregate(reps, { from, weeks, today }) });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
}

module.exports = { registerDashboard, aggregate, riskLevel };
