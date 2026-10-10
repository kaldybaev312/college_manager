// Еженедельные и ежемесячные отчёты по посещаемости групп.
// Excel по одной группе, общий Excel по всем группам и авторассылка в Telegram:
//   понедельник 08:00 — за прошлую неделю, 1-е число 08:10 — за прошлый месяц (Бишкек).
// Отключить авторассылку: REPORTS_CRON=off

const xlsx = require("xlsx");
const cron = require("node-cron");
const {
  pad2,
  shiftMinutes,
  dayStatus,
  isWeekend,
  isValidKey,
  addDays,
  datesBetween,
  weekRange,
  monthRange,
} = require("./attendance-utils");
const { reasonsIndex } = require("./absence");

const RU_WD = ["Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"];
const RU_MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"];
const ATTENTION_LIMIT = 3; // пропусков без уважительной причины за период
const XLSX_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

const fmtDay = (dk) => `${dk.slice(8, 10)}.${dk.slice(5, 7)}`;
const isStudentGroup = (g) => !g.kind || g.kind === "students";

// Период отчёта: week (dk — любой день недели) или month (dk — любой день месяца или "ГГГГ-ММ")
function resolvePeriod(period, dk) {
  if (period === "month") {
    const month = String(dk || "").slice(0, 7);
    if (!/^\d{4}-\d{2}$/.test(month)) throw new Error("Неверный месяц");
    const [y, m] = month.split("-").map(Number);
    return { period, ...monthRange(month), title: `${RU_MONTHS[m - 1]} ${y}` };
  }
  if (!isValidKey(dk)) throw new Error("Неверная дата");
  const r = weekRange(dk);
  return { period: "week", ...r, title: `${fmtDay(r.from)}–${fmtDay(r.to)}.${r.to.slice(0, 4)}` };
}

// Прошедший период относительно сегодняшнего дня (для авторассылки)
function previousPeriod(period, todayKey) {
  if (period === "month") {
    const [y, m] = todayKey.split("-").map(Number);
    const py = m === 1 ? y - 1 : y,
      pm = m === 1 ? 12 : m - 1;
    return resolvePeriod("month", `${py}-${pad2(pm)}`);
  }
  return resolvePeriod("week", addDays(todayKey, -7));
}

function createReports({ Group, Student, Curator, TgSub, getTodayKey }) {
  // Данные отчёта по группе за период
  async function buildGroupReport(group, range) {
    const today = getTodayKey();
    const days = datesBetween(range.from, range.to);
    const workdays = days.filter((dk) => !isWeekend(dk) && dk <= today);
    const shiftMin = shiftMinutes(group.shiftStart);
    const students = await Student.find({ group: group._id }).sort({ fio: 1 }).select("fio skudDaily attendance").lean();
    const reasons = await reasonsIndex({ group: group._id }, range.from, range.to);

    const rows = students.map((stu) => {
      const sid = String(stu._id);
      const cells = [];
      let present = 0,
        late = 0,
        excused = 0,
        unexcused = 0,
        pending = 0;
      for (const dk of days) {
        const st = dayStatus(stu, dk, shiftMin);
        const reason = reasons.get(`${sid}|${dk}`);
        if (st.present) {
          present++;
          if (st.late) late++;
          cells.push(st.late ? "О" : "П");
        } else if (isWeekend(dk) || dk > today) {
          cells.push("");
        } else if (reason) {
          if (!reason.confirmed) pending++;
          if (reason.excused && reason.confirmed) excused++;
          else unexcused++;
          cells.push(reason.code + (reason.confirmed ? "" : "?"));
        } else {
          unexcused++;
          cells.push("Н");
        }
      }
      const pct = workdays.length ? Math.round((Math.min(present, workdays.length) / workdays.length) * 100) : 0;
      return { id: sid, fio: stu.fio, cells, present, late, excused, unexcused, pending, pct };
    });

    const total = rows.length;
    const sum = (k) => rows.reduce((a, r) => a + r[k], 0);
    return {
      group: { id: String(group._id), name: group.name, profRu: group.profRu || "", shiftStart: group.shiftStart || "" },
      range,
      days,
      workdays: workdays.length,
      rows,
      totals: {
        students: total,
        present: sum("present"),
        late: sum("late"),
        excused: sum("excused"),
        unexcused: sum("unexcused"),
        pending: sum("pending"),
        pct: total ? Math.round(rows.reduce((a, r) => a + r.pct, 0) / total) : 0,
      },
      attention: rows.filter((r) => r.unexcused >= ATTENTION_LIMIT).sort((a, b) => b.unexcused - a.unexcused),
    };
  }

  function groupSheet(rep) {
    const kind = rep.range.period === "month" ? "Месячный" : "Недельный";
    const head = ["№", "ФИО", ...rep.days.map((dk) => `${fmtDay(dk)} ${RU_WD[new Date(dk + "T00:00:00Z").getUTCDay()]}`), "Был", "Опозд.", "Уваж.", "Без уваж.", "% посещ."];
    const aoa = [
      [`${kind} отчёт по посещаемости — ПЛ №3`],
      [`Группа: ${rep.group.name}${rep.group.profRu ? " · " + rep.group.profRu : ""}`],
      [`Период: ${rep.range.title} · учебных дней: ${rep.workdays} · начало занятий: ${rep.group.shiftStart || "—"}`],
      [],
      head,
      ...rep.rows.map((r, i) => [i + 1, r.fio, ...r.cells, r.present, r.late, r.excused, r.unexcused, r.pct / 100]),
      [],
      ["", "Итого по группе", ...rep.days.map(() => ""), rep.totals.present, rep.totals.late, rep.totals.excused, rep.totals.unexcused, rep.totals.pct / 100],
      [],
      ["", "Обозначения: П — был, О — опоздал, Н — пропуск без причины, Б — болезнь, С — семейные обст., Л — по направлению лицея, Д — другая причина, ? — ждёт подтверждения куратора"],
    ];
    if (rep.attention.length) {
      aoa.push([], ["", `Требуют внимания (${ATTENTION_LIMIT}+ пропусков без уважительной причины):`]);
      rep.attention.forEach((r) => aoa.push(["", `${r.fio} — ${r.unexcused}`]));
    }
    const ws = xlsx.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 4 }, { wch: 32 }, ...rep.days.map(() => ({ wch: 7 })), { wch: 6 }, { wch: 7 }, { wch: 7 }, { wch: 9 }, { wch: 9 }];
    const pctCol = 2 + rep.days.length + 4;
    for (let r = 5; r < aoa.length; r++) {
      const ref = xlsx.utils.encode_cell({ r, c: pctCol });
      if (ws[ref] && typeof ws[ref].v === "number") ws[ref].z = "0%";
    }
    return ws;
  }

  const sheetName = (name, used) => {
    let base = String(name || "Группа").replace(/[\[\]:*?/\\]/g, "_").slice(0, 28) || "Группа";
    let n = base,
      i = 2;
    while (used.has(n)) n = `${base.slice(0, 26)}_${i++}`;
    used.add(n);
    return n;
  };

  function groupWorkbook(rep) {
    const wb = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(wb, groupSheet(rep), "Посещаемость");
    return xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
  }

  function allWorkbook(reps, range) {
    const wb = xlsx.utils.book_new();
    const aoa = [
      [`${range.period === "month" ? "Месячный" : "Недельный"} отчёт по посещаемости — ПЛ №3`],
      [`Период: ${range.title}`],
      [],
      ["Группа", "Студентов", "Учебных дней", "% посещ.", "Опозданий", "Пропусков уваж.", "Пропусков без уваж.", "Ждут подтверждения", "Требуют внимания"],
      ...reps.map((r) => [r.group.name, r.totals.students, r.workdays, r.totals.pct / 100, r.totals.late, r.totals.excused, r.totals.unexcused, r.totals.pending, r.attention.length]),
    ];
    const ws = xlsx.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 18 }, { wch: 10 }, { wch: 13 }, { wch: 9 }, { wch: 10 }, { wch: 15 }, { wch: 18 }, { wch: 18 }, { wch: 16 }];
    for (let r = 4; r < aoa.length; r++) {
      const ref = xlsx.utils.encode_cell({ r, c: 3 });
      if (ws[ref]) ws[ref].z = "0%";
    }
    const used = new Set(["Итог"]);
    xlsx.utils.book_append_sheet(wb, ws, "Итог");
    reps.forEach((rep) => xlsx.utils.book_append_sheet(wb, groupSheet(rep), sheetName(rep.group.name, used)));
    return xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
  }

  async function studentGroups() {
    return (await Group.find().sort({ name: 1 }).lean()).filter(isStudentGroup);
  }

  function summaryText(rep) {
    const kind = rep.range.period === "month" ? "Месячный" : "Недельный";
    const t = rep.totals;
    let s = `📊 ${kind} отчёт · ${rep.group.name}\n📅 ${rep.range.title}\n\n` +
      `✅ Посещаемость: ${t.pct}%\n⏰ Опозданий: ${t.late}\n🩺 Пропусков по уважительной: ${t.excused}\n❌ Без уважительной: ${t.unexcused}`;
    if (t.pending) s += `\n❓ Причин ждут подтверждения: ${t.pending}`;
    if (rep.attention.length)
      s += `\n\n⚠️ Требуют внимания:\n` + rep.attention.slice(0, 10).map((r) => `• ${r.fio} — ${r.unexcused}`).join("\n");
    return s;
  }

  const fileName = (name, range) =>
    `Otchet_${String(name).replace(/[^a-zA-Z0-9а-яА-ЯёЁ._-]/g, "_")}_${range.from}_${range.to}.xlsx`;

  // Рассылка: кураторам — по их группам, админам — общий файл
  async function sendReports(bot, period) {
    const range = previousPeriod(period, getTodayKey());
    const groups = await studentGroups();
    const reps = [];
    for (const g of groups) reps.push(await buildGroupReport(g, range));

    let sent = 0;
    for (const rep of reps) {
      if (!rep.rows.length) continue;
      const curators = await Curator.find({ groups: rep.group.id }).select("_id").lean();
      if (!curators.length) continue;
      const subs = await TgSub.find({ role: "curator", curatorId: { $in: curators.map((c) => String(c._id)) } }).lean();
      const buf = groupWorkbook(rep);
      for (const s of subs) {
        await bot
          .sendDocument(s.chatId, buf, { caption: summaryText(rep).slice(0, 1000) }, { filename: fileName(rep.group.name, range), contentType: XLSX_TYPE })
          .then(() => sent++)
          .catch((e) => console.error("report → curator:", e.message));
      }
    }

    const admins = await TgSub.find({ role: "admin" }).lean();
    if (admins.length) {
      const withStudents = reps.filter((r) => r.rows.length);
      const avg = withStudents.length ? Math.round(withStudents.reduce((a, r) => a + r.totals.pct, 0) / withStudents.length) : 0;
      const worst = [...withStudents].sort((a, b) => a.totals.pct - b.totals.pct).slice(0, 5);
      const caption =
        `📊 ${period === "month" ? "Месячный" : "Недельный"} отчёт по лицею\n📅 ${range.title}\n\n` +
        `Средняя посещаемость: ${avg}%\nГрупп: ${withStudents.length}\n` +
        (worst.length ? `\nНиже всего:\n` + worst.map((r) => `• ${r.group.name} — ${r.totals.pct}%`).join("\n") : "");
      const buf = allWorkbook(withStudents, range);
      for (const a of admins) {
        await bot
          .sendDocument(a.chatId, buf, { caption: caption.slice(0, 1000) }, { filename: fileName("Vse_gruppy", range), contentType: XLSX_TYPE })
          .then(() => sent++)
          .catch((e) => console.error("report → admin:", e.message));
      }
    }
    console.log(`📊 Отчёт (${period}, ${range.from}..${range.to}) отправлен: ${sent}`);
    return { range, sent };
  }

  function scheduleReports(bot) {
    if (String(process.env.REPORTS_CRON || "").toLowerCase() === "off") return;
    const opts = { timezone: "Asia/Bishkek" };
    cron.schedule("0 8 * * 1", () => sendReports(bot, "week").catch((e) => console.error("weekly report:", e)), opts);
    cron.schedule("10 8 1 * *", () => sendReports(bot, "month").catch((e) => console.error("monthly report:", e)), opts);
    console.log("📊 Авторассылка отчётов: пн 08:00 (неделя), 1-е число 08:10 (месяц)");
  }

  function registerReportRoutes(app, { requireAdmin, canAccessGroup, contentDisposition, getBot }) {
    const wrap = (fn) => (req, res) => fn(req, res).catch((e) => res.status(400).json({ error: e.message }));
    const rangeOf = (req) => resolvePeriod(req.query.period === "month" ? "month" : "week", req.query.date || getTodayKey());
    const sendXlsx = (res, name, buf) => {
      res.setHeader("Content-Disposition", contentDisposition(name));
      res.setHeader("Content-Type", XLSX_TYPE);
      res.send(buf);
    };

    app.get(
      "/api/admin/report/group/:gid",
      wrap(async (req, res) => {
        if (!canAccessGroup(req, req.params.gid)) return res.status(403).json({ error: "Нет доступа к группе" });
        const g = await Group.findById(req.params.gid).lean();
        if (!g) return res.status(404).json({ error: "Группа не найдена" });
        const rep = await buildGroupReport(g, rangeOf(req));
        if (req.query.format === "json") return res.json(rep);
        sendXlsx(res, fileName(g.name, rep.range), groupWorkbook(rep));
      }),
    );

    app.get(
      "/api/admin/report/all",
      requireAdmin,
      wrap(async (req, res) => {
        const range = rangeOf(req);
        const reps = [];
        for (const g of await studentGroups()) reps.push(await buildGroupReport(g, range));
        sendXlsx(res, fileName("Vse_gruppy", range), allWorkbook(reps.filter((r) => r.rows.length), range));
      }),
    );

    // Разослать прямо сейчас (за прошлую неделю/месяц) — для проверки
    app.post(
      "/api/admin/report/send-now",
      requireAdmin,
      wrap(async (req, res) => {
        const bot = getBot();
        if (!bot) return res.status(503).json({ error: "Бот не запущен" });
        res.json(await sendReports(bot, req.body?.period === "month" ? "month" : "week"));
      }),
    );
  }

  return { buildGroupReport, groupWorkbook, allWorkbook, sendReports, scheduleReports, registerReportRoutes, resolvePeriod, previousPeriod, summaryText };
}

module.exports = { createReports, resolvePeriod, previousPeriod };
