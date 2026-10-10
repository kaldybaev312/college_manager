// API мини-приложения Telegram — /api/webapp/*
// Пользователь определяется ТОЛЬКО по подписанному Telegram initData (заголовок x-tg-init-data).
// Подпись проверяется HMAC-ом с токеном бота: подделать chat ID или чужую роль нельзя.

const crypto = require("crypto");
const path = require("path");
const { pad2, shiftMinutes, dayStatus, isWeekend, isValidKey, monthRange } = require("./attendance-utils");
const absence = require("./absence");

const INIT_DATA_MAX_AGE = 7 * 24 * 3600; // сек

// Проверка initData по https://core.telegram.org/bots/webapps#validating-data-received-via-the-mini-app
function verifyInitData(initData, botToken, nowSec = Math.floor(Date.now() / 1000)) {
  if (!initData || !botToken) return null;
  let params;
  try {
    params = new URLSearchParams(String(initData));
  } catch {
    return null;
  }
  const hash = params.get("hash");
  if (!hash || !/^[0-9a-f]{64}$/i.test(hash)) return null;
  params.delete("hash");
  const dataCheck = [...params.entries()]
    .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
    .map(([k, v]) => `${k}=${v}`)
    .join("\n");
  const secret = crypto.createHmac("sha256", "WebAppData").update(botToken).digest();
  const calc = crypto.createHmac("sha256", secret).update(dataCheck).digest();
  if (!crypto.timingSafeEqual(calc, Buffer.from(hash, "hex"))) return null;
  const authDate = Number(params.get("auth_date") || 0);
  if (!authDate || nowSec - authDate > INIT_DATA_MAX_AGE) return null;
  try {
    const user = JSON.parse(params.get("user") || "null");
    if (!user?.id) return null;
    return { user, authDate };
  } catch {
    return null;
  }
}

const isMonth = (m) => /^\d{4}-\d{2}$/.test(String(m || ""));

module.exports = function registerWebapp(app, deps) {
  const { Group, Student, TgSub, Curator, SummaryRecord, ensureGroupSummarySubjects, formatMonthLabel, notifyParents, getTodayKey, getBot, mdv2 } = deps;
  const BOT_TOKEN = process.env.TG_TOKEN;

  // ── Аутентификация ──
  async function tgAuth(req, res, next) {
    const v = verifyInitData(req.headers["x-tg-init-data"], BOT_TOKEN);
    if (!v) return res.status(401).json({ error: "Откройте приложение через Telegram-бота" });
    const chatId = String(v.user.id);
    req.tg = { chatId, user: v.user, sub: await TgSub.findOne({ chatId }) };
    next();
  }
  const wrap = (fn) => (req, res) =>
    Promise.resolve(fn(req, res)).catch((e) => res.status(e.status || 500).json({ error: e.message }));
  const fail = (status, msg) => Object.assign(new Error(msg), { status });

  // Группы, доступные пользователю: null = все (админ)
  async function groupIdsOf(sub) {
    if (sub?.role === "admin") return null;
    if (sub?.role === "curator" && sub.curatorId) {
      const cur = await Curator.findById(sub.curatorId).select("groups");
      return (cur?.groups || []).map(String);
    }
    return [];
  }
  async function needGroup(req, gid) {
    const ids = await groupIdsOf(req.tg.sub);
    if (!gid || (ids !== null && !ids.includes(String(gid)))) throw fail(403, "Нет доступа к группе");
  }
  const needStaff = (req) => {
    if (!["admin", "curator"].includes(req.tg.sub?.role)) throw fail(403, "Только для куратора или администратора");
  };
  // Родитель видит только привязанных детей; админ — любого
  async function needChild(req, inn) {
    const sub = req.tg.sub;
    const clean = String(inn || "").replace(/\D/g, "");
    if (!clean) throw fail(400, "Не указан ИНН");
    if (sub?.role !== "admin" && !(sub?.role === "parent" && (sub.linkedInns || []).includes(clean)))
      throw fail(403, "Нет доступа");
    const stu = await Student.findOne({ inn: clean }).populate("group", "name shiftStart");
    if (!stu) throw fail(404, "Студент не найден");
    return stu;
  }

  const groupBrief = (g) => ({ id: String(g._id), name: g.name, profRu: g.profRu || "", shiftStart: g.shiftStart || "09:30" });

  app.get("/api/webapp/open", (_req, res) => res.sendFile(path.join(__dirname, "public", "miniapp.html")));

  app.use("/api/webapp", (req, res, next) => tgAuth(req, res, next).catch((e) => res.status(500).json({ error: e.message })));

  // ── Кто я ──
  app.get(
    "/api/webapp/me",
    wrap(async (req, res) => {
      const sub = req.tg.sub;
      if (sub?.role === "admin") {
        const groups = await Group.find().sort({ name: 1 }).select("name profRu shiftStart");
        return res.json({ role: "admin", absenceTypes: absence.REASON_TYPES, groups: groups.map(groupBrief) });
      }
      if (sub?.role === "curator" && sub.curatorId) {
        const cur = await Curator.findById(sub.curatorId).populate("groups", "name profRu shiftStart");
        if (!cur) return res.json({ role: null });
        return res.json({ role: "curator", absenceTypes: absence.REASON_TYPES, curatorName: cur.fullName || cur.username || "", groups: cur.groups.map(groupBrief) });
      }
      if (sub?.role === "parent" && sub.linkedInns?.length) {
        const children = await Promise.all(
          sub.linkedInns.map(async (inn) => {
            const stu = await Student.findOne({ inn }).populate("group", "name profRu");
            return stu && { inn, fio: stu.fio || "", groupId: String(stu.group?._id || ""), groupName: stu.group?.name || "", profRu: stu.group?.profRu || "" };
          }),
        );
        return res.json({ role: "parent", absenceTypes: absence.REASON_TYPES, children: children.filter(Boolean) });
      }
      res.json({ role: null });
    }),
  );

  // ── Посещаемость группы сегодня ──
  app.get(
    "/api/webapp/attendance/today",
    wrap(async (req, res) => {
      const gid = req.query.gid;
      await needGroup(req, gid);
      const tdk = getTodayKey();
      const g = await Group.findById(gid);
      const shiftMin = shiftMinutes(g?.shiftStart);
      const students = await Student.find({ group: gid }).sort({ fio: 1 }).select("fio skudDaily attendance");
      const reasons = await absence.reasonsIndex({ group: gid }, tdk, tdk);
      let present = 0,
        absent = 0,
        late = 0;
      const list = students.map((stu) => {
        const st = dayStatus(stu, tdk, shiftMin);
        if (st.present) present++;
        else absent++;
        if (st.late) late++;
        return {
          id: String(stu._id),
          fio: stu.fio,
          status: st.present ? (st.late ? "late" : "present") : "absent",
          time: st.time,
          reason: st.present ? null : reasons.get(`${stu._id}|${tdk}`) || null,
        };
      });
      res.json({ date: tdk, present, absent, late, students: list });
    }),
  );

  // statuses: p — был, l — опоздал, a — нет без причины, e — нет по причине (код в reasons)
  function monthStatuses(stu, days, shiftMin, reasons) {
    const statuses = {},
      codes = {};
    for (const day of days) {
      const st = dayStatus(stu, day, shiftMin);
      const r = !st.present && reasons.get(`${stu._id}|${day}`);
      statuses[day] = st.present ? (st.late ? "l" : "p") : r ? "e" : "a";
      if (r) codes[day] = { code: r.code + (r.confirmed ? "" : "?"), label: r.label, confirmed: r.confirmed, excused: r.excused, comment: r.comment, id: r.id };
    }
    return { statuses, codes };
  }
  const monthDays = (month) => {
    const { from, to } = monthRange(month);
    const days = [];
    for (let d = 1; d <= Number(to.slice(8)); d++) days.push(`${month}-${pad2(d)}`);
    return { from, to, days };
  };

  app.get(
    "/api/webapp/attendance/month",
    wrap(async (req, res) => {
      const { gid, month } = req.query;
      await needGroup(req, gid);
      if (!isMonth(month)) throw fail(400, "Invalid month");
      const { from, to, days } = monthDays(month);
      const g = await Group.findById(gid);
      const shiftMin = shiftMinutes(g?.shiftStart);
      const students = await Student.find({ group: gid }).sort({ fio: 1 }).select("fio skudDaily attendance");
      const reasons = await absence.reasonsIndex({ group: gid }, from, to);
      res.json({
        days,
        students: students.map((stu) => {
          const { statuses, codes } = monthStatuses(stu, days, shiftMin, reasons);
          return { id: String(stu._id), fio: stu.fio, statuses, reasons: codes };
        }),
      });
    }),
  );

  // ── Предметы, студенты, оценки ──
  app.get(
    "/api/webapp/subjects",
    wrap(async (req, res) => {
      await needGroup(req, req.query.gid);
      const g = await Group.findById(req.query.gid);
      if (!g) throw fail(404, "Group not found");
      res.json({ subjects: await ensureGroupSummarySubjects(g) });
    }),
  );

  app.get(
    "/api/webapp/students",
    wrap(async (req, res) => {
      await needGroup(req, req.query.gid);
      const stus = await Student.find({ group: req.query.gid }).sort({ fio: 1 }).select("fio inn lyceumId birthDate");
      res.json({
        students: stus.map((s) => ({ id: String(s._id), fio: s.fio, inn: s.inn || "", lyceumId: s.lyceumId || "", birthDate: s.birthDate || "" })),
      });
    }),
  );

  const gradesToObj = (raw) => {
    const out = {};
    if (raw instanceof Map) raw.forEach((v, k) => v && (out[k] = v));
    else if (raw) Object.entries(raw).forEach(([k, v]) => v && !k.startsWith("$") && (out[k] = v));
    return out;
  };

  app.get(
    "/api/webapp/grades",
    wrap(async (req, res) => {
      const { gid, month } = req.query;
      await needGroup(req, gid);
      const g = await Group.findById(gid);
      const subjects = await ensureGroupSummarySubjects(g);
      const records = await SummaryRecord.find({ group: gid, month }).lean();
      const grades = {};
      records.forEach((r) => (grades[String(r.student)] = gradesToObj(r.grades)));
      res.json({ grades, subjects });
    }),
  );

  app.post(
    "/api/webapp/grades/set",
    wrap(async (req, res) => {
      needStaff(req);
      const { gid, studentId, month, grades } = req.body || {};
      await needGroup(req, gid);
      if (!isMonth(month)) throw fail(400, "Invalid month");
      const stu = await Student.findOne({ _id: studentId, group: gid }).select("fio inn");
      if (!stu) throw fail(404, "Студент не найден в этой группе");
      const g = await Group.findById(gid);
      const keys = new Set((await ensureGroupSummarySubjects(g)).map((s) => s.key));
      const clean = {};
      for (const [k, v] of Object.entries(grades || {})) if (keys.has(k)) clean[k] = String(v ?? "").slice(0, 8);
      await SummaryRecord.findOneAndUpdate(
        { group: gid, student: studentId, month },
        { group: gid, student: studentId, month, grades: clean },
        { upsert: true, new: true },
      );
      if (stu.inn && Object.values(clean).some(Boolean))
        notifyParents(stu.inn, `📊 *${mdv2("Новые оценки!")}*\n👤 ${mdv2(stu.fio)}\n📅 ${mdv2(formatMonthLabel(month))}`).catch(() => {});
      res.json({ ok: true });
    }),
  );

  // ── Родитель ──
  const linkFails = new Map(); // chatId → [время неудачных попыток]
  const failsOf = (id) => (linkFails.get(id) || []).filter((t) => Date.now() - t < 3600 * 1000);

  app.post(
    "/api/webapp/parent/link",
    wrap(async (req, res) => {
      const chatId = req.tg.chatId;
      if (["admin", "curator"].includes(req.tg.sub?.role))
        throw fail(400, "Этот Telegram уже вошёл как куратор или администратор");
      if (failsOf(chatId).length >= 5) throw fail(429, "Слишком много попыток. Попробуйте через час");
      const clean = String(req.body?.inn || "").replace(/\D/g, "");
      if (clean.length < 12) throw fail(400, "ИНН минимум 12 цифр");
      const stu = await Student.findOne({ inn: clean }).populate("group");
      if (!stu) {
        linkFails.set(chatId, [...failsOf(chatId), Date.now()]);
        throw fail(404, "Студент с таким ИНН не найден");
      }
      linkFails.delete(chatId);
      const name = [req.tg.user.first_name, req.tg.user.last_name].filter(Boolean).join(" ");
      await TgSub.findOneAndUpdate(
        { chatId },
        { $set: { chatId, role: "parent", ...(name ? { name } : {}) }, $addToSet: { linkedInns: clean } },
        { upsert: true, new: true },
      );
      res.json({ ok: true, fio: stu.fio, groupName: stu.group?.name || "", groupId: String(stu.group?._id || "") });
    }),
  );

  app.get(
    "/api/webapp/parent/grades",
    wrap(async (req, res) => {
      const stu = await needChild(req, req.query.inn);
      const g = await Group.findById(stu.group?._id || stu.group);
      const subjects = await ensureGroupSummarySubjects(g);
      const record = await SummaryRecord.findOne({ group: g?._id, student: stu._id, month: req.query.month }).lean();
      res.json({ grades: gradesToObj(record?.grades), subjects });
    }),
  );

  app.get(
    "/api/webapp/parent/attendance",
    wrap(async (req, res) => {
      const stu = await needChild(req, req.query.inn);
      const month = req.query.month;
      if (!isMonth(month)) throw fail(400, "Invalid month");
      const { from, to, days } = monthDays(month);
      const shiftMin = shiftMinutes(stu.group?.shiftStart);
      const reasons = await absence.reasonsIndex({ student: stu._id }, from, to);
      const attMap = {};
      for (const day of days) {
        const st = dayStatus(stu, day, shiftMin);
        const r = !st.present && reasons.get(`${stu._id}|${day}`);
        attMap[day] = {
          status: st.present ? (st.late ? "l" : "p") : r ? "e" : "a",
          time: st.time || null,
          reason: r ? { label: r.label, code: r.code, confirmed: r.confirmed, comment: r.comment } : null,
        };
      }
      res.json({ fio: stu.fio, attMap });
    }),
  );

  // ── Причины отсутствия ──
  async function notifyCuratorsPlain(groupId, text) {
    const bot = getBot();
    if (!bot) return;
    const curs = await Curator.find({ groups: groupId }).select("_id").lean();
    const subs = await TgSub.find({ role: "curator", curatorId: { $in: curs.map((c) => String(c._id)) } }).lean();
    for (const s of subs) bot.sendMessage(s.chatId, text).catch(() => null);
  }

  app.get(
    "/api/webapp/absences",
    wrap(async (req, res) => {
      needStaff(req);
      const { gid, from, to } = req.query;
      await needGroup(req, gid);
      if (!isValidKey(from) || !isValidKey(to)) throw fail(400, "Неверный период");
      res.json((await absence.findReasons({ group: gid }, from, to)).map(absence.serialize));
    }),
  );

  app.post(
    "/api/webapp/absences",
    wrap(async (req, res) => {
      const sub = req.tg.sub;
      const body = req.body || {};
      if (sub?.role === "parent") {
        const stu = await needChild(req, body.inn);
        const r = await absence.createReason({ student: stu, body, source: "parent", createdBy: sub.name || req.tg.chatId });
        const period = r.dateFrom === r.dateTo ? r.dateFrom : `${r.dateFrom} — ${r.dateTo}`;
        notifyCuratorsPlain(
          stu.group?._id || stu.group,
          `📝 Родитель сообщил об отсутствии\n👤 ${stu.fio} (${stu.group?.name || ""})\n📅 ${period}\n📌 ${r.label}${r.comment ? "\n💬 " + r.comment : ""}\n\nПодтвердите в мини-приложении: Посещаемость → Сегодня.`,
        ).catch(() => {});
        return res.json(r);
      }
      needStaff(req);
      const stu = await Student.findById(body.studentId).select("group fio");
      if (!stu) throw fail(404, "Студент не найден");
      await needGroup(req, stu.group);
      res.json(await absence.createReason({ student: stu, body, source: sub.role, createdBy: sub.name || req.tg.chatId }));
    }),
  );

  async function staffReason(req) {
    needStaff(req);
    const r = await absence.AbsenceReason.findById(req.params.id).lean();
    if (!r) throw fail(404, "Запись не найдена");
    await needGroup(req, r.group);
    return r;
  }

  app.put(
    "/api/webapp/absences/:id",
    wrap(async (req, res) => {
      await staffReason(req);
      res.json(await absence.updateReason(req.params.id, { ...req.body, confirmed: true }));
    }),
  );

  app.delete(
    "/api/webapp/absences/:id",
    wrap(async (req, res) => {
      await staffReason(req);
      await absence.AbsenceReason.deleteOne({ _id: req.params.id });
      res.json({ ok: true });
    }),
  );

  // ── Статистика сегодня (админ) ──
  app.get(
    "/api/webapp/stats",
    wrap(async (req, res) => {
      if (req.tg.sub?.role !== "admin") throw fail(403, "Admin only");
      const tdk = getTodayKey();
      const groups = (await Group.find().sort({ name: 1 })).filter((g) => !g.kind || g.kind === "students");
      const totals = { total: 0, present: 0, absent: 0, late: 0 };
      const out = [];
      for (const g of groups) {
        const shiftMin = shiftMinutes(g.shiftStart);
        const students = await Student.find({ group: g._id }).select("skudDaily attendance");
        let present = 0,
          late = 0;
        for (const stu of students) {
          const st = dayStatus(stu, tdk, shiftMin);
          if (st.present) present++;
          if (st.late) late++;
        }
        const absent = isWeekend(tdk) ? 0 : students.length - present;
        totals.total += students.length;
        totals.present += present;
        totals.absent += absent;
        totals.late += late;
        out.push({ id: String(g._id), name: g.name, present, absent, late, total: students.length });
      }
      res.json({ ...totals, groups: out });
    }),
  );
};

module.exports.verifyInitData = verifyInitData;
