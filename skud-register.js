// Регистрация студентов в Hikvision СКУД: серверная часть (очередь заданий для локального агента).
// Подключается через skud.js. Администратор видит все группы, куратор только свои.
const multer = require("multer");
const mongoose = require("mongoose");
const xlsx = require("xlsx");

module.exports = function mountRegister(app, ctx, core) {
  const { Student, Group, normalizeINN, getSession, canAccessGroup, getAccessibleGroupIds, requireAdmin } = ctx;
  const { SkudJob, SkudAudit, buildDeviceName, jobStates, checkPhoto, endOfDayBishkek, STATUS_RU, escRe, audit, actorOf } = core;

  const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 1e6, files: 1 } });
  const uploadPhoto = (req, res, next) =>
    upload.single("photo")(req, res, (err) => (err ? res.status(400).json({ error: "Файл слишком большой или повреждён" }) : next()));

  /* ------------------------------ доступ ------------------------------ */

  const roleOf = (req) => (getSession(req) || {}).role;
  const isAdmin = (req) => roleOf(req) === "admin";
  const requireStaffRole = (req, res, next) =>
    ["admin", "curator"].includes(roleOf(req)) ? next() : res.status(403).json({ error: "Нет доступа" });
  const validId = (id) => mongoose.isValidObjectId(String(id || ""));
  const fail = (res, e) => res.status(e.status || 500).json({ error: e.status ? e.message : "Ошибка сервера: " + e.message });
  const httpErr = (status, message) => Object.assign(new Error(message), { status });

  const isReady = (s) => !!normalizeINN(s.inn) && /^\d{1,32}$/.test(String(s.lyceumId || "").trim());

  /* --------------------- удаление с терминалов ----------------------- */

  // Поставить задание «удалить человека с терминалов». Невыполненные регистрации этого номера отменяются.
  async function enqueueDelete(p, by) {
    if (!p.employeeNo) return null;
    await SkudJob.deleteMany({ employeeNo: p.employeeNo, status: "pending" });
    return SkudJob.create({
      type: "delete", student: p.student, fio: p.fio, inn: p.inn, employeeNo: String(p.employeeNo), by,
    });
  }

  // Автоудаление только если человека на терминалы поставила эта система (есть её задание).
  // Остальных (добавленных вручную в iVMS) покажет сверка: удалять их можно кнопкой.
  async function enqueueDeleteForStudent(st, req) {
    const no = String(st.lyceumId || "").trim();
    if (!no) return null;
    const last = await SkudJob.find({ employeeNo: no }, "-photo").sort({ createdAt: -1, _id: -1 }).limit(1);
    if (!last.length || last[0].type === "delete") return null;
    const job = await enqueueDelete({ employeeNo: no, student: st._id, fio: st.fio, inn: st.inn }, actorOf(req).actor);
    await audit(req, "terminal_remove_auto", `${st.fio} №${no}`, "студент удалён из базы");
    return job;
  }

  async function enqueueDeleteForGroup(groupId, req) {
    const students = await Student.find({ group: groupId }, "fio inn lyceumId");
    let n = 0;
    for (const st of students) if (await enqueueDeleteForStudent(st, req)) n++;
    return n;
  }

  /* ------------------------------ списки ------------------------------ */

  app.get("/api/admin/skud-reg/me", requireStaffRole, (req, res) => {
    const s = getSession(req) || {};
    res.json({ role: s.role, username: s.username || "" });
  });

  app.get("/api/admin/skud-reg/groups", requireStaffRole, async (req, res) => {
    try {
      const ids = getAccessibleGroupIds(req);
      res.json(await Group.find(ids === null ? {} : { _id: { $in: ids } }, "name kind").sort({ name: 1 }).lean());
    } catch (e) { fail(res, e); }
  });

  app.get("/api/admin/skud-reg/students", requireStaffRole, async (req, res) => {
    try {
      const q = String(req.query.q || "").trim();
      const gid = req.query.groupId;
      let filter;
      if (gid) {
        if (!validId(gid)) throw httpErr(400, "Некорректная группа");
        if (!canAccessGroup(req, gid)) throw httpErr(403, "Нет доступа к этой группе");
        filter = { group: gid };
      } else if (q.length >= 2) {
        filter = { fio: new RegExp(escRe(q), "i") };
        const ids = getAccessibleGroupIds(req);
        if (ids !== null) filter.group = { $in: ids };
      } else return res.json([]);

      const students = await Student.find(filter, "fio inn lyceumId group skudConsentAt").sort({ fio: 1 }).limit(300).lean();
      const gmap = new Map((await Group.find({ _id: { $in: [...new Set(students.map((s) => String(s.group)))] } }, "name").lean()).map((g) => [String(g._id), g.name]));
      const states = await jobStates(students.map((s) => s._id));
      res.json(students.map((s) => {
        const st = states.get(String(s._id)) || { status: "none", results: [] };
        return {
          id: s._id, fio: s.fio, inn: s.inn || "", lyceumId: s.lyceumId || "",
          group: gmap.get(String(s.group)) || "", groupId: String(s.group || ""),
          ready: isReady(s), consent: !!s.skudConsentAt,
          status: st.status, results: st.results, retrying: !!st.retrying, jobId: st.jobId || "",
        };
      }));
    } catch (e) { fail(res, e); }
  });

  /* ------------------------- сводка и выгрузка ------------------------- */

  async function collect(req, groupId) {
    let groups;
    if (groupId) {
      if (!validId(groupId)) throw httpErr(400, "Некорректная группа");
      if (!canAccessGroup(req, groupId)) throw httpErr(403, "Нет доступа к этой группе");
      groups = await Group.find({ _id: groupId }, "name kind").lean();
    } else {
      const ids = getAccessibleGroupIds(req);
      groups = await Group.find(ids === null ? {} : { _id: { $in: ids } }, "name kind").sort({ name: 1 }).lean();
    }
    const students = groups.length
      ? await Student.find({ group: { $in: groups.map((g) => g._id) } }, "fio inn lyceumId group skudConsentAt").sort({ fio: 1 }).lean()
      : [];
    const states = await jobStates(students.map((s) => s._id));
    return { groups, students, states };
  }

  function rowOf(s, st, gname) {
    const ready = isReady(s);
    const bad = (st.results || []).find((r) => !r.ok);
    return {
      id: s._id, fio: s.fio, group: gname, inn: s.inn || "", lyceumId: s.lyceumId || "",
      status: st.status, ready, consent: !!s.skudConsentAt,
      reason: !ready ? "нет ИНН или поименного №" : bad ? `${bad.ip}: ${bad.error || "ошибка"}` : "",
    };
  }

  app.get("/api/admin/skud-reg/summary", requireStaffRole, async (req, res) => {
    try {
      const { groups, students, states } = await collect(req, req.query.groupId);
      const gname = new Map(groups.map((g) => [String(g._id), g.name]));
      const by = new Map(groups.map((g) => [String(g._id), {
        id: String(g._id), name: g.name, kind: g.kind || "students",
        total: 0, registered: 0, partial: 0, failed: 0, queued: 0, missing: 0, noConsent: 0, notReady: 0,
      }]));
      const detail = [];
      for (const s of students) {
        const st = states.get(String(s._id)) || { status: "none", results: [] };
        const row = by.get(String(s.group));
        if (!row) continue;
        row.total++;
        if (st.status === "done") row.registered++;
        else if (st.status === "partial") row.partial++;
        else if (st.status === "failed" || st.status === "remove_failed") row.failed++;
        else if (["pending", "processing", "removing"].includes(st.status)) row.queued++;
        else row.missing++;
        if (!s.skudConsentAt && st.status !== "done") row.noConsent++;
        if (!isReady(s)) row.notReady++;
        if (req.query.groupId) detail.push(rowOf(s, st, gname.get(String(s.group))));
      }
      const list = [...by.values()];
      const sum = (k) => list.reduce((a, g) => a + g[k], 0);
      res.json({
        groups: list,
        totals: { total: sum("total"), registered: sum("registered"), partial: sum("partial"), failed: sum("failed"), queued: sum("queued"), missing: sum("missing"), notReady: sum("notReady") },
        students: req.query.groupId ? detail : undefined,
      });
    } catch (e) { fail(res, e); }
  });

  app.get("/api/admin/skud-reg/summary.xlsx", requireStaffRole, async (req, res) => {
    try {
      const { groups, students, states } = await collect(req, req.query.groupId);
      const gname = new Map(groups.map((g) => [String(g._id), g.name]));
      const onlyMissing = req.query.onlyMissing === "1";
      const rows = [];
      for (const s of students) {
        const st = states.get(String(s._id)) || { status: "none", results: [] };
        if (onlyMissing && st.status === "done") continue;
        const r = rowOf(s, st, gname.get(String(s.group)));
        rows.push({
          "Группа": r.group, "ФИО": r.fio, "ИНН": r.inn, "Поименный №": r.lyceumId,
          "Статус СКУД": STATUS_RU[r.status] || r.status, "Согласие": r.consent ? "да" : "нет", "Причина": r.reason,
        });
      }
      rows.sort((a, b) => a["Группа"].localeCompare(b["Группа"], "ru") || a["ФИО"].localeCompare(b["ФИО"], "ru"));
      const ws = xlsx.utils.json_to_sheet(rows.length ? rows : [{ "Группа": "", "ФИО": "Все зарегистрированы" }]);
      ws["!cols"] = [{ wch: 14 }, { wch: 34 }, { wch: 16 }, { wch: 14 }, { wch: 24 }, { wch: 10 }, { wch: 40 }];
      const wb = xlsx.utils.book_new();
      xlsx.utils.book_append_sheet(wb, ws, "СКУД");
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="skud_${onlyMissing ? "missing" : "all"}.xlsx"`);
      res.send(xlsx.write(wb, { type: "buffer", bookType: "xlsx" }));
    } catch (e) { fail(res, e); }
  });

  /* --------------------- действия администратора --------------------- */

  app.get("/api/admin/skud-reg/jobs", requireAdmin, async (_req, res) => {
    res.json(await SkudJob.find({}, "-photo").sort({ createdAt: -1, _id: -1 }).limit(50).lean());
  });

  app.get("/api/admin/skud-reg/audit", requireAdmin, async (req, res) => {
    try {
      const filter = req.query.action ? { action: String(req.query.action) } : {};
      const limit = Math.min(Math.max(parseInt(req.query.limit) || 200, 1), 500);
      res.json(await SkudAudit.find(filter).sort({ at: -1, _id: -1 }).limit(limit).lean());
    } catch (e) { fail(res, e); }
  });

  // Удалить с терминалов номер, которого нет в базе студентов (кнопка в отчёте сверки)
  app.post("/api/admin/skud-reg/remove-employee", requireAdmin, async (req, res) => {
    try {
      const no = String(req.body?.employeeNo || "").trim();
      if (!/^\d{1,32}$/.test(no)) throw httpErr(400, "Некорректный номер");
      if (await Student.exists({ lyceumId: no })) throw httpErr(409, "Этот номер принадлежит студенту из базы. Удаляйте через карточку студента.");
      const fio = String(req.body?.name || "").slice(0, 80);
      await enqueueDelete({ employeeNo: no, fio }, actorOf(req).actor);
      await audit(req, "terminal_remove", `№${no} ${fio}`, "лишний на терминале (по сверке)");
      res.json({ ok: true });
    } catch (e) { fail(res, e); }
  });

  app.post("/api/admin/skud-reg/jobs/:id/retry", requireStaffRole, async (req, res) => {
    try {
      if (!validId(req.params.id)) throw httpErr(404, "Не найдено");
      const job = await SkudJob.findById(req.params.id).select("-photo");
      if (!job) throw httpErr(404, "Не найдено");
      if (!isAdmin(req)) {
        const st = job.student && (await Student.findById(job.student));
        if (!st || !canAccessGroup(req, st.group)) throw httpErr(403, "Нет доступа");
      }
      const j = await SkudJob.findOneAndUpdate(
        { _id: req.params.id, status: { $in: ["failed", "partial"] } },
        { status: "pending", attempts: 0, results: [], $unset: { leaseUntil: 1, retryAt: 1 } },
      );
      if (!j) throw httpErr(409, "Нечего повторять");
      await audit(req, "job_retry", `${job.fio} №${job.employeeNo}`);
      res.json({ ok: true });
    } catch (e) { fail(res, e); }
  });

  // Удалить студента с терминалов (остаётся в базе)
  app.post("/api/admin/skud-reg/:studentId/remove", requireAdmin, async (req, res) => {
    try {
      if (!validId(req.params.studentId)) throw httpErr(404, "Студент не найден");
      const st = await Student.findById(req.params.studentId);
      if (!st) throw httpErr(404, "Студент не найден");
      const no = String(st.lyceumId || "").trim();
      if (!no) throw httpErr(400, "У студента нет поименного номера");
      await enqueueDelete({ employeeNo: no, student: st._id, fio: st.fio, inn: st.inn }, actorOf(req).actor);
      await audit(req, "terminal_remove", `${st.fio} №${no}`);
      res.json({ ok: true });
    } catch (e) { fail(res, e); }
  });

  /* ------------------ загрузка фото студента на терминалы ------------------ */

  // Самый последний маршрут с параметром: фиксированные пути объявлены выше
  app.post("/api/admin/skud-reg/:studentId", requireStaffRole, uploadPhoto, async (req, res) => {
    try {
      if (!validId(req.params.studentId)) throw httpErr(404, "Студент не найден");
      const st = await Student.findById(req.params.studentId);
      if (!st) throw httpErr(404, "Студент не найден");
      if (!isAdmin(req) && !canAccessGroup(req, st.group)) throw httpErr(403, "Нет доступа к этой группе");

      const inn = normalizeINN(st.inn);
      const employeeNo = String(st.lyceumId || "").trim();
      if (!inn) throw httpErr(400, "У студента нет корректного ИНН. Заполните его в карточке.");
      if (!/^\d{1,32}$/.test(employeeNo)) throw httpErr(400, "У студента нет поименного номера (только цифры). Заполните его в карточке.");

      // Биометрия: без отметки о согласии (для несовершеннолетних, родителей) не принимаем
      if (!st.skudConsentAt && req.body?.consent !== "1")
        throw httpErr(400, "Подтвердите, что получено согласие на обработку биометрических данных");

      const p = req.file?.buffer;
      const bad = checkPhoto(p);
      if (bad) throw httpErr(400, bad);

      let validUntil;
      if (req.body?.validUntil) {
        validUntil = endOfDayBishkek(req.body.validUntil);
        if (!validUntil || validUntil < new Date() || validUntil > new Date(Date.now() + 10 * 365 * 864e5))
          throw httpErr(400, "Срок доступа: дата должна быть в будущем (не дальше 10 лет)");
      }

      // Тот же Employee ID у другого студента перезапишет чужую карточку на терминале
      if (await Student.exists({ lyceumId: employeeNo, _id: { $ne: st._id } }))
        throw httpErr(409, `Поименный № ${employeeNo} есть у другого студента`);
      if (await SkudJob.exists({ student: st._id, status: "processing", leaseUntil: { $gt: new Date() } }))
        throw httpErr(409, "Этот студент сейчас отправляется");

      if (!st.skudConsentAt)
        await Student.updateOne({ _id: st._id }, { skudConsentAt: new Date(), skudConsentBy: actorOf(req).actor });

      await SkudJob.deleteMany({ student: st._id, status: "pending" });
      const job = await SkudJob.create({
        type: "register", student: st._id, fio: st.fio, inn, employeeNo,
        deviceName: buildDeviceName(st.fio, inn), photo: p, validUntil, by: actorOf(req).actor,
      });
      await audit(req, "photo_upload", `${st.fio} №${employeeNo}`, validUntil ? "до " + req.body.validUntil : "");
      res.json({ ok: true, jobId: job._id, deviceName: job.deviceName });
    } catch (e) { fail(res, e); }
  });

  return { enqueueDelete, enqueueDeleteForStudent, enqueueDeleteForGroup };
};
