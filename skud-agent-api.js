// СКУД: API для локального агента (очередь заданий, сигналы о состоянии терминалов, сверка)
// и страницы администратора «Терминалы» и «Сверка». Агент авторизуется ключом AGENT_API_KEY.
const mongoose = require("mongoose");

module.exports = function mountAgentApi(app, ctx, core) {
  const { Student, requireAdmin, AGENT_API_KEY } = ctx;
  const { SkudJob, SkudDevice, SkudReport, notify, mdEsc, audit, safeEq } = core;

  const LEASE_MS = 2 * 60 * 1000; // сколько агент "держит" задание
  const MAX_ATTEMPTS = 5; // всего попыток на задание
  const RETRY_BASE_MS = 60 * 1000; // пауза перед автоповтором растёт: 1, 2, 3 минуты
  const OFFLINE_NOTIFY_MS = 3 * 60 * 1000; // терминал не отвечает столько, и тогда уведомляем
  const ONLINE_FRESH_MS = 3 * 60 * 1000; // сигнал агента свежий, если не старше
  const SILENT_MS = 5 * 60 * 1000; // агент молчит дольше: считаем, что он не работает

  const keyOk = (req) => safeEq(req.body?.apiKey, AGENT_API_KEY);
  const validId = (id) => mongoose.isValidObjectId(String(id || ""));
  const hhmm = (d) => new Date(d).toLocaleTimeString("ru-RU", { hour: "2-digit", minute: "2-digit", timeZone: "Asia/Bishkek" });

  /* ============================ задания ============================ */

  app.post("/api/agent/jobs/claim", async (req, res) => {
    if (!keyOk(req)) return res.status(403).send("Forbidden");
    try {
      const now = new Date();
      // Задания, у которых кончились попытки и аренда: помечаем ошибкой, чтобы не висели «отправляется»
      await SkudJob.updateMany(
        { status: "processing", leaseUntil: { $lt: now }, attempts: { $gte: MAX_ATTEMPTS } },
        { status: "failed", results: [{ ip: "агент", ok: false, error: "агент не ответил после нескольких попыток" }], $unset: { leaseUntil: 1 } },
      );
      const job = await SkudJob.findOneAndUpdate(
        {
          attempts: { $lt: MAX_ATTEMPTS },
          $or: [
            { status: "pending", $or: [{ retryAt: null }, { retryAt: { $lte: now } }] },
            { status: "processing", leaseUntil: { $lt: now } },
          ],
        },
        { status: "processing", leaseUntil: new Date(now.getTime() + LEASE_MS), $inc: { attempts: 1 } },
        { sort: { createdAt: 1 }, new: true },
      );
      const flag = await SkudReport.findOneAndUpdate({ _id: "last", requested: true }, { requested: false });

      let out = null;
      if (job) {
        const type = job.type || "register";
        if (type === "register" && !job.photo) {
          await SkudJob.updateOne({ _id: job._id }, { status: "failed", results: [{ ip: "сервер", ok: false, error: "нет фото" }], $unset: { leaseUntil: 1 } });
        } else {
          out = {
            id: job._id, type, employeeNo: job.employeeNo, name: job.deviceName,
            photo: job.photo ? job.photo.toString("base64") : null,
            validUntil: job.validUntil ? new Date(job.validUntil).toISOString() : null,
          };
        }
      }
      res.json({ job: out, reconcile: !!flag });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  app.post("/api/agent/jobs/:id/result", async (req, res) => {
    if (!keyOk(req)) return res.status(403).send("Forbidden");
    try {
      if (!validId(req.params.id)) return res.status(404).json({ error: "Не найдено" });
      const results = (Array.isArray(req.body.results) ? req.body.results : []).slice(0, 50).map((r) => ({
        ip: String(r.ip || "").slice(0, 64), ok: !!r.ok, error: r.ok ? undefined : String(r.error || "").slice(0, 300),
      }));
      const okN = results.filter((r) => r.ok).length;
      let status = results.length && okN === results.length ? "done" : okN ? "partial" : "failed";

      const job = await SkudJob.findOne({ _id: req.params.id, status: "processing" }, "-photo");
      if (!job) return res.json({ ok: true, ignored: true }); // повторная отправка результата
      const type = job.type || "register";

      // Сбой терминала (перезагрузка, сеть) обычно временный: повторяем сами с нарастающей паузой
      const retry = status !== "done" && job.attempts < 3;
      const update = { status: retry ? "pending" : status, results, $unset: { leaseUntil: 1 } };
      if (retry) { update.retryAt = new Date(Date.now() + RETRY_BASE_MS * job.attempts); }
      else { update.$unset.retryAt = 1; }
      if (status === "done") update.$unset.photo = 1; // биометрия не хранится после успеха
      await SkudJob.updateOne({ _id: job._id }, update);

      if (!retry && status !== "done") {
        const errs = results.filter((r) => !r.ok).map((r) => `${r.ip}: ${r.error || "ошибка"}`).join("\n");
        await notify(
          `⚠️ *СКУД*: ${type === "delete" ? "не удалён с терминалов" : "не записан на терминалы"}\n` +
          `${mdEsc(job.fio || "—")} (№${mdEsc(job.employeeNo)})\n${mdEsc(errs)}`,
        );
      }
      await audit(null, "job_result", `${job.fio || ""} №${job.employeeNo}`, `${type}: ${retry ? "повтор" : status}`);
      res.json({ ok: true, status: update.status });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  /* ================== состояние терминалов (сигналы) ================== */

  app.post("/api/agent/heartbeat", async (req, res) => {
    if (!keyOk(req)) return res.status(403).send("Forbidden");
    try {
      const list = (Array.isArray(req.body.devices) ? req.body.devices : []).slice(0, 50);
      const now = new Date();
      for (const d of list) {
        const ip = String(d.ip || "").slice(0, 64);
        if (!ip) continue;
        const was = await SkudDevice.findOne({ ip });
        const online = !!d.online;
        const set = {
          online, lastSeen: now, silentNotified: false,
          latencyMs: Number.isFinite(d.ms) ? Math.round(d.ms) : undefined,
          users: Number.isFinite(d.users) ? d.users : was?.users,
          error: online ? "" : String(d.error || "нет ответа").slice(0, 200),
        };
        if (online) {
          set.lastOk = now;
          set.offlineSince = null;
          set.offlineNotified = false;
          if (was?.offlineNotified) await notify(`✅ *СКУД*: терминал ${mdEsc(ip)} снова в сети`);
        } else {
          const since = was && was.online === false && was.offlineSince ? was.offlineSince : now;
          set.offlineSince = since;
          if (!was?.offlineNotified && now - new Date(since) >= OFFLINE_NOTIFY_MS) {
            set.offlineNotified = true;
            await notify(`🔴 *СКУД*: терминал ${mdEsc(ip)} не отвечает с ${hhmm(since)}\n${mdEsc(set.error)}`);
          } else set.offlineNotified = was?.offlineNotified || false;
        }
        await SkudDevice.findOneAndUpdate({ ip }, { $set: set }, { upsert: true });
      }
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Агент перестал присылать сигналы (выключен ПК, нет интернета в лицее)
  async function checkSilence(now = new Date()) {
    const stale = await SkudDevice.find({ lastSeen: { $lt: new Date(now.getTime() - SILENT_MS) }, silentNotified: { $ne: true } });
    if (!stale.length) return 0;
    for (const d of stale) await SkudDevice.updateOne({ _id: d._id }, { silentNotified: true });
    const last = stale.map((d) => new Date(d.lastSeen).getTime()).sort((a, b) => b - a)[0];
    await notify(`🟠 *СКУД*: агент в лицее не выходит на связь с ${hhmm(last)}. Проверьте компьютер и интернет.`);
    return stale.length;
  }
  setInterval(() => checkSilence().catch(() => {}), 60 * 1000).unref();

  app.get("/api/admin/skud-reg/devices", requireAdmin, async (_req, res) => {
    const now = Date.now();
    const list = await SkudDevice.find({}).sort({ ip: 1 }).lean();
    res.json(list.map((d) => {
      const age = d.lastSeen ? now - new Date(d.lastSeen).getTime() : Infinity;
      const state = age > SILENT_MS ? "silent" : d.online && age <= ONLINE_FRESH_MS ? "online" : d.online === false ? "offline" : "online";
      return { ip: d.ip, state, lastSeen: d.lastSeen, lastOk: d.lastOk, offlineSince: d.offlineSince, latencyMs: d.latencyMs, users: d.users, error: d.error || "" };
    }));
  });

  /* =========================== сверка =========================== */

  app.post("/api/agent/reconcile", async (req, res) => {
    if (!keyOk(req)) return res.status(403).send("Forbidden");
    try {
      const reports = (Array.isArray(req.body.devices) ? req.body.devices : []).slice(0, 50);

      const students = await Student.find({ lyceumId: { $exists: true } }, "fio lyceumId").lean();
      const byNo = new Map(students.filter((s) => s.lyceumId).map((s) => [String(s.lyceumId).trim(), s]));

      // Кто должен быть на терминалах: у кого самое свежее задание «записать» выполнено
      const jobs = await SkudJob.find({ status: "done" }, "-photo").sort({ createdAt: -1, _id: -1 }).lean();
      const newest = new Map();
      for (const j of jobs) if (!newest.has(j.employeeNo)) newest.set(j.employeeNo, j);
      const expected = [...newest.values()].filter((j) => (j.type || "register") === "register");

      const devices = reports.map((r) => {
        const ip = String(r.ip || "").slice(0, 64);
        if (!r.ok) return { ip, ok: false, error: String(r.error || "нет ответа").slice(0, 200), total: 0, extra: [], missing: [] };
        const people = (Array.isArray(r.people) ? r.people : []).slice(0, 20000);
        const have = new Set(people.map((p) => String(p.employeeNo)));
        const extra = people
          .filter((p) => !byNo.has(String(p.employeeNo)))
          .slice(0, 500)
          .map((p) => ({ employeeNo: String(p.employeeNo), name: String(p.name || "").slice(0, 80) }));
        const missing = expected
          .filter((j) => !have.has(String(j.employeeNo)))
          .slice(0, 500)
          .map((j) => ({ employeeNo: String(j.employeeNo), fio: byNo.get(String(j.employeeNo))?.fio || j.fio || "" }));
        return { ip, ok: true, total: people.length, extra, missing };
      });

      const sig = devices.map((d) => `${d.ip}:${d.ok}:${d.extra.length}:${d.missing.length}`).join("|");
      const prev = await SkudReport.findOne({ _id: "last" });
      await SkudReport.findOneAndUpdate({ _id: "last" }, { $set: { at: new Date(), devices, sig } }, { upsert: true });

      const extraN = devices.reduce((a, d) => a + d.extra.length, 0);
      const missN = devices.reduce((a, d) => a + d.missing.length, 0);
      if (sig !== prev?.sig && (extraN || missN))
        await notify(`🔎 *СКУД*: сверка терминалов с базой\nлишних на терминалах: ${extraN}, не хватает: ${missN}\nПодробности: страница «СКУД: сводка» → «Сверка»`);
      res.json({ ok: true, extra: extraN, missing: missN });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  app.get("/api/admin/skud-reg/reconcile", requireAdmin, async (_req, res) => {
    const r = await SkudReport.findOne({ _id: "last" }).lean();
    res.json({ at: r?.at || null, requested: !!r?.requested, devices: r?.devices || [] });
  });

  app.post("/api/admin/skud-reg/reconcile/run", requireAdmin, async (req, res) => {
    await SkudReport.findOneAndUpdate({ _id: "last" }, { $set: { requested: true, requestedAt: new Date() } }, { upsert: true });
    await audit(req, "reconcile_request", "терминалы");
    res.json({ ok: true });
  });

  return { checkSilence };
};
