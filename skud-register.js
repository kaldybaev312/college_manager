// Регистрация студентов в Hikvision СКУД: серверная часть (очередь заданий для локального агента).
// Подключение в server.js: require("./skud-register")(app, { Student, Group, requireAdmin, normalizeINN, AGENT_API_KEY });
const crypto = require("crypto");
const multer = require("multer");
const mongoose = require("mongoose");

const NAME_MAX = Number(process.env.SKUD_NAME_MAX || 32); // лимит поля Name на терминале
const LEASE_MS = 2 * 60 * 1000; // сколько агент "держит" задание
const MAX_ATTEMPTS = 5;
const MAX_PHOTO = 200 * 1024; // Hikvision принимает JPEG примерно до 200 КБ

const SkudJob =
  mongoose.models.SkudJob ||
  mongoose.model(
    "SkudJob",
    new mongoose.Schema(
      {
        student: {
          type: mongoose.Schema.Types.ObjectId,
          ref: "Student",
          index: true,
        },
        fio: String,
        inn: String,
        employeeNo: String, // Employee ID = Student.lyceumId (поименный №)
        deviceName: String, // Name на терминале = ФИО + ИНН
        photo: Buffer,
        status: {
          type: String,
          enum: ["pending", "processing", "done", "partial", "failed"],
          default: "pending",
          index: true,
        },
        attempts: { type: Number, default: 0 },
        leaseUntil: Date,
        results: [{ _id: false, ip: String, ok: Boolean, error: String }],
      },
      { timestamps: true },
    ),
  );

// Name на устройстве. ИНН всегда целиком: по нему /api/agent/sync находит студента.
function buildDeviceName(fio, inn) {
  const clean = String(fio || "")
    .replace(/\s+/g, " ")
    .trim();
  const full = `${clean} ${inn}`;
  if (full.length <= NAME_MAX) return full;
  const [last = "", ...rest] = clean.split(" ");
  const initials = rest
    .filter(Boolean)
    .map((w) => w[0].toUpperCase() + ".")
    .join("");
  const room = NAME_MAX - inn.length - 1;
  return `${`${last} ${initials}`.trim().slice(0, Math.max(room, 1))} ${inn}`;
}

const esc = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
const safeEq = (a, b) => {
  const x = Buffer.from(String(a || "")),
    y = Buffer.from(String(b || ""));
  return x.length === y.length && crypto.timingSafeEqual(x, y);
};

module.exports = function mountSkudRegister(
  app,
  { Student, Group, requireAdmin, normalizeINN, AGENT_API_KEY },
) {
  const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 1e6 },
  });

  /* ---------- Админка (после app.use("/api/admin", requireAuth)) ---------- */

  app.get("/api/admin/skud-reg/groups", requireAdmin, async (_req, res) => {
    res.json(await Group.find({}, "name").sort({ name: 1 }).lean());
  });

  // Список студентов группы или поиск по ФИО + статус последней отправки
  app.get("/api/admin/skud-reg/students", requireAdmin, async (req, res) => {
    try {
      const q = String(req.query.q || "").trim();
      const filter = {};
      if (req.query.groupId) filter.group = req.query.groupId;
      else if (q.length >= 2) filter.fio = new RegExp(esc(q), "i");
      else return res.json([]);
      const students = await Student.find(filter, "fio inn lyceumId group")
        .populate("group", "name")
        .sort({ fio: 1 })
        .limit(200)
        .lean();

      const jobs = await SkudJob.find(
        { student: { $in: students.map((s) => s._id) } },
        "student status results createdAt",
      )
        .sort({ createdAt: -1 })
        .lean();
      const last = new Map();
      for (const j of jobs)
        if (!last.has(String(j.student))) last.set(String(j.student), j);

      res.json(
        students.map((s) => {
          const j = last.get(String(s._id));
          return {
            id: s._id,
            fio: s.fio,
            inn: s.inn || "",
            lyceumId: s.lyceumId || "",
            group: s.group?.name || "",
            ready:
              !!normalizeINN(s.inn) &&
              /^\d{1,32}$/.test(String(s.lyceumId || "").trim()),
            status: j?.status || "none",
            results: j?.results || [],
          };
        }),
      );
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Поставить студента в очередь на регистрацию (фото уже уменьшено в браузере)
  app.post(
    "/api/admin/skud-reg/:studentId",
    requireAdmin,
    upload.single("photo"),
    async (req, res) => {
      try {
        const st = await Student.findById(req.params.studentId);
        if (!st) return res.status(404).json({ error: "Студент не найден" });

        const inn = normalizeINN(st.inn);
        const employeeNo = String(st.lyceumId || "").trim();
        if (!inn)
          return res
            .status(400)
            .json({
              error:
                "У студента нет корректного ИНН. Заполните его в карточке.",
            });
        if (!/^\d{1,32}$/.test(employeeNo))
          return res
            .status(400)
            .json({
              error:
                "У студента нет поименного номера (только цифры). Заполните его в карточке.",
            });

        const p = req.file?.buffer;
        if (!p) return res.status(400).json({ error: "Загрузите фото" });
        if (p[0] !== 0xff || p[1] !== 0xd8)
          return res.status(400).json({ error: "Нужен JPEG" });
        if (p.length > MAX_PHOTO)
          return res.status(400).json({ error: "Фото больше 200 КБ" });

        // Тот же Employee ID у другого студента перезапишет чужую карточку на терминале
        if (
          await Student.exists({ lyceumId: employeeNo, _id: { $ne: st._id } })
        )
          return res
            .status(409)
            .json({
              error: `Поименный № ${employeeNo} есть у другого студента`,
            });
        if (
          await SkudJob.exists({
            student: st._id,
            status: "processing",
            leaseUntil: { $gt: new Date() },
          })
        )
          return res
            .status(409)
            .json({ error: "Этот студент сейчас отправляется" });

        await SkudJob.deleteMany({ student: st._id, status: "pending" });
        const job = await SkudJob.create({
          student: st._id,
          fio: st.fio,
          inn,
          employeeNo,
          deviceName: buildDeviceName(st.fio, inn),
          photo: p,
        });
        res.json({ ok: true, jobId: job._id, deviceName: job.deviceName });
      } catch (e) {
        res.status(500).json({ error: e.message });
      }
    },
  );

  app.get("/api/admin/skud-reg/jobs", requireAdmin, async (_req, res) => {
    res.json(
      await SkudJob.find({}, "-photo").sort({ createdAt: -1 }).limit(50).lean(),
    );
  });

  app.post(
    "/api/admin/skud-reg/jobs/:id/retry",
    requireAdmin,
    async (req, res) => {
      const j = await SkudJob.findOneAndUpdate(
        {
          _id: req.params.id,
          status: { $in: ["failed", "partial"] },
          photo: { $exists: true },
        },
        {
          status: "pending",
          attempts: 0,
          results: [],
          $unset: { leaseUntil: 1 },
        },
      );
      res.json(j ? { ok: true } : { error: "Нечего повторять" });
    },
  );

  /* ---------- Локальный агент (авторизация по AGENT_API_KEY, как /api/agent/sync) ---------- */

  app.post("/api/agent/jobs/claim", async (req, res) => {
    if (!safeEq(req.body?.apiKey, AGENT_API_KEY))
      return res.status(403).send("Forbidden");
    const now = new Date();
    const job = await SkudJob.findOneAndUpdate(
      {
        attempts: { $lt: MAX_ATTEMPTS },
        $or: [
          { status: "pending" },
          { status: "processing", leaseUntil: { $lt: now } },
        ],
      },
      {
        status: "processing",
        leaseUntil: new Date(now.getTime() + LEASE_MS),
        $inc: { attempts: 1 },
      },
      { sort: { createdAt: 1 }, new: true },
    );
    if (!job) return res.json({ job: null });
    res.json({
      job: {
        id: job._id,
        employeeNo: job.employeeNo,
        name: job.deviceName,
        photo: job.photo.toString("base64"),
      },
    });
  });

  app.post("/api/agent/jobs/:id/result", async (req, res) => {
    if (!safeEq(req.body?.apiKey, AGENT_API_KEY))
      return res.status(403).send("Forbidden");
    const results = (
      Array.isArray(req.body.results) ? req.body.results : []
    ).map((r) => ({
      ip: String(r.ip || ""),
      ok: !!r.ok,
      error: r.ok ? undefined : String(r.error || "").slice(0, 300),
    }));
    const okN = results.filter((r) => r.ok).length;
    const status =
      results.length && okN === results.length
        ? "done"
        : okN
          ? "partial"
          : "failed";
    const update = { status, results, $unset: { leaseUntil: 1 } };
    if (status === "done") update.$unset.photo = 1; // биометрия не хранится после успеха
    await SkudJob.findByIdAndUpdate(req.params.id, update);
    res.json({ ok: true, status });
  });
};

module.exports.buildDeviceName = buildDeviceName;
