// СКУД: преподаватели, мастера, гости и самостоятельная съёмка студентов по временной ссылке.
// Человек заполняет анкету (или присылает только фото) по одноразовой ссылке, а администратор
// (для студентов также куратор группы) проверяет заявку. На терминалы попадает только одобренное.
const crypto = require("crypto");
const path = require("path");
const multer = require("multer");
const mongoose = require("mongoose");
const QRCode = require("qrcode");

// У сотрудников нет поименных номеров, поэтому Employee ID выдаётся автоматически
const KINDS = {
  teachers: { label: "Преподаватели", group: "Преподаватели", base: 7000000, days: null },
  masters: { label: "Мастера", group: "Мастера", base: 8000000, days: null },
  guests: { label: "Гости и практиканты", group: "Гости", base: 9000000, days: 30 },
};
const STAFF_KINDS = Object.keys(KINDS);
const { Schema } = mongoose;
const OID = Schema.Types.ObjectId;
const model = (name, schema) => mongoose.models[name] || mongoose.model(name, schema);

const Invite = model(
  "StaffInvite",
  new Schema(
    {
      token: { type: String, unique: true },
      kind: { type: String, enum: [...STAFF_KINDS, "student"], required: true },
      student: { type: OID, ref: "Student" }, // для ссылок студентов
      accessDays: Number, // сколько дней действует пропуск после одобрения
      expiresAt: Date,
      maxUses: { type: Number, default: 30 },
      uses: { type: Number, default: 0 },
      revoked: { type: Boolean, default: false },
      createdBy: String,
    },
    { timestamps: true },
  ),
);

const Submission = model(
  "StaffSubmission",
  new Schema(
    {
      link: { type: OID, ref: "StaffInvite" },
      kind: { type: String, enum: [...STAFF_KINDS, "student"], required: true },
      fio: String,
      inn: String,
      photo: Buffer,
      accessDays: Number,
      consentAt: Date,
      ip: String,
      status: { type: String, enum: ["new", "approved", "rejected"], default: "new", index: true },
      student: { type: OID, ref: "Student" },
    },
    { timestamps: true },
  ),
);

const Counter = model("StaffCounter", new Schema({ _id: String, seq: { type: Number, default: 0 } }, { versionKey: false }));

const TOKEN_RE = /^[A-Za-z0-9_-]{16,64}$/;
const NAME_RE = /^[\p{L}][\p{L}\s.'’-]{3,78}$/u;
const NO_STORE = { "Cache-Control": "no-store", "X-Robots-Tag": "noindex, nofollow", "Referrer-Policy": "no-referrer" };

module.exports = function mountStaff(app, ctx, core) {
  const { Student, Group, requireAdmin, normalizeINN, getSession, canAccessGroup, requireAuth } = ctx;
  const { SkudJob, buildDeviceName, checkPhoto, audit, notify, mdEsc, actorOf, ipOf } = core;

  // Мягкий лимит по IP. В лицее все на одном Wi-Fi (общий внешний IP), поэтому порог высокий.
  // Главные ограничения: срок ссылки, число заявок по ней и проверка человеком.
  const IP_LIMIT = Number(process.env.STAFF_IP_LIMIT || 40);
  const hits = new Map();
  const recent = (ip) => (hits.get(ip) || []).filter((t) => Date.now() - t < 3600e3);
  const limited = (ip) => recent(ip).length >= IP_LIMIT;
  const hit = (ip) => hits.set(ip, [...recent(ip), Date.now()]);
  setInterval(() => { for (const ip of hits.keys()) if (!recent(ip).length) hits.delete(ip); }, 600e3).unref();

  const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 1e6, files: 1, fields: 10 } });
  const uploadPhoto = (req, res, next) =>
    upload.single("photo")(req, res, (err) => (err ? res.status(400).json({ error: "Файл слишком большой или повреждён" }) : next()));

  const roleOf = (req) => (getSession(req) || {}).role;
  const isAdmin = (req) => roleOf(req) === "admin";
  const requireStaffRole = (req, res, next) =>
    ["admin", "curator"].includes(roleOf(req)) ? next() : res.status(403).json({ error: "Нет доступа" });
  const validId = (id) => mongoose.isValidObjectId(String(id || ""));
  const httpErr = (status, message) => Object.assign(new Error(message), { status });
  const fail = (res, e) => res.status(e.status || 500).json({ error: e.status ? e.message : "Ошибка сервера: " + e.message });

  const baseUrl = (req) =>
    (process.env.RENDER_EXTERNAL_URL || `${req.headers["x-forwarded-proto"] || req.protocol}://${req.get("host")}`).replace(/\/$/, "");

  async function findValidLink(token) {
    if (!TOKEN_RE.test(String(token || ""))) return null;
    const l = await Invite.findOne({ token, revoked: false, expiresAt: { $gt: new Date() } });
    return l && l.uses < l.maxUses ? l : null;
  }

  const label = (kind) => (kind === "student" ? "Студент" : KINDS[kind].label);

  const linkView = (l, req) => {
    let state = "active";
    if (l.revoked) state = "revoked";
    else if (new Date(l.expiresAt) <= new Date()) state = "expired";
    else if (l.uses >= l.maxUses) state = "full";
    return {
      id: l._id, kind: l.kind, label: label(l.kind), state, accessDays: l.accessDays || null,
      url: `${baseUrl(req)}/join/${l.token}`, expiresAt: l.expiresAt, uses: l.uses, maxUses: l.maxUses,
    };
  };

  const maskName = (fio) => {
    const [last = "", ...rest] = String(fio || "").split(/\s+/);
    return `${last} ${rest.map((w) => w[0] + ".").join("")}`.trim();
  };

  async function nextEmployeeNo(kind) {
    for (let i = 0; i < 50; i++) {
      const c = await Counter.findOneAndUpdate({ _id: "staff_" + kind }, { $inc: { seq: 1 } }, { upsert: true, new: true });
      const no = String(KINDS[kind].base + c.seq);
      if (!(await Student.exists({ lyceumId: no }))) return no;
    }
    throw new Error("Не удалось подобрать свободный номер");
  }

  async function ensureGroup(kind) {
    const cfg = KINDS[kind];
    const run = () => Group.findOneAndUpdate({ name: cfg.group }, { $set: { kind }, $setOnInsert: { shiftStart: "08:00" } }, { upsert: true, new: true });
    try { return await run(); } catch { return run(); } // гонка двух одновременных одобрений
  }

  /* ---------- Публичные: анкета и фото по временной ссылке ---------- */

  app.get("/join/:token", async (req, res) => {
    res.set(NO_STORE);
    const link = await findValidLink(req.params.token).catch(() => null);
    if (!link)
      return res.status(410).send(
        '<!DOCTYPE html><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ссылка недействительна</title>' +
          '<body style="font-family:system-ui,sans-serif;display:grid;place-items:center;min-height:100vh;margin:0;background:#f1f5f9;color:#0f172a;text-align:center;padding:20px">' +
          "<div><h2>Ссылка недействительна</h2><p>Срок действия истёк или лимит заявок исчерпан.<br>Попросите администратора выдать новую ссылку.</p></div>",
      );
    res.sendFile(path.join(__dirname, "views", "join.html"));
  });

  app.get("/api/join/:token/info", async (req, res) => {
    res.set(NO_STORE);
    const link = await findValidLink(req.params.token).catch(() => null);
    if (!link) return res.status(410).json({ error: "Ссылка недействительна или истекла" });
    if (link.kind === "student") {
      const st = await Student.findById(link.student);
      return res.json({ mode: "student", label: "Фото для пропуска", name: maskName(st?.fio), expiresAt: link.expiresAt });
    }
    res.json({ mode: "staff", label: label(link.kind), expiresAt: link.expiresAt });
  });

  app.post("/api/join/:token", uploadPhoto, async (req, res) => {
    res.set(NO_STORE);
    try {
      const link = await findValidLink(req.params.token);
      if (!link) return res.status(410).json({ error: "Ссылка недействительна или истекла" });
      const ip = ipOf(req);
      if (limited(ip)) return res.status(429).json({ error: "Слишком много попыток. Попробуйте позже." });

      if (req.body.consent !== "1") return res.status(400).json({ error: "Нужно согласие на обработку данных и фото" });
      const p = req.file?.buffer;
      const bad = checkPhoto(p);
      if (bad) return res.status(400).json({ error: bad });

      let rec;
      if (link.kind === "student") {
        const st = await Student.findById(link.student);
        if (!st) return res.status(410).json({ error: "Ссылка недействительна" });
        if (await Submission.exists({ student: st._id, kind: "student", status: "new" }))
          return res.status(409).json({ error: "Фото уже отправлено, дождитесь проверки." });
        rec = { kind: "student", student: st._id, fio: st.fio, inn: st.inn };
      } else {
        const fio = String(req.body.fio || "").replace(/\s+/g, " ").trim();
        const inn = String(req.body.inn || "").replace(/\s/g, "");
        if (!NAME_RE.test(fio) || fio.split(" ").length < 2) return res.status(400).json({ error: "Введите фамилию, имя и отчество полностью" });
        if (!/^\d{14}$/.test(inn) || !normalizeINN(inn)) return res.status(400).json({ error: "ИНН должен состоять из 14 цифр" });
        if ((await Student.exists({ inn })) || (await Submission.exists({ inn, status: "new" }))) {
          hit(ip); // защита от перебора ИНН
          return res.status(409).json({ error: "Этот ИНН уже зарегистрирован. Обратитесь к администратору." });
        }
        rec = { kind: link.kind, fio, inn, accessDays: link.accessDays || undefined };
      }

      // Атомарно занимаем место по ссылке (срок и лимит проверяются в самом запросе)
      const claimed = await Invite.findOneAndUpdate(
        { _id: link._id, revoked: false, expiresAt: { $gt: new Date() }, $expr: { $lt: ["$uses", "$maxUses"] } },
        { $inc: { uses: 1 } },
      );
      if (!claimed) return res.status(410).json({ error: "Лимит заявок по ссылке исчерпан" });
      try {
        await Submission.create({ ...rec, link: link._id, photo: p, consentAt: new Date(), ip });
      } catch (e) {
        await Invite.updateOne({ _id: link._id }, { $inc: { uses: -1 } });
        throw e;
      }
      hit(ip);
      if (rec.kind !== "student") await notify(`📥 *СКУД*: новая анкета (${mdEsc(label(rec.kind))})\n${mdEsc(rec.fio)}\nПроверьте на странице «Преподаватели и мастера»`);
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: "Не удалось отправить, попробуйте ещё раз" });
    }
  });

  /* ---------- Администратор: ссылки для сотрудников ---------- */

  app.get("/skud-staff.html", requireAdmin, (_req, res) => res.sendFile(path.join(__dirname, "views", "skud-staff.html")));
  app.get("/skud-status.html", requireAuth, (_req, res) => res.sendFile(path.join(__dirname, "views", "skud-status.html")));

  app.post("/api/admin/staff-links", requireAdmin, async (req, res) => {
    try {
      const kind = String(req.body.kind || "");
      if (!KINDS[kind]) throw httpErr(400, "Выберите: преподаватели, мастера или гости");
      const hours = Math.min(Math.max(Number(req.body.hours) || 24, 1), 24 * 14);
      const maxUses = Math.min(Math.max(parseInt(req.body.maxUses) || 30, 1), 500);
      const rawDays = req.body.accessDays === "" || req.body.accessDays == null ? KINDS[kind].days : parseInt(req.body.accessDays);
      const accessDays = rawDays ? Math.min(Math.max(rawDays, 1), 3650) : undefined;
      const link = await Invite.create({
        token: crypto.randomBytes(18).toString("base64url"), kind, maxUses, accessDays,
        expiresAt: new Date(Date.now() + hours * 3600e3), createdBy: actorOf(req).actor,
      });
      await audit(req, "link_create", label(kind), `${hours} ч, до ${maxUses} заявок${accessDays ? `, пропуск ${accessDays} дн.` : ""}`);
      res.json(linkView(link, req));
    } catch (e) { fail(res, e); }
  });

  app.get("/api/admin/staff-links", requireAdmin, async (req, res) => {
    const links = await Invite.find({ kind: { $in: STAFF_KINDS } }).sort({ createdAt: -1, _id: -1 }).limit(30);
    res.json(links.map((l) => linkView(l, req)));
  });

  app.get("/api/admin/staff-links/:id/qr", requireAdmin, async (req, res) => {
    try {
      const l = validId(req.params.id) && (await Invite.findById(req.params.id));
      if (!l) return res.status(404).end();
      res.type("png").set("Cache-Control", "private, no-store").send(await QRCode.toBuffer(linkView(l, req).url, { width: 360, margin: 2 }));
    } catch { res.status(404).end(); }
  });

  app.delete("/api/admin/staff-links/:id", requireAdmin, async (req, res) => {
    try {
      if (!validId(req.params.id)) throw httpErr(404, "Не найдено");
      await Invite.updateOne({ _id: req.params.id }, { revoked: true });
      await audit(req, "link_revoke", String(req.params.id));
      res.json({ ok: true });
    } catch (e) { fail(res, e); }
  });

  /* ---------- Ссылка студенту для самостоятельной съёмки ---------- */

  app.post("/api/admin/skud-selflink/:studentId", requireStaffRole, async (req, res) => {
    try {
      if (!validId(req.params.studentId)) throw httpErr(404, "Студент не найден");
      const st = await Student.findById(req.params.studentId);
      if (!st) throw httpErr(404, "Студент не найден");
      if (!isAdmin(req) && !canAccessGroup(req, st.group)) throw httpErr(403, "Нет доступа к этой группе");
      if (!normalizeINN(st.inn) || !/^\d{1,32}$/.test(String(st.lyceumId || "").trim()))
        throw httpErr(400, "Сначала заполните ИНН и поименный № в карточке студента");

      const hours = Math.min(Math.max(Number(req.body?.hours) || 72, 1), 24 * 14);
      await Invite.updateMany({ kind: "student", student: st._id, revoked: false }, { revoked: true });
      const link = await Invite.create({
        token: crypto.randomBytes(18).toString("base64url"), kind: "student", student: st._id, maxUses: 2,
        expiresAt: new Date(Date.now() + hours * 3600e3), createdBy: actorOf(req).actor,
      });
      const v = linkView(link, req);
      await audit(req, "selflink_create", `${st.fio} №${st.lyceumId}`, `${hours} ч`);
      res.json({ ...v, fio: st.fio, qr: await QRCode.toDataURL(v.url, { width: 320, margin: 2 }) });
    } catch (e) { fail(res, e); }
  });

  /* ---------- Проверка заявок ---------- */

  // Права на заявку: сотрудников проверяет только администратор, студентов ещё и куратор группы
  async function canModerate(req, sub) {
    if (isAdmin(req)) return true;
    if (roleOf(req) !== "curator" || sub.kind !== "student" || !sub.student) return false;
    const st = await Student.findById(sub.student);
    return !!st && canAccessGroup(req, st.group);
  }

  app.get("/api/admin/staff-submissions", requireStaffRole, async (req, res) => {
    try {
      const scope = req.query.scope === "student" ? "student" : "staff";
      if (scope === "staff" && !isAdmin(req)) throw httpErr(403, "Только для администратора");
      const subs = await Submission.find(scope === "student" ? { status: "new", kind: "student" } : { status: "new", kind: { $in: STAFF_KINDS } }, "-photo")
        .sort({ createdAt: 1 }).limit(150);
      const out = [];
      for (const s of subs) {
        if (scope === "student" && !(await canModerate(req, s))) continue;
        let group = "";
        if (s.student && scope === "student") {
          const st = await Student.findById(s.student);
          group = st ? (await Group.findById(st.group))?.name || "" : "";
        }
        out.push({ id: s._id, kind: s.kind, label: label(s.kind), fio: s.fio, inn: s.inn, group, accessDays: s.accessDays || null, createdAt: s.createdAt });
      }
      res.json(out);
    } catch (e) { fail(res, e); }
  });

  app.get("/api/admin/staff-submissions/:id/photo", requireStaffRole, async (req, res) => {
    try {
      const s = validId(req.params.id) && (await Submission.findById(req.params.id));
      if (!s?.photo || !(await canModerate(req, s))) return res.status(404).end();
      res.type("jpeg").set("Cache-Control", "private, no-store").send(s.photo);
    } catch { res.status(404).end(); }
  });

  app.post("/api/admin/staff-submissions/:id/approve", requireStaffRole, async (req, res) => {
    let sub = null;
    try {
      if (!validId(req.params.id)) throw httpErr(404, "Заявка не найдена");
      const peek = await Submission.findById(req.params.id, "-photo");
      if (!peek || peek.status !== "new") throw httpErr(404, "Заявка уже обработана");
      if (!(await canModerate(req, peek))) throw httpErr(403, "Нет доступа к заявке");

      sub = await Submission.findOneAndUpdate({ _id: peek._id, status: "new" }, { status: "approved" });
      if (!sub) throw httpErr(404, "Заявка уже обработана");
      const by = actorOf(req).actor;

      if (sub.kind === "student") {
        const st = await Student.findById(sub.student);
        const inn = st && normalizeINN(st.inn);
        const no = String(st?.lyceumId || "").trim();
        if (!st || !inn || !/^\d{1,32}$/.test(no)) throw httpErr(409, "У студента нет ИНН или поименного №");
        await Student.updateOne({ _id: st._id }, { skudConsentAt: sub.consentAt || new Date(), skudConsentBy: "ссылка студента" });
        await SkudJob.deleteMany({ student: st._id, status: "pending" });
        await SkudJob.create({ type: "register", student: st._id, fio: st.fio, inn, employeeNo: no, deviceName: buildDeviceName(st.fio, inn), photo: sub.photo, by });
        await Submission.updateOne({ _id: sub._id }, { $unset: { photo: 1 } });
        await audit(req, "student_photo_approve", `${st.fio} №${no}`);
        return res.json({ ok: true, employeeNo: no });
      }

      if (await Student.exists({ inn: sub.inn })) throw httpErr(409, "Сотрудник с таким ИНН уже есть");
      const cfg = KINDS[sub.kind];
      const group = await ensureGroup(sub.kind);
      const employeeNo = await nextEmployeeNo(sub.kind);
      const student = await Student.create({
        fio: sub.fio, inn: sub.inn, lyceumId: employeeNo, group: group._id, skudConsentAt: sub.consentAt || new Date(), skudConsentBy: "анкета по ссылке",
      });
      const validUntil = sub.accessDays ? new Date(Date.now() + sub.accessDays * 864e5) : undefined;
      await SkudJob.create({
        type: "register", student: student._id, fio: sub.fio, inn: sub.inn, employeeNo,
        deviceName: buildDeviceName(sub.fio, sub.inn), photo: sub.photo, validUntil, by,
      });
      await Submission.updateOne({ _id: sub._id }, { $unset: { photo: 1 }, student: student._id });
      await audit(req, "staff_approve", `${sub.fio} №${employeeNo}`, `${cfg.group}${sub.accessDays ? `, пропуск ${sub.accessDays} дн.` : ""}`);
      res.json({ ok: true, employeeNo, group: cfg.group });
    } catch (e) {
      if (sub) await Submission.updateOne({ _id: sub._id }, { status: "new" }).catch(() => {});
      fail(res, e);
    }
  });

  app.post("/api/admin/staff-submissions/:id/reject", requireStaffRole, async (req, res) => {
    try {
      if (!validId(req.params.id)) throw httpErr(404, "Не найдено");
      const peek = await Submission.findById(req.params.id, "-photo");
      if (!peek || peek.status !== "new") throw httpErr(404, "Заявка уже обработана");
      if (!(await canModerate(req, peek))) throw httpErr(403, "Нет доступа к заявке");
      await Submission.updateOne({ _id: peek._id, status: "new" }, { status: "rejected", $unset: { photo: 1 } });
      await audit(req, "submission_reject", `${peek.fio || ""}`, label(peek.kind));
      res.json({ ok: true });
    } catch (e) { fail(res, e); }
  });

  // Последние отправки сотрудников на терминалы
  app.get("/api/admin/staff-jobs", requireAdmin, async (_req, res) => {
    const gids = (await Group.find({ kind: { $in: STAFF_KINDS } }, "_id").lean()).map((g) => g._id);
    const sids = gids.length ? (await Student.find({ group: { $in: gids } }, "_id").lean()).map((s) => s._id) : [];
    const jobs = sids.length ? await SkudJob.find({ student: { $in: sids } }, "-photo").sort({ createdAt: -1, _id: -1 }).limit(20).lean() : [];
    res.json(jobs.map((j) => ({ id: j._id, fio: j.fio, employeeNo: j.employeeNo, status: j.status, results: j.results || [], createdAt: j.createdAt })));
  });

  // Не храним лишнее: обработанные заявки и старые необработанные удаляются
  setInterval(() => {
    const day = new Date(Date.now() - 24 * 3600e3), two = new Date(Date.now() - 14 * 24 * 3600e3);
    Submission.deleteMany({ $or: [{ status: { $ne: "new" }, updatedAt: { $lt: day } }, { status: "new", createdAt: { $lt: two } }] }).catch(() => {});
  }, 3600e3).unref();
};

module.exports.STAFF_KINDS = STAFF_KINDS;
