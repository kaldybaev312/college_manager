// Причины отсутствия: болезнь, семейные обстоятельства, по направлению лицея, другое.
// Запись покрывает период dateFrom..dateTo (ключи "ГГГГ-ММ-ДД").
// Куратор/админ ставит причину сразу подтверждённой, родитель — на подтверждение куратору.

const mongoose = require("mongoose");
const { isValidKey, datesBetween } = require("./attendance-utils");

const { Schema } = mongoose;

const REASON_TYPES = {
  sick: { label: "Болезнь", code: "Б", excused: true },
  family: { label: "Семейные обстоятельства", code: "С", excused: true },
  official: { label: "По направлению лицея", code: "Л", excused: true },
  other: { label: "Другая причина", code: "Д", excused: false },
};
const MAX_DAYS = 31;

const AbsenceReason =
  mongoose.models.AbsenceReason ||
  mongoose.model(
    "AbsenceReason",
    new Schema(
      {
        student: { type: Schema.Types.ObjectId, ref: "Student", index: true },
        group: { type: Schema.Types.ObjectId, ref: "Group", index: true },
        dateFrom: String,
        dateTo: String,
        type: { type: String, enum: Object.keys(REASON_TYPES), default: "other" },
        comment: { type: String, default: "" },
        excused: { type: Boolean, default: false },
        confirmed: { type: Boolean, default: true },
        source: { type: String, enum: ["admin", "curator", "parent"], default: "curator" },
        createdBy: { type: String, default: "" },
      },
      { timestamps: true },
    ),
  );

function serialize(r) {
  const t = REASON_TYPES[r.type] || REASON_TYPES.other;
  return {
    id: String(r._id),
    studentId: String(r.student),
    groupId: String(r.group || ""),
    dateFrom: r.dateFrom,
    dateTo: r.dateTo,
    type: r.type,
    label: t.label,
    code: t.code,
    excused: !!r.excused,
    confirmed: r.confirmed !== false,
    comment: r.comment || "",
    source: r.source,
    createdBy: r.createdBy || "",
  };
}

// Причины, пересекающие период from..to. filter: { group } или { student: {$in: [...] } }
async function findReasons(filter, from, to) {
  return AbsenceReason.find({ ...filter, dateFrom: { $lte: to }, dateTo: { $gte: from } })
    .sort({ createdAt: 1 })
    .lean();
}

// Map "studentId|ГГГГ-ММ-ДД" → причина (более поздняя запись перекрывает раннюю)
async function reasonsIndex(filter, from, to) {
  const map = new Map();
  for (const r of await findReasons(filter, from, to)) {
    const s = serialize(r);
    const a = r.dateFrom > from ? r.dateFrom : from;
    const b = r.dateTo < to ? r.dateTo : to;
    for (const dk of datesBetween(a, b)) map.set(`${s.studentId}|${dk}`, s);
  }
  return map;
}

// Проверка и нормализация входных данных. Бросает Error с понятным текстом.
function cleanInput(body, { partial = false } = {}) {
  const out = {};
  if (!partial || body.dateFrom !== undefined || body.dateTo !== undefined) {
    const from = String(body.dateFrom || body.date || "");
    const to = String(body.dateTo || from);
    if (!isValidKey(from) || !isValidKey(to)) throw new Error("Неверная дата");
    if (to < from) throw new Error("Дата окончания раньше даты начала");
    if (datesBetween(from, to).length > MAX_DAYS) throw new Error(`Период не больше ${MAX_DAYS} дней`);
    out.dateFrom = from;
    out.dateTo = to;
  }
  if (!partial || body.type !== undefined) {
    const type = String(body.type || "");
    if (!REASON_TYPES[type]) throw new Error("Неизвестный тип причины");
    out.type = type;
    if (body.excused === undefined) out.excused = REASON_TYPES[type].excused;
  }
  if (body.excused !== undefined) out.excused = !!body.excused;
  if (body.comment !== undefined) out.comment = String(body.comment || "").trim().slice(0, 300);
  if (body.confirmed !== undefined) out.confirmed = !!body.confirmed;
  return out;
}

async function createReason({ student, body, source, createdBy }) {
  const data = cleanInput(body);
  const doc = await AbsenceReason.create({
    ...data,
    student: student._id,
    group: student.group?._id || student.group,
    source,
    createdBy: String(createdBy || ""),
    confirmed: source !== "parent",
  });
  return serialize(doc.toObject ? doc.toObject() : doc);
}

async function updateReason(id, body) {
  const data = cleanInput(body, { partial: true });
  const doc = await AbsenceReason.findByIdAndUpdate(id, { $set: data }, { returnDocument: "after" }).lean();
  return doc ? serialize(doc) : null;
}

// Маршруты веб-панели (админ и кураторы). Подключать после app.use("/api/admin", requireAuth).
function registerAbsenceRoutes(app, { Student, canAccessGroup }) {
  const who = (req) => req.auth?.username || req.auth?.role || "";
  const wrap = (fn) => (req, res) =>
    fn(req, res).catch((e) => res.status(e.status || 400).json({ error: e.message }));

  app.get("/api/admin/absence-types", (_req, res) => res.json(REASON_TYPES));

  app.get(
    "/api/admin/absences",
    wrap(async (req, res) => {
      const { gid, from, to } = req.query;
      if (!canAccessGroup(req, gid)) return res.status(403).json({ error: "Нет доступа к группе" });
      if (!isValidKey(from) || !isValidKey(to)) return res.status(400).json({ error: "Неверный период" });
      const list = await findReasons({ group: gid }, from, to);
      res.json(list.map(serialize));
    }),
  );

  app.post(
    "/api/admin/absences",
    wrap(async (req, res) => {
      const student = await Student.findById(req.body.studentId).select("group");
      if (!student) return res.status(404).json({ error: "Студент не найден" });
      if (!canAccessGroup(req, student.group)) return res.status(403).json({ error: "Нет доступа к группе" });
      const source = req.auth.role === "admin" ? "admin" : "curator";
      res.json(await createReason({ student, body: req.body, source, createdBy: who(req) }));
    }),
  );

  const loadOwn = async (req, res) => {
    const r = await AbsenceReason.findById(req.params.id).lean();
    if (!r) {
      res.status(404).json({ error: "Запись не найдена" });
      return null;
    }
    if (!canAccessGroup(req, r.group)) {
      res.status(403).json({ error: "Нет доступа к группе" });
      return null;
    }
    return r;
  };

  app.put(
    "/api/admin/absences/:id",
    wrap(async (req, res) => {
      if (!(await loadOwn(req, res))) return;
      res.json(await updateReason(req.params.id, req.body));
    }),
  );

  app.delete(
    "/api/admin/absences/:id",
    wrap(async (req, res) => {
      if (!(await loadOwn(req, res))) return;
      await AbsenceReason.deleteOne({ _id: req.params.id });
      res.json({ ok: true });
    }),
  );
}

module.exports = {
  REASON_TYPES,
  AbsenceReason,
  serialize,
  findReasons,
  reasonsIndex,
  createReason,
  updateReason,
  cleanInput,
  registerAbsenceRoutes,
};
