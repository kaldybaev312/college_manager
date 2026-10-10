// Производственная практика: договор с организацией, дневник, характеристика.
// Подключение в server.js:
//   require("./practice")(app, { Student, Group, Curator, getSession, getTodayKey });
//
// Роли:
//   admin (вход администратора)   → старший мастер: видит все группы, меняет шаблоны и настройки
//   curator (вход куратора)        → мастер п/о: свои группы, распределение, договоры
//   студент                        → по личной ссылке /p/s/<код> или по ИНН на /p
//   руководитель от предприятия    → по ссылке /p/r/<код>, без пароля
//
// Документы собираются из Word-шаблонов (docxtemplater). Шаблоны по умолчанию лежат в templates/practice_*.docx,
// старший мастер может загрузить свои через страницу /practice → «Шаблоны» (хранятся в MongoDB, переживают деплой).
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const multer = require("multer");
const mongoose = require("mongoose");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");

const { Schema } = mongoose;
const OID = Schema.Types.ObjectId;

/* ============================ модели ============================ */
const OrgSchema = new Schema({
  name: String,
  head: String, // «в лице …» — должность и ФИО в родительном падеже
  address: String,
  phone: String,
  supName: String, // руководитель практики от предприятия
  supPos: String,
  supPhone: String,
  tbName: String, // наименование инструктажа по ТБ
  tbDate: String, // YYYY-MM-DD
  contract: { type: String, enum: ["draft", "printed", "returned"], default: "draft" },
  token: String,
});
const PracticeSchema = new Schema(
  {
    group: { type: OID, ref: "Group", index: true },
    groupName: String,
    course: Number,
    profession: String,
    start: String,
    end: String,
    days: [{ _id: false, date: String, hours: Number }],
    master: { curatorId: String, name: String, position: { type: String, default: "мастер п/о" }, phone: String },
    students: [{ _id: false, student: { type: OID, ref: "Student" }, fio: String, org: { type: OID, default: null }, token: String }],
    orgs: [OrgSchema],
    createdBy: String,
  },
  { timestamps: true },
);
PracticeSchema.index({ "students.token": 1 });
PracticeSchema.index({ "orgs.token": 1 });

const DiaryEntrySchema = new Schema(
  {
    practice: { type: OID, index: true },
    student: { type: OID, index: true },
    date: String,
    text: String,
    status: { type: String, enum: ["draft", "sent", "ok", "back"], default: "draft" },
    grade: Number,
    comment: String,
    media: [{ _id: false, fileId: OID, type: { type: String }, name: String, size: Number }],
  },
  { timestamps: true },
);
DiaryEntrySchema.index({ practice: 1, student: 1, date: 1 }, { unique: true });

const CharacteristicSchema = new Schema(
  {
    practice: { type: OID, index: true },
    student: OID,
    disc: Number,
    skill: Number,
    tb: Number,
    q: [String],
    works: String,
    extra: String,
    rank: String,
    grade: Number,
    done: Boolean,
  },
  { timestamps: true },
);
CharacteristicSchema.index({ practice: 1, student: 1 }, { unique: true });

const PracticeSettings = mongoose.models.PracticeSettings || mongoose.model("PracticeSettings", new Schema({ key: { type: String, unique: true }, data: Schema.Types.Mixed }, { timestamps: true }));
const Practice = mongoose.models.Practice || mongoose.model("Practice", PracticeSchema);
const DiaryEntry = mongoose.models.DiaryEntry || mongoose.model("DiaryEntry", DiaryEntrySchema);
const Characteristic = mongoose.models.Characteristic || mongoose.model("Characteristic", CharacteristicSchema);

/* ============================ настройки ============================ */
const DEFAULT_SETTINGS = {
  ministry: "Министерство науки, высшего образования и инноваций Кыргызской Республики",
  lyceumName: "Профессиональный лицей №3",
  directorGen: "Белялова Ш.А.", // в родительном падеже: «в лице директора … »
  lyceumAddress: "ул. Раззакова, 62",
  lyceumPhone: "0312 62-49-12",
  city: "г. Кант",
  hoursByCourse: { 1: 6, 2: 8, 3: 8 },
  tbDefault: "Инструктаж по ТБ и ППМ при работе с электрооборудованием",
  phrases: {
    disc: {
      5: "Добросовестно относился(ась) к работе, нарушений трудовой дисциплины не имел(а), пропусков без уважительных причин нет.",
      4: "Ответственно относился(ась) к работе, трудовую дисциплину соблюдал(а), имелись единичные опоздания.",
      3: "К работе относился(ась) удовлетворительно, имелись нарушения трудовой дисциплины.",
    },
    skill: {
      5: "Показал(а) хорошую теоретическую подготовку, самостоятельно и качественно выполнял(а) работы по профессии.",
      4: "Теоретическая подготовка достаточная, работы по профессии выполнял(а) с незначительной помощью наставника.",
      3: "Работы по профессии выполнял(а) под постоянным контролем наставника, допускал(а) ошибки.",
    },
    tb: {
      5: "Правила техники безопасности и охраны труда соблюдал(а) в полном объёме, средствами индивидуальной защиты пользовался(ась) правильно.",
      4: "Правила техники безопасности соблюдал(а), получал(а) отдельные замечания по использованию СИЗ.",
      3: "Допускал(а) нарушения правил техники безопасности, требовал(а) постоянного контроля.",
    },
  },
  qualities: ["дисциплинированность", "пунктуальность", "аккуратность", "инициативность", "ответственность", "умение работать в команде", "быстро осваивает новое", "самостоятельность"],
  ranks: ["2-го", "3-го", "4-го"],
};
function mergeDeep(base, over) {
  if (!over || typeof over !== "object" || Array.isArray(over)) return over === undefined ? base : over;
  const out = { ...base };
  for (const k of Object.keys(over)) out[k] = base && typeof base[k] === "object" && !Array.isArray(base[k]) ? mergeDeep(base[k], over[k]) : over[k];
  return out;
}
async function getSettings() {
  const d = await PracticeSettings.findOne({ key: "main" }).lean();
  return mergeDeep(DEFAULT_SETTINGS, d?.data || {});
}

/* ============================ шаблоны ============================ */
const TEMPLATES = {
  contract: { file: "practice_contract.docx", title: "Договор о практике" },
  diary: { file: "practice_diary.docx", title: "Дневник практики" },
  characteristic: { file: "practice_characteristic.docx", title: "Характеристика" },
};
// Метки, которые можно использовать в шаблоне — показываются на странице «Шаблоны»
const TEMPLATE_TAGS = {
  contract: "{directorGen} {orgName} {orgHead} {startText} {endText} {studentsText} {profession} {masterName} {masterPos} {masterPhone} {supName} {supPos} {supPhone} {orgAddress} {orgPhone} {lyceumAddress} {lyceumPhone} {group} · таблица ТБ: {#tb}{n} {fio} {tbName} {tbDate}{/tb}",
  diary: "{ministry} {lyceumName} {studentName} {group} {course} {profession} {orgName} {orgAddress} {startText} {endText} {daysCount} {hoursTotal} {supName} {supPos} {supShort} {masterName} {masterShort} {city} {year} {avgGrade} {code} · строки дней: {#days}{n} {date} {hours} {text} {grade} {media}{/days}",
  characteristic: "{studentName} {group} {course} {profession} {orgName} {startText} {endText} {hoursTotal} {discText} {skillText} {worksText} {tbText} {qualitiesText} {#extra}{extra}{/extra} {conclusionText} {rank} {grade} {gradeText} {supName} {supPos} {supShort} {code}",
};
let _tplBucket = null, _mediaBucket = null;
const tplBucket = () => (_tplBucket ||= new mongoose.mongo.GridFSBucket(mongoose.connection.db, { bucketName: "practiceTemplates" }));
const mediaBucket = () => (_mediaBucket ||= new mongoose.mongo.GridFSBucket(mongoose.connection.db, { bucketName: "practiceMedia" }));

function readStream(stream) {
  return new Promise((res, rej) => { const parts = []; stream.on("data", (c) => parts.push(c)); stream.on("end", () => res(Buffer.concat(parts))); stream.on("error", rej); });
}
async function customTemplate(name) {
  const f = await tplBucket().find({ filename: name }).sort({ uploadDate: -1 }).limit(1).toArray();
  return f[0] || null;
}
async function loadTemplate(name) {
  const c = await customTemplate(name);
  if (c) return readStream(tplBucket().openDownloadStream(c._id));
  return fs.readFileSync(path.join(__dirname, "templates", TEMPLATES[name].file));
}
function renderDocx(buf, data) {
  const doc = new Docxtemplater(new PizZip(buf), { paragraphLoop: true, linebreaks: true, nullGetter: () => "" });
  doc.render(data);
  return doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });
}
function templateErrorText(e) {
  const list = e?.properties?.errors;
  if (Array.isArray(list) && list.length) return list.slice(0, 5).map((x) => x.properties?.explanation || x.message).join("; ");
  return e?.message || String(e);
}
function sendDocx(res, buf, filename) {
  const ascii = filename.replace(/[^\x20-\x7e]/g, "_");
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
  res.setHeader("Content-Disposition", `attachment; filename="${ascii}"; filename*=UTF-8''${encodeURIComponent(filename)}`);
  res.send(buf);
}

/* ============================ утилиты ============================ */
const MG = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"];
const longDate = (d) => { if (!d) return ""; const [y, m, dd] = d.split("-"); return `${+dd} ${MG[m - 1]} ${y} г.`; };
const fmt = (d) => { if (!d) return ""; const [y, m, dd] = d.split("-"); return `${dd}.${m}.${y}`; };
const short = (fio = "") => { const [f, i, o] = String(fio).trim().split(/\s+/); return `${f || ""} ${i ? i[0] + "." : ""}${o ? o[0] + "." : ""}`.trim(); };
const GW = { 2: "неудовлетворительно", 3: "удовлетворительно", 4: "хорошо", 5: "отлично" };
const token = () => crypto.randomBytes(9).toString("base64url");
const isDate = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || ""));
const safe = (s, max = 300) => String(s ?? "").trim().slice(0, max);
const iso = (d) => `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
function buildDays(start, end, hours, sat) {
  const out = []; const d = new Date(start + "T00:00:00Z"); const e = new Date(end + "T00:00:00Z");
  while (d <= e && out.length < 250) { const g = d.getUTCDay(); if (g !== 0 && (sat || g !== 6)) out.push({ date: iso(d), hours }); d.setUTCDate(d.getUTCDate() + 1); }
  return out;
}
const sumHours = (p) => p.days.reduce((a, d) => a + Number(d.hours || 0), 0);
const orgById = (p, id) => (id ? p.orgs.find((o) => String(o._id) === String(id)) : null);
const studentRow = (p, sid) => p.students.find((s) => String(s.student) === String(sid));
const codeOf = (p, org, sid) => `${p.groupName}-${(org?.token || "").slice(0, 4).toUpperCase()}-${String(sid).slice(-4).toUpperCase()}`;

/* ============================ данные для документов ============================ */
function contractData(p, org, S) {
  const st = p.students.filter((s) => String(s.org) === String(org._id));
  return {
    directorGen: S.directorGen, lyceumAddress: S.lyceumAddress, lyceumPhone: S.lyceumPhone, group: p.groupName,
    orgName: org.name, orgHead: org.head, orgAddress: org.address, orgPhone: org.phone,
    startText: longDate(p.start), endText: longDate(p.end), profession: p.profession,
    studentsText: st.map((s) => s.fio).join(", "),
    masterName: p.master?.name || "", masterPos: p.master?.position || "мастер п/о", masterPhone: p.master?.phone || "",
    supName: org.supName, supPos: org.supPos, supPhone: org.supPhone,
    tb: st.map((s, i) => ({ n: i + 1, fio: s.fio, tbName: org.tbName || S.tbDefault, tbDate: fmt(org.tbDate || p.start) })),
  };
}
function diaryData(p, srow, org, entries, S) {
  const E = Object.fromEntries(entries.map((e) => [e.date, e]));
  const days = p.days.map((d, i) => {
    const e = E[d.date]; const show = e && (e.status === "ok" || e.status === "sent");
    return { n: i + 1, date: fmt(d.date), hours: d.hours, text: show ? e.text : "", grade: e?.status === "ok" ? e.grade : "", media: show && e.media?.length ? `фото/видео: ${e.media.length}` : "" };
  });
  const g = entries.filter((e) => e.status === "ok" && e.grade).map((e) => e.grade);
  return {
    ministry: S.ministry, lyceumName: S.lyceumName, studentName: srow.fio, group: p.groupName, course: p.course, profession: p.profession,
    orgName: org?.name || "", orgAddress: org?.address || "", startText: longDate(p.start), endText: longDate(p.end),
    daysCount: p.days.length, hoursTotal: sumHours(p), supName: org?.supName || "", supPos: org?.supPos || "", supShort: short(org?.supName),
    masterName: p.master?.name || "", masterShort: short(p.master?.name), year: (p.start || "").slice(0, 4), city: S.city,
    days, avgGrade: g.length ? (g.reduce((a, b) => a + b, 0) / g.length).toFixed(1) : "", code: codeOf(p, org, srow.student),
  };
}
function charData(p, srow, org, ch, S) {
  ch = ch || {};
  const ph = S.phrases || {};
  const rankText = ch.rank ? ` Рекомендуется присвоение квалификации «${p.profession}» ${ch.rank} разряда.` : "";
  return {
    studentName: srow.fio, group: p.groupName, course: p.course, profession: p.profession, orgName: org?.name || "",
    startText: longDate(p.start), endText: longDate(p.end), hoursTotal: sumHours(p),
    discText: ph.disc?.[ch.disc] || "", skillText: ph.skill?.[ch.skill] || "", tbText: ph.tb?.[ch.tb] || "",
    worksText: ch.works ? ` Освоил(а): ${ch.works}.` : "", qualitiesText: (ch.q || []).join(", ") || "—",
    extra: ch.extra || "", conclusionText: `Программа производственной практики выполнена${ch.grade && ch.grade <= 2 ? " не в полном объёме" : ""}.${rankText}`,
    rank: ch.rank || "", grade: ch.grade || "", gradeText: ch.grade ? `${ch.grade} (${GW[ch.grade]})` : "",
    supName: org?.supName || "", supPos: org?.supPos || "", supShort: short(org?.supName), code: codeOf(p, org, srow.student) + "-Х",
  };
}

/* ============================ модуль ============================ */
module.exports = function practiceModule(app, { Student, Group, Curator, getSession, getTodayKey }) {
  const memUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 30 * 1024 * 1024 } });
  const tplUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 5 * 1024 * 1024 } });
  const wrap = (fn) => (req, res) => Promise.resolve(fn(req, res)).catch((e) => { console.error("practice:", e); if (!res.headersSent) res.status(500).json({ error: e.message }); });

  /* ---------- доступ сотрудников ---------- */
  function staff(req, res, next) {
    const s = getSession(req);
    if (!s?.role) return res.status(401).json({ error: "Войдите в систему" });
    req.auth = s; next();
  }
  function admin(req, res, next) {
    if (req.auth?.role !== "admin") return res.status(403).json({ error: "Только для старшего мастера (администратора)" });
    next();
  }
  const groupIdsOf = (s) => (s.role === "admin" ? null : (Array.isArray(s.groupIds) ? s.groupIds : s.groupId ? [s.groupId] : []).map(String));
  function canSee(s, p) { const ids = groupIdsOf(s); return ids === null || ids.includes(String(p.group)); }
  async function loadPractice(req, res) {
    if (!mongoose.isValidObjectId(req.params.id)) { res.status(404).json({ error: "Практика не найдена" }); return null; }
    const p = await Practice.findById(req.params.id);
    if (!p) { res.status(404).json({ error: "Практика не найдена" }); return null; }
    if (!canSee(req.auth, p)) { res.status(403).json({ error: "Нет доступа к этой группе" }); return null; }
    return p;
  }
  async function statsFor(ids) {
    const rows = await DiaryEntry.aggregate([{ $match: { practice: { $in: ids } } }, { $group: { _id: { p: "$practice", s: "$student", st: "$status" }, n: { $sum: 1 } } }]);
    const chars = await Characteristic.find({ practice: { $in: ids }, done: true }).select("practice student grade").lean();
    const out = {};
    for (const r of rows) { const k = `${r._id.p}:${r._id.s}`; (out[k] ||= { ok: 0, sent: 0, back: 0, draft: 0 })[r._id.st] = r.n; }
    for (const c of chars) { const k = `${c.practice}:${c.student}`; (out[k] ||= { ok: 0, sent: 0, back: 0, draft: 0 }).char = c.grade || true; }
    return out;
  }
  function practiceJSON(p, st) {
    const o = p.toObject ? p.toObject() : p;
    return {
      ...o, hoursTotal: sumHours(o),
      students: o.students.map((s) => ({ ...s, stats: st[`${o._id}:${s.student}`] || { ok: 0, sent: 0, back: 0, draft: 0 } })),
    };
  }

  /* ---------- страницы ---------- */
  app.get(["/practice", "/practice.html"], (req, res) => {
    if (!getSession(req)?.role) return res.redirect("/login");
    res.sendFile(path.join(__dirname, "public", "practice.html"));
  });
  app.get(["/p", "/p/s/:t", "/p/r/:t"], (req, res) => res.sendFile(path.join(__dirname, "public", "p.html")));

  /* ---------- сотрудники: справочники ---------- */
  app.get("/api/practice/me", staff, wrap(async (req, res) => {
    const s = req.auth; let name = s.username;
    if (s.curatorId && mongoose.isValidObjectId(s.curatorId)) { const c = await Curator.findById(s.curatorId).select("fullName").lean(); if (c?.fullName) name = c.fullName; }
    let phone = "";
    const last = await Practice.findOne(s.curatorId ? { "master.curatorId": s.curatorId } : { createdBy: s.username }).sort({ createdAt: -1 }).select("master").lean();
    if (last?.master) { phone = last.master.phone || ""; if (last.master.name) name = last.master.name; }
    res.json({ role: s.role, username: s.username, curatorId: s.curatorId || null, name, phone, groupNames: s.groupNames || [], today: getTodayKey() });
  }));
  app.get("/api/practice/groups", staff, wrap(async (req, res) => {
    const ids = groupIdsOf(req.auth);
    res.json(await Group.find(ids === null ? {} : { _id: { $in: ids } }).sort({ name: 1 }).select("name profRu duration").lean());
  }));
  app.get("/api/practice/groups/:gid/students", staff, wrap(async (req, res) => {
    const ids = groupIdsOf(req.auth);
    if (ids !== null && !ids.includes(String(req.params.gid))) return res.status(403).json({ error: "Нет доступа к этой группе" });
    res.json(await Student.find({ group: req.params.gid }).sort({ fio: 1 }).select("fio").lean());
  }));

  /* ---------- настройки и шаблоны (старший мастер) ---------- */
  app.get("/api/practice/settings", staff, wrap(async (req, res) => res.json(await getSettings())));
  app.put("/api/practice/settings", staff, admin, wrap(async (req, res) => {
    const data = mergeDeep(DEFAULT_SETTINGS, req.body || {});
    await PracticeSettings.updateOne({ key: "main" }, { $set: { data } }, { upsert: true });
    res.json(data);
  }));
  app.get("/api/practice/templates", staff, wrap(async (req, res) => {
    const out = [];
    for (const [name, t] of Object.entries(TEMPLATES)) { const c = await customTemplate(name); out.push({ name, title: t.title, custom: !!c, uploadedAt: c?.uploadDate || null, tags: TEMPLATE_TAGS[name] }); }
    res.json(out);
  }));
  app.get("/api/practice/templates/:name", staff, wrap(async (req, res) => {
    const t = TEMPLATES[req.params.name]; if (!t) return res.status(404).json({ error: "Нет такого шаблона" });
    sendDocx(res, await loadTemplate(req.params.name), `Шаблон — ${t.title}.docx`);
  }));
  app.post("/api/practice/templates/:name", staff, admin, tplUpload.single("file"), wrap(async (req, res) => {
    const name = req.params.name; if (!TEMPLATES[name]) return res.status(404).json({ error: "Нет такого шаблона" });
    if (!req.file) return res.status(400).json({ error: "Выберите файл .docx" });
    try {
      const xml = new PizZip(req.file.buffer).file("word/document.xml")?.asText() || "";
      if (!/\{[^{}]*\}/.test(xml.replace(/<[^>]+>/g, ""))) return res.status(400).json({ error: "В файле нет ни одной метки в фигурных скобках — похоже, это готовый документ, а не шаблон. Скачайте шаблон и правьте его." });
      renderDocx(req.file.buffer, {});
    }
    catch (e) { return res.status(400).json({ error: "Шаблон не принят: " + templateErrorText(e) + ". Проверьте, что каждая метка в фигурных скобках набрана целиком, без смены шрифта внутри." }); }
    const old = await tplBucket().find({ filename: name }).toArray();
    await new Promise((ok, fail) => { const w = tplBucket().openUploadStream(name, { metadata: { by: req.auth.username } }); w.on("finish", ok); w.on("error", fail); w.end(req.file.buffer); });
    for (const f of old) await tplBucket().delete(f._id).catch(() => {});
    res.json({ ok: true });
  }));
  app.delete("/api/practice/templates/:name", staff, admin, wrap(async (req, res) => {
    for (const f of await tplBucket().find({ filename: req.params.name }).toArray()) await tplBucket().delete(f._id).catch(() => {});
    res.json({ ok: true });
  }));

  /* ---------- практики ---------- */
  app.get("/api/practice/practices", staff, wrap(async (req, res) => {
    const ids = groupIdsOf(req.auth);
    const list = await Practice.find(ids === null ? {} : { group: { $in: ids } }).sort({ start: -1 });
    const st = await statsFor(list.map((p) => p._id));
    res.json(list.map((p) => practiceJSON(p, st)));
  }));
  app.post("/api/practice/practices", staff, wrap(async (req, res) => {
    const b = req.body || {};
    const ids = groupIdsOf(req.auth);
    if (!mongoose.isValidObjectId(b.groupId)) return res.status(400).json({ error: "Выберите группу" });
    if (ids !== null && !ids.includes(String(b.groupId))) return res.status(403).json({ error: "Нет доступа к этой группе" });
    if (!isDate(b.start) || !isDate(b.end) || b.end < b.start) return res.status(400).json({ error: "Проверьте даты практики" });
    const g = await Group.findById(b.groupId).lean(); if (!g) return res.status(404).json({ error: "Группа не найдена" });
    const studs = await Student.find({ group: g._id, _id: { $in: (b.studentIds || []).filter(mongoose.isValidObjectId) } }).select("fio").sort({ fio: 1 }).lean();
    if (!studs.length) return res.status(400).json({ error: "Отметьте хотя бы одного студента" });
    const S = await getSettings(); const course = Number(b.course) || 1;
    const hours = Number(b.hours) || S.hoursByCourse?.[course] || 6;
    const p = await Practice.create({
      group: g._id, groupName: g.name, course, profession: safe(b.profession) || g.profRu || "",
      start: b.start, end: b.end, days: buildDays(b.start, b.end, hours, b.sat !== false),
      master: { curatorId: req.auth.curatorId || null, name: safe(b.master?.name, 120), position: safe(b.master?.position, 80) || "мастер п/о", phone: safe(b.master?.phone, 40) },
      students: studs.map((s) => ({ student: s._id, fio: s.fio, org: null, token: token() })), orgs: [], createdBy: req.auth.username,
    });
    res.json(practiceJSON(p, {}));
  }));
  app.put("/api/practice/practices/:id", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return; const b = req.body || {};
    if (b.profession !== undefined) p.profession = safe(b.profession);
    if (b.course !== undefined) p.course = Number(b.course) || p.course;
    if (b.master) p.master = { ...p.master, name: safe(b.master.name, 120), position: safe(b.master.position, 80) || "мастер п/о", phone: safe(b.master.phone, 40) };
    if (Array.isArray(b.days)) {
      const days = b.days.filter((d) => isDate(d.date)).map((d) => ({ date: d.date, hours: Math.max(1, Math.min(12, Number(d.hours) || 1)) })).sort((a, z) => a.date.localeCompare(z.date));
      if (!days.length) return res.status(400).json({ error: "В практике должен остаться хотя бы один день" });
      p.days = days; p.start = days[0].date; p.end = days[days.length - 1].date;
    }
    if (b.rebuild && isDate(b.rebuild.start) && isDate(b.rebuild.end) && b.rebuild.end >= b.rebuild.start) {
      p.days = buildDays(b.rebuild.start, b.rebuild.end, Number(b.rebuild.hours) || p.days[0]?.hours || 6, b.rebuild.sat !== false);
      p.start = b.rebuild.start; p.end = b.rebuild.end;
    }
    await p.save(); res.json(practiceJSON(p, await statsFor([p._id])));
  }));
  app.delete("/api/practice/practices/:id", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return;
    const entries = await DiaryEntry.find({ practice: p._id }).select("media").lean();
    for (const e of entries) for (const m of e.media || []) await mediaBucket().delete(m.fileId).catch(() => {});
    await DiaryEntry.deleteMany({ practice: p._id }); await Characteristic.deleteMany({ practice: p._id }); await p.deleteOne();
    res.json({ ok: true });
  }));

  /* студенты в практике */
  app.post("/api/practice/practices/:id/students", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return;
    const have = new Set(p.students.map((s) => String(s.student)));
    const add = await Student.find({ group: p.group, _id: { $in: (req.body?.studentIds || []).filter(mongoose.isValidObjectId) } }).select("fio").lean();
    for (const s of add) if (!have.has(String(s._id))) p.students.push({ student: s._id, fio: s.fio, org: null, token: token() });
    p.students.sort((a, z) => a.fio.localeCompare(z.fio, "ru"));
    await p.save(); res.json(practiceJSON(p, await statsFor([p._id])));
  }));
  app.delete("/api/practice/practices/:id/students/:sid", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return;
    p.students = p.students.filter((s) => String(s.student) !== req.params.sid);
    await p.save(); res.json(practiceJSON(p, await statsFor([p._id])));
  }));
  app.put("/api/practice/practices/:id/students/:sid/org", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return;
    const row = studentRow(p, req.params.sid); if (!row) return res.status(404).json({ error: "Студент не найден в практике" });
    const oid = req.body?.orgId || null; if (oid && !orgById(p, oid)) return res.status(404).json({ error: "Организация не найдена" });
    row.org = oid; await p.save(); res.json(practiceJSON(p, await statsFor([p._id])));
  }));

  /* организации */
  const ORG_FIELDS = ["name", "head", "address", "phone", "supName", "supPos", "supPhone", "tbName", "tbDate"];
  const pickOrg = (b) => Object.fromEntries(ORG_FIELDS.filter((k) => b[k] !== undefined).map((k) => [k, safe(b[k])]));
  app.post("/api/practice/practices/:id/orgs", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return; const o = pickOrg(req.body || {});
    if (!o.name || !o.supName) return res.status(400).json({ error: "Нужны название организации и ФИО руководителя практики" });
    p.orgs.push({ ...o, contract: "draft", token: token() }); await p.save(); res.json(practiceJSON(p, await statsFor([p._id])));
  }));
  app.put("/api/practice/practices/:id/orgs/:oid", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return; const o = orgById(p, req.params.oid);
    if (!o) return res.status(404).json({ error: "Организация не найдена" });
    Object.assign(o, pickOrg(req.body || {}));
    if (["draft", "printed", "returned"].includes(req.body?.contract)) o.contract = req.body.contract;
    if (req.body?.newToken) o.token = token();
    await p.save(); res.json(practiceJSON(p, await statsFor([p._id])));
  }));
  app.delete("/api/practice/practices/:id/orgs/:oid", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return;
    if (p.students.some((s) => String(s.org) === req.params.oid)) return res.status(400).json({ error: "Сначала переведите студентов в другую организацию" });
    p.orgs = p.orgs.filter((o) => String(o._id) !== req.params.oid); await p.save(); res.json(practiceJSON(p, await statsFor([p._id])));
  }));

  /* документы для сотрудников */
  async function docOr400(res, name, data, filename) {
    let buf; try { buf = renderDocx(await loadTemplate(name), data); }
    catch (e) { return res.status(400).json({ error: `Ошибка в шаблоне «${TEMPLATES[name].title}»: ${templateErrorText(e)}` }); }
    sendDocx(res, buf, filename);
  }
  app.get("/api/practice/practices/:id/orgs/:oid/contract.docx", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return; const o = orgById(p, req.params.oid);
    if (!o) return res.status(404).json({ error: "Организация не найдена" });
    if (o.contract === "draft") { o.contract = "printed"; await p.save(); }
    await docOr400(res, "contract", contractData(p, o, await getSettings()), `Договор ${p.groupName} — ${o.name}.docx`);
  }));
  app.get("/api/practice/practices/:id/students/:sid/entries", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return;
    res.json(await DiaryEntry.find({ practice: p._id, student: req.params.sid }).lean());
  }));
  app.get("/api/practice/practices/:id/students/:sid/diary.docx", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return; const row = studentRow(p, req.params.sid);
    if (!row) return res.status(404).json({ error: "Студент не найден" });
    const entries = await DiaryEntry.find({ practice: p._id, student: row.student }).lean();
    await docOr400(res, "diary", diaryData(p, row, orgById(p, row.org), entries, await getSettings()), `Дневник — ${short(row.fio)}.docx`);
  }));
  app.get("/api/practice/practices/:id/students/:sid/characteristic.docx", staff, wrap(async (req, res) => {
    const p = await loadPractice(req, res); if (!p) return; const row = studentRow(p, req.params.sid);
    if (!row) return res.status(404).json({ error: "Студент не найден" });
    const ch = await Characteristic.findOne({ practice: p._id, student: row.student }).lean();
    await docOr400(res, "characteristic", charData(p, row, orgById(p, row.org), ch, await getSettings()), `Характеристика — ${short(row.fio)}.docx`);
  }));

  /* ============================ студент по ссылке ============================ */
  async function byStudentToken(t) {
    const p = await Practice.findOne({ "students.token": String(t || "") });
    if (!p) return null;
    const row = p.students.find((s) => s.token === t);
    return { p, row, org: orgById(p, row.org) };
  }
  function publicPractice(p) {
    return { id: p._id, group: p.groupName, course: p.course, profession: p.profession, start: p.start, end: p.end, days: p.days, hoursTotal: sumHours(p), master: { name: p.master?.name, phone: p.master?.phone } };
  }
  const publicOrg = (o) => (o ? { id: o._id, name: o.name, address: o.address, supName: o.supName, supPos: o.supPos, supPhone: o.supPhone } : null);

  app.post("/api/p/login", wrap(async (req, res) => {
    const inn = String(req.body?.inn || "").replace(/\D/g, "");
    if (inn.length < 12) return res.status(400).json({ error: "Введите ИНН полностью (14 цифр)" });
    const year = String(req.body?.year || "").replace(/\D/g, "");
    const studs = await Student.find({ inn: new RegExp(inn + "$") }).select("_id birthDate").lean();
    const st = studs.find((s) => { const y = String(s.birthDate || "").match(/(19|20)\d{2}/)?.[0]; return !y || y === year; });
    if (!st) return res.status(404).json({ error: studs.length ? "Год рождения не совпадает с данными лицея" : "Студент с таким ИНН не найден" });
    const p = await Practice.findOne({ "students.student": st._id }).sort({ start: -1 });
    if (!p) return res.status(404).json({ error: "Вы пока не направлены на практику. Обратитесь к мастеру п/о." });
    res.json({ token: p.students.find((s) => String(s.student) === String(st._id)).token });
  }));

  app.get("/api/p/s/:t", wrap(async (req, res) => {
    const x = await byStudentToken(req.params.t); if (!x) return res.status(404).json({ error: "Ссылка недействительна. Попросите у мастера п/о новую." });
    const entries = await DiaryEntry.find({ practice: x.p._id, student: x.row.student }).lean();
    const ch = await Characteristic.findOne({ practice: x.p._id, student: x.row.student, done: true }).select("grade").lean();
    res.json({ role: "student", today: getTodayKey(), practice: publicPractice(x.p), student: { id: x.row.student, fio: x.row.fio }, org: publicOrg(x.org), entries, characteristic: ch ? { grade: ch.grade } : null });
  }));
  async function editableEntry(req, res) {
    const x = await byStudentToken(req.params.t); if (!x) { res.status(404).json({ error: "Ссылка недействительна" }); return null; }
    if (!x.org) { res.status(400).json({ error: "Мастер п/о ещё не направил вас в организацию" }); return null; }
    const date = req.params.date;
    if (!x.p.days.some((d) => d.date === date)) { res.status(400).json({ error: "Этого дня нет в практике" }); return null; }
    if (date > getTodayKey()) { res.status(400).json({ error: "Этот день ещё не наступил" }); return null; }
    let e = await DiaryEntry.findOne({ practice: x.p._id, student: x.row.student, date });
    if (e && (e.status === "ok" || e.status === "sent")) { res.status(400).json({ error: e.status === "ok" ? "Запись уже подтверждена" : "Запись на проверке у руководителя" }); return null; }
    if (!e) e = new DiaryEntry({ practice: x.p._id, student: x.row.student, date, text: "", status: "draft", media: [] });
    return { x, e };
  }
  app.put("/api/p/s/:t/entries/:date", wrap(async (req, res) => {
    const r = await editableEntry(req, res); if (!r) return;
    r.e.text = safe(req.body?.text, 3000);
    if (req.body?.send) { if (!r.e.text) return res.status(400).json({ error: "Опишите выполненную работу" }); r.e.status = "sent"; r.e.comment = ""; }
    await r.e.save(); res.json(r.e);
  }));
  app.post("/api/p/s/:t/entries/:date/media", memUpload.single("file"), wrap(async (req, res) => {
    const r = await editableEntry(req, res); if (!r) return;
    const f = req.file; if (!f) return res.status(400).json({ error: "Файл не получен" });
    const type = f.mimetype.startsWith("video/") ? "video" : f.mimetype.startsWith("image/") ? "image" : null;
    if (!type) return res.status(400).json({ error: "Можно прикрепить только фото или видео" });
    if ((r.e.media || []).length >= 6) return res.status(400).json({ error: "Не больше 6 файлов на один день" });
    const id = await new Promise((ok, fail) => {
      const w = mediaBucket().openUploadStream(safe(f.originalname, 120) || type, { contentType: f.mimetype, metadata: { practice: r.x.p._id, student: r.x.row.student, date: r.e.date } });
      w.on("finish", () => ok(w.id)); w.on("error", fail); w.end(f.buffer);
    });
    r.e.media.push({ fileId: id, type, name: safe(f.originalname, 120), size: f.size });
    await r.e.save(); res.json(r.e);
  }));
  app.delete("/api/p/s/:t/entries/:date/media/:fid", wrap(async (req, res) => {
    const r = await editableEntry(req, res); if (!r) return;
    const before = r.e.media.length;
    r.e.media = r.e.media.filter((m) => String(m.fileId) !== req.params.fid);
    if (r.e.media.length !== before) await mediaBucket().delete(new mongoose.Types.ObjectId(req.params.fid)).catch(() => {});
    await r.e.save(); res.json(r.e);
  }));
  app.get("/api/p/s/:t/diary.docx", wrap(async (req, res) => {
    const x = await byStudentToken(req.params.t); if (!x) return res.status(404).json({ error: "Ссылка недействительна" });
    const entries = await DiaryEntry.find({ practice: x.p._id, student: x.row.student }).lean();
    await docOr400(res, "diary", diaryData(x.p, x.row, x.org, entries, await getSettings()), `Дневник — ${short(x.row.fio)}.docx`);
  }));
  app.get("/api/p/s/:t/characteristic.docx", wrap(async (req, res) => {
    const x = await byStudentToken(req.params.t); if (!x) return res.status(404).json({ error: "Ссылка недействительна" });
    const ch = await Characteristic.findOne({ practice: x.p._id, student: x.row.student, done: true }).lean();
    if (!ch) return res.status(400).json({ error: "Руководитель ещё не заполнил характеристику" });
    await docOr400(res, "characteristic", charData(x.p, x.row, x.org, ch, await getSettings()), `Характеристика — ${short(x.row.fio)}.docx`);
  }));

  /* ============================ руководитель по ссылке ============================ */
  async function bySupToken(t) {
    const p = await Practice.findOne({ "orgs.token": String(t || "") });
    if (!p) return null;
    const org = p.orgs.find((o) => o.token === t);
    return { p, org, students: p.students.filter((s) => String(s.org) === String(org._id)) };
  }
  app.get("/api/p/r/:t", wrap(async (req, res) => {
    const x = await bySupToken(req.params.t); if (!x) return res.status(404).json({ error: "Ссылка недействительна. Попросите у мастера п/о новую." });
    const ids = x.students.map((s) => s.student);
    const entries = await DiaryEntry.find({ practice: x.p._id, student: { $in: ids } }).lean();
    const chars = await Characteristic.find({ practice: x.p._id, student: { $in: ids } }).lean();
    res.json({
      role: "supervisor", today: getTodayKey(), practice: publicPractice(x.p), org: publicOrg(x.org), settings: await getSettings(),
      students: x.students.map((s) => ({ id: s.student, fio: s.fio, entries: entries.filter((e) => String(e.student) === String(s.student)), characteristic: chars.find((c) => String(c.student) === String(s.student)) || null })),
    });
  }));
  app.put("/api/p/r/:t/entries/:eid", wrap(async (req, res) => {
    const x = await bySupToken(req.params.t); if (!x) return res.status(404).json({ error: "Ссылка недействительна" });
    if (!mongoose.isValidObjectId(req.params.eid)) return res.status(404).json({ error: "Запись не найдена" });
    const e = await DiaryEntry.findOne({ _id: req.params.eid, practice: x.p._id, student: { $in: x.students.map((s) => s.student) } });
    if (!e) return res.status(404).json({ error: "Запись не найдена" });
    const b = req.body || {};
    if (b.action === "ok") { const g = Number(b.grade); if (![2, 3, 4, 5].includes(g)) return res.status(400).json({ error: "Поставьте оценку от 2 до 5" }); e.status = "ok"; e.grade = g; e.comment = ""; }
    else if (b.action === "back") { if (!safe(b.comment)) return res.status(400).json({ error: "Напишите замечание" }); e.status = "back"; e.grade = null; e.comment = safe(b.comment, 500); }
    else return res.status(400).json({ error: "Неизвестное действие" });
    await e.save(); res.json(e);
  }));
  app.put("/api/p/r/:t/characteristic/:sid", wrap(async (req, res) => {
    const x = await bySupToken(req.params.t); if (!x) return res.status(404).json({ error: "Ссылка недействительна" });
    const row = x.students.find((s) => String(s.student) === req.params.sid); if (!row) return res.status(404).json({ error: "Студент не найден" });
    const b = req.body || {}; const lv = (v) => ([3, 4, 5].includes(Number(v)) ? Number(v) : 5);
    const grade = Number(b.grade); if (![2, 3, 4, 5].includes(grade)) return res.status(400).json({ error: "Поставьте итоговую оценку" });
    const data = { disc: lv(b.disc), skill: lv(b.skill), tb: lv(b.tb), q: (b.q || []).map((s) => safe(s, 80)).slice(0, 20), works: safe(b.works, 400), extra: safe(b.extra, 1500), rank: safe(b.rank, 20), grade, done: true };
    const ch = await Characteristic.findOneAndUpdate({ practice: x.p._id, student: row.student }, { $set: data }, { upsert: true, returnDocument: "after" });
    res.json(ch);
  }));
  app.get("/api/p/r/:t/students/:sid/:doc.docx", wrap(async (req, res) => {
    const x = await bySupToken(req.params.t); if (!x) return res.status(404).json({ error: "Ссылка недействительна" });
    const row = x.students.find((s) => String(s.student) === req.params.sid); if (!row) return res.status(404).json({ error: "Студент не найден" });
    const S = await getSettings();
    if (req.params.doc === "diary") {
      const entries = await DiaryEntry.find({ practice: x.p._id, student: row.student }).lean();
      return docOr400(res, "diary", diaryData(x.p, row, x.org, entries, S), `Дневник — ${short(row.fio)}.docx`);
    }
    if (req.params.doc === "characteristic") {
      const ch = await Characteristic.findOne({ practice: x.p._id, student: row.student }).lean();
      return docOr400(res, "characteristic", charData(x.p, row, x.org, ch, S), `Характеристика — ${short(row.fio)}.docx`);
    }
    res.status(404).json({ error: "Нет такого документа" });
  }));

  /* ============================ фото и видео ============================ */
  app.get("/api/p/media/:fid", wrap(async (req, res) => {
    if (!mongoose.isValidObjectId(req.params.fid)) return res.status(404).end();
    const fid = new mongoose.Types.ObjectId(req.params.fid);
    const e = await DiaryEntry.findOne({ "media.fileId": fid }).select("practice student").lean();
    if (!e) return res.status(404).end();
    const p = await Practice.findById(e.practice).select("group students orgs").lean();
    const t = String(req.query.t || "");
    const row = p?.students.find((s) => String(s.student) === String(e.student));
    const s = getSession(req);
    const allowed = (t && row?.token === t) || (t && p?.orgs.some((o) => o.token === t && String(o._id) === String(row?.org))) || (s?.role && canSee(s, p));
    if (!allowed) return res.status(403).end();
    const file = (await mediaBucket().find({ _id: fid }).toArray())[0]; if (!file) return res.status(404).end();
    const size = file.length; const range = /bytes=(\d*)-(\d*)/.exec(req.headers.range || "");
    res.setHeader("Content-Type", file.contentType || "application/octet-stream");
    res.setHeader("Accept-Ranges", "bytes"); res.setHeader("Cache-Control", "private, max-age=86400");
    if (range) {
      const start = range[1] ? Number(range[1]) : 0; const end = range[2] ? Math.min(Number(range[2]), size - 1) : size - 1;
      if (start >= size || start > end) { res.status(416).setHeader("Content-Range", `bytes */${size}`); return res.end(); }
      res.status(206); res.setHeader("Content-Range", `bytes ${start}-${end}/${size}`); res.setHeader("Content-Length", end - start + 1);
      return mediaBucket().openDownloadStream(fid, { start, end: end + 1 }).pipe(res);
    }
    res.setHeader("Content-Length", size);
    mediaBucket().openDownloadStream(fid).pipe(res);
  }));

  console.log("✅ Модуль практики подключён: /practice, /p");
};
