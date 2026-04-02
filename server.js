require("dotenv").config(); // загружаем .env
const express = require("express");
const mongoose = require("mongoose");
const path = require("path");
const fs = require("fs");
const crypto = require("crypto");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const QRCode = require("qrcode");
const multer = require("multer");
const xlsx = require("xlsx");
const TelegramBot = require("node-telegram-bot-api");
const cron = require("node-cron");

const app = express();
const upload = multer({ dest: "uploads/" });
if (!fs.existsSync("uploads")) fs.mkdirSync("uploads");
app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));

const ADMIN_LOGIN = process.env.ADMIN_LOGIN || "admin";
const ADMIN_PASS = process.env.ADMIN_PASS || "pl3-2026";
const AUTH_COOKIE = "pl3_auth";

function getCookie(req, name) {
  const raw = req.headers.cookie;
  if (!raw) return null;
  const found = raw
    .split(";")
    .map((s) => s.trim())
    .find((s) => s.startsWith(name + "="));
  return found ? found.slice(name.length + 1) : null;
}
function hashPassword(p) {
  return crypto
    .createHash("sha256")
    .update(String(p || ""))
    .digest("hex");
}
function getSession(req) {
  const raw = getCookie(req, AUTH_COOKIE);
  if (!raw) return null;
  try {
    return JSON.parse(
      Buffer.from(decodeURIComponent(raw), "base64").toString("utf8"),
    );
  } catch {
    return null;
  }
}
function setAuthCookie(res, session) {
  const value = Buffer.from(JSON.stringify(session), "utf8").toString("base64");
  res.cookie(AUTH_COOKIE, value, {
    maxAge: 30 * 24 * 60 * 60 * 1000,
    httpOnly: true,
    sameSite: "lax",
  });
}
function clearAuthCookies(res) {
  res.clearCookie("admin_auth");
  res.clearCookie(AUTH_COOKIE);
}
function requireAuth(req, res, next) {
  const s = getSession(req);
  if (s?.role) {
    req.auth = s;
    return next();
  }
  res.redirect("/login");
}
function requireAdmin(req, res, next) {
  const s = req.auth || getSession(req);
  if (s?.role === "admin") {
    req.auth = s;
    return next();
  }
  // Если это API-запрос — 403 JSON, иначе — редирект на /login
  const isApi =
    req.xhr ||
    req.path.startsWith("/api/") ||
    (req.headers["accept"] || "").includes("application/json") ||
    (req.headers["content-type"] || "").includes("json");
  if (isApi) return res.status(403).json({ error: "Admin only" });
  // Куратор → на свою страницу, неавторизован → логин
  if (s?.role === "curator") return res.redirect("/summary.html");
  res.redirect("/login");
}
function canAccessGroup(req, gid) {
  const s = req.auth || getSession(req);
  if (!s?.role) return false;
  if (s.role === "admin") return true;
  const ids = Array.isArray(s.groupIds)
    ? s.groupIds
    : s.groupId
      ? [s.groupId]
      : [];
  return ids.map(String).includes(String(gid || ""));
}
function getAccessibleGroupIds(req) {
  const s = req.auth || getSession(req);
  if (!s?.role) return [];
  if (s.role === "admin") return null;
  return Array.isArray(s.groupIds)
    ? s.groupIds.map(String)
    : s.groupId
      ? [String(s.groupId)]
      : [];
}
function requireGroupAccess(param = "id") {
  return (req, res, next) => {
    if (canAccessGroup(req, req.params[param])) return next();
    res.status(403).json({ error: "No access to this group" });
  };
}

app.get("/login", (req, res) => {
  const s = getSession(req);
  if (s?.role === "admin") return res.redirect("/");
  if (s?.role === "curator") return res.redirect("/summary.html");
  const err = req.query.error
    ? '<div style="color:#dc2626;font-size:13px;margin-bottom:15px;text-align:center;background:#fee2e2;padding:10px;border-radius:8px;">Неверный логин или пароль!</div>'
    : "";
  res.send(
    `<!DOCTYPE html><html lang="ru"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Вход | ПЛ №3</title><style>body{background:#f1f5f9;font-family:"Inter","Segoe UI",sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}.c{background:white;padding:40px 30px;border-radius:20px;box-shadow:0 10px 25px rgba(0,0,0,.05);width:100%;max-width:380px;border:1px solid #e2e8f0}h2{margin:0;font-size:24px;color:#0f172a;font-weight:800;text-align:center}p{margin:5px 0 30px;color:#64748b;font-size:14px;text-align:center}.fg{margin-bottom:20px}label{display:block;margin-bottom:8px;font-size:13px;font-weight:600;color:#475569}input{width:100%;padding:12px 15px;border:2px solid #e2e8f0;border-radius:10px;font-size:15px;outline:none;box-sizing:border-box}input:focus{border-color:#4f46e5}button{width:100%;padding:14px;background:#4f46e5;color:white;border:none;border-radius:10px;font-size:16px;font-weight:bold;cursor:pointer}button:hover{background:#4338ca}</style></head><body><div class="c"><h2>ПРОФЛИЦЕЙ №3</h2><p>Система управления</p>${err}<form action="/login" method="POST"><div class="fg"><label>Логин</label><input type="text" name="username" required autocomplete="off"></div><div class="fg"><label>Пароль</label><input type="password" name="password" required></div><button type="submit">Войти</button></form></div></body></html>`,
  );
});

app.post("/login", (req, res) => {
  const { username, password } = req.body;
  if (username === ADMIN_LOGIN && password === ADMIN_PASS) {
    setAuthCookie(res, { role: "admin", username: ADMIN_LOGIN });
    return res.redirect("/");
  }
  Curator.findOne({ username: String(username || "").trim() })
    .populate("groups")
    .then((curator) => {
      if (
        !curator ||
        curator.passwordHash !== hashPassword(password) ||
        !curator.groups?.length
      )
        return res.redirect("/login?error=1");
      const groups = curator.groups.map((g) => ({
        id: String(g._id),
        name: g.name,
      }));
      setAuthCookie(res, {
        role: "curator",
        username: curator.username,
        curatorId: String(curator._id),
        groupIds: groups.map((g) => g.id),
        groupId: groups[0]?.id || null,
        groupNames: groups.map((g) => g.name),
        groupName: groups[0]?.name || "",
      });
      res.redirect("/summary.html");
    })
    .catch(() => res.redirect("/login?error=1"));
});

app.get("/logout", (req, res) => {
  clearAuthCookies(res);
  res.redirect("/login");
});

const SELF_URL =
  process.env.RENDER_EXTERNAL_URL || "https://pl3service.onrender.com";
setInterval(
  () => {
    fetch(`${SELF_URL}/api/health`).catch(() => {});
  },
  10 * 60 * 1000,
);
app.get("/api/health", (req, res) => res.json({ ok: true }));

const MONGO_URI =
  process.env.MONGO_URI ||
  "mongodb+srv://ilimkaldybaev5_db_user:liceyStudents@riestr.uki8ep8.mongodb.net/PL3_Database?retryWrites=true&w=majority&appName=riestr";
mongoose
  .connect(MONGO_URI)
  .then(async () => {
    console.log("MongoDB connected");
    // Удалить устаревший уникальный индекс group_1 из коллекции curators
    try {
      await mongoose.connection.collection("curators").dropIndex("group_1");
      console.log("✅ Dropped old index: curators.group_1");
    } catch (e) {
      // Индекс уже удалён или не существует — ок
      if (e.code !== 27) console.log("index drop info:", e.message);
    }
    // Также сбросить любые другие проблемные индексы если есть
    try {
      await mongoose.connection.collection("curators").dropIndex("groups_1");
      console.log("✅ Dropped old index: curators.groups_1");
    } catch (e) {
      /* не существует — ок */
    }
  })
  .catch((e) => console.log("Mongo error:", e.message));

const AGENT_API_KEY = process.env.AGENT_API_KEY || "pl3_secret_key";

const PROFESSIONS = [
  { id: 1, ru: "Токарь", kg: "Токарь", duration: "2 года" },
  {
    id: 2,
    ru: "Электрогазосварщик",
    kg: "Электргазоширетүүчү",
    duration: "2 года",
  },
  {
    id: 3,
    ru: "Электромонтер по ремонту и обслуживанию электрооборудования",
    kg: "Электр жабдууларын оңдоо жана тейлөө боюнча электромонтер",
    duration: "2 года",
  },
  {
    id: 4,
    ru: "Мастер по ремонту и обслуживанию бытовой техники",
    kg: "Турмуш-тиричилик техникасын оңдоо жана тейлөө боюнча мастер",
    duration: "2 года",
  },
  {
    id: 5,
    ru: "Электрик по ремонту автомобильного электрооборудования",
    kg: "Автоунаанын электр жабдууларын оңдоо боюнча электрик",
    duration: "2 года",
  },
  {
    id: 6,
    ru: "Разработчик Web и мультимедийных приложений",
    kg: "Web жана мультимедиалык тиркемелерди иштеп чыгуучу",
    duration: "2 года",
  },
  {
    id: 7,
    ru: "Оператор цифровой печати",
    kg: "Санариптик басма оператору",
    duration: "2 года",
  },
  { id: 8, ru: "Повар", kg: "Ашпозчу", duration: "2 года" },
  { id: 9, ru: "Переплетчик", kg: "Түптөөчү", duration: "10 месяцев" },
  {
    id: 10,
    ru: "Электромонтер (10 м.)",
    kg: "Электромонтер",
    duration: "10 месяцев",
  },
  {
    id: 11,
    ru: "Автослесарь-Автоэлектрик",
    kg: "Автослесарь-Автоэлектрик",
    duration: "10 месяцев",
  },
  { id: 12, ru: "Программист", kg: "Программист", duration: "10 месяцев" },
  {
    id: 13,
    ru: "Оператор печатного оборудования",
    kg: "Басма жабдууларын оператору",
    duration: "2 года",
  },
];

const SUMMARY_BASE_SUBJECTS = [
  { key: "kg_lang", label: "Кыргыз тили", section: "theory" },
  { key: "kg_lit", label: "Кыргыз адабияты", section: "theory" },
  { key: "ru_lang", label: "Русский язык", section: "theory" },
  { key: "ru_lit", label: "Русская литература", section: "theory" },
  { key: "en_lang", label: "Английский язык", section: "theory" },
  { key: "history", label: "История", section: "theory" },
  { key: "society", label: "Человек и общество", section: "theory" },
  { key: "algebra", label: "Алгебра", section: "theory" },
  { key: "geometry", label: "Геометрия", section: "theory" },
  { key: "biology", label: "Биология", section: "theory" },
  { key: "geography", label: "География", section: "theory" },
  { key: "physics", label: "Физика", section: "theory" },
  { key: "astronomy", label: "Астрономия", section: "theory" },
  { key: "chemistry", label: "Химия", section: "theory" },
  { key: "pe", label: "Физкультура", section: "theory" },
];
const SUMMARY_SPECIAL_SUBJECTS = {
  6: [
    { key: "dpm", label: "ДПМ", section: "practice" },
    { key: "obip", label: "ОБиП", section: "practice" },
    { key: "law", label: "Правоведение", section: "practice" },
    { key: "oop", label: "ООП", section: "practice" },
    {
      key: "programming",
      label: "Основы языков программирования",
      section: "practice",
    },
  ],
  default: [
    { key: "special_1", label: "Спецтехнология", section: "practice" },
    {
      key: "special_2",
      label: "Производственное обучение",
      section: "practice",
    },
  ],
};

const Group = mongoose.model(
  "Group",
  new mongoose.Schema({
    name: { type: String, unique: true },
    profId: Number,
    profRu: String,
    profKg: String,
    duration: String,
    shiftStart: { type: String, default: "09:30" },
    summarySubjects: [
      {
        key: String,
        label: String,
        section: {
          type: String,
          enum: ["theory", "practice"],
          default: "theory",
        },
      },
    ],
  }),
);
const Student = mongoose.model(
  "Student",
  new mongoose.Schema({
    fio: String,
    birthDate: String,
    inn: String,
    lyceumId: String,
    group: { type: mongoose.Schema.Types.ObjectId, ref: "Group" },
    attendance: [{ date: String, present: Boolean }],
    skudDaily: [
      new mongoose.Schema(
        {
          date: String,
          firstIn: String,
          lastIn: String,
          firstOut: String,
          lastOut: String,
          inCount: { type: Number, default: 0 },
          outCount: { type: Number, default: 0 },
          present: { type: Boolean, default: false },
        },
        { _id: false },
      ),
    ],
  }),
);
const TgSub = mongoose.model(
  "TgSub",
  new mongoose.Schema(
    {
      chatId: { type: String, unique: true },
      name: String,
      role: {
        type: String,
        enum: ["admin", "curator", "parent"],
        default: "parent",
      },
      linkedInns: [{ type: String }], // только parent
      curatorId: { type: String, default: null }, // только curator
      notifyLate: { type: Boolean, default: true }, // curator: получать ли опоздания
    },
    { timestamps: true },
  ),
);
const Curator = mongoose.model(
  "Curator",
  new mongoose.Schema(
    {
      username: { type: String, unique: true, sparse: true },
      passwordHash: String,
      fullName: String,
      groups: [{ type: mongoose.Schema.Types.ObjectId, ref: "Group" }],
    },
    { timestamps: true, autoIndex: false },
  ),
);
const SummaryRecord = mongoose.model(
  "SummaryRecord",
  new mongoose.Schema(
    {
      group: {
        type: mongoose.Schema.Types.ObjectId,
        ref: "Group",
        index: true,
      },
      student: {
        type: mongoose.Schema.Types.ObjectId,
        ref: "Student",
        index: true,
      },
      month: { type: String, index: true },
      grades: { type: Map, of: String, default: {} },
    },
    { timestamps: true },
  ),
);

// Уведомить ТОЛЬКО кураторов конкретной группы (у которых notifyLate=true)
// ── NOTIFICATION STUBS ──
// Объявляем как let — будут заменены реальными функциями из бота после init
let notifyParents = async () => {};
let notifyAdmins = async () => {};
let notifyCuratorsOfGroup = async () => {};
let notifyAll = async () => {};

async function notifyTg(message) {
  try {
    await notifyAdmins(message);
  } catch (e) {
    console.error(e);
  }
}

function pad2(n) {
  return String(n).padStart(2, "0");
}
function isoToDateKey(iso) {
  const d = new Date(iso);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
function normalizeINN(s) {
  const d = String(s || "").replace(/\D/g, "");
  return d.length >= 12 ? d : null;
}
function normalizeFIO(s) {
  return String(s || "")
    .replace(/[.,]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}
function extractINN(s) {
  const m = String(s || "").match(/\b(\d{12,16})\b/);
  return m ? m[1] : null;
}
function extractFIO(s) {
  const inn = extractINN(s);
  let fio = String(s || "");
  if (inn) fio = fio.replace(inn, "");
  return fio.replace(/,/g, " ").replace(/\s+/g, " ").trim();
}
function getBishkekNow() {
  const n = new Date();
  return new Date(n.getTime() + n.getTimezoneOffset() * 60000 + 6 * 3600000);
}
function getTodayKey() {
  const d = getBishkekNow();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
function getCurrentMonth() {
  const d = getBishkekNow();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
}
function sanitizeSubjectKey(v, fb = "subject") {
  const k = String(v || "")
    .toLowerCase()
    .replace(/[^a-z0-9а-яё]+/gi, "_")
    .replace(/^_+|_+$/g, "");
  return k || fb;
}
function inferProfId(g) {
  if (g?.profId) return Number(g.profId);
  return PROFESSIONS.find((p) => p.ru === g?.profRu)?.id || null;
}
function buildDefaultSummarySubjects(g) {
  const pid = inferProfId(g),
    special = SUMMARY_SPECIAL_SUBJECTS[pid] || SUMMARY_SPECIAL_SUBJECTS.default;
  return [...SUMMARY_BASE_SUBJECTS, ...special].map((s, i) => ({
    key: sanitizeSubjectKey(s.key || s.label, `subject_${i + 1}`),
    label: s.label,
    section: s.section === "practice" ? "practice" : "theory",
  }));
}
async function ensureGroupSummarySubjects(g) {
  if (!g) return [];
  if (Array.isArray(g.summarySubjects) && g.summarySubjects.length)
    return g.summarySubjects.map((s, i) => ({
      key: sanitizeSubjectKey(s.key || s.label, `subject_${i + 1}`),
      label: s.label,
      section: s.section === "practice" ? "practice" : "theory",
    }));
  const ss = buildDefaultSummarySubjects(g);
  g.summarySubjects = ss;
  if (!g.profId) g.profId = inferProfId(g);
  await g.save();
  return ss;
}
const MRU = [
  "Январь",
  "Февраль",
  "Март",
  "Апрель",
  "Май",
  "Июнь",
  "Июль",
  "Август",
  "Сентябрь",
  "Октябрь",
  "Ноябрь",
  "Декабрь",
];
function formatMonthLabel(m) {
  const [y, mn] = String(m || "")
    .split("-")
    .map(Number);
  if (!y || !mn || mn < 1 || mn > 12) return String(m || "");
  return `${MRU[mn - 1]} ${y}`;
}

async function getSummaryDocumentData(groupId, month) {
  const group = await Group.findById(groupId);
  if (!group) throw new Error("Группа не найдена");
  const subjects = await ensureGroupSummarySubjects(group);
  const students = await Student.find({ group: groupId })
    .sort({ fio: 1 })
    .select("fio inn lyceumId");
  const records = await SummaryRecord.find({ group: groupId, month }).lean();
  const curator = await Curator.findOne({ groups: groupId })
    .sort({ createdAt: -1 })
    .select("fullName");
  const recMap = new Map(
    records.map((r) => {
      let g = {};
      if (r.grades instanceof Map) {
        r.grades.forEach((v, k) => {
          g[k] = v;
        });
      } else if (r.grades && typeof r.grades === "object") {
        const raw = r.grades;
        if (typeof raw.get === "function") {
          raw.forEach((v, k) => {
            g[k] = v;
          });
        } else {
          Object.entries(raw).forEach(([k, v]) => {
            if (!k.startsWith("$") && !k.startsWith("_")) g[k] = v;
          });
        }
      }
      return [String(r.student), g];
    }),
  );
  const theory = subjects.filter((s) => s.section !== "practice"),
    practice = subjects.filter((s) => s.section === "practice");
  const [y, mn] = String(month || "")
    .split("-")
    .map(Number);
  const studyYear = y && mn ? `${y}-${y + 1}` : String(month || "");
  const rows = students.map((s, i) => ({
    num: i + 1,
    fio: s.fio,
    inn: s.inn || "—",
    lyceumId: s.lyceumId || "—",
    grades: recMap.get(String(s._id)) || {},
  }));
  return {
    group,
    curatorFullName: curator?.fullName || "",
    month,
    monthLabel: formatMonthLabel(month),
    studyYear,
    subjects,
    theorySubjects: theory,
    practiceSubjects: practice,
    rows,
  };
}

async function buildSummaryDocumentHtml(groupId, month) {
  const data = await getSummaryDocumentData(groupId, month);
  const esc = (v) =>
    String(v || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  const all = [...data.theorySubjects, ...data.practiceSubjects];
  const tc = Math.max(data.theorySubjects.length, 1),
    pc = Math.max(data.practiceSubjects.length, 0);
  const mu = String(data.monthLabel || "").toUpperCase();
  const rowsHtml = data.rows
    .map(
      (row) =>
        `<tr><td class="num">${row.num}.</td><td class="fio">${esc(row.fio)}</td>${all
          .map((s) => {
            const g =
              row.grades instanceof Map
                ? row.grades.get(s.key) || ""
                : row.grades?.[s.key] || "";
            return `<td class="grade">${esc(g)}</td>`;
          })
          .join("")}</tr>`,
    )
    .join("");
  return `<!DOCTYPE html><html lang="ru"><head><meta charset="UTF-8"><title>Сводная ведомость</title><style>@page{size:A4 landscape;margin:10mm}*{box-sizing:border-box}body{margin:0;color:#000;font-family:"Times New Roman",serif;font-size:12pt;background:#fff}.sheet{width:100%;margin:0 auto}.title{margin:0 0 10px;text-align:center;font-size:21pt;font-weight:700;text-transform:uppercase}.subtitle,.speciality,.period{margin:0;text-align:center;font-size:13pt;font-weight:700;line-height:1.25}.speciality{font-style:italic}table{width:100%;border-collapse:collapse;table-layout:fixed;margin-top:14px}th,td{border:1px solid #000;text-align:center;vertical-align:middle;padding:4px 3px;font-size:10.5pt;line-height:1.15}th{font-weight:700}.num{width:4%;white-space:nowrap}.fio{width:22%;text-align:left;font-weight:700}.subject-group{font-size:11pt;font-weight:700}.subject{font-size:9pt;font-weight:700;word-break:break-word}.grade{font-size:11pt}.signatures{display:flex;justify-content:space-between;gap:24px;margin-top:18px;font-size:12pt}.signatures div{width:48%}@media print{body{background:#fff}}</style></head><body><div class="sheet"><p class="title">Сводная ведомость</p><p class="subtitle">успеваемости учащихся группы «${esc(data.group.name)}»</p><p class="speciality">Специальность: «${esc(data.group.profRu || "Не указана")}»</p><p class="period">ПЛ № 3 за ${esc(mu)} ${esc(data.studyYear)} гг.</p><table><thead><tr><th rowspan="2" class="num">№<br>п/п</th><th rowspan="2" class="fio">Фамилия, имя и отчество</th><th colspan="${tc}" class="subject-group">Предметы теоретического обучения</th>${pc ? `<th colspan="${pc}" class="subject-group">Производст. обучения</th>` : ""}</tr><tr>${data.theorySubjects.map((s) => `<th class="subject">${esc(s.label)}</th>`).join("") || '<th class="subject">—</th>'}${data.practiceSubjects.map((s) => `<th class="subject">${esc(s.label)}</th>`).join("")}</tr></thead><tbody>${rowsHtml}</tbody></table><div class="signatures"><div>Зам. директора по УПМР ___________________</div><div>Куратор группы ${esc(data.curatorFullName || "___________________")}</div></div></div></body></html>`;
}

const SKUD_EVENTS = [];

/* ── FILENAME HELPERS ── */
function safeFilename(name) {
  return String(name || "file").replace(/[^a-zA-Z0-9._-]/g, "_");
}
function contentDisposition(filename) {
  const ascii = String(filename).replace(/[^ -~]/g, "_");
  const encoded = encodeURIComponent(filename);
  return `attachment; filename="${ascii}"; filename*=UTF-8''${encoded}`;
}

function pushSkudEvent(ev) {
  SKUD_EVENTS.unshift(ev);
  if (SKUD_EVENTS.length > 500) SKUD_EVENTS.pop();
}

/* === ПУБЛИЧНЫЕ МАРШРУТЫ === */
app.get("/journal.html", (req, res) =>
  res.sendFile(path.join(__dirname, "public", "journal.html")),
);

app.get("/api/admin/groups-list", async (req, res) => {
  try {
    const ids = getAccessibleGroupIds(req);
    const filter = ids === null ? {} : ids.length ? { _id: { $in: ids } } : {};
    res.json(await Group.find(filter).sort({ name: 1 }));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/admin/global-stats", async (req, res) => {
  try {
    const tk = getTodayKey();
    const c = await Student.countDocuments({
      $or: [
        { skudDaily: { $elemMatch: { date: tk, present: true } } },
        { attendance: { $elemMatch: { date: tk, present: true } } },
      ],
    });
    res.json({ presentToday: c });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/admin/attendance-matrix/:gid/:month", async (req, res) => {
  try {
    const session = getSession(req);
    if (session?.role && !canAccessGroup(req, req.params.gid))
      return res.status(403).json({ error: "Нет доступа к этой группе" });
    const { gid, month } = req.params,
      [year, mon] = month.split("-").map(Number),
      dim = new Date(year, mon, 0).getDate();
    const group = await Group.findById(gid),
      shiftTime = group?.shiftStart || "09:30";
    const [sH, sM] = shiftTime.split(":").map(Number),
      stm = sH * 60 + sM;
    const students = await Student.find({ group: gid }).sort({ fio: 1 });
    const bd = getBishkekNow(),
      tdk = `${bd.getFullYear()}-${pad2(bd.getMonth() + 1)}-${pad2(bd.getDate())}`,
      tiw = bd.getDay() === 0 || bd.getDay() === 6;
    let ptc = 0,
      tml = 0;
    const matrix = students.map((s) => {
      const days = [];
      let ipt = false,
        sl = 0,
        sp = 0,
        sa = 0;
      for (let d = 1; d <= dim; d++) {
        const dk = `${year}-${pad2(mon)}-${pad2(d)}`,
          do2 = new Date(Date.UTC(year, mon - 1, d)),
          iw = do2.getUTCDay() === 0 || do2.getUTCDay() === 6;
        const skud = s.skudDaily?.find((r) => r.date === dk),
          att = s.attendance?.find((a) => a.date === dk);
        let present = false,
          late = false,
          ts = "";
        if (skud) {
          present = skud.present;
          if (skud.firstIn) {
            const dt = new Date(skud.firstIn);
            if (!isNaN(dt.getTime())) {
              const tm = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360,
                lh = Math.floor(tm / 60) % 24,
                lm = tm % 60;
              ts = `${pad2(lh)}:${pad2(lm)}`;
              late = lh * 60 + lm > stm;
            }
          }
        } else if (att) {
          present = att.present;
        }
        if (present) sp++;
        else if (!iw && dk <= tdk) sa++;
        if (late) {
          sl++;
          tml++;
        }
        if (dk === tdk && present) ipt = true;
        days.push({ present, late, time: ts, isWeekend: iw });
      }
      if (ipt) ptc++;
      return {
        fio: s.fio,
        days,
        stats: { presences: sp, lates: sl, absences: sa },
      };
    });
    res.json({
      daysInMonth: dim,
      matrix,
      shiftStart: shiftTime,
      stats: {
        totalStudents: students.length,
        presentToday: ptc,
        absentToday: tiw ? 0 : students.length - ptc,
        totalLates: tml,
      },
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post("/api/agent/sync", async (req, res) => {
  const { apiKey, data } = req.body;
  if (apiKey !== AGENT_API_KEY) return res.status(403).send("Forbidden");
  if (!Array.isArray(data))
    return res.status(400).json({ error: "array expected" });
  let ok = 0,
    fail = 0;
  for (const entry of data) {
    try {
      const employeeNo = entry.id?.toString().trim(),
        deviceIp = entry.deviceIp,
        direction = entry.direction,
        rawHikTime = entry.time || "",
        eventIso = rawHikTime;
      let dateKey;
      const dt = new Date(rawHikTime);
      if (!isNaN(dt.getTime())) {
        dt.setUTCHours(dt.getUTCHours() + 6);
        dateKey = `${dt.getUTCFullYear()}-${pad2(dt.getUTCMonth() + 1)}-${pad2(dt.getUTCDate())}`;
      } else {
        dateKey =
          rawHikTime.length >= 10
            ? rawHikTime.substring(0, 10)
            : isoToDateKey(new Date().toISOString());
      }
      const rawName = String(entry.rawName || "").trim();
      if (!rawName) {
        pushSkudEvent({
          ok: false,
          reason: "person_not_found",
          employeeNo,
          deviceIp,
          at: eventIso,
        });
        fail++;
        continue;
      }
      const inn = normalizeINN(extractINN(rawName)),
        fioNorm = normalizeFIO(extractFIO(rawName));
      let student = null;
      if (inn) student = await Student.findOne({ inn });
      if (!student && fioNorm) {
        const fp = fioNorm.split(" ")[0],
          cands = await Student.find({ fio: new RegExp("^" + fp, "i") }).limit(
            50,
          );
        student = cands.find((c) => normalizeFIO(c.fio) === fioNorm) || null;
      }
      if (!student) {
        pushSkudEvent({
          ok: false,
          reason: "student_not_found",
          rawName,
          innFromHik: inn,
          employeeNo,
          deviceIp,
          at: eventIso,
        });
        fail++;
        continue;
      }
      const exist = student.attendance.find((a) => a.date === dateKey);
      if (exist) exist.present = true;
      else student.attendance.push({ date: dateKey, present: true });
      let rec = student.skudDaily.find((r) => r.date === dateKey);
      if (!rec) {
        student.skudDaily.push({
          date: dateKey,
          inCount: 0,
          outCount: 0,
          present: false,
        });
        rec = student.skudDaily[student.skudDaily.length - 1];
      }
      let isFirst = false,
        timeStr = "—",
        lh = 0,
        lm = 0;
      if (direction === "in") {
        isFirst = rec.inCount === 0;
        rec.inCount++;
        rec.present = true;
        rec.firstIn = rec.firstIn || eventIso;
        rec.lastIn = eventIso;
        const dtE = new Date(eventIso);
        if (!isNaN(dtE.getTime())) {
          const tm = dtE.getUTCHours() * 60 + dtE.getUTCMinutes() + 360;
          lh = Math.floor(tm / 60) % 24;
          lm = tm % 60;
          timeStr = `${pad2(lh)}:${pad2(lm)}`;
        }
        if (isFirst) {
          await student.populate("group");
          const st = student.group?.shiftStart || "09:30",
            [shH, shM] = st.split(":").map(Number);
          if (lh * 60 + lm > shH * 60 + shM) {
            // Сохраняем опоздание — куратор видит через кнопку "Опоздавшие сегодня" в боте
            pushSkudEvent({
              ok: true,
              type: "late",
              studentFio: student.fio,
              studentInn: student.inn || "",
              groupId: String(student.group?._id || ""),
              groupName: student.group?.name || "—",
              time: timeStr,
              at: eventIso,
            });
          }
          if (student.inn) {
            notifyParents(
              student.inn,
              `✅ *Вход в лицей*\n🧑‍🎓 ${student.fio}\n🕒 *${timeStr}*`,
            ).catch(() => null);
          }
        }
      } else {
        rec.outCount++;
        rec.firstOut = rec.firstOut || eventIso;
        rec.lastOut = eventIso;
        const dtE = new Date(eventIso);
        if (!isNaN(dtE.getTime()) && student.inn) {
          const tm = dtE.getUTCHours() * 60 + dtE.getUTCMinutes() + 360,
            tStr = `${pad2(Math.floor(tm / 60) % 24)}:${pad2(tm % 60)}`;
          if (student.inn) {
            notifyParents(
              student.inn,
              `🏃‍♂️ *Выход из лицея*\n🧑‍🎓 ${student.fio}\n🕒 *${tStr}*`,
            ).catch(() => null);
          }
        }
      }
      student.markModified("attendance");
      student.markModified("skudDaily");
      await student.save();
      pushSkudEvent({
        ok: true,
        rawName,
        studentFio: student.fio,
        studentInn: student.inn,
        direction,
        deviceIp,
        at: eventIso,
      });
      ok++;
    } catch (e) {
      fail++;
      pushSkudEvent({ ok: false, reason: e.message });
    }
  }
  res.json({ success: true, okCount: ok, failCount: fail });
});

/* === ЗАЩИТА СТРАНИЦ === */
app.get(["/", "/index.html"], requireAuth, requireAdmin, (req, res) =>
  res.sendFile(path.join(__dirname, "public", "index.html")),
);
app.get("/skud-events.html", requireAuth, requireAdmin, (req, res) =>
  res.sendFile(path.join(__dirname, "public", "skud-events.html")),
);
app.get("/admin/events", requireAuth, requireAdmin, (req, res) =>
  res.sendFile(path.join(__dirname, "public", "skud-events.html")),
);
app.get("/summary.html", requireAuth, (req, res) =>
  res.sendFile(path.join(__dirname, "public", "summary.html")),
);
app.use("/api/admin", requireAuth);
app.use(express.static("public"));

app.get("/api/admin/me", async (req, res) => {
  try {
    const s = req.auth || getSession(req);
    if (!s?.role) return res.status(401).json({ error: "Не авторизован" });
    let groups = [];
    if (s.role === "admin")
      groups = await Group.find()
        .sort({ name: 1 })
        .select("name profRu profId duration shiftStart");
    else if (s.groupIds?.length)
      groups = await Group.find({ _id: { $in: s.groupIds } })
        .sort({ name: 1 })
        .select("name profRu profId duration shiftStart");
    res.json({
      role: s.role,
      username: s.username,
      groupId: s.groupId || null,
      groupName: s.groupName || null,
      groupIds: s.groupIds || (s.groupId ? [s.groupId] : []),
      groupNames: s.groupNames || (s.groupName ? [s.groupName] : []),
      groups,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/admin/skud-events", requireAdmin, (req, res) =>
  res.json({ ok: true, items: SKUD_EVENTS }),
);

app.delete("/api/admin/group/:id", requireAdmin, async (req, res) => {
  try {
    await Curator.updateMany({}, { $pull: { groups: req.params.id } });
    await SummaryRecord.deleteMany({ group: req.params.id });
    await Student.deleteMany({ group: req.params.id });
    await Group.findByIdAndDelete(req.params.id);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.delete("/api/admin/group/:id/clear", requireAdmin, async (req, res) => {
  try {
    const ss = await Student.find({ group: req.params.id }).select("_id");
    await SummaryRecord.deleteMany({ student: { $in: ss.map((s) => s._id) } });
    await Student.deleteMany({ group: req.params.id });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.put("/api/admin/group/:id/settings", requireAdmin, async (req, res) => {
  try {
    const { profId, shiftStart } = req.body;
    let ud = { shiftStart: shiftStart || "09:30" };
    if (profId) {
      const prof = PROFESSIONS.find((p) => p.id == profId);
      if (prof) {
        ud.profId = Number(profId);
        ud.profRu = prof.ru;
        ud.profKg = prof.kg;
        ud.duration = prof.duration;
        ud.summarySubjects = buildDefaultSummarySubjects({
          profId: Number(profId),
          profRu: prof.ru,
        });
      }
    }
    await Group.findByIdAndUpdate(req.params.id, ud);
    res.json({ ok: true, updateData: ud });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.get(
  "/api/admin/group/:id/summary-config",
  requireGroupAccess("id"),
  async (req, res) => {
    try {
      const g = await Group.findById(req.params.id);
      if (!g) return res.status(404).json({ error: "Группа не найдена" });
      const subjects = await ensureGroupSummarySubjects(g);
      res.json({
        group: {
          _id: g._id,
          name: g.name,
          profRu: g.profRu || "",
          profId: g.profId || inferProfId(g),
        },
        subjects,
        theorySubjects: subjects.filter((s) => s.section !== "practice"),
        practiceSubjects: subjects.filter((s) => s.section === "practice"),
      });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);
app.put(
  "/api/admin/group/:id/summary-subjects",
  requireAdmin,
  async (req, res) => {
    try {
      const subjects = Array.isArray(req.body.subjects)
        ? req.body.subjects
        : [];
      const norm = subjects
        .map((s, i) => ({
          key: sanitizeSubjectKey(s.key || s.label, `subject_${i + 1}`),
          label: String(s.label || "").trim(),
          section: s.section === "practice" ? "practice" : "theory",
        }))
        .filter((s) => s.label);
      if (!norm.length)
        return res.status(400).json({ error: "Список предметов пуст" });
      await Group.findByIdAndUpdate(req.params.id, { summarySubjects: norm });
      res.json({ ok: true, subjects: norm });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);

app.get("/api/admin/curators", requireAdmin, async (req, res) => {
  try {
    const curators = await Curator.find()
      .populate("groups", "name")
      .sort({ createdAt: -1 });
    res.json(
      curators.map((c) => ({
        _id: c._id,
        username: c.username,
        fullName: c.fullName || "",
        groupIds: c.groups?.map((g) => String(g._id)) || [],
        groupNames: c.groups?.map((g) => g.name) || [],
        createdAt: c.createdAt,
      })),
    );
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.post("/api/admin/curators", requireAdmin, async (req, res) => {
  try {
    const username = String(req.body.username || "").trim();
    const password = String(req.body.password || "").trim();
    const fullName = String(req.body.fullName || "").trim();
    const rawCurId = req.body.curatorId;
    const curatorId =
      rawCurId &&
      String(rawCurId) !== "undefined" &&
      String(rawCurId) !== "null"
        ? String(rawCurId).trim()
        : "";
    const groupIds = (
      Array.isArray(req.body.groupIds) ? req.body.groupIds : [req.body.groupId]
    )
      .map((id) => String(id || "").trim())
      .filter(Boolean);

    if (!username || !password || !groupIds.length)
      return res
        .status(400)
        .json({ error: "Логин, пароль и группа обязательны" });

    const groups = await Group.find({ _id: { $in: groupIds } }).select("name");
    if (!groups.length)
      return res.status(404).json({ error: "Группы не найдены" });

    const groupObjectIds = groups.map((g) => g._id);
    const passwordHash = hashPassword(password);

    let finalDoc;

    if (curatorId && mongoose.Types.ObjectId.isValid(curatorId)) {
      // ── Редактирование по ID ──
      finalDoc = await Curator.findByIdAndUpdate(
        curatorId,
        { $set: { username, passwordHash, fullName, groups: groupObjectIds } },
        { new: true, runValidators: false },
      );
      if (!finalDoc)
        return res.status(404).json({ error: "Куратор не найден" });
    } else {
      // ── Создание или обновление по username ──
      // Используем нативный драйвер чтобы избежать unique index errors от mongoose
      const col = Curator.collection;
      await col.updateOne(
        { username }, // фильтр
        {
          $set: {
            username,
            passwordHash,
            fullName,
            groups: groupObjectIds,
            updatedAt: new Date(),
          },
          $setOnInsert: { createdAt: new Date() },
        },
        { upsert: true }, // создать если нет, обновить если есть
      );
      finalDoc = await Curator.findOne({ username });
    }

    res.json({
      ok: true,
      curator: {
        _id: finalDoc._id,
        username: finalDoc.username,
        fullName: finalDoc.fullName || "",
        groupIds: groups.map((g) => String(g._id)),
        groupNames: groups.map((g) => g.name),
      },
    });
  } catch (e) {
    console.error("curator error:", e.code, e.message);
    res.status(500).json({ error: e.message });
  }
});

// Очистить дублирующихся кураторов (оставить последнего по дате)
app.delete("/api/admin/curators/cleanup", requireAdmin, async (req, res) => {
  try {
    const all = await Curator.find().sort({ createdAt: 1 });
    const seen = new Map();
    const toDelete = [];
    for (const c of all) {
      if (seen.has(c.username)) toDelete.push(c._id);
      else seen.set(c.username, c._id);
    }
    if (toDelete.length) await Curator.deleteMany({ _id: { $in: toDelete } });
    res.json({ ok: true, deleted: toDelete.length });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.delete("/api/admin/curators/:id", requireAdmin, async (req, res) => {
  try {
    await Curator.findByIdAndDelete(req.params.id);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get(
  "/api/admin/group-students/:id",
  requireGroupAccess("id"),
  async (req, res) => {
    try {
      res.json(await Student.find({ group: req.params.id }).sort({ fio: 1 }));
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);
app.get("/api/admin/search", requireAdmin, async (req, res) => {
  try {
    const q = (req.query.q || "").trim(),
      page = Math.max(1, parseInt(req.query.page) || 1),
      limit = Math.min(100, parseInt(req.query.limit) || 60),
      skip = (page - 1) * limit;
    const filter = q.length >= 2 ? { fio: new RegExp(q, "i") } : {};
    const [total, students] = await Promise.all([
      Student.countDocuments(filter),
      Student.find(filter)
        .populate("group")
        .sort({ fio: 1 })
        .skip(skip)
        .limit(limit),
    ]);
    res.json({ students, total, page, limit });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.post("/api/admin/student", requireAdmin, async (req, res) => {
  try {
    const { fio, birthDate, inn, lyceumId, groupId } = req.body;
    if (!fio || !fio.trim())
      return res.status(400).json({ error: "ФИО обязательно" });
    if (!groupId) return res.status(400).json({ error: "Группа обязательна" });
    const student = await Student.create({
      fio: fio.trim(),
      birthDate,
      inn,
      lyceumId,
      group: groupId,
    });
    res.json({ ok: true, student });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.put("/api/admin/student/:id", requireAdmin, async (req, res) => {
  try {
    const { fio, birthDate, inn, lyceumId } = req.body;
    await Student.findByIdAndUpdate(req.params.id, {
      fio,
      birthDate,
      inn,
      lyceumId,
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.delete("/api/admin/student/:id", requireAdmin, async (req, res) => {
  try {
    await SummaryRecord.deleteMany({ student: req.params.id });
    await Student.findByIdAndDelete(req.params.id);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get(
  "/api/admin/summary/:gid/:month",
  requireGroupAccess("gid"),
  async (req, res) => {
    try {
      const g = await Group.findById(req.params.gid);
      if (!g) return res.status(404).json({ error: "Группа не найдена" });
      const subjects = await ensureGroupSummarySubjects(g);
      const students = await Student.find({ group: req.params.gid })
        .sort({ fio: 1 })
        .select("fio birthDate inn lyceumId");
      const records = await SummaryRecord.find({
        group: req.params.gid,
        month: req.params.month,
      }).lean();
      const recMap = new Map(
        records.map((r) => {
          let g = {};
          if (r.grades instanceof Map) {
            r.grades.forEach((v, k) => {
              g[k] = v;
            });
          } else if (r.grades && typeof r.grades === "object") {
            const raw = r.grades;
            if (typeof raw.get === "function") {
              raw.forEach((v, k) => {
                g[k] = v;
              });
            } else {
              Object.entries(raw).forEach(([k, v]) => {
                if (!k.startsWith("$") && !k.startsWith("_")) g[k] = v;
              });
            }
          }
          return [String(r.student), g];
        }),
      );
      res.json({
        group: {
          _id: g._id,
          name: g.name,
          profRu: g.profRu || "",
          profId: g.profId || inferProfId(g),
        },
        month: req.params.month,
        subjects,
        rows: students.map((s) => ({
          studentId: s._id,
          fio: s.fio,
          birthDate: s.birthDate || "",
          inn: s.inn || "",
          lyceumId: s.lyceumId || "",
          grades: recMap.get(String(s._id)) || {},
        })),
      });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);
app.put(
  "/api/admin/summary/:gid/:month",
  requireGroupAccess("gid"),
  async (req, res) => {
    try {
      const g = await Group.findById(req.params.gid);
      if (!g) return res.status(404).json({ error: "Группа не найдена" });
      const rows = Array.isArray(req.body.rows) ? req.body.rows : [];
      const subjects = await ensureGroupSummarySubjects(g),
        ak = new Set(subjects.map((s) => s.key));
      const sids = rows.map((r) => String(r.studentId || "")).filter(Boolean);
      const ss = await Student.find({
        _id: { $in: sids },
        group: req.params.gid,
      }).select("_id");
      const as = new Set(ss.map((s) => String(s._id)));
      for (const row of rows) {
        if (!as.has(String(row.studentId || ""))) continue;
        const grades = {};
        Object.entries(row.grades || {}).forEach(([k, v]) => {
          if (!ak.has(k)) return;
          const n = String(v || "").trim();
          if (n) grades[k] = n;
        });
        await SummaryRecord.findOneAndUpdate(
          {
            group: req.params.gid,
            student: row.studentId,
            month: req.params.month,
          },
          {
            group: req.params.gid,
            student: row.studentId,
            month: req.params.month,
            grades,
          },
          { upsert: true, new: true, setDefaultsOnInsert: true },
        );
      }
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);
app.get(
  "/api/admin/summary-print/:gid/:month",
  requireGroupAccess("gid"),
  async (req, res) => {
    try {
      const html = await buildSummaryDocumentHtml(
        req.params.gid,
        req.params.month,
      );
      res.send(`${html}<script>window.onload=()=>window.print()</script>`);
    } catch (e) {
      res.status(500).send(e.message);
    }
  },
);

app.get(
  "/api/admin/export-journal/:gid/:month",
  requireGroupAccess("gid"),
  async (req, res) => {
    try {
      const { gid, month } = req.params,
        [year, mon] = month.split("-").map(Number),
        dim = new Date(year, mon, 0).getDate();
      const g = await Group.findById(gid);
      if (!g) return res.status(404).send("Группа не найдена");
      const [sH, sM] = (g.shiftStart || "09:30").split(":").map(Number),
        stm = sH * 60 + sM;
      const students = await Student.find({ group: gid }).sort({ fio: 1 });
      const bd = getBishkekNow(),
        tdk = `${bd.getFullYear()}-${pad2(bd.getMonth() + 1)}-${pad2(bd.getDate())}`;
      const excelData = [],
        hdr = ["№", "ФИО Студента"];
      for (let d = 1; d <= dim; d++) hdr.push(String(d));
      hdr.push("Присутствий", "Отсутствий", "Опозданий");
      excelData.push(hdr);
      students.forEach((s, i) => {
        const row = [i + 1, s.fio];
        let sp = 0,
          sa = 0,
          sl = 0;
        for (let d = 1; d <= dim; d++) {
          const dk = `${year}-${pad2(mon)}-${pad2(d)}`,
            do2 = new Date(Date.UTC(year, mon - 1, d)),
            iw = do2.getUTCDay() === 0 || do2.getUTCDay() === 6;
          const skud = s.skudDaily?.find((r) => r.date === dk),
            att = s.attendance?.find((a) => a.date === dk);
          let present = false,
            late = false;
          if (skud) {
            present = skud.present;
            if (skud.firstIn) {
              const dt = new Date(skud.firstIn);
              if (!isNaN(dt.getTime())) {
                const tm = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
                if ((Math.floor(tm / 60) % 24) * 60 + (tm % 60) > stm)
                  late = true;
              }
            }
          } else if (att) {
            present = att.present;
          }
          if (present) {
            sp++;
            if (late) {
              sl++;
              row.push("О");
            } else row.push("П");
          } else if (!iw && dk <= tdk) {
            sa++;
            row.push("Н");
          } else row.push("-");
        }
        row.push(sp, sa, sl);
        excelData.push(row);
      });
      const wb = xlsx.utils.book_new(),
        ws = xlsx.utils.aoa_to_sheet(excelData);
      const cw = [{ wch: 4 }, { wch: 35 }];
      for (let d = 1; d <= dim; d++) cw.push({ wch: 4 });
      cw.push({ wch: 12 }, { wch: 12 }, { wch: 12 });
      ws["!cols"] = cw;
      xlsx.utils.book_append_sheet(wb, ws, "Журнал");
      const buf = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
      res.setHeader(
        "Content-Disposition",
        contentDisposition(`Journal_${g.name}_${month}.xlsx`),
      );
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      );
      res.send(buf);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);

app.get("/api/admin/professions", (req, res) => res.json(PROFESSIONS));

app.post(
  "/api/admin/import",
  requireAdmin,
  upload.single("file"),
  async (req, res) => {
    try {
      const { groupName, profId } = req.body,
        prof = PROFESSIONS.find((p) => p.id == profId);
      if (!prof) return res.status(400).json({ error: "Профессия не найдена" });
      let g = await Group.findOne({ name: groupName });
      if (!g)
        g = await Group.create({
          name: groupName,
          profId: Number(profId),
          profRu: prof.ru,
          profKg: prof.kg,
          duration: prof.duration,
          shiftStart: "09:30",
          summarySubjects: buildDefaultSummarySubjects({
            profId: Number(profId),
            profRu: prof.ru,
          }),
        });
      const wb = xlsx.readFile(req.file.path),
        ws = wb.Sheets[wb.SheetNames[0]],
        rows = xlsx.utils.sheet_to_json(ws, { header: 1 });
      let created = 0;
      for (const row of rows) {
        const fio = String(row[0] || "").trim(),
          birthDate = String(row[1] || "").trim(),
          inn = String(row[2] || "").trim(),
          lyceumId = String(row[3] || "").trim();
        if (!fio || fio.length < 3) continue;
        await Student.findOneAndUpdate(
          { inn: inn || null, fio },
          { fio, birthDate, inn, lyceumId, group: g._id },
          { upsert: true, new: true },
        );
        created++;
      }
      fs.unlinkSync(req.file.path);
      res.json({ ok: true, created });
    } catch (e) {
      if (req.file?.path && fs.existsSync(req.file.path))
        fs.unlinkSync(req.file.path);
      res.status(500).json({ error: e.message });
    }
  },
);

app.get("/api/admin/print/:id/:type", requireAdmin, async (req, res) => {
  try {
    const student = await Student.findById(req.params.id).populate("group");
    if (!student) return res.status(404).send("Студент не найден");
    const tplMap = {
      common: "template_common.docx",
      army: "template_army.docx",
      social: "template_social.docx",
    };
    const tplPath = path.join(
      __dirname,
      "templates",
      tplMap[req.params.type] || tplMap.common,
    );
    if (!fs.existsSync(tplPath))
      return res.json({ message: `Шаблон не найден.` });
    const content = fs.readFileSync(tplPath, "binary"),
      zip = new PizZip(content);
    const qrDataUrl = await QRCode.toDataURL(
      `ФИО: ${student.fio}\nИНН: ${student.inn || "—"}\nГруппа: ${student.group?.name || "—"}`,
    );
    const qrBuffer = Buffer.from(qrDataUrl.split(",")[1], "base64");
    const imageModule = new ImageModule({
      centered: false,
      getImage: (tv) => (tv === "qr" ? qrBuffer : fs.readFileSync(tv)),
      getSize: () => [80, 80],
    });
    const doc = new Docxtemplater(zip, {
      modules: [imageModule],
      paragraphLoop: true,
      linebreaks: true,
    });
    const nowD = new Date(),
      mRu = [
        "января",
        "февраля",
        "марта",
        "апреля",
        "мая",
        "июня",
        "июля",
        "августа",
        "сентября",
        "октября",
        "ноября",
        "декабря",
      ];
    doc.render({
      fio: student.fio,
      birth: student.birthDate || "—",
      inn: student.inn || "—",
      lycId: student.lyceumId || "—",
      group: student.group?.name || "—",
      profRu: student.group?.profRu || "—",
      profKg: student.group?.profKg || "—",
      duration: student.group?.duration || "—",
      prof_ru: student.group?.profRu || "—",
      prof_kg: student.group?.profKg || "—",
      prof: student.group?.profRu || "—",
      profession: student.group?.profRu || "—",
      date: `${nowD.getDate()} ${mRu[nowD.getMonth()]} ${nowD.getFullYear()}`,
      year: nowD.getFullYear(),
      qr: "qr",
    });
    const buf = doc.getZip().generate({ type: "nodebuffer" });
    res.setHeader(
      "Content-Disposition",
      contentDisposition(`spravka_${student.lyceumId || student._id}.docx`),
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get(
  "/api/admin/report-food/:gid/:month",
  requireGroupAccess("gid"),
  async (req, res) => {
    try {
      const { gid, month } = req.params,
        [year, mon] = month.split("-").map(Number),
        dim = new Date(year, mon, 0).getDate();
      const g = await Group.findById(gid),
        students = await Student.find({ group: gid }).sort({ fio: 1 });
      if (!g) return res.status(404).json({ error: "Группа не найдена" });
      if (!students.length)
        return res.status(404).json({ error: "Студентов нет" });
      const rows = students.map((s, i) => {
        let pd = 0,
          dc = [];
        for (let d = 1; d <= dim; d++) {
          const dk = `${year}-${pad2(mon)}-${pad2(d)}`,
            skud = s.skudDaily?.find((r) => r.date === dk),
            att = s.attendance?.find((a) => a.date === dk),
            present = skud ? skud.present : att ? att.present : false;
          if (present) pd++;
          dc.push(present ? "+" : " ");
        }
        return {
          num: i + 1,
          fio: s.fio,
          lyceumId: s.lyceumId || "—",
          days: dc,
          totalDays: pd,
          totalAmount: pd * 60,
        };
      });
      const tplPath = path.join(__dirname, "templates", "template_food.docx");
      if (!fs.existsSync(tplPath))
        return res.json({ message: "Шаблон template_food.docx не найден." });
      const MU = [
        "ЯНВАРЬ",
        "ФЕВРАЛЬ",
        "МАРТ",
        "АПРЕЛЬ",
        "МАЙ",
        "ИЮНЬ",
        "ИЮЛЬ",
        "АВГУСТ",
        "СЕНТЯБРЬ",
        "ОКТЯБРЬ",
        "НОЯБРЬ",
        "ДЕКАБРЬ",
      ];
      const gt = rows.reduce((s, r) => s + r.totalAmount, 0);
      const content = fs.readFileSync(tplPath, "binary"),
        zip = new PizZip(content),
        doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
      doc.render({
        month_name: MU[mon - 1],
        year: String(year),
        group: g.name,
        groupName: g.name,
        st: rows.map((r) => ({
          fio: r.fio,
          lyceum_id: r.lyceumId,
          days: String(r.totalDays),
          price: "60",
          sum: String(r.totalAmount),
          num: String(r.num),
          lyceumId: r.lyceumId,
          totalDays: String(r.totalDays),
          totalAmount: String(r.totalAmount),
        })),
        total: String(gt),
        grandTotal: String(gt),
        profRu: g.profRu || "—",
        profKg: g.profKg || "—",
        prof_ru: g.profRu || "—",
        prof_kg: g.profKg || "—",
        duration: g.duration || "—",
        month: `${pad2(mon)}.${year}`,
        daysInMonth: String(dim),
      });
      const buf = doc.getZip().generate({ type: "nodebuffer" });
      res.setHeader(
        "Content-Disposition",
        contentDisposition(`food_${g.name}_${month}.docx`),
      );
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      );
      res.send(buf);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);

/* === API: СГЕНЕРИРОВАТЬ ССЫЛКУ КУРАТОРА ДЛЯ TELEGRAM === */
app.get("/api/admin/curator-link/:id", requireAdmin, async (req, res) => {
  try {
    const cur = await Curator.findById(req.params.id).populate(
      "groups",
      "name",
    );
    if (!cur) return res.status(404).json({ error: "Куратор не найден" });
    const curId = String(cur._id);
    // HMAC подпись — защита от подбора
    const hmac = require("crypto")
      .createHmac("sha256", process.env.ADMIN_TG_SECRET || "pl3admin2026")
      .update(curId)
      .digest("hex")
      .slice(0, 8);
    const token = `${curId}_${hmac}`;
    const botUser = process.env.BOT_USERNAME || "pl3_school_bot";
    const link = `https://t.me/${botUser}?start=curator_${token}`;
    res.json({
      ok: true,
      link,
      curatorName: cur.fullName || cur.username,
      groups: cur.groups.map((g) => g.name),
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* === API: НАСТРОЙКИ УВЕДОМЛЕНИЙ КУРАТОРА В ТГ === */
app.put("/api/admin/curator-tg-notify/:id", requireAdmin, async (req, res) => {
  try {
    const curId = String(req.params.id);
    const { notifyLate } = req.body;
    await TgSub.updateMany(
      { role: "curator", curatorId: curId },
      { notifyLate: !!notifyLate },
    );
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* ============================================================
   ГЕНЕРАЦИЯ СВОДНОЙ ВЕДОМОСТИ — ДИНАМИЧЕСКИЕ КОЛОНКИ
============================================================ */
function buildSummaryDocx(data) {
  const theory = data.subjects.filter((s) => s.section !== "practice");
  const practice = data.subjects.filter((s) => s.section === "practice");
  const all = [...theory, ...practice];

  const MONTHS_RU = [
    "ЯНВАРЬ",
    "ФЕВРАЛЬ",
    "МАРТ",
    "АПРЕЛЬ",
    "МАЙ",
    "ИЮНЬ",
    "ИЮЛЬ",
    "АВГУСТ",
    "СЕНТЯБРЬ",
    "ОКТЯБРЬ",
    "НОЯБРЬ",
    "ДЕКАБРЬ",
  ];
  const [y, mn] = String(data.month || "")
    .split("-")
    .map(Number);
  const monthUp = mn ? MONTHS_RU[mn - 1] : "";
  const studyYear = y && mn ? `${y}-${y + 1}` : String(data.month || "");

  function esc(v) {
    return String(v || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }
  function getGrade(grades, key) {
    if (!grades) return "";
    if (typeof grades.get === "function") return grades.get(key) || "";
    return grades[key] || "";
  }

  // A4 landscape, margins from original template
  const PW = 16838,
    PH = 11906;
  const ML = 1134,
    MR = 851,
    MT = 1134,
    MB = 1134;
  const usableTwips = PW - ML - MR;
  const COL_NO = 530;
  const COL_FIO = 2600;
  const restTwips = usableTwips - COL_NO - COL_FIO;
  const COL_SUBJ = Math.max(
    280,
    Math.floor(restTwips / Math.max(all.length, 1)),
  );

  const RFonts = `<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>`;

  // tcBorders (all sides)
  const BDR = `<w:tcBorders>
          <w:top    w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:left   w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:right  w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        </w:tcBorders>`;

  // CORRECT tcPr order: tcW → gridSpan → vMerge → tcBorders → shd → textDirection → vAlign
  // CORRECT pPr order: spacing → jc
  function tc(text, width, opts = {}) {
    const span = opts.span ? `<w:gridSpan w:val="${opts.span}"/>` : "";
    const vmStart = opts.vmStart ? `<w:vMerge w:val="restart"/>` : "";
    const vmCont = opts.vmCont ? `<w:vMerge/>` : "";
    const shade = opts.fill
      ? `<w:shd w:val="clear" w:color="auto" w:fill="${opts.fill}"/>`
      : "";
    const tdir = opts.vert ? `<w:textDirection w:val="btLr"/>` : "";
    const va = opts.va || (opts.vert ? "bottom" : "center");
    const jc = opts.jc || "center";
    const sz = opts.sz || 18;
    const bold = opts.bold ? "<w:b/><w:bCs/>" : "";
    const italic = opts.italic ? "<w:i/><w:iCs/>" : "";

    // pPr: spacing BEFORE jc (OOXML spec order)
    const pPr = `<w:pPr>
            <w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>
            <w:jc w:val="${jc}"/>
          </w:pPr>`;

    // tcPr: tcW → gridSpan → vMerge → tcBorders → shd → textDirection → vAlign
    return `<w:tc>
        <w:tcPr>
          <w:tcW w:w="${width}" w:type="dxa"/>
          ${span}${vmStart}${vmCont}${BDR}${shade}${tdir}
          <w:vAlign w:val="${va}"/>
        </w:tcPr>
        <w:p>${pPr}<w:r>
            <w:rPr>${RFonts}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>${bold}${italic}</w:rPr>
            <w:t xml:space="preserve">${esc(text)}</w:t>
          </w:r>
        </w:p>
      </w:tc>`;
  }

  // HEADER ROW 1
  const hr1 = `<w:tr>
      <w:trPr><w:trHeight w:val="500" w:hRule="atLeast"/><w:cantSplit/></w:trPr>
      ${tc("№\nп/п", COL_NO, { bold: true, sz: 18, vmStart: true })}
      ${tc("Фамилия, имя и отчество", COL_FIO, { bold: true, sz: 18, vmStart: true, jc: "left" })}
      ${
        theory.length
          ? tc("Предметы теоретического обучения", COL_SUBJ * theory.length, {
              bold: true,
              sz: 18,
              span: theory.length,
              fill: "EEF2FF",
            })
          : ""
      }
      ${
        practice.length
          ? tc(
              "Производственное обучение / Спец. предметы",
              COL_SUBJ * practice.length,
              { bold: true, sz: 18, span: practice.length, fill: "FEF3C7" },
            )
          : ""
      }
    </w:tr>`;

  // HEADER ROW 2 (subject names vertical)
  const hr2 = `<w:tr>
      <w:trPr><w:trHeight w:val="1700" w:hRule="exact"/><w:cantSplit/></w:trPr>
      ${tc("", COL_NO, { vmCont: true })}
      ${tc("", COL_FIO, { vmCont: true })}
      ${theory.map((s) => tc(s.label, COL_SUBJ, { bold: true, sz: 16, vert: true, fill: "EEF2FF" })).join("")}
      ${practice.map((s) => tc(s.label, COL_SUBJ, { bold: true, sz: 16, vert: true, fill: "FEF9E7" })).join("")}
    </w:tr>`;

  // DATA ROWS
  const GRADE_FILL = { 5: "DCFCE7", 4: "DBEAFE", 2: "FEE2E2" };
  const dataRows = data.rows
    .map((row, i) => {
      const gradeCells = all
        .map((s) => {
          const g = getGrade(row.grades, s.key);
          const fill = GRADE_FILL[g] || "FFFFFF";
          return tc(g, COL_SUBJ, { sz: 20, bold: !!g, fill });
        })
        .join("");
      return `<w:tr>
        <w:trPr><w:trHeight w:val="350" w:hRule="atLeast"/></w:trPr>
        ${tc(`${i + 1}.`, COL_NO, { sz: 18 })}
        ${tc(row.fio, COL_FIO, { sz: 18, jc: "left" })}
        ${gradeCells}
      </w:tr>`;
    })
    .join("\n");

  // GRID columns
  let grid = `<w:gridCol w:w="${COL_NO}"/><w:gridCol w:w="${COL_FIO}"/>`;
  for (let i = 0; i < all.length; i++) grid += `<w:gridCol w:w="${COL_SUBJ}"/>`;

  // CORRECT tblPr order: tblW → tblBorders → tblLayout (per OOXML spec)
  const tableXml = `<w:tbl>
      <w:tblPr>
        <w:tblW w:w="0" w:type="auto"/>
        <w:tblBorders>
          <w:top    w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:left   w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:right  w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
          <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        </w:tblBorders>
        <w:tblLayout w:type="fixed"/>
      </w:tblPr>
      <w:tblGrid>${grid}</w:tblGrid>
      ${hr1}${hr2}${dataRows}
    </w:tbl>`;

  function para(text, opts = {}) {
    const sz = opts.sz || 24;
    const bold = opts.bold ? "<w:b/><w:bCs/>" : "";
    const ital = opts.ital ? "<w:i/><w:iCs/>" : "";
    const after = opts.after !== undefined ? opts.after : 0;
    // pPr: spacing BEFORE jc
    return `<w:p>
      <w:pPr>
        <w:spacing w:before="0" w:after="${after}"/>
        <w:jc w:val="center"/>
      </w:pPr>
      <w:r>
        <w:rPr>${RFonts}<w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>${bold}${ital}</w:rPr>
        <w:t xml:space="preserve">${esc(text)}</w:t>
      </w:r>
    </w:p>`;
  }

  const signaturesXml = `
    <w:p><w:pPr><w:spacing w:before="280" w:after="0"/><w:jc w:val="left"/></w:pPr></w:p>
    <w:p>
      <w:pPr><w:spacing w:before="0" w:after="80"/><w:jc w:val="left"/></w:pPr>
      <w:r><w:rPr>${RFonts}<w:sz w:val="22"/><w:szCs w:val="22"/><w:i/><w:iCs/></w:rPr>
        <w:t>Зам. директора по УПМР _____________________</w:t>
      </w:r>
    </w:p>
    <w:p>
      <w:pPr><w:spacing w:before="0" w:after="40"/><w:jc w:val="left"/></w:pPr>
      <w:r><w:rPr>${RFonts}<w:sz w:val="22"/><w:szCs w:val="22"/><w:i/><w:iCs/></w:rPr>
        <w:t xml:space="preserve">Куратор группы: ${esc(data.curatorFullName || "_____________________")}</w:t>
      </w:r>
    </w:p>`;

  const bodyXml = [
    para("СВОДНАЯ ВЕДОМОСТЬ", { bold: true, sz: 32, after: 80 }),
    para(`успеваемости учащихся группы «${esc(data.group.name)}»`, {
      bold: true,
      sz: 24,
    }),
    para(`Специальность: «${esc(data.group.profRu || "Не указана")}»`, {
      bold: true,
      ital: true,
      sz: 24,
    }),
    para(`ПЛ № 3 за ${monthUp} ${studyYear} гг.`, {
      bold: true,
      sz: 24,
      after: 200,
    }),
    tableXml,
    signaturesXml,
  ].join("\n");

  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
  mc:Ignorable="w14">
<w:body>
${bodyXml}
<w:sectPr>
  <w:pgSz w:w="${PW}" w:h="${PH}" w:orient="landscape"/>
  <w:pgMar w:top="${MT}" w:right="${MR}" w:bottom="${MB}" w:left="${ML}"
           w:header="0" w:footer="0" w:gutter="0"/>
</w:sectPr>
</w:body>
</w:document>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults><w:rPrDefault><w:rPr>
    <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
    <w:sz w:val="20"/><w:szCs w:val="20"/>
  </w:rPr></w:rPrDefault></w:docDefaults>
</w:styles>`;

  const zip = new PizZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    Target="styles.xml"/>
</Relationships>`,
  );
  zip.file("word/document.xml", documentXml);
  zip.file("word/styles.xml", stylesXml);
  return zip.generate({ type: "nodebuffer", compression: "DEFLATE" });
}

/* ===========================
   СКАЧАТЬ WORD .docx СВОДНОЙ ВЕДОМОСТИ
=========================== */
app.get(
  "/api/admin/summary-word/:gid/:month",
  requireGroupAccess("gid"),
  async (req, res) => {
    try {
      const data = await getSummaryDocumentData(
        req.params.gid,
        req.params.month,
      );
      const buf = buildSummaryDocx(data);
      const sn = String(data.group.name || "group").replace(/[^\w\s-]/g, "_");
      res.setHeader(
        "Content-Disposition",
        contentDisposition(`Vedomost_${sn}_${req.params.month}.docx`),
      );
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      );
      res.send(buf);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);

/* ===========================
   ОТПРАВИТЬ ВЕДОМОСТЬ В TELEGRAM (.docx)
=========================== */
app.post(
  "/api/admin/summary-send-tg/:gid/:month",
  requireGroupAccess("gid"),
  async (req, res) => {
    try {
      const { gid, month } = req.params;
      const data = await getSummaryDocumentData(gid, month);
      const buf = buildSummaryDocx(data);
      const sn = String(data.group.name || "group").replace(
        /[^a-zA-Z0-9._-]/g,
        "_",
      );
      const filename = `Vedomost_${sn}_${month}.docx`;
      const MONTHS_RU = [
        "Январь",
        "Февраль",
        "Март",
        "Апрель",
        "Май",
        "Июнь",
        "Июль",
        "Август",
        "Сентябрь",
        "Октябрь",
        "Ноябрь",
        "Декабрь",
      ];
      const [y, mn] = String(month || "")
        .split("-")
        .map(Number);
      const monthLabel = mn ? `${MONTHS_RU[mn - 1]} ${y}` : month;
      const subs = await TgSub.find({ role: "admin" });
      if (!subs.length)
        return res
          .status(404)
          .json({ error: "Нет администраторов в Telegram" });
      const caption = `📝 *Сводная ведомость*\n📂 Группа: *${data.group.name}*\n🎓 ${data.group.profRu || "—"}\n📅 ${monthLabel}\n👥 Студентов: *${data.rows.length}*`;
      let sent = 0,
        failed = 0;
      for (const sub of subs) {
        try {
          await bot.sendDocument(
            sub.chatId,
            buf,
            { caption, parse_mode: "Markdown" },
            {
              filename,
              contentType:
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            },
          );
          sent++;
        } catch (e) {
          failed++;
        }
      }
      res.json({ ok: true, sent, failed, total: subs.length });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);

/* ===========================
   ПОЛУЧИТЬ СПИСОК СОХРАНЁННЫХ МЕСЯЦЕВ ДЛЯ ГРУППЫ
=========================== */
app.get(
  "/api/admin/summary-months/:gid",
  requireGroupAccess("gid"),
  async (req, res) => {
    try {
      const records = await SummaryRecord.find({
        group: req.params.gid,
      }).distinct("month");
      records.sort().reverse();
      res.json(records);
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  },
);

// Все месяцы по всем группам (для поиска по месяцу)
app.get("/api/admin/summary-all-months", requireAuth, async (req, res) => {
  try {
    const ids = getAccessibleGroupIds(req);
    const filter =
      ids === null ? {} : ids.length ? { group: { $in: ids } } : {};
    const months = await SummaryRecord.distinct("month", filter);
    months.sort().reverse();
    res.json(months);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Группы и данные ведомости за конкретный месяц
app.get("/api/admin/summary-by-month/:month", requireAuth, async (req, res) => {
  try {
    const month = req.params.month;
    const ids = getAccessibleGroupIds(req);
    // Find all groups that have records for this month
    const filter = {
      month,
      ...(ids !== null && ids.length ? { group: { $in: ids } } : {}),
    };
    const groupIds = await SummaryRecord.distinct("group", filter);
    if (!groupIds.length) return res.json([]);
    const groups = await Group.find({ _id: { $in: groupIds } }).sort({
      name: 1,
    });
    const result = await Promise.all(
      groups.map(async (g) => {
        const students = await Student.find({ group: g._id }).select("_id");
        const subjects = await ensureGroupSummarySubjects(g);
        const records = await SummaryRecord.find({
          group: g._id,
          month,
        }).lean();
        let filled = 0;
        records.forEach((r) => {
          const grades =
            r.grades instanceof Map
              ? Object.fromEntries(r.grades)
              : r.grades && typeof r.grades === "object"
                ? r.grades
                : {};
          subjects.forEach((s) => {
            if (grades[s.key]) filled++;
          });
        });
        const total = students.length * subjects.length;
        const pct = total > 0 ? Math.round((filled / total) * 100) : 0;
        return {
          groupId: String(g._id),
          groupName: g.name,
          profRu: g.profRu || "—",
          students: students.length,
          subjects: subjects.length,
          filled,
          total,
          pct,
        };
      }),
    );
    res.json(result);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* ================================================================
   TELEGRAM BOT — 3 уровня доступа (Админ | Куратор | Родитель)
   Подключается из bot.js
================================================================ */
const createBot = require("./bot");

// SKUD events buffer (shared with bot)
const SKUD_EVENTS_REF = SKUD_EVENTS;

const _botExports = createBot(
  // models
  { Group, Student, TgSub, Curator, SummaryRecord },
  // helpers
  {
    pad2,
    getBishkekNow,
    getTodayKey,
    ensureGroupSummarySubjects,
    getSummaryDocumentData,
    formatMonthLabel,
    getSkudEvents: () => SKUD_EVENTS_REF,
  },
  // buildSummaryDocx
  buildSummaryDocx,
);
const {
  bot,
  notifyAll: _notifyAll,
  notifyParents: _notifyParents,
  notifyAdmins: _notifyAdmins,
  notifyCurators: _notifyCurators,
} = _botExports;
// Wire real functions from bot (replacing stubs)
notifyAll = _notifyAll;
notifyParents = _notifyParents;
notifyAdmins = _notifyAdmins;
notifyCuratorsOfGroup = _notifyCurators;

/* ================================================================
   TELEGRAM MINI APP API  — /api/webapp/*
   Аутентификация через chatId из заголовка x-tg-chat-id
================================================================ */

async function resolveWebappUser(req) {
  const chatId = String(req.headers["x-tg-chat-id"] || "").trim();
  if (!chatId) return null;
  return await TgSub.findOne({ chatId });
}

// Кто я?
app.get("/api/webapp/me", async (req, res) => {
  try {
    const sub = await resolveWebappUser(req);
    if (!sub) return res.json({ role: null });

    if (sub.role === "admin") {
      const groups = await Group.find().sort({ name: 1 }).select("name profRu shiftStart");
      return res.json({
        role: "admin",
        groups: groups.map((g) => ({
          id: String(g._id),
          name: g.name,
          profRu: g.profRu || "",
          shiftStart: g.shiftStart || "09:30",
        })),
      });
    }

    if (sub.role === "curator" && sub.curatorId) {
      const cur = await Curator.findById(sub.curatorId).populate(
        "groups",
        "name profRu shiftStart",
      );
      if (!cur) return res.json({ role: null });
      return res.json({
        role: "curator",
        curatorName: cur.fullName || cur.username || "",
        groups: cur.groups.map((g) => ({
          id: String(g._id),
          name: g.name,
          profRu: g.profRu || "",
          shiftStart: g.shiftStart || "09:30",
        })),
      });
    }

    if (sub.role === "parent" && sub.linkedInns?.length) {
      const children = await Promise.all(
        sub.linkedInns.map(async (inn) => {
          const stu = await Student.findOne({ inn }).populate("group", "name profRu");
          if (!stu) return null;
          return {
            inn,
            fio: stu.fio || "",
            groupId: String(stu.group?._id || ""),
            groupName: stu.group?.name || "",
            profRu: stu.group?.profRu || "",
          };
        }),
      );
      return res.json({ role: "parent", children: children.filter(Boolean) });
    }

    return res.json({ role: null });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Посещаемость сегодня для куратора
app.get("/api/webapp/attendance/today", async (req, res) => {
  try {
    const sub = await resolveWebappUser(req);
    if (!sub) return res.status(401).json({ error: "Unauthorized" });
    const gid = req.query.gid;
    if (sub.role === "curator") {
      const cur = await Curator.findById(sub.curatorId);
      const gids = cur?.groups?.map(String) || [];
      if (!gids.includes(gid))
        return res.status(403).json({ error: "No access" });
    }
    const tdk = getTodayKey();
    const g = await Group.findById(gid);
    const [sH, sM] = (g?.shiftStart || "09:30").split(":").map(Number),
      stm = sH * 60 + sM;
    const students = await Student.find({ group: gid })
      .sort({ fio: 1 })
      .select("fio inn skudDaily attendance");
    let present = 0,
      absent = 0,
      late = 0;
    const list = students.map((stu) => {
      const skud = stu.skudDaily?.find((r) => r.date === tdk),
        att = stu.attendance?.find((a) => a.date === tdk);
      let here = false,
        isLate = false,
        time = "";
      if (skud) {
        here = skud.present;
        if (skud.firstIn) {
          const dt = new Date(skud.firstIn);
          if (!isNaN(dt)) {
            const tm = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
            const h = Math.floor(tm / 60) % 24,
              m = tm % 60;
            time = `${pad2(h)}:${pad2(m)}`;
            isLate = h * 60 + m > stm;
          }
        }
      } else if (att) {
        here = att.present;
      }
      if (here && isLate) {
        present++;
        late++;
      } else if (here) {
        present++;
      } else {
        absent++;
      }
      return {
        fio: stu.fio,
        status: here ? (isLate ? "late" : "present") : "absent",
        time,
      };
    });
    res.json({ present, absent, late, students: list });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Посещаемость группы за месяц
app.get("/api/webapp/attendance/month", async (req, res) => {
  try {
    const sub = await resolveWebappUser(req);
    if (!sub) return res.status(401).json({ error: "Unauthorized" });
    const { gid, month } = req.query;
    if (sub.role === "curator") {
      const cur = await Curator.findById(sub.curatorId);
      const gids = cur?.groups?.map(String) || [];
      if (!gids.includes(gid))
        return res.status(403).json({ error: "No access" });
    }
    const [yr, mn] = (month || "").split("-").map(Number);
    if (!yr || !mn) return res.status(400).json({ error: "Invalid month" });
    const dim = new Date(yr, mn, 0).getDate();
    const days = [];
    for (let d = 1; d <= dim; d++) days.push(`${yr}-${pad2(mn)}-${pad2(d)}`);
    const g = await Group.findById(gid);
    const [sH, sM] = (g?.shiftStart || "09:30").split(":").map(Number);
    const stm = sH * 60 + sM;
    const students = await Student.find({ group: gid })
      .sort({ fio: 1 })
      .select("fio skudDaily attendance");
    const result = students.map((stu) => {
      const statuses = {};
      for (const day of days) {
        const skud = stu.skudDaily?.find((r) => r.date === day);
        const att = stu.attendance?.find((a) => a.date === day);
        let status = "a";
        if (skud?.present) {
          if (skud.firstIn) {
            const dt = new Date(skud.firstIn);
            if (!isNaN(dt)) {
              const tm = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
              const h = Math.floor(tm / 60) % 24, m = tm % 60;
              status = h * 60 + m > stm ? "l" : "p";
            } else {
              status = "p";
            }
          } else {
            status = "p";
          }
        } else if (att?.present) {
          status = "p";
        }
        statuses[day] = status;
      }
      return { fio: stu.fio, statuses };
    });
    res.json({ days, students: result });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Предметы группы
app.get("/api/webapp/subjects", async (req, res) => {
  try {
    const g = await Group.findById(req.query.gid);
    if (!g) return res.status(404).json({ error: "Group not found" });
    const subjects = await ensureGroupSummarySubjects(g);
    res.json({ subjects });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Студенты группы
app.get("/api/webapp/students", async (req, res) => {
  try {
    const stus = await Student.find({ group: req.query.gid })
      .sort({ fio: 1 })
      .select("fio inn lyceumId birthDate");
    res.json({
      students: stus.map((s) => ({
        id: String(s._id),
        fio: s.fio,
        inn: s.inn || "",
        lyceumId: s.lyceumId || "",
        birthDate: s.birthDate || "",
      })),
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Оценки группы за месяц
app.get("/api/webapp/grades", async (req, res) => {
  try {
    const { gid, month } = req.query;
    const g = await Group.findById(gid);
    const subjects = await ensureGroupSummarySubjects(g);
    const records = await SummaryRecord.find({ group: gid, month }).lean();
    // grades: {studentId: {subjectKey: value}}
    const grades = {};
    records.forEach((r) => {
      const sid = String(r.student);
      grades[sid] = {};
      const raw = r.grades;
      if (raw instanceof Map)
        raw.forEach((v, k) => {
          if (v) grades[sid][k] = v;
        });
      else if (raw)
        Object.entries(raw).forEach(([k, v]) => {
          if (v && !k.startsWith("$")) grades[sid][k] = v;
        });
    });
    res.json({ grades, subjects });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Сохранить оценки одного студента
app.post("/api/webapp/grades/set", async (req, res) => {
  try {
    const sub = await resolveWebappUser(req);
    if (!sub || (sub.role !== "curator" && sub.role !== "admin"))
      return res.status(403).json({ error: "Forbidden" });
    const { gid, studentId, month, grades } = req.body;
    // Verify curator has access
    if (sub.role === "curator") {
      const cur = await Curator.findById(sub.curatorId);
      if (!cur?.groups?.map(String).includes(gid))
        return res.status(403).json({ error: "No access" });
    }
    await SummaryRecord.findOneAndUpdate(
      { group: gid, student: studentId, month },
      { group: gid, student: studentId, month, grades },
      { upsert: true, new: true },
    );
    // Notify parents
    const stu = await Student.findById(studentId).select("fio inn");
    if (stu?.inn) {
      const entries = Object.entries(grades || {}).filter(([, v]) => v);
      if (entries.length) {
        const g = await Group.findById(gid);
        const subjects = await ensureGroupSummarySubjects(g);
        const subj = subjects.find((s) => s.key === entries[0][0]);
        notifyParents(
          stu.inn,
          `📊 *Новые оценки\!*
👤 ${stu.fio}
📅 ${formatMonthLabel(month)}`,
        ).catch(() => {});
      }
    }
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Привязать ребёнка (родитель)
app.post("/api/webapp/parent/link", async (req, res) => {
  try {
    const { inn, chatId } = req.body;
    const clean = String(inn || "").replace(/\D/g, "");
    if (clean.length < 12)
      return res.status(400).json({ error: "ИНН минимум 12 цифр" });
    const stu = await Student.findOne({ inn: clean }).populate("group");
    if (!stu)
      return res.status(404).json({ error: "Студент с таким ИНН не найден" });
    await TgSub.findOneAndUpdate(
      { chatId: String(chatId) },
      {
        chatId: String(chatId),
        role: "parent",
        $addToSet: { linkedInns: clean },
      },
      { upsert: true, new: true },
    );
    res.json({
      ok: true,
      fio: stu.fio,
      groupName: stu.group?.name || "",
      groupId: String(stu.group?._id || ""),
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Оценки ребёнка (родитель)
app.get("/api/webapp/parent/grades", async (req, res) => {
  try {
    const { inn, month } = req.query;
    const stu = await Student.findOne({ inn }).select("_id group");
    if (!stu) return res.status(404).json({ error: "Не найдено" });
    const g = await Group.findById(stu.group);
    const subjects = await ensureGroupSummarySubjects(g);
    const record = await SummaryRecord.findOne({
      group: stu.group,
      student: stu._id,
      month,
    }).lean();
    const grades = {};
    if (record?.grades instanceof Map)
      record.grades.forEach((v, k) => {
        grades[k] = v;
      });
    else if (record?.grades)
      Object.entries(record.grades).forEach(([k, v]) => {
        if (!k.startsWith("$")) grades[k] = v;
      });
    res.json({ grades, subjects });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Посещаемость ребёнка за месяц (родитель)
app.get("/api/webapp/parent/attendance", async (req, res) => {
  try {
    const { inn, month } = req.query;
    const stu = await Student.findOne({ inn })
      .populate("group", "shiftStart")
      .select("fio skudDaily attendance group");
    if (!stu) return res.status(404).json({ error: "Не найдено" });
    const [yr, mn] = (month || "").split("-").map(Number);
    if (!yr || !mn) return res.status(400).json({ error: "Invalid month" });
    const dim = new Date(yr, mn, 0).getDate();
    const days = [];
    for (let d = 1; d <= dim; d++) days.push(`${yr}-${pad2(mn)}-${pad2(d)}`);
    const [sH, sM] = (stu.group?.shiftStart || "09:30").split(":").map(Number);
    const stm = sH * 60 + sM;
    const attMap = {};
    for (const day of days) {
      const skud = stu.skudDaily?.find((r) => r.date === day);
      const att  = stu.attendance?.find((a) => a.date === day);
      let status = "a", time = null;
      if (skud?.present) {
        if (skud.firstIn) {
          const dt = new Date(skud.firstIn);
          if (!isNaN(dt)) {
            const tm = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
            const h  = Math.floor(tm / 60) % 24, m = tm % 60;
            time     = `${pad2(h)}:${pad2(m)}`;
            status   = h * 60 + m > stm ? "l" : "p";
          } else {
            status = "p";
          }
        } else {
          status = "p";
        }
      } else if (att?.present) {
        status = "p";
      }
      attMap[day] = { status, time };
    }
    res.json({ fio: stu.fio, attMap });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Глобальная статистика сегодня (только admin)
app.get("/api/webapp/stats", async (req, res) => {
  try {
    const sub = await resolveWebappUser(req);
    if (!sub || sub.role !== "admin")
      return res.status(403).json({ error: "Admin only" });
    const tdk = getTodayKey();
    const allGroups = await Group.find().sort({ name: 1 });
    let totalAll = 0, presentAll = 0, absentAll = 0, lateAll = 0;
    const groups = await Promise.all(
      allGroups.map(async (g) => {
        const [sH, sM] = (g.shiftStart || "09:30").split(":").map(Number);
        const stm = sH * 60 + sM;
        const students = await Student.find({ group: g._id }).select(
          "skudDaily attendance",
        );
        let present = 0, absent = 0, late = 0;
        for (const stu of students) {
          const skud = stu.skudDaily?.find((r) => r.date === tdk);
          const att = stu.attendance?.find((a) => a.date === tdk);
          let here = false, isLate = false;
          if (skud) {
            here = skud.present;
            if (skud.firstIn) {
              const dt = new Date(skud.firstIn);
              if (!isNaN(dt)) {
                const tm = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
                const h = Math.floor(tm / 60) % 24, m = tm % 60;
                isLate = h * 60 + m > stm;
              }
            }
          } else if (att) {
            here = att.present;
          }
          if (here && isLate) { present++; late++; }
          else if (here) { present++; }
          else { absent++; }
        }
        totalAll += students.length;
        presentAll += present;
        absentAll += absent;
        lateAll += late;
        return {
          id: String(g._id),
          name: g.name,
          present, absent, late,
          total: students.length,
        };
      }),
    );
    res.json({ total: totalAll, present: presentAll, absent: absentAll, late: lateAll, groups });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Подать Mini App через бота — обработка кнопки
app.get("/api/webapp/open", (req, res) => {
  res.sendFile(require("path").join(__dirname, "public", "miniapp.html"));
});

app.listen(3000, () => console.log("Server running on http://localhost:3000"));
