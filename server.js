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
const att = require("./attendance-utils");
const absence = require("./absence");
const { createReports } = require("./reports");
const { createWeeklyCards } = require("./weekly-card");

// Секреты берутся ТОЛЬКО из переменных окружения (.env / Render → Environment).
// Если переменной нет, сервер не запускается и сразу говорит какой.
function requireEnv(name) {
  const v = process.env[name];
  if (!v) {
    console.error(`❌ Не задана переменная окружения ${name}`);
    process.exit(1);
  }
  return v;
}

const app = express();
app.set("trust proxy", 1); // за прокси Render: реальный IP клиента в req.ip
const upload = multer({ dest: "uploads/" });
if (!fs.existsSync("uploads")) fs.mkdirSync("uploads");
app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));

const ADMIN_LOGIN = requireEnv("ADMIN_LOGIN");
const ADMIN_PASS = requireEnv("ADMIN_PASS");
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
// Пароли кураторов: scrypt с солью ("scrypt$соль$хеш").
// Старые хеши SHA-256 без соли ещё принимаются и заменяются на scrypt при входе.
function hashPassword(p) {
  const salt = crypto.randomBytes(16).toString("hex");
  const hash = crypto.scryptSync(String(p || ""), salt, 64).toString("hex");
  return `scrypt$${salt}$${hash}`;
}
function legacySha256(p) {
  return crypto.createHash("sha256").update(String(p || "")).digest("hex");
}
function verifyPassword(p, stored) {
  stored = String(stored || "");
  if (stored.startsWith("scrypt$")) {
    const [, salt, hash] = stored.split("$");
    if (!salt || !hash) return false;
    const calc = crypto.scryptSync(String(p || ""), salt, 64);
    const want = Buffer.from(hash, "hex");
    return want.length === calc.length && crypto.timingSafeEqual(calc, want);
  }
  return !!stored && safeEq(stored, legacySha256(p));
}
const isLegacyHash = (stored) => !!stored && !String(stored).startsWith("scrypt$");

// Сессия подписана HMAC: подделать cookie без SESSION_SECRET нельзя.
// Без SESSION_SECRET ключ выводится из других секретов (см. .env.example), но лучше задать его явно.
const SESSION_SECRET =
  process.env.SESSION_SECRET ||
  crypto
    .createHash("sha256")
    .update(
      [ADMIN_PASS, process.env.MONGO_URI, process.env.AGENT_API_KEY].join("|"),
    )
    .digest("hex");
if (!process.env.SESSION_SECRET)
  console.warn(
    "⚠️ SESSION_SECRET не задан: ключ сессий выведен из других секретов. Задайте SESSION_SECRET в .env / Render.",
  );
const SESSION_MS = 30 * 24 * 60 * 60 * 1000;
const sign = (payload) =>
  crypto.createHmac("sha256", SESSION_SECRET).update(payload).digest("base64url");
function safeEq(a, b) {
  const h = (x) =>
    crypto
      .createHash("sha256")
      .update(String(x ?? ""))
      .digest();
  return crypto.timingSafeEqual(h(a), h(b));
}

// Проверенная сессия запроса (см. refreshSession ниже) имеет приоритет над cookie
function getSession(req) {
  if (req._sess !== undefined) return req._sess;
  return readSessionCookie(req);
}
function readSessionCookie(req) {
  const raw = getCookie(req, AUTH_COOKIE);
  if (!raw) return null;
  try {
    const [payload, sig] = decodeURIComponent(raw).split(".");
    if (!payload || !sig || !safeEq(sig, sign(payload))) return null;
    const s = JSON.parse(Buffer.from(payload, "base64url").toString("utf8"));
    if (!s.exp || s.exp < Date.now()) return null;
    return s;
  } catch {
    return null;
  }
}
function setAuthCookie(res, session) {
  const payload = Buffer.from(
    JSON.stringify({ ...session, exp: Date.now() + SESSION_MS }),
    "utf8",
  ).toString("base64url");
  res.cookie(AUTH_COOKIE, `${payload}.${sign(payload)}`, {
    maxAge: SESSION_MS,
    httpOnly: true,
    sameSite: "lax",
    secure: !!process.env.RENDER,
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

// Сессии кураторов сверяются с базой: изменения групп действуют сразу,
// а после смены логина/пароля, «выйти везде» или удаления куратора старый вход перестаёт работать.
const curatorCache = new Map(); // id → { at, doc }
const CURATOR_CACHE_MS = 15 * 1000;
function forgetCuratorCache(id) {
  if (id) curatorCache.delete(String(id));
  else curatorCache.clear();
}
async function loadCuratorCached(id) {
  const hit = curatorCache.get(String(id));
  if (hit && Date.now() - hit.at < CURATOR_CACHE_MS) return hit.doc;
  const doc = mongoose.Types.ObjectId.isValid(id)
    ? await Curator.findById(id).populate("groups", "name").lean()
    : null;
  curatorCache.set(String(id), { at: Date.now(), doc });
  return doc;
}
app.use(async (req, res, next) => {
  const s = readSessionCookie(req);
  if (!s || s.role !== "curator") return next();
  try {
    const cur = await loadCuratorCached(s.curatorId);
    if (!cur || (cur.sessionVer || 0) !== (s.v || 0) || cur.username !== s.username) {
      req._sess = null;
      res.clearCookie(AUTH_COOKIE);
      return next();
    }
    const groups = (cur.groups || []).filter(Boolean);
    req._sess = {
      ...s,
      groupIds: groups.map((g) => String(g._id)),
      groupId: groups[0] ? String(groups[0]._id) : null,
      groupNames: groups.map((g) => g.name),
      groupName: groups[0]?.name || "",
    };
  } catch (e) {
    // база недоступна — пускаем по cookie, как раньше
  }
  next();
});

app.get("/login", (req, res) => {
  const s = getSession(req);
  if (s?.role === "admin") return res.redirect("/");
  if (s?.role === "curator") return res.redirect("/summary.html");
  const err = req.query.error
    ? '<div style="color:#dc2626;font-size:13px;margin-bottom:15px;text-align:center;background:#fee2e2;padding:10px;border-radius:8px;">Неверный логин или пароль!</div>'
    : "";
  res.send(
    `<!DOCTYPE html><html lang="ru"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Вход | ПЛ №3</title><link rel="manifest" href="/manifest.webmanifest"><meta name="theme-color" content="#13285c"><link rel="icon" type="image/png" href="/icons/favicon-32.png"><link rel="apple-touch-icon" href="/icons/apple-touch-icon.png"><meta name="apple-mobile-web-app-capable" content="yes"><meta name="apple-mobile-web-app-title" content="ПЛ №3"><script src="/pwa.js" defer></script><style>body{background:#f1f5f9;font-family:"Inter","Segoe UI",sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}.c{background:white;padding:40px 30px;border-radius:20px;box-shadow:0 10px 25px rgba(0,0,0,.05);width:100%;max-width:380px;border:1px solid #e2e8f0}h2{margin:0;font-size:24px;color:#0f172a;font-weight:800;text-align:center}p{margin:5px 0 30px;color:#64748b;font-size:14px;text-align:center}.fg{margin-bottom:20px}label{display:block;margin-bottom:8px;font-size:13px;font-weight:600;color:#475569}input{width:100%;padding:12px 15px;border:2px solid #e2e8f0;border-radius:10px;font-size:15px;outline:none;box-sizing:border-box}input:focus{border-color:#4f46e5}button{width:100%;padding:14px;background:#4f46e5;color:white;border:none;border-radius:10px;font-size:16px;font-weight:bold;cursor:pointer}button:hover{background:#4338ca}</style></head><body><div class="c"><h2>ПРОФЛИЦЕЙ №3</h2><p>Система управления</p>${err}<form action="/login" method="POST"><div class="fg"><label>Логин</label><input type="text" name="username" required autocomplete="off"></div><div class="fg"><label>Пароль</label><input type="password" name="password" required></div><button type="submit">Войти</button></form></div></body></html>`,
  );
});

// Защита от подбора пароля: не больше 10 неудачных входов с одного IP за 15 минут
const loginFails = new Map();
const failsOf = (ip) =>
  (loginFails.get(ip) || []).filter((t) => Date.now() - t < 15 * 60 * 1000);
setInterval(() => {
  for (const ip of loginFails.keys()) if (!failsOf(ip).length) loginFails.delete(ip);
}, 10 * 60 * 1000).unref();

app.post("/login", (req, res) => {
  const ip = req.ip;
  if (failsOf(ip).length >= 10)
    return res
      .status(429)
      .send("Слишком много попыток входа. Подождите 15 минут.");
  const bad = () => {
    loginFails.set(ip, [...failsOf(ip), Date.now()]);
    return res.redirect("/login?error=1");
  };
  const { username, password } = req.body;
  if (safeEq(username, ADMIN_LOGIN) && safeEq(password, ADMIN_PASS)) {
    loginFails.delete(ip);
    setAuthCookie(res, { role: "admin", username: ADMIN_LOGIN });
    return res.redirect("/");
  }
  Curator.findOne({ username: String(username || "").trim() })
    .populate("groups")
    .then((curator) => {
      if (
        !curator ||
        !verifyPassword(password, curator.passwordHash) ||
        !curator.groups?.length
      )
        return bad();
      loginFails.delete(ip);
      if (isLegacyHash(curator.passwordHash))
        Curator.updateOne(
          { _id: curator._id },
          { $set: { passwordHash: hashPassword(password) } },
        ).catch(() => null);
      const groups = curator.groups.map((g) => ({
        id: String(g._id),
        name: g.name,
      }));
      setAuthCookie(res, {
        role: "curator",
        v: curator.sessionVer || 0,
        username: curator.username,
        curatorId: String(curator._id),
        groupIds: groups.map((g) => g.id),
        groupId: groups[0]?.id || null,
        groupNames: groups.map((g) => g.name),
        groupName: groups[0]?.name || "",
      });
      res.redirect("/summary.html");
    })
    .catch(() => bad());
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

const MONGO_URI = requireEnv("MONGO_URI");
mongoose
  .connect(MONGO_URI)
  .then(async () => {
    console.log("MongoDB connected");
    restoreSkudEvents();
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

const AGENT_API_KEY = requireEnv("AGENT_API_KEY");

const PROFESSIONS = [
  { id: 1, ru: "Токарь", kg: "Токарь", duration: "2 года" },
  { id: 2, ru: "Электрогазосварщик", kg: "Электргазоширетүүчү", duration: "2 года" },
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
  { id: 7, ru: "Оператор цифровой печати", kg: "Санариптик басма оператору", duration: "2 года" },
  { id: 8, ru: "Повар", kg: "Ашпозчу", duration: "2 года" },
  { id: 9, ru: "Переплетчик", kg: "Түптөөчү", duration: "10 месяцев" },
  { id: 10, ru: "Электромонтер (10 м.)", kg: "Электромонтер", duration: "10 месяцев" },
  { id: 11, ru: "Автослесарь-Автоэлектрик", kg: "Автослесарь-Автоэлектрик", duration: "10 месяцев" },
  { id: 12, ru: "Программист", kg: "Программист", duration: "10 месяцев" },
  { id: 13, ru: "Оператор печатного оборудования", kg: "Басма жабдууларын оператору", duration: "2 года" },
  { id: 14, ru: "Компютерная диагностика", kg: "Компьютердик диагностика", duration: "10 месяцев" },
  { id: 15, ru: "Системный администратор", kg: "Системалык администратор", duration: "10 месяцев" },
  {
    id: 16,
    ru: "Разработчик Web и мультимедийных приложений",
    kg: "Web жана мультимедиалык тиркемелерди иштеп чыгуучу",
    duration: "10 месяцев",
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
    shiftStart: { type: String, default: "08:00" },
    saturday: { type: Boolean, default: false }, // учебная суббота
    // students: обычная группа; teachers/masters/admin/guests: сотрудники и гости СКУД
    kind: {
      type: String,
      enum: ["students", "teachers", "masters", "admin", "guests"],
      default: "students",
    },
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
    fio: { type: String, index: true },
    birthDate: String,
    inn: { type: String, index: true },
    lyceumId: String,
    skudConsentAt: Date, // согласие на обработку биометрии (СКУД)
    skudConsentBy: String,
    group: { type: mongoose.Schema.Types.ObjectId, ref: "Group", index: true },
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
          manual: Boolean, // поставлено вручную администратором
          manualBy: String,
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
      // растёт при смене логина/пароля или «выйти везде» — старые входы перестают действовать
      sessionVer: { type: Number, default: 0 },
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
let forgetBotCurator = () => {};

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
// Экранирование для Telegram MarkdownV2
function mdv2(v) {
  return String(v ?? "").replace(/[_*[\]()~`>#+=|{}.!\-\\]/g, "\\$&");
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
  "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
  "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
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

// Лента событий СКУД: последние 500 в памяти + копия в Mongo (хранится 14 дней),
// чтобы после перезапуска Render лента и "опоздавшие сегодня" не пропадали.
const SkudEventLog = mongoose.model(
  "SkudEventLog",
  new mongoose.Schema(
    { ev: Object, createdAt: { type: Date, default: Date.now, expires: 14 * 24 * 3600 } },
    { versionKey: false },
  ),
);
function pushSkudEvent(ev) {
  SKUD_EVENTS.unshift(ev);
  if (SKUD_EVENTS.length > 500) SKUD_EVENTS.pop();
  SkudEventLog.create({ ev }).catch(() => null);
}
async function restoreSkudEvents() {
  try {
    const last = await SkudEventLog.find().sort({ createdAt: -1 }).limit(500).lean();
    if (!SKUD_EVENTS.length) SKUD_EVENTS.push(...last.map((d) => d.ev));
    console.log(`🗂 Лента СКУД восстановлена: ${last.length}`);
  } catch (e) {
    console.log("skud events restore:", e.message);
  }
}

/* === ЖУРНАЛ: доступ по входу или по секретной ссылке для просмотра === */
// Ссылка /journal.html?k=<ключ> даёт только просмотр журнала (без ИНН и комментариев).
// Ключ выводится из SESSION_SECRET: смена SESSION_SECRET отзывает старые ссылки.
const JOURNAL_KEY = sign("journal-share-v1");
function journalAccess(req, res, next) {
  const s = getSession(req);
  if (s?.role) {
    req.auth = s;
    return next();
  }
  if (req.query.k && safeEq(req.query.k, JOURNAL_KEY)) {
    req.journalShare = true;
    return next();
  }
  res.status(401).json({ error: "Нужен вход или ссылка для просмотра" });
}

app.get("/journal.html", (req, res) =>
  res.sendFile(path.join(__dirname, "public", "journal.html")),
);

app.get("/api/admin/journal-link", requireAuth, requireAdmin, (req, res) =>
  res.json({ url: `/journal.html?k=${JOURNAL_KEY}` }),
);

app.get("/api/admin/groups-list", journalAccess, async (req, res) => {
  try {
    const ids = req.journalShare ? null : getAccessibleGroupIds(req);
    const filter = ids === null ? {} : { _id: { $in: ids } };
    res.json(await Group.find(filter).sort({ name: 1 }));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/admin/global-stats", requireAuth, requireAdmin, async (req, res) => {
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

app.get("/api/admin/attendance-matrix/:gid/:month", journalAccess, async (req, res) => {
  try {
    if (!req.journalShare && !canAccessGroup(req, req.params.gid))
      return res.status(403).json({ error: "Нет доступа к этой группе" });
    const { gid, month } = req.params;
    if (!/^\d{4}-\d{2}$/.test(month))
      return res.status(400).json({ error: "Неверный месяц" });
    const { from, to } = att.monthRange(month);
    const dim = Number(to.slice(8));
    const group = await Group.findById(gid),
      shiftTime = group?.shiftStart || "09:30",
      shiftMin = att.shiftMinutes(shiftTime);
    const students = await Student.find({ group: gid })
      .sort({ fio: 1 })
      .select("fio skudDaily attendance")
      .lean();
    const reasons = await absence.reasonsIndex({ group: gid }, from, to);
    const tdk = getTodayKey(),
      sat = !!group?.saturday,
      tiw = att.isWeekend(tdk, sat);
    let ptc = 0,
      tml = 0;
    const matrix = students.map((s) => {
      const days = [];
      let ipt = false,
        sl = 0,
        sp = 0,
        sa = 0,
        se = 0;
      for (let d = 1; d <= dim; d++) {
        const dk = `${month}-${pad2(d)}`,
          iw = att.isWeekend(dk, sat);
        const st = att.dayStatus(s, dk, shiftMin);
        let reason = null;
        if (!st.present && !iw && dk <= tdk) {
          const r = reasons.get(`${s._id}|${dk}`);
          if (r) {
            reason = {
              id: r.id,
              code: r.code,
              label: r.label,
              excused: r.excused,
              confirmed: r.confirmed,
              comment: req.journalShare ? "" : r.comment,
              dateFrom: r.dateFrom,
              dateTo: r.dateTo,
              type: r.type,
            };
            if (r.excused && r.confirmed) se++;
            else sa++;
          } else sa++;
        }
        if (st.present) sp++;
        if (st.late) {
          sl++;
          tml++;
        }
        if (dk === tdk && st.present) ipt = true;
        days.push({ date: dk, present: st.present, late: st.late, time: st.time, isWeekend: iw, absent: !st.present && !iw && dk <= tdk, reason });
      }
      if (ipt) ptc++;
      return {
        id: String(s._id),
        fio: s.fio,
        days,
        stats: { presences: sp, lates: sl, absences: sa, excused: se },
      };
    });
    res.json({
      daysInMonth: dim,
      matrix,
      shiftStart: shiftTime,
      saturday: sat,
      canEdit: !req.journalShare,
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
  if (!safeEq(apiKey, AGENT_API_KEY)) return res.status(403).send("Forbidden");
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
      if (inn) student = await Student.findOne({ inn }).select("fio inn group");
      if (!student && fioNorm) {
        const fp = fioNorm.split(" ")[0],
          cands = await Student.find({ fio: new RegExp("^" + fp.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i") })
            .select("fio inn group")
            .limit(50);
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
      // Атомарные обновления: не перезаписываем массивы целиком и не теряем
      // события, если два запроса агента пришли одновременно.
      const sid = student._id;
      await Student.updateOne(
        { _id: sid, "attendance.date": { $ne: dateKey } },
        { $push: { attendance: { date: dateKey, present: true } } },
      );
      await Student.updateOne(
        { _id: sid, "attendance.date": dateKey },
        { $set: { "attendance.$.present": true } },
      );
      await Student.updateOne(
        { _id: sid, "skudDaily.date": { $ne: dateKey } },
        { $push: { skudDaily: { date: dateKey, inCount: 0, outCount: 0, present: false } } },
      );
      const local = att.localTime(eventIso);
      const timeStr = local ? local.str : "—";
      if (direction === "in") {
        // старая версия записи дня: по ней понимаем, первый ли это вход
        const prev = await Student.findOneAndUpdate(
          { _id: sid, "skudDaily.date": dateKey },
          {
            $inc: { "skudDaily.$.inCount": 1 },
            $set: { "skudDaily.$.present": true, "skudDaily.$.lastIn": eventIso },
          },
          { returnDocument: "before", projection: { skudDaily: { $elemMatch: { date: dateKey } } } },
        ).lean();
        const before = prev?.skudDaily?.[0] || {};
        if (!before.firstIn)
          await Student.updateOne(
            { _id: sid, skudDaily: { $elemMatch: { date: dateKey, firstIn: null } } },
            { $set: { "skudDaily.$.firstIn": eventIso } },
          );
        if (!before.inCount) {
          const group = await Group.findById(student.group).select("name shiftStart").lean();
          if (local && local.min > att.shiftMinutes(group?.shiftStart)) {
            // Опоздание — куратор видит через кнопку "Опоздавшие сегодня" в боте
            pushSkudEvent({
              ok: true,
              type: "late",
              studentFio: student.fio,
              studentInn: student.inn || "",
              groupId: String(group?._id || ""),
              groupName: group?.name || "—",
              time: timeStr,
              at: eventIso,
            });
          }
          if (student.inn)
            notifyParents(
              student.inn,
              `✅ *Вход в лицей*\n🧑‍🎓 ${mdv2(student.fio)}\n🕒 *${mdv2(timeStr)}*`,
            ).catch(() => null);
        }
      } else {
        await Student.updateOne(
          { _id: sid, "skudDaily.date": dateKey },
          { $inc: { "skudDaily.$.outCount": 1 }, $set: { "skudDaily.$.lastOut": eventIso } },
        );
        await Student.updateOne(
          { _id: sid, skudDaily: { $elemMatch: { date: dateKey, firstOut: null } } },
          { $set: { "skudDaily.$.firstOut": eventIso } },
        );
        if (local && student.inn)
          notifyParents(
            student.inn,
            `🏃‍♂️ *Выход из лицея*\n🧑‍🎓 ${mdv2(student.fio)}\n🕒 *${mdv2(timeStr)}*`,
          ).catch(() => null);
      }
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

// Восстановление истории с терминалов (history-sync.js): идемпотентно, без уведомлений родителям
require("./history-sync")(app, { Student, AGENT_API_KEY, safeEq, normalizeINN, extractINN, normalizeFIO, extractFIO });

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
// Администратор видит все группы, куратор только свои (проверяется в API)
app.get("/skud-register.html", requireAuth, (req, res) =>
  res.sendFile(path.join(__dirname, "public", "skud-register.html")),
);
// Производственная практика (страницы /practice и /p, API /api/practice и /api/p).
// Подключаем до express.static, чтобы /practice.html открывался только после входа.
require("./practice")(app, { Student, Group, Curator, getSession, getTodayKey });
app.use("/api/admin", requireAuth);
app.use(express.static("public"));
absence.registerAbsenceRoutes(app, { Student, canAccessGroup });
const skud = require("./skud")(app, {
  Student,
  Group,
  requireAuth,
  requireAdmin,
  normalizeINN,
  AGENT_API_KEY,
  getSession,
  canAccessGroup,
  getAccessibleGroupIds,
  notify: (m) => notifyAdmins(m),
});

// Перевод студента между группами (student-transfer.js)
require("./student-transfer")(app, { Student, Group, AbsenceReason: absence.AbsenceReason, requireAdmin, audit: skud.core.audit });

// Ручная правка посещаемости, только админ (attendance-manual.js)
require("./attendance-manual")(app, { Student, requireAdmin, audit: skud.core.audit, att, getTodayKey, getSession });

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
    await skud.enqueueDeleteForGroup(req.params.id, req).catch(() => {});
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
    await skud.enqueueDeleteForGroup(req.params.id, req).catch(() => {});
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
    const { profId, shiftStart, saturday } = req.body;
    let ud = { shiftStart: shiftStart || "09:30" };
    if (saturday !== undefined) ud.saturday = !!saturday;
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

/* === ДОСТУП КУРАТОРОВ (curators.js) === */
require("./curators")(app, {
  Curator,
  Group,
  TgSub,
  hashPassword,
  adminLogin: ADMIN_LOGIN,
  requireAdmin,
  onChange: (id) => {
    forgetCuratorCache(id);
    try {
      forgetBotCurator?.(String(id));
    } catch {}
  },
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
    const st = await Student.findById(req.params.id);
    if (st) await skud.enqueueDeleteForStudent(st, req).catch(() => {});
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
          { upsert: true, returnDocument: "after", setDefaultsOnInsert: true },
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
      const { gid, month } = req.params;
      if (!/^\d{4}-\d{2}$/.test(month)) return res.status(400).send("Неверный месяц");
      const { from, to } = att.monthRange(month);
      const dim = Number(to.slice(8));
      const g = await Group.findById(gid);
      if (!g) return res.status(404).send("Группа не найдена");
      const shiftMin = att.shiftMinutes(g.shiftStart);
      const students = await Student.find({ group: gid })
        .sort({ fio: 1 })
        .select("fio skudDaily attendance")
        .lean();
      const reasons = await absence.reasonsIndex({ group: gid }, from, to);
      const tdk = getTodayKey();
      const excelData = [],
        hdr = ["№", "ФИО Студента"];
      for (let d = 1; d <= dim; d++) hdr.push(String(d));
      hdr.push("Присутствий", "Опозданий", "Пропусков уваж.", "Пропусков без уваж.");
      excelData.push(hdr);
      students.forEach((s, i) => {
        const row = [i + 1, s.fio];
        let sp = 0,
          sa = 0,
          se = 0,
          sl = 0;
        for (let d = 1; d <= dim; d++) {
          const dk = `${month}-${pad2(d)}`;
          const st = att.dayStatus(s, dk, shiftMin);
          if (st.present) {
            sp++;
            if (st.late) sl++;
            row.push(st.late ? "О" : "П");
          } else if (!att.isWeekend(dk, g.saturday) && dk <= tdk) {
            const r = reasons.get(`${s._id}|${dk}`);
            if (r && r.excused && r.confirmed) se++;
            else sa++;
            row.push(r ? r.code + (r.confirmed ? "" : "?") : "Н");
          } else row.push("-");
        }
        row.push(sp, sl, se, sa);
        excelData.push(row);
      });
      excelData.push(
        [],
        ["", "П — был, О — опоздал, Н — без причины, Б — болезнь, С — семейные, Л — по направлению лицея, Д — другая причина, ? — не подтверждено"],
      );
      const wb = xlsx.utils.book_new(),
        ws = xlsx.utils.aoa_to_sheet(excelData);
      const cw = [{ wch: 4 }, { wch: 35 }];
      for (let d = 1; d <= dim; d++) cw.push({ wch: 4 });
      cw.push({ wch: 12 }, { wch: 11 }, { wch: 15 }, { wch: 18 });
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
          { upsert: true, returnDocument: "after" },
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
        "января", "февраля", "марта", "апреля", "мая", "июня",
        "июля", "августа", "сентября", "октября", "ноября", "декабря",
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
        "ЯНВАРЬ", "ФЕВРАЛЬ", "МАРТ", "АПРЕЛЬ", "МАЙ", "ИЮНЬ",
        "ИЮЛЬ", "АВГУСТ", "СЕНТЯБРЬ", "ОКТЯБРЬ", "НОЯБРЬ", "ДЕКАБРЬ",
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
    if (!process.env.ADMIN_TG_SECRET || !process.env.BOT_USERNAME)
      return res.status(500).json({ error: "Не заданы ADMIN_TG_SECRET или BOT_USERNAME в настройках сервера" });
    const hmac = require("crypto")
      .createHmac("sha256", process.env.ADMIN_TG_SECRET)
      .update(curId)
      .digest("hex")
      .slice(0, 8);
    const token = `${curId}_${hmac}`;
    const botUser = process.env.BOT_USERNAME;
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
    "ЯНВАРЬ", "ФЕВРАЛЬ", "МАРТ", "АПРЕЛЬ", "МАЙ", "ИЮНЬ",
    "ИЮЛЬ", "АВГУСТ", "СЕНТЯБРЬ", "ОКТЯБРЬ", "НОЯБРЬ", "ДЕКАБРЬ",
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
        "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
        "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь",
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

// Отчёты и «итог недели» не зависят от бота: бот передаётся при отправке
const reports = createReports({ Group, Student, Curator, TgSub, getTodayKey });
const weeklyCards = createWeeklyCards({ Group, Student, TgSub, SummaryRecord, ensureGroupSummarySubjects, getTodayKey });

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
    weeklyCards,
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
forgetBotCurator = _botExports.forgetCurator;

/* ================================================================
   TELEGRAM MINI APP API — /api/webapp/* (webapp.js)
   Причины отсутствия — absence.js, отчёты — reports.js
================================================================ */
require("./webapp")(app, {
  Group,
  Student,
  TgSub,
  Curator,
  SummaryRecord,
  ensureGroupSummarySubjects,
  formatMonthLabel,
  notifyParents: (...a) => notifyParents(...a),
  getTodayKey,
  getBot: () => bot,
  mdv2,
  weeklyCards,
});

reports.registerReportRoutes(app, {
  requireAdmin,
  canAccessGroup,
  contentDisposition,
  getBot: () => bot,
});
reports.scheduleReports(bot);

// «Не пришёл»: уведомления родителям и кураторам (absent-alerts.js)
const absentAlerts = require("./absent-alerts").createAbsentAlerts({
  Group,
  Student,
  Curator,
  TgSub,
  getTodayKey,
  getBishkekMinutes: () => {
    const d = getBishkekNow();
    return d.getHours() * 60 + d.getMinutes();
  },
  // студенты на производственной практике по графику в этот день — их в лицее не ждём
  practiceStudentIds: async (day) => {
    const P = mongoose.models.Practice;
    if (!P) return [];
    const list = await P.find({ "days.date": day }).select("students.student").lean();
    return list.flatMap((p) => (p.students || []).map((s) => s.student));
  },
});
absentAlerts.schedule(bot);
app.post("/api/admin/absent-alerts/run", requireAdmin, async (req, res) => {
  try {
    // по умолчанию — предпросмотр (кому ушло бы уведомление); send=1 — отправить сейчас,
    // но не повторно: группы, по которым сегодня уже отправляли, пропускаются
    const send = req.query.send === "1";
    res.json(await absentAlerts.run(bot, send ? { ignoreTime: true } : { dry: true, force: true }));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Общий журнал по лицею (листы по курсам)
require("./journal-all").create({ Group, Student, absence, att, getTodayKey }).registerRoutes(app, { requireAdmin, contentDisposition });

// Учёт сотрудников: преподаватели по расписанию Kundoluk, мастера — присутствие (staff-attendance.js)
const staffAttendance = require("./staff-attendance").createStaffAttendance({
  Group,
  Student,
  TgSub,
  getTodayKey,
  getBishkekMinutes: () => {
    const d = getBishkekNow();
    return d.getHours() * 60 + d.getMinutes();
  },
});
staffAttendance.registerRoutes(app, { requireAdmin, requireAuth, safeEq, contentDisposition });
staffAttendance.schedule(bot);

// Еженедельный бэкап базы в Telegram администраторам (backup.js)
const backup = require("./backup").createBackup({ getDb: () => mongoose.connection.db, TgSub, getTodayKey });
backup.schedule(bot);
backup.registerRoutes(app, { requireAdmin, getBot: () => bot });

// Дашборд посещаемости и «зона риска» (dashboard.js)
require("./dashboard").registerDashboard(app, { Group, reports, getTodayKey, getAccessibleGroupIds, requireAuth });
weeklyCards.schedule(bot);

// «Итог недели»: предпросмотр карточки и ручная рассылка
app.get("/api/admin/weekly-card/:studentId.png", async (req, res) => {
  try {
    const stu = await Student.findById(req.params.studentId).select("group").lean();
    if (!stu) return res.status(404).json({ error: "Студент не найден" });
    if (!canAccessGroup(req, stu.group)) return res.status(403).json({ error: "Нет доступа к группе" });
    const c = await weeklyCards.cardFor({ _id: stu._id }, att.isValidKey(req.query.date) ? req.query.date : undefined);
    res.setHeader("Content-Type", "image/png");
    res.send(c.png);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});
app.post("/api/admin/weekly-cards/send-now", requireAdmin, async (req, res) => {
  try {
    res.json(await weeklyCards.sendAll(bot));
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

const PORT = Number(process.env.PORT) || 3000;
app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));
