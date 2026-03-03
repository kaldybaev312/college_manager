const express = require("express");
const mongoose = require("mongoose");
const path = require("path");
const fs = require("fs");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const ImageModule = require("docxtemplater-image-module-free");
const QRCode = require("qrcode");
const multer = require("multer");
const xlsx = require("xlsx");
const DigestClient = require("digest-fetch").default;
const TelegramBot = require("node-telegram-bot-api");
const cron = require("node-cron");

const app = express();
const upload = multer({ dest: "uploads/" });
if (!fs.existsSync("uploads")) fs.mkdirSync("uploads");

app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));

/* ===========================
   ЗАЩИТА АДМИНКИ (ПАРОЛЬ)
=========================== */
const ADMIN_LOGIN = process.env.ADMIN_LOGIN || "admin";
const ADMIN_PASS = process.env.ADMIN_PASS || "pl3-2026";

// --- 1. ПУБЛИЧНЫЕ API (ДЛЯ АГЕНТА И ЖУРНАЛА) ---
// Эти маршруты должны быть ПЕРЕД requireAuth

// СИНХРОНИЗАЦИЯ АГЕНТА (защищена только API ключом)
app.post("/api/agent/sync", async (req, res) => {
  const { apiKey, data } = req.body;
  if (apiKey !== AGENT_API_KEY) {
    console.log("⚠️ Попытка доступа агента с неверным ключом!");
    return res.status(403).send("Forbidden: Invalid API Key");
  }
  // ... ваш код обработки data ...
  res.json({ success: true });
});

// ПУБЛИЧНЫЕ ДАННЫЕ ДЛЯ РОДИТЕЛЕЙ (journal.html)
app.get("/api/admin/groups-list", async (req, res, next) => next());
app.get("/api/admin/global-stats", async (req, res, next) => next());
app.get("/api/admin/attendance-matrix/:gid/:month", async (req, res, next) =>
  next(),
);

// Вспомогательная функция для чтения Cookie (чтобы запоминать вход)
function getCookie(req, name) {
  const cookieHeader = req.headers.cookie;
  if (!cookieHeader) return null;
  const match = cookieHeader.match(new RegExp("(^| )" + name + "=([^;]+)"));
  return match ? match[2] : null;
}

// Проверка: вошел ли пользователь?
function requireAuth(req, res, next) {
  if (getCookie(req, "admin_auth") === "true") return next();
  res.redirect("/login");
}

// Разрешаем серверу понимать данные из форм
app.use(express.urlencoded({ extended: true }));

// 1. ОТДАЕМ КРАСИВУЮ СТРАНИЦУ ЛОГИНА
app.get("/login", (req, res) => {
  // Если ввели неверный пароль, покажем красную ошибку
  const errorMsg = req.query.error
    ? '<div style="color: #dc2626; font-size: 13px; margin-bottom: 15px; text-align: center; background: #fee2e2; padding: 10px; border-radius: 8px; font-weight: 500;">Неверный логин или пароль!</div>'
    : "";

  res.send(`
    <!DOCTYPE html>
    <html lang="ru">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Вход | ПЛ №3</title>
      <style>
        body { background: #f1f5f9; font-family: "Inter", "Segoe UI", sans-serif; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; }
        .login-card { background: white; padding: 40px 30px; border-radius: 20px; box-shadow: 0 10px 25px rgba(0,0,0,0.05); width: 100%; max-width: 380px; border: 1px solid #e2e8f0; }
        .logo { text-align: center; margin-bottom: 30px; }
        .logo h2 { margin: 0; font-size: 24px; color: #0f172a; font-weight: 800; }
        .logo p { margin: 5px 0 0; color: #64748b; font-size: 14px; }
        .form-group { margin-bottom: 20px; }
        label { display: block; margin-bottom: 8px; font-size: 13px; font-weight: 600; color: #475569; }
        input { width: 100%; padding: 12px 15px; border: 2px solid #e2e8f0; border-radius: 10px; font-size: 15px; outline: none; transition: border-color 0.2s; box-sizing: border-box; }
        input:focus { border-color: #4f46e5; }
        button { width: 100%; padding: 14px; background: #4f46e5; color: white; border: none; border-radius: 10px; font-size: 16px; font-weight: bold; cursor: pointer; transition: background 0.2s; }
        button:hover { background: #4338ca; }
      </style>
    </head>
    <body>
      <div class="login-card">
        <div class="logo">
          <h2>ПРОФЛИЦЕЙ №3</h2>
          <p>Система Управления</p>
        </div>
        ${errorMsg}
        <form action="/login" method="POST">
          <div class="form-group">
            <label>Логин администратора</label>
            <input type="text" name="username" placeholder="Введите логин" required autocomplete="off">
          </div>
          <div class="form-group">
            <label>Пароль</label>
            <input type="password" name="password" placeholder="Введите пароль" required>
          </div>
          <button type="submit">Войти в панель</button>
        </form>
      </div>
    </body>
    </html>
  `);
});

// 2. ОБРАБОТКА КНОПКИ "ВОЙТИ"
app.post("/login", (req, res) => {
  const { username, password } = req.body;
  if (username === ADMIN_LOGIN && password === ADMIN_PASS) {
    res.cookie("admin_auth", "true", {
      maxAge: 30 * 24 * 60 * 60 * 1000,
      httpOnly: true,
    });
    res.redirect("/");
  } else {
    res.redirect("/login?error=1");
  }
});

// 3. ВЫХОД ИЗ АДМИНКИ
app.get("/logout", (req, res) => {
  res.clearCookie("admin_auth");
  res.redirect("/login");
});

// 4. ЗАЩИЩАЕМ СТРАНИЦЫ
// Разрешаем journal.html открываться без пароля
app.get("/journal.html", (req, res, next) => next());

// Все остальные страницы (index.html и т.д.) требуют пароль
app.get(
  ["/", "/index.html", "/skud-events.html", "/admin/events"],
  requireAuth,
);

// Все остальные API методы (удаление, редактирование) требуют пароль
app.use("/api/admin", requireAuth);

app.use(express.static("public"));

/* ===========================
   MONGODB
=========================== */
const MONGO_URI =
  process.env.MONGO_URI ||
  "mongodb+srv://ilimkaldybaev5_db_user:liceyStudents@riestr.uki8ep8.mongodb.net/PL3_Database?retryWrites=true&w=majority&appName=riestr";

mongoose
  .connect(MONGO_URI)
  .then(() => console.log("✅ MongoDB connected"))
  .catch((e) => console.log("❌ Mongo error:", e.message));

/* ===========================
   CONFIG
=========================== */
const AGENT_API_KEY = process.env.AGENT_API_KEY || "pl3_secret_key";
const HIK_USER = process.env.HIK_USER || "admin";
const HIK_PASS = process.env.HIK_PASS || "Qwerty#12";

/* ===========================
   PROFESSIONS
=========================== */
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

/* ===========================
   MODELS
=========================== */
const Group = mongoose.model(
  "Group",
  new mongoose.Schema({
    name: { type: String, unique: true },
    profRu: String,
    profKg: String,
    duration: String,
    shiftStart: { type: String, default: "09:30" }, // Настройка времени смены для группы
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
  new mongoose.Schema({
    chatId: { type: String, unique: true },
    name: String,
    linkedInns: [{ type: String }], // <--- Здесь бот будет хранить ИНН детей
  }),
);

// Функция для рассылки всем подписчикам (завучу)
async function notifyTg(message) {
  try {
    const subs = await TgSub.find();
    for (const sub of subs) {
      bot
        .sendMessage(sub.chatId, message, { parse_mode: "Markdown" })
        .catch(() => null);
    }
  } catch (e) {
    console.error("TG Notify Error:", e);
  }
}

/* ===========================
   SKUD EVENTS MEMORY LOG
=========================== */
const SKUD_EVENTS = [];
const SKUD_LIMIT = 500;

function pushSkudEvent(ev) {
  SKUD_EVENTS.unshift(ev);
  if (SKUD_EVENTS.length > SKUD_LIMIT) SKUD_EVENTS.pop();
}

app.get("/api/admin/skud-events", (req, res) => {
  res.json({ ok: true, items: SKUD_EVENTS });
});

app.get("/admin/events", (req, res) => {
  res.sendFile(path.join(__dirname, "public", "skud-events.html"));
});

/* ===========================
   HELPERS
=========================== */
function pad2(n) {
  return String(n).padStart(2, "0");
}

function isoToDateKey(iso) {
  const d = new Date(iso);
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function digitsOnly(s) {
  return String(s || "").replace(/\D/g, "");
}

function normalizeINN(s) {
  const d = digitsOnly(s);
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
  return fio.trim();
}

/* ===========================
   AGENT SYNC
=========================== */
app.post("/api/agent/sync", async (req, res) => {
  const { apiKey, data } = req.body;

  if (apiKey !== AGENT_API_KEY) return res.status(403).send("Forbidden");
  if (!Array.isArray(data))
    return res.status(400).json({ error: "data must be array" });

  let okCount = 0;
  let failCount = 0;

  for (const entry of data) {
    try {
      const employeeNo = entry.id?.toString().trim();
      const deviceIp = entry.deviceIp;
      const direction = entry.direction;

      const rawHikTime = entry.time || "";
      const eventIso = rawHikTime;

      // 🔥 ФИКС ТАЙМЗОНЫ: Вычисляем правильную дату (+6 часов) для записи в базу
      let dateKey;
      const dt = new Date(rawHikTime);
      if (!isNaN(dt.getTime())) {
        dt.setUTCHours(dt.getUTCHours() + 6); // +6 часов (Бишкек)
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
        failCount++;
        continue;
      }

      const inn = normalizeINN(extractINN(rawName));
      const fioNorm = normalizeFIO(extractFIO(rawName));

      let student = null;

      if (inn) {
        student = await Student.findOne({ inn });
      }

      if (!student && fioNorm) {
        const firstNamePart = fioNorm.split(" ")[0];
        const candidates = await Student.find({
          fio: new RegExp("^" + firstNamePart, "i"),
        }).limit(50);

        student =
          candidates.find((c) => normalizeFIO(c.fio) === fioNorm) || null;
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
        failCount++;
        continue;
      }

      const exist = student.attendance.find((a) => a.date === dateKey);
      if (exist) exist.present = true;
      else student.attendance.push({ date: dateKey, present: true });

      let rec = student.skudDaily.find((r) => r.date === dateKey);
      if (!rec) {
        rec = { date: dateKey, inCount: 0, outCount: 0, present: false };
        student.skudDaily.push(rec);
        rec = student.skudDaily[student.skudDaily.length - 1];
      }

      let isFirstEntryToday = false;
      let timeStr = "—";
      let localH = 0,
        localM = 0;
      if (direction === "in") {
        isFirstEntryToday = rec.inCount === 0;
        rec.inCount++;
        rec.present = true;
        rec.firstIn = rec.firstIn || eventIso;
        rec.lastIn = eventIso;

        const dt = new Date(eventIso);
        if (!isNaN(dt.getTime())) {
          const totalMin = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
          localH = Math.floor(totalMin / 60) % 24;
          localM = totalMin % 60;
          timeStr = `${pad2(localH)}:${pad2(localM)}`;
        }

        // 2. Если это первый вход за день — запускаем проверки
        if (isFirstEntryToday) {
          await student.populate("group");
          const shiftTime = student.group?.shiftStart || "09:30";
          const [shiftH, shiftM] = shiftTime.split(":").map(Number);

          let lateInfo = "";
          // Проверка на опоздание
          if (localH * 60 + localM > shiftH * 60 + shiftM) {
            lateInfo = ` (Опоздание! Смена: ${shiftTime})`;
            notifyTg(
              `⚠️ *Опоздание!*\n🧑‍🎓 ${student.fio}\n📂 Группа: ${student.group?.name || "—"}\n⏰ Прибыл в: *${timeStr}*`,
            );
          }

          // Уведомление привязанным родителям
          if (student.inn) {
            const parents = await TgSub.find({ linkedInns: student.inn });
            parents.forEach((p) => {
              bot
                .sendMessage(
                  p.chatId,
                  `✅ *Вход в лицей*\n🧑‍🎓 ${student.fio}\n🕒 Время: *${timeStr}*${lateInfo}`,
                  { parse_mode: "Markdown" },
                )
                .catch(() => null);
            });
          }
        }
      } else {
        // --- ЛОГИКА ВЫХОДА ---
        rec.outCount++;
        rec.firstOut = rec.firstOut || eventIso;
        rec.lastOut = eventIso;

        // Здесь можно добавить аналогичное уведомление о выходе для родителей
        const dt = new Date(eventIso);
        if (!isNaN(dt.getTime())) {
          const totalMin = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
          const timeStr = `${pad2(Math.floor(totalMin / 60) % 24)}:${pad2(totalMin % 60)}`;

          // УВЕДОМЛЕНИЕ ДЛЯ РОДИТЕЛЕЙ О ВЫХОДЕ
          if (student.inn) {
            TgSub.find({ linkedInns: student.inn }).then((parents) => {
              parents.forEach((p) => {
                bot
                  .sendMessage(
                    p.chatId,
                    `🏃‍♂️ *Выход из лицея*\n🧑‍🎓 ${student.fio}\n🕒 Время: *${timeStr}*`,
                    { parse_mode: "Markdown" },
                  )
                  .catch(() => null);
              });
            });
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
      okCount++;
    } catch (e) {
      failCount++;
      pushSkudEvent({ ok: false, reason: e.message });
    }
  }

  res.json({ success: true, okCount, failCount });
});

/* ===========================
   ADMIN — GROUPS
=========================== */
app.get("/api/admin/groups-list", async (req, res) => {
  try {
    const groups = await Group.find().sort({ name: 1 });
    res.json(groups);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.delete("/api/admin/group/:id", async (req, res) => {
  try {
    await Student.deleteMany({ group: req.params.id });
    await Group.findByIdAndDelete(req.params.id);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.delete("/api/admin/group/:id/clear", async (req, res) => {
  try {
    await Student.deleteMany({ group: req.params.id });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Новый роут: Сохранение настроек группы (Профессия и Время смены)
app.put("/api/admin/group/:id/settings", async (req, res) => {
  try {
    const { profId, shiftStart } = req.body;
    let updateData = { shiftStart: shiftStart || "09:30" };

    if (profId) {
      const prof = PROFESSIONS.find((p) => p.id == profId);
      if (prof) {
        updateData.profRu = prof.ru;
        updateData.profKg = prof.kg;
        updateData.duration = prof.duration;
      }
    }

    await Group.findByIdAndUpdate(req.params.id, updateData);
    res.json({ ok: true, updateData });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* ===========================
   ADMIN — STUDENTS
=========================== */
app.get("/api/admin/group-students/:id", async (req, res) => {
  try {
    const students = await Student.find({ group: req.params.id }).sort({
      fio: 1,
    });
    res.json(students);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/admin/search", async (req, res) => {
  try {
    const q = (req.query.q || "").trim();
    const page = Math.max(1, parseInt(req.query.page) || 1);
    const limit = Math.min(100, parseInt(req.query.limit) || 60);
    const skip = (page - 1) * limit;

    const filter = q.length >= 2 ? { fio: new RegExp(q, "i") } : {};

    // Считаем total и тянем страницу параллельно
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

app.delete("/api/admin/student/:id", async (req, res) => {
  try {
    await Student.findByIdAndDelete(req.params.id);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.put("/api/admin/student/:id", async (req, res) => {
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

/* ===========================
   ADMIN — ЖУРНАЛ ПОСЕЩАЕМОСТИ
=========================== */
app.get("/api/admin/attendance-matrix/:gid/:month", async (req, res) => {
  try {
    const { gid, month } = req.params;
    const [year, mon] = month.split("-").map(Number);
    const daysInMonth = new Date(year, mon, 0).getDate();

    const group = await Group.findById(gid);
    const shiftTime = group?.shiftStart || "09:30";
    const [shiftH, shiftM] = shiftTime.split(":").map(Number);
    const shiftTotalMinutes = shiftH * 60 + shiftM;

    const students = await Student.find({ group: gid }).sort({ fio: 1 });

    // 🔥 БРОНЕБОЙНЫЙ ФИКС ВРЕМЕНИ (Точное время Бишкека +6)
    const nowDt = new Date();
    const utcMs = nowDt.getTime() + nowDt.getTimezoneOffset() * 60000;
    const bishkekDt = new Date(utcMs + 6 * 3600000);

    const bYear = bishkekDt.getFullYear();
    const bMonth = pad2(bishkekDt.getMonth() + 1);
    const bDay = pad2(bishkekDt.getDate());
    const todayDateKey = `${bYear}-${bMonth}-${bDay}`; // Например: "2026-03-03"
    const todayIsWeekend = bishkekDt.getDay() === 0 || bishkekDt.getDay() === 6;

    let presentTodayCount = 0;
    let totalMonthLates = 0;

    const matrix = students.map((s) => {
      const days = [];
      let isPresentToday = false;
      let studentLates = 0;
      let studentPresences = 0;
      let studentAbsences = 0;

      for (let d = 1; d <= daysInMonth; d++) {
        const dateKey = `${year}-${pad2(mon)}-${pad2(d)}`;

        // 🔥 Надежное определение выходного (без влияния часового пояса сервера)
        // Используем строгий UTC, чтобы день недели никогда не съезжал
        const dateObj = new Date(Date.UTC(year, mon - 1, d));
        const dayOfWeek = dateObj.getUTCDay();
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

        const skud = s.skudDaily?.find((r) => r.date === dateKey);
        const att = s.attendance?.find((a) => a.date === dateKey);

        let present = false;
        let late = false;
        let timeStr = "";

        if (skud) {
          present = skud.present;
          if (skud.firstIn) {
            // Возвращаем время терминала в Бишкекское (+360 минут)
            const dt = new Date(skud.firstIn);
            if (!isNaN(dt.getTime())) {
              const totalMin = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
              const localH = Math.floor(totalMin / 60) % 24;
              const localM = totalMin % 60;
              timeStr = `${pad2(localH)}:${pad2(localM)}`;
              late = localH * 60 + localM > shiftTotalMinutes;
            }
          }
        } else if (att) {
          present = att.present;
        }

        // --- ЛОГИКА СТАТИСТИКИ ---
        if (present) {
          studentPresences++;
        } else if (!isWeekend && dateKey <= todayDateKey) {
          // Отсутствие засчитывается ТОЛЬКО если это НЕ СУББОТА/ВОСКРЕСЕНЬЕ
          // и этот день уже наступил (или прошел)
          studentAbsences++;
        }

        if (late) {
          studentLates++;
          totalMonthLates++;
        }

        if (dateKey === todayDateKey && present) {
          isPresentToday = true;
        }

        days.push({ present, late, time: timeStr, isWeekend });
      }

      if (isPresentToday) presentTodayCount++;

      return {
        fio: s.fio,
        days,
        stats: {
          presences: studentPresences,
          lates: studentLates,
          absences: studentAbsences,
        },
      };
    });

    res.json({
      daysInMonth,
      matrix,
      shiftStart: shiftTime,
      stats: {
        totalStudents: students.length,
        presentToday: presentTodayCount,
        // Считаем группу отсутствующей сегодня, только если сегодня будний день
        absentToday: todayIsWeekend ? 0 : students.length - presentTodayCount,
        totalLates: totalMonthLates,
      },
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* ===========================
   ADMIN — ЭКСПОРТ ЖУРНАЛА В EXCEL
=========================== */
app.get("/api/admin/export-journal/:gid/:month", async (req, res) => {
  try {
    const { gid, month } = req.params;
    const [year, mon] = month.split("-").map(Number);
    const daysInMonth = new Date(year, mon, 0).getDate();

    const group = await Group.findById(gid);
    if (!group) return res.status(404).send("Группа не найдена");

    const shiftTime = group.shiftStart || "09:30";
    const [shiftH, shiftM] = shiftTime.split(":").map(Number);
    const shiftTotalMinutes = shiftH * 60 + shiftM;

    const students = await Student.find({ group: gid }).sort({ fio: 1 });

    // Точное время Бишкека +6
    const nowDt = new Date();
    const utcMs = nowDt.getTime() + nowDt.getTimezoneOffset() * 60000;
    const bishkekDt = new Date(utcMs + 6 * 3600000);
    const bYear = bishkekDt.getFullYear();
    const bMonth = pad2(bishkekDt.getMonth() + 1);
    const bDay = pad2(bishkekDt.getDate());
    const todayDateKey = `${bYear}-${bMonth}-${bDay}`;

    // Подготавливаем данные для Excel (массив массивов)
    const excelData = [];

    // Формируем заголовок таблицы
    const headerRow = ["№", "ФИО Студента"];
    for (let d = 1; d <= daysInMonth; d++) headerRow.push(String(d));
    headerRow.push("Присутствий", "Отсутствий", "Опозданий");
    excelData.push(headerRow);

    // Заполняем строки студентов
    students.forEach((s, idx) => {
      const row = [idx + 1, s.fio];
      let studentPresences = 0;
      let studentAbsences = 0;
      let studentLates = 0;

      for (let d = 1; d <= daysInMonth; d++) {
        const dateKey = `${year}-${pad2(mon)}-${pad2(d)}`;
        const dateObj = new Date(Date.UTC(year, mon - 1, d));
        const dayOfWeek = dateObj.getUTCDay();
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;

        const skud = s.skudDaily?.find((r) => r.date === dateKey);
        const att = s.attendance?.find((a) => a.date === dateKey);

        let present = false;
        let late = false;

        if (skud) {
          present = skud.present;
          if (skud.firstIn) {
            const dt = new Date(skud.firstIn);
            if (!isNaN(dt.getTime())) {
              const totalMin = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
              const localH = Math.floor(totalMin / 60) % 24;
              const localM = totalMin % 60;
              if (localH * 60 + localM > shiftTotalMinutes) late = true;
            }
          }
        } else if (att) {
          present = att.present;
        }

        // Подсчет и вывод символа в ячейку
        if (present) {
          studentPresences++;
          if (late) {
            studentLates++;
            row.push("О"); // Опоздал
          } else {
            row.push("П"); // Присутствовал
          }
        } else if (!isWeekend && dateKey <= todayDateKey) {
          studentAbsences++;
          row.push("Н"); // Не был
        } else {
          // Выходной или будущий день
          row.push("-");
        }
      }

      // Добавляем итоговую статистику студента в конец строки
      row.push(studentPresences, studentAbsences, studentLates);
      excelData.push(row);
    });

    // Создаем книгу Excel
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet(excelData);

    // Настраиваем ширину колонок
    const colWidths = [{ wch: 4 }, { wch: 35 }]; // № и ФИО
    for (let d = 1; d <= daysInMonth; d++) colWidths.push({ wch: 4 }); // Дни
    colWidths.push({ wch: 12 }, { wch: 12 }, { wch: 12 }); // Итоги
    ws["!cols"] = colWidths;

    xlsx.utils.book_append_sheet(wb, ws, `Журнал`);
    const buf = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });

    // Отправляем файл пользователю
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="Journal_${group.name}_${month}.xlsx"`,
    );
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* ===========================
   ADMIN — ИМПОРТ EXCEL
=========================== */
app.get("/api/admin/professions", (req, res) => {
  res.json(PROFESSIONS);
});

app.post("/api/admin/import", upload.single("file"), async (req, res) => {
  try {
    const { groupName, profId } = req.body;
    const prof = PROFESSIONS.find((p) => p.id == profId);
    if (!prof) return res.status(400).json({ error: "Профессия не найдена" });

    let group = await Group.findOne({ name: groupName });
    if (!group) {
      group = await Group.create({
        name: groupName,
        profRu: prof.ru,
        profKg: prof.kg,
        duration: prof.duration,
        shiftStart: "09:30",
      });
    }

    const wb = xlsx.readFile(req.file.path);
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1 });

    let created = 0;
    for (const row of rows) {
      const fio = String(row[0] || "").trim();
      const birthDate = String(row[1] || "").trim();
      const inn = String(row[2] || "").trim();
      const lyceumId = String(row[3] || "").trim();

      if (!fio || fio.length < 3) continue;

      await Student.findOneAndUpdate(
        { inn: inn || null, fio },
        { fio, birthDate, inn, lyceumId, group: group._id },
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
});

/* ===========================
   ADMIN — СПРАВКИ & ПИТАНИЕ 
=========================== */
app.get("/api/admin/print/:id/:type", async (req, res) => {
  try {
    const student = await Student.findById(req.params.id).populate("group");
    if (!student) return res.status(404).send("Студент не найден");

    const type = req.params.type;
    const tplMap = {
      common: "template_common.docx",
      army: "template_army.docx",
      social: "template_social.docx",
    };

    const tplFile = tplMap[type] || tplMap.common;
    const tplPath = path.join(__dirname, "templates", tplFile);

    if (!fs.existsSync(tplPath)) {
      return res.json({
        message: `Шаблон ${tplFile} не найден. Положите его в папку templates/`,
      });
    }

    const content = fs.readFileSync(tplPath, "binary");
    const zip = new PizZip(content);

    const qrDataUrl = await QRCode.toDataURL(
      `ФИО: ${student.fio}\nИНН: ${student.inn || "—"}\nГруппа: ${student.group?.name || "—"}`,
    );
    const qrBuffer = Buffer.from(qrDataUrl.split(",")[1], "base64");

    const imageModule = new ImageModule({
      centered: false,
      getImage: (tagValue) =>
        tagValue === "qr" ? qrBuffer : fs.readFileSync(tagValue),
      getSize: () => [80, 80],
    });

    const doc = new Docxtemplater(zip, {
      modules: [imageModule],
      paragraphLoop: true,
      linebreaks: true,
    });

    const nowD = new Date();
    const monthsRu = [
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
    const dateStr =
      nowD.getDate() +
      " " +
      monthsRu[nowD.getMonth()] +
      " " +
      nowD.getFullYear();

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
      date: dateStr,
      year: nowD.getFullYear(),
      qr: "qr",
    });

    const buf = doc.getZip().generate({ type: "nodebuffer" });

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="spravka_${student.lyceumId || student._id}.docx"`,
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

app.get("/api/admin/report-food/:gid/:month", async (req, res) => {
  try {
    const { gid, month } = req.params;
    const [year, mon] = month.split("-").map(Number);
    const daysInMonth = new Date(year, mon, 0).getDate();

    const group = await Group.findById(gid);
    const students = await Student.find({ group: gid }).sort({ fio: 1 });

    if (!group) return res.status(404).json({ error: "Группа не найдена" });
    if (!students.length)
      return res.status(404).json({ error: "Студентов нет" });

    const rows = students.map((s, idx) => {
      let presentDays = 0;
      const dayCells = [];

      for (let d = 1; d <= daysInMonth; d++) {
        const dateKey = `${year}-${pad2(mon)}-${pad2(d)}`;
        const skud = s.skudDaily?.find((r) => r.date === dateKey);
        const att = s.attendance?.find((a) => a.date === dateKey);
        const present = skud ? skud.present : att ? att.present : false;
        if (present) presentDays++;
        dayCells.push(present ? "+" : "");
      }

      return {
        num: idx + 1,
        fio: s.fio,
        lyceumId: s.lyceumId || "—",
        days: dayCells,
        totalDays: presentDays,
        totalAmount: presentDays * 60,
      };
    });

    const tplPath = path.join(__dirname, "templates", "template_food.docx");

    if (!fs.existsSync(tplPath)) {
      return res.json({
        message:
          "Шаблон template_food.docx не найден. Положите его в папку templates/",
      });
    }

    const content = fs.readFileSync(tplPath, "binary");
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
    });

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
    const PRICE_PER_DAY = 60;
    const grandTotal = rows.reduce((s, r) => s + r.totalAmount, 0);

    doc.render({
      month_name: MONTHS_RU[mon - 1],
      year: String(year),
      group: group.name,
      groupName: group.name,
      st: rows.map((r) => ({
        fio: r.fio,
        lyceum_id: r.lyceumId || "—",
        days: String(r.totalDays),
        price: String(PRICE_PER_DAY),
        sum: String(r.totalAmount),
        num: String(r.num),
        lyceumId: r.lyceumId || "—",
        totalDays: String(r.totalDays),
        totalAmount: String(r.totalAmount),
      })),
      total: String(grandTotal),
      grandTotal: String(grandTotal),
      profRu: group.profRu || "—",
      profKg: group.profKg || "—",
      prof_ru: group.profRu || "—",
      prof_kg: group.profKg || "—",
      duration: group.duration || "—",
      month: `${pad2(mon)}.${year}`,
      daysInMonth: String(daysInMonth),
    });

    const buf = doc.getZip().generate({ type: "nodebuffer" });

    res.setHeader(
      "Content-Disposition",
      `attachment; filename="food_${group.name}_${month}.docx"`,
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

/* ===========================
   TELEGRAM BOT (ПЛ №3) v3.0
   Reply Keyboard + пошаговое меню
   Замените ВЕСЬ старый блок бота на этот
=========================== */
const TG_TOKEN =
  process.env.TG_TOKEN || "8552643412:AAGfcarX8UI02vHNDGU4HIBHdKiCe10SNuQ";
const bot = new TelegramBot(TG_TOKEN, { polling: true });

// =========================
// СОСТОЯНИЯ ПОЛЬЗОВАТЕЛЕЙ
// Хранит в памяти шаг диалога для каждого chatId
// =========================
const userState = new Map();
// Структура: { step: "idle"|"await_group_report"|"await_month_report"|"await_group_view"|"await_link_inn", groupName?: string }

function getState(chatId) {
  return userState.get(chatId) || { step: "idle" };
}
function setState(chatId, state) {
  userState.set(chatId, state);
}
function resetState(chatId) {
  userState.set(chatId, { step: "idle" });
}

// =========================
// КЛАВИАТУРЫ
// =========================

// Главное меню — всегда внизу
const MAIN_KEYBOARD = {
  reply_markup: {
    keyboard: [
      [{ text: "📊 Сводка сегодня" }, { text: "📂 Состояние группы" }],
      [{ text: "📥 Скачать Excel" }, { text: "🔍 Найти студента" }],
      [{ text: "🔔 Привязать ребёнка" }],
    ],
    resize_keyboard: true,
    persistent: true,
  },
};

// Клавиатура выбора группы (строится динамически из БД)
async function buildGroupKeyboard(extraButton = null) {
  const groups = await Group.find().sort({ name: 1 });
  const rows = [];
  // По 2 группы в ряд
  for (let i = 0; i < groups.length; i += 2) {
    const row = [{ text: groups[i].name }];
    if (groups[i + 1]) row.push({ text: groups[i + 1].name });
    rows.push(row);
  }
  // Кнопка отмены всегда последняя
  rows.push([{ text: "❌ Отмена" }]);
  return {
    reply_markup: {
      keyboard: rows,
      resize_keyboard: true,
      one_time_keyboard: false,
    },
  };
}

// Клавиатура выбора месяца
function buildMonthKeyboard() {
  const months = [];
  const now = getBishkekNow();
  // Текущий месяц + 5 предыдущих
  for (let i = 0; i < 6; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    const val = `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
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
    const label = `${MONTHS_RU[d.getMonth()]} ${d.getFullYear()}`;
    months.push({ text: label, value: val });
  }
  const rows = [];
  for (let i = 0; i < months.length; i += 2) {
    const row = [{ text: months[i].text }];
    if (months[i + 1]) row.push({ text: months[i + 1].text });
    rows.push(row);
  }
  rows.push([{ text: "❌ Отмена" }]);
  return {
    months,
    keyboard: { reply_markup: { keyboard: rows, resize_keyboard: true } },
  };
}

// Убрать клавиатуру
const REMOVE_KEYBOARD = { reply_markup: { remove_keyboard: true } };

// =========================
// HELPERS — время Бишкека
// =========================
function getBishkekNow() {
  const nowDt = new Date();
  const utcMs = nowDt.getTime() + nowDt.getTimezoneOffset() * 60000;
  return new Date(utcMs + 6 * 3600000);
}
function getTodayKey() {
  const d = getBishkekNow();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}
function getCurrentMonth() {
  const d = getBishkekNow();
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`;
}
// Парсим "Март 2026" → "2026-03"
function parseMonthLabel(label) {
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
  const parts = label.trim().split(" ");
  if (parts.length < 2) return null;
  const mIdx = MONTHS_RU.findIndex((m) => m === parts[0]);
  if (mIdx === -1) return null;
  return `${parts[1]}-${pad2(mIdx + 1)}`;
}

// =========================
// /start — подписка + главное меню
// =========================
bot.onText(/\/start/, async (msg) => {
  const chatId = String(msg.chat.id);
  const name = msg.from.first_name || "Коллега";
  await TgSub.findOneAndUpdate({ chatId }, { chatId, name }, { upsert: true });
  resetState(chatId);

  bot.sendMessage(
    chatId,
    `👋 Добро пожаловать, *${name}*!\n\n` +
      `Вы подписаны на уведомления СКУД ПЛ №3.\n\n` +
      `Используйте кнопки меню ниже 👇`,
    { parse_mode: "Markdown", ...MAIN_KEYBOARD },
  );
});

// =========================
// ГЛАВНЫЙ ОБРАБОТЧИК СООБЩЕНИЙ
// Перехватывает кнопки меню и обычный текст
// =========================
bot.on("message", async (msg) => {
  const chatId = String(msg.chat.id);
  const text = (msg.text || "").trim();

  if (!text || text.startsWith("/")) return;

  const state = getState(chatId);

  // ── КНОПКИ ГЛАВНОГО МЕНЮ ──
  if (text === "📊 Сводка сегодня") {
    resetState(chatId);
    return handleStats(chatId);
  }

  if (text === "📂 Состояние группы") {
    setState(chatId, { step: "await_group_view" });
    const kb = await buildGroupKeyboard();
    return bot.sendMessage(chatId, "📂 Выберите группу:", kb);
  }

  if (text === "📥 Скачать Excel") {
    setState(chatId, { step: "await_group_report" });
    const kb = await buildGroupKeyboard();
    return bot.sendMessage(
      chatId,
      "📥 Выберите группу для скачивания журнала:",
      kb,
    );
  }

  if (text === "🔍 Найти студента") {
    setState(chatId, { step: "idle" });
    return bot.sendMessage(chatId, "🔍 Введите ФИО или ИНН студента:", {
      reply_markup: {
        keyboard: [[{ text: "❌ Отмена" }]],
        resize_keyboard: true,
      },
    });
  }

  if (text === "🔔 Привязать ребёнка") {
    setState(chatId, { step: "await_link_inn" });
    return bot.sendMessage(
      chatId,
      "🔔 Введите *ИНН* вашего ребёнка (14 цифр):\n\n_После привязки бот будет сообщать о каждом входе и выходе из лицея._",
      {
        parse_mode: "Markdown",
        reply_markup: {
          keyboard: [[{ text: "❌ Отмена" }]],
          resize_keyboard: true,
        },
      },
    );
  }

  if (text === "❌ Отмена") {
    resetState(chatId);
    return bot.sendMessage(chatId, "Главное меню 👇", MAIN_KEYBOARD);
  }

  // ── ПОШАГОВЫЕ ДИАЛОГИ ──

  // Шаг 1: пользователь выбрал группу для просмотра состояния
  if (state.step === "await_group_view") {
    resetState(chatId);
    return handleGroupReport(chatId, text);
  }

  // Шаг 1: пользователь выбрал группу для Excel
  if (state.step === "await_group_report") {
    // Проверяем что группа реально существует
    const group = await Group.findOne({ name: new RegExp(`^${text}$`, "i") });
    if (!group) {
      return bot.sendMessage(
        chatId,
        `❌ Группа *${text}* не найдена. Попробуйте ещё раз.`,
        { parse_mode: "Markdown" },
      );
    }
    // Переходим к выбору месяца
    setState(chatId, { step: "await_month_report", groupName: group.name });
    const { months, keyboard } = buildMonthKeyboard();
    return bot.sendMessage(
      chatId,
      `📅 Выберите месяц для группы *${group.name}*:`,
      { parse_mode: "Markdown", ...keyboard },
    );
  }

  // Шаг 2: пользователь выбрал месяц → отправляем Excel
  if (state.step === "await_month_report") {
    const month = parseMonthLabel(text);
    if (!month) {
      return bot.sendMessage(
        chatId,
        "⚠️ Пожалуйста, выберите месяц из списка.",
      );
    }
    const groupName = state.groupName;
    resetState(chatId);
    // Возвращаем главное меню
    bot.sendMessage(chatId, "⏳ Формирую файл...", MAIN_KEYBOARD);
    return sendExcelReport(chatId, groupName, month);
  }

  // Привязка ребёнка — ввод ИНН
  if (state.step === "await_link_inn") {
    resetState(chatId);
    return handleLink(chatId, text, msg.from.first_name);
  }

  // Если нет активного диалога — поиск студента по ФИО/ИНН
  return handleStudentSearch(chatId, text);
});

// =========================
// /stats команда (оставляем для совместимости)
// =========================
bot.onText(/\/stats/, async (msg) => {
  resetState(String(msg.chat.id));
  handleStats(String(msg.chat.id));
});

// =========================
// HANDLER: Сводка сегодня
// =========================
async function handleStats(chatId) {
  try {
    const totalStudents = await Student.countDocuments();
    const todayDateKey = getTodayKey();
    const bishkekDt = getBishkekNow();
    const todayIsWeekend = bishkekDt.getDay() === 0 || bishkekDt.getDay() === 6;

    const studentsWithData = await Student.find({
      $or: [
        { "skudDaily.date": todayDateKey },
        { "attendance.date": todayDateKey },
      ],
    }).populate("group");

    let presentToday = 0,
      latesToday = 0;

    studentsWithData.forEach((s) => {
      const skud = s.skudDaily?.find((r) => r.date === todayDateKey);
      const att = s.attendance?.find((a) => a.date === todayDateKey);
      let present = false,
        late = false;

      if (skud) {
        present = skud.present;
        if (skud.firstIn) {
          const shiftTime = s.group?.shiftStart || "09:30";
          const [shiftH, shiftM] = shiftTime.split(":").map(Number);
          const dt = new Date(skud.firstIn);
          if (!isNaN(dt.getTime())) {
            const totalMin = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
            if (totalMin > shiftH * 60 + shiftM) late = true;
          }
        }
      } else if (att) {
        present = att.present;
      }

      if (present) presentToday++;
      if (late) latesToday++;
    });

    const absentToday = todayIsWeekend ? 0 : totalStudents - presentToday;
    const percent =
      totalStudents > 0 ? Math.round((presentToday / totalStudents) * 100) : 0;

    // Строим мини-прогресс бар
    const filled = Math.round(percent / 10);
    const empty = 10 - filled;
    const bar = "█".repeat(filled) + "░".repeat(empty);

    bot.sendMessage(
      chatId,
      `📊 *Сводка на ${pad2(bishkekDt.getDate())}.${pad2(bishkekDt.getMonth() + 1)}*\n\n` +
        `${bar} *${percent}%*\n\n` +
        `👥 Всего в базе: *${totalStudents}*\n` +
        `✅ Присутствуют: *${presentToday}*\n` +
        `❌ Отсутствуют: *${absentToday}*\n` +
        `⏰ Опозданий: *${latesToday}*` +
        (todayIsWeekend ? "\n\n_📅 Сегодня выходной день_" : ""),
      { parse_mode: "Markdown", ...MAIN_KEYBOARD },
    );
  } catch (e) {
    bot.sendMessage(
      chatId,
      "❌ Ошибка формирования статистики.",
      MAIN_KEYBOARD,
    );
  }
}

// =========================
// HANDLER: Состояние группы
// =========================
async function handleGroupReport(chatId, groupName) {
  try {
    const group = await Group.findOne({
      name: new RegExp(`^${groupName}$`, "i"),
    });
    if (!group) {
      return bot.sendMessage(chatId, `❌ Группа *${groupName}* не найдена.`, {
        parse_mode: "Markdown",
        ...MAIN_KEYBOARD,
      });
    }

    const students = await Student.find({ group: group._id }).sort({ fio: 1 });
    if (!students.length) {
      return bot.sendMessage(
        chatId,
        `В группе *${group.name}* пока нет студентов.`,
        { parse_mode: "Markdown", ...MAIN_KEYBOARD },
      );
    }

    const todayDateKey = getTodayKey();
    const bishkekDt = getBishkekNow();
    const shiftTime = group.shiftStart || "09:30";
    const [shiftH, shiftM] = shiftTime.split(":").map(Number);
    const shiftTotalMin = shiftH * 60 + shiftM;

    let presentList = [],
      lateList = [],
      absentList = [];

    students.forEach((s) => {
      const skud = s.skudDaily?.find((r) => r.date === todayDateKey);
      const att = s.attendance?.find((a) => a.date === todayDateKey);
      let present = false,
        late = false,
        timeStr = "";

      if (skud) {
        present = skud.present;
        if (skud.firstIn) {
          const dt = new Date(skud.firstIn);
          if (!isNaN(dt.getTime())) {
            const totalMin = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
            const localH = Math.floor(totalMin / 60) % 24;
            const localM = totalMin % 60;
            timeStr = `${pad2(localH)}:${pad2(localM)}`;
            if (totalMin > shiftTotalMin) late = true;
          }
        }
      } else if (att) {
        present = att.present;
      }

      if (present) {
        if (late) lateList.push(`• ${s.fio} _(${timeStr})_`);
        else presentList.push(`• ${s.fio}${timeStr ? ` _(${timeStr})_` : ""}`);
      } else {
        absentList.push(`• ${s.fio}`);
      }
    });

    const total = students.length;
    const presentN = presentList.length + lateList.length;
    const percent = Math.round((presentN / total) * 100);
    const filled = Math.round(percent / 10);
    const bar = "█".repeat(filled) + "░".repeat(10 - filled);

    let reply =
      `📂 *${group.name}* — ${pad2(bishkekDt.getDate())}.${pad2(bishkekDt.getMonth() + 1)}\n` +
      `${bar} *${percent}%* явка\n` +
      `⏰ Начало смены: ${shiftTime}\n\n`;

    if (presentList.length)
      reply += `✅ *Вовремя (${presentList.length}):*\n${presentList.join("\n")}\n\n`;
    if (lateList.length)
      reply += `⏰ *Опоздали (${lateList.length}):*\n${lateList.join("\n")}\n\n`;
    if (absentList.length)
      reply += `❌ *Отсутствуют (${absentList.length}):*\n${absentList.join("\n")}`;

    if (reply.length > 4000)
      reply = reply.substring(0, 4000) + "\n_...список сокращён_";

    // Inline кнопка быстрого скачивания Excel
    const currentMonth = getCurrentMonth();
    bot.sendMessage(chatId, reply, {
      parse_mode: "Markdown",
      ...MAIN_KEYBOARD,
      reply_markup: {
        ...MAIN_KEYBOARD.reply_markup,
        inline_keyboard: [
          [
            {
              text: `📥 Excel за ${currentMonth}`,
              callback_data: `report:${group.name}:${currentMonth}`,
            },
          ],
        ],
      },
    });
  } catch (e) {
    console.error("handleGroupReport error:", e);
    bot.sendMessage(chatId, "⚠️ Ошибка генерации отчёта.", MAIN_KEYBOARD);
  }
}

// =========================
// HANDLER: Поиск студента
// =========================
async function handleStudentSearch(chatId, text) {
  try {
    let students = await Student.find({ inn: text }).populate("group");
    if (!students.length)
      students = await Student.find({ fio: new RegExp(text, "i") }).populate(
        "group",
      );

    if (!students.length) {
      return bot.sendMessage(
        chatId,
        "❌ *Студент не найден.*\nВведите ИНН (14 цифр) или Фамилию.",
        { parse_mode: "Markdown", ...MAIN_KEYBOARD },
      );
    }
    if (students.length > 5) {
      return bot.sendMessage(
        chatId,
        `⚠️ Найдено *${students.length}* студентов — уточните запрос.`,
        { parse_mode: "Markdown", ...MAIN_KEYBOARD },
      );
    }

    const todayDateKey = getTodayKey();
    const bishkekDt = getBishkekNow();
    const currentMonthPfx = `${bishkekDt.getFullYear()}-${pad2(bishkekDt.getMonth() + 1)}`;

    for (const student of students) {
      const skud = student.skudDaily?.find((r) => r.date === todayDateKey);
      const att = student.attendance?.find((a) => a.date === todayDateKey);

      let status = "❌ Отсутствует",
        timeStr = "—",
        lateInfo = "";
      let monthPresences = 0;

      student.skudDaily?.forEach((r) => {
        if (r.date.startsWith(currentMonthPfx) && r.present) monthPresences++;
      });
      student.attendance?.forEach((a) => {
        if (a.date.startsWith(currentMonthPfx) && a.present) monthPresences++;
      });

      if (skud?.present) {
        status = "✅ Пришёл вовремя";
        if (skud.firstIn) {
          const shiftTime = student.group?.shiftStart || "09:30";
          const [shiftH, shiftM] = shiftTime.split(":").map(Number);
          const dt = new Date(skud.firstIn);
          if (!isNaN(dt.getTime())) {
            const totalMin = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
            timeStr = `${pad2(Math.floor(totalMin / 60) % 24)}:${pad2(totalMin % 60)}`;
            if (totalMin > shiftH * 60 + shiftM) {
              status = "⏰ Опоздал";
              lateInfo = ` _(смена: ${shiftTime})_`;
            }
          }
        }
      } else if (att?.present) {
        status = "✅ Присутствует (ручная отметка)";
      }

      bot.sendMessage(
        chatId,
        `🧑‍🎓 *${student.fio}*\n` +
          `📂 Группа: ${student.group?.name || "Без группы"}\n` +
          `🪪 ИНН: \`${student.inn || "—"}\`\n\n` +
          `📅 *Сегодня:* ${status}\n` +
          `🕒 Прибыл в: *${timeStr}*${lateInfo}\n\n` +
          `📊 Присутствий за месяц: *${monthPresences}*`,
        { parse_mode: "Markdown", ...MAIN_KEYBOARD },
      );
    }
  } catch (e) {
    console.error("handleStudentSearch error:", e);
    bot.sendMessage(chatId, "⚠️ Ошибка при поиске.", MAIN_KEYBOARD);
  }
}

// =========================
// HANDLER: Привязка ребёнка
// =========================
async function handleLink(chatId, inn, firstName) {
  try {
    const student = await Student.findOne({ inn });
    if (!student) {
      return bot.sendMessage(
        chatId,
        `❌ Студент с ИНН *${inn}* не найден.\nПроверьте правильность ввода.`,
        { parse_mode: "Markdown", ...MAIN_KEYBOARD },
      );
    }
    await TgSub.findOneAndUpdate(
      { chatId },
      { $addToSet: { linkedInns: inn }, name: firstName },
      { upsert: true },
    );
    bot.sendMessage(
      chatId,
      `🔔 *Подписка оформлена!*\n\n` +
        `Студент: *${student.fio}*\n` +
        `Группа: ${student.group ? (await Group.findById(student.group))?.name || "—" : "—"}\n\n` +
        `Теперь бот будет автоматически сообщать о входе и выходе из лицея.`,
      { parse_mode: "Markdown", ...MAIN_KEYBOARD },
    );
  } catch (e) {
    bot.sendMessage(chatId, "⚠️ Ошибка при привязке.", MAIN_KEYBOARD);
  }
}

// =========================
// INLINE CALLBACK: кнопка "Excel за месяц" под сообщением группы
// =========================
bot.on("callback_query", async (query) => {
  const chatId = String(query.message.chat.id);
  const data = query.data || "";

  if (data.startsWith("report:")) {
    const parts = data.split(":");
    const groupName = parts[1];
    const month = parts[2];
    bot.answerCallbackQuery(query.id, { text: "⏳ Формирую файл..." });
    await sendExcelReport(chatId, groupName, month);
  }
});

// =========================
// CORE: генерация и отправка Excel
// =========================
async function sendExcelReport(chatId, groupName, month) {
  try {
    const group = await Group.findOne({
      name: new RegExp(`^${groupName}$`, "i"),
    });
    if (!group) {
      return bot.sendMessage(chatId, `❌ Группа *${groupName}* не найдена.`, {
        parse_mode: "Markdown",
        ...MAIN_KEYBOARD,
      });
    }

    const [year, mon] = month.split("-").map(Number);
    const daysInMonth = new Date(year, mon, 0).getDate();
    const shiftTime = group.shiftStart || "09:30";
    const [shiftH, shiftM] = shiftTime.split(":").map(Number);
    const shiftTotalMin = shiftH * 60 + shiftM;
    const students = await Student.find({ group: group._id }).sort({ fio: 1 });

    if (!students.length) {
      return bot.sendMessage(
        chatId,
        `❌ В группе *${group.name}* нет студентов.`,
        { parse_mode: "Markdown", ...MAIN_KEYBOARD },
      );
    }

    const bishkekDt = getBishkekNow();
    const todayDateKey = `${bishkekDt.getFullYear()}-${pad2(bishkekDt.getMonth() + 1)}-${pad2(bishkekDt.getDate())}`;

    const excelData = [];
    const headerRow = ["№", "ФИО Студента"];
    for (let d = 1; d <= daysInMonth; d++) {
      const dateObj = new Date(Date.UTC(year, mon - 1, d));
      const dayNames = ["Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"];
      headerRow.push(`${d}\n${dayNames[dateObj.getUTCDay()]}`);
    }
    headerRow.push("Присутствий", "Опозданий", "Отсутствий");
    excelData.push(headerRow);

    let grandPresences = 0,
      grandAbsences = 0,
      grandLates = 0;

    students.forEach((s, idx) => {
      const row = [idx + 1, s.fio];
      let presences = 0,
        absences = 0,
        lates = 0;

      for (let d = 1; d <= daysInMonth; d++) {
        const dateKey = `${year}-${pad2(mon)}-${pad2(d)}`;
        const dateObj = new Date(Date.UTC(year, mon - 1, d));
        const isWeekend =
          dateObj.getUTCDay() === 0 || dateObj.getUTCDay() === 6;
        const skud = s.skudDaily?.find((r) => r.date === dateKey);
        const att = s.attendance?.find((a) => a.date === dateKey);
        let present = false,
          late = false;

        if (skud) {
          present = skud.present;
          if (skud.firstIn) {
            const dt = new Date(skud.firstIn);
            if (!isNaN(dt.getTime())) {
              const totalMin = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
              if (totalMin > shiftTotalMin) late = true;
            }
          }
        } else if (att) {
          present = att.present;
        }

        if (present) {
          presences++;
          if (late) {
            lates++;
            row.push("О");
          } else row.push("П");
        } else if (!isWeekend && dateKey <= todayDateKey) {
          absences++;
          row.push("Н");
        } else {
          row.push(isWeekend ? "вых" : "");
        }
      }

      row.push(presences, lates, absences);
      grandPresences += presences;
      grandAbsences += absences;
      grandLates += lates;
      excelData.push(row);
    });

    // Итоговая строка
    const totalRow = ["", "ИТОГО ПО ГРУППЕ"];
    for (let d = 0; d < daysInMonth; d++) totalRow.push("");
    totalRow.push(grandPresences, grandLates, grandAbsences);
    excelData.push(totalRow);

    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet(excelData);
    const colWidths = [{ wch: 4 }, { wch: 32 }];
    for (let d = 0; d < daysInMonth; d++) colWidths.push({ wch: 5 });
    colWidths.push({ wch: 11 }, { wch: 11 }, { wch: 11 });
    ws["!cols"] = colWidths;
    xlsx.utils.book_append_sheet(wb, ws, `Журнал ${month}`);
    const buf = xlsx.write(wb, { type: "buffer", bookType: "xlsx" });

    const MONTHS_RU = [
      "январь",
      "февраль",
      "март",
      "апрель",
      "май",
      "июнь",
      "июль",
      "август",
      "сентябрь",
      "октябрь",
      "ноябрь",
      "декабрь",
    ];
    const caption =
      `📊 *Журнал посещаемости*\n` +
      `📂 Группа: *${group.name}*\n` +
      `📅 ${MONTHS_RU[mon - 1]} ${year}\n` +
      `👥 Студентов: *${students.length}*\n\n` +
      `П — присутствовал, О — опоздал, Н — не был, вых — выходной`;

    bot.sendDocument(
      chatId,
      buf,
      { caption, parse_mode: "Markdown" },
      {
        filename: `Journal_${group.name}_${month}.xlsx`,
        contentType:
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      },
    );
  } catch (e) {
    console.error("sendExcelReport error:", e);
    bot.sendMessage(
      chatId,
      "⚠️ Ошибка при формировании Excel файла.",
      MAIN_KEYBOARD,
    );
  }
}

// =========================
// Утренняя сводка пн-пт в 10:00 Бишкек
// =========================
cron.schedule("0 4 * * 1-5", async () => {
  notifyTg(
    "🔔 *Доброе утро!*\n\nНажмите кнопку *📊 Сводка сегодня* для актуальных данных на 10:00.",
  );
});
/* ===========================
   START
=========================== */
app.listen(3000, () =>
  console.log("🚀 Server running on http://localhost:3000"),
);
