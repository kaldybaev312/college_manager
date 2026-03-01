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

const app = express();
const upload = multer({ dest: "uploads/" });
if (!fs.existsSync("uploads")) fs.mkdirSync("uploads");

app.use(express.json({ limit: "2mb" }));
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
   PERSON RESOLVE (HIKVISION)
=========================== */

async function fetchPersonName(deviceIp, employeeNo) {
  const client = new DigestClient(HIK_USER, HIK_PASS, { algorithm: "MD5" });

  const url = `http://${deviceIp}/ISAPI/AccessControl/UserInfo/Search?format=json`;

  const payload = {
    UserInfoSearchCond: {
      searchID: "1",
      searchResultPosition: 0,
      maxResults: 5,
      EmployeeNoList: [{ employeeNo: String(employeeNo) }],
    },
  };

  const res = await client.fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!res.ok) return null;

  const json = await res.json();
  return json?.UserInfoSearch?.UserInfo?.[0]?.name || null;
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
      const eventIso = new Date(entry.time).toISOString();
      const dateKey = isoToDateKey(eventIso);

      const rawName = await fetchPersonName(deviceIp, employeeNo);
      if (!rawName) {
        pushSkudEvent({
          ok: false,
          reason: "person_not_found_in_terminal",
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
        const candidates = await Student.find({
          fio: new RegExp(fioNorm.split(" ")[0], "i"),
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

      // attendance
      const exist = student.attendance.find((a) => a.date === dateKey);
      if (exist) exist.present = true;
      else student.attendance.push({ date: dateKey, present: true });

      // skudDaily
      let rec = student.skudDaily.find((r) => r.date === dateKey);
      if (!rec) {
        rec = { date: dateKey, inCount: 0, outCount: 0, present: false };
        student.skudDaily.push(rec);
        rec = student.skudDaily[student.skudDaily.length - 1];
      }

      if (direction === "in") {
        rec.inCount++;
        rec.present = true;
        rec.firstIn = rec.firstIn || eventIso;
        rec.lastIn = eventIso;
      } else {
        rec.outCount++;
        rec.firstOut = rec.firstOut || eventIso;
        rec.lastOut = eventIso;
      }

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

// Список всех групп
app.get("/api/admin/groups-list", async (req, res) => {
  try {
    const groups = await Group.find().sort({ name: 1 });
    res.json(groups);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Удалить группу и всех её студентов
app.delete("/api/admin/group/:id", async (req, res) => {
  try {
    await Student.deleteMany({ group: req.params.id });
    await Group.findByIdAndDelete(req.params.id);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Очистить студентов группы (группа остаётся)
app.delete("/api/admin/group/:id/clear", async (req, res) => {
  try {
    await Student.deleteMany({ group: req.params.id });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/* ===========================
   ADMIN — STUDENTS
=========================== */

// Студенты конкретной группы
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

// Поиск студентов по ФИО
app.get("/api/admin/search", async (req, res) => {
  try {
    const q = (req.query.q || "").trim();
    const filter = q.length >= 2 ? { fio: new RegExp(q, "i") } : {};
    const students = await Student.find(filter)
      .populate("group")
      .sort({ fio: 1 })
      .limit(50);
    res.json(students);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Обновить профессию группы
app.put("/api/admin/group/:id/set-prof", async (req, res) => {
  try {
    const { profId } = req.body;
    const prof = PROFESSIONS.find((p) => p.id == profId);
    if (!prof) return res.status(404).json({ error: "Профессия не найдена" });

    await Group.findByIdAndUpdate(req.params.id, {
      profRu: prof.ru,
      profKg: prof.kg,
      duration: prof.duration,
    });

    res.json({
      ok: true,
      profRu: prof.ru,
      profKg: prof.kg,
      duration: prof.duration,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Редактировать студента
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
    const { gid, month } = req.params; // month = "2026-02"
    const [year, mon] = month.split("-").map(Number);
    const daysInMonth = new Date(year, mon, 0).getDate();

    const students = await Student.find({ group: gid }).sort({ fio: 1 });

    const matrix = students.map((s) => {
      const days = [];
      for (let d = 1; d <= daysInMonth; d++) {
        const dateKey = `${year}-${pad2(mon)}-${pad2(d)}`;

        const skud = s.skudDaily?.find((r) => r.date === dateKey);
        const att = s.attendance?.find((a) => a.date === dateKey);

        if (skud) {
          let timeStr = "";
          let late = false;

          if (skud.firstIn) {
            const dt = new Date(skud.firstIn);
            // UTC+6 (Bishkek)
            const totalMin = dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360;
            const localH = Math.floor(totalMin / 60) % 24;
            const localM = totalMin % 60;
            timeStr = `${pad2(localH)}:${pad2(localM)}`;
            late = localH * 60 + localM > 8 * 60 + 30; // опоздание после 08:30
          }

          days.push({ present: skud.present, late, time: timeStr });
        } else if (att) {
          days.push({ present: att.present, late: false, time: "" });
        } else {
          days.push({ present: false, late: false, time: "" });
        }
      }
      return { fio: s.fio, days };
    });

    res.json({ daysInMonth, matrix });
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
   ADMIN — СПРАВКИ (DOCX)
=========================== */

app.get("/api/admin/print/:id/:type", async (req, res) => {
  try {
    const student = await Student.findById(req.params.id).populate("group");
    if (!student) return res.status(404).send("Студент не найден");

    const type = req.params.type; // common | army | social
    const tplMap = {
      common: "template_common.docx",
      army: "template_army.docx",
      social: "template_social.docx",
    };

    const tplFile = tplMap[type] || tplMap.common;
    const tplPath = path.join(__dirname, "templates", tplFile);

    if (!fs.existsSync(tplPath)) {
      // Шаблон ещё не создан — отдаём заглушку JSON
      return res.json({
        message: `Шаблон ${tplFile} не найден. Положите его в папку templates/`,
        student: {
          fio: student.fio,
          birthDate: student.birthDate,
          inn: student.inn,
          lyceumId: student.lyceumId,
        },
        group: student.group
          ? {
              name: student.group.name,
              profRu: student.group.profRu,
              profKg: student.group.profKg,
              duration: student.group.duration,
            }
          : null,
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

    // Формируем дату в формате «ДД» месяц ГГГГ
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
      // camelCase варианты
      profRu: student.group?.profRu || "—",
      profKg: student.group?.profKg || "—",
      duration: student.group?.duration || "—",
      // snake_case варианты (для совместимости с шаблонами)
      prof_ru: student.group?.profRu || "—",
      prof_kg: student.group?.profKg || "—",
      // другие возможные имена тегов
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

/* ===========================
   ADMIN — ВЕДОМОСТЬ ПИТАНИЯ (DOCX)
=========================== */

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

    // Строим таблицу посещаемости по дням
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
        days: dayCells, // массив строк
        totalDays: presentDays,
        totalAmount: presentDays * 60,
      };
    });

    const tplPath = path.join(__dirname, "templates", "template_food.docx");

    if (!fs.existsSync(tplPath)) {
      // Шаблон отсутствует — отдаём JSON-заглушку
      return res.json({
        message:
          "Шаблон template_food.docx не найден. Положите его в папку templates/",
        group: group.name,
        month,
        rows: rows.map((r) => ({
          fio: r.fio,
          days: r.totalDays,
          amount: r.totalAmount,
        })),
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
      // ── Заголовок (теги шаблона: {month_name}, {year}, {group}) ──
      month_name: MONTHS_RU[mon - 1],
      year: String(year),
      group: group.name,
      groupName: group.name, // запасной вариант

      // ── Массив строк (тег: {#st}...{/st}) ──
      // Каждая строка: {fio}, {lyceum_id}, {days}, {price}, {sum}
      st: rows.map((r) => ({
        fio: r.fio,
        lyceum_id: r.lyceumId || "—",
        days: String(r.totalDays),
        price: String(PRICE_PER_DAY),
        sum: String(r.totalAmount),
        // запасные имена на случай других тегов в шаблоне
        num: String(r.num),
        lyceumId: r.lyceumId || "—",
        totalDays: String(r.totalDays),
        totalAmount: String(r.totalAmount),
      })),

      // ── Итого (тег: {total}) ──
      total: String(grandTotal),
      grandTotal: String(grandTotal),

      // ── Дополнительные поля ──
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
   START
=========================== */

app.listen(3000, () =>
  console.log("🚀 Server running on http://localhost:3000"),
);
