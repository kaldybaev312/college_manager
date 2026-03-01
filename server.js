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

// ✅ Digest auth for Hikvision
const DigestClient = require("digest-fetch").default;

const app = express();
const upload = multer({ dest: "uploads/" });
if (!fs.existsSync("uploads")) fs.mkdirSync("uploads");

// --- ПОДКЛЮЧЕНИЕ К MONGODB ---
const MONGO_URI =
  process.env.MONGO_URI ||
  "mongodb+srv://ilimkaldybaev5_db_user:liceyStudents@riestr.uki8ep8.mongodb.net/PL3_Database?retryWrites=true&w=majority&appName=riestr";

mongoose
  .connect(MONGO_URI)
  .then(() => console.log("✅ База данных подключена"))
  .catch((e) => console.log("❌ Ошибка MongoDB:", e.message));

// --- СЕКРЕТ ДЛЯ АГЕНТА ---
const AGENT_API_KEY = process.env.AGENT_API_KEY || "pl3_secret_key";

// --- Hikvision учётка (лучше в ENV) ---
const HIK_USER = process.env.HIK_USER || "admin";
const HIK_PASS = process.env.HIK_PASS || "Qwerty#12";

// --- СПРАВОЧНИК ПРОФЕССИЙ ---
const PROFESSIONS = [
  { id: 1, ru: "Токарь", kg: "Токарь", duration: "2 года" },
  { id: 2, ru: "Электрогазосварщик", kg: "Электргазоширетүүчү", duration: "2 года" },
  { id: 3, ru: "Электромонтер по ремонту и обслуживанию электрооборудования", kg: "Электр жабдууларын оңдоо жана тейлөө боюнча электромонтер", duration: "2 года" },
  { id: 4, ru: "Мастер по ремонту и обслуживанию бытовой техники", kg: "Турмуш-тиричилик техникасын оңдоо жана тейлөө боюнча мастер", duration: "2 года" },
  { id: 5, ru: "Электрик по ремонту автомобильного электрооборудования", kg: "Автоунаанын электр жабдууларын оңдоо боюнча электрик", duration: "2 года" },
  { id: 6, ru: "Разработчик Web и мультимедийных приложений", kg: "Web жана мультимедиалык тиркемелерди иштеп чыгуучу", duration: "2 года" },
  { id: 7, ru: "Оператор цифровой печати", kg: "Санариптик басма оператору", duration: "2 года" },
  { id: 8, ru: "Повар", kg: "Ашпозчу", duration: "2 года" },
  { id: 9, ru: "Переплетчик", kg: "Түптөөчү", duration: "10 месяцев" },
  { id: 10, ru: "Электромонтер (10 м.)", kg: "Электромонтер", duration: "10 месяцев" },
  { id: 11, ru: "Автослесарь-Автоэлектрик", kg: "Автослесарь-Автоэлектрик", duration: "10 месяцев" },
  { id: 12, ru: "Программист", kg: "Программист", duration: "10 месяцев" },
  { id: 13, ru: "Оператор печатного оборудования", kg: "Басма жабдууларын оператору", duration: "2 года" },
];

// --- МОДЕЛИ ---
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

    // старое поле (оставляем, ваш журнал работает)
    attendance: [{ date: String, present: Boolean }],

    // ✅ новое: расширенный СКУД-день
    skudDaily: [
      new mongoose.Schema(
        {
          date: String, // YYYY-MM-DD
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

app.use(express.json({ limit: "2mb" }));
app.use(express.static("public"));

// =======================
// HELPERS
// =======================
function formatExcelDate(serial) {
  if (!serial) return "—";
  if (!isNaN(serial) && typeof serial === "number") {
    const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
    return date.toLocaleDateString("ru-RU");
  }
  return String(serial).trim();
}

function pad2(n) {
  return String(n).padStart(2, "0");
}

function isoToDateKey(iso) {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return null;
  return `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
}

function minIso(a, b) {
  if (!a) return b;
  if (!b) return a;
  return new Date(a) <= new Date(b) ? a : b;
}
function maxIso(a, b) {
  if (!a) return b;
  if (!b) return a;
  return new Date(a) >= new Date(b) ? a : b;
}

function digitsOnly(s) {
  return String(s || "").replace(/\D/g, "");
}
function normalizeINN(s) {
  const d = digitsOnly(s);
  if (d.length < 12) return null;
  return d;
}
function normalizeFIO(s) {
  return String(s || "")
    .replace(/[.,]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

// parse INN / FIO from Hikvision "name" field ("ФИО ИНН")
function extractINN(s) {
  const m = String(s || "").match(/\b(\d{12,16})\b/);
  return m ? m[1] : null;
}
function extractFIO(s) {
  const inn = extractINN(s);
  let fio = String(s || "");
  if (inn) fio = fio.replace(inn, " ");
  return fio.replace(/\s+/g, " ").trim();
}

// ===========================
// CACHE employeeNo -> { name, inn, fio, fioNorm, ts }
// ===========================
const PERSON_CACHE = new Map();
const CACHE_TTL_MS = 24 * 60 * 60 * 1000;

function cacheGet(employeeNo) {
  const v = PERSON_CACHE.get(employeeNo);
  if (!v) return null;
  if (Date.now() - v.ts > CACHE_TTL_MS) {
    PERSON_CACHE.delete(employeeNo);
    return null;
  }
  return v;
}
function cacheSet(employeeNo, value) {
  PERSON_CACHE.set(employeeNo, { ...value, ts: Date.now() });
}

// Hikvision UserInfo/Search по employeeNo
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

  const text = await res.text();
  if (!res.ok) {
    console.log("❌ HIK UserInfo/Search:", deviceIp, res.status, text.slice(0, 250));
    return null;
  }

  const json = JSON.parse(text);
  const user = json?.UserInfoSearch?.UserInfo?.[0];
  return user?.name || null;
}

async function resolvePerson(deviceIp, employeeNo) {
  const cached = cacheGet(employeeNo);
  if (cached?.inn) return cached;

  const name = await fetchPersonName(deviceIp, employeeNo);
  if (!name) return null;

  const innRaw = extractINN(name);
  const fio = extractFIO(name);

  const inn = normalizeINN(innRaw);
  const fioNorm = normalizeFIO(fio);

  const v = { name, inn, fio, fioNorm };
  cacheSet(employeeNo, v);
  return v;
}

// =======================
// HEALTH
// =======================
app.get("/health", (req, res) => res.json({ ok: true }));

// =======================
// SKUD: прием данных от агента
// payload: { apiKey, data:[{id,deviceIp,direction,time,minor}] }
// =======================
app.post("/api/agent/sync", async (req, res) => {
  const { apiKey, data } = req.body;

  if (apiKey !== AGENT_API_KEY) return res.status(403).send("Forbidden");
  if (!Array.isArray(data))
    return res.status(400).json({ success: false, error: "data must be array" });

  let okCount = 0;
  let failCount = 0;

  for (const entry of data) {
    try {
      const employeeNo = entry?.id?.toString?.().trim();
      const deviceIp = String(entry?.deviceIp || "").trim();
      const direction = String(entry?.direction || "unknown").toLowerCase(); // in/out
      const eventIso = entry?.time ? new Date(entry.time).toISOString() : new Date().toISOString();
      const dateKey = isoToDateKey(eventIso);

      if (!employeeNo || !deviceIp || !dateKey) {
        failCount++;
        continue;
      }

      // 1) Resolve INN/FIO from Hikvision by employeeNo
      const person = await resolvePerson(deviceIp, employeeNo);
      if (!person?.inn) {
        console.log("❌ Person not resolved or no INN:", { employeeNo, deviceIp, name: person?.name });
        failCount++;
        continue;
      }

      const innNorm = person.inn;      // уже цифры
      const fioNorm = person.fioNorm;  // уже нормализовано

      // 2) Find student: 2 этапа (INN -> fallback FIO)
      let student = null;

      // 2.1 быстрый: точное совпадение inn (если база чистая)
      student = await Student.findOne({ inn: innNorm });

      // 2.2 грязная база: ищем кандидатов и сравниваем normalizeINN(c.inn)
      if (!student) {
        const anchor = innNorm.slice(-6); // последние 6 цифр
        const candidates = await Student.find({ inn: { $regex: anchor } }).limit(80);
        student = candidates.find((c) => normalizeINN(c.inn) === innNorm) || null;
      }

      // 2.3 fallback: по ФИО (если ИНН в базе пустой/битый)
      if (!student && fioNorm) {
        const parts = fioNorm.split(" ").filter(Boolean);
        const must = parts.slice(0, 2); // фамилия+имя
        const re = must.length ? new RegExp(must.join(".*"), "i") : null;

        if (re) {
          const candidates = await Student.find({ fio: re }).limit(80);
          student = candidates.find((c) => normalizeFIO(c.fio) === fioNorm) || null;
        }
      }

      if (!student) {
        console.log("❌ Student not found:", { innNorm, fioFromHik: person.fio, employeeNo });

        // подсказка: 5 ближайших по фамилии
        try {
          const lastName = fioNorm ? fioNorm.split(" ")[0] : null;
          if (lastName) {
            const near = await Student.find({ fio: new RegExp(lastName, "i") })
              .limit(5)
              .select("fio inn lyceumId");
            console.log("🔎 Near:", near.map(x => ({ fio: x.fio, inn: x.inn, lyceumId: x.lyceumId })));
          }
        } catch {}

        failCount++;
        continue;
      }

      // 3) Update old attendance (present=true) — чтобы ваш журнал работал
      const att = student.attendance || [];
      const exist = att.find((a) => a.date === dateKey);
      if (exist) exist.present = true;
      else att.push({ date: dateKey, present: true });
      student.attendance = att;

      // 4) Update skudDaily
      const daily = student.skudDaily || [];
      let rec = daily.find((x) => x.date === dateKey);
      if (!rec) {
        rec = {
          date: dateKey,
          firstIn: null,
          lastIn: null,
          firstOut: null,
          lastOut: null,
          inCount: 0,
          outCount: 0,
          present: false,
        };
        daily.push(rec);
      }

      if (direction === "in") {
        rec.inCount += 1;
        rec.firstIn = minIso(rec.firstIn, eventIso);
        rec.lastIn = maxIso(rec.lastIn, eventIso);
        rec.present = true;
      } else if (direction === "out") {
        rec.outCount += 1;
        rec.firstOut = minIso(rec.firstOut, eventIso);
        rec.lastOut = maxIso(rec.lastOut, eventIso);
      }

      student.skudDaily = daily;
      await student.save();

      okCount++;
      console.log("✅ SKUD:", direction.toUpperCase(), student.fio, "INN:", student.inn, "emp:", employeeNo, "from", deviceIp);
    } catch (e) {
      failCount++;
      console.log("❌ sync error:", e.message);
    }
  }

  res.json({ success: true, okCount, failCount });
});

// (опционально) просмотр skudDaily по INN
app.get("/api/admin/skud-daily-by-inn/:inn", async (req, res) => {
  const inn = normalizeINN(req.params.inn);
  if (!inn) return res.status(400).json({ ok: false, error: "bad inn" });

  const s = await Student.findOne({ inn });
  if (!s) return res.status(404).json({ ok: false, error: "not found" });

  res.json({ ok: true, fio: s.fio, inn: s.inn, skudDaily: s.skudDaily?.slice(-40) || [] });
});

// --- ЖУРНАЛ: МАТРИЦА ПОСЕЩАЕМОСТИ ---
app.get("/api/admin/attendance-matrix/:groupId/:month", async (req, res) => {
  const students = await Student.find({ group: req.params.groupId }).sort({ fio: 1 });
  const [year, month] = req.params.month.split("-").map(Number);
  const daysInMonth = new Date(year, month, 0).getDate();

  const matrix = students.map((s) => {
    const days = [];
    for (let d = 1; d <= daysInMonth; d++) {
      const dStr = `${req.params.month}-${d.toString().padStart(2, "0")}`;
      days.push({ day: d, present: s.attendance.some((a) => a.date === dStr) });
    }
    return { fio: s.fio, days };
  });

  res.json({ daysInMonth, matrix });
});

// --- ПИТАНИЕ: ГЕНЕРАЦИЯ ВЕДОМОСТИ В WORD ---
app.get("/api/admin/report-food/:groupId/:month", async (req, res) => {
  const group = await Group.findById(req.params.groupId);
  const students = await Student.find({ group: group._id }).sort({ fio: 1 });

  const monthNames = ["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
  const monthIndex = parseInt(req.params.month.split("-")[1]) - 1;
  const monthName = monthNames[monthIndex];

  const tableData = students.map((s, i) => {
    const days = s.attendance.filter((a) => a.date.startsWith(req.params.month) && a.present).length;
    return { idx: i + 1, fio: s.fio, lyceum_id: s.lyceumId, days, price: 60, sum: days * 60 };
  });

  const total = tableData.reduce((a, b) => a + b.sum, 0);
  const zip = new PizZip(fs.readFileSync(path.resolve(__dirname, "templates", "template_food.docx"), "binary"));
  const doc = new Docxtemplater(zip);

  doc.render({
    group: group.name,
    month_name: monthName,
    year: req.params.month.split("-")[0],
    st: tableData,
    total,
  });

  res.set({ "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  res.send(doc.getZip().generate({ type: "nodebuffer" }));
});

// --- ИМПОРТ ---
app.post("/api/admin/import", upload.single("file"), async (req, res) => {
  try {
    const { groupName, profId } = req.body;
    const prof = PROFESSIONS.find((p) => p.id == profId);

    let g = await Group.findOneAndUpdate(
      { name: groupName },
      { profRu: prof.ru, profKg: prof.kg, duration: prof.duration },
      { upsert: true, new: true },
    );

    const wb = xlsx.readFile(req.file.path);
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet);

    const studentData = rows.map((r) => ({
      fio: r["ФИО"] ? String(r["ФИО"]).trim() : "Без имени",
      birthDate: formatExcelDate(r["Дата рождение"]),
      inn: normalizeINN(r["ИНН"]) || "—", // ✅ нормализуем
      lyceumId: r["Поименный номер"] ? String(r["Поименный номер"]).trim() : "—",
      group: g._id,
    }));

    await Student.insertMany(studentData);
    fs.unlinkSync(req.file.path);
    res.json({ success: true });
  } catch (e) {
    console.log("❌ Import error:", e.message);
    res.status(500).send("Ошибка импорта");
  }
});

// --- ПЕЧАТЬ ---
app.get("/api/admin/print/:id/:type", async (req, res) => {
  try {
    const student = await Student.findById(req.params.id).populate("group");
    const templatePath = path.resolve(__dirname, "templates", `template_${req.params.type}.docx`);

    const zip = new PizZip(fs.readFileSync(templatePath, "binary"));
    const doc = new Docxtemplater(zip, {
      modules: [
        new ImageModule({
          centered: false,
          getImage: (v) => Buffer.from(v, "base64"),
          getSize: () => [100, 100],
        }),
      ],
    });

    const qr = await QRCode.toBuffer(`http://${req.headers.host}/verify/${student._id}`);

    doc.render({
      fio: student.fio,
      group: student.group ? student.group.name : "—",
      prof_ru: student.group ? student.group.profRu : "—",
      prof_kg: student.group ? student.group.profKg : "—",
      duration: student.group ? student.group.duration : "—",
      inn: student.inn,
      birth: student.birthDate,
      lyceum_id: student.lyceumId,
      date: new Date().toLocaleDateString("ru-RU"),
      qr_code: qr.toString("base64"),
    });

    res.set({ "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
    res.send(doc.getZip().generate({ type: "nodebuffer" }));
  } catch (e) {
    console.log("❌ Print error:", e.message);
    res.status(500).send("Ошибка печати");
  }
});

app.get("/api/admin/professions", (req, res) => res.json(PROFESSIONS));
app.get("/api/admin/groups-list", async (req, res) => res.json(await Group.find().sort({ name: 1 })));
app.get("/api/admin/group-students/:id", async (req, res) => res.json(await Student.find({ group: req.params.id }).sort({ fio: 1 })));
app.get("/api/admin/search", async (req, res) =>
  res.json(
    await Student.find(req.query.q ? { fio: new RegExp(req.query.q, "i") } : {})
      .populate("group")
      .sort({ fio: 1 }),
  ),
);
app.put("/api/admin/student/:id", async (req, res) => {
  await Student.findByIdAndUpdate(req.params.id, req.body);
  res.json({ success: true });
});
app.delete("/api/admin/group/:id/clear", async (req, res) => {
  await Student.deleteMany({ group: req.params.id });
  res.json({ success: true });
});
app.delete("/api/admin/group/:id", async (req, res) => {
  await Student.deleteMany({ group: req.params.id });
  await Group.findByIdAndDelete(req.params.id);
  res.json({ success: true });
});

app.listen(3000, () => console.log("🚀 Server: http://localhost:3000"));
