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
  "mongodb+srv://ilimkaldybaev5_db_user:liceyStudents>@riestr.uki8ep8.mongodb.net/?appName=riestr";

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
        rec = {
          date: dateKey,
          inCount: 0,
          outCount: 0,
          present: false,
        };
        student.skudDaily.push(rec);
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
      pushSkudEvent({
        ok: false,
        reason: e.message,
      });
    }
  }

  res.json({ success: true, okCount, failCount });
});

/* ===========================
   START
=========================== */

app.listen(3000, () =>
  console.log("🚀 Server running on http://localhost:3000"),
);
