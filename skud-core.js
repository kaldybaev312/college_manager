// СКУД: общее ядро. Модели, журнал действий, уведомления, вспомогательные функции.
// Используется модулями skud-register, skud-agent-api и skud-staff (подключаются через skud.js).
const crypto = require("crypto");
const mongoose = require("mongoose");

const { Schema } = mongoose;
const OID = Schema.Types.ObjectId;

const NAME_MAX = Number(process.env.SKUD_NAME_MAX || 32); // лимит поля Name на терминале
const MIN_SIDE = 300; // минимальная сторона фото (px)
const MAX_PHOTO = 200 * 1024; // терминал Hikvision принимает JPEG примерно до 200 КБ

const model = (name, schema) => mongoose.models[name] || mongoose.model(name, schema);

/* ============================ модели ============================ */

// Задание для локального агента в лицее: записать человека на терминалы или удалить с них
const SkudJob = model(
  "SkudJob",
  new Schema(
    {
      type: { type: String, enum: ["register", "delete"], default: "register" },
      student: { type: OID, ref: "Student", index: true },
      fio: String,
      inn: String,
      employeeNo: { type: String, index: true }, // Employee ID на терминале (= Student.lyceumId)
      deviceName: String, // Name на терминале = ФИО + ИНН
      photo: Buffer, // хранится только пока задание не выполнено
      validUntil: Date, // до какого дня действует пропуск (иначе до 2035)
      status: {
        type: String,
        enum: ["pending", "processing", "done", "partial", "failed"],
        default: "pending",
        index: true,
      },
      attempts: { type: Number, default: 0 },
      leaseUntil: Date,
      retryAt: Date, // автоповтор после сбоя терминала
      results: [{ _id: false, ip: String, ok: Boolean, error: String }],
      by: String, // кто поставил задание
    },
    { timestamps: true },
  ),
);

// Журнал действий (кто, кого и когда). Хранится год.
const SkudAudit = model(
  "SkudAudit",
  new Schema(
    {
      at: { type: Date, default: Date.now, expires: 365 * 24 * 3600 },
      actor: String,
      role: String,
      action: { type: String, index: true },
      target: String,
      details: String,
      ip: String,
    },
    { versionKey: false },
  ),
);

// Состояние терминалов (сигналы от агента раз в минуту)
const SkudDevice = model(
  "SkudDevice",
  new Schema(
    {
      ip: { type: String, unique: true },
      online: Boolean,
      lastSeen: Date, // последний сигнал агента
      lastOk: Date, // последний раз был в сети
      offlineSince: Date,
      offlineNotified: Boolean,
      silentNotified: Boolean,
      latencyMs: Number,
      users: Number,
      error: String,
    },
    { timestamps: true },
  ),
);

// Результат сверки терминалов с базой (одна запись _id="last")
const SkudReport = model(
  "SkudReport",
  new Schema(
    {
      _id: String,
      requested: Boolean,
      requestedAt: Date,
      at: Date,
      sig: String,
      devices: [
        {
          _id: false,
          ip: String,
          ok: Boolean,
          error: String,
          total: Number,
          extra: [{ _id: false, employeeNo: String, name: String }],
          missing: [{ _id: false, employeeNo: String, fio: String }],
        },
      ],
    },
    { versionKey: false },
  ),
);

/* ======================= вспомогательные ======================== */

const STATUS_RU = {
  none: "не отправлен",
  pending: "в очереди",
  processing: "отправляется",
  done: "в СКУД",
  partial: "не на всех терминалах",
  failed: "ошибка",
  removing: "удаляется",
  removed: "удалён с терминалов",
  remove_failed: "ошибка удаления",
};

// Name на терминале: ФИО + ИНН. ИНН всегда целиком: по нему /api/agent/sync находит студента.
function buildDeviceName(fio, inn) {
  const clean = String(fio || "").replace(/\s+/g, " ").trim();
  const full = `${clean} ${inn}`;
  if (full.length <= NAME_MAX) return full;
  const [last = "", ...rest] = clean.split(" ");
  const initials = rest.filter(Boolean).map((w) => w[0].toUpperCase() + ".").join("");
  const room = NAME_MAX - inn.length - 1;
  return `${`${last} ${initials}`.trim().slice(0, Math.max(room, 1))} ${inn}`;
}

const sha = (s) => crypto.createHash("sha256").update(String(s ?? "")).digest();
const safeEq = (a, b) => crypto.timingSafeEqual(sha(a), sha(b));
const escRe = (s) => String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
// Экранирование для Telegram (Markdown): иначе ФИО с "_" или "*" ломает всё сообщение
const mdEsc = (s) => String(s ?? "").replace(/([_*`\[])/g, "\\$1");

// Размер JPEG из заголовка (без внешних библиотек). null, если файл повреждён.
function jpegSize(b) {
  if (!b || b.length < 12 || b[0] !== 0xff || b[1] !== 0xd8) return null;
  let i = 2;
  while (i + 9 < b.length) {
    if (b[i] !== 0xff) { i++; continue; }
    const m = b[i + 1];
    if (m === 0xff) { i++; continue; }
    if (m >= 0xc0 && m <= 0xcf && m !== 0xc4 && m !== 0xc8 && m !== 0xcc)
      return { h: b.readUInt16BE(i + 5), w: b.readUInt16BE(i + 7) };
    i += 2 + b.readUInt16BE(i + 2);
  }
  return null;
}

// Проверка фото для терминала. Возвращает текст ошибки или null.
function checkPhoto(buf) {
  if (!buf || !buf.length) return "Загрузите фото";
  if (buf[0] !== 0xff || buf[1] !== 0xd8) return "Нужен JPEG";
  if (buf.length > MAX_PHOTO) return "Фото больше 200 КБ";
  const sz = jpegSize(buf);
  if (!sz) return "Файл повреждён или это не фото";
  if (Math.min(sz.w, sz.h) < MIN_SIDE) return `Фото слишком маленькое (${sz.w}×${sz.h}), нужно не меньше ${MIN_SIDE} px по короткой стороне`;
  return null;
}

// Состояние студента по самому свежему заданию
function stateOf(j) {
  if (!j) return { status: "none", results: [] };
  let status = j.status;
  if (j.type === "delete")
    status = j.status === "done" ? "removed" : j.status === "failed" || j.status === "partial" ? "remove_failed" : "removing";
  return {
    status,
    results: (j.results || []).map((r) => ({ ip: r.ip, ok: r.ok, error: r.error })),
    at: j.updatedAt || j.createdAt,
    jobId: String(j._id),
    retrying: j.status === "pending" && (j.attempts || 0) > 0,
  };
}

async function jobStates(studentIds) {
  const map = new Map();
  if (!studentIds.length) return map;
  const jobs = await SkudJob.find({ student: { $in: studentIds } }, "-photo").sort({ createdAt: -1, _id: -1 });
  for (const j of jobs) {
    const k = String(j.student);
    if (!map.has(k)) map.set(k, stateOf(j));
  }
  return map;
}

// Конец дня YYYY-MM-DD по времени Бишкека (UTC+6). null, если дата некорректна.
function endOfDayBishkek(ymd) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(ymd || ""))) return null;
  const d = new Date(`${ymd}T23:59:59+06:00`);
  return isNaN(d.getTime()) ? null : d;
}

module.exports = function makeCore(ctx = {}) {
  // req.ip учитывает trust proxy (Render); заголовок X-Forwarded-For напрямую подделывается клиентом
  const ipOf = (req) => req?.ip || req?.socket?.remoteAddress || "";

  const actorOf = (req) => {
    const s = (req && ctx.getSession && ctx.getSession(req)) || {};
    return { actor: s.username || "система", role: s.role || "" };
  };

  // Запись в журнал не должна ломать основное действие
  async function audit(req, action, target, details) {
    try {
      const { actor, role } = actorOf(req);
      await SkudAudit.create({
        actor, role, action,
        target: String(target ?? "").slice(0, 200),
        details: details ? String(details).slice(0, 500) : "",
        ip: ipOf(req),
      });
    } catch (e) {
      console.error("skud audit:", e.message);
    }
  }

  // Уведомление администраторам в Telegram (если бот подключён)
  async function notify(text) {
    try {
      if (ctx.notify) await ctx.notify(text);
    } catch (e) {
      console.error("skud notify:", e.message);
    }
  }

  return {
    SkudJob, SkudAudit, SkudDevice, SkudReport,
    buildDeviceName, stateOf, jobStates, checkPhoto, jpegSize, endOfDayBishkek,
    STATUS_RU, MAX_PHOTO, MIN_SIDE,
    safeEq, escRe, mdEsc, ipOf, actorOf, audit, notify,
  };
};
