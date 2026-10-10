// Еженедельный бэкап базы в Telegram администраторам: воскресенье 03:00 (Бишкек).
// Все коллекции → EJSON (сохраняет ObjectId и даты) → gzip; при BACKUP_PASSWORD — ещё и AES-256-GCM.
// Файлы фото/видео практики (GridFS *.chunks) не входят: они большие, а данные о них сохраняются.
// Восстановление: node scripts/restore-backup.js <файл> <MONGO_URI> [--drop]
// Отключить: BACKUP_CRON=off
const zlib = require("zlib");
const crypto = require("crypto");
const cron = require("node-cron");
const { BSON } = require("mongoose").mongo;
const { EJSON } = BSON;

const MAGIC = Buffer.from("PL3BKP1"); // признак зашифрованного файла
const TG_LIMIT = 49 * 1024 * 1024; // бот Telegram отправляет файлы до 50 МБ

const skip = (name) => name.endsWith(".chunks") || name.startsWith("system.");

// db — объект драйвера MongoDB (mongoose.connection.db)
async function dumpDatabase(db) {
  const cols = (await db.listCollections().toArray()).map((c) => c.name).filter((n) => !skip(n)).sort();
  const out = { createdAt: new Date().toISOString(), format: "pl3-backup-v1", collections: {} };
  const counts = {};
  for (const name of cols) {
    const docs = await db.collection(name).find({}).toArray();
    out.collections[name] = docs;
    counts[name] = docs.length;
  }
  const json = EJSON.stringify(out, { relaxed: false });
  return { buffer: zlib.gzipSync(Buffer.from(json, "utf8"), { level: 9 }), counts };
}

function encrypt(buf, password) {
  const salt = crypto.randomBytes(16), iv = crypto.randomBytes(12);
  const key = crypto.scryptSync(password, salt, 32);
  const c = crypto.createCipheriv("aes-256-gcm", key, iv);
  const data = Buffer.concat([c.update(buf), c.final()]);
  return Buffer.concat([MAGIC, salt, iv, c.getAuthTag(), data]);
}
function decrypt(buf, password) {
  if (!buf.subarray(0, MAGIC.length).equals(MAGIC)) return buf; // не зашифрован
  if (!password) throw new Error("Файл зашифрован: укажите пароль в BACKUP_PASSWORD");
  let o = MAGIC.length;
  const salt = buf.subarray(o, (o += 16)), iv = buf.subarray(o, (o += 12)), tag = buf.subarray(o, (o += 16));
  const d = crypto.createDecipheriv("aes-256-gcm", crypto.scryptSync(password, salt, 32), iv);
  d.setAuthTag(tag);
  return Buffer.concat([d.update(buf.subarray(o)), d.final()]);
}
// Файл бэкапа → { collections: { имя: [документы] } }
function readBackup(buf, password) {
  const plain = zlib.gunzipSync(decrypt(buf, password));
  const data = EJSON.parse(plain.toString("utf8"), { relaxed: false });
  if (data.format !== "pl3-backup-v1") throw new Error("Это не файл бэкапа ПЛ №3");
  return data;
}

function createBackup({ getDb, TgSub, getTodayKey }) {
  const fileName = () => `pl3-backup-${getTodayKey()}.json.gz${process.env.BACKUP_PASSWORD ? ".enc" : ""}`;

  async function make() {
    const db = getDb();
    if (!db) throw new Error("База не подключена");
    const { buffer, counts } = await dumpDatabase(db);
    const file = process.env.BACKUP_PASSWORD ? encrypt(buffer, process.env.BACKUP_PASSWORD) : buffer;
    return { file, counts, name: fileName() };
  }

  async function sendToAdmins(bot) {
    const b = await make();
    const admins = await TgSub.find({ role: "admin" }).lean();
    const total = Object.values(b.counts).reduce((a, n) => a + n, 0);
    const mb = (b.file.length / 1024 / 1024).toFixed(2);
    const caption =
      `💾 Бэкап базы ПЛ №3 · ${getTodayKey()}\n` +
      `Коллекций: ${Object.keys(b.counts).length}, записей: ${total}, размер: ${mb} МБ\n` +
      (process.env.BACKUP_PASSWORD ? "🔒 Зашифрован паролем BACKUP_PASSWORD" : "⚠️ Без шифрования — в файле личные данные студентов, не пересылайте его") +
      `\nВосстановление: node scripts/restore-backup.js <файл> <MONGO_URI>`;
    if (b.file.length > TG_LIMIT) {
      for (const a of admins) await bot.sendMessage(a.chatId, `💾 Бэкап ${mb} МБ — больше лимита Telegram (50 МБ). Скачайте его в админке: /api/admin/backup`).catch(() => null);
      return { sent: 0, tooBig: true, size: b.file.length, counts: b.counts };
    }
    let sent = 0;
    for (const a of admins) {
      await bot
        .sendDocument(a.chatId, b.file, { caption: caption.slice(0, 1000) }, { filename: b.name, contentType: "application/octet-stream" })
        .then(() => sent++)
        .catch((e) => console.error("backup → admin:", e.message));
    }
    console.log(`💾 Бэкап отправлен администраторам: ${sent} (${mb} МБ)`);
    return { sent, size: b.file.length, counts: b.counts };
  }

  function schedule(bot) {
    if (String(process.env.BACKUP_CRON || "").toLowerCase() === "off") return;
    cron.schedule("0 3 * * 0", () => sendToAdmins(bot).catch((e) => console.error("backup:", e.message)), { timezone: "Asia/Bishkek" });
    console.log(`💾 Бэкап базы администраторам: вс 03:00${process.env.BACKUP_PASSWORD ? ", с шифрованием" : ""}`);
  }

  function registerRoutes(app, { requireAdmin, getBot }) {
    app.get("/api/admin/backup", requireAdmin, async (_req, res) => {
      try {
        const b = await make();
        res.setHeader("Content-Type", "application/octet-stream");
        res.setHeader("Content-Disposition", `attachment; filename="${b.name}"`);
        res.send(b.file);
      } catch (e) {
        res.status(500).json({ error: e.message });
      }
    });
    app.post("/api/admin/backup/send-now", requireAdmin, async (_req, res) => {
      try {
        const bot = getBot();
        if (!bot) return res.status(503).json({ error: "Бот не запущен" });
        res.json(await sendToAdmins(bot));
      } catch (e) {
        res.status(500).json({ error: e.message });
      }
    });
  }

  return { make, sendToAdmins, schedule, registerRoutes };
}

module.exports = { createBackup, dumpDatabase, readBackup, encrypt, decrypt, BSON };
