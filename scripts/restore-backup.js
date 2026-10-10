// Восстановление бэкапа ПЛ №3 в базу MongoDB.
//   node scripts/restore-backup.js pl3-backup-2026-10-11.json.gz "mongodb+srv://..." [--drop]
// Без --drop восстанавливает только в пустые коллекции (чтобы случайно не затереть рабочую базу).
// Для зашифрованного файла (.enc) задайте пароль: BACKUP_PASSWORD=... node scripts/restore-backup.js ...
const fs = require("fs");
const { MongoClient } = require("mongoose").mongo;
const { readBackup } = require("../backup");

(async () => {
  const [file, uri] = process.argv.slice(2).filter((a) => !a.startsWith("--"));
  const drop = process.argv.includes("--drop");
  if (!file || !uri) {
    console.log("Использование: node scripts/restore-backup.js <файл бэкапа> <MONGO_URI> [--drop]");
    process.exit(1);
  }
  const data = readBackup(fs.readFileSync(file), process.env.BACKUP_PASSWORD);
  console.log(`Бэкап от ${data.createdAt}, коллекций: ${Object.keys(data.collections).length}`);
  const client = await MongoClient.connect(uri);
  const db = client.db();
  for (const [name, docs] of Object.entries(data.collections)) {
    const col = db.collection(name);
    const existing = await col.estimatedDocumentCount();
    if (existing && !drop) {
      console.log(`  — ${name}: пропущено, в базе уже ${existing} записей (добавьте --drop, чтобы заменить)`);
      continue;
    }
    if (existing) await col.deleteMany({});
    if (docs.length) await col.insertMany(docs, { ordered: false });
    console.log(`  ✓ ${name}: ${docs.length}`);
  }
  await client.close();
  console.log("Готово. Перезапустите сервер, чтобы он пересоздал индексы.");
})().catch((e) => {
  console.error("Ошибка:", e.message);
  process.exit(1);
});
