// Тесты «не пришёл» (absent-alerts.js) и бэкапа (backup.js). БД — в памяти.  Запуск: npm test
const Module = require("module");
const path = require("path");
const assert = require("assert/strict");

const fake = require("./fake-mongoose");
const origResolve = Module._resolveFilename;
Module._resolveFilename = function (req, ...rest) {
  // mongoose.mongo (EJSON, MongoClient) берём из настоящего пакета, модели — из имитации
  return req === "mongoose" && !rest[0]?.filename?.endsWith("backup.js") && !rest[0]?.filename?.endsWith("restore-backup.js")
    ? path.join(__dirname, "fake-mongoose.js")
    : origResolve.call(this, req, ...rest);
};
const { Schema } = fake;
const Group = fake.model("Group", new Schema({ name: String, shiftStart: { type: String, default: "08:30" }, kind: { type: String, default: "students" }, saturday: { type: Boolean, default: false } }));
const Student = fake.model("Student", new Schema({ fio: String, inn: String, group: Schema.Types.ObjectId, attendance: [], skudDaily: [] }));
const TgSub = fake.model("TgSub", new Schema({ chatId: String, role: String, linkedInns: [], curatorId: String }));
const Curator = fake.model("Curator", new Schema({ username: String, groups: [] }));
const absence = require("../absence");
const { createAbsentAlerts, AlertLog } = require("../absent-alerts");
const backupMod = require("../backup");

const TODAY = "2026-10-14"; // среда
const at = (dk, hhmm) => new Date(`${dk}T${hhmm}:00+06:00`).toISOString();
const came = (hhmm = "08:20") => [{ date: TODAY, present: true, firstIn: at(TODAY, hhmm), inCount: 1 }];

let passed = 0,
  failed = 0;
async function t(name, fn) {
  try {
    await fn();
    passed++;
    console.log("  ✓", name);
  } catch (e) {
    failed++;
    console.log("  ✗", name, "\n    ", e.message);
  }
}

(async () => {
  const gA = await Group.create({ name: "ИТ-21", shiftStart: "08:30" });
  const gB = await Group.create({ name: "СВ-11", shiftStart: "13:00" }); // вторая смена
  const gC = await Group.create({ name: "ПК-31", shiftStart: "08:30" }); // сегодня никто не пришёл
  await Group.create({ name: "Учителя", kind: "teachers" });
  await Student.create({ fio: "Асанов Бакыт", inn: "111", group: gA._id, skudDaily: came() });
  const bek = await Student.create({ fio: "Бекова Айгуль", inn: "222", group: gA._id, skudDaily: [] });
  const sick = await Student.create({ fio: "Болот Эмир", inn: "333", group: gA._id, skudDaily: [] });
  const prac = await Student.create({ fio: "Практикант Нурбек", inn: "444", group: gA._id, skudDaily: [] });
  await Student.create({ fio: "Вторая Смена", inn: "555", group: gB._id, skudDaily: [] });
  await Student.create({ fio: "Пустая Группа", inn: "666", group: gC._id, skudDaily: [] });
  await absence.createReason({ student: sick, body: { type: "sick", dateFrom: TODAY }, source: "curator" });
  const cur = await Curator.create({ username: "cur", groups: [String(gA._id)] });
  await TgSub.create({ chatId: "P2", role: "parent", linkedInns: ["222"] });
  await TgSub.create({ chatId: "P3", role: "parent", linkedInns: ["333"] });
  await TgSub.create({ chatId: "P4", role: "parent", linkedInns: ["444"] });
  await TgSub.create({ chatId: "C1", role: "curator", curatorId: String(cur._id) });

  let nowMin = 9 * 60 + 40;
  const alerts = createAbsentAlerts({ Group, Student, Curator, TgSub, getTodayKey: () => TODAY, getBishkekMinutes: () => nowMin, practiceStudentIds: async () => [prac._id] });
  const sent = [];
  const bot = { sendMessage: async (chatId, text) => sent.push({ chatId, text }) };

  console.log("«Не пришёл»:");
  await t("до часа после начала занятий — тишина", async () => {
    nowMin = 9 * 60 + 20;
    const r = await alerts.run(bot);
    assert.equal(r.groups, 0);
    assert.equal(sent.length, 0);
  });
  await t("через час: родителю отсутствующего и куратору — список", async () => {
    nowMin = 9 * 60 + 40;
    const r = await alerts.run(bot);
    assert.equal(r.groups, 1, JSON.stringify(r));
    assert.deepEqual(sent.map((m) => m.chatId).sort(), ["C1", "P2"]);
    assert.match(sent.find((m) => m.chatId === "P2").text, /Бекова Айгуль.*не отмечен/);
    assert.match(sent.find((m) => m.chatId === "C1").text, /1\. Бекова Айгуль/);
  });
  await t("не тревожим: есть причина, на практике, вторая смена ещё не началась", async () => {
    assert.ok(!sent.some((m) => ["P3", "P4"].includes(m.chatId)));
    assert.ok(!sent.some((m) => /Вторая Смена/.test(m.text)));
  });
  await t("группа, где не пришёл никто, пропускается (праздник/выезд/сбой)", async () => {
    const r = await alerts.run(bot, { dry: true, force: true });
    assert.ok(r.skipped.some((x) => x.group === "ПК-31"));
    assert.ok(!sent.some((m) => /Пустая Группа/.test(m.text)));
  });
  await t("повторно за день не отправляем", async () => {
    sent.length = 0;
    await alerts.run(bot);
    assert.equal(sent.length, 0);
    assert.equal(AlertLog._store.filter((x) => x.date === TODAY).length, 1);
  });
  await t("предпросмотр показывает имена и ничего не отправляет", async () => {
    const r = await alerts.run(bot, { dry: true, force: true });
    assert.deepEqual(r.preview.find((p) => p.group === "ИТ-21").absent, ["Бекова Айгуль"]);
    assert.equal(sent.length, 0);
  });
  await t("в выходной и когда в лицее нет ни одного входа — молчим", async () => {
    const sat = createAbsentAlerts({ Group, Student, Curator, TgSub, getTodayKey: () => "2026-10-17", getBishkekMinutes: () => 600, practiceStudentIds: async () => [] });
    assert.equal((await sat.run(bot, { dry: true, force: true })).groups, 0);
    const empty = createAbsentAlerts({ Group, Student, Curator, TgSub, getTodayKey: () => "2026-10-15", getBishkekMinutes: () => 600, practiceStudentIds: async () => [] });
    const r = await empty.run(bot, { dry: true, force: true });
    assert.equal(r.groups, 0);
    assert.ok(r.skipped.every((x) => /нет ни одного входа|никто не пришёл/.test(x.why)));
  });

  console.log("Бэкап:");
  const { ObjectId } = backupMod.BSON;
  const oid = new ObjectId();
  const fakeDb = {
    listCollections: () => ({ toArray: async () => [{ name: "students" }, { name: "groups" }, { name: "practiceMedia.chunks" }] }),
    collection: (n) => ({
      find: () => ({
        toArray: async () =>
          n === "students" ? [{ _id: oid, fio: "Асанов Бакыт", inn: "111", createdAt: new Date("2026-10-01T00:00:00Z") }] : n === "groups" ? [{ _id: new ObjectId(), name: "ИТ-21" }] : [{ data: "big" }],
      }),
    }),
  };
  await t("дамп сохраняет типы (ObjectId, даты) и пропускает файлы GridFS", async () => {
    const { buffer, counts } = await backupMod.dumpDatabase(fakeDb);
    assert.deepEqual(counts, { groups: 1, students: 1 });
    const back = backupMod.readBackup(buffer);
    const s = back.collections.students[0];
    assert.ok(s._id instanceof ObjectId && s._id.equals(oid));
    assert.ok(s.createdAt instanceof Date);
    assert.equal(s.fio, "Асанов Бакыт");
  });
  await t("шифрование паролем: без пароля и с чужим паролем не читается", async () => {
    const { buffer } = await backupMod.dumpDatabase(fakeDb);
    const enc = backupMod.encrypt(buffer, "секрет-123");
    assert.ok(!enc.includes(Buffer.from("Асанов")));
    assert.equal(backupMod.readBackup(enc, "секрет-123").collections.students[0].fio, "Асанов Бакыт");
    assert.throws(() => backupMod.readBackup(enc), /пароль/);
    assert.throws(() => backupMod.readBackup(enc, "не тот"));
  });
  await t("отправка администраторам в Telegram", async () => {
    await TgSub.create({ chatId: "A1", role: "admin" });
    const docs = [];
    const b = backupMod.createBackup({ getDb: () => fakeDb, TgSub, getTodayKey: () => TODAY });
    process.env.BACKUP_PASSWORD = "секрет-123";
    const r = await b.sendToAdmins({ sendDocument: async (chatId, buf, opts, file) => docs.push({ chatId, buf, file, caption: opts.caption }), sendMessage: async () => {} });
    delete process.env.BACKUP_PASSWORD;
    assert.equal(r.sent, 1);
    assert.equal(docs[0].chatId, "A1");
    assert.equal(docs[0].file.filename, `pl3-backup-${TODAY}.json.gz.enc`);
    assert.match(docs[0].caption, /Зашифрован/);
    assert.equal(backupMod.readBackup(docs[0].buf, "секрет-123").collections.groups[0].name, "ИТ-21");
  });

  console.log(`\nИтог: ${passed} прошло, ${failed} упало`);
  process.exit(failed ? 1 : 0);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
