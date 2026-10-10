// Тесты модулей СКУД без MongoDB и без настоящих терминалов:
// БД имитируется в памяти, терминал — mock-hik.js (с Digest-авторизацией).  Запуск: npm test
const Module = require("module");
const path = require("path");
const fs = require("fs");
const assert = require("assert/strict");

const fake = require("./fake-mongoose");
const origResolve = Module._resolveFilename;
Module._resolveFilename = function (req, ...rest) {
  return req === "mongoose" ? path.join(__dirname, "fake-mongoose.js") : origResolve.call(this, req, ...rest);
};

const express = require("express");
const startMock = require("./mock-hik");
const agentSkud = require("../agent-skud");

const FACE = fs.readFileSync(path.join(__dirname, "fixtures/face.jpg"));
const SMALL = fs.readFileSync(path.join(__dirname, "fixtures/small.jpg"));
const KEY = "test-agent-key";

const { Schema } = fake;
const Group = fake.model("Group", new Schema({ name: { type: String, unique: true }, shiftStart: { type: String, default: "08:00" }, kind: { type: String, enum: ["students", "teachers", "masters", "guests"], default: "students" } }));
const Student = fake.model("Student", new Schema({ fio: String, inn: String, lyceumId: String, skudConsentAt: Date, skudConsentBy: String, group: Schema.Types.ObjectId }));

const sessionOf = (req) => { try { return JSON.parse(req.headers["x-test-session"]); } catch { return null; } };
const requireAuth = (req, res, next) => (sessionOf(req)?.role ? next() : res.status(401).json({ error: "auth" }));
const requireAdmin = (req, res, next) => (sessionOf(req)?.role === "admin" ? next() : res.status(403).json({ error: "Admin only" }));
const ids = (req) => { const s = sessionOf(req); return !s?.role ? [] : s.role === "admin" ? null : s.groupIds || []; };
const canAccessGroup = (req, g) => { const i = ids(req); return i === null || i.map(String).includes(String(g)); };
const normalizeINN = (s) => { const d = String(s || "").replace(/\D/g, ""); return d.length >= 12 ? d : null; };

const notes = [];
const app = express();
app.set("trust proxy", 1);
app.use(express.json({ limit: "2mb" }));
app.use("/api/admin", requireAuth);
const skud = require("../skud")(app, {
  Student, Group, requireAuth, requireAdmin, normalizeINN, AGENT_API_KEY: KEY,
  getSession: sessionOf, canAccessGroup, getAccessibleGroupIds: ids, notify: async (m) => { notes.push(m); },
});
const { SkudJob, SkudDevice, SkudReport, SkudAudit } = skud.core;

const ADMIN = { role: "admin", username: "boss" };
let base, gA, gB, curA, curB;

async function api(who, method, url, opt = {}) {
  const headers = {};
  if (who) headers["x-test-session"] = JSON.stringify(who);
  let body;
  if (opt.json !== undefined && method !== "GET") { headers["content-type"] = "application/json"; body = JSON.stringify(opt.json); }
  if (opt.form) body = opt.form;
  const r = await fetch(base + url, { method, headers, body });
  const ct = r.headers.get("content-type") || "";
  const data = ct.includes("json") ? await r.json() : Buffer.from(await r.arrayBuffer());
  return { status: r.status, data, headers: r.headers };
}
const photoForm = (buf = FACE, fields = {}) => {
  const f = new FormData();
  f.append("photo", new Blob([buf], { type: "image/jpeg" }), "p.jpg");
  for (const [k, v] of Object.entries(fields)) f.append(k, v);
  return f;
};
const agentPost = (p, body, key = KEY) => api(null, "POST", p, { json: { apiKey: key, ...body } });

let passed = 0, failed = 0;
const t = async (name, fn) => {
  try { await fn(); passed++; console.log("  ✓", name); }
  catch (e) { failed++; console.log("  ✗", name, "\n    ", e.stack.split("\n").slice(0, 4).join("\n     ")); }
};

(async () => {
  const srv = app.listen(0); await new Promise((r) => srv.once("listening", r));
  base = `http://127.0.0.1:${srv.address().port}`;
  const hik = await startMock();
  const hik2 = await startMock();
  const ag = agentSkud.create({ serverUrl: base, apiKey: KEY, devices: [{ ip: hik.ip }, { ip: hik2.ip }], user: "admin", pass: "secret" });

  gA = await Group.create({ name: "ЭС-21" }); gB = await Group.create({ name: "СВ-22" });
  curA = { role: "curator", username: "curA", groupIds: [String(gA._id)] };
  curB = { role: "curator", username: "curB", groupIds: [String(gB._id)] };
  const s1 = await Student.create({ fio: "Иванов Иван Иванович", inn: "12345678901234", lyceumId: "2024001", group: gA._id });
  const s2 = await Student.create({ fio: "Петрова Анна Сергеевна", inn: "22345678901234", lyceumId: "2024002", group: gA._id });
  const s3 = await Student.create({ fio: "Сидоров Пётр", inn: "", group: gA._id });
  const sB = await Student.create({ fio: "Чужой Студент Группы", inn: "32345678901234", lyceumId: "2024100", group: gB._id });
  const drain = async () => { for (let i = 0; i < 10; i++) await ag.tick(); };

  console.log("Доступ");
  await t("без сессии API закрыт, агентский ключ обязателен", async () => {
    assert.equal((await api(null, "GET", "/api/admin/skud-reg/groups")).status, 401);
    assert.equal((await agentPost("/api/agent/jobs/claim", {}, "bad")).status, 403);
    assert.equal((await agentPost("/api/agent/heartbeat", {}, "")).status, 403);
  });
  await t("куратор видит только свои группы и студентов", async () => {
    const g = await api(curA, "GET", "/api/admin/skud-reg/groups");
    assert.deepEqual(g.data.map((x) => x.name), ["ЭС-21"]);
    assert.equal((await api(curA, "GET", `/api/admin/skud-reg/students?groupId=${gB._id}`)).status, 403);
    const own = await api(curA, "GET", `/api/admin/skud-reg/students?groupId=${gA._id}`);
    assert.equal(own.data.length, 3);
    const search = await api(curA, "GET", "/api/admin/skud-reg/students?q=Чужой");
    assert.equal(search.data.length, 0);
    assert.equal((await api(ADMIN, "GET", "/api/admin/skud-reg/students?q=Чужой")).data.length, 1);
  });
  await t("куратор не может админские действия", async () => {
    for (const [m, u] of [["GET", "/api/admin/skud-reg/audit"], ["GET", "/api/admin/skud-reg/jobs"], ["POST", `/api/admin/skud-reg/${s1._id}/remove`], ["GET", "/api/admin/skud-reg/devices"], ["POST", "/api/admin/skud-reg/reconcile/run"], ["POST", "/api/admin/staff-links"], ["GET", "/api/admin/staff-submissions?scope=staff"]])
      assert.equal((await api(curA, m, u, { json: {} })).status, 403, m + u);
  });

  console.log("Регистрация фото");
  await t("нужно согласие, JPEG ≥300px, ИНН и номер", async () => {
    let r = await api(curA, "POST", `/api/admin/skud-reg/${s1._id}`, { form: photoForm() });
    assert.equal(r.status, 400); assert.match(r.data.error, /согласие/);
    r = await api(curA, "POST", `/api/admin/skud-reg/${s1._id}`, { form: photoForm(SMALL, { consent: "1" }) });
    assert.equal(r.status, 400); assert.match(r.data.error, /маленькое/);
    r = await api(curA, "POST", `/api/admin/skud-reg/${s1._id}`, { form: photoForm(Buffer.from("not a jpeg at all"), { consent: "1" }) });
    assert.equal(r.status, 400);
    r = await api(curA, "POST", `/api/admin/skud-reg/${s3._id}`, { form: photoForm(FACE, { consent: "1" }) });
    assert.equal(r.status, 400); assert.match(r.data.error, /ИНН/);
    r = await api(curB, "POST", `/api/admin/skud-reg/${s1._id}`, { form: photoForm(FACE, { consent: "1" }) });
    assert.equal(r.status, 403);
    assert.equal(await SkudJob.countDocuments({}), 0);
  });
  await t("успешная отправка: терминал получает имя ФИО+ИНН, номер, срок; фото стирается", async () => {
    const r = await api(curA, "POST", `/api/admin/skud-reg/${s1._id}`, { form: photoForm(FACE, { consent: "1" }) });
    assert.equal(r.status, 200, JSON.stringify(r.data));
    assert.ok(r.data.deviceName.endsWith("12345678901234"));
    assert.ok(r.data.deviceName.length <= 32);
    await drain();
    for (const h of [hik, hik2]) {
      const u = h.users.get("2024001");
      assert.ok(u, "пользователь на терминале");
      assert.equal(u.name, r.data.deviceName);
      assert.equal(u.Valid.endTime, "2035-12-31T23:59:59");
      assert.ok(h.faces.get("2024001"));
    }
    const job = await SkudJob.findOne({ employeeNo: "2024001" });
    assert.equal(job.status, "done"); assert.equal(job.photo, undefined);
    const st = (await api(curA, "GET", `/api/admin/skud-reg/students?groupId=${gA._id}`)).data.find((x) => x.id === String(s1._id));
    assert.equal(st.status, "done"); assert.equal(st.consent, true);
  });
  await t("повторная отправка заменяет фото (employeeNo уже есть)", async () => {
    const r = await api(curA, "POST", `/api/admin/skud-reg/${s1._id}`, { form: photoForm() }); // согласие уже сохранено
    assert.equal(r.status, 200);
    await drain();
    assert.equal((await SkudJob.findOne({ _id: r.data.jobId })).status, "done");
    assert.ok(hik.log.includes("PUT Intelligent/FDLib/FDModify"));
  });
  await t("срок доступа переводится в UTC+6", async () => {
    const day = new Date(Date.now() + 5 * 864e5).toISOString().slice(0, 10);
    const r = await api(ADMIN, "POST", `/api/admin/skud-reg/${s2._id}`, { form: photoForm(FACE, { consent: "1", validUntil: day }) });
    assert.equal(r.status, 200);
    await drain();
    assert.equal(hik.users.get("2024002").Valid.endTime, `${day}T23:59:59`);
    const bad = await api(ADMIN, "POST", `/api/admin/skud-reg/${s2._id}`, { form: photoForm(FACE, { validUntil: "2020-01-01" }) });
    assert.equal(bad.status, 400);
  });
  await t("дубликат поименного номера у другого студента отклоняется", async () => {
    const dup = await Student.create({ fio: "Двойник Номера Иванович", inn: "42345678901234", lyceumId: "2024001", group: gA._id });
    const r = await api(ADMIN, "POST", `/api/admin/skud-reg/${dup._id}`, { form: photoForm(FACE, { consent: "1" }) });
    assert.equal(r.status, 409);
    await Student.updateOne({ _id: dup._id }, { lyceumId: "2024777" });
    await Student.updateOne({ _id: dup._id }, { group: gB._id });
  });

  console.log("Сбои и повторы");
  await t("ошибка распознавания лица: автоповторы, затем failed и Telegram; ручной повтор", async () => {
    const s = await Student.create({ fio: "Плохое Фото Тестович", inn: "52345678901234", lyceumId: "2024050", group: gA._id });
    hik.failFace = true; notes.length = 0;
    const r = await api(ADMIN, "POST", `/api/admin/skud-reg/${s._id}`, { form: photoForm(FACE, { consent: "1" }) });
    for (let i = 0; i < 3; i++) {
      await drain();
      await SkudJob.updateMany({ status: "pending" }, { retryAt: new Date(Date.now() - 1000) });
    }
    const job = await SkudJob.findOne({ _id: r.data.jobId });
    hik.failFace = false;
    assert.equal(job.status, "partial", job.status); // на одном терминале записан, на другом нет
    assert.equal(job.attempts, 3);
    assert.ok(job.results.some((x) => /лицо не распознано/.test(x.error)));
    assert.ok(notes.some((n) => /не записан/.test(n)));
    const rr = await api(ADMIN, "POST", `/api/admin/skud-reg/jobs/${job._id}/retry`);
    assert.equal(rr.status, 200);
    await drain();
    assert.equal((await SkudJob.findOne({ _id: job._id })).status, "done");
  });
  await t("один терминал не отвечает: partial, затем done после восстановления", async () => {
    const s = await Student.create({ fio: "Частичный Успех Иванович", inn: "62345678901234", lyceumId: "2024060", group: gA._id });
    hik2.down = true;
    const r = await api(ADMIN, "POST", `/api/admin/skud-reg/${s._id}`, { form: photoForm(FACE, { consent: "1" }) });
    await ag.tick();
    let job = await SkudJob.findOne({ _id: r.data.jobId });
    assert.equal(job.status, "pending"); assert.ok(job.retryAt);
    hik2.down = false;
    await SkudJob.updateOne({ _id: job._id }, { retryAt: new Date(Date.now() - 1) });
    await drain();
    job = await SkudJob.findOne({ _id: r.data.jobId });
    assert.equal(job.status, "done");
    assert.ok(hik2.users.has("2024060"));
  });
  await t("зависшее задание (аренда истекла) берётся заново", async () => {
    const j = await SkudJob.create({ type: "register", fio: "x", employeeNo: "999", deviceName: "x", photo: FACE, status: "processing", attempts: 1, leaseUntil: new Date(Date.now() - 1000) });
    const r = await agentPost("/api/agent/jobs/claim", {});
    assert.equal(String(r.data.job.id), String(j._id));
    await SkudJob.deleteMany({ employeeNo: "999" });
  });

  console.log("Удаление");
  await t("удаление студента из базы убирает его с терминалов; чужих не трогает", async () => {
    assert.ok(hik.users.has("2024002"));
    const no = await skud.enqueueDeleteForStudent(await Student.findById(s3._id), { headers: {} });
    assert.equal(no, null); // у s3 нет номера
    const job = await skud.enqueueDeleteForStudent(await Student.findById(s2._id), { headers: {} });
    assert.ok(job);
    await drain();
    assert.ok(!hik.users.has("2024002") && !hik2.users.has("2024002"));
    assert.equal((await SkudJob.findOne({ _id: job._id })).status, "done");
    // повторное удаление не ставится: последнее задание уже delete
    assert.equal(await skud.enqueueDeleteForStudent(await Student.findById(s2._id), { headers: {} }), null);
    // студент, которого система на терминал не ставила: автоудаления нет
    const manual = await Student.create({ fio: "Вручную Добавленный Иванович", inn: "72345678901234", lyceumId: "5555", group: gA._id });
    assert.equal(await skud.enqueueDeleteForStudent(manual, { headers: {} }), null);
  });
  await t("кнопка «с терминалов» (админ) и статус removed", async () => {
    const r = await api(ADMIN, "POST", `/api/admin/skud-reg/${s1._id}/remove`);
    assert.equal(r.status, 200);
    await drain();
    assert.ok(!hik.users.has("2024001"));
    const st = (await api(ADMIN, "GET", `/api/admin/skud-reg/students?groupId=${gA._id}`)).data.find((x) => x.id === String(s1._id));
    assert.equal(st.status, "removed");
  });
  await t("удаление несуществующего на терминале считается успехом", async () => {
    const ok = await ag.remove(hik.ip, { employeeNo: "000000" });
    assert.equal(ok, undefined);
  });

  console.log("Сотрудники и гости");
  let guestLink, sub;
  await t("ссылка для гостей: срок пропуска 30 дней по умолчанию", async () => {
    const r = await api(ADMIN, "POST", "/api/admin/staff-links", { json: { kind: "guests", hours: 24, maxUses: 2 } });
    assert.equal(r.status, 200); assert.equal(r.data.accessDays, 30);
    guestLink = r.data;
    const qr = await api(ADMIN, "GET", `/api/admin/staff-links/${r.data.id}/qr`);
    assert.equal(qr.status, 200); assert.equal(qr.data[1], 0x50); // PNG
  });
  await t("публичная анкета: проверки и отправка", async () => {
    const tok = guestLink.url.split("/join/")[1];
    assert.equal((await api(null, "GET", `/api/join/${tok}/info`)).data.mode, "staff");
    assert.equal((await api(null, "GET", "/join/" + tok)).status, 200);
    assert.equal((await api(null, "GET", "/join/" + "x".repeat(20))).status, 410);
    const f = (o) => photoForm(FACE, o);
    let r = await api(null, "POST", `/api/join/${tok}`, { form: f({ fio: "Гостев Гость Гостевич", inn: "11111111111111" }) });
    assert.equal(r.status, 400); assert.match(r.data.error, /согласие/);
    r = await api(null, "POST", `/api/join/${tok}`, { form: f({ fio: "Гостев", inn: "11111111111111", consent: "1" }) });
    assert.equal(r.status, 400);
    r = await api(null, "POST", `/api/join/${tok}`, { form: f({ fio: "Гостев Гость Гостевич", inn: "123", consent: "1" }) });
    assert.equal(r.status, 400);
    r = await api(null, "POST", `/api/join/${tok}`, { form: photoForm(SMALL, { fio: "Гостев Гость Гостевич", inn: "11111111111111", consent: "1" }) });
    assert.equal(r.status, 400);
    notes.length = 0;
    r = await api(null, "POST", `/api/join/${tok}`, { form: f({ fio: "Гостев Гость Гостевич", inn: "11111111111111", consent: "1" }) });
    assert.equal(r.status, 200, JSON.stringify(r.data));
    assert.ok(notes.some((n) => /новая анкета/.test(n)));
    r = await api(null, "POST", `/api/join/${tok}`, { form: f({ fio: "Другой Человек Иванович", inn: "11111111111111", consent: "1" }) });
    assert.equal(r.status, 409); // тот же ИНН уже подан
  });
  await t("модерация: только админ; одобрение создаёт группу «Гости», номер 9…, пропуск 30 дней", async () => {
    assert.equal((await api(curA, "GET", "/api/admin/staff-submissions")).status, 403);
    const list = await api(ADMIN, "GET", "/api/admin/staff-submissions");
    assert.equal(list.data.length, 1); sub = list.data[0];
    assert.equal((await api(ADMIN, "GET", `/api/admin/staff-submissions/${sub.id}/photo`)).status, 200);
    assert.equal((await api(curA, "POST", `/api/admin/staff-submissions/${sub.id}/approve`)).status, 403);
    const r = await api(ADMIN, "POST", `/api/admin/staff-submissions/${sub.id}/approve`);
    assert.equal(r.status, 200, JSON.stringify(r.data));
    assert.match(r.data.employeeNo, /^9000\d{3}$/);
    assert.equal((await Group.findOne({ name: "Гости" })).kind, "guests");
    assert.equal((await api(ADMIN, "POST", `/api/admin/staff-submissions/${sub.id}/approve`)).status, 404); // уже обработана
    await drain();
    const u = hik.users.get(r.data.employeeNo);
    assert.ok(u && hik.faces.get(r.data.employeeNo));
    const end = new Date(u.Valid.endTime + "+06:00").getTime();
    const days = (end - Date.now()) / 864e5;
    assert.ok(days > 29.5 && days < 30.5, "дней: " + days);
    assert.equal((await Student.findOne({ lyceumId: r.data.employeeNo })).skudConsentBy, "анкета по ссылке");
  });
  await t("учителя и мастера получают 7… и 8…, ссылка отзывается", async () => {
    const mk = await api(ADMIN, "POST", "/api/admin/staff-links", { json: { kind: "teachers", hours: 1, maxUses: 5 } });
    assert.equal(mk.data.accessDays, null);
    const tok = mk.data.url.split("/join/")[1];
    await api(null, "POST", `/api/join/${tok}`, { form: photoForm(FACE, { fio: "Учителев Учитель Учителевич", inn: "21111111111111", consent: "1" }) });
    const s = (await api(ADMIN, "GET", "/api/admin/staff-submissions")).data[0];
    const r = await api(ADMIN, "POST", `/api/admin/staff-submissions/${s.id}/approve`);
    assert.match(r.data.employeeNo, /^7000\d{3}$/);
    await drain();
    assert.equal(hik.users.get(r.data.employeeNo).Valid.endTime, "2035-12-31T23:59:59");
    await api(ADMIN, "DELETE", `/api/admin/staff-links/${mk.data.id}`);
    assert.equal((await api(null, "GET", `/api/join/${tok}/info`)).status, 410);
  });
  await t("отклонение удаляет фото; лимит заявок по ссылке", async () => {
    const mk = await api(ADMIN, "POST", "/api/admin/staff-links", { json: { kind: "masters", hours: 1, maxUses: 1 } });
    const tok = mk.data.url.split("/join/")[1];
    assert.equal((await api(null, "POST", `/api/join/${tok}`, { form: photoForm(FACE, { fio: "Мастеров Мастер Мастерович", inn: "31111111111111", consent: "1" }) })).status, 200);
    assert.equal((await api(null, "POST", `/api/join/${tok}`, { form: photoForm(FACE, { fio: "Второй Мастер Иванович", inn: "41111111111111", consent: "1" }) })).status, 410);
    const s = (await api(ADMIN, "GET", "/api/admin/staff-submissions")).data[0];
    assert.equal((await api(ADMIN, "POST", `/api/admin/staff-submissions/${s.id}/reject`)).status, 200);
    assert.equal((await api(ADMIN, "GET", `/api/admin/staff-submissions/${s.id}/photo`)).status, 404);
  });

  console.log("Самостоятельная съёмка студента");
  let selfTok;
  await t("куратор создаёт QR-ссылку только для своих студентов", async () => {
    assert.equal((await api(curB, "POST", `/api/admin/skud-selflink/${s1._id}`, { json: {} })).status, 403);
    assert.equal((await api(curA, "POST", `/api/admin/skud-selflink/${s3._id}`, { json: {} })).status, 400); // нет ИНН
    const r = await api(curA, "POST", `/api/admin/skud-selflink/${s1._id}`, { json: {} });
    assert.equal(r.status, 200); assert.match(r.data.qr, /^data:image\/png/);
    selfTok = r.data.url.split("/join/")[1];
    const info = (await api(null, "GET", `/api/join/${selfTok}/info`)).data;
    assert.equal(info.mode, "student"); assert.ok(!/12345678901234/.test(JSON.stringify(info)));
  });
  await t("студент отправляет фото; куратор принимает; согласие сохраняется", async () => {
    let r = await api(null, "POST", `/api/join/${selfTok}`, { form: photoForm(FACE, {}) });
    assert.equal(r.status, 400);
    r = await api(null, "POST", `/api/join/${selfTok}`, { form: photoForm(FACE, { consent: "1" }) });
    assert.equal(r.status, 200);
    assert.equal((await api(curA, "GET", "/api/admin/staff-submissions?scope=student")).data.length, 1);
    assert.equal((await api(curB, "GET", "/api/admin/staff-submissions?scope=student")).data.length, 0);
    const sub = (await api(curA, "GET", "/api/admin/staff-submissions?scope=student")).data[0];
    assert.equal((await api(curB, "POST", `/api/admin/staff-submissions/${sub.id}/approve`)).status, 403);
    assert.equal((await api(curB, "GET", `/api/admin/staff-submissions/${sub.id}/photo`)).status, 404);
    assert.equal((await api(curA, "POST", `/api/admin/staff-submissions/${sub.id}/approve`)).status, 200);
    await drain();
    assert.ok(hik.users.has("2024001"));
    assert.ok((await Student.findById(s1._id)).skudConsentAt);
  });

  console.log("Терминалы и сверка");
  await t("heartbeat: online/offline, уведомление об отключении ≥3 мин и о возврате", async () => {
    notes.length = 0;
    await ag.heartbeat();
    let dv = (await api(ADMIN, "GET", "/api/admin/skud-reg/devices")).data;
    assert.equal(dv.length, 2); assert.ok(dv.every((d) => d.state === "online"));
    assert.ok(dv.find((d) => d.ip === hik.ip).users >= 3);
    hik2.down = true;
    await ag.heartbeat();
    assert.equal(notes.filter((n) => /не отвечает/.test(n)).length, 0); // ещё рано
    await SkudDevice.updateOne({ ip: hik2.ip }, { offlineSince: new Date(Date.now() - 4 * 60e3) });
    await ag.heartbeat();
    assert.equal(notes.filter((n) => /не отвечает/.test(n)).length, 1);
    await ag.heartbeat();
    assert.equal(notes.filter((n) => /не отвечает/.test(n)).length, 1); // без повторов
    dv = (await api(ADMIN, "GET", "/api/admin/skud-reg/devices")).data;
    assert.equal(dv.find((d) => d.ip === hik2.ip).state, "offline");
    hik2.down = false;
    await ag.heartbeat();
    assert.ok(notes.some((n) => /снова в сети/.test(n)));
  });
  await t("молчание агента >5 мин даёт одно уведомление", async () => {
    notes.length = 0;
    await SkudDevice.updateMany({}, { lastSeen: new Date(Date.now() - 6 * 60e3) });
    await skud.checkSilence();
    await skud.checkSilence();
    assert.equal(notes.filter((n) => /агент/.test(n)).length, 1);
    assert.ok((await api(ADMIN, "GET", "/api/admin/skud-reg/devices")).data.every((d) => d.state === "silent"));
    await ag.heartbeat();
  });
  await t("сверка находит лишних и недостающих; лишнего можно удалить", async () => {
    hik.users.set("424242", { employeeNo: "424242", name: "Левый Человек" });
    hik2.users.delete("2024001"); hik2.faces.delete("2024001"); // пропал на втором терминале
    for (let i = 0; i < 5; i++) hik.users.set("7777" + i, { employeeNo: "7777" + i, name: "Лишний " + i }); // проверка постраничного чтения (по 2 записи)
    notes.length = 0;
    await api(ADMIN, "POST", "/api/admin/skud-reg/reconcile/run");
    const claim = await agentPost("/api/agent/jobs/claim", {});
    assert.equal(claim.data.reconcile, true);
    const rec = await ag.reconcile();
    assert.ok(rec.extra >= 6 && rec.missing >= 1);
    const rep = (await api(ADMIN, "GET", "/api/admin/skud-reg/reconcile")).data;
    const d1 = rep.devices.find((d) => d.ip === hik.ip), d2 = rep.devices.find((d) => d.ip === hik2.ip);
    assert.ok(d1.extra.some((x) => x.employeeNo === "424242"));
    assert.equal(d1.extra.length, 6);
    assert.ok(d2.missing.some((x) => x.employeeNo === "2024001"));
    assert.ok(notes.some((n) => /сверка/.test(n)));
    assert.equal((await api(ADMIN, "POST", "/api/admin/skud-reg/remove-employee", { json: { employeeNo: "2024001" } })).status, 409); // это студент из базы
    assert.equal((await api(ADMIN, "POST", "/api/admin/skud-reg/remove-employee", { json: { employeeNo: "424242", name: "Левый" } })).status, 200);
    await drain();
    assert.ok(!hik.users.has("424242"));
  });

  console.log("Сводка, Excel, журнал");
  await t("сводка по группам: кто без лица; куратор видит только свою группу", async () => {
    const all = (await api(ADMIN, "GET", "/api/admin/skud-reg/summary")).data;
    const a = all.groups.find((g) => g.name === "ЭС-21");
    assert.ok(a.total >= 4 && a.registered >= 1 && a.notReady === 1);
    const c = (await api(curA, "GET", "/api/admin/skud-reg/summary")).data;
    assert.ok(c.groups.every((g) => g.name === "ЭС-21"));
    const det = (await api(curA, "GET", `/api/admin/skud-reg/summary?groupId=${gA._id}`)).data;
    assert.ok(det.students.find((s) => s.fio.startsWith("Сидоров")).reason);
    assert.equal((await api(curA, "GET", `/api/admin/skud-reg/summary?groupId=${gB._id}`)).status, 403);
  });
  await t("Excel с теми, кого нет в СКУД", async () => {
    const x = await api(curA, "GET", "/api/admin/skud-reg/summary.xlsx?onlyMissing=1");
    assert.equal(x.status, 200);
    const wb = require("xlsx").read(x.data);
    const rows = require("xlsx").utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]]);
    assert.ok(rows.some((r) => r["ФИО"].startsWith("Сидоров")));
    assert.ok(!rows.some((r) => r["ФИО"].startsWith("Иванов")));
    assert.ok(!rows.some((r) => r["ФИО"].startsWith("Чужой")));
  });
  await t("журнал действий пишется и доступен админу", async () => {
    const log = (await api(ADMIN, "GET", "/api/admin/skud-reg/audit")).data;
    for (const a of ["photo_upload", "terminal_remove", "staff_approve", "link_create", "selflink_create", "student_photo_approve"]) assert.ok(log.some((x) => x.action === a), a);
    assert.ok(log.every((x) => !/photo|inn/i.test(Object.keys(x).join())));
    assert.ok((await api(ADMIN, "GET", "/api/admin/skud-reg/audit?action=link_create")).data.every((x) => x.action === "link_create"));
  });
  await t("страницы отдаются по ролям", async () => {
    assert.equal((await api(curA, "GET", "/skud-status.html")).status, 200);
    assert.equal((await api(curA, "GET", "/skud-staff.html")).status, 403);
    assert.equal((await api(ADMIN, "GET", "/skud-staff.html")).status, 200);
  });
  await t("Telegram: спецсимволы в ФИО экранируются", async () => {
    assert.equal(skud.core.mdEsc("a_b*c[d`"), "a\\_b\\*c\\[d\\`");
  });
  await t("имя на терминале укладывается в лимит, ИНН всегда целиком", async () => {
    const n = skud.core.buildDeviceName("Абдрахманова-Кадырбекова Айгерим Нурлановна", "12345678901234");
    assert.ok(n.length <= 32 && n.endsWith(" 12345678901234"), n);
    assert.equal(skud.core.buildDeviceName("Иванов Иван", "12345678901234"), "Иванов Иван 12345678901234");
  });

  console.log(`\nИтог: ${passed} прошло, ${failed} упало`);
  await hik.close(); await hik2.close(); srv.close();
  process.exit(failed ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
