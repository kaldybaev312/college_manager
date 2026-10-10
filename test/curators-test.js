// Тесты управления доступом кураторов (curators.js). БД — в памяти.  Запуск: npm test
const Module = require("module");
const path = require("path");
const crypto = require("crypto");
const assert = require("assert/strict");

const fake = require("./fake-mongoose");
const origResolve = Module._resolveFilename;
Module._resolveFilename = function (req, ...rest) {
  return req === "mongoose" ? path.join(__dirname, "fake-mongoose.js") : origResolve.call(this, req, ...rest);
};
const express = require("express");
const { Schema } = fake;

const Group = fake.model("Group", new Schema({ name: String }));
const Curator = fake.model(
  "Curator",
  new Schema({ username: String, passwordHash: String, fullName: String, groups: [{ type: "ObjectId", ref: "Group" }], sessionVer: { type: Number, default: 0 } }, { timestamps: true }),
);
const TgSub = fake.model("TgSub", new Schema({ chatId: String, role: String, curatorId: String }));

const hashPassword = (p) => "h:" + crypto.createHash("sha256").update(p).digest("hex");
const changed = [];

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
  const gA = await Group.create({ name: "ИТ-21" });
  const gB = await Group.create({ name: "СВ-11" });
  const gC = await Group.create({ name: "ПК-31" });

  const app = express();
  app.use(express.json());
  const isAdmin = (req) => req.headers["x-role"] === "admin";
  require("../curators")(app, {
    Curator,
    Group,
    TgSub,
    hashPassword,
    adminLogin: "admin",
    requireAdmin: (req, res, next) => (isAdmin(req) ? next() : res.status(403).json({ error: "Admin only" })),
    onChange: (id) => changed.push(String(id)),
  });
  const srv = app.listen(0);
  const base = `http://127.0.0.1:${srv.address().port}`;
  const call = async (method, url, body, role = "admin") => {
    const r = await fetch(base + url, { method, headers: { "Content-Type": "application/json", "x-role": role }, body: body ? JSON.stringify(body) : undefined });
    return { status: r.status, data: await r.json().catch(() => null) };
  };

  console.log("Доступ кураторов:");
  let id;
  await t("создание: проверка логина, пароля, групп", async () => {
    assert.equal((await call("POST", "/api/admin/curators", { username: "ив", password: "123456", groupIds: [String(gA._id)] })).status, 400);
    assert.equal((await call("POST", "/api/admin/curators", { username: "ivanova", password: "123", groupIds: [String(gA._id)] })).status, 400);
    assert.equal((await call("POST", "/api/admin/curators", { username: "ivanova", password: "123456", groupIds: [] })).status, 400);
    assert.equal((await call("POST", "/api/admin/curators", { username: "admin", password: "123456", groupIds: [String(gA._id)] })).status, 400);
    const r = await call("POST", "/api/admin/curators", { username: "ivanova", password: "secret1", fullName: "Иванова А.А.", groupIds: [String(gA._id), String(gB._id)] });
    assert.equal(r.status, 200, JSON.stringify(r.data));
    id = r.data.curator._id;
    assert.deepEqual(r.data.curator.groupNames, ["ИТ-21", "СВ-11"]);
  });
  await t("тот же логин второй раз не перезаписывает куратора", async () => {
    const r = await call("POST", "/api/admin/curators", { username: "ivanova", password: "other12", groupIds: [String(gC._id)] });
    assert.equal(r.status, 400);
    assert.match(r.data.error, /занят/);
    const list = (await call("GET", "/api/admin/curators")).data;
    assert.equal(list.length, 1);
    assert.deepEqual(list[0].groupNames, ["ИТ-21", "СВ-11"]);
  });
  await t("не-админ не может управлять кураторами", async () => {
    assert.equal((await call("GET", "/api/admin/curators", null, "curator")).status, 403);
    assert.equal((await call("PUT", `/api/admin/curators/${id}`, { groupIds: [] }, "curator")).status, 403);
  });
  await t("смена групп: дать и убрать — без выхода из системы", async () => {
    changed.length = 0;
    const r = await call("PUT", `/api/admin/curators/${id}`, { groupIds: [String(gB._id), String(gC._id)] });
    assert.deepEqual(r.data.curator.groupNames, ["СВ-11", "ПК-31"]);
    assert.equal(r.data.loggedOut, false);
    assert.deepEqual(changed, [id]);
    const cur = Curator._store.find((c) => String(c._id) === id);
    assert.equal(cur.sessionVer, 0);
  });
  await t("смена ФИО без пароля не трогает пароль и вход", async () => {
    const before = Curator._store.find((c) => String(c._id) === id).passwordHash;
    const r = await call("PUT", `/api/admin/curators/${id}`, { fullName: "Иванова Айгуль" });
    assert.equal(r.data.curator.fullName, "Иванова Айгуль");
    assert.equal(Curator._store.find((c) => String(c._id) === id).passwordHash, before);
    assert.equal(r.data.loggedOut, false);
  });
  await t("смена логина и пароля завершает старые входы", async () => {
    await Curator.create({ username: "petrov", passwordHash: "x", groups: [] });
    assert.equal((await call("PUT", `/api/admin/curators/${id}`, { username: "petrov" })).status, 400);
    assert.equal((await call("PUT", `/api/admin/curators/${id}`, { password: "12" })).status, 400);
    const r = await call("PUT", `/api/admin/curators/${id}`, { username: "aigul_it", password: "newpass1" });
    assert.equal(r.status, 200);
    assert.equal(r.data.loggedOut, true);
    const cur = Curator._store.find((c) => String(c._id) === id);
    assert.equal(cur.username, "aigul_it");
    assert.equal(cur.passwordHash, hashPassword("newpass1"));
    assert.equal(cur.sessionVer, 1);
  });
  await t("«выйти везде» отвязывает Telegram и сбрасывает входы", async () => {
    await TgSub.create({ chatId: "55", role: "curator", curatorId: id });
    const r = await call("POST", `/api/admin/curators/${id}/logout`);
    assert.equal(r.data.telegramUnlinked, 1);
    assert.equal(Curator._store.find((c) => String(c._id) === id).sessionVer, 2);
    assert.equal(TgSub._store.length, 0);
  });
  await t("удаление убирает куратора и его Telegram", async () => {
    await TgSub.create({ chatId: "56", role: "curator", curatorId: id });
    assert.equal((await call("DELETE", `/api/admin/curators/${id}`)).status, 200);
    assert.ok(!Curator._store.some((c) => String(c._id) === id));
    assert.equal(TgSub._store.length, 0);
  });

  console.log(`\nИтог: ${passed} прошло, ${failed} упало`);
  srv.close();
  process.exit(failed ? 1 : 0);
})().catch((e) => {
  console.error(e);
  process.exit(1);
});
