// Доступ кураторов: список, создание, изменение (ФИО, логин, пароль, группы),
// «выйти на всех устройствах», удаление. Только для администратора.
// Смена логина/пароля и «выйти везде» увеличивают sessionVer — старые входы перестают работать.
const mongoose = require("mongoose");

module.exports = function registerCuratorRoutes(app, { Curator, Group, TgSub, hashPassword, adminLogin, requireAdmin, onChange = () => {} }) {
  const USERNAME_RE = /^[a-zA-Z0-9._-]{3,32}$/;
  function cleanGroupIds(raw) {
    return [...new Set((Array.isArray(raw) ? raw : []).map((id) => String(id || "").trim()))].filter((id) =>
      mongoose.isValidObjectId(id),
    );
  }
  async function curatorJSON(c) {
    const tg = await TgSub.countDocuments({ role: "curator", curatorId: String(c._id) });
    const groups = (c.groups || []).filter((g) => g && g._id);
    return {
      _id: c._id,
      username: c.username,
      fullName: c.fullName || "",
      groupIds: groups.map((g) => String(g._id)),
      groupNames: groups.map((g) => g.name),
      telegram: tg,
      createdAt: c.createdAt,
      updatedAt: c.updatedAt,
    };
  }
  const afterCuratorChange = (id) => onChange(id);
  async function checkUsername(username, exceptId) {
    if (!USERNAME_RE.test(username))
      return "Логин: 3–32 символа — латинские буквы, цифры, точка, дефис, подчёркивание";
    const other = await Curator.findOne({ username }).select("_id").lean();
    if (other && String(other._id) !== String(exceptId || "")) return "Такой логин уже занят";
    if (username === adminLogin) return "Этот логин занят администратором";
    return null;
  }
  const checkPassword = (p) => (String(p).length < 6 ? "Пароль — минимум 6 символов" : null);

  app.get("/api/admin/curators", requireAdmin, async (req, res) => {
    try {
      const curators = await Curator.find().populate("groups", "name").sort({ fullName: 1, username: 1 }).lean();
      res.json(await Promise.all(curators.map(curatorJSON)));
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Создать куратора
  app.post("/api/admin/curators", requireAdmin, async (req, res) => {
    try {
      const username = String(req.body.username || "").trim();
      const password = String(req.body.password || "");
      const fullName = String(req.body.fullName || "").trim().slice(0, 120);
      const groupIds = cleanGroupIds(req.body.groupIds || (req.body.groupId ? [req.body.groupId] : []));
      const err = (await checkUsername(username)) || checkPassword(password) || (!groupIds.length && "Выберите хотя бы одну группу");
      if (err) return res.status(400).json({ error: err });
      const groups = await Group.find({ _id: { $in: groupIds } }).select("_id");
      if (!groups.length) return res.status(404).json({ error: "Группы не найдены" });
      const doc = await Curator.create({ username, passwordHash: hashPassword(password), fullName, groups: groups.map((g) => g._id) });
      res.json({ ok: true, curator: await curatorJSON(await Curator.findById(doc._id).populate("groups", "name").lean()) });
    } catch (e) {
      res.status(e.code === 11000 ? 400 : 500).json({ error: e.code === 11000 ? "Такой логин уже занят" : e.message });
    }
  });

  // Изменить: любые из полей fullName, username, password, groupIds
  app.put("/api/admin/curators/:id", requireAdmin, async (req, res) => {
    try {
      const cur = await Curator.findById(req.params.id);
      if (!cur) return res.status(404).json({ error: "Куратор не найден" });
      const b = req.body || {};
      const set = {};
      let bump = false;
      if (b.fullName !== undefined) set.fullName = String(b.fullName || "").trim().slice(0, 120);
      if (b.username !== undefined) {
        const username = String(b.username || "").trim();
        if (username !== cur.username) {
          const err = await checkUsername(username, cur._id);
          if (err) return res.status(400).json({ error: err });
          set.username = username;
          bump = true;
        }
      }
      if (b.password) {
        const err = checkPassword(b.password);
        if (err) return res.status(400).json({ error: err });
        set.passwordHash = hashPassword(String(b.password));
        bump = true;
      }
      if (b.groupIds !== undefined) {
        const ids = cleanGroupIds(b.groupIds);
        const groups = ids.length ? await Group.find({ _id: { $in: ids } }).select("_id") : [];
        set.groups = groups.map((g) => g._id);
      }
      const upd = { $set: set };
      if (bump) upd.$inc = { sessionVer: 1 };
      await Curator.updateOne({ _id: cur._id }, upd);
      afterCuratorChange(cur._id);
      res.json({
        ok: true,
        loggedOut: bump,
        curator: await curatorJSON(await Curator.findById(cur._id).populate("groups", "name").lean()),
      });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Выйти везде: сайт (все браузеры) и Telegram-бот
  app.post("/api/admin/curators/:id/logout", requireAdmin, async (req, res) => {
    try {
      const cur = await Curator.findByIdAndUpdate(req.params.id, { $inc: { sessionVer: 1 } });
      if (!cur) return res.status(404).json({ error: "Куратор не найден" });
      const tg = await TgSub.deleteMany({ role: "curator", curatorId: String(cur._id) });
      afterCuratorChange(cur._id);
      res.json({ ok: true, telegramUnlinked: tg.deletedCount || 0 });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });

  // Очистить дублирующихся кураторов (оставить первого по дате)
  app.delete("/api/admin/curators/cleanup", requireAdmin, async (req, res) => {
    try {
      const all = await Curator.find().sort({ createdAt: 1 });
      const seen = new Map();
      const toDelete = [];
      for (const c of all) {
        if (seen.has(c.username)) toDelete.push(c._id);
        else seen.set(c.username, c._id);
      }
      if (toDelete.length) {
        await Curator.deleteMany({ _id: { $in: toDelete } });
        await TgSub.deleteMany({ role: "curator", curatorId: { $in: toDelete.map(String) } });
        toDelete.forEach(afterCuratorChange);
      }
      res.json({ ok: true, deleted: toDelete.length });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
  app.delete("/api/admin/curators/:id", requireAdmin, async (req, res) => {
    try {
      await Curator.findByIdAndDelete(req.params.id);
      await TgSub.deleteMany({ role: "curator", curatorId: String(req.params.id) });
      afterCuratorChange(req.params.id);
      res.json({ ok: true });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
};
