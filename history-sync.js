// Восстановление истории: агент присылает уже сведённые по дням данные (первый/последний вход и выход).
// Запись идемпотентна (можно запускать повторно), уведомления родителям и «опоздавшие» не отправляются.
module.exports = function mountHistorySync(app, { Student, AGENT_API_KEY, safeEq, normalizeINN, extractINN, normalizeFIO, extractFIO }) {
  app.post("/api/agent/history-sync", async (req, res) => {
  const { apiKey, days } = req.body || {};
  if (!safeEq(apiKey, AGENT_API_KEY)) return res.status(403).send("Forbidden");
  if (!Array.isArray(days)) return res.status(400).json({ error: "array expected" });
  const ms = (x) => (x ? new Date(x).getTime() : NaN);
  const pickMin = (a, b) => (a && (!b || !(ms(b) <= ms(a))) ? a : b);
  const pickMax = (a, b) => (a && (!b || !(ms(b) >= ms(a))) ? a : b);
  const cache = new Map();
  const find = async (rawName) => {
    if (cache.has(rawName)) return cache.get(rawName);
    const inn = normalizeINN(extractINN(rawName)),
      fioNorm = normalizeFIO(extractFIO(rawName));
    let st = null;
    if (inn) st = await Student.findOne({ inn }).select("fio inn group");
    if (!st && fioNorm) {
      const fp = fioNorm.split(" ")[0],
        cands = await Student.find({ fio: new RegExp("^" + fp.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i") })
          .select("fio inn group")
          .limit(50);
      st = cands.find((c) => normalizeFIO(c.fio) === fioNorm) || null;
    }
    cache.set(rawName, st);
    return st;
  };
  let ok = 0, fail = 0;
  const unknown = new Set();
  for (const e of days.slice(0, 500)) {
    try {
      const rawName = String(e.rawName || "").trim(),
        dateKey = String(e.date || "");
      if (!rawName || !/^\d{4}-\d{2}-\d{2}$/.test(dateKey)) { fail++; continue; }
      const student = await find(rawName);
      if (!student) { unknown.add(rawName); fail++; continue; }
      const full = await Student.findById(student._id).select("skudDaily attendance").lean();
      const row = (full?.skudDaily || []).find((d) => d.date === dateKey);
      const inN = Math.max(0, Number(e.inCount) || 0), outN = Math.max(0, Number(e.outCount) || 0);
      const merged = {
        date: dateKey,
        inCount: Math.max(row?.inCount || 0, inN),
        outCount: Math.max(row?.outCount || 0, outN),
        present: !!(row?.present || inN > 0),
        firstIn: pickMin(inN ? e.firstIn : null, row?.firstIn) || undefined,
        lastIn: pickMax(inN ? e.lastIn : null, row?.lastIn) || undefined,
        firstOut: pickMin(outN ? e.firstOut : null, row?.firstOut) || undefined,
        lastOut: pickMax(outN ? e.lastOut : null, row?.lastOut) || undefined,
      };
      if (!row) {
        await Student.updateOne({ _id: student._id, "skudDaily.date": { $ne: dateKey } }, { $push: { skudDaily: merged } });
      } else {
        const set = {};
        for (const [k, v] of Object.entries(merged)) if (k !== "date" && v !== undefined) set[`skudDaily.$.${k}`] = v;
        await Student.updateOne({ _id: student._id, "skudDaily.date": dateKey }, { $set: set });
      }
      if (merged.present) {
        const a = (full?.attendance || []).find((d) => d.date === dateKey);
        if (!a) await Student.updateOne({ _id: student._id, "attendance.date": { $ne: dateKey } }, { $push: { attendance: { date: dateKey, present: true } } });
        else if (!a.present) await Student.updateOne({ _id: student._id, "attendance.date": dateKey }, { $set: { "attendance.$.present": true } });
      }
      ok++;
    } catch (err) {
      fail++;
    }
  }
  res.json({ success: true, okCount: ok, failCount: fail, unknown: [...unknown].slice(0, 20) });
});
};
