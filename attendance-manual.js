// Ручная правка посещаемости (только администратор): отметить приход со временем или «не был».
// Правка помечается manual (в журнале значок ✎) и пишется в журнал СКУД с прежним значением.
module.exports = function mountAttendanceManual(app, { Student, requireAdmin, audit, att, getTodayKey, getSession }) {
  const timeOk = (t) => /^([01]\d|2[0-3]):[0-5]\d$/.test(String(t || ""));
  const describe = (row, a) => {
    const present = !!(row?.present || a?.present);
    const t = att.localTime(row?.firstIn)?.str;
    return present ? `был${t ? " " + t : ""}` : "не был";
  };

  app.post("/api/admin/attendance-manual", requireAdmin, async (req, res) => {
    try {
      const { studentId, date, status, time, outTime } = req.body || {};
      if (!att.isValidKey(date)) return res.status(400).json({ error: "Неверная дата" });
      if (date > getTodayKey()) return res.status(400).json({ error: "Нельзя менять будущие дни" });
      if (!["present", "absent"].includes(status)) return res.status(400).json({ error: "Неверный статус" });
      if (status === "present" && !timeOk(time)) return res.status(400).json({ error: "Укажите время прихода (ЧЧ:ММ)" });
      if (outTime && !timeOk(outTime)) return res.status(400).json({ error: "Неверное время ухода" });
      const st = await Student.findById(studentId).select("fio skudDaily attendance");
      if (!st) return res.status(404).json({ error: "Студент не найден" });

      const row = (st.skudDaily || []).find((r) => r.date === date);
      const a = (st.attendance || []).find((r) => r.date === date);
      const before = describe(row, a);
      const who = (getSession && getSession(req)?.username) || "admin";

      // строка дня должна существовать
      if (!row) await Student.updateOne({ _id: st._id, "skudDaily.date": { $ne: date } }, { $push: { skudDaily: { date, inCount: 0, outCount: 0, present: false } } });
      const set = { "skudDaily.$.manual": true, "skudDaily.$.manualBy": who };
      const unset = {};
      if (status === "present") {
        const iso = `${date}T${time}:00+06:00`;
        set["skudDaily.$.present"] = true;
        set["skudDaily.$.firstIn"] = iso;
        set["skudDaily.$.lastIn"] = iso;
        set["skudDaily.$.inCount"] = Math.max(1, row?.inCount || 0);
        if (outTime) {
          const o = `${date}T${outTime}:00+06:00`;
          set["skudDaily.$.firstOut"] = o;
          set["skudDaily.$.lastOut"] = o;
          set["skudDaily.$.outCount"] = Math.max(1, row?.outCount || 0);
        }
      } else {
        Object.assign(set, { "skudDaily.$.present": false, "skudDaily.$.inCount": 0, "skudDaily.$.outCount": 0 });
        for (const k of ["firstIn", "lastIn", "firstOut", "lastOut"]) unset[`skudDaily.$.${k}`] = 1;
      }
      await Student.updateOne({ _id: st._id, "skudDaily.date": date }, Object.keys(unset).length ? { $set: set, $unset: unset } : { $set: set });

      // «attendance» — отдельная отметка; dayStatus считает присутствие, если true хотя бы в одной
      const present = status === "present";
      if (!a) await Student.updateOne({ _id: st._id, "attendance.date": { $ne: date } }, { $push: { attendance: { date, present } } });
      else await Student.updateOne({ _id: st._id, "attendance.date": date }, { $set: { "attendance.$.present": present } });

      const after = present ? `был ${time}${outTime ? `, ушёл ${outTime}` : ""}` : "не был";
      if (audit) await audit(req, "attendance_manual", st.fio, `${date}: ${before} → ${after}`).catch(() => {});
      res.json({ ok: true, before, after });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
};
