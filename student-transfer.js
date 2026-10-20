// Перевод студента из одной группы в другую (только администратор).
// Переезжают: группа студента и его причины отсутствия (чтобы журнал новой группы показывал Б/С/Л/Д за прошлые дни).
// Остаются как есть: посещаемость (хранится у студента), оценки сводки (принадлежат предметам старой группы),
// запись на терминалах (ФИО + ИНН от группы не зависят).
module.exports = function mountStudentTransfer(app, { Student, Group, AbsenceReason, requireAdmin, audit }) {
  const isStudentGroup = (g) => !g.kind || g.kind === "students";

  app.post("/api/admin/student/:id/transfer", requireAdmin, async (req, res) => {
    try {
      const st = await Student.findById(req.params.id);
      if (!st) return res.status(404).json({ error: "Студент не найден" });
      const to = req.body?.groupId ? await Group.findById(req.body.groupId).catch(() => null) : null;
      if (!to) return res.status(400).json({ error: "Выберите группу" });
      if (!isStudentGroup(to)) return res.status(400).json({ error: "Переводить можно только в группу студентов" });
      if (String(st.group) === String(to._id)) return res.status(400).json({ error: "Студент уже в этой группе" });
      const from = st.group ? await Group.findById(st.group).catch(() => null) : null;
      if (from && !isStudentGroup(from)) return res.status(400).json({ error: "Это не студент" });

      await Student.updateOne({ _id: st._id }, { $set: { group: to._id } });
      await AbsenceReason.updateMany({ student: st._id }, { $set: { group: to._id } });
      if (audit) await audit(req, "student_transfer", st.fio, `${from?.name || "—"} → ${to.name}`).catch(() => {});
      res.json({ ok: true, from: from?.name || "", to: to.name });
    } catch (e) {
      res.status(500).json({ error: e.message });
    }
  });
};
