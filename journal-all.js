// Общий журнал посещаемости по всему лицею: листы «2 курс», «1 курс», «10 месяцев».
// Формат как в образце: название группы (объединённая шапка), №, ФИО, рабочие дни; П/О/Н/Б/С/Л/Д/-.
const ExcelJS = require("exceljs");

function courseOf(name, month) {
  const m = /\/\s*(\d{2})\b/.exec(name) || /(\d{2})\D*$/.exec(name);
  if (!m) return null;
  const [y, mo] = month.split("-").map(Number);
  const start = mo >= 8 ? y : y - 1; // год начала учебного года
  const c = start - (2000 + Number(m[1])) + 1;
  return c >= 1 && c <= 4 ? c : null;
}

function sheetKey(g, month) {
  if (/10\s*мес/i.test(g.duration || "")) return "10 месяцев";
  const c = courseOf(g.name || "", month);
  return c ? `${c} курс` : "Прочие";
}
const ORDER = ["2 курс", "1 курс", "10 месяцев", "3 курс", "4 курс", "Прочие"];

function create({ Group, Student, absence, att, getTodayKey }) {
  const thin = { style: "thin" };
  const med = { style: "medium" };

  async function build(month) {
    const { from, to } = att.monthRange(month);
    const dim = Number(to.slice(8));
    const tdk = getTodayKey();
    const groups = (await Group.find({}).lean()).filter((g) => !g.kind || g.kind === "students");
    groups.sort((a, b) => String(a.name).localeCompare(String(b.name), "ru", { numeric: true }));
    const reasons = await absence.reasonsIndex({}, from, to);
    const sheets = new Map();
    for (const g of groups) {
      const k = sheetKey(g, month);
      if (!sheets.has(k)) sheets.set(k, []);
      sheets.get(k).push(g);
    }
    const wb = new ExcelJS.Workbook();
    const keys = [...sheets.keys()].sort((a, b) => ORDER.indexOf(a) - ORDER.indexOf(b));
    if (!keys.length) wb.addWorksheet("Журнал");
    for (const k of keys) {
      const gs = sheets.get(k);
      const withSat = gs.some((g) => g.saturday);
      const days = [];
      for (let d = 1; d <= dim; d++) {
        const dk = `${month}-${att.pad2(d)}`;
        if (dk > tdk) break;
        const wd = new Date(`${dk}T00:00:00Z`).getUTCDay();
        if (wd === 0 || (wd === 6 && !withSat)) continue;
        days.push({ d, dk });
      }
      const ws = wb.addWorksheet(k, { views: [{ zoomScale: 60 }] });
      const lastCol = 2 + Math.max(days.length, 1);
      ws.getColumn(1).width = 4.8;
      ws.getColumn(2).width = 35.8;
      for (let c = 3; c <= lastCol; c++) ws.getColumn(c).width = 4.8;
      let r = 1;
      const font = { name: "Calibri", size: 12 };
      for (const g of gs) {
        const students = await Student.find({ group: g._id }).sort({ fio: 1 }).select("fio skudDaily attendance").lean();
        const shiftMin = att.shiftMinutes(g.shiftStart);
        ws.mergeCells(r, 1, r, lastCol);
        const t = ws.getCell(r, 1);
        t.value = g.name;
        t.font = { ...font, bold: true };
        t.alignment = { horizontal: "center", vertical: "middle" };
        for (let c = 1; c <= lastCol; c++) ws.getCell(r, c).border = { top: med, bottom: med, left: c === 1 ? med : undefined, right: c === lastCol ? med : undefined };
        r++;
        const h = ["№", "ФИО Студента", ...days.map((x) => x.d)];
        h.forEach((v, i) => {
          const c = ws.getCell(r, i + 1);
          c.value = v;
          c.font = font;
          c.alignment = { horizontal: i === 1 ? "left" : "center" };
          c.border = { top: thin, bottom: thin, left: i === 0 ? med : thin, right: i === 0 ? med : thin };
        });
        r++;
        students.forEach((s, i) => {
          const row = [i + 1, s.fio];
          for (const { dk } of days) {
            const st = att.dayStatus(s, dk, shiftMin);
            if (st.present) row.push(st.late ? "О" : "П");
            else if (!att.isWeekend(dk, g.saturday)) {
              const x = reasons.get(`${s._id}|${dk}`);
              row.push(x ? x.code + (x.confirmed ? "" : "?") : "Н");
            } else row.push("-");
          }
          row.forEach((v, j) => {
            const c = ws.getCell(r, j + 1);
            c.value = v;
            c.font = font;
            c.alignment = { horizontal: j === 1 ? "left" : "center" };
            c.border = { top: thin, bottom: thin, left: j === 0 ? med : thin, right: j === 0 ? med : thin };
          });
          r++;
        });
        r++; // пустая строка между группами
      }
      const lg = ws.getCell(r, 2);
      lg.value = "П — был, О — опоздал, Н — без причины, Б — болезнь, С — семейные, Л — по направлению лицея, Д — другая причина, ? — не подтверждено";
      lg.font = { name: "Calibri", size: 10, italic: true };
    }
    return wb.xlsx.writeBuffer();
  }

  function registerRoutes(app, { requireAdmin, contentDisposition }) {
    app.get("/api/admin/export-journal-all/:month", requireAdmin, async (req, res) => {
      try {
        const { month } = req.params;
        if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(month)) return res.status(400).send("Неверный месяц");
        const buf = await build(month);
        res.setHeader("Content-Disposition", contentDisposition(`Journal_all_${month}.xlsx`));
        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.send(Buffer.from(buf));
      } catch (e) {
        res.status(500).send(e.message);
      }
    });
  }
  return { build, registerRoutes };
}

module.exports = { create, courseOf, sheetKey };
