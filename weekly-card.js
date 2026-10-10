// «Итог недели» для родителей: карточка-картинка (PNG) с посещаемостью и оценками ребёнка.
// Рассылка: пятница 17:00 (обычные группы), суббота 15:00 (группы с учебной субботой), Бишкек.
// Отключить: WEEKLY_CARD_CRON=off
const path = require("path");
const fs = require("fs");
const cron = require("node-cron");
const { Resvg } = require("@resvg/resvg-js");
const { shiftMinutes, dayStatus, isWeekend, weekRange, datesBetween, parseKey } = require("./attendance-utils");
const { reasonsIndex } = require("./absence");

const FONT_DIR = path.join(__dirname, "assets", "fonts");
const FONT_FILES = fs.existsSync(FONT_DIR) ? fs.readdirSync(FONT_DIR).filter((f) => f.endsWith(".ttf")).map((f) => path.join(FONT_DIR, f)) : [];

const WD = ["Вс", "Пн", "Вт", "Ср", "Чт", "Пт", "Сб"];
const MON_NOM = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"];
const MON = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля", "августа", "сентября", "октября", "ноября", "декабря"];
const xml = (v) => String(v ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" })[c]);
const dayNum = (dk) => Number(dk.slice(8, 10));
const rangeTitle = (from, to) => {
  const a = parseKey(from), b = parseKey(to);
  return a.getUTCMonth() === b.getUTCMonth()
    ? `${a.getUTCDate()}–${b.getUTCDate()} ${MON[b.getUTCMonth()]}`
    : `${a.getUTCDate()} ${MON[a.getUTCMonth()]} – ${b.getUTCDate()} ${MON[b.getUTCMonth()]}`;
};
// Подбор размера шрифта, чтобы строка влезла в ширину (приблизительно: средняя ширина символа ~0.56 em)
function fitSize(text, maxW, size, min = 26, k = 0.56) {
  let s = size;
  while (s > min && String(text).length * s * k > maxW) s -= 2;
  return s;
}
function clip(text, maxW, size, k = 0.56) {
  const max = Math.floor(maxW / (size * k));
  const t = String(text || "");
  return t.length > max ? t.slice(0, max - 1) + "…" : t;
}

const COLORS = {
  p: { bg: "#DCFCE7", fg: "#15803D", label: "был" },
  l: { bg: "#FEF3C7", fg: "#B45309", label: "опоздал" },
  a: { bg: "#FEE2E2", fg: "#DC2626", label: "нет" },
  e: { bg: "#E0E7FF", fg: "#4338CA", label: "по причине" },
  f: { bg: "#F1F5F9", fg: "#94A3B8", label: "" },
};

function createWeeklyCards({ Group, Student, TgSub, SummaryRecord, ensureGroupSummarySubjects, getTodayKey }) {
  // Данные карточки для студента за неделю, в которую входит dk
  async function weekData(stu, dk) {
    const group = stu.group && stu.group.name ? stu.group : await Group.findById(stu.group).lean();
    const sat = !!group?.saturday;
    const { from, to } = weekRange(dk);
    const today = getTodayKey();
    const days = datesBetween(from, to).filter((d) => !isWeekend(d, sat));
    const shiftMin = shiftMinutes(group?.shiftStart);
    const reasons = await reasonsIndex({ student: stu._id }, from, to);
    let present = 0, late = 0, absent = 0, excused = 0, past = 0;
    const cells = days.map((d) => {
      if (d > today) return { date: d, st: "f" };
      past++;
      const s = dayStatus(stu, d, shiftMin);
      if (s.present) {
        present++;
        if (s.late) late++;
        return { date: d, st: s.late ? "l" : "p", time: s.time };
      }
      const r = reasons.get(`${stu._id}|${d}`);
      if (r && r.excused && r.confirmed) { excused++; return { date: d, st: "e", code: r.label }; }
      absent++;
      return { date: d, st: r ? "e" : "a", code: r ? r.label + (r.confirmed ? "" : "?") : "" };
    });
    // оценки за месяц, который пересекает неделя (последний учебный день недели)
    const month = (days[days.length - 1] || to).slice(0, 7);
    let grades = [];
    if (group && SummaryRecord) {
      const rec = await SummaryRecord.findOne({ group: group._id, student: stu._id, month }).lean();
      const raw = rec?.grades instanceof Map ? Object.fromEntries(rec.grades) : rec?.grades || {};
      // нужен документ mongoose: функция может сохранить предметы по умолчанию
      const subjects = ensureGroupSummarySubjects ? await ensureGroupSummarySubjects(await Group.findById(group._id)) : [];
      grades = subjects.filter((s) => raw[s.key]).map((s) => ({ label: s.label, value: String(raw[s.key]) }));
    }
    return { fio: stu.fio, group: group?.name || "", from, to, title: rangeTitle(from, to), cells, present, late, absent, excused, past, grades, month };
  }

  function cardSvg(d) {
    const W = 1080, pad = 64;
    const nameSize = fitSize(d.fio, W - pad * 2, 64, 34, 0.62);
    const gRows = Math.max(1, Math.ceil(Math.min(d.grades.length, 8) / 2));
    const H = 900 + 70 + gRows * 64 + 150;
    const allGood = d.past > 0 && d.present === d.past && d.late === 0;
    const verdict = d.past === 0 ? "Неделя только началась" : allGood ? "Отличная неделя: ни одного пропуска и опоздания!" : d.absent === 0 ? "Хорошая неделя: пропусков без причины нет" : `Пропусков без уважительной причины: ${d.absent}`;
    const vColor = d.past === 0 ? "#475569" : allGood ? "#15803D" : d.absent === 0 ? "#1D4ED8" : "#DC2626";

    // плитки дней
    const n = d.cells.length, gap = 16;
    const tw = Math.floor((W - pad * 2 - gap * (n - 1)) / n), th = 230, ty = 420;
    const tiles = d.cells.map((c, i) => {
      const x = pad + i * (tw + gap);
      const col = COLORS[c.st];
      const wd = WD[parseKey(c.date).getUTCDay()];
      const mark = c.st === "p" ? `<path d="M${x + tw / 2 - 26} ${ty + 128} l18 18 l34 -38" stroke="${col.fg}" stroke-width="10" fill="none" stroke-linecap="round" stroke-linejoin="round"/>`
        : c.st === "l" ? `<circle cx="${x + tw / 2}" cy="${ty + 128}" r="26" fill="none" stroke="${col.fg}" stroke-width="8"/><path d="M${x + tw / 2} ${ty + 112} v18 l12 8" stroke="${col.fg}" stroke-width="7" fill="none" stroke-linecap="round"/>`
        : c.st === "a" ? `<path d="M${x + tw / 2 - 22} ${ty + 106} l44 44 M${x + tw / 2 + 22} ${ty + 106} l-44 44" stroke="${col.fg}" stroke-width="10" stroke-linecap="round"/>`
        : c.st === "e" ? `<circle cx="${x + tw / 2}" cy="${ty + 128}" r="24" fill="${col.fg}"/><rect x="${x + tw / 2 - 4}" y="${ty + 112}" width="8" height="32" rx="3" fill="#fff"/><rect x="${x + tw / 2 - 16}" y="${ty + 124}" width="32" height="8" rx="3" fill="#fff"/>`
        : `<circle cx="${x + tw / 2}" cy="${ty + 128}" r="6" fill="${col.fg}"/>`;
      const sub = c.st === "p" || c.st === "l" ? c.time || col.label : c.st === "e" ? clip(c.code, tw - 16, 20) : col.label;
      return `<rect x="${x}" y="${ty}" width="${tw}" height="${th}" rx="26" fill="${col.bg}"/>
        <text x="${x + tw / 2}" y="${ty + 46}" text-anchor="middle" font-family="Inter" font-weight="600" font-size="26" fill="${col.fg}">${wd}</text>
        <text x="${x + tw / 2}" y="${ty + 80}" text-anchor="middle" font-family="Inter" font-size="22" fill="${col.fg}" opacity="0.75">${dayNum(c.date)}</text>
        ${mark}
        <text x="${x + tw / 2}" y="${ty + 200}" text-anchor="middle" font-family="Inter" font-weight="600" font-size="${c.st === "e" ? 20 : 24}" fill="${col.fg}">${xml(sub)}</text>`;
    }).join("");

    // цифры недели
    const stats = [
      [`${d.present} из ${d.past || 0}`, "дней в лицее", "#15803D"],
      [String(d.late), "опозданий", "#B45309"],
      [String(d.absent), "без причины", "#DC2626"],
    ];
    const sw = Math.floor((W - pad * 2 - 32) / 3), sy = 690;
    const statSvg = stats.map(([v, l, c], i) => {
      const x = pad + i * (sw + 16);
      return `<rect x="${x}" y="${sy}" width="${sw}" height="150" rx="24" fill="#F8FAFC" stroke="#E2E8F0" stroke-width="2"/>
        <text x="${x + 28}" y="${sy + 78}" font-family="Inter" font-weight="800" font-size="52" fill="${c}">${xml(v)}</text>
        <text x="${x + 28}" y="${sy + 120}" font-family="Inter" font-size="24" fill="#64748B">${xml(l)}</text>`;
    }).join("");

    // оценки
    const gy = 900;
    const gList = d.grades.slice(0, 8);
    const gColor = (v) => (v === "5" ? "#15803D" : v === "4" ? "#1D4ED8" : v === "3" ? "#B45309" : v === "2" ? "#DC2626" : "#334155");
    const gradesSvg = gList.length
      ? gList.map((g, i) => {
          const col = i % 2, row = Math.floor(i / 2);
          const x = pad + col * ((W - pad * 2) / 2 + 8), y = gy + 70 + row * 64;
          const cw = (W - pad * 2) / 2 - 24;
          return `<text x="${x}" y="${y}" font-family="Inter" font-size="26" fill="#334155">${xml(clip(g.label, cw - 80, 26))}</text>
            <rect x="${x + cw - 56}" y="${y - 36}" width="52" height="48" rx="12" fill="${gColor(g.value)}"/>
            <text x="${x + cw - 30}" y="${y - 2}" text-anchor="middle" font-family="Inter" font-weight="800" font-size="28" fill="#fff">${xml(g.value)}</text>`;
        }).join("")
      : `<text x="${pad}" y="${gy + 70}" font-family="Inter" font-size="26" fill="#94A3B8">Оценок за ${MON_NOM[Number(d.month.slice(5, 7)) - 1]} пока нет</text>`;

    return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}">
      <defs><linearGradient id="hd" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="#1E3A8A"/><stop offset="1" stop-color="#4F46E5"/></linearGradient></defs>
      <rect width="${W}" height="${H}" fill="#FFFFFF"/>
      <rect width="${W}" height="330" fill="url(#hd)"/>
      <circle cx="${W - 90}" cy="70" r="150" fill="#FFFFFF" opacity="0.06"/><circle cx="${W - 40}" cy="250" r="90" fill="#FFFFFF" opacity="0.06"/>
      <text x="${pad}" y="92" font-family="Inter" font-weight="600" font-size="28" fill="#C7D2FE" letter-spacing="2">ПЛ №3 · ИТОГ НЕДЕЛИ</text>
      <text x="${pad}" y="140" font-family="Inter" font-size="30" fill="#E0E7FF">${xml(d.title)}</text>
      <text x="${pad}" y="232" font-family="Inter" font-weight="800" font-size="${nameSize}" fill="#FFFFFF">${xml(clip(d.fio, W - pad * 2, nameSize, 0.62))}</text>
      <text x="${pad}" y="284" font-family="Inter" font-size="30" fill="#E0E7FF">Группа ${xml(d.group)}</text>
      <text x="${pad}" y="388" font-family="Inter" font-weight="600" font-size="30" fill="${vColor}">${xml(verdict)}</text>
      ${tiles}
      ${statSvg}
      <text x="${pad}" y="${gy + 14}" font-family="Inter" font-weight="800" font-size="34" fill="#0F172A">Оценки за месяц</text>
      ${gradesSvg}
      <line x1="${pad}" y1="${H - 110}" x2="${W - pad}" y2="${H - 110}" stroke="#E2E8F0" stroke-width="2"/>
      <text x="${pad}" y="${H - 58}" font-family="Inter" font-size="24" fill="#94A3B8">Данные СКУД и журнала · бот ПЛ №3</text>
      <text x="${W - pad}" y="${H - 58}" text-anchor="end" font-family="Inter" font-weight="600" font-size="24" fill="#4F46E5">Хороших выходных!</text>
    </svg>`;
  }

  function renderPng(d) {
    const r = new Resvg(cardSvg(d), {
      font: { fontFiles: FONT_FILES, loadSystemFonts: FONT_FILES.length === 0, defaultFontFamily: "Inter" },
      fitTo: { mode: "width", value: 1080 },
    });
    return r.render().asPng();
  }

  async function cardFor(studentQuery, dk) {
    const stu = await Student.findOne(studentQuery).select("fio inn group skudDaily attendance").lean();
    if (!stu) return null;
    const data = await weekData(stu, dk || getTodayKey());
    return { data, png: renderPng(data) };
  }

  const caption = (d) => `🗓 Итог недели (${d.title})\n👤 ${d.fio}, ${d.group}\n✅ В лицее: ${d.present} из ${d.past} дн. · ⏰ опозданий: ${d.late} · ❌ без причины: ${d.absent}`;
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // Разослать родителям. onlySaturday: true — только группы с учебной субботой, false — только без неё, null — все
  async function sendAll(bot, { onlySaturday = null } = {}) {
    const subs = await TgSub.find({ role: "parent" }).lean();
    const satGroups = new Set((await Group.find({ saturday: true }).select("_id").lean()).map((g) => String(g._id)));
    let sent = 0;
    for (const sub of subs) {
      for (const inn of sub.linkedInns || []) {
        try {
          const stu = await Student.findOne({ inn }).select("group").lean();
          if (!stu) continue;
          const isSat = satGroups.has(String(stu.group));
          if (onlySaturday !== null && isSat !== onlySaturday) continue;
          const c = await cardFor({ inn });
          if (!c || !c.data.past) continue;
          await bot.sendPhoto(sub.chatId, c.png, { caption: caption(c.data) }, { filename: "itog-nedeli.png", contentType: "image/png" });
          sent++;
          await sleep(60); // лимиты Telegram
        } catch (e) {
          console.error("weekly card:", e.message);
        }
      }
    }
    console.log(`🖼 Итог недели отправлен родителям: ${sent}`);
    return { sent };
  }

  function schedule(bot) {
    if (String(process.env.WEEKLY_CARD_CRON || "").toLowerCase() === "off") return;
    const opts = { timezone: "Asia/Bishkek" };
    cron.schedule("0 17 * * 5", () => sendAll(bot, { onlySaturday: false }).catch((e) => console.error(e)), opts);
    cron.schedule("0 15 * * 6", () => sendAll(bot, { onlySaturday: true }).catch((e) => console.error(e)), opts);
    console.log("🖼 Итог недели родителям: пт 17:00 (сб 15:00 для групп с учебной субботой)");
  }

  return { weekData, cardSvg, renderPng, cardFor, sendAll, schedule, caption };
}

module.exports = { createWeeklyCards };
