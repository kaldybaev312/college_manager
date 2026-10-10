// Учёт сотрудников по СКУД.
//   Преподаватели — строгий контроль по расписанию Kundoluk: опоздание к первому уроку (мин),
//   ранний уход после последнего урока, часы в лицее, отсутствие в день с уроками,
//   уведомления администратору в Telegram, табель за месяц.
//   Мастера — был/не был и часы в лицее (без опозданий и уведомлений).
// Расписание приходит из редактора расписания (POST /api/agent/teacher-schedule по SCHEDULE_API_KEY)
// или загружается в админке файлом Excel из раздела «Отчёты» редактора.
const path = require("path");
const multer = require("multer");
const xlsx = require("xlsx");
const cron = require("node-cron");
const mongoose = require("mongoose");
const { localTime, isWeekend, isValidKey, addDays, monthRange, datesBetween, parseKey, pad2 } = require("./attendance-utils");
const { reasonsIndex } = require("./absence");

const { Schema } = mongoose;
const model = (name, schema) => mongoose.models[name] || mongoose.model(name, schema);

// Один урок одного преподавателя в конкретную дату
const StaffLesson = model(
  "StaffLesson",
  new Schema(
    {
      date: { type: String, index: true },
      gradeName: String,
      lesson: Number,
      start: String, // "08:00"
      end: String, // "09:20"
      subject: String,
      room: String,
      staffName: String, // как записано в расписании
      key: { type: String, index: true }, // нормализованное ФИО для сопоставления
      source: String, // kundoluk | excel
    },
    { timestamps: true, versionKey: false },
  ),
);
// Ручное сопоставление: ФИО из расписания → сотрудник СКУД
const StaffAlias = model("StaffAlias", new Schema({ key: { type: String, unique: true }, name: String, staff: Schema.Types.ObjectId }, { timestamps: true, versionKey: false }));
const StaffImport = model(
  "StaffImport",
  new Schema({ source: String, weeks: [String], groups: Number, lessons: Number, unmatched: [String], by: String }, { timestamps: true, versionKey: false }),
);
const StaffAlertLog = model(
  "StaffAlertLog",
  new Schema({ staff: String, date: String, type: String, at: { type: Date, default: Date.now, expires: 60 * 24 * 3600 } }, { versionKey: false }),
);

const WD = ["вс", "пн", "вт", "ср", "чт", "пт", "сб"];
const DAY_NAMES = { понедельник: 1, вторник: 2, среда: 3, четверг: 4, пятница: 5, суббота: 6, пн: 1, вт: 2, ср: 3, чт: 4, пт: 5, сб: 6 };
const toMin = (hhmm) => {
  const m = /^(\d{1,2}):(\d{2})/.exec(String(hhmm || ""));
  return m ? Number(m[1]) * 60 + Number(m[2]) : null;
};
const fromMin = (m) => (m == null ? "" : `${pad2(Math.floor(m / 60))}:${pad2(m % 60)}`);

// «08:00-09:20», «8.00 – 9.20» → { start, end }
function parseBell(bell) {
  const m = /(\d{1,2})[:.](\d{2})\s*[-–—]\s*(\d{1,2})[:.](\d{2})/.exec(String(bell || ""));
  if (!m) return null;
  return { start: `${pad2(+m[1])}:${m[2]}`, end: `${pad2(+m[3])}:${m[4]}` };
}

// ФИО → ключ «фамилия и о»: «Иванова Айгуль Асановна» и «Иванова А.А.» совпадают
function nameKey(fio) {
  const words = String(fio || "")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/[^a-zа-яөүң\s.-]/gi, " ")
    .replace(/\./g, " ")
    .split(/\s+/)
    .filter(Boolean);
  if (!words.length) return "";
  const [last, ...rest] = words;
  return [last, ...rest.slice(0, 2).map((w) => w[0])].join(" ");
}
// «фамилия и» — запасной ключ, если в одном месте записано без отчества
const shortKey = (key) => String(key || "").split(" ").slice(0, 2).join(" ");
// Учителей в одном уроке может быть несколько (подгруппы): «Иванова А.А. / Петров Б.Б.»
const splitStaff = (s) => String(s || "").split(/\s*[/;,]\s*|\s+и\s+/).map((x) => x.trim()).filter((x) => x && x !== "-" && x !== "—");

const isKind = (kind) => (g) => g && g.kind === kind;

function createStaffAttendance({ Group, Student, TgSub, getTodayKey, getBishkekMinutes }) {
  const LATE_GRACE = () => Math.max(0, parseInt(process.env.TEACHER_LATE_GRACE, 10) || 0);

  /* ---------- импорт расписания ---------- */
  // groups: [{ gradeName, weekStart: "ГГГГ-ММ-ДД" (понедельник), lessons: [{ weekday 1..6 | date, lesson, bell, subjectName, staffName, roomName }] }]
  async function importSchedule(groups, { source, by } = {}) {
    if (!Array.isArray(groups) || !groups.length) throw new Error("Пустое расписание");
    const teachers = await teacherIndex();
    const docs = [];
    const covered = [];
    const weeks = new Set();
    const unmatched = new Set();
    for (const g of groups) {
      const gradeName = String(g.gradeName || "").trim();
      if (!gradeName || !isValidKey(g.weekStart)) throw new Error(`Группа «${gradeName || "?"}»: нет даты начала недели`);
      weeks.add(g.weekStart);
      const dates = datesBetween(g.weekStart, addDays(g.weekStart, 5));
      covered.push(...dates.map((date) => ({ date, gradeName })));
      for (const l of g.lessons || []) {
        const date = isValidKey(l.date) ? l.date : Number(l.weekday) >= 1 && Number(l.weekday) <= 6 ? addDays(g.weekStart, Number(l.weekday) - 1) : null;
        const bell = parseBell(l.bell) || (l.start && l.end ? { start: l.start, end: l.end } : null);
        if (!date || !bell) continue;
        for (const name of splitStaff(l.staffName)) {
          const key = nameKey(name);
          if (!key) continue;
          if (!teachers.byKey.has(key) && !teachers.byKey.has("short:" + shortKey(key))) unmatched.add(name);
          docs.push({ date, gradeName, lesson: Number(l.lesson) || null, start: bell.start, end: bell.end, subject: String(l.subjectName || "").slice(0, 120), room: String(l.roomName || "").slice(0, 40), staffName: name.slice(0, 120), key, source });
        }
      }
    }
    // заменяем только присланные группы и недели — остальное расписание не трогаем
    for (const c of covered) await StaffLesson.deleteMany({ date: c.date, gradeName: c.gradeName });
    if (docs.length) await StaffLesson.insertMany(docs);
    const rec = await StaffImport.create({ source, weeks: [...weeks].sort(), groups: groups.length, lessons: docs.length, unmatched: [...unmatched].sort(), by: by || "" });
    return { groups: groups.length, lessons: docs.length, weeks: rec.weeks, unmatched: rec.unmatched };
  }

  // Excel из раздела «Отчёты» редактора: A2 «с 13.10.2026 по 18.10.2026», строки «Группа» → название → шапка → уроки
  function parseReportWorkbook(buf) {
    const wb = xlsx.read(buf, { type: "buffer" });
    const groups = [];
    for (const sn of wb.SheetNames) {
      const rows = xlsx.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: false, defval: "" });
      let weekStart = null;
      for (const r of rows.slice(0, 6)) {
        const m = /(\d{2})\.(\d{2})\.(\d{4})/.exec(r.join(" "));
        if (m) { weekStart = `${m[3]}-${m[2]}-${m[1]}`; break; }
      }
      if (!weekStart) continue;
      const wd = (parseKey(weekStart).getUTCDay() + 6) % 7; // привести к понедельнику
      weekStart = addDays(weekStart, -wd);
      let cur = null, expectName = false, lastDay = null;
      for (const r of rows) {
        const a = String(r[0] || "").trim();
        if (/^группа$/i.test(a)) { expectName = true; continue; }
        if (expectName && a) { cur = { gradeName: a, weekStart, lessons: [] }; groups.push(cur); expectName = false; lastDay = null; continue; }
        if (!cur || /^дни$/i.test(a)) continue;
        const day = DAY_NAMES[a.toLowerCase()] || (a ? null : lastDay);
        if (a && DAY_NAMES[a.toLowerCase()]) lastDay = day;
        if (!day || !parseBell(r[2])) continue;
        cur.lessons.push({ weekday: day, lesson: r[1], bell: r[2], subjectName: r[3], staffName: r[4], roomName: r[5] });
      }
    }
    return groups.filter((g) => g.lessons.length);
  }

  /* ---------- сопоставление с сотрудниками СКУД ---------- */
  async function staffGroups() {
    const gs = await Group.find().lean();
    return { teachers: gs.filter(isKind("teachers")), masters: gs.filter(isKind("masters")) };
  }
  async function teacherIndex() {
    const { teachers } = await staffGroups();
    const people = teachers.length ? await Student.find({ group: { $in: teachers.map((g) => g._id) } }).sort({ fio: 1 }).select("fio inn lyceumId group skudDaily attendance").lean() : [];
    const byKey = new Map();
    for (const p of people) byKey.set(nameKey(p.fio), p);
    // короткий ключ — только если он однозначный
    const shortCount = new Map();
    for (const p of people) { const k = "short:" + shortKey(nameKey(p.fio)); shortCount.set(k, (shortCount.get(k) || 0) + 1); }
    for (const p of people) { const k = "short:" + shortKey(nameKey(p.fio)); if (shortCount.get(k) === 1) byKey.set(k, p); }
    for (const a of await StaffAlias.find().lean()) {
      const p = people.find((x) => String(x._id) === String(a.staff));
      if (p) byKey.set(a.key, p);
    }
    return { people, byKey, groups: teachers };
  }

  /* ---------- расчёт дня ---------- */
  function teacherDay(p, date, lessons, reason) {
    const sk = p.skudDaily?.find((r) => r.date === date) || {};
    const present = !!(sk.present || p.attendance?.find((a) => a.date === date)?.present);
    const tin = localTime(sk.firstIn), tout = localTime(sk.lastOut || sk.lastIn);
    const inMin = tin?.min ?? null;
    const outMin = tout && (inMin == null || tout.min > inMin) ? tout.min : null;
    const first = lessons.length ? Math.min(...lessons.map((l) => toMin(l.start))) : null;
    const last = lessons.length ? Math.max(...lessons.map((l) => toMin(l.end))) : null;
    const grace = LATE_GRACE();
    const late = first != null && inMin != null && inMin - first > grace ? inMin - first : 0;
    const early = last != null && outMin != null && last - outMin > grace ? last - outMin : 0;
    let status;
    if (!lessons.length) status = present ? "extra" : "free"; // без уроков: пришёл / выходной по расписанию
    else if (!present) status = reason ? "excused" : "absent";
    else status = late ? "late" : "ok";
    return {
      id: String(p._id), fio: p.fio, date, status,
      lessons: lessons.sort((a, b) => toMin(a.start) - toMin(b.start)).map((l) => ({ lesson: l.lesson, start: l.start, end: l.end, group: l.gradeName, subject: l.subject, room: l.room })),
      first: fromMin(first), last: fromMin(last), in: fromMin(inMin), out: fromMin(outMin),
      late, early, minutes: inMin != null && outMin != null ? outMin - inMin : 0,
      reason: reason ? reason.label : null,
    };
  }
  function masterDay(p, date, sat, reason) {
    const sk = p.skudDaily?.find((r) => r.date === date) || {};
    const present = !!(sk.present || p.attendance?.find((a) => a.date === date)?.present);
    const tin = localTime(sk.firstIn), tout = localTime(sk.lastOut || sk.lastIn);
    const outMin = tout && (!tin || tout.min > tin.min) ? tout.min : null;
    const workday = !isWeekend(date, sat);
    return {
      id: String(p._id), fio: p.fio, date, workday,
      status: present ? "ok" : !workday ? "free" : reason ? "excused" : "absent",
      in: tin?.str || "", out: fromMin(outMin), minutes: tin && outMin != null ? outMin - tin.min : 0, reason: reason ? reason.label : null,
    };
  }

  async function lessonsBy(from, to) {
    const map = new Map(); // key|date → [lessons]
    for (const l of await StaffLesson.find({ date: { $gte: from, $lte: to } }).lean()) {
      for (const k of [`${l.key}|${l.date}`, `short:${shortKey(l.key)}|${l.date}`]) (map.get(k) || map.set(k, []).get(k)).push(l);
    }
    return map;
  }
  // ключи расписания, которые относятся к человеку (своё ФИО + ручные сопоставления)
  async function keysByPerson(people) {
    const out = new Map(people.map((p) => [String(p._id), new Set([nameKey(p.fio)])]));
    const shortCount = new Map();
    for (const p of people) { const k = shortKey(nameKey(p.fio)); shortCount.set(k, (shortCount.get(k) || 0) + 1); }
    for (const p of people) { const k = shortKey(nameKey(p.fio)); if (shortCount.get(k) === 1) out.get(String(p._id)).add("short:" + k); }
    for (const a of await StaffAlias.find().lean()) out.get(String(a.staff))?.add(a.key);
    return out;
  }
  // уроки человека за день без повторов (урок может найтись и по полному, и по короткому ключу)
  const lessonsFor = (map, keys, date) => {
    const seen = new Set();
    return [...(keys || [])].flatMap((k) => map.get(`${k}|${date}`) || []).filter((l) => !seen.has(String(l._id)) && seen.add(String(l._id)));
  };

  async function day(date) {
    const { teachers, masters } = await staffGroups();
    const tPeople = teachers.length ? await Student.find({ group: { $in: teachers.map((g) => g._id) } }).sort({ fio: 1 }).lean() : [];
    const mPeople = masters.length ? await Student.find({ group: { $in: masters.map((g) => g._id) } }).sort({ fio: 1 }).lean() : [];
    const map = await lessonsBy(date, date);
    const keys = await keysByPerson(tPeople);
    const reasons = await reasonsIndex({ student: { $in: [...tPeople, ...mPeople].map((p) => p._id) } }, date, date);
    const masterSat = masters.some((g) => g.saturday);
    const scheduleLoaded = (await StaffLesson.countDocuments({ date })) > 0;
    return {
      date, scheduleLoaded,
      teachers: tPeople.map((p) => teacherDay(p, date, lessonsFor(map, keys.get(String(p._id)), date), reasons.get(`${p._id}|${date}`))),
      masters: mPeople.map((p) => masterDay(p, date, masterSat, reasons.get(`${p._id}|${date}`))),
    };
  }

  async function month(monthKey) {
    const { from, to } = monthRange(monthKey);
    const today = getTodayKey();
    const days = datesBetween(from, to);
    const { teachers, masters } = await staffGroups();
    const tPeople = teachers.length ? await Student.find({ group: { $in: teachers.map((g) => g._id) } }).sort({ fio: 1 }).lean() : [];
    const mPeople = masters.length ? await Student.find({ group: { $in: masters.map((g) => g._id) } }).sort({ fio: 1 }).lean() : [];
    const map = await lessonsBy(from, to);
    const keys = await keysByPerson(tPeople);
    const reasons = await reasonsIndex({ student: { $in: [...tPeople, ...mPeople].map((p) => p._id) } }, from, to);
    const masterSat = masters.some((g) => g.saturday);
    const past = days.filter((d) => d <= today);
    const teacherRows = tPeople.map((p) => {
      const ds = past.map((d) => teacherDay(p, d, lessonsFor(map, keys.get(String(p._id)), d), reasons.get(`${p._id}|${d}`)));
      const withLessons = ds.filter((d) => d.lessons.length);
      return {
        id: String(p._id), fio: p.fio, days: ds,
        totals: {
          lessonDays: withLessons.length,
          attended: withLessons.filter((d) => d.status === "ok" || d.status === "late").length,
          absent: ds.filter((d) => d.status === "absent").length,
          excused: ds.filter((d) => d.status === "excused").length,
          late: ds.filter((d) => d.late).length,
          lateMin: ds.reduce((a, d) => a + d.late, 0),
          early: ds.filter((d) => d.early).length,
          earlyMin: ds.reduce((a, d) => a + d.early, 0),
          minutes: ds.reduce((a, d) => a + d.minutes, 0),
        },
      };
    });
    const masterRows = mPeople.map((p) => {
      const ds = past.map((d) => masterDay(p, d, masterSat, reasons.get(`${p._id}|${d}`)));
      return {
        id: String(p._id), fio: p.fio, days: ds,
        totals: { workdays: ds.filter((d) => d.workday).length, present: ds.filter((d) => d.status === "ok").length, absent: ds.filter((d) => d.status === "absent").length, excused: ds.filter((d) => d.status === "excused").length, minutes: ds.reduce((a, d) => a + d.minutes, 0) },
      };
    });
    return { month: monthKey, days, teachers: teacherRows, masters: masterRows };
  }

  const hm = (min) => (min ? `${Math.floor(min / 60)}:${pad2(min % 60)}` : "");
  function monthWorkbook(m) {
    const wb = xlsx.utils.book_new();
    const dayHead = m.days.map((d) => `${+d.slice(8)} ${WD[parseKey(d).getUTCDay()]}`);
    const cellT = (d) => {
      if (!d) return "";
      if (d.status === "absent") return "Н";
      if (d.status === "excused") return "У";
      if (d.status === "free") return "";
      let s = d.in || "+";
      if (d.late) s += ` +${d.late}`;
      if (d.early) s += ` ↑${d.early}`;
      return s;
    };
    const t = [
      [`Табель преподавателей — ${m.month}`],
      ["Время прихода; +N — опоздание к первому уроку (мин); ↑N — ушёл раньше конца последнего урока (мин); Н — не пришёл в день с уроками; У — по уважительной причине"],
      [],
      ["№", "ФИО", ...dayHead, "Дней с уроками", "Пришёл", "Не пришёл", "Уваж.", "Опозданий", "Мин. опозд.", "Ранних уходов", "Мин. раньше", "Часов в лицее"],
      ...m.teachers.map((r, i) => {
        const byDate = Object.fromEntries(r.days.map((d) => [d.date, d]));
        const x = r.totals;
        return [i + 1, r.fio, ...m.days.map((d) => cellT(byDate[d])), x.lessonDays, x.attended, x.absent, x.excused, x.late, x.lateMin, x.early, x.earlyMin, hm(x.minutes)];
      }),
    ];
    const ws1 = xlsx.utils.aoa_to_sheet(t);
    ws1["!cols"] = [{ wch: 4 }, { wch: 30 }, ...m.days.map(() => ({ wch: 9 })), ...Array(9).fill({ wch: 11 })];
    xlsx.utils.book_append_sheet(wb, ws1, "Преподаватели");
    const cellM = (d) => (!d ? "" : d.status === "absent" ? "Н" : d.status === "excused" ? "У" : d.status === "free" ? "" : hm(d.minutes) || "+");
    const mm = [
      [`Табель мастеров — ${m.month}`],
      ["Часы в лицее за день; Н — не было в рабочий день; У — по уважительной причине"],
      [],
      ["№", "ФИО", ...dayHead, "Рабочих дней", "Был", "Не был", "Уваж.", "Часов в лицее"],
      ...m.masters.map((r, i) => {
        const byDate = Object.fromEntries(r.days.map((d) => [d.date, d]));
        const x = r.totals;
        return [i + 1, r.fio, ...m.days.map((d) => cellM(byDate[d])), x.workdays, x.present, x.absent, x.excused, hm(x.minutes)];
      }),
    ];
    const ws2 = xlsx.utils.aoa_to_sheet(mm);
    ws2["!cols"] = [{ wch: 4 }, { wch: 30 }, ...m.days.map(() => ({ wch: 7 })), ...Array(5).fill({ wch: 12 })];
    xlsx.utils.book_append_sheet(wb, ws2, "Мастера");
    return xlsx.write(wb, { type: "buffer", bookType: "xlsx" });
  }

  /* ---------- уведомления администратору ---------- */
  async function alertsTick(bot, { date = getTodayKey(), nowMin = getBishkekMinutes() } = {}) {
    const d = await day(date);
    if (!d.scheduleLoaded) return { sent: 0 };
    const admins = await TgSub.find({ role: "admin" }).lean();
    const msgs = [];
    for (const t of d.teachers) {
      if (!t.lessons.length) continue;
      const first = toMin(t.first);
      const l1 = t.lessons[0];
      const what = `${l1.group}${l1.subject ? ", " + l1.subject : ""}${l1.room ? ", каб. " + l1.room : ""}`;
      if (t.status === "absent" && nowMin >= first + 10 && !(await StaffAlertLog.findOne({ staff: t.id, date, type: "absent" }).lean())) {
        msgs.push({ id: t.id, type: "absent", text: `🔴 ${t.fio} не пришёл к первому уроку ${t.first} (${what}).` });
      } else if (t.late && !(await StaffAlertLog.findOne({ staff: t.id, date, type: "late" }).lean())) {
        msgs.push({ id: t.id, type: "late", text: `🟠 ${t.fio} опоздал на ${t.late} мин: урок в ${t.first}, вход в ${t.in} (${what}).` });
      }
    }
    for (const m of msgs) {
      await StaffAlertLog.create({ staff: m.id, date, type: m.type });
      for (const a of admins) await bot.sendMessage(a.chatId, m.text).catch(() => null);
    }
    return { sent: msgs.length };
  }
  async function eveningSummary(bot, date = getTodayKey()) {
    const d = await day(date);
    if (!d.scheduleLoaded) return { sent: 0 };
    const late = d.teachers.filter((t) => t.late), absent = d.teachers.filter((t) => t.status === "absent"), early = d.teachers.filter((t) => t.early);
    const mAbsent = d.masters.filter((m) => m.status === "absent");
    const withLessons = d.teachers.filter((t) => t.lessons.length).length;
    let text = `📋 Сотрудники за ${date}\nПреподавателей с уроками: ${withLessons}, пришли: ${d.teachers.filter((t) => t.status === "ok" || t.status === "late").length}`;
    if (absent.length) text += `\n\n🔴 Не пришли:\n` + absent.map((t) => `• ${t.fio} (уроки ${t.first}–${t.last})`).join("\n");
    if (late.length) text += `\n\n🟠 Опоздали:\n` + late.map((t) => `• ${t.fio} — ${t.late} мин (урок ${t.first}, вход ${t.in})`).join("\n");
    if (early.length) text += `\n\n🟡 Ушли раньше:\n` + early.map((t) => `• ${t.fio} — на ${t.early} мин (конец ${t.last}, выход ${t.out})`).join("\n");
    if (mAbsent.length) text += `\n\n🔧 Мастера, которых не было:\n` + mAbsent.map((m) => `• ${m.fio}`).join("\n");
    if (!absent.length && !late.length && !early.length && !mAbsent.length) text += "\n\n✅ Без нарушений";
    const admins = await TgSub.find({ role: "admin" }).lean();
    for (let i = 0; i < text.length; i += 3900) for (const a of admins) await bot.sendMessage(a.chatId, text.slice(i, i + 3900)).catch(() => null);
    return { sent: admins.length };
  }
  function schedule(bot) {
    if (String(process.env.STAFF_ALERTS || "").toLowerCase() === "off") return;
    const tz = { timezone: "Asia/Bishkek" };
    let busy = false;
    cron.schedule("*/5 7-17 * * 1-6", async () => {
      if (busy) return;
      busy = true;
      try { await alertsTick(bot); } catch (e) { console.error("staff alerts:", e.message); } finally { busy = false; }
    }, tz);
    cron.schedule("30 18 * * 1-6", () => eveningSummary(bot).catch((e) => console.error("staff summary:", e.message)), tz);
    console.log("👩‍🏫 Контроль преподавателей: уведомления админу об опозданиях, итог дня в 18:30");
  }

  /* ---------- маршруты ---------- */
  function registerRoutes(app, { requireAdmin, requireAuth, safeEq, contentDisposition }) {
    const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 8e6, files: 10 } });
    const wrap = (fn) => (req, res) => Promise.resolve(fn(req, res)).catch((e) => res.status(e.status || 400).json({ error: e.message }));

    app.get("/staff", requireAuth, requireAdmin, (_req, res) => res.sendFile(path.join(__dirname, "public", "staff.html")));

    // Приём из редактора расписания: заголовок X-Api-Key = SCHEDULE_API_KEY
    app.post("/api/agent/teacher-schedule", wrap(async (req, res) => {
      const key = process.env.SCHEDULE_API_KEY;
      if (!key) return res.status(503).json({ error: "На сервере не задан SCHEDULE_API_KEY" });
      if (!safeEq(req.headers["x-api-key"] || req.body?.apiKey, key)) return res.status(403).json({ error: "Неверный ключ" });
      res.json({ ok: true, ...(await importSchedule(req.body?.groups, { source: "kundoluk", by: "редактор расписания" })) });
    }));

    app.post("/api/admin/staff/schedule/upload", requireAdmin, upload.array("files", 10), wrap(async (req, res) => {
      const groups = (req.files || []).flatMap((f) => parseReportWorkbook(f.buffer));
      if (!groups.length) return res.status(400).json({ error: "В файле не найдено расписание. Нужен Excel из раздела «Отчёты» редактора расписания." });
      res.json({ ok: true, ...(await importSchedule(groups, { source: "excel", by: req.auth?.username })) });
    }));

    app.get("/api/admin/staff/schedule/status", requireAdmin, wrap(async (_req, res) => {
      const imports = await StaffImport.find().sort({ createdAt: -1 }).limit(10).lean();
      const { people } = await teacherIndex();
      const aliases = await StaffAlias.find().lean();
      // все ФИО из расписания за 6 недель, которые не нашлись среди сотрудников СКУД
      const from = addDays(getTodayKey(), -21), to = addDays(getTodayKey(), 21);
      const names = await StaffLesson.find({ date: { $gte: from, $lte: to } }).select("key staffName").lean();
      const idx = await teacherIndex();
      const known = (k) => idx.byKey.has(k) || idx.byKey.has("short:" + shortKey(k));
      const unmatched = [...new Map(names.filter((n) => !known(n.key)).map((n) => [n.key, n.staffName])).entries()].map(([key, name]) => ({ key, name })).sort((a, b) => a.name.localeCompare(b.name, "ru"));
      const range = await StaffLesson.find().sort({ date: -1 }).limit(1).lean();
      res.json({
        imports, unmatched, lastDate: range[0]?.date || null,
        teachers: people.map((p) => ({ id: String(p._id), fio: p.fio })),
        aliases: aliases.map((a) => ({ key: a.key, name: a.name, staff: String(a.staff), fio: people.find((p) => String(p._id) === String(a.staff))?.fio || "?" })),
        keyConfigured: !!process.env.SCHEDULE_API_KEY,
      });
    }));
    app.post("/api/admin/staff/alias", requireAdmin, wrap(async (req, res) => {
      const { name, staffId } = req.body || {};
      const key = nameKey(name);
      if (!key || !mongoose.isValidObjectId(staffId)) return res.status(400).json({ error: "Укажите ФИО и сотрудника" });
      if (!(await Student.findById(staffId).select("_id").lean())) return res.status(404).json({ error: "Сотрудник не найден" });
      await StaffAlias.findOneAndUpdate({ key }, { $set: { key, name: String(name).slice(0, 120), staff: staffId } }, { upsert: true });
      res.json({ ok: true });
    }));
    app.delete("/api/admin/staff/alias/:key", requireAdmin, wrap(async (req, res) => {
      await StaffAlias.deleteOne({ key: req.params.key });
      res.json({ ok: true });
    }));

    app.get("/api/admin/staff/day", requireAdmin, wrap(async (req, res) => {
      const date = isValidKey(req.query.date) ? req.query.date : getTodayKey();
      res.json(await day(date));
    }));
    app.get("/api/admin/staff/month", requireAdmin, wrap(async (req, res) => {
      const m = /^\d{4}-\d{2}$/.test(String(req.query.month)) ? req.query.month : getTodayKey().slice(0, 7);
      const data = await month(m);
      if (req.query.format !== "xlsx") return res.json(data);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", contentDisposition(`Tabel_sotrudnikov_${m}.xlsx`));
      res.send(monthWorkbook(data));
    }));
  }

  return { importSchedule, parseReportWorkbook, day, month, monthWorkbook, alertsTick, eveningSummary, schedule, registerRoutes };
}

module.exports = { createStaffAttendance, nameKey, parseBell, splitStaff, StaffLesson, StaffAlias };
