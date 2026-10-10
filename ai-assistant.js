// ИИ-помощник в Telegram-боте для администратора и кураторов.
// Вопросы обычным языком («кто из ИТ-21 пропустил больше 5 дней в октябре?»).
// Модель отвечает ТОЛЬКО через инструменты ниже: они читают базу с учётом прав
// (куратор видит только свои группы) и ничего не меняют.
// Нужен ANTHROPIC_API_KEY; модель — AI_MODEL (по умолчанию claude-haiku-5-5).

const { isValidKey, weekRange, monthRange, addDays, shiftMinutes, dayStatus, isWeekend, datesBetween } = require("./attendance-utils");
const { reasonsIndex } = require("./absence");

const API_URL = "https://api.anthropic.com/v1/messages";
const MAX_STEPS = 6;
const DAILY_LIMIT = Number(process.env.AI_DAILY_LIMIT) || 40; // вопросов в день на одного пользователя
const HISTORY_TURNS = 3;

const TOOLS = [
  {
    name: "list_groups",
    description: "Список групп, доступных пользователю: название, профессия, число студентов, начало занятий, учебная ли суббота.",
    input_schema: { type: "object", properties: {} },
  },
  {
    name: "group_attendance",
    description: "Посещаемость одной группы за период: итоги группы и по каждому студенту (дней был, опозданий, пропусков по уважительной и без, % посещаемости). Учебные дни уже учтены (выходные и будущие дни не считаются).",
    input_schema: {
      type: "object",
      properties: {
        group: { type: "string", description: "Название группы, например ИТ-21" },
        from: { type: "string", description: "Начало периода ГГГГ-ММ-ДД" },
        to: { type: "string", description: "Конец периода ГГГГ-ММ-ДД" },
      },
      required: ["group", "from", "to"],
    },
  },
  {
    name: "lyceum_attendance",
    description: "Сводка посещаемости всех доступных групп за период (по группе: % посещаемости, опоздания, пропуски по уважительной/без, студентов с 3+ пропусками без причины). Отсортировано от худшей к лучшей.",
    input_schema: {
      type: "object",
      properties: { from: { type: "string" }, to: { type: "string" } },
      required: ["from", "to"],
    },
  },
  {
    name: "today_status",
    description: "Кто сегодня пришёл, опоздал (со временем входа) и отсутствует (с причиной, если указана) в группе.",
    input_schema: { type: "object", properties: { group: { type: "string" } }, required: ["group"] },
  },
  {
    name: "find_student",
    description: "Поиск студента по части ФИО. Возвращает до 10 совпадений с id и группой.",
    input_schema: { type: "object", properties: { query: { type: "string" } }, required: ["query"] },
  },
  {
    name: "student_attendance",
    description: "Посещаемость студента по дням за период: был/опоздал (время)/пропуск (причина).",
    input_schema: {
      type: "object",
      properties: { student_id: { type: "string" }, from: { type: "string" }, to: { type: "string" } },
      required: ["student_id", "from", "to"],
    },
  },
  {
    name: "student_grades",
    description: "Оценки студента из сводной ведомости за месяц (ГГГГ-ММ).",
    input_schema: { type: "object", properties: { student_id: { type: "string" }, month: { type: "string" } }, required: ["student_id", "month"] },
  },
  {
    name: "send_excel_report",
    description: "Отправить пользователю Excel-отчёт посещаемости группы (или всех групп — group='all', только админ) за неделю или месяц. Используй, когда просят файл/таблицу/Excel.",
    input_schema: {
      type: "object",
      properties: {
        group: { type: "string", description: "Название группы или 'all'" },
        period: { type: "string", enum: ["week", "month"] },
        date: { type: "string", description: "Любой день нужной недели или месяца, ГГГГ-ММ-ДД" },
      },
      required: ["group", "period", "date"],
    },
  },
];

const WD = ["вс", "пн", "вт", "ср", "чт", "пт", "сб"];

function createAssistant({ Group, Student, Curator, SummaryRecord, ensureGroupSummarySubjects, getTodayKey, reports, callModel }) {
  const apiKey = () => process.env.ANTHROPIC_API_KEY;
  const enabled = () => !!(apiKey() || callModel);
  const model = () => process.env.AI_MODEL || "claude-haiku-5-5";

  const usage = new Map(); // chatId → {day, n}
  const history = new Map(); // chatId → [{role, content}]

  async function anthropic(body) {
    if (callModel) return callModel(body);
    const r = await fetch(API_URL, {
      method: "POST",
      headers: { "content-type": "application/json", "x-api-key": apiKey(), "anthropic-version": "2023-06-01" },
      body: JSON.stringify(body),
    });
    const j = await r.json().catch(() => ({}));
    if (!r.ok) throw new Error(j?.error?.message || `Claude API: ${r.status}`);
    return j;
  }

  // ── доступ ──
  async function groupsFor(ctx) {
    const all = (await Group.find().sort({ name: 1 }).lean()).filter((g) => !g.kind || g.kind === "students");
    if (ctx.role === "admin") return all;
    const ids = (ctx.groupIds || []).map(String);
    return all.filter((g) => ids.includes(String(g._id)));
  }
  const norm = (s) => String(s || "").toLowerCase().replace(/[\s\-–_]/g, "");
  async function findGroup(ctx, name) {
    const gs = await groupsFor(ctx);
    const n = norm(name);
    return gs.find((g) => norm(g.name) === n || String(g._id) === String(name)) || gs.find((g) => norm(g.name).includes(n)) || null;
  }
  async function studentFor(ctx, id) {
    const s = await Student.findById(id).select("fio group skudDaily attendance").lean();
    if (!s) return null;
    const gs = await groupsFor(ctx);
    const g = gs.find((x) => String(x._id) === String(s.group));
    return g ? { s, g } : null;
  }
  const checkRange = (from, to) => {
    if (!isValidKey(from) || !isValidKey(to) || to < from) throw new Error("Неверный период: нужен формат ГГГГ-ММ-ДД");
    if (datesBetween(from, to).length > 370) throw new Error("Период больше года");
  };

  // ── инструменты ──
  const tools = {
    async list_groups(ctx) {
      const gs = await groupsFor(ctx);
      const out = [];
      for (const g of gs) out.push({ name: g.name, profession: g.profRu || "", students: await Student.countDocuments({ group: g._id }), shiftStart: g.shiftStart || "", saturday: !!g.saturday });
      return out;
    },
    async group_attendance(ctx, { group, from, to }) {
      checkRange(from, to);
      const g = await findGroup(ctx, group);
      if (!g) return { error: `Группа «${group}» не найдена или нет доступа` };
      const rep = await reports.buildGroupReport(g, { period: "custom", from, to, title: `${from}..${to}` });
      return {
        group: g.name, from, to, workdays: rep.workdays, totals: rep.totals,
        students: rep.rows.map((r) => ({ id: r.id, fio: r.fio, present: r.present, late: r.late, excused: r.excused, unexcused: r.unexcused, pending: r.pending, pct: r.pct })),
      };
    },
    async lyceum_attendance(ctx, { from, to }) {
      checkRange(from, to);
      const out = [];
      for (const g of await groupsFor(ctx)) {
        const rep = await reports.buildGroupReport(g, { period: "custom", from, to, title: "" });
        if (!rep.rows.length) continue;
        out.push({ group: g.name, students: rep.totals.students, workdays: rep.workdays, pct: rep.totals.pct, late: rep.totals.late, excused: rep.totals.excused, unexcused: rep.totals.unexcused, attention: rep.attention.length });
      }
      return out.sort((a, b) => a.pct - b.pct);
    },
    async today_status(ctx, { group }) {
      const g = await findGroup(ctx, group);
      if (!g) return { error: `Группа «${group}» не найдена или нет доступа` };
      const tdk = getTodayKey();
      if (isWeekend(tdk, !!g.saturday)) return { group: g.name, date: tdk, note: "Сегодня выходной для этой группы" };
      const shiftMin = shiftMinutes(g.shiftStart);
      const studs = await Student.find({ group: g._id }).sort({ fio: 1 }).select("fio skudDaily attendance").lean();
      const reasons = await reasonsIndex({ group: g._id }, tdk, tdk);
      const present = [], late = [], absent = [];
      for (const s of studs) {
        const st = dayStatus(s, tdk, shiftMin);
        if (st.late) late.push(`${s.fio} (${st.time})`);
        else if (st.present) present.push(s.fio);
        else { const r = reasons.get(`${s._id}|${tdk}`); absent.push(r ? `${s.fio} — ${r.label}${r.confirmed ? "" : " (не подтв.)"}` : s.fio); }
      }
      return { group: g.name, date: tdk, shiftStart: g.shiftStart, presentCount: present.length + late.length, total: studs.length, late, absent };
    },
    async find_student(ctx, { query }) {
      const q = String(query || "").trim();
      if (q.length < 2) return { error: "Слишком короткий запрос" };
      const ids = (await groupsFor(ctx)).map((g) => String(g._id));
      const rx = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "i");
      const studs = await Student.find({ fio: rx }).select("fio group").limit(50).lean();
      const gs = Object.fromEntries((await groupsFor(ctx)).map((g) => [String(g._id), g.name]));
      return studs.filter((s) => ids.includes(String(s.group))).slice(0, 10).map((s) => ({ id: String(s._id), fio: s.fio, group: gs[String(s.group)] }));
    },
    async student_attendance(ctx, { student_id, from, to }) {
      checkRange(from, to);
      const x = await studentFor(ctx, student_id);
      if (!x) return { error: "Студент не найден или нет доступа" };
      const shiftMin = shiftMinutes(x.g.shiftStart), tdk = getTodayKey();
      const reasons = await reasonsIndex({ student: x.s._id }, from, to);
      const days = [];
      for (const d of datesBetween(from, to)) {
        if (isWeekend(d, !!x.g.saturday) || d > tdk) continue;
        const st = dayStatus(x.s, d, shiftMin);
        const r = !st.present && reasons.get(`${x.s._id}|${d}`);
        days.push({ date: d, wd: WD[new Date(d + "T00:00:00Z").getUTCDay()], status: st.present ? (st.late ? `опоздал ${st.time}` : `был ${st.time || ""}`.trim()) : r ? `пропуск: ${r.label}${r.confirmed ? "" : " (не подтв.)"}` : "пропуск без причины" });
      }
      return { fio: x.s.fio, group: x.g.name, days };
    },
    async student_grades(ctx, { student_id, month }) {
      if (!/^\d{4}-\d{2}$/.test(String(month))) return { error: "Месяц в формате ГГГГ-ММ" };
      const x = await studentFor(ctx, student_id);
      if (!x) return { error: "Студент не найден или нет доступа" };
      const rec = await SummaryRecord.findOne({ group: x.g._id, student: x.s._id, month }).lean();
      const raw = rec?.grades instanceof Map ? Object.fromEntries(rec.grades) : rec?.grades || {};
      const subjects = await ensureGroupSummarySubjects(await Group.findById(x.g._id));
      return { fio: x.s.fio, group: x.g.name, month, grades: subjects.filter((s) => raw[s.key]).map((s) => ({ subject: s.label, grade: raw[s.key] })) };
    },
    async send_excel_report(ctx, { group, period, date }, files) {
      const range = reports.resolvePeriod(period === "month" ? "month" : "week", date);
      if (String(group).toLowerCase() === "all") {
        if (ctx.role !== "admin") return { error: "Общий отчёт доступен только администратору" };
        const reps = [];
        for (const g of await groupsFor(ctx)) reps.push(await reports.buildGroupReport(g, range));
        files.push({ buf: reports.allWorkbook(reps.filter((r) => r.rows.length), range), filename: `Otchet_vse_gruppy_${range.from}_${range.to}.xlsx` });
        return { ok: true, file: "отчёт по всем группам", period: range.title };
      }
      const g = await findGroup(ctx, group);
      if (!g) return { error: `Группа «${group}» не найдена или нет доступа` };
      const rep = await reports.buildGroupReport(g, range);
      files.push({ buf: reports.groupWorkbook(rep), filename: `Otchet_${g.name}_${range.from}_${range.to}.xlsx` });
      return { ok: true, file: `отчёт по группе ${g.name}`, period: range.title };
    },
  };

  function systemPrompt(ctx, groups) {
    const t = getTodayKey();
    const w = weekRange(t), m = monthRange(t.slice(0, 7));
    return [
      "Ты — помощник по посещаемости и успеваемости Профессионального лицея №3 (г. Кант). Отвечаешь в Telegram сотруднику лицея.",
      `Сегодня ${t} (${WD[new Date(t + "T00:00:00Z").getUTCDay()]}). Эта неделя: ${w.from}..${w.to}. Этот месяц: ${m.from}..${m.to}. Прошлая неделя: ${addDays(w.from, -7)}..${addDays(w.from, -1)}.`,
      ctx.role === "admin" ? "Пользователь — администратор, видит все группы." : `Пользователь — куратор, видит только свои группы: ${groups.map((g) => g.name).join(", ") || "нет групп"}.`,
      "Правила:",
      "- Факты бери ТОЛЬКО из инструментов. Не придумывай цифры, имена и даты. Если данных нет — так и скажи.",
      "- Если группа или студент не указаны однозначно — сначала найди через list_groups / find_student.",
      "- Отвечай по-русски, кратко и по делу: сначала ответ, потом детали. Списки — короткие строки с «•». Без таблиц и без Markdown-разметки (*, #, _) — это обычный текст Telegram.",
      "- «Пропуск без причины» и «по уважительной причине» — разные вещи, не смешивай. Проценты округляй до целых.",
      "- Если просят файл/Excel/таблицу — вызови send_excel_report и скажи, что файл ниже.",
      "- Ты не можешь ничего менять в базе; на просьбы изменить данные отвечай, что это делается в журнале или мини-приложении.",
      "- На вопросы не про лицей, посещаемость, оценки и студентов вежливо отвечай, что помогаешь только с этим.",
    ].join("\n");
  }

  function takeQuota(chatId) {
    const day = getTodayKey();
    const u = usage.get(chatId);
    if (!u || u.day !== day) { usage.set(chatId, { day, n: 1 }); return true; }
    if (u.n >= DAILY_LIMIT) return false;
    u.n++;
    return true;
  }

  // Главный вход: вопрос → {text, files}
  async function ask(chatId, ctx, question) {
    if (!enabled()) return { text: "ИИ-помощник не настроен: администратору нужно добавить ANTHROPIC_API_KEY в настройки сервера.", files: [] };
    if (!["admin", "curator"].includes(ctx.role)) return { text: "ИИ-помощник доступен администратору и кураторам.", files: [] };
    if (!takeQuota(String(chatId))) return { text: `Лимит ${DAILY_LIMIT} вопросов в день исчерпан. Попробуйте завтра.`, files: [] };
    const groups = await groupsFor(ctx);
    const files = [];
    const past = history.get(String(chatId)) || [];
    const messages = [...past, { role: "user", content: String(question).slice(0, 1500) }];
    let answer = "";
    for (let step = 0; step < MAX_STEPS; step++) {
      const res = await anthropic({ model: model(), max_tokens: 1200, system: systemPrompt(ctx, groups), tools: TOOLS, messages });
      const blocks = res.content || [];
      messages.push({ role: "assistant", content: blocks });
      const calls = blocks.filter((b) => b.type === "tool_use");
      answer = blocks.filter((b) => b.type === "text").map((b) => b.text).join("\n").trim();
      if (res.stop_reason !== "tool_use" || !calls.length) break;
      const results = [];
      for (const c of calls) {
        let out;
        try {
          const fn = tools[c.name];
          out = fn ? await fn(ctx, c.input || {}, files) : { error: "Нет такого инструмента" };
        } catch (e) {
          out = { error: e.message };
        }
        results.push({ type: "tool_result", tool_use_id: c.id, content: JSON.stringify(out).slice(0, 60000) });
      }
      messages.push({ role: "user", content: results });
      if (step === MAX_STEPS - 1) answer = answer || "Не удалось собрать ответ за отведённое число шагов. Уточните вопрос.";
    }
    answer = answer || (files.length ? "Готово, файл ниже." : "Не получилось ответить. Переформулируйте вопрос.");
    // короткая история только из текста — для уточняющих вопросов
    const turns = [...past, { role: "user", content: String(question).slice(0, 1500) }, { role: "assistant", content: answer.slice(0, 2000) }];
    history.set(String(chatId), turns.slice(-HISTORY_TURNS * 2));
    return { text: answer, files };
  }

  const reset = (chatId) => history.delete(String(chatId));
  return { ask, reset, enabled, tools, TOOLS, systemPrompt };
}

module.exports = { createAssistant, TOOLS };
