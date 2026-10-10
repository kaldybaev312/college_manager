// «Не пришёл»: через час после начала занятий группы (по умолчанию) родителю приходит сообщение,
// если ребёнка нет в лицее и причина отсутствия не указана; куратору — список отсутствующих.
// Проверка идёт каждые 10 минут с 7:00 до 16:00 (Бишкек), по каждой группе — один раз в день.
// Не тревожим, если: выходной для группы, студент на практике по графику, есть причина отсутствия,
// в группе сегодня не пришёл никто (праздник/выезд/сбой СКУД) или в лицее вообще нет ни одного входа.
// Настройки: ABSENT_ALERTS=off — выключить; ABSENT_ALERT_DELAY — минут после начала занятий (60).
const mongoose = require("mongoose");
const cron = require("node-cron");
const { shiftMinutes, dayStatus, isWeekend } = require("./attendance-utils");
const { reasonsIndex } = require("./absence");

const AlertLog =
  mongoose.models.AbsentAlertLog ||
  mongoose.model(
    "AbsentAlertLog",
    new mongoose.Schema({ group: String, date: String, sent: Number, at: { type: Date, default: Date.now, expires: 40 * 24 * 3600 } }, { versionKey: false }),
  );

const isStudentGroup = (g) => !g.kind || g.kind === "students";
const delayMin = () => {
  const n = parseInt(process.env.ABSENT_ALERT_DELAY, 10);
  return Number.isFinite(n) && n >= 0 && n <= 300 ? n : 60;
};
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function createAbsentAlerts({ Group, Student, Curator, TgSub, getTodayKey, getBishkekMinutes, practiceStudentIds }) {
  // Кто сейчас должен получить уведомления. nowMin — минуты от полуночи по Бишкеку.
  // force — без проверки времени и журнала отправок (для предпросмотра); ignoreTime — без проверки времени
  async function collect({ today = getTodayKey(), nowMin = getBishkekMinutes(), force = false, ignoreTime = false } = {}) {
    const groups = (await Group.find().sort({ name: 1 }).lean()).filter(isStudentGroup);
    const onPractice = new Set((await practiceStudentIds(today)).map(String));
    const plans = [];
    let lyceumPresent = 0;
    for (const g of groups) {
      if (isWeekend(today, !!g.saturday)) continue;
      const due = shiftMinutes(g.shiftStart) + delayMin();
      if (!force && !ignoreTime && nowMin < due) continue;
      if (!force && (await AlertLog.findOne({ group: String(g._id), date: today }).lean())) continue;
      const shiftMin = shiftMinutes(g.shiftStart);
      const studs = await Student.find({ group: g._id }).sort({ fio: 1 }).select("fio inn skudDaily attendance").lean();
      const reasons = await reasonsIndex({ group: g._id }, today, today);
      const absent = [];
      let present = 0;
      for (const s of studs) {
        if (onPractice.has(String(s._id))) continue;
        if (dayStatus(s, today, shiftMin).present) { present++; continue; }
        if (reasons.get(`${s._id}|${today}`)) continue;
        absent.push({ id: String(s._id), fio: s.fio, inn: s.inn || "" });
      }
      lyceumPresent += present;
      plans.push({ group: g, present, absent, skipped: present === 0 ? "в группе сегодня никто не пришёл" : null });
    }
    // Никто во всём лицее не прошёл через турникет — скорее праздник или сбой СКУД: молчим
    if (lyceumPresent === 0) plans.forEach((p) => (p.skipped = p.skipped || "в лицее нет ни одного входа"));
    return plans;
  }

  const parentText = (s, g) =>
    `❗ ${s.fio} сегодня не отмечен(а) на входе в лицей (группа ${g.name}, занятия с ${g.shiftStart || "—"}).\n\n` +
    `Если ребёнок отсутствует по уважительной причине — сообщите в мини-приложении: «Сообщить об отсутствии». ` +
    `Если он(а) в лицее — возможно, прошёл(а) без отметки, уточните у куратора.`;
  const curatorText = (g, list) =>
    `📋 ${g.name}: нет на занятиях без указанной причины — ${list.length}\n\n` +
    list.map((s, i) => `${i + 1}. ${s.fio}`).join("\n") +
    `\n\nРодителям отправлено уведомление. Причину можно указать в журнале или мини-приложении.`;

  async function run(bot, opts = {}) {
    const plans = await collect(opts);
    const result = { groups: 0, parents: 0, curators: 0, absent: 0, skipped: [] };
    for (const p of plans) {
      if (p.skipped) {
        // не отмечаем как «отправлено»: через 10 минут проверим снова (вдруг СКУД ещё не выгрузил входы)
        result.skipped.push({ group: p.group.name, why: p.skipped });
        continue;
      }
      result.groups++;
      result.absent += p.absent.length;
      if (opts.dry) {
        (result.preview ||= []).push({ group: p.group.name, absent: p.absent.map((x) => x.fio) });
        continue;
      }
      let sent = 0;
      for (const s of p.absent) {
        if (!s.inn) continue;
        const subs = await TgSub.find({ role: "parent", linkedInns: s.inn }).lean();
        for (const sub of subs) {
          await bot.sendMessage(sub.chatId, parentText(s, p.group)).then(() => { sent++; result.parents++; }).catch(() => null);
          await sleep(50);
        }
      }
      if (p.absent.length) {
        const curs = await Curator.find({ groups: p.group._id }).select("_id").lean();
        const subs = curs.length ? await TgSub.find({ role: "curator", curatorId: { $in: curs.map((c) => String(c._id)) } }).lean() : [];
        for (const sub of subs) await bot.sendMessage(sub.chatId, curatorText(p.group, p.absent)).then(() => result.curators++).catch(() => null);
      }
      if (!opts.force) await AlertLog.create({ group: String(p.group._id), date: opts.today || getTodayKey(), sent });
    }
    if (result.groups || result.skipped.length) console.log(`❗ «Не пришёл»: групп ${result.groups}, отсутствуют ${result.absent}, родителям ${result.parents}, кураторам ${result.curators}`);
    return result;
  }

  function schedule(bot) {
    if (String(process.env.ABSENT_ALERTS || "").toLowerCase() === "off") return;
    let busy = false;
    cron.schedule(
      "*/10 7-16 * * 1-6",
      async () => {
        if (busy) return;
        busy = true;
        try { await run(bot); } catch (e) { console.error("absent alerts:", e.message); } finally { busy = false; }
      },
      { timezone: "Asia/Bishkek" },
    );
    console.log(`❗ «Не пришёл»: родителям через ${delayMin()} мин после начала занятий группы`);
  }

  return { collect, run, schedule, parentText, curatorText };
}

module.exports = { createAbsentAlerts, AlertLog };
