// Общие расчёты посещаемости: статус дня, опоздание, выходные, периоды.
// Время СКУД хранится в ISO (UTC); лицей живёт по Бишкеку (UTC+6).

const pad2 = (n) => String(n).padStart(2, "0");

// ISO → минуты от начала суток по Бишкеку и строка "ЧЧ:ММ"
function localTime(iso) {
  if (!iso) return null;
  const dt = new Date(iso);
  if (isNaN(dt.getTime())) return null;
  const min = (dt.getUTCHours() * 60 + dt.getUTCMinutes() + 360) % 1440;
  return { min, str: `${pad2(Math.floor(min / 60))}:${pad2(min % 60)}` };
}

function shiftMinutes(shiftStart) {
  const [h, m] = String(shiftStart || "09:30").split(":").map(Number);
  return (h || 0) * 60 + (m || 0);
}

// Статус студента за день: { present, late, time }
function dayStatus(stu, dk, shiftMin) {
  const skud = stu.skudDaily?.find((r) => r.date === dk);
  const att = stu.attendance?.find((a) => a.date === dk);
  const present = !!(skud?.present || att?.present);
  let late = false,
    time = "";
  const t = localTime(skud?.firstIn);
  if (t) {
    time = t.str;
    late = present && t.min > shiftMin;
  }
  return { present, late, time };
}

const dateKey = (d) => `${d.getUTCFullYear()}-${pad2(d.getUTCMonth() + 1)}-${pad2(d.getUTCDate())}`;
const parseKey = (dk) => {
  const [y, m, d] = String(dk).split("-").map(Number);
  return new Date(Date.UTC(y, m - 1, d));
};
// строгая проверка: 2026-02-30 и 2026-13-01 не проходят
const isValidKey = (dk) => /^\d{4}-\d{2}-\d{2}$/.test(String(dk || "")) && dateKey(parseKey(dk)) === dk;
const isWeekend = (dk) => {
  const wd = parseKey(dk).getUTCDay();
  return wd === 0 || wd === 6;
};
function addDays(dk, n) {
  const d = parseKey(dk);
  d.setUTCDate(d.getUTCDate() + n);
  return dateKey(d);
}
function datesBetween(from, to) {
  const out = [];
  for (let dk = from; dk <= to && out.length < 400; dk = addDays(dk, 1)) out.push(dk);
  return out;
}
// Понедельник–воскресенье недели, в которую входит dk
function weekRange(dk) {
  const wd = (parseKey(dk).getUTCDay() + 6) % 7; // пн = 0
  const from = addDays(dk, -wd);
  return { from, to: addDays(from, 6) };
}
function monthRange(month) {
  const [y, m] = String(month).split("-").map(Number);
  const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return { from: `${y}-${pad2(m)}-01`, to: `${y}-${pad2(m)}-${pad2(dim)}` };
}

module.exports = {
  pad2,
  localTime,
  shiftMinutes,
  dayStatus,
  dateKey,
  parseKey,
  isValidKey,
  isWeekend,
  addDays,
  datesBetween,
  weekRange,
  monthRange,
};
