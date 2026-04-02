/* ================================================================
   TELEGRAM BOT — ПЛ №3  v3.0
   Изоляция данных по ролям:
     admin   → всё
     curator → только своя группа, уведомления только своих
     parent  → только свои дети (по ИНН)
================================================================ */
"use strict";
const TelegramBot = require("node-telegram-bot-api");
const crypto      = require("crypto");

module.exports = function createBot(models, helpers, buildSummaryDocx) {
  const { Group, Student, TgSub, Curator, SummaryRecord } = models;
  const { pad2, getBishkekNow, getTodayKey, ensureGroupSummarySubjects, getSummaryDocumentData, getSkudEvents } = helpers;

  const TG_TOKEN     = process.env.TG_TOKEN;
  if (!TG_TOKEN) throw new Error("❌ TG_TOKEN не задан в .env!");
  const ADMIN_SECRET = process.env.ADMIN_TG_SECRET  || "pl3admin2026";
  const BOT_USERNAME = process.env.BOT_USERNAME     || "pl3_check_bot";

  const bot = new TelegramBot(TG_TOKEN, { polling: true });

  // Обработка ошибок поллинга (сеть, конфликт токенов и т.д.)
  bot.on("polling_error", (err) => {
    console.error("❌ TG Polling error:", err.code, err.message?.slice(0,120));
  });
  bot.on("error", (err) => {
    console.error("❌ TG Bot error:", err.message?.slice(0,120));
  });

  console.log("🤖 Bot polling started. Token:", TG_TOKEN.slice(0,10)+"...");

  /* ──────────────────────────────────────────────────────────────
     IN-MEMORY SESSION CACHE  (роль + рабочий контекст)
  ────────────────────────────────────────────────────────────── */
  const SESS = new Map(); // chatId → { role, name, groupIds?, inns?, step, ...fill }

  function S(id)           { return SESS.get(String(id)) || {}; }
  function set(id, patch)  { SESS.set(String(id), { ...S(id), ...patch }); }
  function clearStep(id)   {
    const s = S(id);
    SESS.set(String(id), {
      role: s.role, name: s.name, groupIds: s.groupIds,
      curatorId: s.curatorId, inns: s.inns,
    });
  }
  function evictCache(id) { SESS.delete(String(id)); }

  /* ──────────────────────────────────────────────────────────────
     HELPERS
  ────────────────────────────────────────────────────────────── */
  const MONTHS = ["Январь","Февраль","Март","Апрель","Май","Июнь",
                  "Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];

  function fmtM(key) {
    const [y,m] = (key||"").split("-").map(Number);
    return m ? `${MONTHS[m-1]} ${y}` : key||"—";
  }
  function last6() {
    const d = getBishkekNow();
    return Array.from({length:6}, (_,i)=>{
      const x = new Date(d.getFullYear(), d.getMonth()-i, 1);
      return `${x.getFullYear()}-${pad2(x.getMonth()+1)}`;
    });
  }
  function parseM(label) {
    const p = label.trim().split(" ");
    if (p.length<2) return null;
    const i = MONTHS.indexOf(p[0]);
    return i>=0 ? `${p[1]}-${pad2(i+1)}` : null;
  }
  function nowKey() { const d=getBishkekNow(); return `${d.getFullYear()}-${pad2(d.getMonth()+1)}`; }

  // Safe MarkdownV2 escape
  function ev(v) { return String(v??"").replace(/[_*[\]()~`>#+=|{}.!\-\\]/g,"\\$&"); }

  async function send(cid, txt, extra={}) {
    try { return await bot.sendMessage(cid, txt, {parse_mode:"MarkdownV2", ...extra}); }
    catch { try { return await bot.sendMessage(cid, txt.replace(/[_*[\]()~`>#+=|{}.!\-\\]/g,""), extra); } catch {} }
  }
  async function edit(cid, mid, txt, extra={}) {
    try { return await bot.editMessageText(txt, {chat_id:cid, message_id:mid, parse_mode:"MarkdownV2", ...extra}); } catch {}
  }
  async function ack(qid, txt="") { try { await bot.answerCallbackQuery(qid, {text:txt}); } catch {} }

  /* ──────────────────────────────────────────────────────────────
     KEYBOARDS
  ────────────────────────────────────────────────────────────── */
  const KB_ADMIN = { reply_markup:{ keyboard:[
    [{text:"📊 Статистика"},           {text:"📅 Посещаемость групп"}],
    [{text:"🔍 Найти студента"},        {text:"📝 Ведомость"}],
    [{text:"⏰ Список опозданий"},      {text:"🔗 Ссылки кураторов"}],
    [{text:"📱 Мини-приложение", web_app:{url: process.env.RENDER_EXTERNAL_URL ? process.env.RENDER_EXTERNAL_URL+"/api/webapp/open" : "https://pl3service.onrender.com/api/webapp/open"}}],
  ], resize_keyboard:true }};

  function KB_CUR() { return { reply_markup:{ keyboard:[
    [{text:"📅 Посещаемость сегодня"}, {text:"🔍 Найти студента"}],
    [{text:"⏰ Опоздания сегодня"},     {text:"📥 Посещаемость Excel"}],
    [{text:"📝 Заполнить ведомость"},  {text:"📚 Предметы группы"}],
    [{text:"📤 Отправить ведомость"},  {text:"📋 Оценки группы"}],
    [{text:"👥 Список студентов"},      {text:"📱 Мини-приложение", web_app:{url: process.env.RENDER_EXTERNAL_URL ? process.env.RENDER_EXTERNAL_URL+"/api/webapp/open" : "https://pl3service.onrender.com/api/webapp/open"}}],
  ], resize_keyboard:true }}; }

  const KB_PAR = { reply_markup:{ keyboard:[
    [{text:"📊 Оценки ребёнка"},    {text:"🗓 Посещаемость"}],
    [{text:"👨‍👩‍👧 Мои дети"},          {text:"➕ Привязать ребёнка"}],
    [{text:"📱 Мини-приложение", web_app:{url: process.env.RENDER_EXTERNAL_URL ? process.env.RENDER_EXTERNAL_URL+"/api/webapp/open" : "https://pl3service.onrender.com/api/webapp/open"}}],
  ], resize_keyboard:true }};

  const KB_X = { reply_markup:{ keyboard:[[{text:"❌ Отмена"}]], resize_keyboard:true }};

  function mkb(role) {
    if (role==="admin")   return KB_ADMIN;
    if (role==="curator") return KB_CUR();
    if (role==="parent")  return KB_PAR;
    return KB_X;
  }

  function monthKb(months) {
    const rows=[];
    for (let i=0;i<months.length;i+=2) {
      const r=[{text:fmtM(months[i])}];
      if(months[i+1]) r.push({text:fmtM(months[i+1])});
      rows.push(r);
    }
    rows.push([{text:"❌ Отмена"}]);
    return {reply_markup:{keyboard:rows,resize_keyboard:true}};
  }

  async function grpKb(gids=null) {
    const f = gids?.length ? {_id:{$in:gids}} : {};
    const gs = await Group.find(f).sort({name:1});
    const rows=[];
    for(let i=0;i<gs.length;i+=2){const r=[{text:gs[i].name}];if(gs[i+1])r.push({text:gs[i+1].name});rows.push(r);}
    rows.push([{text:"❌ Отмена"}]);
    return {groups:gs, kb:{reply_markup:{keyboard:rows,resize_keyboard:true}}};
  }

  /* ──────────────────────────────────────────────────────────────
     ROLE RESOLUTION  (MongoDB → cache)
     ИЗОЛЯЦИЯ: каждая роль видит только своё
  ────────────────────────────────────────────────────────────── */
  async function resolve(cid) {
    const cached = S(cid);
    if (cached.role) return cached;

    const sub = await TgSub.findOne({ chatId:String(cid) });
    if (!sub) return {};

    if (sub.role==="admin") {
      set(cid, {role:"admin", name:sub.name});
      return S(cid);
    }
    if (sub.role==="curator" && sub.curatorId) {
      const cur = await Curator.findById(sub.curatorId).populate("groups");
      if (cur) {
        const gids = cur.groups.map(g=>String(g._id));
        set(cid, {role:"curator", name:cur.fullName||cur.username, username:cur.username,
                  groupIds:gids, curatorId:String(cur._id)});
        return S(cid);
      }
    }
    if (sub.role==="parent") {
      // ИЗОЛЯЦИЯ: parent видит ТОЛЬКО свои ИНН
      // Если ИНН уже привязаны - сбрасываем шаг (не держим link_inn)
      const parentInns = sub.linkedInns||[];
      SESS.set(String(cid), {
        role:"parent", name:sub.name, inns:parentInns,
        // step не устанавливаем - будет undefined/idle
      });
      return S(cid);
    }
    return {};
  }

  /* ──────────────────────────────────────────────────────────────
     GUARD: проверить что groupId входит в groupIds куратора
  ────────────────────────────────────────────────────────────── */
  function canAccessGroup(s, gid) {
    if (s.role==="admin") return true;
    if (s.role==="curator") return (s.groupIds||[]).map(String).includes(String(gid));
    return false;
  }

  /* ══════════════════════════════════════════════════════════════
     /start
  ══════════════════════════════════════════════════════════════ */
  bot.onText(/\/start(.*)/, async (msg, match) => {
    const cid  = String(msg.chat.id);
    const name = msg.from?.first_name || "Пользователь";
    const prm  = (match[1]||"").trim();

    /* ── Admin ── */
    if (prm===`admin_${ADMIN_SECRET}`) {
      await TgSub.findOneAndUpdate({chatId:cid},
        {chatId:cid, name, role:"admin", curatorId:null, linkedInns:[]}, {upsert:true});
      evictCache(cid);
      set(cid,{role:"admin",name});
      await send(cid,`🔐 *Администратор авторизован\\!*\n\nПолный доступ к системе ПЛ №3\\.`,KB_ADMIN);
      return;
    }

    /* ── Curator (HMAC-signed token) ── */
    if (prm.startsWith("curator_")) {
      const token = prm.slice(8);
      const sep   = token.lastIndexOf("_");
      if (sep<0) { await send(cid,"❌ Недействительная ссылка\\."); return; }
      const curId = token.slice(0,sep), hmac = token.slice(sep+1);
      const expect = crypto.createHmac("sha256",ADMIN_SECRET).update(curId).digest("hex").slice(0,8);
      if (hmac!==expect) { await send(cid,"❌ Ссылка недействительна или устарела\\."); return; }
      const cur = await Curator.findById(curId).populate("groups");
      if (!cur) { await send(cid,"❌ Куратор не найден\\."); return; }
      await TgSub.findOneAndUpdate({chatId:cid},
        {chatId:cid, name, role:"curator", curatorId:String(cur._id), linkedInns:[]},
        {upsert:true});
      evictCache(cid);
      const gids = cur.groups.map(g=>String(g._id));
      set(cid,{role:"curator",name:cur.fullName||cur.username,username:cur.username,
               groupIds:gids,curatorId:String(cur._id)});
      await send(cid,
        `✅ *Добро пожаловать, ${ev(cur.fullName||cur.username)}\\!*\n\n` +
        `📂 Ваши группы: *${ev(cur.groups.map(g=>g.name).join(", "))}*\n\n` +
        `_Вы будете получать уведомления об опозданиях только ваших студентов\\._`,
        KB_CUR());
      return;
    }

    /* ── Existing user ── */
    const s = await resolve(cid);
    if (s.role) { await send(cid,`👋 *С возвращением, ${ev(s.name||name)}\\!*`,mkb(s.role)); return; }

    /* ── New user → parent ── */
    await TgSub.findOneAndUpdate({chatId:cid},{chatId:cid,name,role:"parent",linkedInns:[]},{upsert:true});
    set(cid,{step:"link_inn",name});
    await send(cid,
      `👋 *Добро пожаловать в ПЛ №3\\!*\n\n` +
      `Введите *ИНН ребёнка* \\(12–14 цифр\\) для привязки\\.\n\n` +
      `_Куратор и администратор используют специальную ссылку для входа\\._`, KB_X);
  });

  /* ══════════════════════════════════════════════════════════════
     /reset — сбросить свою роль (для тестов / переключения)
  ══════════════════════════════════════════════════════════════ */
  bot.onText(/\/reset/, async (msg) => {
    const cid = String(msg.chat.id);
    await TgSub.deleteOne({chatId:cid});
    evictCache(cid);
    await send(cid,
      `🔄 *Сессия сброшена\\.*\n\nЗапустите /start чтобы начать заново\\.`,
      {reply_markup:{remove_keyboard:true}});
  });


  /* ══════════════════════════════════════════════════════════════
     /admin SECRET — войти как администратор
  ══════════════════════════════════════════════════════════════ */
  bot.onText(/\/admin (.+)/, async (msg, match) => {
    const cid    = String(msg.chat.id);
    const secret = (match[1]||"").trim();
    if (secret !== ADMIN_SECRET) { await send(cid,"❌ Неверный секрет\."); return; }
    const name = msg.from?.first_name || "Admin";
    await TgSub.findOneAndUpdate({chatId:cid},
      {chatId:cid, name, role:"admin", curatorId:null, linkedInns:[]}, {upsert:true});
    evictCache(cid);
    set(cid,{role:"admin",name});
    await send(cid,`🔐 *Администратор авторизован\!*

Полный доступ к системе ПЛ №3\.`,KB_ADMIN);
  });

  /* ── /curator TOKEN — войти как куратор ── */
  bot.onText(/\/curator (.+)/, async (msg, match) => {
    const cid   = String(msg.chat.id);
    const token = (match[1]||"").trim();
    const sep   = token.lastIndexOf("_");
    if (sep < 0) { await send(cid,"❌ Неверный токен\."); return; }
    const curId = token.slice(0,sep), hmac = token.slice(sep+1);
    const expect = crypto.createHmac("sha256",ADMIN_SECRET).update(curId).digest("hex").slice(0,8);
    if (hmac !== expect) { await send(cid,"❌ Токен недействителен\."); return; }
    const cur = await Curator.findById(curId).populate("groups");
    if (!cur) { await send(cid,"❌ Куратор не найден\."); return; }
    const name = msg.from?.first_name || cur.fullName || cur.username;
    await TgSub.findOneAndUpdate({chatId:cid},
      {chatId:cid, name, role:"curator", curatorId:String(cur._id), linkedInns:[]},
      {upsert:true});
    evictCache(cid);
    const gids = cur.groups.map(g=>String(g._id));
    set(cid,{role:"curator",name:cur.fullName||cur.username,username:cur.username,
             groupIds:gids,curatorId:String(cur._id)});
    await send(cid,
      `✅ *Добро пожаловать, ${ev(cur.fullName||cur.username)}\!*

`+
      `📂 Ваши группы: *${ev(cur.groups.map(g=>g.name).join(", "))}*`,
      KB_CUR());
  });

  /* ── /help — инструкция ── */
  bot.onText(/\/help/, async (msg) => {
    const cid = String(msg.chat.id);
    await send(cid,
      `📖 *Как войти в систему:*

`+
      `👨\u200d👩\u200d👧 *Родитель:*
Напишите /start → введите ИНН ребёнка\.

`+
      `📚 *Куратор:*
Получите токен от администратора и введите:
\`/curator ВАШ\_ТОКЕН\`

`+
      `🔐 *Администратор:*
Введите: \`/admin СЕКРЕТ\`

`+
      `🔄 *Сброс роли:* /reset`,
      {reply_markup:{remove_keyboard:true}});
  });


  bot.onText(/\/whoami/, async (msg) => {
    const cid = String(msg.chat.id);
    const s   = await resolve(cid);
    const roles = { admin:"Администратор 🔐", curator:"Куратор 📚", parent:"Родитель 👨‍👩‍👧" };
    if (!s.role) {
      await send(cid,"❓ *Роль не определена*\.Напишите /help\."); return;
    }
    await send(cid,
      `👤 *Роль:* ${ev(roles[s.role]||s.role)}
📛 *Имя:* ${ev(s.name||"—")}`,
      mkb(s.role));
  });

  /* ══════════════════════════════════════════════════════════════
     MAIN ROUTER
  ══════════════════════════════════════════════════════════════ */
  bot.on("message", async (msg) => {
    if (!msg.text) return;
    const cid = String(msg.chat.id);
    const txt = msg.text.trim();
    if (txt.startsWith("/")) return;

    const s  = await resolve(cid);
    const st = S(cid);

    if (txt === "❌ Отмена") { clearStep(cid); await send(cid, "Главное меню 👇", mkb(s.role)); return; }

    // Кнопки меню всегда работают — даже если есть активный step
    // Это исправляет проблему когда родитель нажимает кнопку меню
    // но step:"link_inn" ещё активен из предыдущего действия
    const PARENT_BTNS = ["📊 Оценки ребёнка","🗓 Посещаемость","👨‍👩‍👧 Мои дети","➕ Привязать ребёнка"];
    const CURATOR_BTNS = ["📅 Посещаемость сегодня","🔍 Найти студента","⏰ Опоздания сегодня",
      "📥 Посещаемость Excel","📝 Заполнить ведомость","📚 Предметы группы",
      "📤 Отправить ведомость","📋 Оценки группы","👥 Список студентов"];
    const ADMIN_BTNS = ["📊 Статистика","📅 Посещаемость групп","🔍 Найти студента",
      "📝 Ведомость","⏰ Список опозданий","🔗 Ссылки кураторов"];

    const isMenuBtn = (s.role==="parent"  && PARENT_BTNS.includes(txt))  ||
                      (s.role==="curator" && CURATOR_BTNS.includes(txt)) ||
                      (s.role==="admin"   && ADMIN_BTNS.includes(txt));

    if (isMenuBtn) {
      clearStep(cid); // Сбросить любой активный шаг при нажатии кнопки меню
      if (s.role === "admin")   { await rAdmin(cid, txt, s);   return; }
      if (s.role === "curator") { await rCurator(cid, txt, s); return; }
      if (s.role === "parent")  { await rParent(cid, txt, s);  return; }
    }

    if (st.step && st.step !== "idle") { await handleStep(cid, txt, msg, s, st); return; }

    if (s.role === "admin")   { await rAdmin(cid, txt, s);   return; }
    if (s.role === "curator") { await rCurator(cid, txt, s); return; }
    if (s.role === "parent")  { await rParent(cid, txt, s);  return; }

    set(cid, { step: "link_inn", name: msg.from?.first_name });
    await send(cid, "Введите *ИНН ребёнка* для привязки:", KB_X);
  });

  /* ══════════════════════════════════════════════════════════════
     KEYBOARDS — полный набор
  ══════════════════════════════════════════════════════════════ */
  // (KB_ADMIN, KB_CUR, KB_PAR, KB_X, mkb, monthKb, grpKb — уже определены выше)

  /* ══════════════════════════════════════════════════════════════
     ADMIN — полный доступ
  ══════════════════════════════════════════════════════════════ */
  async function rAdmin(cid, txt, s) {
    switch (txt) {
      case "📊 Статистика":             return adminStats(cid);
      case "📅 Посещаемость групп":     return adminAttAll(cid);
      case "🔍 Найти студента":          set(cid,{step:"a_search"}); return send(cid,"🔍 Введите *ФИО* или *ИНН*:",KB_X);
      case "📝 Ведомость":              return stepGroupThen(cid,"a_ved_g","a_ved_m",null);
      case "⏰ Список опозданий":       return adminLateAll(cid);
      case "🔗 Ссылки кураторов":       return adminCuratorLinks(cid);
      default: await send(cid,"Используйте кнопки меню 👇", KB_ADMIN);
    }
  }

  /* ── Статистика сегодня ── */
  async function adminStats(cid) {
    const total   = await Student.countDocuments();
    const groups  = await Group.countDocuments();
    const tdk     = getTodayKey();
    const bd      = getBishkekNow();
    const wknd    = bd.getDay()===0||bd.getDay()===6;
    const present = await Student.countDocuments({$or:[
      {skudDaily:{$elemMatch:{date:tdk,present:true}}},
      {attendance:{$elemMatch:{date:tdk,present:true}}},
    ]});
    const pct = total ? Math.round(present/total*100) : 0;
    const bar = "█".repeat(Math.round(pct/10))+"░".repeat(10-Math.round(pct/10));
    await send(cid,
      `📊 *Статистика — ${ev(pad2(bd.getDate()))}\\.${ev(pad2(bd.getMonth()+1))}*\n\n`+
      `${bar} *${pct}%*\n\n`+
      `👥 Студентов: *${total}*  📂 Групп: *${groups}*\n`+
      `✅ Присутствуют: *${present}*\n`+
      `❌ Отсутствуют: *${wknd?0:total-present}*`+
      (wknd?"\n\n_📅 Сегодня выходной_":""), KB_ADMIN);
  }

  /* ── Посещаемость всех групп сегодня ── */
  async function adminAttAll(cid) {
    const tdk  = getTodayKey();
    const bd   = getBishkekNow();
    const groups = await Group.find().sort({name:1});
    if (!groups.length) { await send(cid,"Групп нет\\.", KB_ADMIN); return; }
    const lines = await Promise.all(groups.map(async g => {
      const total   = await Student.countDocuments({group:g._id});
      const present = await Student.countDocuments({group:g._id,$or:[
        {skudDaily:{$elemMatch:{date:tdk,present:true}}},
        {attendance:{$elemMatch:{date:tdk,present:true}}},
      ]});
      const pct = total ? Math.round(present/total*100) : 0;
      const bar = "▓".repeat(Math.round(pct/5))+"░".repeat(20-Math.round(pct/5));
      return `📂 *${ev(g.name)}*\n${bar} *${pct}%* \\(${present}/${total}\\)`;
    }));
    const dateStr = `${pad2(bd.getDate())}\\.${pad2(bd.getMonth()+1)}`;
    // Telegram 4096 limit
    let msg = `📅 *Посещаемость — ${dateStr}*\n\n`;
    for (const l of lines) {
      if (msg.length + l.length > 3900) { await send(cid, msg, KB_ADMIN); msg = ""; }
      msg += l + "\n\n";
    }
    if (msg) await send(cid, msg, KB_ADMIN);
  }

  /* ── Список опозданий всех групп сегодня ── */
  async function adminLateAll(cid) {
    const tdk = getTodayKey();
    const bd  = getBishkekNow();
    const groups = await Group.find().sort({name:1});
    let lateList = [];
    for (const g of groups) {
      const [sH,sM] = (g.shiftStart||"09:30").split(":").map(Number), stm=sH*60+sM;
      const stus = await Student.find({group:g._id}).select("fio skudDaily");
      stus.forEach(stu => {
        const skud = stu.skudDaily?.find(r=>r.date===tdk);
        if (!skud?.firstIn) return;
        const dt = new Date(skud.firstIn);
        if (isNaN(dt)) return;
        const tm = dt.getUTCHours()*60+dt.getUTCMinutes()+360;
        const arrH = Math.floor(tm/60)%24, arrM = tm%60;
        if (arrH*60+arrM > stm) {
          const delay = arrH*60+arrM - stm;
          lateList.push(`⏰ *${ev(stu.fio)}* \\(${ev(g.name)}\\) — ${pad2(arrH)}:${pad2(arrM)} \\+${delay} мин`);
        }
      });
    }
    const dateStr = `${pad2(bd.getDate())}\\.${pad2(bd.getMonth()+1)}`;
    if (!lateList.length) {
      await send(cid,`✅ *Опозданий нет* — ${dateStr}`, KB_ADMIN); return;
    }
    let msg = `⏰ *Опоздания — ${dateStr}* \\(${lateList.length}\\)\n\n`;
    for (const l of lateList) {
      if (msg.length + l.length > 3900) { await send(cid, msg); msg = ""; }
      msg += l + "\n";
    }
    if (msg) await send(cid, msg, KB_ADMIN);
  }

  /* ── Поиск студента (admin) ── */
  async function adminSearch(cid, query) {
    const isInn = /^\d{12,16}$/.test(query.replace(/\D/g,""));
    const filter = isInn ? {inn:query.replace(/\D/g,"")} : {fio:new RegExp(query,"i")};
    const stus = await Student.find(filter).populate("group").limit(10);
    if (!stus.length) { await send(cid,`❌ Не найдено: *${ev(query)}*`, KB_ADMIN); return; }
    const tdk = getTodayKey();
    const lines = await Promise.all(stus.map(async stu => {
      const skud = stu.skudDaily?.find(r=>r.date===tdk);
      const att  = stu.attendance?.find(a=>a.date===tdk);
      const here = skud?.present||att?.present||false;
      return `${here?"✅":"❌"} *${ev(stu.fio)}*\n   📂 ${ev(stu.group?.name||"—")} \\| ИНН: \`${ev(stu.inn||"—")}\``;
    }));
    await send(cid,`🔍 *Найдено: ${stus.length}*\n\n${lines.join("\n\n")}`, KB_ADMIN);
  }

  /* ── Ссылки для кураторов ── */
  async function adminCuratorLinks(cid) {
    const curators = await Curator.find().populate("groups","name");
    if (!curators.length) { await send(cid,"Кураторов нет\\.", KB_ADMIN); return; }

    // Отправляем каждого куратора отдельным сообщением чтобы не превысить лимит 4096 символов
    await send(cid, `🔗 *Ссылки для входа кураторов:*

_Перешлите нужную ссылку куратору — он нажимает и авторизуется автоматически\._`, KB_ADMIN);

    for (const c of curators) {
      const curId   = String(c._id);
      const hmac    = crypto.createHmac("sha256", ADMIN_SECRET).update(curId).digest("hex").slice(0,8);
      const token   = `${curId}_${hmac}`;
      // Чистый URL без экранирования — только для href в inline-ссылке
      const cleanUrl = `https://t.me/${BOT_USERNAME}?start=curator_${token}`;
      const groupNames = c.groups.map(g=>g.name).join(", ") || "—";

      // Inline-ссылка в MarkdownV2: [текст](url) — URL не экранируется
      const msg =
        `👤 *${ev(c.fullName||c.username)}*
` +
        `📂 ${ev(groupNames)}

` +
        `[🔗 Нажать для входа](${cleanUrl})

` +
        `Или написать боту:
\`/curator ${token}\``;

      await send(cid, msg);
    }
  }

  /* ── Word ведомость ── */
  async function sendWord(cid, gid, month, kb) {
    try {
      await send(cid,"⏳ Генерирую файл\\.\\.\\.");
      const data = await getSummaryDocumentData(gid, month);
      const buf  = buildSummaryDocx(data);
      const sn   = data.group.name.replace(/[^a-zA-Z0-9._-]/g,"_");
      await bot.sendDocument(cid, buf,
        {caption:`📝 Сводная ведомость\n📂 ${data.group.name}\n📅 ${fmtM(month)}\n👥 ${data.rows.length} студентов`},
        {filename:`Vedomost_${sn}_${month}.docx`, contentType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
      await send(cid,"Главное меню:", kb||KB_ADMIN);
    } catch(e) { await send(cid,`❌ Ошибка: ${ev(e.message)}`, kb||KB_ADMIN); }
  }

  /* ── Excel журнал ── */
  async function sendXlsx(cid, gid, month, kb) {
    try {
      const xlsx = require("xlsx");
      const [yr,mn] = month.split("-").map(Number);
      const dim  = new Date(yr,mn,0).getDate();
      const g    = await Group.findById(gid);
      const stus = await Student.find({group:gid}).sort({fio:1});
      const [sH,sM] = (g?.shiftStart||"09:30").split(":").map(Number), stm=sH*60+sM;
      const tdk  = getTodayKey();
      const data = [["№","ФИО",...Array.from({length:dim},(_,i)=>String(i+1)),"П","О","Н"]];
      stus.forEach((stu,i) => {
        const row=[i+1,stu.fio]; let p=0,a=0,l=0;
        for (let d=1;d<=dim;d++) {
          const dk=`${yr}-${pad2(mn)}-${pad2(d)}`;
          const iw=new Date(Date.UTC(yr,mn-1,d)).getUTCDay()%6===0;
          const skud=stu.skudDaily?.find(r=>r.date===dk),att=stu.attendance?.find(a=>a.date===dk);
          let here=false,late=false;
          if(skud){here=skud.present;if(skud.firstIn){const dt=new Date(skud.firstIn);if(!isNaN(dt)){const tm=dt.getUTCHours()*60+dt.getUTCMinutes()+360;late=Math.floor(tm/60)%24*60+tm%60>stm;}}}
          else if(att){here=att.present;}
          if(here){p++;late?(l++,row.push("О")):row.push("П");}
          else if(!iw&&dk<=tdk){a++;row.push("Н");}else row.push(iw?"вых":"");
        }
        row.push(p,l,a); data.push(row);
      });
      const wb=xlsx.utils.book_new(),ws=xlsx.utils.aoa_to_sheet(data);
      xlsx.utils.book_append_sheet(wb,ws,"Журнал");
      const buf=xlsx.write(wb,{type:"buffer",bookType:"xlsx"});
      await bot.sendDocument(cid,buf,
        {caption:`📥 Журнал посещаемости\n📂 ${g?.name}\n📅 ${fmtM(month)}`},
        {filename:`Journal_${g?.name}_${month}.xlsx`, contentType:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
      await send(cid,"Главное меню:", kb||KB_CUR());
    } catch(e) { await send(cid,`❌ Ошибка: ${ev(e.message)}`); }
  }

  async function stepGroupThen(cid, gStep, mStep, gids) {
    if (gids?.length===1) { set(cid,{step:mStep,gid:gids[0]}); await send(cid,"📅 Выберите месяц:",monthKb(last6())); }
    else { set(cid,{step:gStep}); const{kb}=await grpKb(gids); await send(cid,"📂 Выберите группу:",kb); }
  }

  /* ══════════════════════════════════════════════════════════════
     CURATOR — только своя группа
  ══════════════════════════════════════════════════════════════ */
  async function rCurator(cid, txt, s) {
    const gids = s.groupIds||[];
    switch(txt) {
      case "📅 Посещаемость сегодня":   return curAttToday(cid, s);
      case "🔍 Найти студента":          set(cid,{step:"c_search"}); return send(cid,"🔍 Введите *ФИО* или *ИНН*:",KB_X);
      case "⏰ Опоздания сегодня":       return curLateToday(cid, s);
      case "📥 Посещаемость Excel":      return stepGroupThen(cid,"c_xl_g","c_xl_m",gids);
      case "📝 Заполнить ведомость":     return curFill(cid, s);
      case "📚 Предметы группы":         return curSubjects(cid, s);
      case "📤 Отправить ведомость":     return stepGroupThen(cid,"c_send_g","c_send_m",gids);
      case "📋 Оценки группы":          return stepGroupThen(cid,"c_grades_g","c_grades_m",gids);
      case "👥 Список студентов":        return curStudents(cid, s);
      default: await send(cid,"Используйте кнопки меню 👇", KB_CUR());
    }
  }

  /* ── Посещаемость своих групп сегодня ── */
  async function curAttToday(cid, s) {
    const tdk = getTodayKey(), bd = getBishkekNow();
    for (const gid of (s.groupIds||[])) {
      if (!canAccessGroup(s,gid)) continue;
      const g    = await Group.findById(gid);
      const stus = await Student.find({group:gid}).sort({fio:1});
      const [sH,sM]=(g?.shiftStart||"09:30").split(":").map(Number),stm=sH*60+sM;
      let pL=[],lL=[],aL=[];
      stus.forEach(stu=>{
        const skud=stu.skudDaily?.find(r=>r.date===tdk),att=stu.attendance?.find(a=>a.date===tdk);
        let here=false,late=false,ts="";
        if(skud){here=skud.present;if(skud.firstIn){const dt=new Date(skud.firstIn);if(!isNaN(dt)){const tm=dt.getUTCHours()*60+dt.getUTCMinutes()+360;ts=`${pad2(Math.floor(tm/60)%24)}:${pad2(tm%60)}`;late=Math.floor(tm/60)%24*60+tm%60>stm;}}}
        else if(att){here=att.present;}
        if(here){late?lL.push(`• ${ev(stu.fio)} _\\(${ev(ts)}\\)_`):pL.push(`• ${ev(stu.fio)}${ts?` _\\(${ev(ts)}\\)_`:""}`);}
        else aL.push(`• ${ev(stu.fio)}`);
      });
      const n=pL.length+lL.length,pct=stus.length?Math.round(n/stus.length*100):0;
      const bar="█".repeat(Math.round(pct/10))+"░".repeat(10-Math.round(pct/10));
      let r=`📅 *${ev(g?.name)}* — ${pad2(bd.getDate())}\\.${pad2(bd.getMonth()+1)}\n${bar} *${pct}%*\n\n`;
      if(pL.length) r+=`✅ *Вовремя \\(${pL.length}\\):*\n${pL.join("\n")}\n\n`;
      if(lL.length) r+=`⏰ *Опоздали \\(${lL.length}\\):*\n${lL.join("\n")}\n\n`;
      if(aL.length) r+=`❌ *Отсутствуют \\(${aL.length}\\):*\n${aL.join("\n")}`;
      if(r.length>4000) r=r.slice(0,3800)+"\n_\\.\\.\\.сокращено_";
      await send(cid, r, KB_CUR());
    }
  }

  /* ── Опоздания своей группы ── */
  async function curLateToday(cid, s) {
    const tdk = getTodayKey(), bd = getBishkekNow();
    for (const gid of (s.groupIds||[])) {
      if (!canAccessGroup(s,gid)) continue;
      const g    = await Group.findById(gid);
      const stus = await Student.find({group:gid}).select("fio skudDaily");
      const [sH,sM]=(g?.shiftStart||"09:30").split(":").map(Number),stm=sH*60+sM;
      const lateList=[];
      stus.forEach(stu=>{
        const skud=stu.skudDaily?.find(r=>r.date===tdk);
        if(!skud?.firstIn) return;
        const dt=new Date(skud.firstIn);
        if(isNaN(dt)) return;
        const tm=dt.getUTCHours()*60+dt.getUTCMinutes()+360;
        const arrH=Math.floor(tm/60)%24,arrM=tm%60;
        if(arrH*60+arrM>stm){
          const delay=arrH*60+arrM-stm;
          lateList.push(`⏰ *${ev(stu.fio)}* — ${pad2(arrH)}:${pad2(arrM)} \\(\\+${delay} мин\\)`);
        }
      });
      const dateStr=`${pad2(bd.getDate())}\\.${pad2(bd.getMonth()+1)}`;
      if(!lateList.length){ await send(cid,`✅ *${ev(g?.name)}* — опозданий нет на ${dateStr}`,KB_CUR()); }
      else { await send(cid,`⏰ *Опоздания ${ev(g?.name)}* — ${dateStr}\n\n${lateList.join("\n")}`,KB_CUR()); }
    }
  }

  /* ── Поиск студента (curator, только своя группа) ── */
  async function curSearch(cid, query, s) {
    const gids = s.groupIds||[];
    const isInn = /^\d{12,16}$/.test(query.replace(/\D/g,""));
    const filter = { group:{$in:gids}, ...(isInn?{inn:query.replace(/\D/g,"")}:{fio:new RegExp(query,"i")}) };
    const stus = await Student.find(filter).populate("group").limit(10);
    if (!stus.length) { await send(cid,`❌ Не найдено в ваших группах: *${ev(query)}*`,KB_CUR()); return; }
    const tdk = getTodayKey();
    const lines = stus.map(stu=>{
      const skud=stu.skudDaily?.find(r=>r.date===tdk),att=stu.attendance?.find(a=>a.date===tdk);
      const here=skud?.present||att?.present||false;
      return `${here?"✅":"❌"} *${ev(stu.fio)}*\n   📂 ${ev(stu.group?.name||"—")} \\| ИНН: \`${ev(stu.inn||"—")}\``;
    });
    await send(cid,`🔍 *Найдено: ${stus.length}*\n\n${lines.join("\n\n")}`,KB_CUR());
  }

  /* ── Список студентов с добавлением/удалением ── */
  async function curStudents(cid, s) {
    const gids = s.groupIds||[];
    if (!gids.length) { await send(cid,"❌ Нет групп\\."); return; }
    if (gids.length===1) {
      set(cid,{step:"c_stu_list_show",gid:gids[0]});
      await showStudentList(cid,gids[0]);
    } else {
      set(cid,{step:"c_stu_list_grp"});
      const{kb}=await grpKb(gids);
      await send(cid,"📂 Выберите группу:",kb);
    }
  }

  /* ══════════════════════════════════════════════════════════════
     СПИСОК СТУДЕНТОВ — добавление/удаление
  ══════════════════════════════════════════════════════════════ */
  async function showStudentList(cid, gid) {
    const g    = await Group.findById(gid);
    const stus = await Student.find({group:gid}).sort({fio:1});

    if (!stus.length) {
      await send(cid,
        `👥 *${ev(g?.name||"?")}* — список пуст\n\nНажмите чтобы добавить студента:`,
        {reply_markup:{inline_keyboard:[
          [{text:"➕ Добавить студента", callback_data:`cstu_add:${gid}`}]
        ]}}
      ); return;
    }

    // Студенты — по одному в ряд: [ФИО полностью] [🗑️]
    const rows = stus.map((stu,i) => [
      { text: `${i+1}. ${stu.fio}`,
        callback_data: `cstu_view:${stu._id}:${gid}` },
      { text: "🗑️",
        callback_data: `cstu_delconf:${stu._id}:${gid}` },
    ]);
    rows.push([{text:"➕ Добавить студента", callback_data:`cstu_add:${gid}`}]);

    await send(cid,
      `👥 *Список — ${ev(g?.name||"?")}*\n_${stus.length} студентов_\n\n📋 = просмотр  🗑️ = удалить:`,
      {reply_markup:{inline_keyboard: rows}}
    );
  }

  /* ══════════════════════════════════════════════════════════════
     ЗАПОЛНЕНИЕ ВЕДОМОСТИ — предмет → список студентов
     Логика: выбрать предмет → каждый студент отдельной строкой
             ФИО видно полностью, оценка справа кнопками
  ══════════════════════════════════════════════════════════════ */
  /* ── Предметы группы: просмотр, добавление, удаление ── */
  async function curSubjects(cid, s) {
    const gids = s.groupIds||[];
    if (!gids.length) { await send(cid,"❌ Нет групп\."); return; }
    if (gids.length===1) {
      set(cid,{step:"c_subj_show", gid:gids[0]});
      await showSubjectList(cid, gids[0]);
    } else {
      set(cid,{step:"c_subj_grp"});
      const{kb}=await grpKb(gids);
      await send(cid,"📂 Выберите группу:",kb);
    }
  }

  async function showSubjectList(cid, gid) {
    const g     = await Group.findById(gid);
    const subjs = await ensureGroupSummarySubjects(g);
    const rows  = subjs.map((s,i)=>[
      {text:`${s.section==="practice"?"🔧":"📖"} ${s.label}`, callback_data:`csubj_noop:x`},
      {text:"❌", callback_data:`csubj_del:${gid}:${i}`},
    ]);
    rows.push([{text:"➕ Добавить предмет (теория)",   callback_data:`csubj_add:${gid}:theory`}]);
    rows.push([{text:"🔧 Добавить предмет (практика)", callback_data:`csubj_add:${gid}:practice`}]);
    await send(cid,
      `📚 *Предметы — ${ev(g?.name||"?")}*
_Всего: ${subjs.length}_

📖 = теория  🔧 = практика
Нажмите ❌ чтобы убрать:`,
      {reply_markup:{inline_keyboard:rows}}
    );
  }

  /* ── Оценки группы ── */
  async function showGrades(cid, gid, month, s) {
    if (!canAccessGroup(s,gid)) { await send(cid,"❌ Нет доступа\."); return; }
    const data = await getSummaryDocumentData(gid, month);
    const all  = [...data.theorySubjects,...data.practiceSubjects];
    const emoji= {"5":"🟢","4":"🔵","3":"🟡","2":"🔴"};
    let txt = `📋 *${ev(data.group.name)}* — ${ev(fmtM(month))}

`;
    for (const row of data.rows) {
      const gs = all.map(subj=>{
        const g = row.grades instanceof Map ? row.grades.get(subj.key)||"" : row.grades?.[subj.key]||"";
        return g ? `${emoji[g]||"⚪"}${g}` : null;
      }).filter(Boolean);
      txt += `*${ev(row.fio)}*
${gs.length ? gs.join(" ") : "_нет оценок_"}

`;
    }
    if (txt.length>4000) txt=txt.slice(0,3800)+"_\.\.\.сокращено_";
    await send(cid, txt, KB_CUR());
  }

  /* ── Отправить ведомость администраторам ── */
  async function sendVedToAdmins(cid, gid, month) {
    if (!canAccessGroup(S(cid), gid)) { await send(cid,"❌ Нет доступа\."); return; }
    try {
      const data   = await getSummaryDocumentData(gid, month);
      const buf    = buildSummaryDocx(data);
      const sn     = data.group.name.replace(/[^a-zA-Z0-9._-]/g,"_");
      const admins = await TgSub.find({role:"admin"});
      for (const a of admins) {
        await bot.sendDocument(a.chatId, buf,
          {caption:`📝 Ведомость от куратора
📂 ${data.group.name}
📅 ${fmtM(month)}`},
          {filename:`Vedomost_${sn}_${month}.docx`, contentType:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"}
        ).catch(()=>{});
      }
      await send(cid,`✅ Отправлено ${admins.length} администраторам\.`,KB_CUR());
    } catch(e){ await send(cid,`❌ Ошибка: ${ev(e.message)}`,KB_CUR()); }
  }

  /* ── Родитель: оценки ребёнка ── */
  async function parGrades(cid, s) {
    const inns=s.inns||[];
    if(!inns.length){ await send(cid,"❌ Нет привязанных детей\.",KB_PAR); return; }
    const month=nowKey();
    const emoG={"5":"🟢","4":"🔵","3":"🟡","2":"🔴"};
    for(const inn of inns){
      const stu=await Student.findOne({inn}).populate("group");
      if(!stu) continue;
      const g=await Group.findById(stu.group);
      const subjs=await ensureGroupSummarySubjects(g);
      const rec=await SummaryRecord.findOne({group:g?._id,student:stu._id,month}).lean();
      const grades={};
      if(rec?.grades instanceof Map){rec.grades.forEach((v,k)=>{grades[k]=v;});}
      else if(rec?.grades){Object.entries(rec.grades).forEach(([k,v])=>{if(!k.startsWith("$"))grades[k]=v;});}
      const lines=subjs.map(s=>{
        const gr=grades[s.key]||"";
        return gr?`${emoG[gr]||"⚪"} ${ev(s.label)}: *${gr}*`:null;
      }).filter(Boolean);
      await send(cid,
        `📊 *Оценки — ${ev(stu.fio)}*
📂 ${ev(g?.name||"—")} \| ${ev(fmtM(month))}

`+
        (lines.length?lines.join(""):"_Оценок за этот месяц нет_"),
        KB_PAR
      );
    }
  }


  async function curFill(cid, s) {
    const gids = s.groupIds||[];
    if (!gids.length) { await send(cid,"❌ Нет прикреплённых групп\\."); return; }
    if (gids.length===1) {
      set(cid,{step:"fill_month",gid:gids[0]});
      await send(cid,"📅 Выберите месяц для заполнения:", monthKb(last6()));
    } else {
      set(cid,{step:"fill_group_sel"});
      const{kb}=await grpKb(gids);
      await send(cid,"📂 Выберите группу:",kb);
    }
  }

  async function initFill(cid, gid, month, s) {
    if (!canAccessGroup(s,gid)) { await send(cid,"❌ Нет доступа к этой группе\\."); return; }
    const g     = await Group.findById(gid);
    const subjs = await ensureGroupSummarySubjects(g);
    const stus  = await Student.find({group:gid}).sort({fio:1});
    const recs  = await SummaryRecord.find({group:gid,month}).lean();
    const gmap  = {};
    stus.forEach(stu=>{ gmap[String(stu._id)]={};});
    recs.forEach(r=>{
      const sid=String(r.student); if(!gmap[sid])gmap[sid]={};
      const raw=r.grades;
      if(raw instanceof Map) raw.forEach((v,k)=>{if(v)gmap[sid][k]=v;});
      else if(raw&&typeof raw==="object") Object.entries(raw).forEach(([k,v])=>{if(v&&!k.startsWith("$"))gmap[sid][k]=v;});
    });
    set(cid,{
      step:"fill_subj", fGid:gid, fGname:g.name, fMonth:month,
      fSubjs:subjs,
      fStus: stus.map(s=>({id:String(s._id), fio:s.fio, inn:s.inn||null})),
      fGrades:gmap, fSubjIdx:0,
    });
    await sendFillSubjMenu(cid);
  }

  /* ── Экран 1: Меню предметов ── */
  async function sendFillSubjMenu(cid) {
    const st = S(cid);
    const {fSubjs=[],fStus=[],fGrades={},fGname,fMonth} = st;

    let totalFilled=0, grand=fSubjs.length*fStus.length;
    fStus.forEach(stu=>fSubjs.forEach(s=>{if(fGrades[stu.id]?.[s.key])totalFilled++;}));
    const pct = grand?Math.round(totalFilled/grand*100):0;
    const bar = "▓".repeat(Math.round(pct/10))+"░".repeat(10-Math.round(pct/10));

    const txt =
      `📝 *Заполнение ведомости*\n`+
      `📂 *${ev(fGname)}* — ${ev(fmtM(fMonth))}\n\n`+
      `${bar} *${pct}%* \\(${totalFilled}/${grand}\\)\n\n`+
      `Выберите *предмет*:`;

    // Предметы — по 1 в ряд, показываем прогресс заполнения
    const rows = fSubjs.map((subj,i) => {
      const cnt  = fStus.filter(stu=>fGrades[stu.id]?.[subj.key]).length;
      const icon = cnt===fStus.length ? "✅" : cnt>0 ? "🔸" : "📝";
      const pctS = fStus.length ? `${cnt}/${fStus.length}` : "0";
      return [{
        text: `${icon} ${subj.label} (${pctS})`,
        callback_data: `fsubj:${i}`
      }];
    });

    rows.push([
      {text:"💾 Сохранить всё",  callback_data:"fsave"},
      {text:"📄 Word файл",      callback_data:"fword"},
    ]);

    await send(cid, txt, {reply_markup:{inline_keyboard:rows}});
  }

  /* ── Экран 2: Один предмет — список студентов 2 колонки ──
     Компактно: [оценка ФИО] [2][3][4][5]
     Каждый студент в ОДНОЙ строке из 5 кнопок
  ── */
  async function sendFillSubj(cid, subjIdx, mid) {
    const st = S(cid);
    const {fSubjs=[],fStus=[],fGrades={},fGname,fMonth} = st;
    if (subjIdx<0||subjIdx>=fSubjs.length) { await sendFillSubjMenu(cid); return; }

    const subj   = fSubjs[subjIdx];
    set(cid,{fSubjIdx:subjIdx});

    const filled = fStus.filter(stu=>fGrades[stu.id]?.[subj.key]).length;
    const emoG   = {"5":"🟢","4":"🔵","3":"🟡","2":"🔴"};

    // Текст сообщения — список студентов с текущими оценками
    // ФИО полностью в тексте, кнопки только для выставления оценки
    const stuLines = fStus.map((stu,i) => {
      const cur = fGrades[stu.id]?.[subj.key]||"";
      const mark = cur ? `${emoG[cur]||""}*${cur}*` : "_\\-_";
      return `${i+1}\\. ${ev(stu.fio)} — ${mark}`;
    }).join("\n");

    const txt =
      `📚 *${ev(subj.label)}*\n`+
      `📂 ${ev(fGname)} — ${ev(fmtM(fMonth))}\n`+
      `_Предмет ${subjIdx+1}/${fSubjs.length}_  •  Выставлено: *${filled}/${fStus.length}*\n\n`+
      `${stuLines}\n\n`+
      `⬇️ Нажмите студента для выставления оценки:`;

    // Кнопки: студенты по 2 в ряд — показываем номер + оценку
    // Нажатие открывает выбор оценки
    const stuRows = [];
    for (let i=0; i<fStus.length; i+=2) {
      const row = [];
      for (let j=0; j<2 && i+j<fStus.length; j++) {
        const stu = fStus[i+j];
        const cur = fGrades[stu.id]?.[subj.key]||"";
        const ico = cur ? emoG[cur]||"•" : "○";
        // Показываем номер + фамилию (первое слово) + оценку
        const lastName = stu.fio.split(" ")[0];
        row.push({
          text: `${ico}${i+j+1}.${lastName}${cur?" "+cur:""}`,
          callback_data: `fgrade_pick:${i+j}:${subjIdx}`
        });
      }
      stuRows.push(row);
    }

    // Навигация по предметам
    const nav = [];
    if(subjIdx>0)                nav.push({text:"⬅️",callback_data:`fsubj:${subjIdx-1}`});
    nav.push({text:"📋 Предметы", callback_data:"fsubjmenu"});
    if(subjIdx<fSubjs.length-1) nav.push({text:"➡️",callback_data:`fsubj:${subjIdx+1}`});
    stuRows.push(nav);
    stuRows.push([
      {text:"💾 Сохранить",     callback_data:"fsave"},
      {text:"📄 Word",          callback_data:"fword"},
    ]);

    if(mid) await edit(cid,mid,txt,{reply_markup:{inline_keyboard:stuRows}});
    else    await send(cid,txt,{reply_markup:{inline_keyboard:stuRows}});
  }

  /* ── Экран 3: Выбор оценки для конкретного студента ──
     Показываем полное ФИО + большие кнопки оценок
  ── */
  async function sendGradePick(cid, stuIdx, subjIdx, mid) {
    const st = S(cid);
    const {fSubjs=[],fStus=[],fGrades={},fGname,fMonth} = st;
    if (stuIdx<0||stuIdx>=fStus.length||subjIdx<0||subjIdx>=fSubjs.length) {
      await sendFillSubj(cid,subjIdx??0,mid); return;
    }

    const stu  = fStus[stuIdx];
    const subj = fSubjs[subjIdx];
    const cur  = fGrades[stu.id]?.[subj.key]||"";
    const emoG = {"5":"🟢","4":"🔵","3":"🟡","2":"🔴"};

    const txt =
      `📚 *${ev(subj.label)}*\n\n`+
      `👤 *${ev(stu.fio)}*\n`+
      `_Студент ${stuIdx+1} из ${fStus.length}_\n\n`+
      `Текущая оценка: ${cur ? `*${cur}* ${emoG[cur]||""}` : "_не выставлена_"}\n\n`+
      `Выберите оценку:`;

    const gradeRow = ["2","3","4","5"].map(v=>({
      text: cur===v ? `✅ ${v}` : `  ${v}  `,
      callback_data: `fg:${stu.id}:${subj.key}:${v}:${subjIdx}`
    }));
    gradeRow.push({text:"✕ Убрать", callback_data:`fg:${stu.id}:${subj.key}::${subjIdx}`});

    // Навигация по студентам
    const nav = [];
    if(stuIdx>0)               nav.push({text:"⬅️ Пред.",callback_data:`fgrade_pick:${stuIdx-1}:${subjIdx}`});
    nav.push({text:"↩️ Список", callback_data:`fsubj:${subjIdx}`});
    if(stuIdx<fStus.length-1) nav.push({text:"След. ➡️",callback_data:`fgrade_pick:${stuIdx+1}:${subjIdx}`});

    const kb = {reply_markup:{inline_keyboard:[gradeRow, nav]}};

    if(mid) await edit(cid,mid,txt,kb);
    else    await send(cid,txt,kb);
  }



  /* ══════════════════════════════════════════════════════════════
     CALLBACK QUERY HANDLER
  ══════════════════════════════════════════════════════════════ */
  bot.on("callback_query", async (q) => {
    const cid = String(q.message?.chat?.id);
    const mid = q.message?.message_id;
    const d   = q.data||"";
    await ack(q.id);
    const s  = await resolve(cid);
    const st = S(cid);

    /* ── FILL ── */
    if(d.startsWith("fsubj:"))      { await sendFillSubj(cid,+d.slice(6),mid); return; }
    if(d==="fsubjmenu")             { set(cid,{step:"fill_subj"}); await edit(cid,mid,"⏳"); await sendFillSubjMenu(cid); return; }
    if(d.startsWith("fl:"))         return; // noop

    // fgrade_pick:stuIdx:subjIdx — открыть экран выбора оценки для студента
    if(d.startsWith("fgrade_pick:")) {
      const [,siStr,sjStr] = d.split(":");
      await sendGradePick(cid, +siStr, +sjStr, mid);
      return;
    }

    if(d.startsWith("fg:")) {
      // fg:stuId:key:val:subjIdx
      const parts  = d.split(":");
      const stuId  = parts[1], key=parts[2], val=parts[3]||"", subjIdx=+(parts[4]||st.fSubjIdx||0);
      if(s.role==="curator"){
        const stu=await Student.findById(stuId).select("group");
        if(!canAccessGroup(s,String(stu?.group))){ await ack(q.id,"⛔ Нет доступа"); return; }
      }
      const ng={...st.fGrades};
      if(!ng[stuId])ng[stuId]={};
      if(val) ng[stuId][key]=val; else delete ng[stuId][key];
      set(cid,{fGrades:ng,fSubjIdx:subjIdx});
      // Autosave
      const grades={};
      Object.entries(ng[stuId]||{}).forEach(([k,v])=>{if(v)grades[k]=v;});
      SummaryRecord.findOneAndUpdate(
        {group:st.fGid,student:stuId,month:st.fMonth},
        {group:st.fGid,student:stuId,month:st.fMonth,grades},
        {upsert:true}
      ).then(async()=>{
        if(val){
          const stu =await Student.findById(stuId).select("fio inn");
          const subj=st.fSubjs?.find(s=>s.key===key);
          if(stu?.inn) notifyParents(stu.inn,`📊 *Новая оценка\\!*\n👤 ${ev(stu.fio)}\n📚 ${ev(subj?.label||key)}: *${val}*\n📅 ${ev(fmtM(st.fMonth))}`).catch(()=>{});
        }
      }).catch(()=>{});
      // После выставления оценки — вернуться на экран выбора оценки (не на список предметов)
      // Найти индекс студента по stuId
      const stuIdx = st.fStus?.findIndex(s=>s.id===stuId) ?? 0;
      await sendGradePick(cid, stuIdx, subjIdx, mid);
      return;
    }

    if(d==="fsave") {
      await saveAll(st);
      let filled=0; const tot=(st.fSubjs?.length||0)*(st.fStus?.length||0);
      st.fStus?.forEach(stu=>st.fSubjs?.forEach(s=>{if(st.fGrades?.[stu.id]?.[s.key])filled++;}));
      await edit(cid,mid,`✅ *Ведомость сохранена\\!*\n\n📂 *${ev(st.fGname)}*\n📅 ${ev(fmtM(st.fMonth))}\nЗаполнено: *${filled}/${tot}*`);
      clearStep(cid); await send(cid,"Главное меню:",KB_CUR());
      return;
    }

    if(d==="fword") {
      await saveAll(st);
      await edit(cid,mid,"⏳ *Генерирую Word\\.\\.\\.*");
      await sendWord(cid,st.fGid,st.fMonth,KB_CUR());
      clearStep(cid);
      return;
    }

    /* ── STUDENT LIST ── */
    // cstu_view:stuId:gid — просмотр карточки студента
    if(d.startsWith("cstu_view:")) {
      const parts = d.split(":");
      const stuId = parts[1], gid = parts[2];
      const stu = await Student.findById(stuId);
      if(!stu){ await ack(q.id,"Не найдено"); return; }
      await send(cid,
        `👤 *${ev(stu.fio)}*
`+
        `ИНН: \`${ev(stu.inn||"—")}\`
`+
        `Дата рождения: ${ev(stu.birthDate||"—")}
`+
        `Поименный №: ${ev(stu.lyceumId||"—")}`,
        {reply_markup:{inline_keyboard:[
          [{text:"🗑️ Удалить студента", callback_data:`cstu_delconf:${stuId}:${gid}`}],
          [{text:"↩️ Назад к списку",    callback_data:`cstu_back:${gid}`}],
        ]}}
      );
      return;
    }

    // cstu_delconf:stuId:gid — подтверждение удаления
    if(d.startsWith("cstu_delconf:")) {
      const parts = d.split(":");
      const stuId = parts[1], gid = parts[2];
      if(s.role==="curator"&&!canAccessGroup(s,gid)){ await ack(q.id,"⛔ Нет доступа"); return; }
      const stu = await Student.findById(stuId);
      if(!stu){ await ack(q.id,"Студент не найден"); return; }
      await edit(cid, mid,
        `⚠️ *Удалить студента?*

*${ev(stu.fio)}*

_Все оценки студента тоже будут удалены_`,
        {reply_markup:{inline_keyboard:[
          [
            {text:"✅ Да, удалить", callback_data:`cstu_delok:${stuId}:${gid}`},
            {text:"❌ Отмена",      callback_data:`cstu_back:${gid}`},
          ]
        ]}}
      );
      return;
    }

    // cstu_delok:stuId:gid — выполнить удаление
    if(d.startsWith("cstu_delok:")) {
      const parts = d.split(":");
      const stuId = parts[1], gid = parts[2];
      if(s.role==="curator"&&!canAccessGroup(s,gid)){ await ack(q.id,"⛔ Нет доступа"); return; }
      try {
        await SummaryRecord.deleteMany({student:stuId});
        await Student.findByIdAndDelete(stuId);
        await edit(cid,mid,`✅ *Студент удалён\.*`);
        await showStudentList(cid, gid);
      } catch(e) {
        await send(cid,`❌ Ошибка удаления: ${ev(e.message)}`);
      }
      return;
    }

    // cstu_back:gid — вернуться к списку без изменений
    if(d.startsWith("cstu_back:")) {
      const gid = d.slice(10);
      await showStudentList(cid, gid);
      return;
    }

    // СТАРЫЕ обработчики для совместимости
    if(d.startsWith("cstu_del:")) {
      const parts = d.split(":");
      const stuId = parts[1], gid = parts[2];
      if(s.role==="curator"&&!canAccessGroup(s,gid)){ await ack(q.id,"⛔ Нет доступа"); return; }
      const stu = await Student.findById(stuId);
      if(!stu){ await ack(q.id,"Студент не найден"); return; }
      await edit(cid,mid,
        `⚠️ *Удалить студента?*

*${ev(stu.fio)}*`,
        {reply_markup:{inline_keyboard:[[
          {text:"✅ Да",  callback_data:`cstu_delok:${stuId}:${gid}`},
          {text:"❌ Нет", callback_data:`cstu_back:${gid}`},
        ]]}}
      );
      return;
    }

    if(d.startsWith("cstu_cancel:")) {
      await showStudentList(cid, d.slice(12));
      return;
    }


    if(d.startsWith("cstu_add:")) {
      const gid=d.split(":")[1];
      set(cid,{step:"c_stu_add",gid});
      await send(cid,
        `➕ *Добавить студента*\n\nВведите данные в формате:\n\`ФИО / ИНН / Дата рождения / Поименный №\`\n\n_Пример:_\n\`Иванов Иван Иванович / 11012345678901 / 01\\.01\\.2007 / 265748\`\n\n_ИНН, дата и номер необязательны\\._`,
        KB_X
      );
      return;
    }

    /* ── SUBJECTS ── */
    if(d.startsWith("csubj_add:")) {
      const gid=d.split(":")[1];
      set(cid,{step:"c_subj_add",gid,subjSection:"theory"});
      await send(cid,"📖 *Добавить предмет \\(теория\\)*\n\nВведите название предмета:",KB_X);
      return;
    }

    if(d.startsWith("csubj_addprac:")) {
      const gid=d.split(":")[1];
      set(cid,{step:"c_subj_add",gid,subjSection:"practice"});
      await send(cid,"🔧 *Добавить спецпредмет \\(практика\\)*\n\nВведите название предмета:",KB_X);
      return;
    }

    if(d.startsWith("csubj_del:")) {
      const [,gid,idxStr]=d.split(":");
      if(s.role==="curator"&&!canAccessGroup(s,gid)){ await ack(q.id,"⛔ Нет доступа"); return; }
      const g=await Group.findById(gid);
      const subjs=await ensureGroupSummarySubjects(g);
      const idx=+idxStr;
      if(idx<0||idx>=subjs.length){ await ack(q.id,"Не найдено"); return; }
      const removed=subjs[idx].label;
      subjs.splice(idx,1);
      g.summarySubjects=subjs;
      await g.save();
      await ack(q.id,`✅ Убран: ${removed}`);
      await showSubjectList(cid,gid);
      return;
    }

    if(d.startsWith("csubj_noop:")) return;
  });

  async function saveAll(st) {
    const {fGid,fMonth,fGrades,fStus}=st;
    if(!fGid||!fMonth) return;
    for(const stu of (fStus||[])){
      const grades={};
      Object.entries(fGrades?.[stu.id]||{}).forEach(([k,v])=>{if(v)grades[k]=v;});
      await SummaryRecord.findOneAndUpdate(
        {group:fGid,student:stu.id,month:fMonth},
        {group:fGid,student:stu.id,month:fMonth,grades},
        {upsert:true}
      ).catch(()=>{});
    }
  }

  /* ══════════════════════════════════════════════════════════════
     PARENT — только свои дети
  ══════════════════════════════════════════════════════════════ */
  async function rParent(cid, txt, s) {
    switch(txt){
      case "📊 Оценки ребёнка":     return parGrades(cid, s);
      case "🗓 Посещаемость":        return parAtt(cid, s);
      case "👨‍👩‍👧 Мои дети":           return parKids(cid, s);
      case "➕ Привязать ребёнка":  set(cid,{step:"link_inn"}); return send(cid,"Введите *ИНН ребёнка* \\(12–14 цифр\\):",KB_X);
      default: await send(cid,"Используйте кнопки меню 👇",KB_PAR);
    }
  }

  async function parKids(cid, s) {
    const inns=s.inns||[];
    if(!inns.length){ await send(cid,"Нет привязанных детей\\. Нажмите *Привязать ребёнка*\\.",KB_PAR); return; }
    const stus=await Student.find({inn:{$in:inns}}).populate("group");
    const lines=stus.map(stu=>`👤 *${ev(stu.fio)}*\n   📂 ${ev(stu.group?.name||"—")} \\| ${ev(stu.group?.profRu||"—")}`);
    await send(cid,`👨‍👩‍👧 *Мои дети:*\n\n${lines.join("\n\n")||"Не найдено\\."}`,KB_PAR);
  }

  async function parAtt(cid, s) {
    const inns=s.inns||[];
    if(!inns.length){ await send(cid,"❌ Нет привязанных детей\\."); return; }
    const month=nowKey();
    const [yr,mn]=month.split("-").map(Number);
    const dim=new Date(yr,mn,0).getDate(),tdk=getTodayKey();
    for(const inn of inns){
      const stu=await Student.findOne({inn});
      if(!stu) continue;
      const g=await Group.findById(stu.group);
      const [sH,sM]=(g?.shiftStart||"09:30").split(":").map(Number),stm=sH*60+sM;
      let p=0,a=0,l=0; const missed=[];
      for(let d=1;d<=dim;d++){
        const dk=`${yr}-${pad2(mn)}-${pad2(d)}`;
        const iw=new Date(Date.UTC(yr,mn-1,d)).getUTCDay()%6===0;
        if(iw||dk>tdk) continue;
        const skud=stu.skudDaily?.find(r=>r.date===dk),att=stu.attendance?.find(a=>a.date===dk);
        let here=false,late=false;
        if(skud){here=skud.present;if(skud.firstIn){const dt=new Date(skud.firstIn);if(!isNaN(dt)){const tm=dt.getUTCHours()*60+dt.getUTCMinutes()+360;late=Math.floor(tm/60)%24*60+tm%60>stm;}}}
        else if(att){here=att.present;}
        if(here&&late){p++;l++;missed.push(`${d} ⏰`);}
        else if(here){p++;}
        else{a++;missed.push(`${d} ❌`);}
      }
      await send(cid,
        `🗓 *Посещаемость — ${ev(stu.fio)}*\n📂 ${ev(g?.name||"—")} \\| ${ev(MONTHS[mn-1])} ${yr}\n\n`+
        `✅ Присутствий: *${p}*\n❌ Отсутствий: *${a}*\n⏰ Опозданий: *${l}*\n\n`+
        (missed.length?`Пропуски и опоздания: ${ev(missed.join(", "))}`:"_Посещал все дни ✨_"),
        KB_PAR
      );
    }
  }

  /* ══════════════════════════════════════════════════════════════
     STEP HANDLER
  ══════════════════════════════════════════════════════════════ */
  async function handleStep(cid, txt, msg, s, st) {
    /* ── ИНН ── */
    if(st.step==="link_inn"){
      const inn=txt.replace(/\D/g,"");
      if(inn.length<12){ await send(cid,"❌ ИНН — минимум 12 цифр\\. Попробуйте ещё раз:"); return; }
      const stu=await Student.findOne({inn}).populate("group");
      if(!stu){ await send(cid,`❌ Студент с ИНН *${ev(inn)}* не найден\\. Проверьте и повторите:`); return; }
      await TgSub.findOneAndUpdate({chatId:String(cid)},{$addToSet:{linkedInns:inn},name:st.name||msg.from?.first_name},{upsert:true});
      const inns=[...new Set([...(st.inns||[]),inn])];
      SESS.set(String(cid),{role:"parent",name:st.name||msg.from?.first_name,inns,step:"idle"});
      await send(cid,
        `✅ *Привязка выполнена\\!*\n\n👤 *${ev(stu.fio)}*\n📂 ${ev(stu.group?.name||"—")}\n\n`+
        `Вы будете получать уведомления о входе/выходе ребёнка\\.`,
        KB_PAR
      );
      return;
    }

    /* ── Admin search ── */
    if(st.step==="a_search"){ clearStep(cid); await adminSearch(cid,txt); return; }

    /* ── Curator search ── */
    if(st.step==="c_search"){ clearStep(cid); await curSearch(cid,txt,s); return; }

    /* ── Добавление студента ── */
    if(st.step==="c_stu_add"){
      const parts=txt.split("/").map(p=>p.trim());
      const fio=parts[0],inn=parts[1]||"",birth=parts[2]||"",lycId=parts[3]||"";
      if(!fio||fio.length<3){ await send(cid,"❌ Введите хотя бы ФИО\\. Попробуйте ещё раз:"); return; }
      const gid=st.gid;
      if(!canAccessGroup(s,gid)){ await send(cid,"❌ Нет доступа\\."); clearStep(cid); return; }
      await Student.create({fio,inn:inn||undefined,birthDate:birth||undefined,lyceumId:lycId||undefined,group:gid});
      clearStep(cid);
      await send(cid,`✅ Студент *${ev(fio)}* добавлен\\.`,KB_CUR());
      await showStudentList(cid,gid);
      return;
    }

    /* ── Добавление предмета ── */
    if(st.step==="c_subj_add"){
      const label=txt.trim();
      if(!label||label.length<2){ await send(cid,"❌ Слишком короткое название\\. Попробуйте ещё раз:"); return; }
      const gid=st.gid, section=st.subjSection||"theory";
      if(!canAccessGroup(s,gid)){ await send(cid,"❌ Нет доступа\\."); clearStep(cid); return; }
      const g=await Group.findById(gid);
      const subjs=await ensureGroupSummarySubjects(g);
      const key=label.toLowerCase().replace(/[^a-zа-яё0-9]+/gi,"_").replace(/^_+|_+$/g,"")||`subj_${Date.now()}`;
      subjs.push({key,label,section});
      g.summarySubjects=subjs;
      await g.save();
      clearStep(cid);
      await send(cid,`✅ Предмет *${ev(label)}* добавлен\\.`,KB_CUR());
      await showSubjectList(cid,gid);
      return;
    }

    /* ── Group → month ── */
    const groupSteps = {
      "a_ved_g":"a_ved_m","a_xl_g":"a_xl_m",
      "c_xl_g":"c_xl_m","c_grades_g":"c_grades_m",
      "c_send_g":"c_send_m","fill_group_sel":"fill_month",
      "c_stu_list_grp":"c_stu_list_show","c_subj_grp":"c_subj_show",
    };
    if(groupSteps[st.step]){
      const gids=s.role==="curator"?s.groupIds:null;
      const filter=gids?{name:new RegExp(`^${txt}$`,"i"),_id:{$in:gids}}:{name:new RegExp(`^${txt}$`,"i")};
      const g=await Group.findOne(filter);
      if(!g){ await send(cid,"❌ Группа не найдена\\.",KB_X); return; }
      const next=groupSteps[st.step];
      if(next==="c_stu_list_show"){ set(cid,{gid:String(g._id),step:next}); await showStudentList(cid,String(g._id)); return; }
      if(next==="c_subj_show"){    set(cid,{gid:String(g._id),step:next}); await showSubjectList(cid,String(g._id)); return; }
      set(cid,{step:next,gid:String(g._id)});
      await send(cid,"📅 Выберите месяц:",monthKb(last6()));
      return;
    }

    /* ── Month steps ── */
    const month=parseM(txt);

    if(st.step==="a_ved_m"){
      if(!month){ await send(cid,"⚠️ Выберите месяц из списка\\."); return; }
      clearStep(cid); await sendWord(cid,st.gid,month,KB_ADMIN); return;
    }
    if(st.step==="a_xl_m"){
      if(!month){ await send(cid,"⚠️ Выберите месяц из списка\\."); return; }
      clearStep(cid); await sendXlsx(cid,st.gid,month,KB_ADMIN); return;
    }
    if(st.step==="c_xl_m"){
      if(!month){ await send(cid,"⚠️ Выберите месяц из списка\\."); return; }
      if(!canAccessGroup(s,st.gid)){ await send(cid,"❌ Нет доступа\\."); clearStep(cid); return; }
      clearStep(cid); await sendXlsx(cid,st.gid,month,KB_CUR()); return;
    }
    if(st.step==="c_grades_m"){
      if(!month){ await send(cid,"⚠️ Выберите месяц из списка\\."); return; }
      clearStep(cid); await showGrades(cid,st.gid,month,s); return;
    }
    if(st.step==="c_send_m"){
      if(!month){ await send(cid,"⚠️ Выберите месяц из списка\\."); return; }
      clearStep(cid); await sendVedToAdmins(cid,st.gid,month); return;
    }
    if(st.step==="fill_month"){
      if(!month){ await send(cid,"⚠️ Выберите месяц из списка\\."); return; }
      await initFill(cid,st.gid,month,s); return;
    }

    clearStep(cid); await send(cid,"Главное меню:",mkb(s.role));
  }

  // Уведомить всех (для cron)
  async function notifyAll(msg) {
    try {
      const subs=await TgSub.find();
      for(const s of subs) bot.sendMessage(s.chatId,msg,{parse_mode:"Markdown"}).catch(()=>null);
    } catch(e){console.error(e);}
  }

  // Уведомить родителей конкретного ребёнка (по ИНН) — MarkdownV2
  async function notifyParents(inn, msg) {
    try {
      const subs=await TgSub.find({role:"parent",linkedInns:inn});
      for(const s of subs) bot.sendMessage(s.chatId,msg,{parse_mode:"MarkdownV2"}).catch(()=>null);
    } catch(e){console.error(e);}
  }

  // Уведомить кураторов конкретной группы (только у кого notifyLate=true)
  async function notifyCuratorsOfGroup(groupId, msg) {
    try {
      const cur=await Curator.findOne({groups:groupId}).select("_id");
      if(!cur) return;
      const subs=await TgSub.find({role:"curator",curatorId:String(cur._id)});
      for(const s of subs) bot.sendMessage(s.chatId,msg,{parse_mode:"Markdown"}).catch(()=>null);
    } catch(e){console.error(e);}
  }

  // Уведомить всех админов
  async function notifyAdmins(msg) {
    try {
      const subs=await TgSub.find({role:"admin"});
      for(const s of subs) bot.sendMessage(s.chatId,msg,{parse_mode:"Markdown"}).catch(()=>null);
    } catch(e){console.error(e);}
  }

  console.log("✅ Telegram bot v3.0 запущен");
  console.log("   Token:", TG_TOKEN.slice(0,12)+"...");
  console.log("   Admin secret:", ADMIN_SECRET);
  console.log("   Bot username:", BOT_USERNAME);
  console.log("   Войти как админ: /admin " + ADMIN_SECRET);
  return { bot, notifyAll, notifyParents, notifyAdmins, notifyCurators: notifyCuratorsOfGroup };
};