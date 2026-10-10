// Минимальная имитация mongoose в памяти: только то, что используют модули СКУД.
// Нужна, чтобы гонять тесты без установленной MongoDB. Это не замена настоящей БД.
const crypto = require("crypto");
const isOid = (s) => /^[0-9a-f]{24}$/i.test(String(s || ""));
const newId = () => crypto.randomBytes(12).toString("hex");

class Schema {
  constructor(def, opts = {}) { this.def = def; this.opts = opts; }
  index() { return this; }
}
Schema.Types = { ObjectId: "ObjectId", Mixed: "Mixed" };

const clone = (v) => {
  if (Buffer.isBuffer(v)) return Buffer.from(v);
  if (v instanceof Date) return new Date(v);
  if (Array.isArray(v)) return v.map(clone);
  if (v && typeof v === "object") return Object.fromEntries(Object.entries(v).map(([k, x]) => [k, clone(x)]));
  return v;
};
const get = (o, path) => path.split(".").reduce((a, k) => (a == null ? undefined : a[k]), o);
const norm = (v) => (v instanceof Date ? v.getTime() : v);
const eq = (a, b) => {
  if (b instanceof RegExp) return typeof a === "string" && b.test(a);
  if (b === null) return a === null || a === undefined;
  if (a instanceof Date || b instanceof Date) return norm(a) === norm(b);
  return String(a) === String(b);
};

function evalExpr(e, doc) {
  if (typeof e === "string" && e.startsWith("$")) return get(doc, e.slice(1));
  if (e && typeof e === "object" && !Array.isArray(e)) {
    const [op, args] = Object.entries(e)[0];
    const [x, y] = args.map((a) => evalExpr(a, doc));
    if (op === "$lt") return x < y;
    if (op === "$lte") return x <= y;
    if (op === "$gt") return x > y;
    throw new Error("expr " + op);
  }
  return e;
}

function match(doc, q) {
  for (const [k, c] of Object.entries(q || {})) {
    if (k === "$or") { if (!c.some((x) => match(doc, x))) return false; continue; }
    if (k === "$and") { if (!c.every((x) => match(doc, x))) return false; continue; }
    if (k === "$expr") { if (!evalExpr(c, doc)) return false; continue; }
    const v = get(doc, k);
    if (c && typeof c === "object" && !(c instanceof RegExp) && !(c instanceof Date) && !Buffer.isBuffer(c)) {
      for (const [op, x] of Object.entries(c)) {
        const ok =
          op === "$in" ? x.some((y) => eq(v, y)) :
          op === "$nin" ? !x.some((y) => eq(v, y)) :
          op === "$ne" ? !eq(v, x) :
          op === "$exists" ? (v !== undefined) === !!x :
          op === "$lt" ? v != null && norm(v) < norm(x) :
          op === "$lte" ? v != null && norm(v) <= norm(x) :
          op === "$gt" ? v != null && norm(v) > norm(x) :
          op === "$gte" ? v != null && norm(v) >= norm(x) :
          (() => { throw new Error("op " + op); })();
        if (!ok) return false;
      }
    } else if (!eq(v, c)) return false;
  }
  return true;
}

function applyUpdate(doc, upd, inserting) {
  const ops = { $set: {}, $setOnInsert: {}, $inc: {}, $unset: {}, $addToSet: {} };
  for (const [k, v] of Object.entries(upd || {})) { if (k in ops) Object.assign(ops[k], v); else ops.$set[k] = v; }
  Object.assign(doc, clone(ops.$set));
  if (inserting) Object.assign(doc, clone(ops.$setOnInsert));
  for (const [k, v] of Object.entries(ops.$inc)) doc[k] = (doc[k] || 0) + v;
  for (const k of Object.keys(ops.$unset)) delete doc[k];
  for (const [k, v] of Object.entries(ops.$addToSet)) {
    doc[k] = Array.isArray(doc[k]) ? doc[k] : [];
    if (!doc[k].some((x) => String(x) === String(v))) doc[k].push(clone(v));
  }
}

function sorter(sort) {
  const keys = typeof sort === "string" ? sort.split(" ").filter(Boolean).map((k) => (k[0] === "-" ? [k.slice(1), -1] : [k, 1])) : Object.entries(sort || {});
  return (a, b) => {
    for (const [k, d] of keys) {
      const x = norm(get(a, k)), y = norm(get(b, k));
      if (x === y) continue;
      if (x === undefined) return -d; if (y === undefined) return d;
      return (x < y ? -1 : 1) * d;
    }
    return 0;
  };
}

function project(doc, sel) {
  if (!sel) return doc;
  const fields = String(sel).split(/\s+/).filter(Boolean);
  const out = { ...doc };
  if (fields.every((f) => f[0] === "-")) { for (const f of fields) delete out[f.slice(1)]; return out; }
  const inc = fields.filter((f) => f[0] !== "-");
  return Object.fromEntries(Object.entries(out).filter(([k]) => k === "_id" || inc.includes(k)));
}

const models = {};
function model(name, schema) {
  const store = [];
  const def = schema.def, ts = !!schema.opts.timestamps;
  const defaults = () => {
    const o = {};
    for (const [k, f] of Object.entries(def)) {
      if (f && typeof f === "object" && !Array.isArray(f) && "default" in f) o[k] = typeof f.default === "function" ? f.default() : clone(f.default);
      if (Array.isArray(f)) o[k] = [];
    }
    return o;
  };
  const wrap = (d) => (d ? Object.assign(Object.create({ toObject: () => clone(d) }), clone(d)) : d);

  class Query {
    constructor(run) { this.run = run; this.o = {}; }
    sort(s) { this.o.sort = s; return this; }
    limit(n) { this.o.limit = n; return this; }
    skip(n) { this.o.skip = n; return this; }
    select(s) { this.o.sel = s; return this; }
    lean() { this.o.lean = true; return this; }
    populate(path) { (this.o.pop ||= []).push(path); return this; }
    then(a, b) { return Promise.resolve().then(() => this.run(this.o)).then((r) => populateAll(r, this.o.pop)).then(a, b); }
    catch(b) { return this.then(undefined, b); }
  }
  // Раскрыть ссылки (ref) для указанных полей: только то, что нужно тестам
  const refOf = (path) => { const f = def[path]; const x = Array.isArray(f) ? f[0] : f; return x && typeof x === "object" ? x.ref : null; };
  const populateAll = (r, paths) => {
    if (!paths || r == null) return r;
    const one = (d) => {
      if (!d || typeof d !== "object") return d;
      for (const p of paths) {
        const ref = refOf(p); const M2 = ref && models[ref]; if (!M2) continue;
        const find = (id) => { const x = M2._store.find((y) => String(y._id) === String(id)); return x ? clone(x) : null; };
        d[p] = Array.isArray(d[p]) ? d[p].map(find).filter(Boolean) : d[p] != null ? find(d[p]) : d[p];
      }
      return d;
    };
    return Array.isArray(r) ? r.map(one) : one(r);
  };
  const finish = (list, o) => {
    let r = list;
    if (o.sort) r = [...r].sort(sorter(o.sort));
    if (o.skip) r = r.slice(o.skip);
    if (o.limit) r = r.slice(0, o.limit);
    return r.map((d) => (o.lean ? clone(project(d, o.sel)) : wrap(project(d, o.sel))));
  };

  const M = {
    modelName: name,
    _store: store,
    find: (q, sel) => new Query((o) => finish(store.filter((d) => match(d, q)), { ...o, sel: o.sel || sel })),
    findOne: (q, sel) => new Query((o) => { const r = finish([...store.filter((d) => match(d, q))].sort(o.sort ? sorter(o.sort) : () => 0), { ...o, limit: 1, sel: o.sel || sel }); return r[0] || null; }),
    findById: (id, sel) => M.findOne({ _id: id }, sel),
    exists: (q) => new Query(() => (store.some((d) => match(d, q)) ? { _id: "x" } : null)),
    countDocuments: (q) => new Query(() => store.filter((d) => match(d, q)).length),
    create: async (data) => {
      const arr = Array.isArray(data);
      const out = (arr ? data : [data]).map((x) => {
        const d = { ...defaults(), ...clone(x) };
        if (d._id === undefined) d._id = newId();
        for (const [k, f] of Object.entries(def)) {
          if (f && f.unique && d[k] !== undefined && store.some((s) => String(s[k]) === String(d[k]))) throw Object.assign(new Error("E11000 duplicate key " + k), { code: 11000 });
          if (f && f.required && d[k] == null) throw new Error(`${name}: ${k} required`);
          if (f && Array.isArray(f.enum) && d[k] != null && !f.enum.includes(d[k])) throw new Error(`${name}.${k}: bad enum ${d[k]}`);
        }
        if (ts) d.createdAt = d.updatedAt = new Date();
        store.push(d);
        return wrap(d);
      });
      return arr ? out : out[0];
    },
    updateOne: (q, u) => new Query(() => {
      const d = store.find((x) => match(x, q)); if (!d) return { matchedCount: 0, modifiedCount: 0 };
      applyUpdate(d, u); if (ts) d.updatedAt = new Date(); return { matchedCount: 1, modifiedCount: 1 };
    }),
    updateMany: (q, u) => new Query(() => {
      const l = store.filter((x) => match(x, q));
      for (const d of l) { applyUpdate(d, u); if (ts) d.updatedAt = new Date(); }
      return { matchedCount: l.length, modifiedCount: l.length };
    }),
    deleteOne: (q) => new Query(() => {
      const i = store.findIndex((x) => match(x, q));
      if (i >= 0) store.splice(i, 1);
      return { deletedCount: i >= 0 ? 1 : 0 };
    }),
    deleteMany: (q) => new Query(() => {
      let n = 0;
      for (let i = store.length - 1; i >= 0; i--) if (match(store[i], q)) { store.splice(i, 1); n++; }
      return { deletedCount: n };
    }),
    findOneAndUpdate: (q, u, opt = {}) => new Query((o) => {
      const sort = opt.sort || o.sort;
      let d = [...store.filter((x) => match(x, q))].sort(sort ? sorter(sort) : () => 0)[0];
      let before = d && clone(d);
      if (!d) {
        if (!opt.upsert) return null;
        d = { ...defaults() };
        for (const [k, v] of Object.entries(q)) if (!k.startsWith("$") && (v === null || typeof v !== "object" || v instanceof Date)) d[k] = clone(v);
        if (d._id === undefined) d._id = newId();
        applyUpdate(d, u, true); if (ts) d.createdAt = d.updatedAt = new Date();
        store.push(d);
        return wrap(d);
      }
      applyUpdate(d, u); if (ts) d.updatedAt = new Date();
      return wrap(project(opt.new || opt.returnDocument === "after" ? d : before, o.sel));
    }),
  };
  M.findByIdAndUpdate = (id, u, opt) => M.findOneAndUpdate({ _id: id }, u, opt);
  M.findByIdAndDelete = (id) => new Query(() => { const i = store.findIndex((x) => String(x._id) === String(id)); return i >= 0 ? wrap(store.splice(i, 1)[0]) : null; });
  models[name] = M;
  return M;
}

module.exports = { Schema, model, models, isValidObjectId: isOid, _reset: () => Object.values(models).forEach((m) => (m._store.length = 0)) };
