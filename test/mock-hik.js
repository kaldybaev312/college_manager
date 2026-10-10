// Имитация терминала Hikvision: Digest-авторизация и нужные методы ISAPI.
const http = require("http");
const crypto = require("crypto");
const md5 = (s) => crypto.createHash("md5").update(s).digest("hex");

module.exports = function startMock({ user = "admin", pass = "secret" } = {}) {
  const state = { users: new Map(), faces: new Map(), log: [], failFace: false, down: false, nonce: crypto.randomBytes(8).toString("hex") };

  const server = http.createServer((req, res) => {
    const chunks = [];
    req.on("data", (c) => chunks.push(c));
    req.on("end", () => {
      if (state.down) return res.destroy();
      const body = Buffer.concat(chunks);
      const auth = req.headers.authorization || "";
      const kv = Object.fromEntries([...auth.matchAll(/(\w+)="?([^",]+)"?/g)].map((m) => [m[1], m[2]]));
      let ok = false;
      if (auth.startsWith("Digest") && kv.username === user) {
        const ha1 = md5(`${user}:${kv.realm}:${pass}`), ha2 = md5(`${req.method}:${kv.uri}`);
        ok = kv.response === md5(`${ha1}:${kv.nonce}:${kv.nc}:${kv.cnonce}:${kv.qop}:${ha2}`);
      }
      if (!ok) {
        res.writeHead(401, { "WWW-Authenticate": `Digest realm="mock", qop="auth", nonce="${state.nonce}", opaque="o"` });
        return res.end();
      }
      const url = new URL(req.url, "http://x");
      const p = url.pathname.replace("/ISAPI/", "");
      const send = (code, obj) => { res.writeHead(code, { "Content-Type": "application/json" }); res.end(JSON.stringify(obj)); };
      state.log.push(`${req.method} ${p}`);
      const text = body.toString("latin1");
      const jsonBody = () => { try { return JSON.parse(body.toString()); } catch { return {}; } };

      if (p === "System/deviceInfo") return send(200, { DeviceInfo: { deviceName: "mock" } });
      if (p === "AccessControl/UserInfo/Count") return send(200, { UserInfoCount: { userNumber: state.users.size } });
      if (p === "AccessControl/UserInfo/Record" || p === "AccessControl/UserInfo/Modify") {
        const u = jsonBody().UserInfo;
        const exists = state.users.has(u.employeeNo);
        if (p.endsWith("Record") && exists) return send(400, { statusCode: 6, statusString: "Invalid Content", subStatusCode: "employeeNoAlreadyExist" });
        if (p.endsWith("Modify") && !exists) return send(400, { statusCode: 6, subStatusCode: "employeeNoNotExist" });
        state.users.set(u.employeeNo, u);
        return send(200, { statusCode: 1, statusString: "OK" });
      }
      if (p === "Intelligent/FDLib/FaceDataRecord" || p === "Intelligent/FDLib/FDModify") {
        const m = text.match(/"FPID":\s*"(\d+)"/);
        const no = m && m[1];
        if (state.failFace) return send(400, { statusCode: 6, statusString: "Invalid Content", subStatusCode: "SubpicAnalysisModelingError", errorMsg: "saveFacePic" });
        if (!state.users.has(no)) return send(400, { statusCode: 6, subStatusCode: "employeeNoNotExist" });
        if (p.endsWith("Record") && state.faces.has(no)) return send(400, { statusCode: 6, subStatusCode: "faceAlreadyExist" });
        if (!body.includes(Buffer.from([0xff, 0xd8, 0xff]))) return send(400, { statusCode: 6, subStatusCode: "badImage" });
        state.faces.set(no, true);
        return send(200, { statusCode: 1, statusString: "OK" });
      }
      if (p === "AccessControl/UserInfo/Delete") {
        const list = jsonBody().UserInfoDelCond?.EmployeeNoList || [];
        for (const e of list) { state.users.delete(e.employeeNo); state.faces.delete(e.employeeNo); }
        return send(200, { statusCode: 1, statusString: "OK" });
      }
      if (p === "AccessControl/UserInfo/Search") {
        const c = jsonBody().UserInfoSearchCond;
        const all = [...state.users.values()];
        const page = all.slice(c.searchResultPosition, c.searchResultPosition + Math.min(c.maxResults, 2)); // отдаём по 2, как капризная прошивка
        const more = c.searchResultPosition + page.length < all.length;
        return send(200, { UserInfoSearch: { searchID: c.searchID, responseStatusStrg: more ? "MORE" : "OK", numOfMatches: page.length, totalMatches: all.length, UserInfo: page } });
      }
      send(404, { statusCode: 4, statusString: "Not Found" });
    });
  });
  return new Promise((resolve) => server.listen(0, "127.0.0.1", () => {
    state.ip = `127.0.0.1:${server.address().port}`;
    state.close = () => new Promise((r) => { server.closeAllConnections?.(); server.close(r); });
    resolve(state);
  }));
};
