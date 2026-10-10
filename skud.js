// СКУД: единая точка подключения всех модулей. Вызывается из server.js один раз.
const makeCore = require("./skud-core");

module.exports = function mountSkud(app, ctx) {
  const core = makeCore(ctx);
  const reg = require("./skud-register")(app, ctx, core);
  const agent = require("./skud-agent-api")(app, ctx, core);
  require("./skud-staff")(app, ctx, core);
  return { core, ...reg, ...agent };
};
