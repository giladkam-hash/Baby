/* Loads apps-script/Code.gs with in-memory fakes of the Google services, so the notification rules can be tested. */
const fs = require('fs'), path = require('path');
const CODE = fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'Code.gs'), 'utf8');

function makeEnv(opts = {}) {
  const env = { now: opts.now || Date.now(), sent: [], status: 200, props: Object.assign({}, opts.props || {}), tz: 'Asia/Jerusalem', triggers: [] };
  const mk = init => { const data = init; return { data, getLastRow: () => data.length, setRightToLeft() {}, setFrozenRows() {}, clearContents() { data.length = 0; },
    getRange(r, c, nr, nc) { nr = nr || 1; nc = nc || 1; return { setValues(v) { for (let i = 0; i < nr; i++) { data[r - 1 + i] = data[r - 1 + i] || []; for (let j = 0; j < nc; j++) data[r - 1 + i][c - 1 + j] = v[i][j]; } return this; }, setFontWeight() { return this; }, setNumberFormat() { return this; },
      getValues() { const o = []; for (let i = 0; i < nr; i++) { o.push([]); for (let j = 0; j < nc; j++) o[i].push((data[r - 1 + i] || [])[c - 1 + j] ?? ''); } return o; } }; } }; };
  env.sheets = { 'יומן התינוקת': mk([['header']]) };
  env.log = env.sheets['יומן התינוקת'];
  const fmt = (d, tz, f) => { const p = new Intl.DateTimeFormat('en-CA', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false }).formatToParts(d).reduce((a, x) => (a[x.type] = x.value, a), {});
    const hh = p.hour === '24' ? '00' : p.hour; return f === 'yyyy-MM-dd' ? `${p.year}-${p.month}-${p.day}` : f === 'HH:mm' ? `${hh}:${p.minute}` : `${p.day}.${p.month} ${hh}:${p.minute}`; };
  const ctx = {
    SpreadsheetApp: { getActiveSpreadsheet: () => ({ getSheetByName: n => env.sheets[n] || null, insertSheet: n => env.sheets[n] = mk([]) }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    ContentService: { createTextOutput: t => ({ setMimeType() { return t; } }), MimeType: { JSON: 1 } },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => env.props[k] ?? null, setProperty: (k, v) => { if (String(v).length > 9000) throw new Error('property too big: ' + k); env.props[k] = String(v); }, deleteProperty: k => delete env.props[k], getProperties: () => ({ ...env.props }) }) },
    ScriptApp: { getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' }), getProjectTriggers: () => env.triggers, deleteTrigger() {}, newTrigger: h => ({ timeBased: () => ({ everyMinutes: () => ({ create() { env.triggers.push({ getHandlerFunction: () => h }); } }) }) }) },
    UrlFetchApp: { fetch: (u, o) => { const msg = JSON.parse(o.payload); const code = typeof env.status === 'function' ? env.status(msg) : env.status; if (code === 200) env.sent.push(msg); return { getResponseCode: () => code, getContentText: () => '{"error":"x"}' }; } },
    Utilities: { formatDate: fmt, base64Encode() { return ''; } },
    DriveApp: {},
  };
  const realNow = Date.now; Date.now = () => env.now;
  const api = new Function(...Object.keys(ctx), CODE + ';return {doPost,doGet,setupNotifications,notifyTick,NVER};')(...Object.values(ctx));
  env.api = api; env.restore = () => { Date.now = realNow; };
  const J = (k, id, u, d) => JSON.stringify({ k, id, u, d });
  const row = (id, stamp, j) => [id, stamp || String(Math.random()), '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', j];
  env.post = rows => JSON.parse(api.doPost({ postData: { contents: JSON.stringify({ action: 'append', rows }) } }));
  env.settings = (u, d) => env.post([row('settings', 's' + u + Math.random(), J('set', 'settings', u, d))]);
  env.feed = (id, side, startMin, endMin, dur, T0) => env.post([row(id, 'e' + id, J('ev', id, T0 + endMin * 6e4, { id, type: 'feed', side, lastSide: side, t: T0 + startMin * 6e4, end: T0 + endMin * 6e4, dur }))]);
  env.timer = (on, T0, m) => env.post([row('timers', 't' + m + Math.random(), J('tim', 'timers', T0 + m * 6e4, { feed: on ? { start: T0 + m * 6e4 } : null }))]);
  env.tick = () => api.notifyTick();
  env.titles = () => env.sent.map(m => m.title);
  env.J = J; env.row = row;
  return env;
}
module.exports = { makeEnv };
