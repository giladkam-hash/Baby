const assert = require('assert'), { makeEnv } = require('./gs-harness');
let passed = 0; const test = (name, fn) => { try { fn(); passed++; console.log('ok   ' + name); } catch (e) { console.log('FAIL ' + name + '\n     ' + e.message); process.exitCode = 1; } };
const BASE = { feedEvery: 2, notifyLeads: [], notifyDue: false, notifyEnd: true, notifyImmMin: 30, notifyWaitMin: '', autoWaitMin: 15, mergeMin: 40, ntfy: 'chan', tz: 'Asia/Jerusalem' };
const fresh = (extra = {}) => { const e = makeEnv({ props: { NTFY_TOKEN: 'tk_abcdefghijklmnop' } }); e.settings(e.now, { ...BASE, ...extra }); return e; };
const run = (e, from, to, T0, hooks = {}) => { for (let m = from; m <= to; m++) { e.now = T0 + m * 6e4; (hooks[m] || (() => {}))(); e.tick(); } };

test('split feed (net >= 30) produces ONE status note, at the end of the last part', () => {
  const e = fresh(), T0 = e.now; run(e, 0, 60, T0, { 20: () => e.feed('a', 'R', 0, 20, 20, T0), 25: () => e.timer(1, T0, 25), 40: () => { e.timer(0, T0, 40); e.feed('b', 'L', 25, 40, 15, T0); } });
  const n = e.titles().filter(t => t.startsWith('הנקה הסתיימה')); assert.strictEqual(n.length, 1, n.join(' | ')); // only after the second side ended (min 40), never after the first part alone
});
test('short single feed waits the (auto) 15 minutes, then notifies once', () => {
  const e = fresh(), T0 = e.now; run(e, 0, 40, T0, { 20: () => e.feed('a', 'L', 0, 20, 20, T0) });
  assert.strictEqual(e.sent.length, 1); assert.ok(e.sent[0].message.includes('צד הבא: ימין'), e.sent[0].message);
});
test('no note while a feeding timer is running', () => {
  const e = fresh(), T0 = e.now; run(e, 0, 30, T0, { 5: () => e.feed('a', 'L', 0, 5, 20, T0), 6: () => e.timer(1, T0, 6) }); assert.strictEqual(e.sent.length, 0);
});
test('long feed (>= 30 net) notifies immediately, then every 30 minutes until the next feed', () => {
  const e = fresh(), T0 = e.now; run(e, 0, 130, T0, { 35: () => e.feed('c', 'R', 0, 35, 35, T0) });
  assert.strictEqual(e.sent.length, 4, e.titles().join('|')); assert.ok(e.sent.every(m => m.sequence_id === 'feed-status'));
});
test('pre-feed alert at 15 and 10 minutes, due alert, once each', () => {
  const e = fresh({ notifyLeads: [15, 10], notifyDue: true, notifyEnd: false }), T0 = e.now; run(e, 0, 140, T0, { 20: () => e.feed('a', 'L', 0, 20, 40, T0) });
  assert.deepStrictEqual(e.titles().map(t => t.replace(/\d+ דק׳/, 'N דק׳')), ['הנקה בעוד N דק׳', 'הנקה בעוד N דק׳', 'הגיע הזמן להנקה']);
});
test('a rejected send (429) is retried and delivered exactly once', () => {
  const e = fresh(), T0 = e.now; let fails = 2; e.status = () => (fails-- > 0 ? 429 : 200); run(e, 0, 40, T0, { 20: () => e.feed('a', 'L', 0, 20, 40, T0) });
  assert.strictEqual(e.sent.length, 1); const st = JSON.parse(e.props.nstat); assert.strictEqual(st.fail, 2);
});
test('daily vitamin reminder goes out once, and not if it was already given', () => {
  const e = makeEnv({ props: { NTFY_TOKEN: 'tk_abcdefghijklmnop' } }), T0 = e.now; const hm = new Intl.DateTimeFormat('en-GB', { timeZone: 'Asia/Jerusalem', hour: '2-digit', minute: '2-digit', hour12: false }).format(new Date(T0 - 5 * 6e4));
  e.settings(T0, { ...BASE, notifyMeds: true, daily: [{ id: 'v1', label: 'ויטמין D', time: hm, notify: true }, { id: 'v2', label: 'ברזל', time: hm, who: 'mom', notify: true }, { id: 'v3', label: 'אומגה', time: hm, notify: false }] });
  e.post([e.row('m0', 'x', e.J('ev', 'm0', T0, { id: 'm0', type: 'mmed', name: 'ברזל', t: T0 - 6e4 }))]); run(e, 0, 5, T0);
  assert.deepStrictEqual(e.titles(), ['💊 ויטמין D לתינוקת']);
});
test('water reminder when behind the pace, silent when on pace', () => {
  const e = makeEnv({ props: { NTFY_TOKEN: 'tk_abcdefghijklmnop' } }), T0 = e.now; e.settings(T0, { ...BASE, notifyWater: true, water: { goal: 24, start: '00:00', end: '23:59' } }); e.tick(); assert.ok(e.titles().some(t => t.includes('מים')));
});
test('settings from two phones merge per field: a stale phone cannot wipe the channel', () => {
  const e = makeEnv({ props: { NTFY_TOKEN: 'tk_abcdefghijklmnop' } }), t = e.now;
  e.settings(t, { ...BASE, ntfy: 'chan', fu: { ntfy: t, feedEvery: t } });
  e.settings(t + 10, { ...BASE, ntfy: '', autoWaitMin: 22, fu: { autoWaitMin: t + 10 } }); // stale phone changes only autoWaitMin
  e.tick(); const set = JSON.parse(e.props.nset); assert.strictEqual(set.topic, 'chan'); assert.strictEqual(set.wait, 22);
});
test('exact duplicate rows are not stored twice', () => {
  const e = makeEnv(), r = e.row('x1', 'stamp', e.J('ev', 'x1', 1, { id: 'x1', type: 'feed', t: 1 })); e.post([r]); e.post([r]); e.post([r, r]); assert.strictEqual(e.log.data.length, 2);
});
test('removing rows from the log does not freeze notifications', () => {
  const e = fresh(), T0 = e.now; e.tick(); e.log.data.splice(1, 0); e.log.data.length = 1; e.timer(0, T0, 1); e.tick(); assert.ok(e.props.nrow);
});
test('status fields the app needs are served', () => {
  const e = fresh(); e.tick(); const st = JSON.parse(e.api.doGet({ parameter: { action: 'nstatus', topic: 'chan' } })); assert.strictEqual(st.ok, true); assert.strictEqual(st.token, true); assert.strictEqual(st.topic, true); assert.ok(st.ver && st.last);
});
test('own runtime is measured and the status tab is not rewritten every minute', () => {
  const e = fresh(), T0 = e.now; run(e, 0, 5, T0); const run0 = JSON.parse(e.props.nrun); assert.strictEqual(run0.n, 6); const writes = e.props.ndiagt; run(e, 6, 8, T0); assert.strictEqual(e.props.ndiagt, writes);
});
console.log(`\n${passed} passed`);
