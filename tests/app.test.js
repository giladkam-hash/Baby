/* Browser tests of the journal (needs: npm i playwright; a local static server is started here). */
const { chromium, devices } = require('playwright'); const assert = require('assert'), http = require('http'), fs = require('fs'), path = require('path');
const root = path.join(__dirname, '..'); const types = { '.html': 'text/html', '.js': 'text/javascript', '.png': 'image/png', '.webmanifest': 'application/json', '.gs': 'text/plain' };
const server = http.createServer((q, r) => { let f = path.join(root, decodeURIComponent(q.url.split('?')[0])); if (f.endsWith(path.sep)) f += 'index.html'; fs.readFile(f, (e, d) => { if (e) { r.statusCode = 404; return r.end(); } r.setHeader('content-type', types[path.extname(f)] || 'application/octet-stream'); r.end(d); }); });
let passed = 0; const test = async (name, fn) => { try { await fn(); passed++; console.log('ok   ' + name); } catch (e) { console.log('FAIL ' + name + '\n     ' + e.message); process.exitCode = 1; } };
(async () => {
  await new Promise(r => server.listen(0, r)); const base = 'http://localhost:' + server.address().port + '/';
  const browser = await chromium.launch(); const newPage = async (setup, nstatus) => {
    const ctx = await browser.newContext({ ...devices['iPhone 13'], locale: 'he-IL' }), p = await ctx.newPage(); p.errs = [];
    p.on('pageerror', e => p.errs.push(e.message));
    await p.route('https://script.google.com/**', route => { const u = new URL(route.request().url()), a = u.searchParams.get('action'); const body = a === 'nstatus' ? (nstatus || { ok: true, ver: '10', last: Date.now(), trigger: true, token: true, topic: true, sendOk: 3, sendFail: 0, lastOk: Date.now(), lastFail: 0, runSec: 120 }) : { ok: true, total: 0, rows: [] }; route.fulfill({ contentType: 'application/json', body: JSON.stringify(body) }); });
    await p.goto(base); if (setup) { await p.evaluate(setup); await p.reload(); } await p.waitForTimeout(600); return p;
  };
  const withCfg = extra => `(()=>{localStorage.setItem('babylog.v1.cfg',JSON.stringify({url:'https://script.google.com/macros/s/x/exec'}));localStorage.setItem('babylog.v1',JSON.stringify({days:{},questions:{},settings:${JSON.stringify(extra)},timers:{}}))})()`;

  await test('app loads, tabs open, no script errors', async () => { const p = await newPage(); for (let i = 1; i <= 5; i++) { await p.click(`#tabs button:nth-child(${i})`); await p.waitForTimeout(150); } assert.deepStrictEqual(p.errs, []); });
  await test('settings: remote change of ONE field does not overwrite the others', async () => {
    const p = await newPage(withCfg({ ntfy: 'chan', feedEvery: 3, fu: { ntfy: 1000, feedEvery: 1000 }, u: 1000 }));
    const r = await p.evaluate(() => { applyRemote({ k: 'set', id: 'settings', u: 5000, d: { ntfy: '', feedEvery: 3, autoWaitMin: 22, fu: { autoWaitMin: 5000 }, u: 5000 } }); return S.settings; });
    assert.strictEqual(r.ntfy, 'chan'); assert.strictEqual(r.autoWaitMin, 22);
  });
  await test('settings: a newer change to the same field wins; an older one loses', async () => {
    const p = await newPage(withCfg({ ntfy: 'chan', fu: { ntfy: 3000 }, u: 3000 }));
    const r = await p.evaluate(() => { applyRemote({ k: 'set', id: 'settings', u: 2000, d: { ntfy: 'old', fu: { ntfy: 2000 }, u: 2000 } }); const a = S.settings.ntfy; applyRemote({ k: 'set', id: 'settings', u: 4000, d: { ntfy: 'new', fu: { ntfy: 4000 }, u: 4000 } }); return [a, S.settings.ntfy]; });
    assert.deepStrictEqual(r, ['chan', 'new']);
  });
  await test('settings: saving stamps only the fields that changed', async () => {
    const p = await newPage(withCfg({ ntfy: 'chan', momName: '', fu: { ntfy: 10, momName: 10 }, u: 10 }));
    const fu = await p.evaluate(() => { S.settings = { ...S.settings, momName: 'קרן' }; saveSettings(); return S.settings.fu; });
    assert.strictEqual(fu.ntfy, 10); assert.ok(fu.momName > 10);
  });
  await test('settings: an older app (no per-field times) still syncs', async () => {
    const p = await newPage(withCfg({ ntfy: 'chan', fu: { ntfy: 10 }, u: 10 }));
    const r = await p.evaluate(() => { applyRemote({ k: 'set', id: 'settings', u: 9000, d: { ntfy: 'legacy', u: 9000 } }); return S.settings.ntfy; }); assert.strictEqual(r, 'legacy');
  });
  await test('health: no banner when everything is fine', async () => {
    const p = await newPage(withCfg({ ntfy: 'chan' })); await p.evaluate(() => checkHealth()); await p.evaluate(() => checkHealth()); await p.waitForTimeout(300); assert.strictEqual(await p.locator('.banner:has-text("לא תקינות")').count(), 0);
  });
  await test('health: banner appears (after two failed checks) when the server stopped', async () => {
    const p = await newPage(withCfg({ ntfy: 'chan' }), { ok: true, ver: '10', last: Date.now() - 40 * 6e4, trigger: true, token: true, topic: true });
    await p.evaluate(() => checkHealth()); await p.evaluate(() => checkHealth()); await p.waitForTimeout(300); const t = await p.locator('.banner:has-text("לא תקינות")').first().textContent(); assert.ok(t.includes('לא בודק'), t);
  });
  await test('health: old server code and missing token are reported', async () => {
    const p = await newPage(withCfg({ ntfy: 'chan' }), { ok: true, ver: '7', last: Date.now(), trigger: true, token: false, topic: true });
    const issues = await p.evaluate(async () => (await ntfyHealth()).issues.join('|')); assert.ok(issues.includes('לא מעודכן') && issues.includes('טוקן'), issues);
  });
  await test('system check lists every item', async () => {
    const p = await newPage(withCfg({ ntfy: 'chan' })); await p.click('#tabs button:nth-child(5)'); await p.click('[data-pcat="notif"]'); await p.click('#ntfyStatus'); await p.waitForTimeout(600);
    const t = await p.locator('#ntfyMsg').textContent(); assert.ok(t.includes('✓') && !t.includes('✗'), t);
  });
  await test('a diaper report goes through the review screen and is saved', async () => {
    const p = await newPage(); await p.click('[data-open="diaper"]'); await p.click('[data-pick="pee"]'); await p.waitForTimeout(200); await p.click('#rvSave'); await p.waitForTimeout(300);
    assert.ok((await p.evaluate(() => Object.values(S.days).flatMap(d => Object.values(d.ev)).length)) === 1);
  });
  await test('weight calculation: mother+baby minus mother, with confirmation', async () => {
    const p = await newPage(withCfg({ momName: 'קרן', name: 'נועה' })); await p.click('[data-open="weight"]'); await p.click('[data-wmode="calc"]'); await p.locator('#wMom').pressSequentially('62.4'); await p.locator('#wBoth').pressSequentially('65.5');
    assert.strictEqual(await p.inputValue('#wBaby'), '3.100'); await p.click('#wOk'); await p.click('#wSave'); await p.waitForTimeout(300); assert.strictEqual(await p.evaluate(() => lastWeight().kg), 3.1);
  });
  await browser.close(); server.close(); console.log(`\n${passed} passed`);
})();
