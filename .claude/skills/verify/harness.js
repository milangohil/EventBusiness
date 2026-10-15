// Opens the real Index.html with google.script.run / google.script.history backed by real Code.gs.
const { chromium } = require(process.env.PLAYWRIGHT_PATH || '/opt/node22/lib/node_modules/playwright');
const SHIM = `(() => {
  const listeners = []; let handler = null;
  function runner(succ, fail) {
    return new Proxy({}, { get(_, prop) {
      if (prop === 'withSuccessHandler') return f => runner(f, fail);
      if (prop === 'withFailureHandler') return f => runner(succ, f);
      return (arg) => { window.__gas(prop, arg === undefined ? null : JSON.stringify(arg)).then(res => {
        const r = JSON.parse(res); setTimeout(() => r.ok ? succ && succ(r.value) : fail && fail(new Error(r.error)), 30); }); };
    } });
  }
  window.google = { script: { run: runner(null, null),
    history: { push(state, params, hash) { history.pushState(state, '', '#' + (hash || '')); }, replace(state, params, hash) { history.replaceState(state, '', '#' + (hash || '')); },
      setChangeHandler(fn) { handler = fn; } } } };
  window.addEventListener('popstate', e => handler && handler({ state: e.state, location: { hash: location.hash.slice(1) } }));
})();`;
async function openApp(browser, backend, user, viewport) {
  const ctx = await browser.newContext({ viewport: viewport || { width: 390, height: 844 }, deviceScaleFactor: 2 });
  const page = await ctx.newPage();
  page.errors = [];
  page.on('pageerror', e => page.errors.push('PAGEERR ' + e.message));
  page.on('console', m => m.type() === 'error' && page.errors.push(m.text()));
  page.calls = [];
  await page.exposeFunction('__gas', (fn, argJSON) => {
    page.calls.push(fn);
    try { return JSON.stringify({ ok: true, value: backend.call(user, fn, argJSON === null ? undefined : JSON.parse(argJSON)) }); }
    catch (e) { return JSON.stringify({ ok: false, error: e.message }); }
  });
  await page.addInitScript(SHIM);
  await page.goto('file:///home/user/EventBusiness/apps-script/Index.html');
  await page.waitForSelector('.kpis, .empty', { timeout: 15000 });
  return page;
}
module.exports = { openApp };
