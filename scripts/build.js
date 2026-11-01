'use strict';
// Builds public/index.html from src/ and data/. Run: npm run build
const fs = require('fs');
const path = require('path');
const root = path.join(__dirname, '..');
const read = (p) => fs.readFileSync(path.join(root, p), 'utf8');
const dataFile = process.argv[2] || 'data/demo-data.json';
const data = read(dataFile).replace(/</g, '\\u003c');
const scripts = ['src/shim.js', 'src/engine.js', 'src/model.js', 'src/exports.js', 'src/app.js'].map((f) => `<script>\n${read(f)}\n</script>`).join('\n');
const brand = '<svg class="brand-mark" viewBox="0 0 32 32" aria-hidden="true"><rect width="32" height="32" rx="8" fill="var(--accent)"/><path d="M8 22c4 0 4-12 8-12s4 8 8 8" stroke="var(--accent-ink)" stroke-width="2.4" fill="none" stroke-linecap="round" stroke-dasharray="0.1 4.4"/><circle cx="8" cy="22" r="2.7" fill="var(--accent-ink)"/><circle cx="24" cy="18" r="2.6" fill="none" stroke="var(--accent-ink)" stroke-width="2"/></svg>';
const html = `<!doctype html>
<html lang="sv">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<title>Spårbar Klimatportal (demo)</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700;900&family=Roboto+Mono:wght@400;500&display=swap">
<style>
body { margin: 0; }
img { max-width: 100%; }
[hidden] { display: none !important; }
${read('src/style.css')}
</style>
</head>
<body>
<div class="app">
  <aside class="rail">
    <div class="brand">${brand}<div><div class="brand-name">Spårbar</div><div class="brand-sub">Klimatrapporter för transporter</div></div></div>
    <nav class="nav" id="nav" aria-label="Huvudmeny"></nav>
    <div class="rail-foot" id="railfoot"></div>
  </aside>
  <main class="main">
    <div class="ctxbar" id="ctxbar"></div>
    <div class="view" id="view"></div>
  </main>
</div>
<div id="tip" hidden></div>
<div id="drawer-root"></div>
<div id="toast-root" aria-live="polite"></div>
<script type="application/json" id="sample-data">${data}</script>
${scripts}
</body>
</html>
`;
fs.writeFileSync(path.join(root, 'public', 'index.html'), html);
console.log(`public/index.html built from ${dataFile} (${Math.round(html.length / 1024)} KB)`);
