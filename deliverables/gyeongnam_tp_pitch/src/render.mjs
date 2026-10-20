import { chromium } from 'playwright';
import http from 'http'; import fs from 'fs'; import path from 'path';
const root = process.cwd();
const srv = http.createServer((q, r) => { const p = path.join(root, decodeURIComponent(q.url.split('?')[0])); if (!fs.existsSync(p)) { r.writeHead(404); return r.end(); }
  r.writeHead(200, { 'Content-Type': p.endsWith('.js') ? 'text/javascript' : 'text/html' }); fs.createReadStream(p).pipe(r); }).listen(8765);
const jobs = JSON.parse(process.argv[2]);
const b = await chromium.launch({ args: ['--use-gl=angle', '--use-angle=swiftshader', '--enable-unsafe-swiftshader', '--ignore-gpu-blocklist'] });
const anchors = fs.existsSync('img/anchors.json') ? JSON.parse(fs.readFileSync('img/anchors.json')) : {};
for (const [name, qs] of jobs) {
  const pg = await b.newPage(); pg.on('console', m => console.log(name, m.text())); pg.on('pageerror', e => console.log(name, 'ERR', e.message));
  await pg.goto(`http://localhost:8765/r3d/index.html?${qs}`); await pg.waitForFunction('window.__done', null, { timeout: 300000 });
  const { png, a } = await pg.evaluate(() => ({ png: window.__png, a: window.__anchors }));
  fs.writeFileSync(`img/${name}.png`, Buffer.from(png.split(',')[1], 'base64')); anchors[name] = a; console.log('ok', name); await pg.close();
}
fs.writeFileSync('img/anchors.json', JSON.stringify(anchors, null, 1)); await b.close(); srv.close();
