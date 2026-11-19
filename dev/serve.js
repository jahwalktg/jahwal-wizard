/* 문서 변신 마법사 — 미리보기 서버 (기관판 dev/serve.js와 같은 방식)
   GET  /...           이 폴더의 파일
   POST /_out/이름     화면이 만든 파일을 dev/_out/ 에 저장 (엑셀·한글 COM 검증용) */
const http = require('http'), fs = require('fs'), path = require('path');
const ROOT = path.resolve(__dirname, '..'), OUT = path.join(__dirname, '_out'), PORT = +process.argv[2] || 1520;
const MIME = { html: 'text/html; charset=utf-8', js: 'text/javascript; charset=utf-8', css: 'text/css; charset=utf-8', png: 'image/png', json: 'application/json' };
http.createServer((req, res) => {
  const u = decodeURIComponent(req.url.split('?')[0]);
  if (req.method === 'POST' && u.startsWith('/_out/')) {
    const name = path.basename(u), chunks = [];
    req.on('data', c => chunks.push(c));
    req.on('end', () => { fs.mkdirSync(OUT, { recursive: true }); fs.writeFileSync(path.join(OUT, name), Buffer.concat(chunks)); res.writeHead(200, { 'Content-Type': 'text/plain' }); res.end('ok ' + name); });
    return;
  }
  const p = path.join(ROOT, u === '/' ? 'index.html' : u);
  if (!p.startsWith(ROOT)) { res.writeHead(403); res.end(); return; }
  fs.readFile(p, (e, d) => {
    if (e) { res.writeHead(404); res.end('not found'); return; }
    res.writeHead(200, { 'Content-Type': MIME[path.extname(p).slice(1)] || 'application/octet-stream', 'Cache-Control': 'no-store' }); res.end(d);
  });
}).listen(PORT, '127.0.0.1', () => console.log('문서 변신 마법사 미리보기 http://localhost:' + PORT));
