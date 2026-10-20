// Prévia local do sistema: node ferramentas/preview.js  →  http://localhost:5180
const http = require('http');
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..', 'apps-script');
const PORTA = Number(process.env.PORT) || 5180;

function ler(nome) {
  return fs.readFileSync(path.join(RAIZ, nome), 'utf8');
}

function montarIndex() {
  return ler('Index.html').replace(/<\?!=\s*include\('(\w+)'\);?\s*\?>/g, (_, nome) => {
    const html = ler(nome + '.html');
    if (nome !== 'App') return html;
    return '<script src="/mock.js"></script><script src="/Code.gs"></script>\n' + html;
  });
}

http.createServer((req, res) => {
  try {
    const url = req.url.split('?')[0];
    // Arquivos crus para colar no editor do Apps Script (script.google.com) sem digitar.
    if (url.indexOf('/raw/') === 0) {
      const cors = {
        'Access-Control-Allow-Origin': 'https://script.google.com',
        'Access-Control-Allow-Private-Network': 'true',
        'Access-Control-Allow-Methods': 'GET',
      };
      if (req.method === 'OPTIONS') { res.writeHead(204, cors); return res.end(); }
      const nome = path.basename(url.slice(5));
      res.writeHead(200, Object.assign({ 'Content-Type': 'text/plain; charset=utf-8' }, cors));
      return res.end(ler(nome));
    }
    if (url === '/') {
      res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
      return res.end(montarIndex());
    }
    if (url === '/mock.js') {
      res.writeHead(200, { 'Content-Type': 'text/javascript; charset=utf-8' });
      return res.end(fs.readFileSync(path.join(__dirname, 'mock-apps-script.js'), 'utf8'));
    }
    if (/^\/\w+\.gs$/.test(url)) {
      res.writeHead(200, { 'Content-Type': 'text/javascript; charset=utf-8' });
      return res.end(ler(url.slice(1)));
    }
    res.writeHead(404);
    res.end('não encontrado');
  } catch (e) {
    res.writeHead(500, { 'Content-Type': 'text/plain; charset=utf-8' });
    res.end(String(e.stack || e));
  }
}).listen(PORTA, '127.0.0.1', () => console.log('Prévia em http://localhost:' + PORTA));
