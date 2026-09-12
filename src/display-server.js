/*
 * display-server: 表示用 web (結果ビュー) を配信する静的サーバ。
 * 設定UI(web-server.js, web_ui_port) とは別ポート(display_web_port) で動く。
 * webdist/ 配下の静的ファイルと data/bundle.json を返すだけ。
 */
const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const { webDistDir, displayWebPort } = require('./config');

const MIME = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon'
};

function startDisplayServer() {
    const server = http.createServer((req, res) => {
        try {
            let pathname = decodeURIComponent(url.parse(req.url).pathname);
            if (pathname === '') pathname = '/';

            let filePath = path.resolve(path.join(webDistDir, pathname));
            // ディレクトリトラバーサル対策
            if (!filePath.startsWith(path.resolve(webDistDir))) {
                res.writeHead(403); res.end('Forbidden'); return;
            }

            // ディレクトリは index.html を返す。末尾スラッシュ無しはスラッシュ付きへ誘導
            if (fs.existsSync(filePath) && fs.statSync(filePath).isDirectory()) {
                if (!pathname.endsWith('/')) {
                    res.writeHead(301, { Location: pathname + '/' });
                    res.end();
                    return;
                }
                filePath = path.join(filePath, 'index.html');
            }

            if (fs.existsSync(filePath) && fs.statSync(filePath).isFile()) {
                const ext = path.extname(filePath).toLowerCase();
                res.writeHead(200, {
                    'Content-Type': MIME[ext] || 'application/octet-stream',
                    'Access-Control-Allow-Origin': '*'
                });
                fs.createReadStream(filePath).pipe(res);
            } else {
                res.writeHead(404); res.end('Not Found');
            }
        } catch (e) {
            res.writeHead(500); res.end('Internal Server Error');
        }
    });

    server.on('error', (err) => {
        if (err && err.code === 'EADDRINUSE') {
            console.error(
                `[display-server] Port ${displayWebPort} is already in use. ` +
                `The display web will not start, but other processing continues.\n` +
                `  Fix: stop the existing process, or change display_web_port in config.json.`
            );
        } else {
            console.error('[display-server] server error:', err);
        }
    });

    server.listen(displayWebPort, () => {
        console.log(`Display web (results) running at http://localhost:${displayWebPort}`);
    });
    return server;
}

module.exports = { startDisplayServer };
