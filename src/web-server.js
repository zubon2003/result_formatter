const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const { config, eventsDir, configPath, loadConfig } = require('./config');

function startServer(reprocessCallback) {
    const server = http.createServer(async (req, res) => {
        const method = req.method;
        const baseDir = path.join(__dirname, '..');
        const parsedUrl = url.parse(req.url, true);
        const pathname = parsedUrl.pathname;

        try {
            if (pathname === '/') {
                const filePath = path.join(baseDir, 'index.html');
                const data = await fs.promises.readFile(filePath);
                res.writeHead(200, { 'Content-Type': 'text/html' });
                res.end(data);
            } else if (pathname === '/api/config' && method === 'GET') {
                // config.json が無い/項目欠落でも、既定値で補完して返す
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify(loadConfig()));
            } else if (pathname === '/api/config' && method === 'POST') {
                let body = '';
                req.on('data', chunk => { body += chunk.toString(); });
                req.on('end', async () => {
                    try {
                        const newConfig = JSON.parse(body);
                        const currentConfig = loadConfig();
                        const updatedConfig = { ...currentConfig, ...newConfig };
                        await fs.promises.writeFile(configPath, JSON.stringify(updatedConfig, null, 2), 'utf8');
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ message: 'Config saved successfully' }));
                        console.log('Config updated, reprocessing events...');
                        reprocessCallback(); // データ再処理をトリガー
                    } catch (err) {
                        res.writeHead(400, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ error: 'Invalid JSON' }));
                    }
                });
            } else if (pathname === '/api/events' && method === 'GET') {
                const files = await fs.promises.readdir(eventsDir);
                const eventPromises = files.filter(file => {
                    const eventDir = path.join(eventsDir, file);
                    return fs.statSync(eventDir).isDirectory();
                }).map(async eventId => {
                    const eventJsonPath = path.join(eventsDir, eventId, 'Event.json');
                    if (fs.existsSync(eventJsonPath)) {
                        const eventData = JSON.parse(await fs.promises.readFile(eventJsonPath, 'utf8'));
                        return { id: eventId, name: eventData[0].Name };
                    }
                    return null;
                });
                const events = (await Promise.all(eventPromises)).filter(e => e !== null);
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify(events));
            } else if (pathname === '/api/rounds' && method === 'GET') {
                const currentConfig = loadConfig();
                const selectedEventId = currentConfig.selected_event_id;

                if (!selectedEventId || selectedEventId === 'all') {
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify([{ id: 'all', name: 'すべてのラウンド' }]));
                    return;
                }

                const eventDir = path.join(eventsDir, selectedEventId);
                const roundsJsonPath = path.join(eventDir, 'Rounds.json');

                if (!fs.existsSync(roundsJsonPath)) {
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify([{ id: 'all', name: 'すべてのラウンド' }]));
                    return;
                }

                const roundsData = JSON.parse(await fs.promises.readFile(roundsJsonPath, 'utf8'));
                const rounds = roundsData
                    .filter(round => round.Valid === true)
                    .map(round => ({
                        id: round.ID,
                        name: `${round.EventType}Round${round.RoundNumber}`
                    }));

                const allEventTypeRounds = [
                    { id: 'allRace', name: 'すべてのRaceラウンド' },
                    { id: 'allPractice', name: 'すべてのPracticeラウンド' },
                    { id: 'allTimeTrial', name: 'すべてのTimeTrialラウンド' },
                    { id: 'allEndurance', name: 'すべてのEnduranceラウンド' }
                ];
                const responseRounds = [{ id: 'all', name: 'すべてのラウンド' }, ...allEventTypeRounds, ...rounds];
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify(responseRounds));
            } else {
                const baseRoot = path.resolve(baseDir);
                const safeFilePath = path.resolve(path.join(baseDir, pathname));

                // パストラバーサル対策 (baseDir の外は不可。末尾セパレータ付きで前方一致)
                if (safeFilePath !== baseRoot && !safeFilePath.startsWith(baseRoot + path.sep)) {
                    console.warn(`[403] Forbidden access attempt to: ${pathname}`);
                    res.writeHead(403);
                    res.end('Forbidden');
                    return;
                }

                // 静的配信は「公開して良い拡張子」だけに限定する許可リスト方式。
                // これにより credentials.json / config.json (= .json) は配信されない。
                // さらにソース等を含むディレクトリと機密ファイルは明示的に拒否する。
                // (設定UIは /api/* しか使わないので、静的配信を絞っても実害なし)
                const mimeTypes = {
                    '.html': 'text/html',
                    '.js': 'text/javascript',
                    '.css': 'text/css',
                    '.png': 'image/png',
                    '.jpg': 'image/jpeg',
                    '.jpeg': 'image/jpeg',
                    '.gif': 'image/gif',
                    '.svg': 'image/svg+xml',
                    '.ico': 'image/x-icon'
                };
                const ext = path.extname(safeFilePath).toLowerCase();

                const rel = path.relative(baseRoot, safeFilePath).split(path.sep).join('/');
                const firstSeg = rel.split('/')[0];
                const blockedDirs = new Set(['src', 'node_modules', '.git', 'web', 'webdist']);
                const blockedFiles = new Set(['config.json', 'credentials.json', 'package.json', 'package-lock.json']);
                const isBlocked = blockedDirs.has(firstSeg) || blockedFiles.has(rel) || path.basename(rel).startsWith('.');

                if (!mimeTypes[ext] || isBlocked) {
                    console.warn(`[403] Forbidden static path: ${pathname}`);
                    res.writeHead(403);
                    res.end('Forbidden');
                    return;
                }

                if (fs.existsSync(safeFilePath) && fs.statSync(safeFilePath).isFile()) {
                    res.writeHead(200, { 'Content-Type': mimeTypes[ext] });
                    fs.createReadStream(safeFilePath).pipe(res);
                } else {
                    console.log(`[404] Not Found: ${pathname}`);
                    res.writeHead(404);
                    res.end('Not Found');
                }
            }
        } catch (error) {
            console.error(`Error handling request for ${pathname}:`, error);
            res.writeHead(500);
            res.end('Internal Server Error');
        }
    });

    server.listen(config.web_ui_port, () => {
        console.log(`Web UI running at http://localhost:${config.web_ui_port}`);
    });
}

module.exports = { startServer };
