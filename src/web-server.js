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
                const data = await fs.promises.readFile(configPath, 'utf8');
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(data);
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
                const filePath = path.join(baseDir, pathname);
                const safeFilePath = path.resolve(filePath);

                // Security check to prevent path traversal
                if (!safeFilePath.startsWith(path.resolve(baseDir))) {
                    console.warn(`[403] Forbidden access attempt to: ${pathname}`);
                    res.writeHead(403);
                    res.end('Forbidden');
                    return;
                }

                if (fs.existsSync(safeFilePath) && fs.statSync(safeFilePath).isFile()) {
                    const ext = path.extname(safeFilePath).toLowerCase();
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
                    const contentType = mimeTypes[ext] || 'application/octet-stream';

                    res.writeHead(200, { 'Content-Type': contentType });
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
