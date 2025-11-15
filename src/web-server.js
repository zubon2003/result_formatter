const http = require('http');
const fs = require('fs');
const path = require('path');
const url = require('url');
const puppeteer = require('puppeteer');
const { config, eventsDir, configPath, loadConfig } = require('./config');

// Web UIのキャッシュ
let webCache = {
    pilotBests: {},
    allPilots: {},
    eventName: '',
    latestHeatName: null,
    nextHeatName: null,
    nextHeatPilots: [],
    lastHeatPilotIds: []
};

function updateCache(data) {
    if (data) {
        webCache = {
            pilotBests: data.pilotBests || {},
            allPilots: data.allPilots || {},
            eventName: data.eventName || '',
            latestHeatName: data.latestHeatName || null,
            nextHeatName: data.nextHeatName || null,
            nextHeatPilots: data.nextHeatPilots || [],
            lastHeatPilotIds: data.lastHeatPilotIds || []
        };
        console.log('Web UI cache has been updated.');
    } else {
        console.warn('Web UI cache update received null data.');
    }
}

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
            } else if (pathname === '/leaderboard') {
                const filePath = path.join(baseDir, 'leaderboard.html');
                const data = await fs.promises.readFile(filePath);
                res.writeHead(200, { 'Content-Type': 'text/html' });
                res.end(data);
            } else if (pathname === '/image.png') {
                const filePath = path.join(baseDir, 'image.png');
                if (fs.existsSync(filePath)) {
                    const imageStream = fs.createReadStream(filePath);
                    res.writeHead(200, { 'Content-Type': 'image/png' });
                    imageStream.pipe(res);
                } else {
                    res.writeHead(404);
                    res.end('Not Found');
                }
            } else if (pathname === '/api/nextheat.png') {
                let browser = null;
                try {
                    browser = await puppeteer.launch();
                    const page = await browser.newPage();

                    // Puppeteer内のコンソール出力をNodeのコンソールに転送
                    page.on('console', msg => {
                        console.log(`[Browser Console]: ${msg.text()}`);
                    });
                    
                    await page.setViewport({ width: 1920, height: 1080 });

                    const nextHeatHtmlPath = path.join(baseDir, 'nextHeat.html');
                    const htmlContent = fs.readFileSync(nextHeatHtmlPath, 'utf-8');
                    
                    // Base URLを正しく設定するために、file://プロトコルではなく、ローカルサーバーのURLを基準にする
                    // page.setContentは外部リソースの解決にベースURLを必要とする
                    await page.goto(`http://localhost:${config.web_ui_port}/internal_nextheat.html`, { waitUntil: 'networkidle0' });

                    // データが描画されるのを待つ
                    await page.waitForFunction(() => {
                        const area = document.getElementById('pilot-display-area');
                        const message = document.getElementById('no-heat-message');
                        return (area && area.innerHTML.trim() !== '') || (message && message.style.display !== 'none');
                    }, { timeout: 10000 }); // タイムアウトを設定

                    const imageBuffer = await page.screenshot();
                    
                    res.writeHead(200, { 'Content-Type': 'image/png' });
                    res.end(imageBuffer);
                } catch(e) {
                    console.error("Error generating nextheat.png:", e);
                    res.writeHead(500);
                    res.end("Error generating image");
                } finally {
                    if (browser) {
                        await browser.close();
                    }
                }
            } else if (pathname === '/internal_nextheat.html') { // Puppeteer専用の内部ルート
                const filePath = path.join(baseDir, 'nextHeat.html');
                const data = await fs.promises.readFile(filePath);
                res.writeHead(200, { 'Content-Type': 'text/html' });
                res.end(data);
            } else if (pathname === '/api/pilot_image') {
                const pilotId = parsedUrl.query.id;
                const photoPathFromQuery = parsedUrl.query.path; // New: Get path from query

                let imagePathToLoad = null;
                let pilotNameForLog = 'Unknown';

                const fpvTracksideDir = path.resolve(eventsDir, '..');

                if (photoPathFromQuery) {
                    // If path is provided, use it directly
                    imagePathToLoad = path.join(fpvTracksideDir, photoPathFromQuery);
                    pilotNameForLog = `(from path: ${photoPathFromQuery})`;
                } else if (pilotId) {
                    // Fallback to pilotId if path is not provided
                    const pilot = webCache.allPilots[pilotId];
                    if (!pilot || !pilot.PhotoPath) {
                        console.log(`Pilot or PhotoPath not found in cache for pilotId: ${pilotId}. PhotoPath: ${pilot ? pilot.PhotoPath : 'N/A'}. Raw Pilot Data: ${JSON.stringify(pilot)}. Serving placeholder.`);
                        const placeholderPath = path.join(baseDir, 'image2.png');
                        if (fs.existsSync(placeholderPath)) {
                            res.writeHead(200, { 'Content-Type': 'image/png' });
                            fs.createReadStream(placeholderPath).pipe(res);
                        } else {
                            res.writeHead(204);
                            res.end();
                        }
                        return;
                    }
                    imagePathToLoad = path.join(fpvTracksideDir, pilot.PhotoPath);
                    pilotNameForLog = pilot.Name || 'Unknown';
                } else {
                    res.writeHead(400, { 'Content-Type': 'text/plain' });
                    res.end('Pilot ID or PhotoPath is required');
                    return;
                }

                

                // Security check (path traversal)
                if (!path.resolve(imagePathToLoad).startsWith(fpvTracksideDir)) {
                    console.error(`Forbidden access attempt to: ${imagePathToLoad}`);
                    res.writeHead(403);
                    res.end('Forbidden');
                    return;
                }

                if (fs.existsSync(imagePathToLoad)) {
                    const imageStream = fs.createReadStream(imagePathToLoad);
                    const ext = path.extname(imagePathToLoad).toLowerCase();
                    let contentType = 'application/octet-stream';
                    if (ext === '.png') contentType = 'image/png';
                    if (ext === '.jpg' || ext === '.jpeg') contentType = 'image/jpeg';
                    
                    res.writeHead(200, { 'Content-Type': contentType });
                    imageStream.pipe(res);
                } else {
                    console.log(`Pilot image not found at: ${imagePathToLoad}. Serving placeholder.`);
                    const placeholderPath = path.join(baseDir, 'image2.png');
                    if (fs.existsSync(placeholderPath)) {
                        res.writeHead(200, { 'Content-Type': 'image/png' });
                        fs.createReadStream(placeholderPath).pipe(res);
                    } else {
                        res.writeHead(204);
                        res.end();
                    }
                }
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

            } else if (pathname === '/api/leaderboard' && method === 'GET') {
                const currentConfig = loadConfig();
                const { leaderboard_round, sorted_by } = currentConfig;
                const { pilotBests, allPilots, eventName, latestHeatName, nextHeatName, nextHeatPilots, lastHeatPilotIds } = webCache;

                let roundName = 'すべてのラウンド';
                const allEventTypeRoundNames = {
                    'allRace': 'すべてのRaceラウンド', 'allPractice': 'すべてのPracticeラウンド',
                    'allTimeTrial': 'すべてのTimeTrialラウンド', 'allEndurance': 'すべてのEnduranceラウンド'
                };
                if (allEventTypeRoundNames[leaderboard_round]) {
                    roundName = allEventTypeRoundNames[leaderboard_round];
                } else if (leaderboard_round !== 'all' && currentConfig.selected_event_id !== 'all') {
                    // ラウンド名を取得するロジック...
                }

                const sortedByDisplayNames = {
                    "bestLap": "CONSECUTIVE 1 LAP (WITHOUT HS)", "consecutive2Lap": "CONSECUTIVE 2 LAPS (WITHOUT HS)",
                    "consecutive3Lap": "CONSECUTIVE 3 LAPS (WITHOUT HS)", "raceTime": "Race Time",
                    "first1LapWithoutHs": "First 1 LAP (WITHOUT HS)", "first2LapsWithoutHs": "First 2 LAPS (WITHOUT HS)",
                    "first3LapsWithoutHs": "First 3 LAPS (WITHOUT HS)", "first1LapWithHs": "First 1 LAP (WITH HS)",
                    "first2LapsWithHs": "First 2 LAPS (WITH HS)", "first3LapsWithHs": "First 3 LAPS (WITH HS)"
                };
                const sortedByDisplayName = sortedByDisplayNames[sorted_by] || sorted_by;

                let ranking = Object.keys(pilotBests)
                    .map(pilotId => {
                        const pilotName = allPilots[pilotId] ? allPilots[pilotId].Name : pilotId;
                        const data = pilotBests[pilotId][sorted_by];
                        if (!data || typeof data.time !== 'number' || !isFinite(data.time) || data.time >= 999) {
                            return null;
                        }
                        return { pilotId, pilotName, time: data.time, heatName: data.heatName };
                    })
                    .filter(item => item !== null)
                    .sort((a, b) => a.time - b.time);

                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({
                    eventName, roundName, sortedByDisplayName, ranking,
                    lastHeatName: latestHeatName,
                    nextHeatName: nextHeatName,
                    nextHeatPilots: nextHeatPilots,
                    lastHeatPilotIds: lastHeatPilotIds
                }));
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

module.exports = { startServer, updateCache };