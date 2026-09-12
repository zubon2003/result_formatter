const fs = require('fs');
const path = require('path');
const { loadConfig, eventsDir, webSrcDir, webDistDir } = require('./config');
const { sanitizeRaceResults, updateGoogleSheet, updateAllRankingSheets } = require('./google-sheets');
const { exportWebMulti, slugify } = require('./web-export');
const { publishWeb } = require('./publisher');

// 1 イベント分の生データ(Event/Pilots/Rounds/Stages + races)を読む。
// web の保管イベントを必要時だけ読み込むために使う。
function readEventRaw(eventId) {
    const dir = path.join(eventsDir, eventId);
    const eventJson = path.join(dir, 'Event.json');
    const pilotsJson = path.join(dir, 'Pilots.json');
    const roundsJson = path.join(dir, 'Rounds.json');
    if (!fs.existsSync(eventJson) || !fs.existsSync(pilotsJson) || !fs.existsSync(roundsJson)) return null;
    try {
        const eventData = JSON.parse(fs.readFileSync(eventJson, 'utf8'));
        const pilotsData = JSON.parse(fs.readFileSync(pilotsJson, 'utf8'));
        const roundsData = JSON.parse(fs.readFileSync(roundsJson, 'utf8'));
        let stagesData = [];
        const stagesJson = path.join(dir, 'Stages.json');
        if (fs.existsSync(stagesJson)) {
            try { stagesData = JSON.parse(fs.readFileSync(stagesJson, 'utf8')); } catch (e) { /* ignore */ }
        }
        const races = [];
        for (const f of fs.readdirSync(dir)) {
            const rd = path.join(dir, f);
            if (!fs.statSync(rd).isDirectory()) continue;
            const rj = path.join(rd, 'Race.json');
            if (!fs.existsSync(rj)) continue;
            const raceData = JSON.parse(fs.readFileSync(rj, 'utf8'));
            let resultData = null;
            const rs = path.join(rd, 'Result.json');
            if (fs.existsSync(rs)) resultData = JSON.parse(fs.readFileSync(rs, 'utf8'));
            races.push({ id: raceData[0].ID, raceData, resultData });
        }
        return { eventId, eventData, pilotsData, roundsData, stagesData, races };
    } catch (e) {
        console.warn(`readEventRaw failed (${eventId}): ${e.message}`);
        return null;
    }
}

// FPVTrackside の application settings (ShownDecimalPlaces) を読む。
// プロファイル別 <dir>/data/<profile>/ProfileSettings.xml にあるので、
// 最も最近更新された(=アクティブな)プロファイルの値を採用する。既定 2。
function readShownDecimalPlaces() {
    try {
        const dataDir = path.join(path.dirname(eventsDir), 'data');
        if (!fs.existsSync(dataDir)) return 2;
        let best = null, bestM = -1;
        for (const p of fs.readdirSync(dataDir)) {
            const xml = path.join(dataDir, p, 'ProfileSettings.xml');
            if (!fs.existsSync(xml)) continue;
            const m = fs.statSync(xml).mtimeMs;
            if (m > bestM) { bestM = m; best = xml; }
        }
        if (!best) return 2;
        const mt = /<ShownDecimalPlaces>(\d+)<\/ShownDecimalPlaces>/.exec(fs.readFileSync(best, 'utf8'));
        return mt ? parseInt(mt[1], 10) : 2;
    } catch (e) { return 2; }
}

// RaceResult シートの "Race Time (XLap)" ヘッダに使う Lap 数を決める。
// 複数イベント混在(selected_event_id='all')時に「最後に処理したイベント」依存で
// ヘッダが変わる問題を避け、有効レースで最も多く使われている Lap 数を採用する。
// (各レースの Race Time 値自体はレースごとの lapsToDo で計算済み。ここはラベルのみ)
// 同数の場合は小さい方を選び、結果を決定的にする。
function pickHeaderLaps(races, fallback) {
    const counts = new Map();
    for (const r of races) {
        const n = r.lapsToDo;
        if (typeof n !== 'number') continue;
        counts.set(n, (counts.get(n) || 0) + 1);
    }
    let best = fallback, bestCount = -1;
    for (const [n, c] of counts) {
        if (c > bestCount || (c === bestCount && n < best)) { best = n; bestCount = c; }
    }
    return best;
}

// イベント名を安価に取得 (Event.json だけ読む)
function readEventName(eventId) {
    try {
        const e = JSON.parse(fs.readFileSync(path.join(eventsDir, eventId, 'Event.json'), 'utf8'));
        return (e[0] && e[0].Name) || '';
    } catch (e) { return ''; }
}

// メインの処理を関数としてラップ
async function processEvents() {
    try {
        const config = loadConfig(); // 処理開始時に最新のconfigを読み込む
        const selectedEventId = config.selected_event_id || 'all';

        const files = await fs.promises.readdir(eventsDir);

        let targetEventIds = files.filter(file => {
            const eventDir = path.join(eventsDir, file);
            try {
                return fs.statSync(eventDir).isDirectory();
            } catch (e) {
                console.warn(`Could not stat directory ${eventDir}: ${e.message}`);
                return false;
            }
        });

        if (selectedEventId && selectedEventId !== 'all') {
            targetEventIds = targetEventIds.filter(id => id === selectedEventId);
        }

        console.log('Processing event IDs:', targetEventIds);

        let eventName = '';
        const allRaceResults = [];
        let lapsToDo = 4;

        const pilotBests = {};
        const allPilots = {};
        const allValidLapTimes = [];

        const allRaces = []; // 全イベントの全レース情報を格納
        const eventRawById = {}; // web 出力用: イベントごとの生データ (同じ 1 パスで収集)

        for (const eventId of targetEventIds) {
            const eventDir = path.join(eventsDir, eventId);

            const eventJsonPath = path.join(eventDir, 'Event.json');
            const pilotsJsonPath = path.join(eventDir, 'Pilots.json');
            const roundsJsonPath = path.join(eventDir, 'Rounds.json');

            if (!fs.existsSync(eventJsonPath) || !fs.existsSync(pilotsJsonPath) || !fs.existsSync(roundsJsonPath)) {
                continue;
            }

            const eventData = JSON.parse(fs.readFileSync(eventJsonPath, 'utf8'));
            const pilotsData = JSON.parse(fs.readFileSync(pilotsJsonPath, 'utf8'));
            const roundsData = JSON.parse(fs.readFileSync(roundsJsonPath, 'utf8'));

            eventName = eventData[0].Name; // 最後に処理されたイベント名が使われる
            lapsToDo = eventData[0].Laps;

            // Stages.json も同じパスで読む (web のステージ表示用)
            let stagesData = [];
            const stagesJsonPath = path.join(eventDir, 'Stages.json');
            if (fs.existsSync(stagesJsonPath)) {
                try { stagesData = JSON.parse(fs.readFileSync(stagesJsonPath, 'utf8')); }
                catch (e) { console.warn(`Failed to read Stages.json (${eventId}): ${e.message}`); }
            }
            eventRawById[eventId] = { eventId, eventData, pilotsData, roundsData, stagesData, races: [] };

            const raceDirs = fs.readdirSync(eventDir).filter(file => {
                const raceDir = path.join(eventDir, file);
                return fs.statSync(raceDir).isDirectory();
            });

            for (const raceDir of raceDirs) {
                const raceJsonPath = path.join(eventDir, raceDir, 'Race.json');
                const resultJsonPath = path.join(eventDir, raceDir, 'Result.json');

                if (fs.existsSync(raceJsonPath)) {
                    const raceData = JSON.parse(fs.readFileSync(raceJsonPath, 'utf8'));
                    let resultData = null;
                    if (fs.existsSync(resultJsonPath)) {
                        resultData = JSON.parse(fs.readFileSync(resultJsonPath, 'utf8'));
                    }
                    const round = roundsData.find(r => r.ID === raceData[0].Round);
                    allRaces.push({
                        id: raceData[0].ID,
                        roundNumber: round ? round.RoundNumber : 0,
                        eventType: round ? round.EventType : 'Race',
                        raceNumber: raceData[0].RaceNumber,
                        raceData,
                        resultData,
                        pilotsData,
                        eventName: eventData[0].Name,
                        lapsToDo: eventData[0].Laps
                    });
                    // web 出力用にも同じデータを保持 (再読込しない)
                    eventRawById[eventId].races.push({ id: raceData[0].ID, raceData, resultData });
                }
            }
        }

        // --- 表示用 web の生成 ---
        // モデル: 「保管は複数イベント、更新はアクティブ1イベントだけ」。
        //   - published_event_ids が公開(保管)対象
        //   - アクティブ(selected_event_id) は毎回再生成
        //   - それ以外は出力が無ければ一度だけ生成、あれば保管(再生成しない)
        try {
            let publishedIds = Array.isArray(config.published_event_ids) ? config.published_event_ids.slice() : [];

            if (!publishedIds.length) {
                // フォールバック: selected / 最後に開いたイベントの 1 件
                const sel = config.selected_event_id;
                if (sel && sel !== 'all') publishedIds = [sel];
                else {
                    const list = Object.values(eventRawById);
                    list.sort((a, b) => {
                        const la = (a.eventData[0] && a.eventData[0].LastOpened) || '';
                        const lb = (b.eventData[0] && b.eventData[0].LastOpened) || '';
                        return String(lb).localeCompare(String(la));
                    });
                    if (list[0]) publishedIds = [list[0].eventId];
                }
            }

            const activeId = (config.selected_event_id && config.selected_event_id !== 'all')
                ? config.selected_event_id
                : publishedIds[0];

            // アクティブ(ライブ更新)イベントは必ず公開対象に含める
            if (activeId && !publishedIds.includes(activeId)) publishedIds.unshift(activeId);

            const used = new Set();
            const specs = [];
            for (const pid of publishedIds) {
                const name = (eventRawById[pid] && eventRawById[pid].eventData[0] &&
                    eventRawById[pid].eventData[0].Name) || readEventName(pid);

                // 一意なスラッグ (config の並び順で安定)
                let slug = slugify(name);
                if (used.has(slug)) slug = slug + '-' + String(pid).slice(0, 6);
                let n = 2;
                while (used.has(slug)) slug = slugify(name) + '-' + (n++);
                used.add(slug);

                const outExists = fs.existsSync(path.join(webDistDir, slug, 'data', 'bundle.json'));
                const isActive = (pid === activeId);

                if (isActive || !outExists) {
                    // 再生成: 既読(アクティブ)ならそれを、無ければオンデマンドで読む
                    const raw = eventRawById[pid] || readEventRaw(pid);
                    if (raw) specs.push({ slug, name, eventId: pid, ev: raw });
                } else {
                    // 保管: 既存出力をそのまま使う
                    specs.push({ slug, name, eventId: pid });
                }
            }

            if (specs.length) {
                const decimalPlaces = readShownDecimalPlaces();
                exportWebMulti(specs, webSrcDir, webDistDir, decimalPlaces);
                // 生成した webdist を外部ストレージへ公開 (設定が無ければ no-op)
                await publishWeb(webDistDir);
            } else {
                console.warn('web-export: no target event found.');
            }
        } catch (e) {
            console.error('web-export failed:', e);
        }

        // 全レースをラウンドとレース番号でソート
        allRaces.sort((a, b) => a.roundNumber - b.roundNumber || a.raceNumber - b.raceNumber);

        const validRaces = allRaces.filter(race => race.raceData[0].Valid === true);

        // --- ループ1: スプレッドシート用の全データを作成 ---
        validRaces.forEach(race => {
            const { roundNumber, eventType, raceNumber, raceData, resultData, pilotsData, eventName, lapsToDo } = race;
            const displayRoundNumber = roundNumber === 0 ? 'N/A' : roundNumber;
            const raceName = eventType + ' ' + displayRoundNumber + '-' + raceNumber;

            let raceSerialTimestamp = '';
            const firstLap = [...raceData[0].Laps].sort((a, b) => a.LapNumber - b.LapNumber)[0];
            if (firstLap && firstLap.StartTime) {
                const dateObj = new Date(firstLap.StartTime);
                const year = dateObj.getFullYear();
                const month = dateObj.getMonth();
                const day = dateObj.getDate();
                const hours = dateObj.getHours();
                const minutes = dateObj.getMinutes();
                const seconds = dateObj.getSeconds();
                const utcDate = new Date(Date.UTC(year, month, day, hours, minutes, seconds));
                const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                raceSerialTimestamp = (utcDate.getTime() - excelEpoch.getTime()) / (24 * 60 * 60 * 1000);
            }

            const pilotIdsInRace = [...new Set(raceData[0].Detections.map(d => d.Pilot))];

            pilotIdsInRace.forEach(pilotId => {
                const pilot = pilotsData.find(p => p.ID === pilotId);
                if (pilot) {
                    const result = resultData ? resultData.find(r => r.Pilot === pilotId) : null;
                    const position = result ? result.Position : '';

                    const pilotLaps = raceData[0].Laps.filter(lap => {
                        const detection = raceData[0].Detections.find(d => d.ID === lap.Detection);
                        return detection && detection.Pilot === pilotId && detection.Valid === true;
                    }).sort((a, b) => a.LapNumber - b.LapNumber);

                    const lapTimes = Array(31).fill('');
                    let totalFinishTime = '';
                    let lapCount = pilotLaps.filter(lap => lap.LapNumber > 0).length;
                    
                    let bestLap = 999;
                    let consecutive2Lap = 999;
                    let consecutive3Lap = 999;
                    let raceTimeXLap = 9999;

                    const actualLapTimes = pilotLaps.filter(lap => lap.LapNumber >= 1).map(lap => lap.LengthSeconds);

                    if (pilotLaps.length > 0) {
                        lapTimes[0] = pilotLaps[0].LengthSeconds;
                        totalFinishTime = pilotLaps.reduce((sum, lap) => sum + lap.LengthSeconds, 0);

                        if (actualLapTimes.length > 0) {
                            const minLap = Math.min(...actualLapTimes);
                            if (isFinite(minLap)) bestLap = minLap;
                        }
                        if (actualLapTimes.length >= 2) {
                            let min2Lap = Infinity;
                            for (let i = 0; i < actualLapTimes.length - 1; i++) {
                                min2Lap = Math.min(min2Lap, actualLapTimes[i] + actualLapTimes[i+1]);
                            }
                            if (isFinite(min2Lap)) consecutive2Lap = min2Lap;
                        }
                        if (actualLapTimes.length >= 3) {
                            let min3Lap = Infinity;
                            for (let i = 0; i < actualLapTimes.length - 2; i++) {
                                min3Lap = Math.min(min3Lap, actualLapTimes[i] + actualLapTimes[i+1] + actualLapTimes[i+2]);
                            }
                            if (isFinite(min3Lap)) consecutive3Lap = min3Lap;
                        }

                        const hsLap = pilotLaps.find(lap => lap.LapNumber === 0);
                        if (hsLap) {
                            if (pilotLaps.length >= lapsToDo + 1) {
                                raceTimeXLap = pilotLaps.slice(0, lapsToDo + 1).reduce((sum, lap) => sum + lap.LengthSeconds, 0);
                            }
                        } else {
                            if (pilotLaps.length >= lapsToDo) {
                                raceTimeXLap = pilotLaps.slice(0, lapsToDo).reduce((sum, lap) => sum + lap.LengthSeconds, 0);
                            }
                        }

                        pilotLaps.forEach(lap => {
                            if (lap.LapNumber >= 1 && lap.LapNumber <= 30) {
                                lapTimes[lap.LapNumber] = lap.LengthSeconds;
                            }
                        });
                    }

                    const newRow = [
                        eventName, raceName, raceSerialTimestamp, raceSerialTimestamp, pilot.Name, position,
                        lapCount, totalFinishTime, raceTimeXLap, bestLap, consecutive2Lap, consecutive3Lap,
                        ...lapTimes
                    ];
                    allRaceResults.push(newRow);
                }
            });
        });


        // --- ループ2: ランキングシート用にラウンドで絞り込んだデータを作成 ---
        const currentConfig = loadConfig(); // 最新のconfigを取得
        const leaderboardRound = currentConfig.leaderboard_round;

        const filteredRaces = validRaces.filter(race => {
            if (leaderboardRound === 'all') return true;
            if (leaderboardRound === 'allRace' && race.eventType === 'Race') return true;
            if (leaderboardRound === 'allPractice' && race.eventType === 'Practice') return true;
            if (leaderboardRound === 'allTimeTrial' && race.eventType === 'TimeTrial') return true;
            if (leaderboardRound === 'allEndurance' && race.eventType === 'Endurance') return true;
            return race.raceData[0].Round === leaderboardRound;
        });

        filteredRaces.forEach(race => {
            const { roundNumber, eventType, raceNumber, raceData, resultData, pilotsData, lapsToDo } = race;
            const displayRoundNumber = roundNumber === 0 ? 'N/A' : roundNumber;
            const raceName = eventType + ' ' + displayRoundNumber + '-' + raceNumber;

            let raceSerialTimestamp = '';
            const firstLap = [...raceData[0].Laps].sort((a, b) => a.LapNumber - b.LapNumber)[0];
            if (firstLap && firstLap.StartTime) {
                const dateObj = new Date(firstLap.StartTime);
                const year = dateObj.getFullYear();
                const month = dateObj.getMonth();
                const day = dateObj.getDate();
                const hours = dateObj.getHours();
                const minutes = dateObj.getMinutes();
                const seconds = dateObj.getSeconds();
                const utcDate = new Date(Date.UTC(year, month, day, hours, minutes, seconds));
                const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                raceSerialTimestamp = (utcDate.getTime() - excelEpoch.getTime()) / (24 * 60 * 60 * 1000);
            }

            const pilotIdsInRace = [...new Set(raceData[0].Detections.map(d => d.Pilot))];

            pilotIdsInRace.forEach(pilotId => {
                const pilot = pilotsData.find(p => p.ID === pilotId);
                if (pilot) {
                    if (!allPilots[pilot.ID]) {
                        allPilots[pilot.ID] = pilot; // Store the full pilot object
                    }
                    if (!pilotBests[pilot.ID]) {
                        pilotBests[pilot.ID] = {
                            raceTime: { time: 9999, timestamp: null, heatName: null },
                            bestLap: { time: 999, timestamp: null, heatName: null },
                            consecutive2Lap: { time: 999, timestamp: null, heatName: null },
                            consecutive3Lap: { time: 999, timestamp: null, heatName: null },
                            first1LapWithoutHs: { time: 999, timestamp: null, heatName: null },
                            first2LapsWithoutHs: { time: 999, timestamp: null, heatName: null },
                            first3LapsWithoutHs: { time: 999, timestamp: null, heatName: null },
                            first1LapWithHs: { time: 999, timestamp: null, heatName: null },
                            first2LapsWithHs: { time: 999, timestamp: null, heatName: null },
                            first3LapsWithHs: { time: 999, timestamp: null, heatName: null }
                        };
                    }

                    const updateBestTime = (category, time, timestamp, heatName) => {
                        if (typeof time === 'number' && isFinite(time)) {
                            const currentBest = pilotBests[pilot.ID][category];
                            if (time < currentBest.time) {
                                currentBest.time = time;
                                currentBest.timestamp = timestamp;
                                currentBest.heatName = heatName;
                            }
                        }
                    };

                    const pilotLaps = raceData[0].Laps.filter(lap => {
                        const detection = raceData[0].Detections.find(d => d.ID === lap.Detection);
                        return detection && detection.Pilot === pilotId && detection.Valid === true;
                    }).sort((a, b) => a.LapNumber - b.LapNumber);

                    let bestLap = 999;
                    let consecutive2Lap = 999;
                    let consecutive3Lap = 999;
                    let raceTimeXLap = 9999;

                    const actualLapTimes = pilotLaps.filter(lap => lap.LapNumber >= 1).map(lap => lap.LengthSeconds);
                    const allLapTimesIncludingHs = pilotLaps.filter(lap => lap.LapNumber >= 0).map(lap => lap.LengthSeconds);

                    if (pilotLaps.length > 0) {
                        if (actualLapTimes.length > 0) {
                            const minLap = Math.min(...actualLapTimes);
                            if (isFinite(minLap)) bestLap = minLap;
                        }
                        if (actualLapTimes.length >= 2) {
                            let min2Lap = Infinity;
                            for (let i = 0; i < actualLapTimes.length - 1; i++) {
                                min2Lap = Math.min(min2Lap, actualLapTimes[i] + actualLapTimes[i+1]);
                            }
                            if (isFinite(min2Lap)) consecutive2Lap = min2Lap;
                        }
                        if (actualLapTimes.length >= 3) {
                            let min3Lap = Infinity;
                            for (let i = 0; i < actualLapTimes.length - 2; i++) {
                                min3Lap = Math.min(min3Lap, actualLapTimes[i] + actualLapTimes[i+1] + actualLapTimes[i+2]);
                            }
                            if (isFinite(min3Lap)) consecutive3Lap = min3Lap;
                        }

                        const hsLap = pilotLaps.find(lap => lap.LapNumber === 0);
                        if (hsLap) {
                            if (pilotLaps.length >= lapsToDo + 1) {
                                raceTimeXLap = pilotLaps.slice(0, lapsToDo + 1).reduce((sum, lap) => sum + lap.LengthSeconds, 0);
                            }
                        } else {
                            if (pilotLaps.length >= lapsToDo) {
                                raceTimeXLap = pilotLaps.slice(0, lapsToDo).reduce((sum, lap) => sum + lap.LengthSeconds, 0);
                            }
                        }
                    }

                    if (actualLapTimes.length >= 1) updateBestTime('first1LapWithoutHs', actualLapTimes[0], raceSerialTimestamp, raceName);
                    if (actualLapTimes.length >= 2) updateBestTime('first2LapsWithoutHs', actualLapTimes[0] + actualLapTimes[1], raceSerialTimestamp, raceName);
                    if (actualLapTimes.length >= 3) updateBestTime('first3LapsWithoutHs', actualLapTimes[0] + actualLapTimes[1] + actualLapTimes[2], raceSerialTimestamp, raceName);
                    if (allLapTimesIncludingHs.length >= 2) {
                        updateBestTime('first1LapWithHs', allLapTimesIncludingHs[0] + allLapTimesIncludingHs[1], raceSerialTimestamp, raceName);
                    }
                    if (allLapTimesIncludingHs.length >= 3) {
                        updateBestTime('first2LapsWithHs', allLapTimesIncludingHs[0] + allLapTimesIncludingHs[1] + allLapTimesIncludingHs[2], raceSerialTimestamp, raceName);
                    }
                    if (allLapTimesIncludingHs.length >= 4) {
                        updateBestTime('first3LapsWithHs', allLapTimesIncludingHs[0] + allLapTimesIncludingHs[1] + allLapTimesIncludingHs[2] + allLapTimesIncludingHs[3], raceSerialTimestamp, raceName);
                    }

                    updateBestTime('raceTime', raceTimeXLap, raceSerialTimestamp, raceName);
                    updateBestTime('bestLap', bestLap, raceSerialTimestamp, raceName);
                    updateBestTime('consecutive2Lap', consecutive2Lap, raceSerialTimestamp, raceName);
                    updateBestTime('consecutive3Lap', consecutive3Lap, raceSerialTimestamp, raceName);

                    pilotLaps.forEach(lap => {
                        if (lap.LapNumber >= 1) {
                            allValidLapTimes.push({
                                time: lap.LengthSeconds,
                                pilotName: pilot.Name,
                                heatName: raceName
                            });
                        }
                    });
                }
            });
        });

        // --- Google Sheet の更新 ---
        await updateAllRankingSheets(pilotBests, allPilots, allValidLapTimes);
        console.log('Ranking sheets have been updated.');
        
        allRaceResults.sort((a, b) => {
            const timeA = a[2];
            const timeB = b[2];
            if (typeof timeA === 'number' && typeof timeB === 'number') {
                return timeA - timeB;
            } else if (typeof timeA !== 'number') {
                return 1;
            } else {
                return -1;
            }
        });
        const sanitizedData = sanitizeRaceResults(allRaceResults);
        const headerLaps = pickHeaderLaps(validRaces, lapsToDo);
        await updateGoogleSheet(sanitizedData, headerLaps);
        console.log('RaceResult sheet has been updated.');

    } catch (err) {
        console.error('Error processing events:', err);
    }
}

module.exports = { processEvents };

