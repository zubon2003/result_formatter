const fs = require('fs');
const path = require('path');
const { google } = require('googleapis');

// const eventsDir = path.join(__dirname, '..', 'events'); // config.jsonから読み込むように変更
const config = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8'));
const eventsDir = path.join(config.fpvtrackside_dir_path, 'events').replace(/\\/g, '/'); // fpvtrackside_dir_pathからeventsDirを生成


// Google Sheets APIの認証
const auth = new google.auth.GoogleAuth({
    keyFile: path.join(__dirname, 'credentials.json'), // credentials.json を使用
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

let debounceTimer;
const DEBOUNCE_DELAY = 5000; // 5秒のデバウンス遅延

function sanitizeRaceResults(raceResults) {
    // Return a new array with sanitized data
    return raceResults.map((row, rowIndex) => {
        return row.map((cell, colIndex) => {
            // Check for invalid numeric or general values
            const isInvalid = cell === null || cell === undefined || (typeof cell === 'number' && !isFinite(cell));
            if (isInvalid) {
                console.warn(`[Data Sanitization] Invalid data found in RaceResult at row ${rowIndex + 1}, column ${colIndex + 1}. Value: ${cell}. Replacing with blank.`);
                return ''; // Replace invalid data with an empty string
            }
            return cell; // Keep valid data
        });
    });
}

// メインの処理を関数としてラップ
async function processEvents() {
    try {
        const config = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8'));
        const selectedEventId = config.selected_event_id || 'all'; // ★ configから選択されたイベントIDを取得

        const files = await fs.promises.readdir(eventsDir);

        let targetEventIds = files.filter(file => {
            const eventDir = path.join(eventsDir, file);
            return fs.statSync(eventDir).isDirectory();
        });

        if (selectedEventId && selectedEventId !== 'all') { // ★ 'all'オプションを追加する可能性を考慮
            targetEventIds = targetEventIds.filter(id => id === selectedEventId);
        }

        console.log('Processing event IDs:', targetEventIds);

        let eventName = ''; // ここで初期化
        const allRaceResults = []; // スプレッドシート用の絞り込みなしの結果
        let lapsToDo = 4; // Default laps

        const pilotBests = {};
        const allPilots = {};
        const allValidLapTimes = [];

        for (const eventId of targetEventIds) { // ★ targetEventIds をループ
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

            eventName = eventData[0].Name; // eventNameを更新
            lapsToDo = eventData[0].Laps; // Get laps from event data

            const raceDirs = fs.readdirSync(eventDir).filter(file => {
                const raceDir = path.join(eventDir, file);
                return fs.statSync(raceDir).isDirectory();
            });

            const races = [];
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
                    races.push({
                        roundNumber: round ? round.RoundNumber : 0,
                        eventType: round ? round.EventType : 'Race', // EventTypeを追加
                        raceNumber: raceData[0].RaceNumber,
                        raceData,
                        resultData
                    });
                }
            }

            races.sort((a, b) => a.roundNumber - b.roundNumber || a.raceNumber - b.raceNumber);

            const validRaces = races.filter(race => race.raceData[0].Valid === true);

            // --- ループ1: スプレッドシート用の全データを作成 ---
            validRaces.forEach(race => {
                const { roundNumber, eventType, raceNumber, raceData, resultData } = race;
                const displayRoundNumber = roundNumber === 0 ? 'N/A' : roundNumber;
                const raceName = eventType + ' ' + displayRoundNumber + '-' + raceNumber;

                let raceSerialTimestamp = '';
                const firstLap = raceData[0].Laps.sort((a, b) => a.LapNumber - b.LapNumber)[0];
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


            // --- ループ2: Web表示用の絞り込んだデータを作成 ---
            const filteredRaces = validRaces.filter(race => {
                const config = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8'));
                const leaderboardRound = config.leaderboard_round;
                if (leaderboardRound === 'all') {
                    return true; // すべてのラウンドを対象とする
                }
                // 新しいオプションの処理
                if (leaderboardRound === 'allRace' && race.eventType === 'Race') {
                    return true;
                }
                if (leaderboardRound === 'allPractice' && race.eventType === 'Practice') {
                    return true;
                }
                if (leaderboardRound === 'allTimeTrial' && race.eventType === 'TimeTrial') {
                    return true;
                }
                if (leaderboardRound === 'allEndurance' && race.eventType === 'Endurance') {
                    return true;
                }
                return race.raceData[0].Round === leaderboardRound;
            });

            filteredRaces.forEach(race => {
                const { roundNumber, eventType, raceNumber, raceData, resultData } = race;
                const displayRoundNumber = roundNumber === 0 ? 'N/A' : roundNumber;
                const raceName = eventType + ' ' + displayRoundNumber + '-' + raceNumber;

                let raceSerialTimestamp = '';
                const firstLap = raceData[0].Laps.sort((a, b) => a.LapNumber - b.LapNumber)[0];
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
                            allPilots[pilot.ID] = pilot.Name;
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
                        // First X LAPS (WITH HS) の計算 (B案のロジック)
                        // first1LapWithHs = HS + LAP1
                        if (allLapTimesIncludingHs.length >= 2) {
                            updateBestTime('first1LapWithHs', allLapTimesIncludingHs[0] + allLapTimesIncludingHs[1], raceSerialTimestamp, raceName);
                        }
                        // first2LapsWithHs = HS + LAP1 + LAP2
                        if (allLapTimesIncludingHs.length >= 3) {
                            updateBestTime('first2LapsWithHs', allLapTimesIncludingHs[0] + allLapTimesIncludingHs[1] + allLapTimesIncludingHs[2], raceSerialTimestamp, raceName);
                        }
                        // first3LapsWithHs = HS + LAP1 + LAP2 + LAP3
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

            // 最後に飛行したヒートとパイロットを特定
            let latestRaceTimestamp = 0;
            let latestHeatName = null;
            let pilotsInLatestHeat = new Set();

            // filteredRaces を逆順にソートして、最新のレースから処理する
            const sortedFilteredRaces = [...filteredRaces].sort((a, b) => {
                const timeA = new Date(a.raceData[0].Laps.sort((x, y) => x.LapNumber - y.LapNumber)[0]?.StartTime || 0).getTime();
                const timeB = new Date(b.raceData[0].Laps.sort((x, y) => x.LapNumber - y.LapNumber)[0]?.StartTime || 0).getTime();
                return timeB - timeA; // 降順
            });

            if (sortedFilteredRaces.length > 0) {
                const latestRace = sortedFilteredRaces[0];
                const firstLapOfLatestRace = latestRace.raceData[0].Laps.sort((a, b) => a.LapNumber - b.LapNumber)[0];
                if (firstLapOfLatestRace && firstLapOfLatestRace.StartTime) {
                    latestRaceTimestamp = new Date(firstLapOfLatestRace.StartTime).getTime();
                    latestHeatName = latestRace.eventType + ' ' + (latestRace.roundNumber === 0 ? 'N/A' : latestRace.roundNumber) + '-' + latestRace.raceNumber;
                    pilotsInLatestHeat = new Set(latestRace.raceData[0].Detections.map(d => d.Pilot));
                }
            }

            // 1. Webキャッシュを先に更新
            cachedPilotBests = pilotBests;
            cachedAllPilots = allPilots;
            cachedEventName = eventName;
            cachedLastHeatName = latestHeatName; // 新しいキャッシュ変数
            cachedLastHeatPilotIds = Array.from(pilotsInLatestHeat); // 新しいキャッシュ変数
            console.log('Web UI cache has been updated.');
        }

        // 1. Webキャッシュを先に更新
        cachedPilotBests = pilotBests;
        cachedAllPilots = allPilots;
        cachedEventName = eventName;
        console.log('Web UI cache has been updated.');

        // 2. ランキングシートを更新
        await updateAllRankingSheets(pilotBests, allPilots, allValidLapTimes);
        console.log('Ranking sheets have been updated.');
        
        // 3. RaceResultシートを更新 (時間のかかる処理)
        // スタート時刻でソート (昇順)
        allRaceResults.sort((a, b) => {
            const timeA = a[2]; // スタート時刻のシリアル値
            const timeB = b[2];

            if (typeof timeA === 'number' && typeof timeB === 'number') {
                return timeA - timeB;
            } else if (typeof timeA !== 'number') {
                return 1; // timeAが数値でない場合、bを先に
            } else {
                return -1; // timeBが数値でない場合、aを先に
            }
        });
        const sanitizedRaceResults = sanitizeRaceResults(allRaceResults);
        await updateGoogleSheet(sanitizedRaceResults, lapsToDo);
        console.log('RaceResult sheet has been updated.');

    } catch (err) {
        console.error('Error processing events:', err);
    }
}

async function updateGoogleSheet(raceResults, lapsToDo) {
    try {
        const spreadsheetId = config.google_spreadsheet_id;
        const sheetName = 'RaceResult'; // シート名をハードコード

        // --- シートの存在確認とID取得 ---
        const sheetProperties = await sheets.spreadsheets.get({ spreadsheetId });
        let sheet = sheetProperties.data.sheets.find(s => s.properties.title === sheetName);

        const requests = [];

        // --- シートが存在しない場合は作成 ---
        if (!sheet) {
            console.log(`Sheet "${sheetName}" not found, creating it...`);
            const addSheetRequest = {
                addSheet: {
                    properties: {
                        title: sheetName,
                        gridProperties: {
                            rowCount: 1000,
                            columnCount: 100
                        }
                    },
                },
            };
            const response = await sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: { requests: [addSheetRequest] },
            });
            const newSheetProp = response.data.replies[0].addSheet.properties;
            sheet = { properties: newSheetProp };
            console.log(`Sheet "${sheetName}" created with ID: ${sheet.properties.sheetId}.`);
        } else {
            // --- 既存のゼブラストライプを削除 ---
            const existingBands = sheet.bandedRanges || [];
            if (existingBands.length > 0) {
                const deleteRequests = existingBands.map(band => ({
                    deleteBanding: {
                        bandedRangeId: band.bandedRangeId
                    }
                }));
                try {
                    await sheets.spreadsheets.batchUpdate({
                        spreadsheetId,
                        resource: { requests: deleteRequests },
                    });
                } catch (err) {
                    console.log(`Could not delete old banding for RaceResult, probably already gone. Error: ${err.message}`);
                }
            }
        }

        if (!sheet) {
            console.error(`Failed to find or create sheet "${sheetName}".`);
            return;
        }
        const sheetId = sheet.properties.sheetId;
        const currentColumnCount = sheet.properties.gridProperties ? sheet.properties.gridProperties.columnCount : 0; // Get current column count

        if (currentColumnCount < 100) { // If current columns are less than desired
            requests.push({
                updateSheetProperties: {
                    properties: {
                        sheetId: sheetId,
                        gridProperties: {
                            columnCount: 100 // Resize to 100
                        }
                    },
                    fields: 'gridProperties.columnCount'
                }
            });
        }

        // --- シートの並び順を設定 ---
        requests.push({
            updateSheetProperties: {
                properties: {
                    sheetId: sheetId,
                    index: 0 // RaceResultを0番目に
                },
                fields: 'index'
            }
        });

        // --- RaceResultシートのクリアと更新 ---
        requests.push({
            updateCells: {
                range: { sheetId: sheetId },
                fields: "userEnteredValue,userEnteredFormat"
            }
        });

        

        // --- タイトル行のセル結合 ---
        // requests.push({
        //     mergeCells: {
        //         range: {
        //             sheetId: sheetId,
        //             startRowIndex: 0,
        //             endRowIndex: 1,
        //             startColumnIndex: 0,
        //             endColumnIndex: 43 // 全ての列を結合
        //         },
        //         mergeType: 'MERGE_ALL'
        //     }
        // });

        const raceResultHeaders = [
            'Event名', 'HEAT', '日付', 'スタート時刻', 'Pilot', 'Position', 'Lap数', 'Finish time', `Race Time (${lapsToDo}Lap)`, 'Best LAP', '連続2周', '連続3周', 'HS(LAP0)',
            'LAP1', 'LAP2', 'LAP3', 'LAP4', 'LAP5', 'LAP6', 'LAP7', 'LAP8', 'LAP9', 'LAP10',
            'LAP11', 'LAP12', 'LAP13', 'LAP14', 'LAP15', 'LAP16', 'LAP17', 'LAP18', 'LAP19', 'LAP20',
            'LAP21', 'LAP22', 'LAP23', 'LAP24', 'LAP25', 'LAP26', 'LAP27', 'LAP28', 'LAP29', 'LAP30'
        ];

        

        requests.push({
            updateCells: {
                rows: [{
                    values: raceResultHeaders.map(header => ({
                        userEnteredValue: { stringValue: header },
                        userEnteredFormat: {
                            backgroundColor: { red: 0.4, green: 0.4, blue: 0.4 },
                            textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
                            horizontalAlignment: 'RIGHT'
                        }
                    }))
                }],
                range: {
                    sheetId: sheetId,
                    startRowIndex: 0,
                    startColumnIndex: 0,
                    endColumnIndex: 43 // Explicitly set to 43
                },
                fields: "userEnteredValue,userEnteredFormat"
            }
        });

        if (raceResults.length > 0) {
            requests.push({
                addBanding: {
                    bandedRange: {
                        range: {
                            sheetId: sheetId,
                            startRowIndex: 1,
                            endRowIndex: 1 + raceResults.length,
                            startColumnIndex: 0,
                            endColumnIndex: 43
                        },
                        rowProperties: {
                            firstBandColor: { red: 0.9, green: 0.9, blue: 0.9 }, // Even rows
                            secondBandColor: { red: 1, green: 1, blue: 1 }      // Odd rows
                        }
                    }
                }
            });
        }

        if (raceResults.length > 0) {
            
            requests.push({
                updateCells: {
                    rows: raceResults.map(row => ({
                        values: row.map((cell, colIndex) => {
                            const cellData = { userEnteredValue: {} };
                            cellData.userEnteredFormat = { horizontalAlignment: 'RIGHT' };

                            // 列インデックスに基づいて書式を設定
                            if (colIndex === 2) { // C列: 日付
                                if (typeof cell === 'number' && isFinite(cell) && cell !== '') {
                                    cellData.userEnteredValue.numberValue = cell;
                                    cellData.userEnteredFormat.numberFormat = { type: 'DATE', pattern: 'yyyy-mm-dd' };
                                } else {
                                    cellData.userEnteredValue.stringValue = String(cell);
                                }
                            } else if (colIndex === 3) { // D列: 時刻
                                if (typeof cell === 'number' && isFinite(cell) && cell !== '') {
                                    cellData.userEnteredValue.numberValue = cell;
                                    cellData.userEnteredFormat.numberFormat = { type: 'TIME', pattern: 'hh:mm:ss' };
                                } else {
                                    cellData.userEnteredValue.stringValue = String(cell);
                                }
                            } else if (colIndex >= 7 && colIndex <= 42) { // H列からAQ列: 小数点第3位
                                if (typeof cell === 'number' && isFinite(cell) && cell !== '') {
                                    cellData.userEnteredValue.numberValue = cell;
                                    cellData.userEnteredFormat.numberFormat = { type: 'NUMBER', pattern: '0.000' };
                                } else {
                                    cellData.userEnteredValue.stringValue = String(cell);
                                }
                            }
                            // その他の列 (A, B, E列はここに該当し、文字列として扱われる)
                            else if (typeof cell === 'number' && isFinite(cell) && cell !== '') {
                                cellData.userEnteredValue.numberValue = cell;
                            } else {
                                cellData.userEnteredValue.stringValue = String(cell);
                            }
                            return cellData;
                        })
                    })),
                    start: { sheetId: sheetId, rowIndex: 1, columnIndex: 0 }, // ★ rowIndex を 1 に変更
                    fields: "userEnteredValue,userEnteredFormat"
                }
            });
        }

        const columnWidths = [
            { index: 0, size: 150 }, // A
            { index: 1, size: 130 },  // B
            { index: 2, size: 100 }, // C
            { index: 3, size: 100 }, // D
            { index: 4, size: 150 }, // E
            { index: 5, size: 60 },  // F
            { index: 6, size: 60 },  // G
            { index: 7, size: 80 },  // H
            { index: 8, size: 80 },  // I
            { index: 9, size: 80 },  // J
            { index: 10, size: 80 }, // K
            { index: 11, size: 80 }, // L
            { index: 12, size: 80 }, // M
        ];

        columnWidths.forEach(col => {
            requests.push({
                updateDimensionProperties: {
                    range: {
                        sheetId: sheetId,
                        dimension: 'COLUMNS',
                        startIndex: col.index,
                        endIndex: col.index + 1
                    },
                    properties: {
                        pixelSize: col.size
                    },
                    fields: 'pixelSize'
                }
            });
        });

        requests.push({
            updateDimensionProperties: {
                range: {
                    sheetId: sheetId,
                    dimension: 'COLUMNS',
                    startIndex: 13,
                    endIndex: 43
                },
                properties: {
                    pixelSize: 60
                },
                fields: 'pixelSize'
            }
        });

        // --- batchUpdateの実行 ---
        if (requests.length > 0) {
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: { requests },
            });
        }

        console.log(`Successfully updated Google Sheet "${sheetName}".`);

    } catch (err) {
        console.error('Error updating Google Sheet:', err);
    }
}

async function updateSingleRankingSheet(categoryKey, sheetTitle, sourceData, allPilots, index) {
    try {
        const spreadsheetId = config.google_spreadsheet_id;

        // --- スプレッドシート全体の情報を取得 ---
        const spreadsheet = await sheets.spreadsheets.get({ spreadsheetId });
        let sheet = spreadsheet.data.sheets.find(s => s.properties.title === sheetTitle);

        const requests = [];

        // --- シートが存在しない場合は作成 ---
        if (!sheet) {
            console.log(`Sheet "${sheetTitle}" not found, creating it...`);
            const addSheetRequest = {
                addSheet: {
                    properties: { title: sheetTitle, index: index },
                },
            };
            const response = await sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: { requests: [addSheetRequest] },
            });
            const newSheetProp = response.data.replies[0].addSheet.properties;
            sheet = { properties: newSheetProp };
            console.log(`Sheet "${sheetTitle}" created with ID: ${sheet.properties.sheetId}.`);
        } else {
            // --- 既存のゼブラストライプを削除 ---
            const existingBands = sheet.bandedRanges || [];
            if (existingBands.length > 0) {
                const deleteRequests = existingBands.map(band => ({
                    deleteBanding: {
                        bandedRangeId: band.bandedRangeId
                    }
                }));
                try {
                    await sheets.spreadsheets.batchUpdate({
                        spreadsheetId,
                        resource: { requests: deleteRequests },
                    });
                } catch (err) {
                    console.log(`Could not delete old banding, probably already gone. Error: ${err.message}`);
                }
            }
        }

        if (!sheet) {
            console.error(`Failed to find or create sheet "${sheetTitle}".`);
            return;
        }
        const sheetId = sheet.properties.sheetId;

        // --- ランキングデータの作成とソート ---
        let rankingData;

        if (categoryKey === 'minLap') {
            // Minimum Lap Ranking の処理
            sourceData.sort((a, b) => a.time - b.time);
            rankingData = sourceData.slice(0, 100).map(item => ({
                pilotName: item.pilotName,
                data: {
                    time: item.time,
                    heatName: item.heatName
                }
            }));
        } else {
            // 既存のランキング処理
            rankingData = Object.keys(sourceData)
                .map(pilotId => ({
                    pilotId: pilotId,
                    pilotName: allPilots[pilotId] || pilotId,
                    data: sourceData[pilotId][categoryKey]
                }));

            rankingData.sort((a, b) => {
                const timeA = a.data.time;
                const timeB = b.data.time;
                const timestampA = a.data.timestamp;
                const timestampB = b.data.timestamp;

                if (timeA !== timeB) return timeA - timeB;
                if (timestampA !== timestampB) return timestampA - timestampB;
                return 0;
            });
        }


        // --- シートの並び順を設定 ---
        requests.push({
            updateSheetProperties: {
                properties: {
                    sheetId: sheetId,
                    index: index
                },
                fields: 'index'
            }
        });

        // --- シートのクリア ---
        requests.push({
            updateCells: {
                range: { sheetId: sheetId },
                fields: "userEnteredValue,userEnteredFormat"
            }
        });

        // --- タイトル行の作成 ---
        requests.push({
            updateCells: {
                rows: [{
                    values: [{
                        userEnteredValue: { stringValue: sheetTitle },
                        userEnteredFormat: {
                            textFormat: { fontSize: 14, bold: true },
                            horizontalAlignment: 'CENTER'
                        }
                    }]
                }],
                start: { sheetId: sheetId, rowIndex: 0, columnIndex: 0 },
                fields: "userEnteredValue,userEnteredFormat"
            }
        });

        // --- タイトル行のセル結合 ---
        requests.push({
            mergeCells: {
                range: {
                    sheetId: sheetId,
                    startRowIndex: 0,
                    endRowIndex: 1,
                    startColumnIndex: 0,
                    endColumnIndex: 4 // Rank, Pilot, Time, HEAT の4列分
                },
                mergeType: 'MERGE_ALL'
            }
        });

        // --- ヘッダー行の作成 ---
        const headers = ['Rank', 'Pilot', 'Time', 'HEAT'];
        requests.push({
            updateCells: {
                rows: [{
                    values: headers.map(header => ({
                        userEnteredValue: { stringValue: header },
                        userEnteredFormat: {
                            backgroundColor: { red: 0.4, green: 0.4, blue: 0.4 },
                            textFormat: { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 } },
                            horizontalAlignment: 'RIGHT'
                        }
                    }))
                }],
                start: { sheetId: sheetId, rowIndex: 2, columnIndex: 0 },
                fields: "userEnteredValue,userEnteredFormat"
            }
        });

        // --- ゼブラストライプを追加 ---
        if (rankingData.length > 0) {
            requests.push({
                addBanding: {
                    bandedRange: {
                        range: {
                            sheetId: sheetId,
                            startRowIndex: 3,
                            endRowIndex: 3 + rankingData.length,
                            startColumnIndex: 0,
                            endColumnIndex: 4
                        },
                        rowProperties: {
                            firstBandColor: { red: 0.9, green: 0.9, blue: 0.9 },
                            secondBandColor: { red: 1, green: 1, blue: 1 }
                        }
                    }
                }
            });
        }

        // --- データ行の作成 ---
        if (rankingData.length > 0) {
            const rows = rankingData.map((item, index) => ({
                values: [
                    { userEnteredValue: { numberValue: index + 1 }, userEnteredFormat: { horizontalAlignment: 'RIGHT' } }, // Rank
                    { userEnteredValue: { stringValue: item.pilotName }, userEnteredFormat: { horizontalAlignment: 'RIGHT' } }, // Pilot
                    {
                        userEnteredValue: { numberValue: item.data.time },
                        userEnteredFormat: {
                            numberFormat: { type: 'NUMBER', pattern: '0.000' },
                            horizontalAlignment: 'RIGHT'
                        }
                    }, // Time
                    { userEnteredValue: { stringValue: item.data.heatName || '-' }, userEnteredFormat: { horizontalAlignment: 'RIGHT' } } // HEAT
                ]
            }));
            requests.push({
                updateCells: {
                    rows: rows,
                    start: { sheetId: sheetId, rowIndex: 3, columnIndex: 0 },
                    fields: "userEnteredValue,userEnteredFormat"
                }
            });
        }

        // --- 外枠の罫線を設定 ---
        const borderRange = {
            sheetId: sheetId,
            startRowIndex: 0,
            endRowIndex: 3 + rankingData.length,
            startColumnIndex: 0,
            endColumnIndex: 4
        };
        requests.push(
            { updateBorders: { range: borderRange, top: { style: 'SOLID', width: 1 } } },
            { updateBorders: { range: borderRange, bottom: { style: 'SOLID', width: 1 } } },
            { updateBorders: { range: borderRange, left: { style: 'SOLID', width: 1 } } },
            { updateBorders: { range: borderRange, right: { style: 'SOLID', width: 1 } } }
        );

        // --- 列幅を固定値に設定 ---
        requests.push(
            {
                updateDimensionProperties: {
                    range: {
                        sheetId: sheetId,
                        dimension: 'COLUMNS',
                        startIndex: 0, // A列
                        endIndex: 1
                    },
                    properties: {
                        pixelSize: 50
                    },
                    fields: 'pixelSize'
                }
            },
            {
                updateDimensionProperties: {
                    range: {
                        sheetId: sheetId,
                        dimension: 'COLUMNS',
                        startIndex: 1, // B列
                        endIndex: 2
                    },
                    properties: {
                        pixelSize: 150
                    },
                    fields: 'pixelSize'
                }
            },
            {
                updateDimensionProperties: {
                    range: {
                        sheetId: sheetId,
                        dimension: 'COLUMNS',
                        startIndex: 2, // C列
                        endIndex: 3
                    },
                    properties: {
                        pixelSize: 60
                    },
                    fields: 'pixelSize'
                }
            },
            {
                updateDimensionProperties: {
                    range: {
                        sheetId: sheetId,
                        dimension: 'COLUMNS',
                        startIndex: 3, // D列
                        endIndex: 4
                    },
                    properties: {
                        pixelSize: 130
                    },
                    fields: 'pixelSize'
                }
            }
        );

        // --- batchUpdateの実行 ---
        if (requests.length > 0) {
            await sheets.spreadsheets.batchUpdate({
                spreadsheetId,
                resource: { requests },
            });
        }

        console.log(`Successfully updated Google Sheet "${sheetTitle}".`);
    } catch (err) {
        console.error(`Error updating ${sheetTitle} Sheet:`, err);
    }
}

async function updateAllRankingSheets(pilotBests, allPilots, allValidLapTimes) {
    await updateSingleRankingSheet('bestLap', 'Best Lap', pilotBests, allPilots, 1);
    await updateSingleRankingSheet('consecutive2Lap', 'Best 2-Lap', pilotBests, allPilots, 2);
    await updateSingleRankingSheet('consecutive3Lap', 'Best 3-Lap', pilotBests, allPilots, 3);
    await updateSingleRankingSheet('raceTime', 'Best Race Time', pilotBests, allPilots, 4);
    await updateSingleRankingSheet('minLap', 'Minimum Lap Ranking', allValidLapTimes, null, 5);
}


// グローバル変数としてキャッシュを定義
let cachedPilotBests = {};
let cachedAllPilots = {};
let cachedEventName = '';
let cachedLastHeatName = null; // 新しいキャッシュ変数
let cachedLastHeatPilotIds = []; // 新しいキャッシュ変数

// 初期実行
processEvents();

// events ディレクトリの監視 (デバウンス処理を追加)
fs.watch(eventsDir, { recursive: true }, (eventType, filename) => {
    if (filename) {
        const triggerFiles = ['Event.json', 'Pilots.json', 'Rounds.json', 'Race.json', 'Result.json'];
        // 変更されたファイルが、処理に必要なファイルのいずれかで終わるかチェック
        const isTriggerFile = triggerFiles.some(file => filename.endsWith(file));

        if (!isTriggerFile) {
            return; // 処理対象外のファイルなので何もしない
        }

        console.log(`Detected ${eventType} in ${filename}. Debouncing...`);
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => {
            console.log('Reprocessing events after debounce...');
            processEvents();
        }, DEBOUNCE_DELAY);
    }
});

console.log(`Watching for changes in ${eventsDir}...`);

// Web UIサーバー
const http = require('http');
const server = http.createServer((req, res) => {
    if (req.url === '/') {
        fs.readFile(path.join(__dirname, 'index.html'), (err, data) => {
            if (err) {
                res.writeHead(500);
                res.end('Error loading index.html');
                return;
            }
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(data);
        });
    } else if (req.url === '/api/config') { // ★ /api/config エンドポイントを追加
        if (req.method === 'GET') {
            fs.readFile(path.join(__dirname, 'config.json'), 'utf8', (err, data) => {
                if (err) {
                    res.writeHead(500, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: 'Failed to read config.json' }));
                    return;
                }
                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(data);
            });
        } else if (req.method === 'POST') {
            let body = '';
            req.on('data', chunk => {
                body += chunk.toString();
            });
            req.on('end', () => {
                try {
                    const newConfig = JSON.parse(body);
                    // 既存のconfigを読み込み、新しい設定で上書きする
                    const currentConfig = JSON.parse(fs.readFileSync(path.join(__dirname, 'config.json'), 'utf8'));
                    const updatedConfig = { ...currentConfig, ...newConfig };

                    // 新しい設定項目にデフォルト値を設定（もしnewConfigに含まれていなければ）
                    if (updatedConfig.leaderboard_round === undefined) {
                        updatedConfig.leaderboard_round = "all";
                    }
                    if (updatedConfig.sorted_by === undefined) {
                        updatedConfig.sorted_by = "bestLap";
                    }

                    fs.writeFile(path.join(__dirname, 'config.json'), JSON.stringify(updatedConfig, null, 2), 'utf8', (err) => {
                        if (err) {
                            res.writeHead(500, { 'Content-Type': 'application/json' });
                            res.end(JSON.stringify({ error: 'Failed to write config.json' }));
                            return;
                        }
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify({ message: 'Config saved successfully' }));
                        
                        // ★ 設定変更後にprocessEventsを再実行
                        console.log('Config updated, reprocessing events...');
                        processEvents();
                    });
                } catch (err) {
                    res.writeHead(400, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: 'Invalid JSON' }));
                }
            });
        } else {
            res.writeHead(405); // Method Not Allowed
            res.end('Method Not Allowed');
        }
    } else if (req.url === '/api/events') { // ★ /api/events エンドポイントを追加
        if (req.method === 'GET') {
            fs.promises.readdir(eventsDir)
                .then(files => {
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
                    return Promise.all(eventPromises);
                })
                .then(events => {
                    const validEvents = events.filter(e => e !== null);
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify(validEvents));
                })
                .catch(error => {
                    console.error('Error loading events:', error);
                    res.writeHead(500, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: 'Failed to load events' }));
                });
        } else {
            res.writeHead(405);
            res.end('Method Not Allowed');
        }
    }
    else if (req.url === '/api/rounds') {
        if (req.method === 'GET') {
            fs.readFile(path.join(__dirname, 'config.json'), 'utf8', async (err, configData) => {
                if (err) {
                    res.writeHead(500, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: 'Failed to read config.json' }));
                    return;
                }
                const config = JSON.parse(configData);
                const selectedEventId = config.selected_event_id;
                const fpvtrackside_dir_path = config.fpvtrackside_dir_path;
                const eventsDir = path.join(fpvtrackside_dir_path, 'events').replace(/\\/g, '/');

                if (!selectedEventId || selectedEventId === 'all') {
                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify([{ id: 'all', name: 'すべてのラウンド' }]));
                    return;
                }

                try {
                    const eventDir = path.join(eventsDir, selectedEventId);
                    const roundsJsonPath = path.join(eventDir, 'Rounds.json');

                    if (!fs.existsSync(roundsJsonPath)) {
                        res.writeHead(200, { 'Content-Type': 'application/json' });
                        res.end(JSON.stringify([{ id: 'all', name: 'すべてのラウンド' }]));
                        return;
                    }

                    const roundsData = JSON.parse(await fs.promises.readFile(roundsJsonPath, 'utf8'));
                    const rounds = roundsData
                        .filter(round => round.Valid === true) // Validがtrueのラウンドのみをフィルター
                        .map(round => ({
                            id: round.ID,
                            name: `${round.EventType}Round${round.RoundNumber}` // 表示形式を変更
                        }));

                    // 新しい「すべてのXラウンド」オプションを追加
                    const allEventTypeRounds = [
                        { id: 'allRace', name: 'すべてのRaceラウンド' },
                        { id: 'allPractice', name: 'すべてのPracticeラウンド' },
                        { id: 'allTimeTrial', name: 'すべてのTimeTrialラウンド' },
                        { id: 'allEndurance', name: 'すべてのEnduranceラウンド' }
                    ];
                    const responseRounds = [{ id: 'all', name: 'すべてのラウンド' }, ...allEventTypeRounds, ...rounds];

                    res.writeHead(200, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify(responseRounds));

                } catch (error) {
                    console.error('Error loading rounds:', error);
                    res.writeHead(500, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: 'Failed to load rounds' }));
                }
            });
        } else {
            res.writeHead(405);
            res.end('Method Not Allowed');
        }
    }
    else if (req.url === '/leaderboard') {
        fs.readFile(path.join(__dirname, 'leaderboard.html'), (err, data) => {
            if (err) {
                res.writeHead(500);
                res.end('Error loading leaderboard.html');
                return;
            }
            res.writeHead(200, { 'Content-Type': 'text/html' });
            res.end(data);
        });
    }
    else if (req.url === '/api/leaderboard') {
        if (req.method === 'GET') {
            fs.readFile(path.join(__dirname, 'config.json'), 'utf8', async (err, configData) => {
                if (err) {
                    res.writeHead(500, { 'Content-Type': 'application/json' });
                    res.end(JSON.stringify({ error: 'Failed to read config.json' }));
                    return;
                }
                const config = JSON.parse(configData);
                const selectedEventId = config.selected_event_id;
                const leaderboardRound = config.leaderboard_round;
                const sortedBy = config.sorted_by;

                // キャッシュされたデータを使用
                const pilotBests = cachedPilotBests;
                const allPilots = cachedAllPilots;
                const eventName = cachedEventName;

                let roundName = 'すべてのラウンド'; // デフォルト値
                if (leaderboardRound !== 'all' && selectedEventId !== 'all') {
                    // 新しいオプションの表示名を設定
                    const allEventTypeRoundNames = {
                        'allRace': 'すべてのRaceラウンド',
                        'allPractice': 'すべてのPracticeラウンド',
                        'allTimeTrial': 'すべてのTimeTrialラウンド',
                        'allEndurance': 'すべてのEnduranceラウンド'
                    };
                    if (allEventTypeRoundNames[leaderboardRound]) {
                        roundName = allEventTypeRoundNames[leaderboardRound];
                    } else {
                        // config.selected_event_id と config.leaderboard_round に基づいてラウンド名を取得
                        const fpvtrackside_dir_path = config.fpvtrackside_dir_path;
                        const eventsDir = path.join(fpvtrackside_dir_path, 'events').replace(/\\/g, '/');
                        try {
                            const eventDir = path.join(eventsDir, selectedEventId);
                            const roundsJsonPath = path.join(eventDir, 'Rounds.json');
                            if (fs.existsSync(roundsJsonPath)) {
                                const roundsData = JSON.parse(await fs.promises.readFile(roundsJsonPath, 'utf8'));
                                const targetRound = roundsData.find(r => r.ID === leaderboardRound);
                                if (targetRound) {
                                    roundName = `${targetRound.EventType}Round${targetRound.RoundNumber}`; // 表示形式を合わせる
                                }
                            }
                        } catch (e) {
                            console.error('Error getting round name for leaderboard:', e);
                        }
                    }
                }


                // sortedBy の表示名
                const sortedByDisplayNames = {
                    "bestLap": "CONSECUTIVE 1 LAP (WITHOUT HS)",
                    "consecutive2Lap": "CONSECUTIVE 2 LAPS (WITHOUT HS)",
                    "consecutive3Lap": "CONSECUTIVE 3 LAPS (WITHOUT HS)",
                    "raceTime": "Race Time",
                    "first1LapWithoutHs": "First 1 LAP (WITHOUT HS)",
                    "first2LapsWithoutHs": "First 2 LAPS (WITHOUT HS)",
                    "first3LapsWithoutHs": "First 3 LAPS (WITHOUT HS)",
                    "first1LapWithHs": "First 1 LAP (WITH HS)",
                    "first2LapsWithHs": "First 2 LAPS (WITH HS)",
                    "first3LapsWithHs": "First 3 LAPS (WITH HS)"
                };
                const sortedByDisplayName = sortedByDisplayNames[sortedBy] || sortedBy;


                let ranking = [];

                // sortedBy の値に基づいてランキングを生成
                if (cachedPilotBests && Object.keys(cachedPilotBests).length > 0) {
                    ranking = Object.keys(cachedPilotBests)
                        .map(pilotId => {
                            const pilotName = cachedAllPilots[pilotId] || pilotId;
                            const data = cachedPilotBests[pilotId][sortedBy];

                            // データが存在しない、または無効な場合はランキングから除外
                            if (!data || typeof data.time !== 'number' || !isFinite(data.time) || data.time === 999 || data.time === 9999) {
                                return null;
                            }

                            // heatRoundName と heatEventType はここで一度だけ宣言
                            const heatRoundName = data.heatName ? data.heatName.split(' ')[0] + 'Round' + data.heatName.split(' ')[1].split('-')[0] : '';
                            const heatEventType = data.heatName ? data.heatName.split(' ')[0] : '';

                            if (leaderboardRound !== 'all') {
                                // 「すべてのXラウンド」オプションの場合
                                if (leaderboardRound === 'allRace' && heatEventType !== 'Race') {
                                    return null;
                                }
                                if (leaderboardRound === 'allPractice' && heatEventType !== 'Practice') {
                                    return null;
                                }
                                if (leaderboardRound === 'allTimeTrial' && heatEventType !== 'TimeTrial') {
                                    return null;
                                }
                                if (leaderboardRound === 'allEndurance' && heatEventType !== 'Endurance') {
                                    return null;
                                }
                                // 個別のラウンドIDが選択されている場合
                                if (!leaderboardRound.startsWith('all') && heatRoundName !== roundName) {
                                    return null;
                                }
                            }

                            // ここで新しいフラグを追加
                            const isBestTime = (data.time === cachedPilotBests[pilotId][sortedBy].time);

                            return {
                                pilotId: pilotId,
                                pilotName: pilotName,
                                time: data.time,
                                heatName: data.heatName,
                                isBestTime: isBestTime
                            };
                        })
                        .filter(item => item !== null)
                        .sort((a, b) => a.time - b.time);
                }

                res.writeHead(200, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({
                    eventName: eventName,
                    roundName: roundName,
                    sortedByDisplayName: sortedByDisplayName,
                    ranking: ranking,
                    lastHeatName: cachedLastHeatName,
                    lastHeatPilotIds: cachedLastHeatPilotIds
                }));

            });
        } else {
            res.writeHead(405);
            res.end('Method Not Allowed');
        }
    }
    else {
        res.writeHead(404);
        res.end('Not Found');
    }
});

server.listen(config.web_ui_port, () => {
    console.log(`Web UI running at http://localhost:${config.web_ui_port}`);
});