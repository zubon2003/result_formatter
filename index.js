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

        let allResultsText = '';
        const raceResults = [];
        let lapsToDo = 4; // Default laps

        const pilotBests = {};
        const allPilots = {};

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

            const eventName = eventData[0].Name;
            lapsToDo = eventData[0].Laps; // Get laps from event data

            allResultsText += 'Event: ' + eventName + '\n\n';

            const raceDirs = fs.readdirSync(eventDir).filter(file => {
                const raceDir = path.join(eventDir, file);
                return fs.statSync(raceDir).isDirectory();
            });

            const races = [];
            for (const raceDir of raceDirs) {
                const raceJsonPath = path.join(eventDir, raceDir, 'Race.json');
                const resultJsonPath = path.join(eventDir, raceDir, 'Result.json');

                if (fs.existsSync(raceJsonPath) && fs.existsSync(resultJsonPath)) {
                    const raceData = JSON.parse(fs.readFileSync(raceJsonPath, 'utf8'));
                    const resultData = JSON.parse(fs.readFileSync(resultJsonPath, 'utf8'));
                    const round = roundsData.find(r => r.ID === raceData[0].Round);
                    races.push({
                        roundNumber: round ? round.RoundNumber : 0,
                        raceNumber: raceData[0].RaceNumber,
                        raceData,
                        resultData
                    });
                }
            }

            races.sort((a, b) => a.roundNumber - b.roundNumber || a.raceNumber - b.raceNumber);

            const validRaces = races.filter(race => race.resultData && race.resultData.length > 0 && race.raceData[0].Valid === true);

            validRaces.forEach((race, index) => {
                const { roundNumber, raceNumber, raceData, resultData } = race;
                const displayRoundNumber = roundNumber === 0 ? 'N/A' : roundNumber;
                
                // resultDataからResultTypeを取得
                const resultType = resultData[0] && resultData[0].ResultType ? resultData[0].ResultType : 'Race'; // デフォルトは'Race'
                
                const raceName = resultType + ' ' + displayRoundNumber + '-' + raceNumber; // ★変更
                allResultsText += raceName + '\n';

                // 最も小さいLapNumberのラップから日付とスタート時刻を取得
                let raceSerialTimestamp = '';
                const firstLap = raceData[0].Laps.sort((a, b) => a.LapNumber - b.LapNumber)[0];
                if (firstLap && firstLap.StartTime) {
                    const dateObj = new Date(firstLap.StartTime);
                    // Google Sheetsのシリアル値に変換 (UTC基準)
                    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
                    raceSerialTimestamp = (dateObj.getTime() - excelEpoch.getTime()) / (24 * 60 * 60 * 1000);
                }

                resultData.sort((a, b) => a.Position - b.Position);

                resultData.forEach(result => {
                    const pilot = pilotsData.find(p => p.ID === result.Pilot);
                    if (pilot) {
                        if (!allPilots[pilot.ID]) {
                            allPilots[pilot.ID] = pilot.Name;
                        }
                        if (!pilotBests[pilot.ID]) {
                            pilotBests[pilot.ID] = {
                                raceTime: { time: 9999, timestamp: null, heatName: null },
                                bestLap: { time: 999, timestamp: null, heatName: null },
                                consecutive2Lap: { time: 999, timestamp: null, heatName: null },
                                consecutive3Lap: { time: 999, timestamp: null, heatName: null }
                            };
                        }

                        // ラップタイムの収集
                        const pilotLaps = raceData[0].Laps.filter(lap => {
                            const detection = raceData[0].Detections.find(d => d.ID === lap.Detection);
                            return detection && detection.Pilot === result.Pilot && detection.Valid === true;
                        }).sort((a, b) => a.LapNumber - b.LapNumber);

                        const lapTimes = Array(31).fill(''); // HS(LAP0)からLAP30まで、空文字列で初期化
                        let totalFinishTime = '';
                        let lapCount = pilotLaps.length;
                        
                        let bestLap = 999;
                        let consecutive2Lap = 999;
                        let consecutive3Lap = 999;
                        let raceTimeXLap = 9999;

                        // LAP1以上のラップタイムのみを抽出
                        const actualLapTimes = pilotLaps.filter(lap => lap.LapNumber >= 1).map(lap => lap.LengthSeconds);

                        if (pilotLaps.length > 0) {
                            lapTimes[0] = pilotLaps[0].LengthSeconds; // HS(LAP0)
                            totalFinishTime = pilotLaps.reduce((sum, lap) => sum + lap.LengthSeconds, 0);

                            // Best LAPの計算 (LAP1以上のラップから)
                            if (actualLapTimes.length > 0) {
                                const minLap = Math.min(...actualLapTimes);
                                if (isFinite(minLap)) {
                                    bestLap = minLap;
                                }
                            }

                            // 連続2周の計算 (LAP1以上のラップから)
                            if (actualLapTimes.length >= 2) {
                                let min2Lap = Infinity;
                                for (let i = 0; i < actualLapTimes.length - 1; i++) {
                                    min2Lap = Math.min(min2Lap, actualLapTimes[i] + actualLapTimes[i+1]);
                                }
                                if (isFinite(min2Lap)) {
                                    consecutive2Lap = min2Lap;
                                }
                            }

                            // 連続3周の計算 (LAP1以上のラップから)
                            if (actualLapTimes.length >= 3) {
                                let min3Lap = Infinity;
                                for (let i = 0; i < actualLapTimes.length - 2; i++) {
                                    min3Lap = Math.min(min3Lap, actualLapTimes[i] + actualLapTimes[i+1] + actualLapTimes[i+2]);
                                }
                                if (isFinite(min3Lap)) {
                                    consecutive3Lap = min3Lap;
                                }
                            }

                            const hsLap = pilotLaps.find(lap => lap.LapNumber === 0);

                            if (hsLap) {
                                // HSありの場合: HS(LAP0)からLAPxまで完走しているか
                                if (pilotLaps.length >= lapsToDo + 1) {
                                    raceTimeXLap = pilotLaps.slice(0, lapsToDo + 1).reduce((sum, lap) => sum + lap.LengthSeconds, 0);
                                }
                            } else {
                                // HSなしの場合: LAP1からLAPxまで完走しているか
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

                        updateBestTime('raceTime', raceTimeXLap, raceSerialTimestamp, raceName);
                        updateBestTime('bestLap', bestLap, raceSerialTimestamp, raceName);
                        updateBestTime('consecutive2Lap', consecutive2Lap, raceSerialTimestamp, raceName);
                        updateBestTime('consecutive3Lap', consecutive3Lap, raceSerialTimestamp, raceName);

                        allResultsText += `  Position: ${result.Position}, Pilot: ${pilot.Name}, Points: ${result.Points}\n`;
                        raceResults.push([
                            eventName,
                            raceName,
                            raceSerialTimestamp, // 日付用
                            raceSerialTimestamp, // 時刻用
                            pilot.Name,
                            result.Position,
                            lapCount, // Lap数
                            totalFinishTime, // Finish time
                            raceTimeXLap,
                            bestLap,
                            consecutive2Lap,
                            consecutive3Lap,
                            ...lapTimes // HS(LAP0)からLAP30まで
                        ]);

                    }
                });
                 if (index < validRaces.length - 1) {
                    allResultsText += '\n';
                }
            });
        }

        

        // Google Spreadsheetへの書き込み
        await updateGoogleSheet(raceResults, lapsToDo);
        await updateAllRankingSheets(pilotBests, allPilots);

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
                start: { sheetId: sheetId, rowIndex: 0, columnIndex: 0 },
                fields: "userEnteredValue,userEnteredFormat"
            }
        });

        if (raceResults.length > 0) {
            const colors = [
                { red: 0.9, green: 0.9, blue: 0.9 }, // 灰色 (light gray)
                { red: 1, green: 1, blue: 1 }       // 白
            ];
            let colorIndex = 0;
            let currentHeatStartRow = 1; // データは2行目 (0-indexedで1) から始まる
            let previousHeatName = raceResults[0][1]; // 最初のHEAT名

            for (let i = 0; i < raceResults.length; i++) {
                const heatName = raceResults[i][1]; // レース名（HEAT）は2列目（インデックス1）
                const sheetRowIndex = i + 1; // シートの現在のデータ行インデックス (0-indexed)

                // HEATが変わった場合、または最終行の場合
                if (heatName !== previousHeatName || i === raceResults.length - 1) {
                    // HEATが変わる前の行に罫線を追加 (最初のHEAT以外で、HEATが変わる場合)
                    if (i > 0 && heatName !== previousHeatName) { // 最初のHEAT以外で、HEATが変わる場合
                        requests.push({
                            updateBorders: {
                                range: {
                                    sheetId: sheetId,
                                    startRowIndex: sheetRowIndex - 1, // 前のHEATの最終データ行
                                    endRowIndex: sheetRowIndex, // 罫線を入れる行
                                    startColumnIndex: 0,
                                    endColumnIndex: 43 // 全ての列
                                },
                                bottom: { style: 'SOLID', width: 1 }
                            }
                        });
                    }

                    requests.push({
                        addBanding: {
                            bandedRange: {
                                range: {
                                    sheetId: sheetId,
                                    startRowIndex: currentHeatStartRow,
                                    endRowIndex: sheetRowIndex, // このHEATの最終行まで
                                    startColumnIndex: 0, // A列から
                                    endColumnIndex: 43    // すべての列
                                },
                                rowProperties: {
                                    firstBandColor: colors[colorIndex],
                                    secondBandColor: colors[(colorIndex + 1) % colors.length] // 白
                                }
                            }
                        }
                    });

                    // 次のHEATのためにリセット
                    currentHeatStartRow = sheetRowIndex;
                    previousHeatName = heatName;
                    colorIndex = (colorIndex + 1) % colors.length;
                }
            }
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
            { index: 1, size: 80 },  // B
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

async function updateAllRankingSheets(pilotBests, allPilots) {
    await updateSingleRankingSheet('bestLap', 'Best Lap', pilotBests, allPilots, 1);
    await updateSingleRankingSheet('consecutive2Lap', 'Best 2-Lap', pilotBests, allPilots, 2);
    await updateSingleRankingSheet('consecutive3Lap', 'Best 3-Lap', pilotBests, allPilots, 3);
    await updateSingleRankingSheet('raceTime', 'Best Race Time', pilotBests, allPilots, 4);
}

async function updateSingleRankingSheet(categoryKey, sheetTitle, pilotBests, allPilots, index) {
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
        let rankingData = Object.keys(pilotBests)
            .map(pilotId => ({
                pilotId: pilotId,
                pilotName: allPilots[pilotId] || pilotId,
                data: pilotBests[pilotId][categoryKey]
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
                        pixelSize: 80
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


// 初期実行
processEvents();

// events ディレクトリの監視 (デバウンス処理を追加)
fs.watch(eventsDir, { recursive: true }, (eventType, filename) => {
    if (filename) {
        // recordinfo.xmlで終わるファイルは無視する
        if (filename.endsWith('recordinfo.xml')) {
            return;
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
                    fs.writeFile(path.join(__dirname, 'config.json'), JSON.stringify(newConfig, null, 2), 'utf8', (err) => {
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
    else {
        res.writeHead(404);
        res.end('Not Found');
    }
});

server.listen(config.web_ui_port, () => {
    console.log(`Web UI running at http://localhost:${config.web_ui_port}`);
});