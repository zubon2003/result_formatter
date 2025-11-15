const { google } = require('googleapis');
const { config, credentialsPath } = require('./config');

// Google Sheets APIの認証
const auth = new google.auth.GoogleAuth({
    keyFile: credentialsPath,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

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

async function updateGoogleSheet(raceResults, lapsToDo) {
    try {
        const spreadsheetId = config.google_spreadsheet_id;
        if (!spreadsheetId) {
            console.warn('google_spreadsheet_id is not set in config.json. Skipping Google Sheet update.');
            return;
        }
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
                    start: { sheetId: sheetId, rowIndex: 1, columnIndex: 0 },
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
        if (!spreadsheetId) {
            // google_spreadsheet_id がなければ何もしない
            return;
        }

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
                    pilotName: (allPilots[pilotId] ? allPilots[pilotId].Name : pilotId),
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

module.exports = {
    sanitizeRaceResults,
    updateGoogleSheet,
    updateAllRankingSheets
};
