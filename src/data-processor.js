const fs = require('fs');
const path = require('path');
const { loadConfig, eventsDir } = require('./config');
const { sanitizeRaceResults, updateGoogleSheet, updateAllRankingSheets } = require('./google-sheets');

// メインの処理を関数としてラップ
async function processEvents(updateCacheCallback) {
    try {
        const config = loadConfig(); // 処理開始時に最新のconfigを読み込む
        const selectedEventId = config.selected_event_id || 'all';

        // --- Read GLOBAL Channels.json ---
        const channelsJsonPath = path.join(eventsDir, '..', 'httpfiles', 'Channels.json');
        const channelMap = new Map();
        if (fs.existsSync(channelsJsonPath)) {
            const channelsData = JSON.parse(fs.readFileSync(channelsJsonPath, 'utf8'));
            channelsData.forEach(channel => {
                channelMap.set(channel.ID, channel.DisplayName); // DisplayName を使用
            });
        } else {
            console.warn(`Warning: Global Channels.json not found at ${channelsJsonPath}. Band info will be unavailable.`);
        }

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
        
        let pilotsInLatestHeat = new Set();
        const allRaces = []; // 全イベントの全レース情報を格納

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
                }
            }
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
        let currentConfig = loadConfig(); // ループ内で最新のconfigを取得
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

        // --- latestHeatName と nextHeatName を決定する ---
        const allFinishedRacesWithTime = validRaces.filter(r => {
            return r.raceData[0].Laps && r.raceData[0].Laps.length > 0;
        }).map(r => {
            const firstLap = r.raceData[0].Laps.sort((a, b) => a.LapNumber - b.LapNumber)[0];
            return { race: r, startTime: firstLap ? new Date(firstLap.StartTime).getTime() : 0 };
        }).sort((a, b) => b.startTime - a.startTime);


        let latestRace = null;
        let latestHeatName = null;
        if (allFinishedRacesWithTime.length > 0) {
            latestRace = allFinishedRacesWithTime[0].race;
            latestHeatName = `${latestRace.eventType} ${latestRace.roundNumber === 0 ? 'N/A' : latestRace.roundNumber}-${latestRace.raceNumber}`;
            pilotsInLatestHeat = new Set(latestRace.raceData[0].Detections.map(d => d.Pilot));
        }

        // --- Next Heat を決定するロジック ---
        let nextHeatName = null;
        let nextHeatPilots = [];
        const findNextHeat = (startIndex) => {
            for (let i = startIndex; i < allRaces.length; i++) {
                const race = allRaces[i];
                const isStarted = race.raceData[0].Laps && race.raceData[0].Laps.length > 0;
                if (race.raceData[0].Valid === true && !isStarted) {
                    return race; // レースオブジェクト全体を返す
                }
            }
            return null;
        };

        let nextRace = null;
        if (latestRace) {
            const latestRaceIndex = allRaces.findIndex(r => r.id === latestRace.id);
            if (latestRaceIndex !== -1) {
                nextRace = findNextHeat(latestRaceIndex + 1);
            } else {
                nextRace = findNextHeat(0); // フォールバック
            }
        } else {
            // 完了したValidなレースがない場合
            nextRace = findNextHeat(0);
        }

        if (nextRace) {
            nextHeatName = `${nextRace.eventType} ${nextRace.roundNumber === 0 ? 'N/A' : nextRace.roundNumber}-${nextRace.raceNumber}`;
            const pilotsInRace = nextRace.raceData[0].PilotChannels; 
            if (pilotsInRace) {
                nextHeatPilots = pilotsInRace.map(racePilot => {
                    const pilot = nextRace.pilotsData.find(p => p.ID === racePilot.Pilot);
                    const bandInfo = channelMap.get(racePilot.Channel) || 'N/A';
                    return {
                        pilotId: pilot ? pilot.ID : null,
                    pilotName: pilot ? pilot.Name : 'Unknown Pilot',
                    photopath: pilot ? pilot.PhotoPath : null, // Added photopath
                    band: bandInfo
                    };
                });
            }
        }

        // --- Add Leaderboard data to nextHeatPilots ---
        currentConfig = loadConfig(); // Get latest config for sorted_by
        const sortedBy = currentConfig.sorted_by || 'bestLap'; // Default to bestLap

        // Create a temporary ranking from pilotBests to determine ranks
        let tempRanking = Object.keys(pilotBests)
            .map(pilotId => {
                const data = pilotBests[pilotId][sortedBy];
                if (!data || typeof data.time !== 'number' || !isFinite(data.time) || data.time >= 999) {
                    return null;
                }
                return { pilotId, time: data.time };
            })
            .filter(item => item !== null)
            .sort((a, b) => a.time - b.time);

        // Map pilotId to its rank and time for quick lookup
        const pilotRankAndTimeMap = new Map();
        tempRanking.forEach((item, index) => {
            pilotRankAndTimeMap.set(item.pilotId, { rank: index + 1, time: item.time });
        });

        // Update nextHeatPilots with rank and time
        nextHeatPilots = nextHeatPilots.map(pilot => {
            const rankAndTime = pilotRankAndTimeMap.get(pilot.pilotId);
            return {
                ...pilot,
                rank: rankAndTime ? rankAndTime.rank : null,
                time: rankAndTime ? rankAndTime.time : null
            };
        });

        

        // --- Web表示を先に更新 ---
        const webData = {
            pilotBests,
            allPilots,
            eventName,
            latestHeatName,
            nextHeatName,
            nextHeatPilots,
            lastHeatPilotIds: Array.from(pilotsInLatestHeat)
        };

        if (updateCacheCallback) {
            updateCacheCallback(webData);
        }

        // --- 時間のかかるGoogle Sheetの更新を後で行う ---
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
        await updateGoogleSheet(sanitizedData, lapsToDo);
        console.log('RaceResult sheet has been updated.');

    } catch (err) {
        console.error('Error processing events:', err);
        // エラー時もコールバックを呼ぶことで、サーバーが古い情報を持ち続けないようにする（オプション）
        if (updateCacheCallback) {
            updateCacheCallback(null);
        }
    }
}

module.exports = { processEvents };

