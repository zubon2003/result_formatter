/*
 * FPVTrackside 公開ビュー (モバイル対応)
 *  - ステージ → ラウンド → レース の階層
 *  - ラウンドは <details> で開閉
 *  - 全レースを表示 (未実施 / 進行中 / 確定)
 *  - ステージごとの結果(standings)を集計表示
 * データ層は FPVTrackside Webb の EventManager / Accessor を流用。
 */
(function () {
    "use strict";

    // bundle モード: EventManager は eventDirectory="" で生成し、
    // accessor を BundleAccessor に差し替えて単一 bundle から供給する。
    let em = null;

    let LOADED_VERSION = null;   // 読み込んだ bundle の generatedAt
    let pendingReload = false;   // 詳細表示中に更新が来たら、閉じてからリロード
    const AUTO_RELOAD_MS = 10000; // version.json のポーリング間隔 (エッジキャッシュ10sと整合)

    // チャンネル変更(Δ)判定用: pilotId -> [{order, channelId}] (order 昇順)
    // 本体FPVTracksideと同じく RaceOrder = Round.Order + RaceNumber で前後を判定する。
    let pilotHistory = {};

    function raceOrderOf(round, race) {
        return (round.Order || 0) + (race.RaceNumber || 0);
    }

    // そのレース(order)でのパイロットのチャンネルが、直前レースから変わっているか
    function pilotChannelChanged(pilotId, channelId, order) {
        const hist = pilotHistory[pilotId];
        if (!hist) return false;
        let prev = null;
        for (const e of hist) {
            if (e.order < order) prev = e; // 昇順なので最後に拾ったものが直前
            else break;
        }
        return prev ? (prev.channelId !== channelId) : false;
    }

    const el = (tag, cls, html) => {
        const e = document.createElement(tag);
        if (cls) e.className = cls;
        if (html != null) e.innerHTML = html;
        return e;
    };
    const esc = (s) => String(s == null ? "" : s).replace(/[&<>"]/g, (c) =>
        ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
    let DP = 2; // application settings の ShownDecimalPlaces (bundle から設定)
    const fmtTime = (s) => (s == null || !isFinite(s)) ? "" : Number(s).toFixed(DP);

    let EVENT = null; // GetEvent() の結果を保持

    const pilotLink = (pilot) => pilot
        ? `<a href="#pilot=${encodeURIComponent(pilot.ID)}" class="pilot-link" data-pilot="${esc(pilot.ID)}">${esc(pilot.Name)}</a>`
        : "";

    async function fetchJsonArray(url) {
        try {
            const r = await fetch(url);
            if (!r.ok) return [];
            return await r.json();
        } catch (e) { return []; }
    }

    function raceState(race) {
        const notStarted = !race.Start || String(race.Start).startsWith("0001");
        const notEnded = !race.End || String(race.End).startsWith("0001");
        if (notStarted) return { key: "pending", label: "Pending" };
        if (notEnded) return { key: "live", label: "Live" };
        return { key: "done", label: "Finished" };
    }

    function channelChip(channel) {
        if (!channel) return "";
        const color = channel.Color || "#888";
        const label = esc((channel.ShortBand || "") + (channel.Number != null ? channel.Number : ""));
        return `<span class="chip"><span class="dot" style="background:${esc(color)}"></span><span class="ch-label">${label}</span></span>`;
    }

    function posClass(p) {
        if (p === 1) return "pos-1";
        if (p === 2) return "pos-2";
        if (p === 3) return "pos-3";
        return "";
    }

    // ---- 1 レース分の結果テーブル -------------------------------------
    async function buildRace(race, round) {
        const state = raceState(race);
        const order = raceOrderOf(round, race);

        const card = el("div", "race");
        const head = el("div", "race-head");
        const raceName = `${round.EventType} ${round.RoundNumber}-${race.RaceNumber}`;
        // 確定(Finished)レースのみ詳細へのリンクにする
        if (state.key === "done") {
            const nameLink = el("a", "race-name race-link");
            nameLink.href = "#race=" + encodeURIComponent(race.ID);
            nameLink.textContent = raceName;
            head.appendChild(nameLink);
        } else {
            head.appendChild(el("span", "race-name", esc(raceName)));
        }
        head.appendChild(el("span", "badge " + state.key, state.label));
        card.appendChild(head);

        // 結果の取得 (未実施なら空)
        const results = await em.GetResults(race.ID);
        const byPilot = {};
        if (Array.isArray(results)) {
            for (const r of results) {
                if (r && r.Valid) byPilot[r.Pilot] = r;
            }
        }

        // 出走パイロット
        const rows = [];
        for (const pc of race.PilotChannels) {
            const pilot = await em.GetPilot(pc.Pilot);
            const channel = await em.GetChannel(pc.Channel);
            if (!pilot || !channel) continue;
            const changed = pilotChannelChanged(pc.Pilot, pc.Channel, order);
            rows.push({ pilot, channel, res: byPilot[pc.Pilot] || null, changed });
        }

        // 埋まっていないチャンネル枠を空欄行として追加（本体の表示に合わせる）
        const filledIds = new Set(race.PilotChannels.map((pc) => pc.Channel));
        const evChannels = await em.GetEventChannels();
        for (const slot of evChannels) {
            if (filledIds.has(slot.ID)) continue;
            const channel = await em.GetChannel(slot.ID);
            if (channel) rows.push({ pilot: null, channel, res: null, empty: true });
        }

        const hasResults = rows.some((r) => r.res != null);

        rows.sort((a, b) => {
            if (hasResults) {
                const pa = a.res ? (a.res.DNF ? 9990 : a.res.Position || 9991) : 9999;
                const pb = b.res ? (b.res.DNF ? 9990 : b.res.Position || 9991) : 9999;
                if (pa !== pb) return pa - pb;
            }
            return a.channel.Frequency - b.channel.Frequency;
        });

        if (rows.length === 0) {
            card.appendChild(el("div", "empty", "No pilots"));
            return card;
        }

        let html = '<table class="grid"><thead><tr>';
        html += hasResults ? '<th class="pos">#</th>' : "<th></th>";
        html += '<th class="pilot">Pilot</th><th class="chcol">Ch</th>';
        if (hasResults) html += '<th class="num">Laps</th>';
        html += "</tr></thead><tbody>";

        for (const r of rows) {
            const res = r.res;
            let posCell = "";
            if (res) {
                posCell = res.DNF
                    ? '<span class="dnf">DNF</span>'
                    : `<span class="${posClass(res.Position)}">${res.Position || ""}</span>`;
            }
            html += r.empty ? '<tr class="slot-empty">' : "<tr>";
            html += hasResults ? `<td class="pos">${posCell}</td>` : "<td></td>";
            html += `<td class="pilot">${pilotLink(r.pilot)}</td>`;
            html += `<td class="chcol"><span class="chwrap"><span class="chg-slot">${r.changed ? '<span class="chg" title="Channel changed from previous race">Δ</span>' : ''}</span>${channelChip(r.channel)}</span></td>`;
            if (hasResults) {
                html += `<td class="num">${res && res.LapsFinished != null ? res.LapsFinished : ""}</td>`;
            }
            html += "</tr>";
        }
        html += "</tbody></table>";
        card.appendChild(el("div", null, html));
        return card;
    }

    // ---- ステージの standings 集計 ------------------------------------
    async function buildStandings(stage, races, isRace, event) {
        const pbLaps = event.PBLaps || 1;
        const acc = {}; // pilotId -> {pilot, points, wins, bestLap, pb, races}

        for (const race of races) {
            const results = await em.GetResults(race.ID);
            const resByPilot = {};
            if (Array.isArray(results)) for (const r of results) if (r && r.Valid) resByPilot[r.Pilot] = r;

            for (const pc of race.PilotChannels) {
                const pilot = await em.GetPilot(pc.Pilot);
                if (!pilot) continue;
                let a = acc[pilot.ID];
                if (!a) a = acc[pilot.ID] = { pilot, points: 0, wins: 0, bestLap: Infinity, pb: Infinity, races: 0 };
                a.races++;

                const res = resByPilot[pilot.ID];
                if (res) {
                    a.points += res.Points || 0;
                    if (!res.DNF && res.Position === 1) a.wins++;
                }

                const laps = em.GetValidLapsPilot(race, pilot.ID);
                const bl = em.BestLap(laps);
                if (bl < a.bestLap) a.bestLap = bl;

                const nonHole = em.ExcludeHoleshot(laps);
                const pbSet = em.BestConsecutive(nonHole, pbLaps);
                const pbTime = em.TotalTime(pbSet);
                if (pbTime < a.pb) a.pb = pbTime;
            }
        }

        const list = Object.values(acc);
        if (list.length === 0) return null;

        list.sort((a, b) => {
            if (isRace) {
                if (b.points !== a.points) return b.points - a.points;
                if (b.wins !== a.wins) return b.wins - a.wins;
                return a.pb - b.pb;
            }
            if (a.pb !== b.pb) return a.pb - b.pb;
            return a.bestLap - b.bestLap;
        });

        const details = el("details", "standings");
        details.appendChild(el("summary", null, "Stage standings"));

        let html = '<div><table class="grid"><thead><tr>';
        html += '<th class="pos">#</th><th class="pilot">Pilot</th>';
        if (isRace) html += '<th class="num">Wins</th>';
        html += `<th class="num">PB(${pbLaps})</th><th class="num">BestLap</th>`;
        html += "</tr></thead><tbody>";

        let rank = 0;
        for (const a of list) {
            rank++;
            const pb = isFinite(a.pb) ? fmtTime(a.pb) : "";
            const bl = isFinite(a.bestLap) ? fmtTime(a.bestLap) : "";
            html += "<tr>";
            html += `<td class="pos"><span class="${posClass(rank)}">${rank}</span></td>`;
            html += `<td class="pilot">${pilotLink(a.pilot)}</td>`;
            if (isRace) html += `<td class="num">${a.wins}</td>`;
            html += `<td class="num">${pb}</td><td class="num">${bl}</td>`;
            html += "</tr>";
        }
        html += "</tbody></table></div>";
        details.appendChild(el("div", null, html));
        return details;
    }

    // ---- 集計対象(サマリ)の判定 ---------------------------------------
    // FPVTrackside ではステージ/ラウンドに PointSummary・TimeSummary が付き、
    // それが付いたものが「集計対象」。無ければステージ結果は出さない。
    function getStageSummary(group, allRounds) {
        const st = group.stage || {};
        let type = null, summaryObj = null, scope = null, summaryRounds = null;

        if (st.PointSummary) { type = "points"; summaryObj = st.PointSummary; scope = "stage"; }
        else if (st.TimeSummary) { type = "time"; summaryObj = st.TimeSummary; scope = "stage"; }
        else {
            const pr = group.rounds.filter((r) => r.PointSummary);
            const tr = group.rounds.filter((r) => r.TimeSummary);
            if (pr.length) { type = "points"; scope = "rounds"; summaryRounds = pr; }
            else if (tr.length) { type = "time"; scope = "rounds"; summaryRounds = tr; }
        }
        if (!type) return null;

        let targetRounds;
        if (scope === "stage") {
            const includeAll = !!(summaryObj && summaryObj.IncludeAllRounds);
            targetRounds = includeAll ? allRounds.slice() : group.rounds.slice();
        } else {
            targetRounds = summaryRounds;
        }
        return { type, targetRounds };
    }

    // ---- ラウンド ------------------------------------------------------
    async function buildRound(round, event) {
        const races = await em.GetRoundRaces(round.ID);

        const details = el("details", "round");
        details.open = true;

        const summary = el("summary");
        const title = el("span", "round-title", esc(`${round.EventType} Round ${round.RoundNumber}`));
        const sub = el("span", "round-sub", `${races.length} races`);
        summary.appendChild(title);
        summary.appendChild(sub);
        details.appendChild(summary);

        const body = el("div", "round-body");
        if (races.length === 0) {
            body.appendChild(el("div", "empty", "No races"));
        } else {
            for (const race of races) {
                body.appendChild(await buildRace(race, round));
            }
        }
        details.appendChild(body);
        return details;
    }

    // ---- メイン --------------------------------------------------------
    async function loadBundle() {
        const r = await fetch(BUNDLE_URL, { cache: "no-store" });
        if (!r.ok) throw new Error("Failed to load bundle.json (" + r.status + ")");
        return await r.json();
    }

    async function main() {
        const app = document.getElementById("app");

        // 単一 bundle を 1 回だけ取得し、以降は全てメモリから供給
        const bundle = await loadBundle();
        em = new EventManager("", TOO_OLD);
        em.accessor = new BundleAccessor(bundle);
        LOADED_VERSION = bundle.generatedAt || null;
        if (bundle.decimalPlaces != null) DP = bundle.decimalPlaces;

        const event = await em.GetEvent();
        if (!event) {
            app.innerHTML = '<div class="empty">Failed to load event.</div>';
            return;
        }
        EVENT = event;

        document.getElementById("eventName").textContent = event.Name || "Event";
        const rounds = await em.GetRounds(); // valid, Order 順
        document.getElementById("eventMeta").textContent =
            `${rounds.length} rounds / ${(event.Races || []).length} races`;

        // チャンネル変更(Δ)判定用に、全レースのパイロット→チャンネル履歴を構築
        pilotHistory = {};
        for (const r of rounds) {
            const rs = await em.GetRoundRaces(r.ID);
            for (const race of rs) {
                const order = raceOrderOf(r, race);
                for (const pc of race.PilotChannels) {
                    if (!pilotHistory[pc.Pilot]) pilotHistory[pc.Pilot] = [];
                    pilotHistory[pc.Pilot].push({ order, channelId: pc.Channel });
                }
            }
        }
        for (const k in pilotHistory) pilotHistory[k].sort((a, b) => a.order - b.order);

        // ステージ (bundle から)
        let stages = await em.accessor.GetJSON("Stages.json");
        stages = (Array.isArray(stages) ? stages : []).filter((s) => s && s.Valid)
            .sort((a, b) => (a.Order || 0) - (b.Order || 0));

        // ラウンドをステージへ割当 (Stage が無いものは "その他")
        const groups = [];
        const stageOf = {};
        for (const s of stages) { stageOf[s.ID] = { stage: s, rounds: [] }; groups.push(stageOf[s.ID]); }
        let other = null;
        for (const r of rounds) {
            if (r.Stage && stageOf[r.Stage]) {
                stageOf[r.Stage].rounds.push(r);
            } else {
                if (!other) { other = { stage: { Name: "Other" }, rounds: [] }; groups.push(other); }
                other.rounds.push(r);
            }
        }

        app.innerHTML = "";

        let rendered = 0;
        for (const g of groups) {
            if (!g.rounds.length) continue;
            rendered++;

            const isRace = g.rounds.some((r) => r.EventType === "Race");

            const section = el("section", "stage");
            const head = el("div", "stage-head");
            head.appendChild(el("span", "stage-title", esc(g.stage.Name || "Stage")));
            head.appendChild(el("span", "stage-kind", isRace ? "Race" : "Time"));
            section.appendChild(head);

            // ステージ結果: サマリ(集計対象)が定義されたステージのみ表示
            const summary = getStageSummary(g, rounds);
            if (summary) {
                const stageRaces = [];
                for (const r of summary.targetRounds) {
                    const rr = await em.GetRoundRaces(r.ID);
                    for (const race of rr) stageRaces.push(race);
                }
                const standings = await buildStandings(
                    g.stage, stageRaces, summary.type === "points", event
                );
                if (standings) section.appendChild(standings);
            }

            // ラウンド
            for (const r of g.rounds) {
                section.appendChild(await buildRound(r, event));
            }
            app.appendChild(section);
        }

        if (rendered === 0) {
            app.innerHTML = '<div class="empty">No rounds to display.</div>';
        }

        document.getElementById("updated").textContent =
            "Generated: " + new Date().toLocaleString();

        wireToolbar();
        initPilotRouter();
        startAutoReload();
    }

    // ---- パイロット詳細 -----------------------------------------------
    async function buildPilotDetail(pilotId) {
        const event = EVENT;
        const pbLaps = event.PBLaps || 1;
        const lapCount = event.Laps || 2;
        const showHoleshot = event.PrimaryTimingSystemLocation === "Holeshot";

        const pilot = await em.GetPilot(pilotId);
        const wrap = el("div", "pilot-detail");
        if (!pilot) {
            wrap.appendChild(el("div", "empty", "Pilot not found."));
            return { node: wrap, name: "?" };
        }

        const rounds = await em.GetRounds();

        // 自己ベスト集計 + ラウンドごとのラップ
        let bestLap = Infinity;
        let bestPb = Infinity;
        let bestLaps = Infinity;   // lapCount 連続
        let bestRace = Infinity;   // レースタイム(目標周回ちょうど)
        let totalLaps = 0;
        let raceCount = 0;

        const roundBlocks = [];

        for (const round of rounds) {
            const races = await em.GetRoundRaces(round.ID);
            for (const race of races) {
                if (!em.RaceHasPilot(race, pilotId)) continue;
                raceCount++;

                const laps = em.GetValidLapsPilot(race, pilotId);
                const nonHole = em.ExcludeHoleshot(laps);

                const pb = em.TotalTime(em.BestConsecutive(nonHole, pbLaps));
                if (pb < bestPb) bestPb = pb;
                const lc = em.TotalTime(em.BestConsecutive(nonHole, lapCount));
                if (lc < bestLaps) bestLaps = lc;

                const target = race.TargetLaps || lapCount;
                const raceLaps = (showHoleshot ? laps : nonHole);
                const need = showHoleshot ? target + 1 : target;
                if (raceLaps.length === need) {
                    const rt = em.TotalTime(raceLaps);
                    if (rt < bestRace) bestRace = rt;
                }

                const result = await em.GetPilotResult(race.ID, pilotId);

                let rows = "";
                for (const lap of laps) {
                    if (lap.LapNumber !== 0) {
                        totalLaps++;
                        if (lap.LengthSeconds < bestLap) bestLap = lap.LengthSeconds;
                    }
                    const name = lap.LapNumber === 0 ? "HS" : "Lap " + lap.LapNumber;
                    rows += `<tr><td>${name}</td><td class="num">${fmtTime(lap.LengthSeconds)}</td></tr>`;
                }
                if (!rows) rows = '<tr><td colspan="2" class="empty">No laps</td></tr>';

                let resBadge = "";
                if (result) {
                    resBadge = result.DNF
                        ? '<span class="dnf">DNF</span>'
                        : `<span class="${posClass(result.Position)}">P${result.Position || ""}</span>`;
                }

                roundBlocks.push(
                    `<div class="pd-race">
                        <div class="pd-race-head">
                            <span>${esc(round.EventType)} ${round.RoundNumber}-${race.RaceNumber}</span>
                            <span class="pd-res">${resBadge}</span>
                        </div>
                        <table class="grid"><tbody>${rows}</tbody></table>
                     </div>`
                );
            }
        }

        // 自己ベストカード
        const recRows = [];
        if (showHoleshot) recRows.push(["Holeshot", null]); // ホールショット最良は下のbest計算に含めない簡易版
        recRows.push(["Best Lap", bestLap]);
        recRows.push([`PB (${pbLaps} Lap)`, bestPb]);
        recRows.push([`Best ${lapCount} Laps`, bestLaps]);
        if (event.EventType === "Race") recRows.push(["Race Time", bestRace]);

        let recHtml = '<table class="grid"><tbody>';
        for (const [label, val] of recRows) {
            if (label === "Holeshot") continue;
            recHtml += `<tr><td>${esc(label)}</td><td class="num">${isFinite(val) ? fmtTime(val) : "—"}</td></tr>`;
        }
        recHtml += `<tr><td>Lap count</td><td class="num">${totalLaps}</td></tr>`;
        recHtml += `<tr><td>Races</td><td class="num">${raceCount}</td></tr>`;
        recHtml += "</tbody></table>";

        wrap.appendChild(el("h3", "pd-h", "Personal Bests"));
        wrap.appendChild(el("div", "pd-card", recHtml));

        wrap.appendChild(el("h3", "pd-h", "Laps by Round"));
        if (roundBlocks.length) {
            wrap.appendChild(el("div", "pd-races", roundBlocks.join("")));
        } else {
            wrap.appendChild(el("div", "empty", "No race history."));
        }

        return { node: wrap, name: pilot.Name };
    }

    // ---- レース詳細 ---------------------------------------------------
    function fmtClock(value) {
        const t = Date.parse(value);
        if (!isFinite(t) || t < 1000000000000) return null; // 0001/01/01 等は無効
        return new Date(t).toLocaleTimeString();
    }

    async function buildRaceDetail(raceId) {
        const race = await em.GetRace(raceId);
        const wrap = el("div", "race-detail");
        if (!race) {
            wrap.appendChild(el("div", "empty", "Race not found."));
            return { node: wrap, name: "?" };
        }
        const round = await em.GetRound(race.Round);
        const raceName = round
            ? `${round.EventType} ${round.RoundNumber}-${race.RaceNumber}`
            : `Race ${race.RaceNumber}`;
        const state = raceState(race);

        // 出走者 (pilotId -> {pilot, channel})
        const pilots = {};
        for (const pc of race.PilotChannels) {
            const pilot = await em.GetPilot(pc.Pilot);
            const channel = await em.GetChannel(pc.Channel);
            if (pilot && channel) pilots[pilot.ID] = { pilot, channel };
        }

        // メタ情報
        const start = fmtClock(race.Start);
        const end = fmtClock(race.End);
        let length = "";
        if (start && end) {
            const sec = (Date.parse(race.End) - Date.parse(race.Start)) / 1000;
            if (isFinite(sec) && sec > 0) length = fmtTime(sec) + "s";
        }
        let meta = `<span class="badge ${state.key}">${state.label}</span>`;
        meta += `<div class="rd-meta">Target Laps: ${race.TargetLaps != null ? race.TargetLaps : "-"}`;
        if (start) meta += ` · Start ${esc(start)}`;
        if (end) meta += ` · End ${esc(end)}`;
        if (length) meta += ` · ${esc(length)}`;
        meta += `</div>`;
        wrap.appendChild(el("div", "rd-head", meta));

        // 結果テーブル
        const results = await em.GetResults(race.ID);
        const byPilot = {};
        if (Array.isArray(results)) for (const r of results) if (r && r.Valid) byPilot[r.Pilot] = r;

        const rows = Object.values(pilots).map((p) => ({ ...p, res: byPilot[p.pilot.ID] || null }));
        const hasResults = rows.some((r) => r.res != null);
        rows.sort((a, b) => {
            if (hasResults) {
                const pa = a.res ? (a.res.DNF ? 9990 : a.res.Position || 9991) : 9999;
                const pb = b.res ? (b.res.DNF ? 9990 : b.res.Position || 9991) : 9999;
                if (pa !== pb) return pa - pb;
            }
            return a.channel.Frequency - b.channel.Frequency;
        });

        let rhtml = '<table class="grid"><thead><tr><th class="pos">#</th><th class="pilot">Pilot</th><th class="chcol">Ch</th><th class="num">Laps</th></tr></thead><tbody>';
        for (const r of rows) {
            const res = r.res;
            let posCell = res ? (res.DNF ? '<span class="dnf">DNF</span>' : `<span class="${posClass(res.Position)}">${res.Position || ""}</span>`) : "";
            rhtml += "<tr>";
            rhtml += `<td class="pos">${posCell}</td>`;
            rhtml += `<td class="pilot">${pilotLink(r.pilot)}</td>`;
            rhtml += `<td class="chcol"><span class="chwrap">${channelChip(r.channel)}</span></td>`;
            rhtml += `<td class="num">${res && res.LapsFinished != null ? res.LapsFinished : ""}</td>`;
            rhtml += "</tr>";
        }
        rhtml += "</tbody></table>";
        wrap.appendChild(el("h3", "pd-h", "Results"));
        wrap.appendChild(el("div", "pd-card", rhtml));

        // ラップ別の内訳
        const maxLapNumber = (round && round.EventType === "Race") ? (race.TargetLaps || 999) : 999;
        const allLaps = em.GetValidLaps(race);
        const grouped = {};
        for (const lap of allLaps) {
            if (lap.LapNumber > maxLapNumber) continue;
            if (!grouped[lap.LapNumber]) grouped[lap.LapNumber] = [];
            grouped[lap.LapNumber].push(lap);
        }

        const lapNumbers = Object.keys(grouped).map(Number).sort((a, b) => a - b);
        if (lapNumbers.length) {
            wrap.appendChild(el("h3", "pd-h", "Lap details"));
            const lapsWrap = el("div", "pd-races");
            for (const n of lapNumbers) {
                const laps = grouped[n];
                laps.sort((a, b) => Date.parse(a.detectionObject.Time) - Date.parse(b.detectionObject.Time));

                let html = `<div class="pd-race"><div class="pd-race-head"><span>${n === 0 ? "Holeshot" : "Lap " + n}</span></div>`;
                html += '<table class="grid"><thead><tr><th class="pilot">Pilot</th><th class="num">Time</th><th class="num">Behind</th><th class="pos">#</th></tr></thead><tbody>';
                let pos = 1;
                let lastLap = null;
                for (const lap of laps) {
                    const pid = lap.detectionObject.Pilot;
                    const p = pilots[pid];
                    if (!p) continue;
                    const behind = lastLap ? (Date.parse(lap.EndTime) - Date.parse(lastLap.EndTime)) / 1000 : 0;
                    html += "<tr>";
                    html += `<td class="pilot">${pilotLink(p.pilot)}</td>`;
                    html += `<td class="num">${fmtTime(lap.LengthSeconds)}</td>`;
                    html += `<td class="num">${pos === 1 ? "" : "+" + fmtTime(behind)}</td>`;
                    html += `<td class="pos">${pos}</td>`;
                    html += "</tr>";
                    pos++;
                    lastLap = lap;
                }
                html += "</tbody></table></div>";
                lapsWrap.appendChild(el("div", null, html));
            }
            wrap.appendChild(lapsWrap);
        } else {
            wrap.appendChild(el("div", "empty", "No lap records."));
        }

        return { node: wrap, name: raceName };
    }

    // ---- オーバーレイ + ルーティング ----------------------------------
    let overlayEl = null;

    function ensureOverlay() {
        if (overlayEl) return overlayEl;
        overlayEl = el("div", "overlay");
        overlayEl.innerHTML =
            `<div class="overlay-bar">
                <button type="button" class="back-btn" id="pdBack">‹ Back</button>
                <span class="overlay-title" id="pdTitle"></span>
             </div>
             <div class="overlay-body" id="pdBody"></div>`;
        document.body.appendChild(overlayEl);
        overlayEl.querySelector("#pdBack").addEventListener("click", () => history.back());
        return overlayEl;
    }

    function closeOverlay() {
        if (overlayEl) overlayEl.classList.remove("open");
        document.body.classList.remove("no-scroll");
        // 詳細表示中に更新があった場合は、閉じたタイミングでリロード
        if (pendingReload) {
            pendingReload = false;
            location.reload();
        }
    }

    // パイロット詳細・レース詳細で共通のオーバーレイ描画
    async function openDetail(loader) {
        const ov = ensureOverlay();
        ov.classList.add("open");
        document.body.classList.add("no-scroll");
        const body = ov.querySelector("#pdBody");
        const title = ov.querySelector("#pdTitle");
        title.textContent = "";
        body.innerHTML = '<div class="loading">Loading…</div>';
        try {
            const { node, name } = await loader();
            title.textContent = name;
            body.innerHTML = "";
            body.appendChild(node);
            body.scrollTop = 0;
        } catch (e) {
            body.innerHTML = '<div class="empty">Error: ' + esc(e && e.message ? e.message : e) + "</div>";
        }
    }

    function showPilot(pilotId) { return openDetail(() => buildPilotDetail(pilotId)); }
    function showRace(raceId) { return openDetail(() => buildRaceDetail(raceId)); }

    function pilotFromHash() {
        const m = /^#pilot=(.+)$/.exec(location.hash);
        return m ? decodeURIComponent(m[1]) : null;
    }
    function raceFromHash() {
        const m = /^#race=(.+)$/.exec(location.hash);
        return m ? decodeURIComponent(m[1]) : null;
    }

    function syncFromHash() {
        const pid = pilotFromHash();
        const rid = raceFromHash();
        if (pid) showPilot(pid);
        else if (rid) showRace(rid);
        else closeOverlay();
    }

    function initPilotRouter() {
        // パイロット名リンク(<a href="#pilot=...">)は hash を変えるだけ。
        // hashchange を購読して表示/非表示を切り替える (端末の戻る操作にも対応)。
        window.addEventListener("hashchange", syncFromHash);
        // ディープリンク対応 (URL に #pilot=... 付きで開かれた場合)
        syncFromHash();
    }

    // ---- オートリロード -----------------------------------------------
    // version.json (generatedAt) を定期取得し、再生成を検知したらリロード。
    // パイロット詳細を開いている間は閉じるまで保留する。
    async function checkVersion() {
        try {
            const r = await fetch(VERSION_URL, { cache: "no-store" });
            if (!r.ok) return;
            const v = await r.json();
            if (!v || !v.generatedAt) return;
            if (LOADED_VERSION && v.generatedAt !== LOADED_VERSION) {
                if (overlayEl && overlayEl.classList.contains("open")) {
                    pendingReload = true;
                } else {
                    location.reload();
                }
            }
        } catch (e) { /* ネットワーク一時失敗は無視 */ }
    }

    function startAutoReload() {
        setInterval(checkVersion, AUTO_RELOAD_MS);
    }

    function wireToolbar() {
        const setAll = (open) => {
            document.querySelectorAll("details.round").forEach((d) => { d.open = open; });
        };
        document.getElementById("expandAll").addEventListener("click", () => setAll(true));
        document.getElementById("collapseAll").addEventListener("click", () => setAll(false));
    }

    main().catch((e) => {
        document.getElementById("app").innerHTML =
            '<div class="empty">Error: ' + esc(e && e.message ? e.message : e) + "</div>";
    });
})();
