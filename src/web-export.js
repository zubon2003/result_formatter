/*
 * web-export: processEvents() が「1 回の FS 読み込み」で得たデータから、
 * 表示用 web (結果ビュー) の publish フォルダ(webdist) を生成する。
 *
 * 複数イベント対応:
 *   webdist/
 *     index.html              ← イベント一覧(ランディング)
 *     app.js, style.css, httpfiles/   ← 共有アセット
 *     events.json             ← [{slug,name,eventId,version}]
 *     <slug>/
 *       index.html            ← イベント結果ページ(共通シェル / 相対 data を参照)
 *       data/bundle.json      ← 全データを 1 ファイルに集約
 *       data/version.json     ← {generatedAt:<contentHash>, eventId, eventName}
 *
 *  - bundle/version はブラウザが 1 回 fetch するだけ
 *  - version(generatedAt) は「実行時刻」ではなく「そのイベントのデータ内容ハッシュ」。
 *    内容が変わった時だけ version が変わる → 変わったイベントだけ再アップロード/リロード。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

// web/httpfiles/Channels.json を 1 度だけ読み込んでキャッシュ
let _channelsCache = null;
function loadChannels(webSrcDir) {
    if (_channelsCache) return _channelsCache;
    try {
        const p = path.join(webSrcDir, 'httpfiles', 'Channels.json');
        _channelsCache = JSON.parse(fs.readFileSync(p, 'utf8'));
    } catch (e) {
        console.warn('Could not read Channels.json; channel display may be empty:', e.message);
        _channelsCache = [];
    }
    return _channelsCache;
}

// イベント名から URL 用スラッグを作る
function slugify(name) {
    let s = String(name == null ? '' : name).trim().toLowerCase();
    s = s.replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '');
    return s || 'event';
}

/**
 * 既読データから bundle オブジェクトを構築する。
 * @param {object} ev { eventId, eventData, pilotsData, roundsData, stagesData, races:[{id, raceData, resultData}] }
 */
function buildBundle(ev, webSrcDir, decimalPlaces) {
    const dp = (decimalPlaces == null ? 2 : decimalPlaces);
    const files = {};
    files['Event.json'] = ev.eventData || [];
    files['Pilots.json'] = ev.pilotsData || [];
    files['Rounds.json'] = ev.roundsData || [];
    files['Stages.json'] = ev.stagesData || [];
    files['httpfiles/Channels.json'] = loadChannels(webSrcDir);

    for (const race of ev.races || []) {
        if (race.raceData) files[`${race.id}/Race.json`] = race.raceData;
        if (race.resultData) files[`${race.id}/Result.json`] = race.resultData;
    }

    const eventName = (ev.eventData && ev.eventData[0] && ev.eventData[0].Name) || '';
    // version はデータ内容のハッシュ (decimalPlaces も含めて、変化した時だけ変わる)
    const version = crypto.createHash('sha256')
        .update(JSON.stringify(files) + '|dp=' + dp).digest('hex').slice(0, 16);

    return {
        eventId: ev.eventId,
        eventName,
        generatedAt: version,
        decimalPlaces: dp,
        files
    };
}

function copyFileIfExists(src, dst) {
    if (fs.existsSync(src)) { fs.mkdirSync(path.dirname(dst), { recursive: true }); fs.copyFileSync(src, dst); }
}

// 共有アセット(app.js / style.css / httpfiles)を webdist ルートへ
function copySharedAssets(webSrcDir, webDistDir) {
    fs.mkdirSync(webDistDir, { recursive: true });
    copyFileIfExists(path.join(webSrcDir, 'app.js'), path.join(webDistDir, 'app.js'));
    copyFileIfExists(path.join(webSrcDir, 'style.css'), path.join(webDistDir, 'style.css'));
    const srcHttp = path.join(webSrcDir, 'httpfiles');
    const dstHttp = path.join(webDistDir, 'httpfiles');
    fs.mkdirSync(dstHttp, { recursive: true });
    if (fs.existsSync(srcHttp)) {
        for (const f of fs.readdirSync(srcHttp)) {
            const s = path.join(srcHttp, f);
            if (fs.statSync(s).isFile()) fs.copyFileSync(s, path.join(dstHttp, f));
        }
    }
}

// 現在の公開対象に無い古いイベントフォルダを掃除する
function cleanupOldEvents(webDistDir, keepSlugs) {
    const eventsJsonPath = path.join(webDistDir, 'events.json');
    let prev = [];
    try { prev = JSON.parse(fs.readFileSync(eventsJsonPath, 'utf8')); } catch (e) { prev = []; }
    for (const e of prev) {
        if (e && e.slug && !keepSlugs.has(e.slug)) {
            const dir = path.join(webDistDir, e.slug);
            try { fs.rmSync(dir, { recursive: true, force: true }); } catch (ex) { /* ignore */ }
        }
    }
}

/**
 * 複数イベントを webdist へ書き出す。
 * 各 spec は { slug, name, eventId, ev? }:
 *   - ev あり  → そのイベントを再生成 (アクティブ/初回)
 *   - ev なし  → 既存の出力を保管 (再生成しない。一覧には載せる)
 * @param {Array} specs
 */
function exportWebMulti(specs, webSrcDir, webDistDir, decimalPlaces) {
    copySharedAssets(webSrcDir, webDistDir);

    const shellHtml = fs.readFileSync(path.join(webSrcDir, 'index.html'), 'utf8');
    const landingSrc = path.join(webSrcDir, 'landing.html');
    const landingHtml = fs.existsSync(landingSrc) ? fs.readFileSync(landingSrc, 'utf8') : null;

    const keepSlugs = new Set();
    const index = [];
    let regenerated = 0;

    for (const spec of specs) {
        const slug = spec.slug;
        keepSlugs.add(slug);
        const eventDir = path.join(webDistDir, slug);
        const dataDir = path.join(eventDir, 'data');
        fs.mkdirSync(dataDir, { recursive: true });

        // シェルは毎回 BASE を埋めて更新 (共有アセット変更にも追従させる)
        const eventHtml = shellHtml.replace(/%BASE%/g, '/' + slug + '/');
        fs.writeFileSync(path.join(eventDir, 'index.html'), eventHtml, 'utf8');

        if (spec.ev) {
            // 再生成
            const bundle = buildBundle(spec.ev, webSrcDir, decimalPlaces);
            const races = Object.keys(bundle.files).filter(k => k.endsWith('/Race.json')).length;
            fs.writeFileSync(path.join(dataDir, 'bundle.json'), JSON.stringify(bundle), 'utf8');
            fs.writeFileSync(path.join(dataDir, 'version.json'), JSON.stringify({
                generatedAt: bundle.generatedAt,
                eventId: bundle.eventId,
                eventName: bundle.eventName,
                races
            }), 'utf8');
            index.push({ slug, name: bundle.eventName, eventId: bundle.eventId, version: bundle.generatedAt, races });
            regenerated++;
        } else {
            // 保管: 既存の version.json から一覧情報だけ拾う (bundle はそのまま)
            let v = {};
            try { v = JSON.parse(fs.readFileSync(path.join(dataDir, 'version.json'), 'utf8')); } catch (e) { v = {}; }
            index.push({
                slug,
                name: v.eventName || spec.name || slug,
                eventId: spec.eventId || v.eventId,
                version: v.generatedAt || '',
                races: v.races || 0
            });
        }
    }

    // published に無い古いイベントだけ掃除 (保管中のものは消さない)
    cleanupOldEvents(webDistDir, keepSlugs);
    if (landingHtml) fs.writeFileSync(path.join(webDistDir, 'index.html'), landingHtml, 'utf8');
    fs.writeFileSync(path.join(webDistDir, 'events.json'), JSON.stringify(index, null, 2), 'utf8');

    console.log(`web-export: ${index.length} event(s) [regenerated ${regenerated}, kept ${index.length - regenerated}]: ${index.map(e => e.slug).join(', ')}`);
    return index;
}

module.exports = { buildBundle, exportWebMulti, slugify };
