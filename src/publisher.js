/*
 * publisher: webdist を S3 互換ストレージ(Cloudflare R2 など)へアップロードして外部公開する。
 *
 *  - data/bundle.json と data/version.json は毎回アップロード (no-cache)
 *  - 静的アセット(index.html / app.js / style.css / httpfiles/*) は内容ハッシュが
 *    変わった時だけアップロード (short cache)
 *
 * 設定:
 *   config.json の "publish" ブロック (機密でない値):
 *     {
 *       "enabled": true,
 *       "provider": "r2",
 *       "endpoint": "https://<accountid>.r2.cloudflarestorage.com",
 *       "bucket": "fpv-results",
 *       "region": "auto",
 *       "prefix": ""
 *     }
 *   アクセスキー(機密)は環境変数で渡す:
 *     R2_ACCESS_KEY_ID, R2_SECRET_ACCESS_KEY
 *
 * 依存なし: src/s3.js (Node 標準の https/crypto だけで SigV4 署名してアップロード)。
 */
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { loadConfig } = require('./config');
const { putObject, deleteObject } = require('./s3');

const CONTENT_TYPES = {
    '.html': 'text/html; charset=utf-8',
    '.js': 'text/javascript; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.svg': 'image/svg+xml',
    '.ico': 'image/x-icon'
};

// 常に最新を取りに行かせたいファイル (auto-reload の要)
const NO_CACHE = 'no-cache, max-age=0, must-revalidate';
// 静的アセットは短めにキャッシュ
const SHORT_CACHE = 'public, max-age=300';

function sha256(buf) {
    return crypto.createHash('sha256').update(buf).digest('hex');
}

function listFiles(dir, base = dir) {
    const out = [];
    for (const name of fs.readdirSync(dir)) {
        const full = path.join(dir, name);
        const st = fs.statSync(full);
        if (st.isDirectory()) out.push(...listFiles(full, base));
        else out.push(path.relative(base, full).split(path.sep).join('/'));
    }
    return out;
}

function loadState(statePath) {
    try { return JSON.parse(fs.readFileSync(statePath, 'utf8')); }
    catch (e) { return {}; }
}

/**
 * webdist を S3 互換ストレージへ公開する。設定/認証が無ければ何もしない。
 */
async function publishWeb(webDistDir) {
    const cfg = (loadConfig().publish) || {};
    if (!cfg.enabled) return; // 既定では無効 (Sheets と同じくオプトイン)

    const accessKeyId = process.env.R2_ACCESS_KEY_ID || cfg.accessKeyId;
    const secretAccessKey = process.env.R2_SECRET_ACCESS_KEY || cfg.secretAccessKey;
    if (!cfg.endpoint || !cfg.bucket || !accessKeyId || !secretAccessKey) {
        console.warn('publish: endpoint/bucket/credentials not set; skipping.' +
            ' (check the "publish" block in config.json and env vars R2_ACCESS_KEY_ID / R2_SECRET_ACCESS_KEY)');
        return;
    }

    // 依存ゼロの S3 クライアント (src/s3.js, SigV4 自前実装) を使う
    const opt = {
        endpoint: cfg.endpoint,
        bucket: cfg.bucket,
        region: cfg.region || 'auto',
        accessKeyId,
        secretAccessKey
    };

    const prefix = cfg.prefix || '';
    const statePath = path.join(webDistDir, '.publish-state.json');

    // アップロード済み記録は「宛先(endpoint/bucket/prefix)」ごとに有効。
    // バケット等を変えた場合、旧宛先の記録で再送がスキップされ、新バケットが
    // 空のまま(=公開URLが404)になるのを防ぐため、宛先が変わったら全ファイルを
    // 送り直す。新フォーマット: { __dest, files }。旧フラット形式も全送扱い。
    const destId = `${cfg.endpoint}|${cfg.bucket}|${prefix}`;
    const savedState = loadState(statePath);
    const prevState = (savedState && savedState.__dest === destId && savedState.files)
        ? savedState.files
        : {};
    if (savedState && savedState.__dest && savedState.__dest !== destId) {
        console.log(`publish: destination changed (${savedState.__dest} -> ${destId}); re-uploading all files.`);
    }

    // Web から明示的に隠すイベント (ローカルには残し、R2 からは消す/出さない)
    const hiddenIds = new Set((loadConfig().web_unpublished_event_ids) || []);
    let localEvents = [];
    try { localEvents = JSON.parse(fs.readFileSync(path.join(webDistDir, 'events.json'), 'utf8')); }
    catch (e) { localEvents = []; }
    const hiddenSlugs = new Set(
        localEvents.filter(e => e && hiddenIds.has(e.eventId)).map(e => e.slug)
    );

    // 「R2 にあるべき状態」(desired) を計算する。
    // 公開側は「他イベントを参照させない」ため、各イベントを単独表示にする:
    //   - 隠すイベントのフォルダは除外
    //   - 一覧(landing) ルート index.html と events.json は公開しない
    //   - 各イベントページから一覧へ戻る「≡」リンクを除去
    const desired = {}; // rel -> { body, hash, noCache }
    for (const rel of listFiles(webDistDir).filter((f) => f !== '.publish-state.json')) {
        if (rel === 'index.html' || rel === 'events.json') continue; // landing と一覧は公開しない
        if (hiddenSlugs.has(rel.split('/')[0])) continue; // 隠すイベントは出さない
        let body = fs.readFileSync(path.join(webDistDir, rel));
        if (rel.endsWith('/index.html')) {
            // 公開ページからは「≡」(一覧へ戻る) リンクを除去
            body = Buffer.from(String(body).replace(/\s*<a class="home-link"[\s\S]*?<\/a>/, ''), 'utf8');
        }
        const noCache = /(^|\/)data\/(bundle|version)\.json$/.test(rel);
        desired[rel] = { body, hash: sha256(body), noCache };
    }

    const nextState = {};
    let uploaded = 0, skipped = 0, deleted = 0;

    // 変更分だけアップロード
    for (const rel of Object.keys(desired)) {
        const d = desired[rel];
        if (prevState[rel] === d.hash) { nextState[rel] = d.hash; skipped++; continue; }
        const ext = path.extname(rel).toLowerCase();
        try {
            await putObject(opt, prefix + rel, d.body, {
                contentType: CONTENT_TYPES[ext] || 'application/octet-stream',
                cacheControl: d.noCache ? NO_CACHE : SHORT_CACHE
            });
            nextState[rel] = d.hash; // アップロード成功時だけ記録 (失敗は次回再送される)
            uploaded++;
        } catch (e) {
            console.error(`publish: upload failed for ${rel}:`, e.message);
        }
    }

    // desired に無いもの(=ローカル削除 or Web から隠した)は R2 から削除
    for (const oldRel of Object.keys(prevState)) {
        if (nextState[oldRel] !== undefined) continue;
        try {
            await deleteObject(opt, prefix + oldRel);
            deleted++;
        } catch (e) {
            console.error(`publish: delete failed for ${oldRel}:`, e.message);
        }
    }

    try { fs.writeFileSync(statePath, JSON.stringify({ __dest: destId, files: nextState }, null, 2), 'utf8'); }
    catch (e) { /* ignore */ }

    console.log(`publish: uploaded ${uploaded}, skipped ${skipped}, deleted ${deleted}` +
        (hiddenSlugs.size ? ` (web-hidden: ${[...hiddenSlugs].join(', ')})` : '') + ` -> ${cfg.bucket}`);

    // 公開URL(閲覧用)をログに表示
    const base = String(cfg.public_base_url || '').replace(/\/+$/, '');
    const visible = localEvents.filter(e => e && !hiddenSlugs.has(e.slug));
    if (base && visible.length) {
        console.log('Public URLs:');
        for (const e of visible) {
            console.log(`  ${e.name || e.slug}: ${base}/${e.slug}/index.html`);
        }
    } else if (!base && visible.length) {
        console.log('publish: public_base_url が未設定です（設定UIの「公開URL（閲覧用）」に pub-xxxx.r2.dev を入れると URL を表示します）。');
    }
}

module.exports = { publishWeb };
