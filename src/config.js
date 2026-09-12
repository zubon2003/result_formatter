const fs = require('fs');
const path = require('path');

const configPath = path.join(__dirname, '..', 'config.json');

// 全項目の既定値（フォールバック）。config.json が無い/項目が欠けていても、
// ここで補完される。設定UIで保存すると config.json が生成・更新される。
const DEFAULTS = {
    fpvtrackside_dir_path: '',
    google_spreadsheet_id: '',
    web_ui_port: 8087,
    selected_event_id: 'all',
    leaderboard_round: 'all',
    // Minimum Lap Ranking シートに書き出す最大件数 (0 = 無制限)
    min_lap_ranking_limit: 0
};

// FPVTrackside の標準データ場所 (Windows: %LOCALAPPDATA%\FPVTrackside)
function defaultFpvDir() {
    const la = process.env.LOCALAPPDATA;
    return la ? path.join(la, 'FPVTrackside').replace(/\\/g, '/') : '';
}

function withDefaults(parsed) {
    const merged = {
        ...DEFAULTS,
        ...parsed
    };
    // FPVTrackside ディレクトリが空欄なら標準の保存先に自動補完
    if (!merged.fpvtrackside_dir_path || !String(merged.fpvtrackside_dir_path).trim()) {
        merged.fpvtrackside_dir_path = defaultFpvDir();
    }
    return merged;
}

let warnedNoConfig = false;

function loadConfig() {
    try {
        if (fs.existsSync(configPath)) {
            return withDefaults(JSON.parse(fs.readFileSync(configPath, 'utf8')));
        }
        if (!warnedNoConfig) {
            console.warn('Warning: config.json not found. Using built-in defaults ' +
                '(configure via the settings UI and Save to create config.json).');
            warnedNoConfig = true;
        }
        return withDefaults({});
    } catch (error) {
        console.error('Error reading or parsing config.json:', error);
        return withDefaults({});
    }
}

const config = loadConfig();
const eventsDir = path.join(config.fpvtrackside_dir_path, 'events').replace(/\\/g, '/');
const credentialsPath = path.join(__dirname, '..', 'credentials.json');

module.exports = {
    loadConfig,
    config,
    eventsDir,
    credentialsPath,
    configPath
};
