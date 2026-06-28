const fs = require('fs');
const path = require('path');

const configPath = path.join(__dirname, '..', 'config.json');

// 全項目の既定値（フォールバック）。config.json が無い/項目が欠けていても、
// ここで補完される。設定UIで保存すると config.json が生成・更新される。
const DEFAULTS = {
    fpvtrackside_dir_path: '',
    google_spreadsheet_id: '',
    web_ui_port: 8087,
    display_web_port: 8089,
    selected_event_id: 'all',
    published_event_ids: [],
    web_unpublished_event_ids: [],
    leaderboard_round: 'all',
    publish: {
        enabled: false,
        provider: 'r2',
        endpoint: '',
        bucket: '',
        region: 'auto',
        prefix: '',
        public_base_url: ''
    }
};

function withDefaults(parsed) {
    return {
        ...DEFAULTS,
        ...parsed,
        // publish は項目欠落を防ぐため個別にもマージ
        publish: { ...DEFAULTS.publish, ...((parsed && parsed.publish) || {}) }
    };
}

function loadConfig() {
    try {
        if (fs.existsSync(configPath)) {
            return withDefaults(JSON.parse(fs.readFileSync(configPath, 'utf8')));
        }
        console.warn('Warning: config.json not found. Using built-in defaults ' +
            '(configure via the settings UI and Save to create config.json).');
        return withDefaults({});
    } catch (error) {
        console.error('Error reading or parsing config.json:', error);
        return withDefaults({});
    }
}

const config = loadConfig();
const eventsDir = path.join(config.fpvtrackside_dir_path, 'events').replace(/\\/g, '/');
const credentialsPath = path.join(__dirname, '..', 'credentials.json');

// 表示用 web (結果ビュー) のソース/出力ディレクトリと、設定UIとは別のポート
const webSrcDir = path.join(__dirname, '..', 'web');
const webDistDir = path.join(__dirname, '..', 'webdist');
const displayWebPort = config.display_web_port || 8089;

module.exports = {
    loadConfig,
    config,
    eventsDir,
    credentialsPath,
    configPath,
    webSrcDir,
    webDistDir,
    displayWebPort
};
