const fs = require('fs');
const path = require('path');

const configPath = path.join(__dirname, '..', 'config.json');

function loadConfig() {
    try {
        if (fs.existsSync(configPath)) {
            return JSON.parse(fs.readFileSync(configPath, 'utf8'));
        }
        // config.json がない場合、デフォルト設定を返すか、エラーを投げる
        // ここでは空のオブジェクトを返す例
        console.warn('Warning: config.json not found. Using default values.');
        return {
            fpvtrackside_dir_path: '.',
            selected_event_id: 'all',
            leaderboard_round: 'all',
            sorted_by: 'bestLap',
            google_spreadsheet_id: '',
            web_ui_port: 3000
        };
    } catch (error) {
        console.error('Error reading or parsing config.json:', error);
        // エラー発生時もデフォルト値を返す
        return {
            fpvtrackside_dir_path: '.',
            selected_event_id: 'all',
            leaderboard_round: 'all',
            sorted_by: 'bestLap',
            google_spreadsheet_id: '',
            web_ui_port: 3000
        };
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
