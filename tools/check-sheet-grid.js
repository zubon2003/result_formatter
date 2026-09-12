// 書き出し先スプレッドシートの各シートの行数/列数を表示する診断スクリプト (読み取りのみ)
//   node tools/check-sheet-grid.js [spreadsheetId]
// 引数省略時は config.json の google_spreadsheet_id を使用。
const { google } = require('googleapis');
const { loadConfig, credentialsPath } = require('../src/config');

(async () => {
    const spreadsheetId = process.argv[2] || loadConfig().google_spreadsheet_id;
    if (!spreadsheetId) {
        console.error('スプレッドシートIDが未設定です。引数で渡すか config.json に設定してください。');
        process.exit(1);
    }

    const auth = new google.auth.GoogleAuth({
        keyFile: credentialsPath,
        scopes: ['https://www.googleapis.com/auth/spreadsheets.readonly'],
    });
    const sheets = google.sheets({ version: 'v4', auth });

    try {
        const res = await sheets.spreadsheets.get({ spreadsheetId });
        console.log(`Spreadsheet: ${res.data.properties.title} (${spreadsheetId})`);
        for (const s of res.data.sheets) {
            const g = s.properties.gridProperties || {};
            console.log(`  ${String(s.properties.title).padEnd(24)} rows=${g.rowCount} cols=${g.columnCount}`);
        }
    } catch (err) {
        console.error('取得に失敗しました:', (err && err.message) || err);
        if (err && err.errors) console.error(JSON.stringify(err.errors, null, 2));
        process.exit(1);
    }
})();
