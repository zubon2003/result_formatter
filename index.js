const fs = require('fs');
const path = require('path');
const { eventsDir } = require('./src/config');
const { processEvents } = require('./src/data-processor');
const { startServer } = require('./src/web-server');
const { startDisplayServer } = require('./src/display-server');

const DEBOUNCE_DELAY = 3000; // 3秒 (変更が静まってから再生成)
let debounceTimer;
let isProcessing = false; // 処理中フラグ
let pendingRun = false;   // 処理中に来た要求を保留

// メインの実行関数
async function run() {
    // 処理中に呼ばれたら破棄せず保留し、完了後にもう一度走らせる
    if (isProcessing) {
        pendingRun = true;
        console.log('Already processing. Queued a rerun for when it finishes.');
        return;
    }
    isProcessing = true;
    try {
        do {
            pendingRun = false;
            console.log('Starting to process events...');
            try {
                await processEvents();
                console.log('Processing finished successfully.');
            } catch (error) {
                console.error('An error occurred during processing:', error);
            }
            if (pendingRun) {
                console.log('Pending request detected. Rerunning...');
            }
        } while (pendingRun);
    } finally {
        isProcessing = false;
    }
}

// ファイル監視とデバウンス処理
function watchFiles() {
    // eventsDirが存在するか確認
    if (!fs.existsSync(eventsDir)) {
        console.error(`Error: Monitored directory not found: ${eventsDir}`);
        console.error('Please check the fpvtrackside_dir_path in your config.json.');
        // 終了する代わりに、ユーザーに修正を促すメッセージを表示
        return; 
    }

    console.log(`Watching for changes in ${eventsDir}...`);
    fs.watch(eventsDir, { recursive: true }, (eventType, filename) => {
        if (filename) {
            const triggerFiles = ['Event.json', 'Pilots.json', 'Rounds.json', 'Stages.json', 'Race.json', 'Result.json'];
            const isTriggerFile = triggerFiles.some(file => filename.endsWith(file));

            if (!isTriggerFile) {
                return;
            }

            console.log(`Detected ${eventType} in ${filename}. Debouncing...`);
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(() => {
                console.log('Debounce timer elapsed. Triggering run.');
                run();
            }, DEBOUNCE_DELAY);
        }
    });
}

// アプリケーションの開始
function main() {
    // 設定UI(別ポート)を起動し、設定変更時のコールバックとして run を渡す
    startServer(run);

    // 表示用 web (結果ビュー) を別ポートで配信
    startDisplayServer();

    // 初回実行
    run();

    // ファイル監視を開始
    watchFiles();
}

main();
