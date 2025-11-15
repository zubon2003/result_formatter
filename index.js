const fs = require('fs');
const path = require('path');
const { eventsDir } = require('./src/config');
const { processEvents } = require('./src/data-processor');
const { startServer, updateCache } = require('./src/web-server');

const DEBOUNCE_DELAY = 5000; // 5秒
let debounceTimer;
let isProcessing = false; // 処理中フラグ

// メインの実行関数
async function run() {
    if (isProcessing) {
        console.log('Already processing. Skipping new run.');
        return;
    }
    isProcessing = true;
    console.log('Starting to process events...');
    try {
        // processEventsにWebキャッシュの更新を任せる
        await processEvents(updateCache);
        console.log('Processing finished successfully.');
    } catch (error) {
        console.error('An error occurred during processing:', error);
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
            const triggerFiles = ['Event.json', 'Pilots.json', 'Rounds.json', 'Race.json', 'Result.json'];
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
    // Webサーバーを起動し、設定変更時のコールバックとして run を渡す
    startServer(run);
    
    // 初回実行
    run();

    // ファイル監視を開始
    watchFiles();
}

main();
