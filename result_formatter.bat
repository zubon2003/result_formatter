@echo off
setlocal
rem Console to UTF-8 so the Japanese diagnostics from src/json-file.js
rem (broken JSON reports) render instead of turning into mojibake.
chcp 65001 >nul
node index.js
