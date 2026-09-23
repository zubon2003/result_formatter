/*
 * events/ 配下の JSON を「壊れていたら読み飛ばせる」形で読む。
 *
 * FPVTrackside のデータには、サイズは正しいのに中身が全部 NUL (0x00) という
 * ファイルが混ざることがある。NTFS がファイルサイズとクラスタ割り当てだけを
 * 確定し、実データがディスクに届かないまま OS / ドライブが止まったときに残る
 * 跡で、電源断・BSOD・ドライブの突然の取り外し・SSD のキャッシュ喪失などで
 * 起きる。アプリが単体でクラッシュしただけでは起きない。
 *
 * 以前はこれを JSON.parse がそのまま投げていたため、
 *   - レース1件の破損でイベント全体 (呼び出し方によっては処理全体) が落ちる
 *   - メッセージが "Unexpected token '', \"\"... is not valid JSON" だけで、
 *     どのファイルが壊れているのか分からない
 * という二重の問題があった。ここで壊れ方を判定してファイル名付きで報告し、
 * 呼び出し側は null を見てその1件だけを飛ばす。
 */
const fs = require('fs');

// 同じファイルを何度も読む (変更検知のたびに再走査する) ので、詳細な説明は
// 内容が変わるまで1回だけ出す。2回目以降は1行に畳む。
const reported = new Map();

function fmt(n) {
    return n.toLocaleString('en-US');
}

// 壊れ方を人が読める理由に変換する。parseError は JSON.parse が投げた例外。
function diagnose(buf, parseError) {
    if (buf.length === 0) return '0 バイト (中身なし)';

    let firstNonZero = -1;
    for (let i = 0; i < buf.length; i++) {
        if (buf[i] !== 0) { firstNonZero = i; break; }
    }
    if (firstNonZero < 0) {
        return `全 ${fmt(buf.length)} バイトが NUL (0x00)。サイズだけが確定し、中身がディスクに書かれていない`;
    }
    if (buf.includes(0)) {
        return 'NUL (0x00) バイトを含む。書き込みの一部が失われている';
    }

    const text = buf.toString('utf8').replace(/^﻿/, '').trim();
    if (!text) return '空白のみ';
    if (text[0] !== '[' && text[0] !== '{') {
        return `先頭が JSON ではない (先頭バイト 0x${buf[0].toString(16).padStart(2, '0')})`;
    }
    if (text[text.length - 1] !== ']' && text[text.length - 1] !== '}') {
        return '末尾が欠けている (書き込みが途中で切れた)';
    }
    return parseError ? parseError.message : '不明';
}

function report(filePath, reason) {
    let info = '';
    try {
        const st = fs.statSync(filePath);
        info = `${fmt(st.size)} バイト / 最終更新 ${st.mtime.toLocaleString('ja-JP')}`;
    } catch (e) { /* 消えていても報告自体は続ける */ }

    const sig = `${reason}|${info}`;
    if (reported.get(filePath) === sig) {
        console.warn(`JSON 破損のためスキップ: ${filePath}`);
        return;
    }
    reported.set(filePath, sig);

    console.error('');
    console.error('---------------- JSON が壊れています ----------------');
    console.error(`  ファイル: ${filePath}`);
    if (info) console.error(`  情報    : ${info}`);
    console.error(`  原因    : ${reason}`);
    console.error('  説明    : 書き込み中に電源断・BSOD・ドライブの取り外し・SSD の');
    console.error('            キャッシュ喪失などが起きると、この状態になります。');
    console.error('            中身はディスク上に存在しないため復元できません。');
    console.error('  対処    : このファイルは読み飛ばして処理を続けます。この警告を');
    console.error('            出したくない場合はリネーム (例: Race.json.corrupt) して');
    console.error('            ください。check-events.bat でも一括で退避できます。');
    console.error('-----------------------------------------------------');
}

// 読めたらパース結果を返す。壊れていたら null を返し、理由をログに出す。
// 呼び出し側は null を「この1件は無かったことにして続行」として扱う。
function readJsonOrNull(filePath) {
    let buf;
    try {
        buf = fs.readFileSync(filePath);
    } catch (e) {
        report(filePath, `読み取りに失敗: ${e.message}`);
        return null;
    }
    try {
        return JSON.parse(buf.toString('utf8'));
    } catch (e) {
        report(filePath, diagnose(buf, e));
        return null;
    }
}

module.exports = { readJsonOrNull };
