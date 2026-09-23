<#
  FPVTrackside の events/ 配下にある JSON の破損チェック。

  「サイズは正しいのに中身が全部 NUL (0x00)」という壊れ方を主に探す。これは
  NTFS がファイルサイズとクラスタ割り当てだけを確定し、実データがディスクに
  届かないまま OS / ドライブが止まったときに残る跡で、電源断・BSOD・USB の
  突然の取り外し・SSD のキャッシュ喪失などで発生する。アプリが単体で落ちた
  だけでは起きない。

  agent も result_formatter も、この1ファイルで JSON.parse に失敗してイベント
  全体を捨ててしまうため、どのファイルが壊れているかを先に特定する。

  使い方:
    check-events.bat                  診断のみ
    check-events.bat -Fix             壊れたファイルを .corrupt-<日時> にリネーム
    check-events.bat -EventId <GUID>  そのイベントだけ調べる
    check-events.bat -EventsDir <path>
#>
[CmdletBinding()]
param(
    [string]$EventsDir,
    [string]$EventId,
    [switch]$Fix,
    [switch]$NoPrompt,
    [switch]$SkipSystem
)

$ErrorActionPreference = 'Stop'

# ---- 出力先ログ -------------------------------------------------------------
$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$desktop = [Environment]::GetFolderPath('Desktop')
if (-not $desktop -or -not (Test-Path $desktop)) { $desktop = $env:TEMP }
$logPath = Join-Path $desktop "fpvtrackside-check-$stamp.txt"
try { Start-Transcript -Path $logPath -Force | Out-Null } catch { $logPath = $null }

function Write-Head($text) {
    Write-Host ''
    Write-Host ('=' * 74) -ForegroundColor DarkGray
    Write-Host " $text" -ForegroundColor Cyan
    Write-Host ('=' * 74) -ForegroundColor DarkGray
}

# ---- events ディレクトリの決定 ----------------------------------------------
# result_formatter / agent の config.json にある fpvtrackside_dir_path を優先し、
# 無ければ FPVTrackside の標準保存先を使う (src/config.js の defaultFpvDir と同じ)。
function Get-ConfiguredFpvDir {
    $candidates = @(
        (Join-Path $PSScriptRoot '..\config.json'),
        (Join-Path $PSScriptRoot 'config.json'),
        'C:\result_formatter-main\config.json',
        'C:\agent-main\config.json'
    )
    foreach ($c in $candidates) {
        if (-not (Test-Path -LiteralPath $c)) { continue }
        try {
            $j = Get-Content -LiteralPath $c -Raw -Encoding UTF8 | ConvertFrom-Json
            $p = $j.fpvtrackside_dir_path
            if ($p -and $p.ToString().Trim()) {
                Write-Host "設定ファイルから取得: $c" -ForegroundColor DarkGray
                return $p.ToString().Trim()
            }
        } catch { }
    }
    return $null
}

if (-not $EventsDir) {
    $fpv = Get-ConfiguredFpvDir
    if (-not $fpv) { $fpv = Join-Path $env:LOCALAPPDATA 'FPVTrackside' }
    $EventsDir = Join-Path $fpv 'events'
}
$EventsDir = $EventsDir -replace '/', '\'

Write-Head 'FPVTrackside イベントデータ 破損チェック'
Write-Host "対象     : $EventsDir"
Write-Host "実行日時 : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
if ($logPath) { Write-Host "ログ     : $logPath" }

if (-not (Test-Path -LiteralPath $EventsDir)) {
    Write-Host ''
    Write-Host "エラー: events フォルダが見つかりません: $EventsDir" -ForegroundColor Red
    Write-Host '  -EventsDir <パス> で明示的に指定してください。' -ForegroundColor Yellow
    try { Stop-Transcript | Out-Null } catch { }
    exit 1
}

$scanRoot = $EventsDir
if ($EventId) {
    $scanRoot = Join-Path $EventsDir $EventId
    if (-not (Test-Path -LiteralPath $scanRoot)) {
        Write-Host "エラー: イベントが見つかりません: $scanRoot" -ForegroundColor Red
        try { Stop-Transcript | Out-Null } catch { }
        exit 1
    }
}

# ---- 1ファイルの判定 ---------------------------------------------------------
# 戻り値 $null = 正常。文字列を返したらそれが破損の理由。
$DEEP_PARSE_LIMIT = 4MB

function Test-JsonFile {
    param([string]$Path, [long]$Length)

    if ($Length -eq 0) { return '0 バイト (中身なし)' }
    if ($Length -gt 100MB) { return $null }   # 巨大ファイルは読まない

    $bytes = [IO.File]::ReadAllBytes($Path)

    # 先頭から最初の非ゼロバイトを探す。見つからなければ全部 NUL。
    $firstNonZero = -1
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -ne 0) { $firstNonZero = $i; break }
    }
    if ($firstNonZero -lt 0) {
        return ('NUL 埋め: 全 {0:N0} バイトが 0x00。書き込みがディスクに届いていない' -f $bytes.Length)
    }
    if ([Array]::IndexOf($bytes, [byte]0) -ge 0) {
        return 'NUL バイト (0x00) を含む'
    }

    $text = [Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF).Trim()
    if ($text.Length -eq 0) { return '空白のみ' }

    $head = $text[0]
    $tail = $text[$text.Length - 1]
    if ($head -ne '[' -and $head -ne '{') {
        return ('先頭が JSON ではない (先頭バイト 0x{0:x2})' -f $bytes[0])
    }
    if ($tail -ne ']' -and $tail -ne '}') {
        return '末尾が欠けている (書き込み途中で切れた可能性)'
    }

    if ($Length -le $DEEP_PARSE_LIMIT) {
        try { $null = ConvertFrom-Json $text } catch {
            return ('JSON として解釈できない: {0}' -f $_.Exception.Message)
        }
    }
    return $null
}

# ---- イベント名の解決 (表示用) ----------------------------------------------
$eventNameCache = @{}
function Get-EventLabel {
    param([string]$FullPath)
    $rel = $FullPath.Substring($EventsDir.Length).TrimStart('\')
    $id  = ($rel -split '\\')[0]
    if (-not $eventNameCache.ContainsKey($id)) {
        $name = $id
        $ej = Join-Path (Join-Path $EventsDir $id) 'Event.json'
        if (Test-Path -LiteralPath $ej) {
            try {
                $e = Get-Content -LiteralPath $ej -Raw -Encoding UTF8 | ConvertFrom-Json
                if ($e[0].Name) { $name = $e[0].Name }
            } catch { }
        }
        $eventNameCache[$id] = $name
    }
    return $eventNameCache[$id]
}

# ---- 走査 -------------------------------------------------------------------
Write-Head 'JSON ファイルの走査'
$files = @(Get-ChildItem -LiteralPath $scanRoot -Recurse -File -Filter *.json -ErrorAction SilentlyContinue)
Write-Host ('対象ファイル数: {0:N0}' -f $files.Count)

$bad = New-Object System.Collections.ArrayList
$n = 0
foreach ($f in $files) {
    $n++
    if ($n % 200 -eq 0) {
        Write-Progress -Activity 'JSON をチェック中' -Status "$n / $($files.Count)" -PercentComplete (100 * $n / [Math]::Max(1, $files.Count))
    }
    $reason = $null
    try { $reason = Test-JsonFile -Path $f.FullName -Length $f.Length }
    catch { $reason = ('読み取りに失敗: {0}' -f $_.Exception.Message) }
    if ($reason) {
        $null = $bad.Add([pscustomobject]@{
            Path   = $f.FullName
            Rel    = $f.FullName.Substring($EventsDir.Length).TrimStart('\')
            Size   = $f.Length
            Event  = (Get-EventLabel $f.FullName)
            Reason = $reason
        })
    }
}
Write-Progress -Activity 'JSON をチェック中' -Completed

Write-Head '結果'
if ($bad.Count -eq 0) {
    Write-Host ('正常 {0:N0} 件 / 破損 0 件。JSON に問題はありません。' -f $files.Count) -ForegroundColor Green
} else {
    Write-Host ('正常 {0:N0} 件 / 破損 {1} 件' -f ($files.Count - $bad.Count), $bad.Count) -ForegroundColor Red
    foreach ($b in $bad) {
        Write-Host ''
        Write-Host ('  ファイル : {0}' -f $b.Rel) -ForegroundColor Yellow
        Write-Host ('  イベント : {0}' -f $b.Event)
        Write-Host ('  サイズ   : {0:N0} バイト' -f $b.Size)
        Write-Host ('  理由     : {0}' -f $b.Reason)
    }
    Write-Host ''
    Write-Host '※ agent / result_formatter は壊れたファイルを自動で読み飛ばします。' -ForegroundColor DarkGray
    Write-Host '   Race.json / Result.json の破損 -> そのレースだけ除外。他のレースは処理されます。' -ForegroundColor DarkGray
    Write-Host '   Event.json / Pilots.json / Rounds.json の破損 -> そのイベント全体が除外されます。' -ForegroundColor DarkGray
    Write-Host '   リネームは必須ではありません。毎回の警告を止めたい場合だけ実行してください。' -ForegroundColor DarkGray
}

# ---- リネーム ---------------------------------------------------------------
if ($bad.Count -gt 0) {
    $doFix = $Fix.IsPresent
    if (-not $doFix -and -not $NoPrompt) {
        Write-Host ''
        $ans = Read-Host '壊れたファイルを .corrupt-<日時> にリネームして退避しますか? (y/N)'
        $doFix = ($ans -eq 'y' -or $ans -eq 'Y')
    }
    if ($doFix) {
        Write-Head 'リネーム'
        foreach ($b in $bad) {
            $new = "$(Split-Path $b.Path -Leaf).corrupt-$stamp"
            try {
                Rename-Item -LiteralPath $b.Path -NewName $new -ErrorAction Stop
                Write-Host ('  OK   {0}  ->  {1}' -f $b.Rel, $new) -ForegroundColor Green
            } catch {
                Write-Host ('  NG   {0}  : {1}' -f $b.Rel, $_.Exception.Message) -ForegroundColor Red
            }
        }
        Write-Host ''
        Write-Host '元に戻したい場合は拡張子 .corrupt-<日時> を取り除いてください。' -ForegroundColor DarkGray
        Write-Host '(中身は失われているため、戻しても同じエラーになります)' -ForegroundColor DarkGray
    }
}

# ---- 原因の切り分け ----------------------------------------------------------
if (-not $SkipSystem) {
    Write-Head '原因の切り分け: ドライブ'
    try {
        $qualified = [IO.Path]::GetFullPath($EventsDir)
        $root = [IO.Path]::GetPathRoot($qualified)
        Write-Host "events のあるドライブ: $root"
        if ($root -match '^\\\\') {
            Write-Host '  => ネットワーク上のパスです。書き込み中の切断で同じ壊れ方をします。' -ForegroundColor Yellow
        } else {
            $letter = $root.Substring(0, 1)
            $v = Get-Volume -DriveLetter $letter
            Write-Host ('  種別       : {0}' -f $v.DriveType)
            Write-Host ('  状態       : {0} / {1}' -f $v.HealthStatus, $v.OperationalStatus)
            Write-Host ('  空き容量   : {0:N1} GB / {1:N1} GB' -f ($v.SizeRemaining / 1GB), ($v.Size / 1GB))
            if ($v.DriveType -ne 'Fixed') {
                Write-Host '  => 内蔵以外のドライブです。書き込み中の取り外し / 接触不良で同じ症状になります。' -ForegroundColor Yellow
            }
            if (($v.SizeRemaining / 1GB) -lt 5) {
                Write-Host '  => 空き容量が少ないです。書き込み失敗による破損の原因になり得ます。' -ForegroundColor Yellow
            }
        }
    } catch { Write-Host "  ドライブ情報を取得できませんでした: $($_.Exception.Message)" -ForegroundColor DarkGray }

    Write-Host ''
    Write-Host '物理ディスク:'
    try {
        Get-PhysicalDisk | Select-Object FriendlyName, BusType, MediaType, HealthStatus, OperationalStatus |
            Format-Table -AutoSize | Out-String | Write-Host
    } catch { Write-Host "  取得できませんでした: $($_.Exception.Message)" -ForegroundColor DarkGray }

    Write-Head '原因の切り分け: 予期しないシャットダウン (直近30日)'
    # 41   = Kernel-Power  正常な終了手順を踏まずに再起動した
    # 6008 = EventLog      前回のシャットダウンが予期しないものだった
    # 1001 = BugCheck      ブルースクリーン
    try {
        $ev = @(Get-WinEvent -FilterHashtable @{
            LogName   = 'System'
            Id        = 41, 1001, 6008
            StartTime = (Get-Date).AddDays(-30)
        } -ErrorAction Stop)
        if ($ev.Count -eq 0) {
            Write-Host '  該当なし。' -ForegroundColor Green
        } else {
            foreach ($e in $ev) {
                Write-Host ('  {0}  Id={1}  {2}' -f $e.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'), $e.Id, $e.ProviderName) -ForegroundColor Yellow
            }
            Write-Host ''
            Write-Host '  => この時刻付近で書き込み中だったファイルが NUL 埋めになります。' -ForegroundColor Yellow
        }
    } catch {
        Write-Host '  該当なし、またはイベントログを読めませんでした。' -ForegroundColor DarkGray
    }

    Write-Head '原因の切り分け: ディスク / NTFS / chkdsk のエラー (直近30日)'
    try {
        $ev2 = @(Get-WinEvent -FilterHashtable @{
            LogName   = 'System'
            Level     = 1, 2, 3
            StartTime = (Get-Date).AddDays(-30)
        } -ErrorAction Stop | Where-Object { $_.ProviderName -match 'disk|Ntfs|volmgr|storahci|stornvme|Chkdsk|Wininit' })
        if ($ev2.Count -eq 0) {
            Write-Host '  該当なし。' -ForegroundColor Green
        } else {
            # 同じエラーが数十件並ぶので、内容ごとにまとめて件数と発生期間だけ出す。
            $groups = $ev2 | Group-Object { '{0}|{1}|{2}' -f $_.ProviderName, $_.Id, (($_.Message -replace '\s+', ' ') -replace '0x[0-9a-fA-F]+|\d+', '#') }
            Write-Host ('  {0} 件 / {1} 種類' -f $ev2.Count, @($groups).Count)
            foreach ($g in ($groups | Sort-Object Count -Descending | Select-Object -First 15)) {
                $s = $g.Group[0]
                $msg = ($s.Message -replace '\s+', ' ')
                if ($msg.Length -gt 100) { $msg = $msg.Substring(0, 100) + '...' }
                $newest = ($g.Group | Measure-Object TimeCreated -Maximum).Maximum
                $oldest = ($g.Group | Measure-Object TimeCreated -Minimum).Minimum
                Write-Host ''
                Write-Host ('  [{0} 回] {1}(Id={2})' -f $g.Count, $s.ProviderName, $s.Id) -ForegroundColor Yellow
                Write-Host ('    期間: {0} 〜 {1}' -f $oldest.ToString('MM-dd HH:mm'), $newest.ToString('MM-dd HH:mm'))
                Write-Host ('    {0}' -f $msg) -ForegroundColor DarkGray
            }
            if (@($groups).Count -gt 15) { Write-Host ('  ... 他 {0} 種類' -f (@($groups).Count - 15)) -ForegroundColor DarkGray }
            Write-Host ''
            Write-Host '  => events フォルダのあるドライブに対するエラーかどうかを確認してください。' -ForegroundColor Yellow
            Write-Host '     別ドライブのエラーはこの破損とは無関係です。' -ForegroundColor DarkGray
        }
    } catch {
        Write-Host '  該当なし、またはイベントログを読めませんでした。' -ForegroundColor DarkGray
    }
}

Write-Head '完了'
if ($logPath) {
    Write-Host "ログを保存しました: $logPath"
    Write-Host '(このファイルをそのまま送れば、こちらで内容を確認できます)'
}
try { Stop-Transcript | Out-Null } catch { }
