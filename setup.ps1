# Sidenote のセットアップ（Windows）。setup.bat をダブルクリックすると、これが動きます。
# 何度動かしても大丈夫です（入っているものは飛ばします）。
#   1. Python      … 無ければ winget で入れる
#   2. Python の部品 … requirements.txt
#   3. speech.cpp  … 公式のインストーラーで入れ（古い版は新しくし）、使うモデル4つを先にダウンロードする
#   4. ffmpeg      … 無ければ winget で入れる
#   5. Claude Code … 無ければ、入れるかどうか聞く。ログインは本人が行う
#   6. ショートカット … スタートメニューとスタートアップ
$ErrorActionPreference = "Stop"
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -LiteralPath $Here
$Models = @("haas-ja-v1.0:q4_k", "silero-vad", "nemotron-3-diarization", "re-use")  # meeting_notes.py が使うモデル
$Speech = Join-Path $env:LOCALAPPDATA "Programs\speech.cpp\current\speech.exe"
$SpeechVersion = "0.10.1"  # 動作を確かめた版

function Step($n, $text) { Write-Host ""; Write-Host "──── [$n/6] $text ────" -ForegroundColor Cyan }

Write-Host ""
Write-Host "Sidenote のセットアップを始めます。" -ForegroundColor Green
Write-Host "全体で 5〜15 分ほどかかります（ネットの速さによります）。英語の文字が流れますが、そのまま待ってください。"
Write-Host "途中で質問するのは、Claude Code を入れるか・ログインするか、の2回だけです。"
function Has($cmd) { [bool](Get-Command $cmd -ErrorAction SilentlyContinue) }
function Refresh-Path {
    # winget などで入れたものを、このウィンドウでもすぐ使えるようにする
    $env:Path = [Environment]::GetEnvironmentVariable("Path", "Machine") + ";" + [Environment]::GetEnvironmentVariable("Path", "User")
}
function Need-Winget {
    if (-not (Has "winget")) {
        throw "winget が見つかりません。Microsoft Store で「アプリ インストーラー」を入れてから、もう一度実行してください。"
    }
}
function Winget-Install($id) {
    winget install -e --id $id @args --accept-package-agreements --accept-source-agreements
    # すでに入っているときも 0 以外を返すので、ここでは止めない。入ったかどうかは、呼んだ側が確かめる
    if ($LASTEXITCODE -ne 0) { Write-Host "winget の終了コード: $LASTEXITCODE" }
}
function Install-Speech {
    # 版を固定して入れる。インストーラー自身が、落とした本体を配布元の SHA256SUMS で確かめる。
    # 別のプロセスで動かすので、この窓の設定には触らない
    $url = "https://raw.githubusercontent.com/nyosegawa/speech.cpp/v$SpeechVersion/install.ps1"
    $file = Join-Path $env:TEMP "speech-cpp-install-$SpeechVersion.ps1"
    Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $file
    & powershell -NoProfile -ExecutionPolicy Bypass -File $file -Version $SpeechVersion
    $code = $LASTEXITCODE
    Remove-Item -LiteralPath $file -Force -ErrorAction SilentlyContinue
    if ($code -ne 0 -or -not (Test-Path -LiteralPath $Speech)) {
        throw "speech.cpp を入れられませんでした。上のメッセージを確かめてください（グラフィックのドライバーに入っている Vulkan が必要です）。"
    }
}
function Find-Claude {
    # アプリと同じ場所を探す。PATH に入っていなくても、ここにあれば使える
    if (Has "claude") { return (Get-Command claude).Source }
    foreach ($p in @((Join-Path $env:USERPROFILE ".local\bin\claude.exe"), (Join-Path $env:APPDATA "npm\claude.cmd"))) {
        if (Test-Path -LiteralPath $p) { return $p }
    }
    return $null
}

# 1. Python（3.11 以上）
Step 1 "Python"
# Microsoft Store 版の Python は使わない。%APPDATA% への書き込みが別の場所へ振り替えられ、
# スタートメニューに登録されず、設定も通常の Python から見えなくなるため
function Find-Python {
    $paths = @()
    if (Has "py") {
        try {
            foreach ($line in (& py -0p 2>$null)) {
                if ("$line" -match '([A-Za-z]:\\.*\.exe)\s*$') { $paths += $Matches[1] }
            }
        } catch {}
    }
    if (Has "python") {
        # -X utf8 は、ユーザー名が日本語のときに場所が文字化けしないように
        try { $paths += (& python -X utf8 -c "import sys; print(sys.executable)" 2>$null) } catch {}
    }
    $paths += @(Get-ChildItem (Join-Path $env:LOCALAPPDATA "Programs\Python\Python3*\python.exe") -ErrorAction SilentlyContinue | ForEach-Object { $_.FullName })
    # 3.12 を優先する。新しすぎる版には、部品の完成品がまだ無いことがある
    $good = $null
    foreach ($p in $paths) {
        if (-not $p -or $p -match '\\WindowsApps\\') { continue }
        try {
            $v = & $p -c "import sys; print('%d.%d' % sys.version_info[:2] if sys.version_info >= (3, 11) else '')" 2>$null
            if ($v -eq "3.12") { return $p }
            if ($v -and -not $good) { $good = $p }
        } catch {}
    }
    return $good
}
$py = Find-Python
if (-not $py) {
    Need-Winget
    Write-Host "Python 3.12 を入れます"
    Winget-Install Python.Python.3.12 --scope user
    Refresh-Path
    $py = Find-Python
    if (-not $py) { throw "Python を入れましたが、まだ見つかりません。このウィンドウを閉じて、setup.bat をもう一度実行してください。" }
}
Write-Host "使う Python: $(& $py --version)"

# 2. Python の部品
Step 2 "Python の部品（1〜3分）"
# Sidenote は Python を場所で直接呼ぶので、Scripts が PATH に無くても困らない。不安にさせる警告を出さない
& $py -m pip install --upgrade pip --quiet --no-warn-script-location
& $py -m pip install -r requirements.txt --quiet --no-warn-script-location
if ($LASTEXITCODE -ne 0) { throw "Python の部品を入れられませんでした" }

# 3. speech.cpp とモデル
Step 3 "speech.cpp（文字起こし）とモデル（いちばん時間がかかります。数百MBをダウンロード）"
if (-not (Test-Path -LiteralPath $Speech)) {
    Write-Host "speech.cpp $SpeechVersion を公式のインストーラーで入れます（https://github.com/nyosegawa/speech.cpp）"
    Install-Speech
} else {
    # 雑音除去（RE-USE）は speech.cpp 0.10 以降にしか無い。無ければ新しい版を入れ直す
    $ErrorActionPreference = "Continue"
    & $Speech enhance --help *> $null
    $hasEnhance = ($LASTEXITCODE -eq 0)
    $ErrorActionPreference = "Stop"
    if ($hasEnhance) {
        Write-Host "speech.cpp は入っています: $Speech"
    } else {
        Write-Host "speech.cpp を $SpeechVersion にします（雑音除去に必要です）"
        Install-Speech
    }
}
foreach ($m in $Models) {
    Write-Host "モデル $m を用意します（初回だけダウンロードします）"
    & $Speech pull $m
    if ($LASTEXITCODE -ne 0) { throw "モデル $m をダウンロードできませんでした" }
}

# 4. ffmpeg（マイクの音を取り込む）
Step 4 "ffmpeg（マイクの音を取り込む）"
if (-not (Has "ffmpeg")) {
    Need-Winget
    Winget-Install Gyan.FFmpeg
    Refresh-Path
    $links = Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Links\ffmpeg.exe"
    $pkg = Get-ChildItem (Join-Path $env:LOCALAPPDATA "Microsoft\WinGet\Packages\Gyan.FFmpeg*\*\bin\ffmpeg.exe") -ErrorAction SilentlyContinue
    if (-not (Has "ffmpeg") -and -not (Test-Path -LiteralPath $links) -and -not $pkg) {
        throw "ffmpeg を入れられませんでした。上のメッセージを確かめて、setup.bat をもう一度実行してください。"
    }
} else {
    Write-Host "ffmpeg は入っています"
}

# 5. Claude Code（質問・補正に使う）
Step 5 "Claude Code（質問・補正に使う）"
$claude = Find-Claude
if (-not $claude) {
    Write-Host "API キーやローカル LLM で使う場合は、Claude Code は要りません（Sidenote の設定で接続先を選べます）。"
    $ans = Read-Host "Claude Code が見つかりません。公式のインストーラーで入れますか？ (y/n)"
    if ($ans -eq "y") {
        Invoke-RestMethod https://claude.ai/install.ps1 | Invoke-Expression
        Refresh-Path
        Set-Location -LiteralPath $Here
        $claude = Find-Claude
    }
}
if ($claude) {
    Write-Host "Claude Code は入っています。質問・補正は、ご自分のアカウントの利用枠を使います。"
    $ans = Read-Host "Claude Code にログインしますか？（もう済んでいれば n）(y/n)"
    if ($ans -eq "y") {
        # 別の窓で claude を開く。ブラウザが開くので、そこでログインする
        $msg = "ブラウザが開いたらログインしてください。ログインが終わったら、この窓は閉じてかまいません。"
        Start-Process powershell -ArgumentList "-NoExit", "-Command", "Write-Host '$msg' -ForegroundColor Green; & '$claude'"
        Write-Host "別の窓を開きました。そちらでログインしてください（セットアップはこのまま続きます）。"
    }
} else {
    Write-Host "Claude Code なしでも文字起こしはできます。質問・補正を使うときは、あとで入れてください。" -ForegroundColor Yellow
}

# 6. ショートカット
Step 6 "スタートメニューとスタートアップに登録"
if (-not (Test-Path -LiteralPath (Join-Path $Here "icon.ico"))) { & $py make_icons.py $Here }
& $py -X utf8 install_shortcuts.py
if ($LASTEXITCODE -ne 0) { throw "スタートメニューに登録できませんでした。上のメッセージを確かめてください。" }

Write-Host ""
Write-Host "セットアップが終わりました。スタートメニューの「Sidenote」から起動できます。" -ForegroundColor Green
