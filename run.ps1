# ─────────────────────────────────────────────────────────────
#  상담·행정 채용공고 — 수집하고 깃허브에 올리기
#
#  복지넷(시·군·구 청소년상담복지센터 공고)은 해외 IP 에서 열리지 않아
#  깃허브 서버에서는 못 받는다. 그래서 이 컴퓨터에서 주기적으로 돌린다.
#  윈도 작업 스케줄러가 월·목 오후 2시에 이 파일을 실행한다.
#
#  손으로 돌리려면 '수집하고올리기.bat' 을 두 번 누르면 된다.
#
#  ※ .bat 이 아니라 .ps1 인 이유 — cmd 는 배치 파일을 시스템 코드페이지(cp949)로
#    읽어서, UTF-8 로 적은 한글이 깨지고 명령이 엉킨다. PowerShell 은 안 그런다.
#
#  ※ 겹침(충돌)을 푸는 대신 아예 안 나게 한다.
#    깃허브 봇도 docs/·data/ 를 고치므로, 먼저 받아서 맞춰 놓고(rebase) 그다음에
#    프로그램을 돌려 결과를 새로 만든다. 그러면 겹칠 것이 없다.
#    그 사이에 봇이 끼어드는 드문 경우만 -X theirs 로 이번 실행분을 택한다.
# ─────────────────────────────────────────────────────────────
param([switch]$Auto)          # -Auto : 작업 스케줄러가 부를 때(끝나고 안 멈춤)

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $here
$기록파일 = Join-Path $here "data\작업기록.txt"

function 남기기($t) {
    $줄 = "{0}  {1}" -f (Get-Date -Format "HH:mm:ss"), $t
    Write-Output $줄
    $줄 | Out-File -FilePath $기록파일 -Append -Encoding utf8
}

function 끝내기($코드) {
    if ($코드 -eq 0) { 남기기 "끝 (정상)" } else { 남기기 "끝 (실패)" }
    if (-not $Auto) { Read-Host "엔터를 누르면 닫힙니다" }
    exit $코드
}

# 기록은 덮어쓴다 — 쌓아 두면 파일만 커진다
"===== 상담·행정 채용공고 수집 : {0} =====" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss") |
    Out-File -FilePath $기록파일 -Encoding utf8

# ── 1. 깃허브에 올라온 것을 먼저 받아 맞춘다 ─────────────────
남기기 "깃허브에서 최신 내용을 받는 중"
git fetch -q origin
if ($LASTEXITCODE -ne 0) { 남기기 "[!] 인터넷이나 깃허브 로그인을 확인하세요."; 끝내기 1 }

git diff --quiet HEAD
if ($LASTEXITCODE -ne 0) {
    남기기 "  (지난번 결과물이 남아 있어 먼저 담아 둡니다)"
    git add -A docs data
    git commit -q -m ("수집 " + (Get-Date -Format "yyyy-MM-dd HH:mm"))
}
git -c rebase.autoStash=true rebase -X theirs origin/main
if ($LASTEXITCODE -ne 0) {
    남기기 "[!] 깃허브 내용과 맞추지 못했습니다. 사람이 봐야 합니다."
    git rebase --abort
    끝내기 1
}

# ── 2. 공고를 모은다 ───────────────────────────────────────
남기기 "공고 모으는 중 (복지넷까지 훑느라 3분쯤 걸립니다)"
python collect.py | Out-Null
if ($LASTEXITCODE -ne 0) {
    남기기 "[!] 수집 중 오류가 났습니다. data\log.txt 를 보세요."
    끝내기 1
}
$끝줄 = Get-Content (Join-Path $here "data\log.txt") -Tail 1 -Encoding UTF8 -ErrorAction SilentlyContinue
if ($끝줄) { 남기기 ("  " + $끝줄) }

# ── 3. 올린다 ─────────────────────────────────────────────
git add -A docs data
git diff --staged --quiet
if ($LASTEXITCODE -eq 0) { 남기기 "바뀐 내용이 없습니다."; 끝내기 0 }

git commit -q -m ("수집 " + (Get-Date -Format "yyyy-MM-dd HH:mm"))
git push -q origin main
if ($LASTEXITCODE -ne 0) {
    # 모으는 사이에 봇이 끼어든 드문 경우. 한 번만 다시 맞추고 올린다.
    남기기 "  (그 사이 깃허브가 바뀌어 한 번 더 맞춥니다)"
    git fetch -q origin
    git -c rebase.autoStash=true rebase -X theirs origin/main
    if ($LASTEXITCODE -ne 0) { git rebase --abort; 남기기 "[!] 맞추지 못했습니다."; 끝내기 1 }
    git push -q origin main
    if ($LASTEXITCODE -ne 0) { 남기기 "[!] 올리지 못했습니다."; 끝내기 1 }
}

남기기 "올렸습니다 → https://hongyul67-cpu.github.io/sangdam-chaeyong/"
남기기 "  (깃허브가 화면을 다시 만드는 데 1~2분 걸립니다)"
끝내기 0
