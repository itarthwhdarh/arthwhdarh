# ══════════════════════════════════════════════════════════════
# سكريبت تشغيل وفحص نظام WhatsApp و OpenWA المحصن
# جمعية إرث وحضارة بالقريات
# ══════════════════════════════════════════════════════════════

$OutputEncoding = [System.Text.Encoding]::UTF8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$projectRoot = Split-Path -Parent $PSScriptRoot
Set-Location $projectRoot

Clear-Host
Write-Host "============================================================" -ForegroundColor Green
Write-Host "      جمعية إرث وحضارة بالقريات — تشغيل وفحص خدمة WhatsApp  " -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""

function Exit-WithError($msg) {
    Write-Host "`n[✕ خطأ فادح]: $msg" -ForegroundColor Red
    Write-Host "`nتوقف التشغيل: لن يتم إعلان النجاح حتى يتم حل المشكلة أعلاه." -ForegroundColor Yellow
    Write-Host "اضغط على أي مفتاح للإغلاق..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}

# ── 1. فحص تشغيل محرك Docker Desktop ──
Write-Host "[1/5] فحص تشغيل محرك Docker Desktop..." -ForegroundColor Cyan
$dockerOk = $false
try {
    $null = docker info 2>&1
    if ($LASTEXITCODE -eq 0) { $dockerOk = $true }
} catch {}

if (-not $dockerOk) {
    Write-Host "  > Docker غير مشتغل. جارٍ محاولة إطلاقه تلقائياً..." -ForegroundColor Yellow
    $dockerCandidates = @(
        "$env:ProgramFiles\Docker\Docker\Docker Desktop.exe",
        "$env:LOCALAPPDATA\Docker\Docker Desktop.exe"
    )
    $dockerStarted = $false
    foreach ($cand in $dockerCandidates) {
        if (Test-Path $cand) {
            Start-Process $cand
            $dockerStarted = $true
            break
        }
    }
    if ($dockerStarted) {
        Write-Host "  > تم إطلاق Docker Desktop. في انتظار اكتمال التهيئة (حتى 60 ثانية)..." -ForegroundColor Yellow
        $waitCount = 60
        while ($waitCount -gt 0) {
            Start-Sleep -Seconds 2
            try {
                $null = docker info 2>&1
                if ($LASTEXITCODE -eq 0) { $dockerOk = $true; break }
            } catch {}
            $waitCount -= 2
            Write-Host -NoNewline "."
        }
        Write-Host ""
    }
}

if (-not $dockerOk) {
    Exit-WithError "محرك Docker Desktop غير مشتغل. يرجى فتح برنامج Docker Desktop أولاً والتأكد من ظهور اللون الأخضر (Engine Running)."
}
Write-Host "  ✓ محرك Docker Desktop يعمل بنجاح." -ForegroundColor Green

# ── 2. فحص وتشغيل حاوية OpenWA ──
Write-Host "`n[2/5] فحص حاوية OpenWA (openwa-api)..." -ForegroundColor Cyan
$containerStatus = (docker inspect openwa-api --format '{{.State.Status}}' 2>$null)

if ($containerStatus -eq 'running') {
    Write-Host "  ✓ حاوية OpenWA (openwa-api) تعمل بالفعل." -ForegroundColor Green
} elseif ($containerStatus) {
    Write-Host "  > الحاوية موجودة ولكنها متوقفة. جارٍ تشغيلها..." -ForegroundColor Yellow
    docker start openwa-api | Out-Null
    Start-Sleep -Seconds 4
    $containerStatus = (docker inspect openwa-api --format '{{.State.Status}}' 2>$null)
    if ($containerStatus -ne 'running') {
        Exit-WithError "فشل تشغيل حاوية openwa-api. يرجى مراجعة سجلات Docker عبر: docker logs openwa-api"
    }
    Write-Host "  ✓ تم تشغيل حاوية OpenWA بنجاح." -ForegroundColor Green
} else {
    Exit-WithError "لم يتم العثور على الحاوية openwa-api في Docker. يرجى التأكد من إنشاء الحاوية مسبقاً."
}

# ── 3. فحص استجابة OpenWA API وجاهزية جلسة الواتساب ──
Write-Host "`n[3/5] فحص استجابة واجهة OpenWA وجاهزية الجلسة على المنفذ 2785..." -ForegroundColor Cyan

$apiReady = $false
$docsUrl = "http://127.0.0.1:2785/api/docs"

for ($i = 0; $i -lt 15; $i++) {
    try {
        $resp = Invoke-WebRequest -Uri $docsUrl -UseBasicParsing -TimeoutSec 3 -ErrorAction Stop
        if ($resp.StatusCode -eq 200) {
            $apiReady = $true
            break
        }
    } catch {
        # تجربة curl.exe كبديل موثوق في حال وجود قيود بـ PowerShell
        $curlCode = (curl.exe -s -o /dev/null -w "%{http_code}" --connect-timeout 2 "http://127.0.0.1:2785/api/docs" 2>$null)
        if ($curlCode -eq "200") {
            $apiReady = $true
            break
        }
    }
    Start-Sleep -Seconds 1
    Write-Host -NoNewline "."
}
Write-Host ""

if (-not $apiReady) {
    Exit-WithError "تعذر الاتصال بـ OpenWA على المنفذ 2785. الحاوية تعمل ولكن المنفذ لا يستجيب على 127.0.0.1:2785."
}
Write-Host "  ✓ واجهة OpenWA API مستجيبة بنجاح على http://127.0.0.1:2785" -ForegroundColor Green

# قراءة مفتاح API والتحقق من الجلسة الفعلية
$keyPath = "C:\Users\huawei\OpenWA\data\.api-key"
if (!(Test-Path $keyPath)) {
    Exit-WithError "ملف مفتاح API غير موجود في المسار: $keyPath"
}
$apiKey = (Get-Content $keyPath -Raw).Trim()

$sessionId = "2831a802-6a68-4a62-a062-5d9a23736f93"
$sessionReady = $false
$sessionDetails = $null

Write-Host "  > التحقق من حالة جلسة واتساب ($sessionId)..." -ForegroundColor Gray

for ($s = 0; $s -lt 10; $s++) {
    try {
        $sessionDetails = Invoke-RestMethod -Uri "http://127.0.0.1:2785/api/sessions/$sessionId" -Headers @{ "X-API-Key" = $apiKey } -UseBasicParsing -TimeoutSec 5 -ErrorAction Stop
        if ($sessionDetails.status -eq 'ready') {
            $sessionReady = $true
            break
        }
    } catch {
        Start-Sleep -Seconds 1
    }
}

if (-not $sessionReady) {
    $currentSt = if ($sessionDetails) { $sessionDetails.status } else { "غير معروف" }
    Exit-WithError "جلسة الواتساب ليست بحالة جاهزة (الحالة الحالية: $currentSt). يرجى فتح لوحة OpenWA والتأكد من مسح QR واتصال الجلسة."
}

Write-Host "  ✓ جلسة واتساب متصلة وجاهزة 100%!" -ForegroundColor Green
Write-Host "    - اسم الجلسة: $($sessionDetails.name)" -ForegroundColor Gray
Write-Host "    - رقم الهاتف المتصل: $($sessionDetails.phone)" -ForegroundColor Gray
Write-Host "    - الاسم المستعار: $($sessionDetails.pushName)" -ForegroundColor Gray

# ── 4. تشغيل وفحص خدمة Worker المحلية ──
Write-Host "`n[4/5] تشغيل وفحص خادم الخدمة الوسيطة (Port 8787)..." -ForegroundColor Cyan

# تنظيف أي عمليات سابقة على المنفذ
Get-Process -Name "cloudflared" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

$serverScript = Join-Path $projectRoot "scripts\local-worker-server.mjs"
$workerProcess = Start-Process -FilePath "node" -ArgumentList "`"$serverScript`"" -PassThru -NoNewWindow
Start-Sleep -Seconds 2

$workerOk = $false
for ($w = 0; $w -lt 6; $w++) {
    try {
        $wCheck = Invoke-RestMethod -Uri "http://127.0.0.1:8787/health" -UseBasicParsing -TimeoutSec 3 -ErrorAction Stop
        if ($wCheck.ok -eq $true) {
            $workerOk = $true
            break
        }
    } catch {
        Start-Sleep -Seconds 1
    }
}

if (-not $workerOk) {
    if ($workerProcess -and -not $workerProcess.HasExited) { Stop-Process -Id $workerProcess.Id -Force }
    Exit-WithError "تعذر تشغيل خدمة Worker المحلية على المنفذ 8787."
}
Write-Host "  ✓ خدمة Worker المحلية تعمل ومستجيبة بنجاح." -ForegroundColor Green

# ── 5. تشغيل Cloudflare Tunnel الآمن عبر HTTPS ──
Write-Host "`n[5/5] إنشاء نفق Cloudflare Tunnel الآمن عبر HTTPS..." -ForegroundColor Cyan
$cloudflaredPath = Join-Path $projectRoot "tools\cloudflared.exe"
if (!(Test-Path $cloudflaredPath)) {
    Exit-WithError "أداة cloudflared.exe مفقودة من مجلد tools."
}

$tunnelLog = Join-Path $projectRoot "scratch\tunnel.log"
if (!(Test-Path (Join-Path $projectRoot "scratch"))) { New-Item -ItemType Directory -Path (Join-Path $projectRoot "scratch") | Out-Null }
if (Test-Path $tunnelLog) { Remove-Item $tunnelLog -Force }

$tunnelProcess = Start-Process -FilePath $cloudflaredPath -ArgumentList "tunnel --url http://127.0.0.1:8787" -PassThru -NoNewWindow -RedirectStandardError $tunnelLog

$tunnelUrl = $null
for ($t = 0; $t -lt 30; $t++) {
    Start-Sleep -Seconds 1
    if (Test-Path $tunnelLog) {
        $content = Get-Content $tunnelLog -Raw
        if ($content -match 'https://[a-zA-Z0-9-]+\.trycloudflare\.com') {
            $tunnelUrl = $matches[0]
            break
        }
    }
    Write-Host -NoNewline "."
}
Write-Host ""

if (-not $tunnelUrl) {
    if ($tunnelProcess -and -not $tunnelProcess.HasExited) { Stop-Process -Id $tunnelProcess.Id -Force }
    if ($workerProcess -and -not $workerProcess.HasExited) { Stop-Process -Id $workerProcess.Id -Force }
    Exit-WithError "فشل استخراج رابط نفق Cloudflare Tunnel بعد 30 ثانية. يرجى التحقق من اتصال الإنترنت."
}

# حفظ الرابط ونسخه للحافظة
Set-Content -Path (Join-Path $projectRoot "tools\tunnel-url.txt") -Value $tunnelUrl -Encoding UTF8
try { Set-Clipboard -Value $tunnelUrl } catch {}

# ══════════════════════════════════════════════════════════════
# الإعلان النهائي بالنجاح التام (فقط بعد اجتياز الـ 5 مراحل 100%)
# ══════════════════════════════════════════════════════════════
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  ✓ تـم التحـقـق والـتـشـغـيـل بـنـجـاح تـام 100%!          " -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
Write-Host "  • حالة واتساب: " -NoNewline -ForegroundColor White
Write-Host "مـتـصـل ومـفـعـل وجـاهـز (رقم: $($sessionDetails.phone))" -ForegroundColor Green

Write-Host "  • رابط النفق المشفر (HTTPS):" -ForegroundColor White
Write-Host "    $tunnelUrl" -ForegroundColor Yellow -BackgroundColor Black
Write-Host "    (تم نسخ الرابط تلقائياً إلى الحافظة Clipboard)" -ForegroundColor Gray
Write-Host ""
Write-Host "  • كيف ترسل الرسائل الآن؟" -ForegroundColor Cyan
Write-Host "    1. افتح صفحة الورش في لوحة الإدارة." -ForegroundColor White
Write-Host "    2. اضغط على زر [إرسال واتساب] الأخضر بجانب اسم أي مسجل." -ForegroundColor White
Write-Host "    3. ستصل الرسالة مباشرة إلى جوال المستلم خلال ثوانٍ!" -ForegroundColor White
Write-Host ""
Write-Host "============================================================" -ForegroundColor DarkGray
Write-Host "  [تنبيه]: اترك هذه النافذة مفتوحة أثناء إرسال الرسائل." -ForegroundColor Yellow
Write-Host "  للإيقاف في أي وقت: أغلق هذه النافذة أو شغّل STOP-WHATSAPP.bat" -ForegroundColor Gray
Write-Host "============================================================" -ForegroundColor DarkGray
Write-Host ""

try {
    while ($true) {
        Start-Sleep -Seconds 3
        if ($tunnelProcess.HasExited -or $workerProcess.HasExited) {
            Write-Host "`n[!] توقفت إحدى الخدمات. جارٍ إنهاء الجلسة..." -ForegroundColor Yellow
            break
        }
    }
} finally {
    Write-Host "`nجارٍ إيقاف الخدمات بأمان..." -ForegroundColor Gray
    if ($tunnelProcess -and -not $tunnelProcess.HasExited) { Stop-Process -Id $tunnelProcess.Id -Force -ErrorAction SilentlyContinue }
    if ($workerProcess -and -not $workerProcess.HasExited) { Stop-Process -Id $workerProcess.Id -Force -ErrorAction SilentlyContinue }
}
