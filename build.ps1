# Axon intelligence — one-command build.
# Rebuilds the Outlook add-in, the dot (PyInstaller), bundles open-computer-use, and compiles the
# installer into installer\Output\Axon-Setup.exe.
#
#   powershell -ExecutionPolicy Bypass -File build.ps1
# (or just double-click build.bat)

$ErrorActionPreference = "Stop"
$root  = Split-Path -Parent $MyInvocation.MyCommand.Path
$asst  = Join-Path $root "assistant"
$addin = Join-Path $root "outlook_addin"
$dist  = Join-Path $asst "dist\AxonIntelligence"

function Step($n) { Write-Host "`n=== $n ===" -ForegroundColor Cyan }

# --- 0. Close Outlook (it locks the add-in DLL while running) -----------------
Step "0/4  Closing Outlook (it holds the add-in DLL; reopen it after the build)"
try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application"); $ol.Quit() } catch {}
Start-Sleep 3
Get-Process OUTLOOK -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep 1

# --- 1. Compile the Outlook add-in DLL ---------------------------------------
Step "1/4  Building the Outlook add-in (AxonAddin.dll)"
$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$ext = (Get-ChildItem "C:\Windows\assembly\GAC\Extensibility" -Recurse -Filter "extensibility.dll" -EA SilentlyContinue | Select-Object -First 1).FullName
$off = (Get-ChildItem "C:\Windows\assembly\GAC_MSIL\office"    -Recurse -Filter "OFFICE.DLL"       -EA SilentlyContinue | Select-Object -First 1).FullName
$std = (Get-ChildItem "C:\Windows\assembly\GAC\stdole"         -Recurse -Filter "stdole.dll"       -EA SilentlyContinue | Select-Object -First 1).FullName
if (-not ($ext -and $off -and $std)) { throw "Office interop assemblies not found in the GAC (need Outlook installed)." }
& $csc /nologo /target:library /out:"$addin\AxonAddin.dll" /link:"$ext" /link:"$off" /link:"$std" `
    /reference:System.Windows.Forms.dll /reference:System.Drawing.dll /reference:System.Web.Extensions.dll /reference:Microsoft.CSharp.dll /reference:System.dll /reference:System.Net.Http.dll `
    "$addin\AxonAddin.cs"
if ($LASTEXITCODE -ne 0) { throw "Add-in compile failed (is Outlook holding AxonAddin.dll? close Outlook and retry)." }
Write-Host "   add-in built." -ForegroundColor Green

# --- 2. Build the dot (PyInstaller) ------------------------------------------
Step "2/4  Building the dot (AxonIntelligence.exe)"
Push-Location $asst
try {
    & ".\.venv\Scripts\pyinstaller.exe" --noconfirm --windowed --name AxonIntelligence `
        --icon "$root\installer\axon.ico" `
        --add-data ".env;." --collect-all browser_use --collect-all certifi --collect-all openai `
        --collect-all pptx --collect-all openpyxl --collect-all docx `
        --collect-all pdfplumber --collect-all pypdf --collect-all pdfminer `
        --collect-all sounddevice --collect-all soundfile overlay.py | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "PyInstaller build failed." }
} finally { Pop-Location }
Write-Host "   dot built." -ForegroundColor Green

# --- 3. Stage the bundled open-computer-use next to the exe -------------------
Step "3/4  Staging open-computer-use"
$ocuSrc = Join-Path $asst "ocu"
if (-not (Test-Path "$ocuSrc\open-computer-use.cmd")) {
    throw "Bundled open-computer-use missing at assistant\ocu. (One-time setup: copy node.exe + ``npm install @qwen-code/open-computer-use --prefix assistant\ocu`` + the launcher.)"
}
Copy-Item $ocuSrc (Join-Path $dist "ocu") -Recurse -Force
Write-Host "   open-computer-use staged." -ForegroundColor Green

# --- 4. Compile the two installers -------------------------------------------
Step "4/4  Compiling the installers (dot + Outlook add-in)"
$iscc = @("$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe",
          "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe",
          "$env:ProgramFiles\Inno Setup 6\ISCC.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $iscc) { throw "Inno Setup not found. Install it: winget install JRSoftware.InnoSetup" }

# Read the keys (from assistant\.env) — the standalone Outlook add-in bakes both into its config.json.
function Read-EnvKey($name) {
    $envFile = Join-Path $asst ".env"
    if (Test-Path $envFile) {
        $m = Select-String -Path $envFile -Pattern ("^\s*" + $name + "\s*=\s*(.+)$") | Select-Object -First 1
        if ($m) { return $m.Matches[0].Groups[1].Value.Trim() }
    }
    return ""
}
$mistral = Read-EnvKey "MISTRAL_API_KEY"
$openai  = Read-EnvKey "OPENAI_API_KEY"

# 4a. The dot (floating dot only) -> Axon-Dot-Setup.exe
& $iscc (Join-Path $root "installer\Axon.iss") | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Dot installer compile failed." }

# 4b. The Outlook add-in (standalone) -> Axon-Outlook-Setup.exe (both keys baked into config.json)
& $iscc "/DMistralKey=$mistral" "/DOpenAIKey=$openai" (Join-Path $root "installer\AxonOutlook.iss") | Out-Null
if ($LASTEXITCODE -ne 0) { throw "Outlook add-in installer compile failed." }

$dot = Join-Path $root "installer\Output\Axon-Dot-Setup.exe"
$olk = Join-Path $root "installer\Output\Axon-Outlook-Setup.exe"
$dotMb = "{0:N0} MB" -f ((Get-Item $dot).Length / 1MB)
$olkMb = "{0:N1} MB" -f ((Get-Item $olk).Length / 1MB)
Write-Host "`nDONE:" -ForegroundColor Green
Write-Host "  Dot     -> $dot  ($dotMb)" -ForegroundColor Green
Write-Host "  Outlook -> $olk  ($olkMb)" -ForegroundColor Green
