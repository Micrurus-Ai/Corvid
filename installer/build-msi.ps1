# Axon Outlook add-in — build the per-user MSI (WiX v3).
# Mirrors AxonOutlook.iss exactly, just as an .msi so IT can deploy it (Intune/GPO) or a user can
# double-click it (no admin). Bakes both API keys into config.json, read from assistant\.env.
#
#   powershell -ExecutionPolicy Bypass -File installer\build-msi.ps1
# Output: installer\Output\Axon-Outlook.msi

$ErrorActionPreference = "Stop"
$root  = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$inst  = Join-Path $root "installer"
$asst  = Join-Path $root "assistant"
$wix   = Join-Path $inst "wix3"
$out   = Join-Path $inst "Output"
New-Item -ItemType Directory -Force $out | Out-Null

# Read the keys the same way build.ps1 does.
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

# Stage config.json (same shape the Inno WriteAddinConfig produces).
if ($mistral) {
    $json = '{"api_base": "https://api.mistral.ai/v1", "api_key": "' + $mistral + '", "model": "mistral-medium-latest",' +
            ' "backup_api_base": "https://api.openai.com/v1", "backup_api_key": "' + $openai + '", "backup_model": "gpt-4o"}'
} else {
    $json = '{"api_base": "https://api.openai.com/v1", "api_key": "' + $openai + '", "model": "gpt-4o"}'
}
$staged = Join-Path $env:TEMP "axon-config.json"
[IO.File]::WriteAllText($staged, $json)

# Compile + link.
$candle = Join-Path $wix "candle.exe"
$light  = Join-Path $wix "light.exe"
$wixobj = Join-Path $env:TEMP "AxonOutlook.wixobj"
$msi    = Join-Path $out "Axon-Outlook.msi"

& $candle -nologo -arch x64 "-dStagedConfig=$staged" -out $wixobj (Join-Path $inst "AxonOutlook.wxs")
if ($LASTEXITCODE -ne 0) { throw "candle failed." }
# -sval: skip ICE validation (per-user profile-dir keypaths trip harmless ICE warnings).
& $light -nologo -sval -b $inst -out $msi $wixobj
if ($LASTEXITCODE -ne 0) { throw "light failed." }

Remove-Item $staged -Force -EA SilentlyContinue
$mb = "{0:N2} MB" -f ((Get-Item $msi).Length / 1MB)
Write-Host "DONE:  $msi  ($mb)" -ForegroundColor Green
