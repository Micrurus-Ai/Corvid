# Registers the Axon Outlook add-in for the CURRENT USER (no admin required).
# Writes the managed-COM class under HKCU\Software\Classes and the Outlook add-in entry.
param([string]$Dll = "$PSScriptRoot\AxonAddin.dll")

$clsid = "{7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60}"
$prog  = "Axon.OutlookAddin"
$asm   = "AxonAddin, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null"
$code  = "file:///" + ((Resolve-Path $Dll).Path -replace '\\', '/')
$mscoree = "$env:WINDIR\System32\mscoree.dll"
$classes = "HKCU:\Software\Classes"

# ProgId -> CLSID
New-Item -Path "$classes\$prog\CLSID" -Force | Out-Null
Set-ItemProperty -Path "$classes\$prog\CLSID" -Name "(default)" -Value $clsid

# CLSID -> managed in-proc server (mscoree shim points at our class/assembly)
$cls = "$classes\CLSID\$clsid"
$inproc = "$cls\InprocServer32"
New-Item -Path $inproc -Force | Out-Null
Set-ItemProperty -Path $inproc -Name "(default)"       -Value $mscoree
Set-ItemProperty -Path $inproc -Name "ThreadingModel"  -Value "Both"
Set-ItemProperty -Path $inproc -Name "Class"           -Value "Axon.OutlookAddin.Connect"
Set-ItemProperty -Path $inproc -Name "Assembly"        -Value $asm
Set-ItemProperty -Path $inproc -Name "RuntimeVersion"  -Value "v4.0.30319"
Set-ItemProperty -Path $inproc -Name "CodeBase"        -Value $code
New-Item -Path "$inproc\0.0.0.0" -Force | Out-Null
Set-ItemProperty -Path "$inproc\0.0.0.0" -Name "Class"          -Value "Axon.OutlookAddin.Connect"
Set-ItemProperty -Path "$inproc\0.0.0.0" -Name "Assembly"       -Value $asm
Set-ItemProperty -Path "$inproc\0.0.0.0" -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty -Path "$inproc\0.0.0.0" -Name "CodeBase"       -Value $code
New-Item -Path "$cls\ProgId" -Force | Out-Null
Set-ItemProperty -Path "$cls\ProgId" -Name "(default)" -Value $prog

# Outlook add-in entry (LoadBehavior 3 = load at startup)
$ak = "HKCU:\Software\Microsoft\Office\Outlook\AddIns\$prog"
New-Item -Path $ak -Force | Out-Null
Set-ItemProperty -Path $ak -Name "FriendlyName" -Value "Axon Email Filer"
Set-ItemProperty -Path $ak -Name "Description"  -Value "File and download emails with Axon"
New-ItemProperty -Path $ak -Name "LoadBehavior" -PropertyType DWord -Value 3 -Force | Out-Null

Write-Output "Axon Outlook add-in registered for current user. Restart Outlook to load it."
