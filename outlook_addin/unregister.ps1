# Removes the Axon Outlook add-in registration for the current user.
$clsid = "{7B2C9E14-6A3D-4F58-9C21-3E5A1B7D4F60}"
$prog  = "Axon.OutlookAddin"
Remove-Item -Path "HKCU:\Software\Microsoft\Office\Outlook\AddIns\$prog" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Classes\CLSID\$clsid" -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path "HKCU:\Software\Classes\$prog" -Recurse -Force -ErrorAction SilentlyContinue
Write-Output "Axon Outlook add-in unregistered. Restart Outlook to fully unload it."
