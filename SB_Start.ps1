$SBGUI = Join-Path -Path $PSScriptRoot ".\SB_GUI.ps1" 
PowerShell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File $SBGUI  -Force

