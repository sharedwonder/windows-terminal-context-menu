# Remove a registry key.
function RemoveKey([Parameter(Mandatory = $true)]$key) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Force -Recurse | Out-Null
    }
}

# Remove all context menus that open Windows Terminal.
function RemoveMenus() {
    Write-Host "Removing all context menus that open Windows Terminal..."

    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalMenu"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalMenuElevated"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalMenu"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalMenuElevated"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalMenu"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalMenuElevated"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalMenu"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalMenuElevated"

    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalMenu"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalMenuElevated"
}

# Remove the storage this software folder
function RemoveStorage() {
    Write-Host "Removing the storage this software folder..."

    $storage = "$Env:LocalAppData\WindowsTerminalMenuContext"
    if (Test-Path $storage) {
        Remove-Item -Path $storage -Force -Recurse | Out-Null
    }
}

RemoveMenus
RemoveStorage
Write-Host "Uninstalled successfully."
exit 0
