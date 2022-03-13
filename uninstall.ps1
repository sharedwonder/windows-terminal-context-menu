# Remove a registry key.
function RemoveKey([Parameter(Mandatory = $true)]$key) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Force -Recurse | Out-Null
    }
}

# Remove all context menus that open Windows Terminal.
function RemoveMenus() {
    Write-Host (Invoke-Expression $translations.RemovingMenus)

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
    Write-Host (Invoke-Expression $translations.RemovingStorage)

    $storage = "$Env:LocalAppData\WindowsTerminalMenuContext"
    if (Test-Path $storage) {
        Remove-Item -Path $storage -Force -Recurse | Out-Null
    }
}

# Get translations.
function GetTranslations() {
    $context = Get-Content -Path "$PSScriptRoot\translations.ini" # Get file translations.
    $context -replace("#.*", "") # Delete comments.
    $language = (Get-ItemProperty 'Registry::HKEY_CURRENT_USER\Control Panel\Desktop' PreferredUILanguages).PreferredUILanguages[0] # Get system language.
    $found = $false # Use to determine if translations corresponding to the system language has been found.

    # Parse file contents.
    do {
        for ($index = 0; $index -lt $context.Count; ++ $index) {
            if ($context[$index] -match "^\[.+\]") {
                if ($context[$index].Equals("[language]".Replace("language", $language))) {
                    $found = $true
                } elseif ($found) {
                    return
                }
            } elseif ($found -and ($context[$index] -match "^\w+=[\s\S]+")) {
                # PowerShell will automatically store and return the results.
                ConvertFrom-StringData -StringData $context[$index]
            }
        }
        # There aren't any translations corresponding to the system language, use default language: English (US).
        if (-not $found) {
            Write-Warning "There aren't any translations corresponding to the system language, use default language: English (US)."
            $language = "en-US"
        }
    } while (-not $found)
}

[System.Text.Encoding]::GetEncoding(65001) | Out-Null # Set the encoding to UTF-8.
$translations = GetTranslations

Write-Host (Invoke-Expression $translations.UninstallingWindowsTerminalContextMenu)

RemoveMenus
RemoveStorage

Write-Host (Invoke-Expression $translations.UninstalledSuccessfully)
exit 0
