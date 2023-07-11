# Removes a registry key.
function RemoveKey([Parameter(Mandatory = $true)]$key) {
    if (Test-Path $key) {
        Remove-Item -Path $key -Force -Recurse | Out-Null
    }
}

# Removes all context menus that open Windows Terminal.
function RemoveMenus() {
    Write-Host (Invoke-Expression $translations.RemovingMenus)

    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenu"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenuElevated"

    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenuElevated"

    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenuElevated"

    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenuElevated"

    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu"
    RemoveKey "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenuElevated"
}

# Removes the storage this software folder
function RemoveStorage() {
    Write-Host (Invoke-Expression $translations.RemovingStorage)

    $storage = "$Env:LocalAppData\WindowsTerminalContextMenuContext"
    if (Test-Path $storage) {
        Remove-Item -Path $storage -Force -Recurse | Out-Null
    }
}

# Gets translation strings.
function GetTranslations() {
    $context = Get-Content -Path "$PSScriptRoot\translations.ini" # Read the translation file.
    $language = (Get-ItemProperty 'Registry::HKEY_CURRENT_USER\Control Panel\Desktop' PreferredUILanguages).PreferredUILanguages[0] # Get the language of the current user.
    $found = $false # Uses to determine if translations corresponding to the system language has been found.

    # Parse file contents.
    do {
        for ($index = 0; $index -lt $context.Count; ++ $index) {
            if ($context[$index] -match "^\[.+\]") {
                if ($context[$index] -eq "[$language]") {
                    $found = $true
                } elseif ($found) {
                    return
                }
            } elseif ($found -and ($context[$index] -match "^\w+=.*")) {
                # Automatically return as a list.
                ConvertFrom-StringData -StringData $context[$index]
            }
        }

        if (-not $found) {
            Write-Warning "There is no translation corresponding to the system language, and the default language is used: English (US)."
            $language = "en-US"
        }
    } while (-not $found)
}

$translations = GetTranslations

Write-Host (Invoke-Expression $translations.UninstallingWindowsTerminalContextMenu)

RemoveMenus
RemoveStorage

Write-Host (Invoke-Expression $translations.UninstalledSuccessfully)

exit 0
