# Removes all context menus that open Windows Terminal.
function RemoveMenus() {
    Write-Host (Invoke-Expression $translations.RemovingMenus)

    $layout = Get-Content "$storage\layout.txt"

    if ($layout -eq 'Unfolded') {
        Get-ChildItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu*" |
            ForEach-Object { Remove-Item -Path "Registry::" + $_.Name -Force -Recurse -ErrorAction Ignore | Out-Null }
        Get-ChildItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu*" |
            ForEach-Object { Remove-Item -Path "Registry::" + $_.Name -Force -Recurse -ErrorAction Ignore | Out-Null }
        Get-ChildItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu*" |
            ForEach-Object { Remove-Item -Path "Registry::" + $_.Name -Force -Recurse -ErrorAction Ignore | Out-Null }
        Get-ChildItem "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu*" | `
            ForEach-Object { Remove-Item -Path "Registry::" + $_.Name -Force -Recurse -ErrorAction Ignore | Out-Null }
    } else {
        Remove-Item -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu" -Force -Recurse -ErrorAction Ignore | Out-Null
        Remove-Item -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenuElevated"  -Force -Recurse -ErrorAction Ignore | Out-Null
        Remove-Item -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu"  -Force -Recurse -ErrorAction Ignore | Out-Null
        Remove-Item -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenuElevated"  -Force -Recurse -ErrorAction Ignore | Out-Null
        Remove-Item -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu"  -Force -Recurse -ErrorAction Ignore | Out-Null
        Remove-Item -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenuElevated"  -Force -Recurse -ErrorAction Ignore | Out-Null
        Remove-Item -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu"  -Force -Recurse -ErrorAction Ignore | Out-Null
        Remove-Item -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenuElevated"  -Force -Recurse -ErrorAction Ignore | Out-Null

        Remove-Item -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenu"  -Force -Recurse -ErrorAction Ignore | Out-Null
        Remove-Item -Path "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenuElevated"  -Force -Recurse -ErrorAction Ignore | Out-Null
    }
}

# Removes the storage this software folder
function RemoveStorage() {
    Write-Host (Invoke-Expression $translations.RemovingStorage)

    if (Test-Path $storage) {
        Remove-Item -Path $storage -Force -Recurse | Out-Null
    }
}

# Gets translation strings.
function GetTranslations() {
    $context = Get-Content -Path "$PSScriptRoot\translations.ini" # Read the translation file.
    # Get the language of the current user.
    $language = (Get-ItemProperty 'Registry::HKEY_CURRENT_USER\Control Panel\Desktop' PreferredUILanguages).PreferredUILanguages[0].ToLower()
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
            $language = "en-us"
        }
    } while (-not $found)
}

$translations = GetTranslations

Write-Host (Invoke-Expression $translations.UninstallingWindowsTerminalContextMenu)

$storage = "$Env:LocalAppData\WindowsTerminalContextMenu"

RemoveMenus
RemoveStorage

Write-Host (Invoke-Expression $translations.UninstalledSuccessfully)

exit 0
