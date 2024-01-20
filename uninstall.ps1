$storage = "$Env:LocalAppData\WindowsTerminalContextMenu"

if (Test-Path $storage) {
    Write-Host "Uninstalling..."

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

    Remove-Item -Path $storage -Force -Recurse | Out-Null

    Write-Host "Done."
} else {
    Write-Host "Not installed."
}

exit 0
