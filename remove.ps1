$Storage = "$env:LocalAppData\WindowsTerminalContextMenu"

if (Test-Path $Storage) {
    $Layout = Get-Content "$Storage\layout"

    if ($Layout -eq "Unfolded") {
        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-*"
        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu-*"
        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu-*"
        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu-*"
    } else {
        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenu"

        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu"
        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu"
        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu"
        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu"

        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenu-Elevated"

        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-Elevated"
        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu-Elevated"
        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu-Elevated"
        Remove-Item -Force -Recurse -ErrorAction Ignore "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu-Elevated"
    }

    Remove-Item -Force -Recurse $Storage

    Write-Host "Done"
} else {
    Write-Host "Nothing to remove"
}

exit 0
