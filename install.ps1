#Requires -Version 6

# Get the edition and the installation folder of Windows Terminal.
function GetInstallationInfo() {
    Write-Host "Looking for Windows Terminal..."

    $edition = 0
    $appx = $null

    # release edition
    if ($null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminal))) {
        Write-Host "Found Windows Terminal:" $appx.Version
        $folder = $appx.InstallLocation
        $edition = 1
    }

    # preview edition
    if ($null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminalPreview))) {
        Write-Host "Found Windows Terminal Preview:" $appx.Version
        if ($edition -ne 0) {
            $edition = 2
            $folder = $appx.InstallLocation
        } else {
            do {
                $edition = Read-Host -Prompt "Select edition [1: Release 2: Preview]"
            } while (($edition -eq 1) -or ($edition -eq 2))

            if ($edition == 2) {
                $folder = $appx.InstallLocation
            }
        }
    }

    # Not found.
    if ($edition -eq 0) {
        Write-Error "Not installed Windows Terminal."
        exit 1
    }

    return ($edition, $folder)
}

# Get active profiles of Windows Terminal.
function GetActiveProfiles([Parameter(Mandatory = $true)][int]$edition) {
    $settings = $null
    if ($edition -eq 1) {
        # Read and parse settings of Windows Terminal.
        $settings = Get-Content "$Env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json" |
            Out-String | ConvertFrom-Json
    } else {
        # Read and parse settings of Windows Terminal Preview.
        $settings = Get-Content "$Env:LocalAppData\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\settings.json" |
            Out-String | ConvertFrom-Json
    }

    # Compatible with older Windows Terminal.
    if ($settings.profiles.PSObject.Properties.name -match "list") {
        $list = $settings.profiles.list
    }
    else {
        $list = $settings.profiles
    }

    # Exclude the disabled profiles and return active profiles.
    return $list | Where-Object {-not $_.hidden} | Where-Object {($null -eq $_.source) -or -not ($settings.disabledProfileSources -contains $_.source)}
}

# Convert PNG to ICO (icon).
function ConvertToIcon([Parameter(Mandatory = $true)][string]$file, [Parameter(Mandatory = $true)][string]$outputFile) {
    Add-Type -AssemblyName System.Drawing

    $resolvedFile = $ExecutionContext.SessionState.Path.GetResolvedPSPathFromPSPath($file)
    if (-not $resolvedFile) {
        return
    }
    $inputBitmap = [Drawing.Image]::FromFile($resolvedFile)
    $width = $inputBitmap.Width
    $height = $inputBitmap.Height
    $size = New-Object Drawing.Size $width, $height
    $newBitmap = New-Object Drawing.Bitmap $inputBitmap, $size

    if ($width -gt 255 -or $height -gt 255) {
        $ratio = ($height, $width | Measure-Object -Maximum).Maximum / 255
        $width /= $ratio
        $height /= $ratio
    }

    $memoryStream = New-Object System.IO.MemoryStream
    $newBitmap.Save($memoryStream, [System.Drawing.Imaging.ImageFormat]::Png)

    $resolvedOutputFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($outputFile)
    $output = [IO.File]::Create("$resolvedOutputFile")

    $iconWriter = New-Object System.IO.BinaryWriter($output)

    $iconWriter.Write([char]0)
    $iconWriter.Write([char]0)
    $iconWriter.Write([int16]1)
    $iconWriter.Write([int16]1)
    $iconWriter.Write([char]$width)
    $iconWriter.Write([char]$height)
    $iconWriter.Write([char]0)
    $iconWriter.Write([char]0)
    $iconWriter.Write([int16]0)
    $iconWriter.Write([int16]32)
    $iconWriter.Write([int]$memoryStream.Length)
    $iconWriter.Write([int]22)
    $iconWriter.Write($memoryStream.ToArray())

    $iconWriter.Flush()
    $output.Close()

    $memoryStream.Dispose()
    $newBitmap.Dispose()
    $inputBitmap.Dispose()
}

# Get the icon of a profile.
function GetProfileIcon([Parameter(Mandatory = $true)]$profile, [Parameter(Mandatory = $true)][String]$folder,
                        [Parameter(Mandatory = $true)][String]$defaultIcon) {
    if ($null -ne $profile.icon) {
        # For profiles with a user-defined icon.
        return $profile.icon
    } else {
        if ($profile.source -eq "Windows.Terminal.Wsl") {
            # For WSL (Windows Subsystem for Linux).
            $guid = "{9acb9455-ca41-5af7-950f-6bca1bc9722f}"
        } else {
            $guid = $profile.guid
        }

        $profilePng = "$folder\ProfileIcons\$guid.scale-200.png"
        if (Test-Path $profilePng) {
            # For automatically generated profiles.
            $cache = "$Env:LocalAppData\WindowsTerminalIconsCache"
            $icon = "$cache\$guid.ico"

            if (-not (Test-Path $cache)) {
                New-Item -Path $cache -ItemType Directory
            }

            ConvertToIcon $profilePng $icon
            return $icon
        } else {
            # For profiles without a icon.
            return $defaultIcon
        }
    }
}

# Add menu subitems for a profile.
function AddProfileMenuItem([Parameter(Mandatory = $true)]$profile, [Parameter(Mandatory = $true)]$index,
                            [Parameter(Mandatory = $true)][String]$folder, [Parameter(Mandatory = $true)][String]$defaultIcon) {
    Write-Host "Adding menu subitems for the profile" $profile.guid ":" $profile.name
    $guid = $profile.guid
    $name = $profile.name
    $icon = GetProfileIcon $profile $folder $defaultIcon

    $rootKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalMenu\shell\$index-$guid"
    $command = "wscript `"$PSScriptRoot\launch.vbs`" `"%V\.`" $guid"

    New-Item -Path $rootKey -Force | Out-Null
    New-ItemProperty -Path $rootKey -Name 'MUIVerb' -PropertyType String -Value $name | Out-Null
    New-ItemProperty -Path $rootKey -Name 'Icon' -PropertyType String -Value $icon | Out-Null
    New-Item -Path "$rootKey\command" -Force | Out-Null
    New-ItemProperty -Path "$rootKey\command" -Name '(Default)' -PropertyType String -Value $command | Out-Null

    $rootKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalMenuElevated\shell\$index-$guid"
    $command = "wscript `"$PSScriptRoot\launch.vbs`" `"%V\.`" $guid elevated"

    New-Item -Path $rootKey -Force | Out-Null
    New-ItemProperty -Path $rootKey -Name 'MUIVerb' -PropertyType String -Value $name | Out-Null
    New-ItemProperty -Path $rootKey -Name 'Icon' -PropertyType String -Value $icon | Out-Null
    New-ItemProperty -Path $rootKey -Name 'HasLUAShield' -PropertyType String -Value '' | Out-Null
    New-Item -Path "$rootKey\command" -Force | Out-Null
    New-ItemProperty -Path "$rootKey\command" -Name '(Default)' -PropertyType String -Value $command | Out-Null
}

# Add a menu that open Windows Terminal.
function AddMenu([Parameter(Mandatory = $true)][String]$rootKey, [Parameter(Mandatory = $true)][String]$icon,
    [Parameter(Mandatory = $true)] $translations, [Parameter(Mandatory = $true)] [bool] $elevated) {
    New-Item -Path $rootKey -Force | Out-Null
    New-ItemProperty -Path $rootKey -Name 'Icon' -PropertyType String -Value $icon | Out-Null

    if ($elevated) {
        New-ItemProperty -Path $rootKey -Name 'MUIVerb' -PropertyType String -Value $translations.WindowsTerminalMenuElevated | Out-Null
        New-ItemProperty -Path $rootKey -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'WindowsTerminalMenuElevated' | Out-Null
        New-ItemProperty -Path $rootKey -Name 'HasLUAShield' -PropertyType String -Value '' | Out-Null
    }
    else {
        New-ItemProperty -Path $rootKey -Name 'MUIVerb' -PropertyType String -Value $translations.WindowsTerminalMenu | Out-Null
        New-ItemProperty -Path $rootKey -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'WindowsTerminalMenu' | Out-Null
    }
}

# Create all context menus that open Windows Terminal.
function CreateMenus([Parameter(Mandatory = $true)][String]$icon, [Parameter(Mandatory = $true)]$translations) {
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalMenu" $icon $translations $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalMenuElevated" $icon $translations $true
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalMenu" $icon $translations $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalMenuElevated" $icon $translations $true
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalMenu" $icon $translations $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalMenuElevated" $icon $translations $true
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalMenu" $icon $translations $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalMenuElevated" $icon $translations $true
}

# Get translations.
function GetTranslations() {
    $context = Get-Content -Path "$PSScriptRoot\lang.ini" # Get file translations.
    $context -replace("#.*", "") # Delete comments.
    $language = (Get-ItemProperty 'Registry::HKEY_CURRENT_USER\Control Panel\Desktop' PreferredUILanguages).PreferredUILanguages[0] # Get system language.
    $flag = $false # Use to determine if translations corresponding to the system language has been found.

    # Parse file contents.
    do {
        for ($index = 1; $index -lt $context.Count; ++ $index) {
            if ($context[$index] -match "\[.+\]") {
                if ($context[$index].Equals("[language]".Replace("language", $language))) {
                    $flag = $true
                } elseif ($flag) {
                    return
                }
            } elseif ($flag) {
                # PowerShell will automatically store and return the results.
                ConvertFrom-StringData -StringData $context[$index]
            }
        }
        # There aren't any translations corresponding to the system language, use default language: English (US).
        if (-not $flag) {
            Write-Warning "There aren't any translations corresponding to the system language, use default language: English (US)."
            $language = "en-US"
        }
    } while (-not $flag)
}

[System.Text.Encoding]::GetEncoding(65001) | Out-Null # Set the encoding to UTF-8.
$translations = GetTranslations

Write-Host "Installing Windows Terminal Context Menu..."

$info = GetInstallationInfo
$edition = $info[0]
$folder = $info[1]
Write-Host "Windows Terminal installtion folder:" $folder
$profiles = GetActiveProfiles $edition
$icon = "$PSScriptRoot\icon.ico"
CreateMenus $icon $translations
for ($index = 0; $index -lt $profiles.Count; ++ $index) {
    AddProfileMenuItem $profiles[$index] $index $folder $icon
}
Write-Host "Installed successfully."
exit 0
