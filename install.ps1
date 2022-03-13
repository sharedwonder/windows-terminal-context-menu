#Requires -Version 6

# Get the edition and the installation folder of Windows Terminal.
function GetInstallationInfo() {
    Write-Host (Invoke-Expression $translations.LookingForWindowsTerminal)

    $edition = 0

    # release edition
    if ($null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminal))) {
        $version = $appx.Version
        Write-Host (Invoke-Expression $translations.FoundWindowsTerminal)
        $folder = $appx.InstallLocation
        $edition = 1
    }

    # preview edition
    if ($null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminalPreview))) {
        $version = $appx.Version
        Write-Host (Invoke-Expression $translations.FoundWindowsTerminalPreview)
        if ($edition -eq 0) {
            $edition = 2
            $folder = $appx.InstallLocation
        } else {
            # Found multiple editions.
            do {
                Write-Host (Invoke-Expression $translations.SelectEdition)
                $edition = Read-Host
            } while (($edition -eq 1) -or ($edition -eq 2))

            if ($edition -eq 2) {
                $folder = $appx.InstallLocation
            }
        }
    }

    # Not found.
    if ($edition -eq 0) {
        Write-Error (Invoke-Expression $translations.NotInstalledWindowsTerminal)
        exit 1
    } elseif ($version -lt "1.0") {
        Write-Warning (Invoke-Expression $translations.WindowsTerminalVersionTooOld)
    }

    Write-Host (Invoke-Expression $translations.WindowsTerminalInstallationFolder)

    return $edition, $folder
}

# Get active profiles of Windows Terminal.
function GetActiveProfiles([Parameter(Mandatory = $true)][int]$edition) {
    if ($edition -eq 1) {
        $file = "$env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
    } else {
        $file = "$env:LocalAppData\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\settings.json"
    }

    if (-not (Test-Path $file)) {
        Write-Error (Invoke-Expression $translations.SettingsNotExist)
        exit 1
    }

    # Read and parse settings of Windows Terminal.
    $settings = Get-Content $file | Out-String | ConvertFrom-Json

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

    $iconWriter.Write([byte]0)
    $iconWriter.Write([byte]0)
    $iconWriter.Write([short]1)
    $iconWriter.Write([short]1)
    $iconWriter.Write([byte]$width)
    $iconWriter.Write([byte]$height)
    $iconWriter.Write([byte]0)
    $iconWriter.Write([byte]0)
    $iconWriter.Write([short]0)
    $iconWriter.Write([short]32)
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
function GetProfileIcon([Parameter(Mandatory = $true)]$profile, [Parameter(Mandatory = $true)][string]$folder,
                        [Parameter(Mandatory = $true)][string]$defaultIcon, [Parameter(Mandatory = $true)][int]$edition) {
    if ($profile.source -eq "Windows.Terminal.Wsl") {
        $guid = "{9acb9455-ca41-5af7-950f-6bca1bc9722f}"
    } else {
        $guid = $profile.guid
    }

    if ($null -ne $profile.icon) {
        if ($profile.icon -match "^ms-appx:///.*") {
            $iconFile = $folder + "\" + ($profile.icon -replace ("ms-appx:///", "") -replace ("/", "\"))
            if (-not ($iconFile -match "^.*\.scale-.*\.png$")) {
                $iconFile = $iconFile -replace ("\.png$", ".scale-200.png")
            }
        } elseif ($profile.icon -match "^ms-appdata:///Local/.*") {
            if ($edition -eq 1) {
                $iconFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\" +
                    ($profile.icon -replace ("ms-appdata:///Local/", "") -replace ("/", "\"))
            } else {
                $iconFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\" +
                    ($profile.icon -replace ("ms-appdata:///Local/", "") -replace ("/", "\"))
            }
        } elseif ($profile.icon -match "^ms-appdata:///Roaming/.*") {
            if ($edition -eq 1) {
                $iconFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\RoamingState\" +
                    ($profile.icon -replace ("ms-appdata:///Roaming/", "") -replace ("/", "\"))
            } else {
                $iconFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\RoamingState\" +
                    ($profile.icon -replace ("ms-appdata:///Roaming/", "") -replace ("/", "\"))
            }
        } else {
            $iconFile = [System.Environment]::ExpandEnvironmentVariables($profile.icon)
        }
    } else {
        $iconFile = "$folder\ProfileIcons\$guid.scale-200.png"
    }

    if (Test-Path $iconFile) {
        if ($iconFile -match ".*\.ico$") {
            return $iconFile
        }

        $newIcon = "$storage\$guid.ico"
        ConvertToIcon $iconFile $newIcon
        return $newIcon
    } else {
        return $defaultIcon
    }
}

# Add menu subitems for a profile.
function AddProfileMenuItem([Parameter(Mandatory = $true)]$profile, [Parameter(Mandatory = $true)]$index,
                            [Parameter(Mandatory = $true)][string]$folder, [Parameter(Mandatory = $true)][string]$defaultIcon,
                            [Parameter(Mandatory = $true)][string]$launcher, [Parameter(Mandatory = $true)][int]$edition) {
    $guid = $profile.guid
    $name = $profile.name
    $icon = GetProfileIcon $profile $folder $defaultIcon $edition

    Write-Host (Invoke-Expression $translations.AddingMenuSubitems)

    $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalMenu\shell\$index-$guid"
    $command = "wscript `"$launcher`" `"%V\.`" $guid"

    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name 'MUIVerb' -PropertyType String -Value $name | Out-Null
    New-ItemProperty -Path $key -Name 'Icon' -PropertyType String -Value $icon | Out-Null
    New-Item -Path "$key\command" -Force | Out-Null
    New-ItemProperty -Path "$key\command" -Name '(Default)' -PropertyType String -Value $command | Out-Null

    $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalMenuElevated\shell\$index-$guid"
    $command = "wscript `"$PSScriptRoot\launch.vbs`" `"%V\.`" $guid elevated"

    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name 'MUIVerb' -PropertyType String -Value $name | Out-Null
    New-ItemProperty -Path $key -Name 'Icon' -PropertyType String -Value $icon | Out-Null
    New-ItemProperty -Path $key -Name 'HasLUAShield' -PropertyType String -Value '' | Out-Null
    New-Item -Path "$key\command" -Force | Out-Null
    New-ItemProperty -Path "$key\command" -Name '(Default)' -PropertyType String -Value $command | Out-Null
}

# Add a menu that open Windows Terminal.
function AddMenu([Parameter(Mandatory = $true)][String]$key, [Parameter(Mandatory = $true)][String]$icon,
                 [Parameter(Mandatory = $true)][bool]$elevated) {
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name 'Icon' -PropertyType String -Value $icon | Out-Null

    if ($elevated) {
        New-ItemProperty -Path $key -Name 'MUIVerb' -PropertyType String -Value (Invoke-Expression $translations.WindowsTerminalMenuElevated) | Out-Null
        New-ItemProperty -Path $key -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'WindowsTerminalMenuElevated' | Out-Null
        New-ItemProperty -Path $key -Name 'HasLUAShield' -PropertyType String -Value '' | Out-Null
    } else {
        New-ItemProperty -Path $key -Name 'MUIVerb' -PropertyType String -Value (Invoke-Expression $translations.WindowsTerminalMenu) | Out-Null
        New-ItemProperty -Path $key -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'WindowsTerminalMenu' | Out-Null
    }
}

# Create all context menus that open Windows Terminal.
function CreateMenus([Parameter(Mandatory = $true)][string]$storage, [Parameter(Mandatory = $true)][string]$icon) {
    Copy-Item $icon "$storage\WindowsTerminal.ico"
    $icon = "$storage\WindowsTerminal.ico"

    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalMenu" $icon $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalMenuElevated" $icon $true
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalMenu" $icon $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalMenuElevated" $icon $true
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalMenu" $icon $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalMenuElevated" $icon $true
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalMenu" $icon $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalMenuElevated" $icon $true
}

# Get translations.
function GetTranslations() {
    $context = Get-Content -Path "$PSScriptRoot\translations.ini" # Read translations file.
    $context -replace ("#.*", "") # Delete comments.
    $language = (Get-ItemProperty 'Registry::HKEY_CURRENT_USER\Control Panel\Desktop' PreferredUILanguages).PreferredUILanguages[0] # Get system language.
    $found = $false # Use to determine if translations corresponding to the system language has been found.

    # Parse file contents.
    do {
        for ($index = 0; $index -lt $context.Count; ++ $index) {
            if ($context[$index] -match "^\[.+\]") {
                if ($context[$index] -eq "[$language]") {
                    $found = $true
                } elseif ($found) {
                    return
                }
            } elseif ($found -and ($context[$index] -match "^\w+=.+")) {
                # PowerShell will automatically store and return the results.
                ConvertFrom-StringData -StringData $context[$index]
            }
        }

        # There aren't any translations corresponding to the system language, using default language: English (US).
        if (-not $found) {
            Write-Warning "There aren't any translations corresponding to the system language, using default language: English (US)."
            $language = "en-US"
        }
    } while (-not $found)
}

[System.Text.Encoding]::GetEncoding(65001) | Out-Null # Set the encoding to UTF-8.
$translations = GetTranslations

Write-Host (Invoke-Expression $translations.InstallingWindowsTerminalContextMenu)

$info = GetInstallationInfo
$edition = $info[0]
$folder = $info[1]
$profiles = GetActiveProfiles $edition
$icon = "$PSScriptRoot\icon.ico"

$storage = "$env:LocalAppData\WindowsTerminalMenuContext"
if (-not (Test-Path $storage)) {
    New-Item -Path $storage -ItemType Directory | Out-Null
}

CreateMenus $storage $icon
Copy-Item "$PSScriptRoot\launch.vbs" "$storage\launch.vbs"
for ($index = 0; $index -lt $profiles.Count; ++ $index) {
    AddProfileMenuItem $profiles[$index] $index $folder $icon "$storage\launch.vbs" $edition
}

Write-Host (Invoke-Expression $translations.InstalledSuccessfully)
exit 0
