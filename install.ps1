#Requires -Version 6

# Get the edition and the installation folder of Windows Terminal.
function GetInstallationInfo() {
    Write-Host (Invoke-Expression $translations.LookingForWindowsTerminal)

    $edition = 0

    # release edition
    if ($null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminal))) {
        Write-Host (Invoke-Expression $translations.FoundWindowsTerminal.Replace('$version', $appx.Version))
        $folder = $appx.InstallLocation
        $edition = 1
    }

    # preview edition
    if ($null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminalPreview))) {
        Write-Host (Invoke-Expression $translations.FoundWindowsTerminalPreview.Replace('$version', $appx.Version))
        if ($edition -ne 0) {
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
    }

    Write-Host (Invoke-Expression $translations.WindowsTerminalInstallationFolder.Replace('$folder', $folder))

    return $edition, $folder
}

# Get active profiles of Windows Terminal.
function GetActiveProfiles([Parameter(Mandatory = $true)][int]$edition) {
    if ($edition -eq 1) {
        $file = "$Env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"

    } else {
        $file = "$Env:LocalAppData\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\settings.json"
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
                        [Parameter(Mandatory = $true)][string]$defaultIcon) {
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
            $icon = "$storage\$guid.ico"

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
                            [Parameter(Mandatory = $true)][string]$folder, [Parameter(Mandatory = $true)][string]$defaultIcon,
                            [Parameter(Mandatory = $true)][string]$launcher) {
    Write-Host (Invoke-Expression $translations.AddingMenuSubitems.Replace('$guid', $profile.guid).Replace('$name', $profile.name))

    $guid = $profile.guid
    $name = $profile.name
    $icon = GetProfileIcon $profile $folder $defaultIcon

    $rootKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalMenu\shell\$index-$guid"
    $command = "wscript `"$launcher`" `"%V\.`" $guid"

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
                 [Parameter(Mandatory = $true)][bool]$elevated) {
    New-Item -Path $rootKey -Force | Out-Null
    New-ItemProperty -Path $rootKey -Name 'Icon' -PropertyType String -Value $icon | Out-Null

    if ($elevated) {
        New-ItemProperty -Path $rootKey -Name 'MUIVerb' -PropertyType String -Value (Invoke-Expression $translations.WindowsTerminalMenuElevated) | Out-Null
        New-ItemProperty -Path $rootKey -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'WindowsTerminalMenuElevated' | Out-Null
        New-ItemProperty -Path $rootKey -Name 'HasLUAShield' -PropertyType String -Value '' | Out-Null
    }
    else {
        New-ItemProperty -Path $rootKey -Name 'MUIVerb' -PropertyType String -Value (Invoke-Expression $translations.WindowsTerminalMenu) | Out-Null
        New-ItemProperty -Path $rootKey -Name 'ExtendedSubCommandsKey' -PropertyType String -Value 'WindowsTerminalMenu' | Out-Null
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
    $context = Get-Content -Path "$PSScriptRoot\translations.ini" # Get file translations.
    $context -replace("#.*", "") # Delete comments.
    $language = (Get-ItemProperty 'Registry::HKEY_CURRENT_USER\Control Panel\Desktop' PreferredUILanguages).PreferredUILanguages[0] # Get system language.
    $flag = $false # Use to determine if translations corresponding to the system language has been found.

    # Parse file contents.
    do {
        for ($index = 1; $index -lt $context.Count; ++ $index) {
            if ($context[$index] -match "^\[.+\]") {
                if ($context[$index].Equals("[$language]")) {
                    $flag = $true
                } elseif ($flag) {
                    return
                }
            } elseif ($flag -and ($context[$index] -match "^\w+=.+")) {
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

Write-Host (Invoke-Expression $translations.InstallingWindowsTerminalContextMenu)

$info = GetInstallationInfo
$edition = $info[0]
$folder = $info[1]
$profiles = GetActiveProfiles $edition
$icon = "$PSScriptRoot\icon.ico"

$storage = "$Env:LocalAppData\WindowsTerminalMenuContext"
if (-not (Test-Path $storage)) {
    New-Item -Path $storage -ItemType Directory | Out-Null
}

CreateMenus $storage $icon
Copy-Item "$PSScriptRoot\launch.vbs" "$storage\launch.vbs"
for ($index = 0; $index -lt $profiles.Count; ++ $index) {
    AddProfileMenuItem $profiles[$index] $index $folder $icon "$storage\launch.vbs"
}
Write-Host (Invoke-Expression $translations.InstalledSuccessfully)
exit 0
