#Requires -Version 6

# Generates the launch script.
function GenerateLaunchScript() {
    Write-Output @"
If Wscript.Arguments.Count > 1 Then
    Set shell = WScript.CreateObject("Shell.Application")
    folder = WScript.Arguments(0)
    profile = WScript.Arguments(1)
    If Wscript.Arguments.Count = 2 Then
        shell.ShellExecute "wt", "-p " & profile & " -d """ & folder & """", "", "", 1
    ElseIf WScript.Arguments(2) = "-elevated" Then
        shell.ShellExecute "wt", "-p " & profile & " -d """ & folder & """", "", "RunAs", 1
    End If
End If
"@ > "$storage\launch.vbs"
}

# Gets the edition and the installation folder of Windows Terminal.
function GetInstallationInfo() {
    Write-Host (Invoke-Expression $translations.LookingForWindowsTerminal)

    $edition = 0

    if ($null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminal))) {
        # release edition
        $version = $appx.Version
        Write-Host (Invoke-Expression $translations.FoundWindowsTerminal)
        $folder = $appx.InstallLocation
        $edition = 1
        $selectVersion = $version
    }

    if ($null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminalPreview))) {
        # preview edition
        $version = $appx.Version
        Write-Host (Invoke-Expression $translations.FoundWindowsTerminalPreview)
        if ($edition -eq 0) {
            $edition = 2
            $folder = $appx.InstallLocation
            $selectVersion = $version
        } else {
            # Found multiple editions.
            do {
                Write-Host (Invoke-Expression $translations.SelectEdition)
                $edition = Read-Host
            } while (($edition -eq 1) -or ($edition -eq 2))

            if ($edition -eq 2) {
                $folder = $appx.InstallLocation
                $selectVersion = $version
            }
        }
    }

    # Not found.
    if ($edition -eq 0) {
        Write-Error (Invoke-Expression $translations.NotInstalledWindowsTerminal)
        exit 1
    }
    elseif ($selectVersion -lt "1.0") {
        Write-Warning (Invoke-Expression $translations.WindowsTerminalVersionTooOld)
    }

    Write-Host (Invoke-Expression $translations.WindowsTerminalInstallationFolder)

    return $edition, $folder
}

# Gets active profiles of Windows Terminal.
function GetActiveProfiles() {
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

    if ($settings.profiles.PSObject.Properties.name -match "list") {
        # Old Windows Terminal
        $list = $settings.profiles.list
    } else {
        $list = $settings.profiles
    }

    # Exclude the disabled profiles and return active profiles.
    return $list | Where-Object { -not $_.hidden } | Where-Object { ($null -eq $_.source) -or -not ($settings.disabledProfileSources -contains $_.source) }
}

# Converts PNG to ICO.
function ConvertToIcon([Parameter(Mandatory)][string]$file, [Parameter(Mandatory)][string]$outputFile) {
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

# Gets the icon of the provided profile.
function GetProfileIcon([Parameter(Mandatory)]$wtProfile) {
    if ($null -ne $wtProfile.icon) {
        if ($wtProfile.icon -match "^ms-appx:///.*") {
            $iconFile = $folder + "\" + ($wtProfile.icon -replace ("ms-appx:///", "") -replace ("/", "\"))
            if (-not ($iconFile -match "^.*\.scale-.*\.png$")) {
                $iconFile = $iconFile -replace ("\.png$", ".scale-200.png")
            }
        } elseif ($wtProfile.icon -match "^ms-appdata:///Local/.*") {
            if ($edition -eq 1) {
                $iconFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\" +
                    ($wtProfile.icon -replace ("ms-appdata:///Local/", "") -replace ("/", "\"))
            } else {
                $iconFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\" +
                    ($wtProfile.icon -replace ("ms-appdata:///Local/", "") -replace ("/", "\"))
            }
        } elseif ($wtProfile.icon -match "^ms-appdata:///Roaming/.*") {
            if ($edition -eq 1) {
                $iconFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\RoamingState\" +
                    ($wtProfile.icon -replace ("ms-appdata:///Roaming/", "") -replace ("/", "\"))
            } else {
                $iconFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\RoamingState\" +
                    ($wtProfile.icon -replace ("ms-appdata:///Roaming/", "") -replace ("/", "\"))
            }
        } else {
            $iconFile = [System.Environment]::ExpandEnvironmentVariables($wtProfile.icon)
        }
    } else {
        if ($wtProfile.source -eq "Windows.Terminal.Wsl") {
            $iconFile = "$folder\ProfileIcons\{9acb9455-ca41-5af7-950f-6bca1bc9722f}.scale-200.png"
        }
        elseif ($wtProfile.source -eq "Git") {
            $gitIcon = Convert-Path ((Get-Command git).Path + "\..\..\mingw64\share\git\git-for-windows.ico")

            if (-not (Test-Path $gitIcon)) {
                $gitIcon = Convert-Path ((Get-Command git).Path + "\..\..\mingw32\share\git\git-for-windows.ico")
            }
            $iconFile = $gitIcon
        }
        else {
            $iconFile = "$folder\ProfileIcons\$guid.scale-200.png"
        }
    }

    if (Test-Path $iconFile) {
        if ($iconFile -match ".*\.ico$") {
            return $iconFile
        } elseif ($iconFile -match ".*\.png$") {    
            $newIcon = "$storage\$guid.ico"
            ConvertToIcon $iconFile $newIcon
            return $newIcon
        } elseif ($iconFile -match ".*\.exe$") {
            $tempPng = "$env:TEMP\temp.png"
            $newIcon = "$storage\$guid.ico"
            [System.Drawing.Icon]::ExtractAssociatedIcon($iconFile).ToBitmap().Save($tempPng)
            ConvertToIcon $tempPng $newIcon
            Remove-Item $tempPng
            return $newIcon
        } else {
            Write-Warning (Invoke-Expression $translations.UnknownIconFile)
            return $terminalIcon
        }
    } else {
        Write-Warning (Invoke-Expression $translations.IconFileNotFound)
        return $terminalIcon
    }
}

# Adds a menu subitem of the provided profile.
function AddProfileMenuItem([Parameter(Mandatory)]$wtProfile, [Parameter(Mandatory)]$index) {
    $guid = $wtProfile.guid
    $name = $wtProfile.name
    $icon = GetProfileIcon $wtProfile

    Write-Host (Invoke-Expression $translations.AddingMenuSubitems)

    $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenu\shell\$index-$guid"
    $command = "wscript `"$storage\launch.vbs`" `"%V\.`" $guid"

    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value $name | Out-Null
    New-ItemProperty -Path $key -Name "Icon" -PropertyType String -Value $icon | Out-Null
    New-Item -Path "$key\command" -Force | Out-Null
    New-ItemProperty -Path "$key\command" -Name "(Default)" -PropertyType String -Value $command | Out-Null

    $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenuElevated\shell\$index-$guid"
    $command = "wscript `"$storage\launch.vbs`" `"%V\.`" $guid -elevated"

    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value $name | Out-Null
    New-ItemProperty -Path $key -Name "Icon" -PropertyType String -Value $icon | Out-Null
    New-ItemProperty -Path $key -Name "HasLUAShield" -PropertyType String -Value "" | Out-Null
    New-Item -Path "$key\command" -Force | Out-Null
    New-ItemProperty -Path "$key\command" -Name "(Default)" -PropertyType String -Value $command | Out-Null
}

# Adds a menu that open Windows Terminal.
function AddMenu([Parameter(Mandatory)][String]$key, [Parameter(Mandatory)][bool]$elevated) {
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name "Icon" -PropertyType String -Value $terminalIcon | Out-Null

    if ($elevated) {
        New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value (Invoke-Expression $translations.WindowsTerminalContextMenuElevated) | Out-Null
        New-ItemProperty -Path $key -Name "ExtendedSubCommandsKey" -PropertyType String -Value "WindowsTerminalContextMenuElevated" | Out-Null
        New-ItemProperty -Path $key -Name "HasLUAShield" -PropertyType String -Value "" | Out-Null
    }
    else {
        New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value (Invoke-Expression $translations.WindowsTerminalContextMenu) | Out-Null
        New-ItemProperty -Path $key -Name "ExtendedSubCommandsKey" -PropertyType String -Value "WindowsTerminalContextMenu" | Out-Null
    }
}

# Creates all context menus that open Windows Terminal.
function CreateMenus() {
    # Directory
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu" $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenuElevated" $true

    # Directory background
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu" $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenuElevated" $true

    # Drive
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu" $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenuElevated" $true

    # Library background
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu" $false
    AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenuElevated" $true
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
                }
                elseif ($found) {
                    return
                }
            }
            elseif ($found -and ($context[$index] -match "^\w+=.*")) {
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

Write-Host (Invoke-Expression $translations.InstallingWindowsTerminalContextMenu)

$storage = "$env:LocalAppData\WindowsTerminalContextMenu"
if (-not (Test-Path $storage)) {
    New-Item -Path $storage -ItemType Directory | Out-Null
}

$info = GetInstallationInfo
$edition = $info[0]
$folder = $info[1]
$wtProfiles = GetActiveProfiles
if ($edition -eq 1) {
    $terminalIcon = "$storage\terminal.ico"   
} else {
    $terminalIcon = "$storage\terminal-preview.ico"
}

$tempPng = "$env:TEMP\temp.png"
[System.Drawing.Icon]::ExtractAssociatedIcon("$folder\WindowsTerminal.exe").ToBitmap().Save($tempPng)
ConvertToIcon $tempPng $terminalIcon
Remove-Item $tempPng

CreateMenus
GenerateLaunchScript
for ($index = 0; $index -lt $wtProfiles.Count; ++$index) {
    AddProfileMenuItem $wtProfiles[$index] $index
}

Write-Host (Invoke-Expression $translations.InstalledSuccessfully)

exit 0
