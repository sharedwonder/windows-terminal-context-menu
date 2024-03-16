#Requires -Version 6

param (
    [ValidateSet("Folded", "Unfolded", "Mini")]
    [string]$Layout = "Folded",
    [ValidateSet("No", "Yes", "AdminOnly")]
    [string]$Extended = "No",
    [ValidateSet("Inquiry", "Stable", "Preview")]
    [string]$SelectedEdition = "Inquiry",
    [string]$Language = (Get-Culture).Name
)

function GenerateLaunchScript() {
    Write-Output @"
If Wscript.Arguments.Count > 1 Then
    Set shell = WScript.CreateObject("Shell.Application")
    dir = WScript.Arguments(0)
    profile = WScript.Arguments(1)
    If Wscript.Arguments.Count = 2 Then
        shell.ShellExecute "wt", "-p " & profile & " -d """ & dir & """", "", "", 1
    ElseIf WScript.Arguments(2) = "-elevated" Then
        shell.ShellExecute "wt", "-p " & profile & " -d """ & dir & """", "", "RunAs", 1
    End If
End If
"@ > "$storage\launch.vbs"
}

function ExtractTerminalIcon() {
    if ($edition -eq 1) {
        $terminalIcon = "$storage\terminal.ico"
    } else {
        $terminalIcon = "$storage\terminal-preview.ico"
    }
    $tempPng = "$env:TEMP\temp.png"
    [System.Drawing.Icon]::ExtractAssociatedIcon("$installDir\WindowsTerminal.exe").ToBitmap().Save($tempPng)
    ConvertToIcon $tempPng $terminalIcon
    Remove-Item $tempPng
    return $terminalIcon
}

function GetTranslations() {
    $content = Get-Content -Path "$PSScriptRoot\translations.txt"
    $found = $false

    do {
        for ($index = 0; $index -lt $content.Count; ++$index) {
            if ($content[$index].StartsWith(":")) {
                if ($found) {
                    return
                }
                if ($Language -match $content[$index].Substring(1)) {
                    $found = $true
                }
            } elseif ($found -and ($content[$index] -match "^\w+=.*")) {
                ConvertFrom-StringData -StringData $content[$index]
            }
        }

        if (-not $found) {
            Write-Warning "There is no translation corresponding to the specified/system language, so using English."
            $Language = "en"
        }
    } while (-not $found)
}

function GetInstallationInfo() {
    # Release edition
    if ($SelectedEdition -eq "Inquiry" -or $SelectedEdition -eq "Stable") {
        if ($null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminal))) {
            $installDir = $appx.InstallLocation
            Write-Host "Found Windows Terminal (Version $($appx.Version)): $installDir"
            $edition = 1
        }
    }

    # Preview edition
    if ($SelectedEdition -eq "Inquiry" -or $SelectedEdition -eq "Preview") {
        if ($null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminalPreview))) {
            $installDir = $appx.InstallLocation
            Write-Host "Found Windows Terminal Preview (Version $($appx.Version)): $installDir"
            if ($edition -eq 0) {
                $edition = 2
            } else {
                do {
                    $edition = Read-Host "Select edition [1: stable, 2: preview]"
                } while (($edition -eq 1) -or ($edition -eq 2))

                if ($edition -eq 2) {
                    $installDir = $appx.InstallLocation
                }
            }
        }
    }

    if ($edition -eq 0) {
        Write-Error "Not installed (the specified edition of) Windows Terminal."
        exit 1
    }

    return $edition, $installDir
}

function GetSettingsFile() {
    if ($edition -eq 1) {
        $settingsFile = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
    } else {
        $settingsFile = "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\settings.json"
    }

    if (-not (Test-Path $settingsFile)) {
        Write-Error "The settings file isn't exist."
        exit 1
    }

    return $settingsFile
}

function GetActiveProfiles() {
    $settings = Get-Content $settingsFile | Out-String | ConvertFrom-Json

    if ($settings.profiles.PSObject.Properties.name -match "list") {
        $list = $settings.profiles.list
    } else {
        $list = $settings.profiles
    }

    return $list | Where-Object { -not $_.hidden } | Where-Object { ($null -eq $_.source) -or -not ($settings.disabledProfileSources -contains $_.source) }
}

function ConvertToIcon([Parameter(Mandatory)][string]$file, [Parameter(Mandatory)][string]$outputFile) {
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

function GetProfileIcon([Parameter(Mandatory)]$wtProfile) {
    if ($null -ne $wtProfile.icon) {
        if ($wtProfile.icon -match "^ms-appx:///.*") {
            $iconFile = $installDir + "\" + ($wtProfile.icon -replace ("ms-appx:///", "") -replace ("/", "\"))
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
            $iconFile = "$installDir\ProfileIcons\{9acb9455-ca41-5af7-950f-6bca1bc9722f}.scale-200.png"
        } elseif ($wtProfile.source -eq "Git") {
            $gitIcon = Convert-Path ((Get-Command git).Path + "\..\..\mingw64\share\git\git-for-windows.ico")
            if (-not (Test-Path $gitIcon)) {
                $gitIcon = Convert-Path ((Get-Command git).Path + "\..\..\mingw32\share\git\git-for-windows.ico")
            }
            $iconFile = $gitIcon
        } else {
            $iconFile = "$installDir\ProfileIcons\$guid.scale-200.png"
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
            Write-Warning "Unknown icon file type: $iconFile"
            return $terminalIcon
        }
    } else {
        Write-Warning "Icon file not found: $iconFile"
        return $terminalIcon
    }
}

function AddMenuItemForProfile([Parameter(Mandatory)]$wtProfile, [Parameter(Mandatory)]$index) {
    $guid = $wtProfile.guid
    $name = $wtProfile.name

    Write-Host "Profile ${guid}: $name"

    $digits = 1
    for ($x = [math]::Floor($index / 10); $x -ne 0; $x = [math]::Floor($x / 10)) {
        ++$digits
    }
    $prefix = "0" * (10 - $digits) + $index

    $icon = GetProfileIcon $wtProfile
    if ($Layout -ne "Folded" -or $index -ge 36) {
        $display = $name
    } elseif ($index -ge 10) {
        $display = "&" + [char]($index - 10 + 65) + ". $name"
    } elseif ($index -eq 9) {
        $display = "&0. $name"
    } else {
        $display = "&" + [char]($index + 1 + 48) + ". $name"
    }

    if ($Layout -eq "Folded") {
        $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenu\shell\$prefix-$guid"
    } elseif ($Layout -eq "Unfolded") {
        $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-$prefix-$guid"
    }
    $command = "wscript `"$storage\launch.vbs`" `"%V\.`" $guid"

    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value $display | Out-Null
    New-ItemProperty -Path $key -Name "Icon" -PropertyType String -Value $icon | Out-Null
    New-Item -Path "$key\command" -Force | Out-Null
    New-ItemProperty -Path "$key\command" -Name "(Default)" -PropertyType String -Value $command | Out-Null

    if ($Layout -eq "Folded") {
        $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenuElevated\shell\$prefix-$guid"
    } elseif ($Layout -eq "Unfolded") {
        $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenuElevated-$prefix-$guid"
    }
    $command = "wscript `"$storage\launch.vbs`" `"%V\.`" $guid -elevated"

    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value $display | Out-Null
    New-ItemProperty -Path $key -Name "Icon" -PropertyType String -Value $icon | Out-Null
    New-ItemProperty -Path $key -Name "HasLUAShield" -PropertyType String -Value "" | Out-Null
    New-Item -Path "$key\command" -Force | Out-Null
    New-ItemProperty -Path "$key\command" -Name "(Default)" -PropertyType String -Value $command | Out-Null

    if ($Layout -eq "Unfolded") {
        Copy-Item -Force -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-$prefix-$guid" `
            "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu-$prefix-$guid" | Out-Null
        Copy-Item -Force -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenuElevated-$prefix-$guid" `
            "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenuElevated-$prefix-$guid" | Out-Null

        Copy-Item -Force -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-$prefix-$guid" `
            "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu-$prefix-$guid" | Out-Null
        Copy-Item -Force -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenuElevated-$prefix-$guid" `
            "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenuElevated-$prefix-$guid" | Out-Null

        Copy-Item -Force -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-$prefix-$guid" `
            "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu-$prefix-$guid" | Out-Null
        Copy-Item -Force -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenuElevated-$prefix-$guid" `
            "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenuElevated-$prefix-$guid" | Out-Null
    }
}

function AddMenu([Parameter(Mandatory)][String]$key, [Parameter(Mandatory)][bool]$elevated) {
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name "Icon" -PropertyType String -Value $terminalIcon | Out-Null

    if ($Layout -eq "Mini") {
        $defaultProfile = (Get-Content $settingsFile | Out-String | ConvertFrom-Json).defaultProfile
    }

    if (-not $elevated) {
        New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value (Invoke-Expression $translations.a) | Out-Null
        if ($Extended -eq "Yes") {
            New-ItemProperty -Path $key -Name "Extended" -PropertyType String -Value "" | Out-Null
        }
        if ($Layout -eq "Folded") {
            New-ItemProperty -Path $key -Name "ExtendedSubCommandsKey" -PropertyType String -Value "WindowsTerminalContextMenu" | Out-Null
        } elseif ($Layout -eq "Mini") {
            New-Item -Path "$key\command" -Force | Out-Null
            New-ItemProperty -Path "$key\command" -Name "(Default)" -PropertyType String -Value "wscript `"$storage\launch.vbs`" `"%V\.`" $defaultProfile" | Out-Null
        }
    } else {
        New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value (Invoke-Expression $translations.b) | Out-Null
        New-ItemProperty -Path $key -Name "HasLUAShield" -PropertyType String -Value "" | Out-Null
        if ($Extended -eq "Yes" -or $Extended -eq "AdminOnly") {
            New-ItemProperty -Path $key -Name "Extended" -PropertyType String -Value "" | Out-Null
        }
        if ($Layout -eq "Folded") {
            New-ItemProperty -Path $key -Name "ExtendedSubCommandsKey" -PropertyType String -Value "WindowsTerminalContextMenuElevated" | Out-Null
        } elseif ($Layout -eq "Mini") {
            New-Item -Path "$key\command" -Force | Out-Null
            New-ItemProperty -Path "$key\command" -Name "(Default)" -PropertyType String -Value "wscript `"$storage\launch.vbs`" `"%V\.`" $defaultProfile" | Out-Null
        }
    }
}

function CreateMenus() {
    if ($Layout -ne "Unfolded") {
        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu" $false
        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenuElevated" $true

        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu" $false
        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenuElevated" $true

        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu" $false
        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenuElevated" $true

        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu" $false
        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenuElevated" $true
    }

    if ($Layout -ne "Mini") {
        $wtProfiles = GetActiveProfiles
        for ($index = 0; $index -lt $wtProfiles.Count; ++$index) {
            AddMenuItemForProfile $wtProfiles[$index] $index
        }
    }
}

Write-Host "Installing..."

$storage = "$env:LOCALAPPDATA\WindowsTerminalContextMenu"
if (-not (Test-Path $storage)) {
    New-Item -Path $storage -ItemType Directory | Out-Null
}

$info = GetInstallationInfo
$edition = $info[0]
$installDir = $info[1]
$settingsFile = GetSettingsFile

Write-Output $Layout > "$storage\layout.txt"
$terminalIcon = ExtractTerminalIcon
GenerateLaunchScript
$translations = GetTranslations
CreateMenus

Write-Host "Done."

exit 0
