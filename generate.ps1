#Requires -Version 6

param (
    [ValidateSet("Folded", "Unfolded", "Minimal")]
    [string]$Layout = "Folded",

    [ValidateSet("No", "Yes", "Admin")]
    [string]$Extended = "No",

    [switch]$NoAdmin = $false,

    [string]$Language = (Get-UICulture).IetfLanguageTag
)

function GetInstallationInfo() {
    $edition = 0

    if ($null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminal))) {
        $installDir = $appx.InstallLocation
        Write-Host "Found Windows Terminal (Version $($appx.Version)): $installDir"
        $edition = 1
    }

    if ($edition -eq 0 -and $null -ne ($appx = (Get-AppxPackage Microsoft.WindowsTerminalPreview))) {
        $installDir = $appx.InstallLocation
        Write-Host "Found Windows Terminal Preview (Version $($appx.Version)): $installDir"
        $edition = 2
    }

    if ($edition -eq 0) {
        Write-Error "Not installed (the specified edition of) Windows Terminal."
        exit 1
    }

    return $edition, $installDir
}

function GetTranslations() {
    $content = Get-Content -Path "$PSScriptRoot\translations.json" | ConvertFrom-Json -AsHashtable
    $result = $null

    foreach ($key in $content.Keys) {
        if ($Language -match $key) {
            $result = $content[$key]
            break
        }
    }

    if ($null -eq $result) {
        $result = $content["^en"]
        Write-Warning "There is no translation corresponding to the specified/system language, so English is used."
    }

    return $result
}

function GenerateLaunchScript() {
    if ($Layout -ne "Minimal") {
        Write-Output @"
If WScript.Arguments.Count > 1 Then
    Set shell = WScript.CreateObject("Shell.Application")
    dir = WScript.Arguments(0)
    profile = WScript.Arguments(1)
    If WScript.Arguments.Count = 2 Then
        shell.ShellExecute "wt", "-p " & profile & " -d """ & dir & """", "", "", 1
    ElseIf WScript.Arguments(2) = "-elevated" Then
        shell.ShellExecute "wt", "-p " & profile & " -d """ & dir & """", "", "RunAs", 1
    End If
End If
"@ > "$Storage\launch.vbs"
    } else {
        Write-Output @"
If WScript.Arguments.Count > 0 Then
    Set shell = WScript.CreateObject("Shell.Application")
    dir = WScript.Arguments(0)
    If WScript.Arguments.Count = 1 Then
        shell.ShellExecute "wt", " -d """ & dir & """", "", "", 1
    ElseIf WScript.Arguments(1) = "-elevated" Then
        shell.ShellExecute "wt", " -d """ & dir & """", "", "RunAs", 1
    End If
End If
"@ > "$Storage\launch.vbs"
    }
}

function GetTerminalSettings() {
    if ($Edition -eq 1) {
        $settingFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json"
    } else {
        $settingFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\settings.json"
    }

    if (-not (Test-Path $settingFile)) {
        Write-Error "The setting file isn't exist."
        exit 1
    }

    return Get-Content $settingFile | Out-String | ConvertFrom-Json
}

function ExtractTerminalIcon() {
    $iconFile = "$Storage\terminal.ico"
    $tempPng = "$env:TEMP\temp.png"
    [System.Drawing.Icon]::ExtractAssociatedIcon("$InstallDir\WindowsTerminal.exe").ToBitmap().Save($tempPng)
    ConvertToIcon $tempPng $iconFile
    Remove-Item $tempPng
    return $iconFile
}

function GetActiveProfiles() {
    if ($Settings.profiles.PSObject.Properties.name -eq "list") {
        $list = $Settings.profiles.list
    } else {
        $list = $Settings.profiles
    }

    return $list | Where-Object { -not $_.hidden }
}

function ConvertToIcon([Parameter(Mandatory)][string]$file, [Parameter(Mandatory)][string]$outputFile) {
    $inputBitmap = [Drawing.Image]::FromFile($file)

    $width = $inputBitmap.Width
    $height = $inputBitmap.Height

    $newBitmap = [Drawing.Bitmap]::new($inputBitmap, $width, $height)
    $memoryStream = New-Object System.IO.MemoryStream
    $newBitmap.Save($memoryStream, [System.Drawing.Imaging.ImageFormat]::Png)

    if ($width -gt 255 -or $height -gt 255) {
        $ratio = ($height, $width | Measure-Object -Maximum).Maximum / 255
        $width /= $ratio
        $height /= $ratio
    }

    $output = [IO.File]::Create($outputFile)
    $iconWriter = [System.IO.BinaryWriter]::new($output)
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
            $iconFile = $InstallDir + "\" + ($wtProfile.icon -replace ("ms-appx:///", "") -replace ("/", "\"))
            if (-not ($iconFile -match "^.*\.scale-.*\.png$")) {
                $iconFile = $iconFile -replace ("\.png$", ".scale-200.png")
            }
        } elseif ($wtProfile.icon -match "^ms-appdata:///Local/.*") {
            if ($Edition -eq 1) {
                $iconFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\" +
                    ($wtProfile.icon -replace ("ms-appdata:///Local/", "") -replace ("/", "\"))
            } else {
                $iconFile = "$env:LocalAppData\Packages\Microsoft.WindowsTerminalPreview_8wekyb3d8bbwe\LocalState\" +
                    ($wtProfile.icon -replace ("ms-appdata:///Local/", "") -replace ("/", "\"))
            }
        } elseif ($wtProfile.icon -match "^ms-appdata:///Roaming/.*") {
            if ($Edition -eq 1) {
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
            $iconFile = "$InstallDir\ProfileIcons\{9acb9455-ca41-5af7-950f-6bca1bc9722f}.scale-200.png"
        } elseif ($wtProfile.source -eq "Git") {
            $gitIcon = Convert-Path ((Get-Command git).Path + "\..\..\mingw64\share\git\git-for-windows.ico")
            if (-not (Test-Path $gitIcon)) {
                $gitIcon = Convert-Path ((Get-Command git).Path + "\..\..\mingw32\share\git\git-for-windows.ico")
            }
            $iconFile = $gitIcon
        } else {
            $iconFile = "$InstallDir\ProfileIcons\$guid.scale-200.png"
        }
    }

    if (Test-Path $iconFile) {
        if ($iconFile -match "\.ico$") {
            return $iconFile
        } elseif ($iconFile -match "\.png$") {
            $iconCopy = "$Storage\icon-$guid.ico"
            ConvertToIcon $iconFile $iconCopy
            return $iconCopy
        } elseif ($iconFile -match "\.exe$") {
            $tempPng = "$env:TEMP\temp.png"
            [System.Drawing.Icon]::ExtractAssociatedIcon($iconFile).ToBitmap().Save($tempPng)
            $iconCopy = "$Storage\icon-$guid.ico"
            ConvertToIcon $tempPng $iconCopy
            Remove-Item $tempPng
            return $iconCopy
        } else {
            Write-Warning "Unknown icon file type: $iconFile"
            return $TerminalIcon
        }
    } else {
        Write-Warning "Icon file not found: $iconFile"
        return $TerminalIcon
    }
}

function AddMenuItemForProfile([Parameter(Mandatory)]$wtProfile, [Parameter(Mandatory)][int]$index) {
    $guid = $wtProfile.guid
    $name = $wtProfile.name

    Write-Host "Profile ${guid}: $name"

    $digits = 1
    for ($x = [math]::Floor($index / 10); $x -ne 0; $x = [math]::Floor($x / 10)) {
        ++$digits
    }
    $order = "0" * (10 - $digits) + $index

    $icon = GetProfileIcon $wtProfile
    if ($Layout -ne "Folded" -or $index -ge 36) {
        $display = $name
    } elseif ($index -ge 10) {
        $display = "$name (&$([char]($index - 10 + 65)))"
    } elseif ($index -eq 9) {
        $display = "$name (&0)"
    } else {
        $display = "$name (&$([char]($index + 1 + 48)))"
    }

    if ($Layout -eq "Folded") {
        $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenu\shell\$order-$guid"
    } elseif ($Layout -eq "Unfolded") {
        $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-$order"
    }
    $command = "wscript `"$Storage\launch.vbs`" `"%V\.`" $guid"

    New-Item -Path $key -Force | Out-Null
    if ($Layout -eq "Unfolded") {
        New-ItemProperty -Path $key -Name "(Default)" -PropertyType String -Value $display | Out-Null
    } else {
        New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value $display | Out-Null
    }
    New-ItemProperty -Path $key -Name "Icon" -PropertyType String -Value $icon | Out-Null
    New-Item -Path "$key\command" | Out-Null
    New-ItemProperty -Path "$key\command" -Name "(Default)" -PropertyType String -Value $command | Out-Null

    if ($Layout -eq "Unfolded") {
        Copy-Item -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-$order" `
            "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu-$order" | Out-Null

        Copy-Item -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-$order" `
            "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu-$order" | Out-Null

        Copy-Item -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-$order" `
            "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu-$order" | Out-Null
    }

    if (-not $NoAdmin) {
        if ($Layout -eq "Folded") {
            $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalContextMenu-Elevated\shell\$order-$guid"
        } elseif ($Layout -eq "Unfolded") {
            $key = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-Elevated-$order"
        }
        $command = "wscript `"$Storage\launch.vbs`" `"%V\.`" $guid -elevated"

        New-Item -Path $key -Force | Out-Null
        if ($Layout -eq "Unfolded") {
            New-ItemProperty -Path $key -Name "(Default)" -PropertyType String -Value ("$display" + $Translations["admin-postfix"]) | Out-Null
        } else {
            New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value $display | Out-Null
        }
        New-ItemProperty -Path $key -Name "Icon" -PropertyType String -Value $icon | Out-Null
        New-ItemProperty -Path $key -Name "HasLUAShield" -PropertyType String -Value "" | Out-Null
        New-Item -Path "$key\command" | Out-Null
        New-ItemProperty -Path "$key\command" -Name "(Default)" -PropertyType String -Value $command | Out-Null

        if ($Layout -eq "Unfolded") {
            Copy-Item -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-Elevated-$order" `
                "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu-Elevated-$order" | Out-Null

            Copy-Item -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-Elevated-$order" `
                "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu-Elevated-$order" | Out-Null

            Copy-Item -Recurse "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-Elevated-$order" `
                "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu-Elevated-$order" | Out-Null
        }
    }
}

function AddMenu([Parameter(Mandatory)][string]$key, [Parameter(Mandatory)][int]$elevated) {
    New-Item -Path $key -Force | Out-Null
    New-ItemProperty -Path $key -Name "Icon" -PropertyType String -Value $TerminalIcon | Out-Null

    if (-not $elevated) {
        New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value ($Translations["standard"]) | Out-Null
        if ($Extended -eq "Yes") {
            New-ItemProperty -Path $key -Name "Extended" -PropertyType String -Value "" | Out-Null
        }
        if ($Layout -eq "Folded") {
            New-ItemProperty -Path $key -Name "ExtendedSubCommandsKey" -PropertyType String -Value "WindowsTerminalContextMenu" | Out-Null
        } elseif ($Layout -eq "Minimal") {
            New-Item -Path "$key\command" | Out-Null
            New-ItemProperty -Path "$key\command" -Name "(Default)" -PropertyType String -Value "wscript `"$Storage\launch.vbs`" `"%V\.`"" | Out-Null
        }
    } else {
        New-ItemProperty -Path $key -Name "MUIVerb" -PropertyType String -Value ($Translations["admin"]) | Out-Null
        New-ItemProperty -Path $key -Name "HasLUAShield" -PropertyType String -Value "" | Out-Null
        if ($Extended -eq "Yes" -or $Extended -eq "Admin") {
            New-ItemProperty -Path $key -Name "Extended" -PropertyType String -Value "" | Out-Null
        }
        if ($Layout -eq "Folded") {
            New-ItemProperty -Path $key -Name "ExtendedSubCommandsKey" -PropertyType String -Value "WindowsTerminalContextMenu-Elevated" | Out-Null
        } elseif ($Layout -eq "Minimal") {
            New-Item -Path "$key\command" | Out-Null
            New-ItemProperty -Path "$key\command" -Name "(Default)" -PropertyType String -Value "wscript `"$Storage\launch.vbs`" `"%V\.`" -elevated" | Out-Null
        }
    }
}

function CreateMenus() {
    if ($Layout -ne "Unfolded") {
        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu" $false
        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu" $false
        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu" $false
        AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu" $false

        if (-not $NoAdmin) {
            AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\shell\WindowsTerminalContextMenu-Elevated" $true
            AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Directory\Background\shell\WindowsTerminalContextMenu-Elevated" $true
            AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\Drive\shell\WindowsTerminalContextMenu-Elevated" $true
            AddMenu "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\LibraryFolder\Background\shell\WindowsTerminalContextMenu-Elevated" $true
        }
    }

    if ($Layout -ne "Minimal") {
        $wtProfiles = GetActiveProfiles
        for ($index = 0; $index -lt $wtProfiles.Count; ++$index) {
            AddMenuItemForProfile $wtProfiles[$index] $index
        }
    }
}

$Storage = "$env:LocalAppData\WindowsTerminalContextMenu"
if (Test-Path "$Storage\layout") {
    Write-Error "The context menus already exists"
    exit 1
}
if (-not (Test-Path $Storage)) {
    New-Item -Path $Storage -ItemType Directory | Out-Null
}

$Edition, $InstallDir = GetInstallationInfo
$Settings = GetTerminalSettings
$Translations = GetTranslations
$TerminalIcon = ExtractTerminalIcon
Write-Output $Layout > "$Storage\layout"
GenerateLaunchScript
CreateMenus

Write-Host "Done"

exit 0
