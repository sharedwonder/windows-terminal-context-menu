function GetProgramFilesFolder() {
    return "$Env:ProgramFiles\WindowsApps\" + (Get-ChildItem "Registry::HKEY_CLASSES_ROOT\ActivatableClasses\Package\Microsoft.WindowsTerminal_*_*__8wekyb3d8bbwe").Name.Split("\")[-1]
}

function GetActiveProfiles() {
    $settings = Get-Content "$Env:LocalAppData\Packages\Microsoft.WindowsTerminal_8wekyb3d8bbwe\LocalState\settings.json" | Out-String | ConvertFrom-Json

    if ($settings.profiles.PSObject.Properties.name -match "list") {
        $list = $settings.profiles.list
    }
    else {
        $list = $settings.profiles
    }

    return $list | Where-Object {-not $_.hidden} | Where-Object {($null -eq $_.source) -or -not ($settings.disabledProfileSources -contains $_.source)}
}

function ConvertToIcon([Parameter(Mandatory = $true)][string]$File, [Parameter(Mandatory = $true)][string]$OutputFile) {
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
    $iconWriter.Write([int](6 + 16))
    $iconWriter.Write($memoryStream.ToArray())

    $iconWriter.Flush()
    $output.Close()

    $memoryStream.Dispose()
    $newBitmap.Dispose()
    $inputBitmap.Dispose()
}

function GetProfileIcon([Parameter(Mandatory = $true)]$profile, [Parameter(Mandatory = $true)][String]$folder, [Parameter(Mandatory = $true)][String]$defaultIcon) {
    if ($null -eq $profile.icon) {
        if ($profile.source -eq "Windows.Terminal.Wsl") {
            $guid = "{9acb9455-ca41-5af7-950f-6bca1bc9722f}"
        } else {
            $guid = $profile.guid
        }

        $profilePng = "$folder\ProfileIcons\$guid.scale-200.png"
        $cache = "$Env:LocalAppData\WindowsTerminalIconsCache"
        $icon = "$cache\$guid.ico"

        if (Test-Path $profilePng) {
            if (-not (Test-Path $cache)) {
                New-Item -Path $cache -ItemType Directory
            }

            ConvertToIcon $profilePng $icon
            return $icon
        } else {
            return $defaultIcon
        }
    } else {
        return $profile.icon
    }
}

function AddProfileMenuItem([Parameter(Mandatory = $true)]$profile, [Parameter(Mandatory = $true)][String]$folder, [Parameter(Mandatory = $true)][String]$defaultIcon) {
    $guid = $profile.guid
    $name = $profile.name
    $icon = GetProfileIcon $profile $folder $defaultIcon

    $rootKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalMenu\shell\$guid"
    $command = "wscript `"$PSScriptRoot\launch.vbs`" `"%V\.`" $guid"

    New-Item -Path $rootKey -Force | Out-Null
    New-ItemProperty -Path $rootKey -Name 'MUIVerb' -PropertyType String -Value $name | Out-Null
    New-ItemProperty -Path $rootKey -Name 'Icon' -PropertyType String -Value $icon | Out-Null
    New-Item -Path "$rootKey\command" -Force | Out-Null
    New-ItemProperty -Path "$rootKey\command" -Name '(Default)' -PropertyType String -Value $command | Out-Null

    $rootKey = "Registry::HKEY_CURRENT_USER\SOFTWARE\Classes\WindowsTerminalMenuElevated\shell\$guid"
    $command = "wscript `"$PSScriptRoot\launch.vbs`" `"%V\.`" $guid elevated"

    New-Item -Path $rootKey -Force | Out-Null
    New-ItemProperty -Path $rootKey -Name 'MUIVerb' -PropertyType String -Value $name | Out-Null
    New-ItemProperty -Path $rootKey -Name 'Icon' -PropertyType String -Value $icon | Out-Null
    New-ItemProperty -Path $rootKey -Name 'HasLUAShield' -PropertyType String -Value '' | Out-Null
    New-Item -Path "$rootKey\command" -Force | Out-Null
    New-ItemProperty -Path "$rootKey\command" -Name '(Default)' -PropertyType String -Value $command | Out-Null
}

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

function GetTranslations() {
    $context = Get-Content -Path "$PSScriptRoot\lang.ini"
    $language = (Get-ItemProperty 'Registry::HKEY_CURRENT_USER\Control Panel\Desktop' PreferredUILanguages).PreferredUILanguages[0]
    $flag = $false

    do {
        for ($index = 1; $index -lt $context.Count; ++ $index) {
            if ($context[$index] -match "\[.+\]") {
                if ($context[$index].Equals("[language]".Replace("language", $language))) {
                    $flag = $true
                }
                elseif ($flag) {
                    return
                }
            }
            elseif ($flag) {
                ConvertFrom-StringData -StringData $context[$index]
            }
        }
        if (-not $flag) {
            $language = "en-US"
        }
    } while (-not $flag)
}

function Main() {
    $folder = GetProgramFilesFolder
    $icon = "$PSScriptRoot\icon.ico"
    $translations = GetTranslations
    $profiles = GetActiveProfiles
    CreateMenus $icon $translations
    foreach ($profile in $profiles) {
        Write-Output $profile
        AddProfileMenuItem $profile $folder $icon
    }
}

Main
