[English](#windows-terminal-context-menu) | [简体中文](#windows-terminal-上下文菜单)

# Windows Terminal Context Menu

This software is licensed under MIT licence.

## 1.Introducing

This software can automatically generate and open the right-click menu of Windows Terminal here.

The script can automatically recognize the system language.

- Currently supported languages: see [the translation file](translations.ini).

## 2.Install/Uninstall

Note: These scripts require [new PowerShell](https://github.com/PowerShell/PowerShell) (version at least 6).

First, clone this repository:

```powershell
git clone https://github.com/sharedwonder/windows-terminal-context-menu.git
```

- or download the archive.

If your system does not allow executing unsigned PowerShell scripts, execute the following command to allow the current PowerShell session execute unsignified scripts:

```powershell
Set-ExecutionPolicy -Scope Process Bypass
```

- To install: run install.ps1
- To uninstall: run uninstall.ps1

## 3.Parameters

The parameters used for install.ps1

`-Layout` : Indicates the menus layout

- `Folded`: The profiles are in the tier-2 menu (default)
- `Unfolded`: The profiles are in the tier-1 menu
- `Mini`: Mini layout with only two items (the standard permissions and the administrator permissions, run with the default profile)

`-Extended`: Whether the menus are extended

- `No`: No (default)
- `Yes`: Yes
- `AdminOnly`: Only the menus that run the terminal as administrator

`-SelectedEdition`: Selected version

- `Inquiry`: Ask the user (default)
- `Stable`: Stable version
- `Preview`: Preview version

## 4.Thanks

Special thanks to lextm. This project is modified on lextm's [windowsterminal-shell](https://github.com/lextm/windowsterminal-shell).

## 5.Improvements over the windowsterminal-shell

This script supports multiple languages, the language file contains several languages currently, if there is no language that you use, just add it by the language code (if you can, please [submit a pull requset on GitHub](https://github.com/sharedwonder/windows-terminal-context-menu/pulls) for me to add a language).

Functionally, this script sorts the menu items by the order of profiles, supports shortcut keys (only 'Folded' layout, first 36 items).

And it uses the PowerShell cmdlet to get the installation directory of Windows Terminal directly without the administrator permissions.

---

# Windows Terminal 上下文菜单

此软件根据麻省理工学院许可证（MIT Licence）许可使用。

## 1.介绍

本软件可在此自动生成在此处打开Windows Terminal的右键菜单。

脚本会自动识别系统语言。

- 当前支持的语言：请查看[翻译文件](translations.ini)。

## 2.安装/卸载

注意：这些脚本需要[新的PowerShell](https://github.com/PowerShell/PowerShell) （版本至少为6）。

首先，克隆这个仓库：

```powershell
git clone https://github.com/sharedwonder/windows-terminal-context-menu.git
```

- 或者下载存根。

如果你的系统不允许执行未签名的PowerShell脚本，执行以下命令以允许当前PowerShell会话执行未签名脚本：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
```

- 安装：运行install.ps1
- 卸载：运行uninstall.ps1

## 3.参数

用于install.ps1的参数

`-Layout`：菜单布局

- `Folded`：配置文件在二级菜单（默认）
- `Unfolded`：配置文件在一级菜单
- `Mini`：迷你布局，只有两项（标准权限和管理员权限，以默认配置文件运行）

`-Extended`：是否为扩展菜单

- `No`：否（默认）
- `Yes`：是
- `AdminOnly`：仅以管理员身份运行终端的菜单

`-SelectedEdition`：选定的版本

- `Inquiry`：询问用户（默认）
- `Stable`：正式版
- `Preview`：预览版

## 4.感谢

特别感谢lextm，此项目修改于lextm的[windowsterminal-shell](https://github.com/lextm/windowsterminal-shell)。

## 5.相比windowsterminal-shell的改进

这个脚本支持用语言文件支持多语言，目前语言文件包含了几种语言，如果没有你使用的语言，只需按语言代码添加即可（如果可以，请[在GitHub上提交一个拉取请求（pull requset）](https://github.com/sharedwonder/windows-terminal-context-menu/pulls)来让我添加语言）。

功能上，这个脚本会按配置文件的顺序排序菜单项目，并且支持快捷键（仅Folded布局，前36项）。

另外它使用PowerShell cmdlet直接获取Windows Terminal的安装目录而无需管理员权限。
