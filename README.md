[English](#windows-terminal-context-menu) | [简体中文](#windows-terminal-上下文菜单)

# Windows Terminal Context Menu

This software is licensed under MIT licence.

## Introducing

This software can automatically generate `Terminal here` right-click menus for Windows Terminal.

## Install

Note: The scripts require [new PowerShell](https://github.com/PowerShell/PowerShell) (version at least 6).

First, clone this repository:

```powershell
git clone https://github.com/sharedwonder/windows-terminal-context-menu.git
```

- or download [the archive](https://github.com/sharedwonder/windows-terminal-context-menu/archive/main.zip).

Then open a PowerShell session and run `install.ps1`.

*If your system configuration does not allow executing unsigned PowerShell scripts, execute the following command to allow the current PowerShell session execute unsignified scripts:*

```powershell
Set-ExecutionPolicy -Scope Process Bypass
```

## Uninstall

Run `uninstall.ps1`.

## Script Parameters

The parameters used for `install.ps1`:

`-Layout` : Indicates the menus layout

- `Folded`: The profiles are in the tier-2 menu (default)
- `Unfolded`: The profiles are in the tier-1 menu
- `Mini`: Mini layout with only two items (the standard permissions and the administrator permissions, run with the default profile)

`-Extended`: Whether the menus are extended

- `No`: No (default)
- `Yes`: Yes
- `AdminOnly`: Only the items that run the terminal as administrator

`-SelectedEdition`: Specifies the edition

- `Inquiry`: Ask the user when found 2 editions (default)
- `Stable`: Stable edition
- `Preview`: Preview edition

`-Language`: Specifies which language to use in the menu, if not specified, the system language will be used

## Thanks

Special thanks to [lextm](https://github.com/lextm). This project is modified on lextm's [windowsterminal-shell](https://github.com/lextm/windowsterminal-shell).

---

# Windows Terminal 上下文菜单

此软件以 MIT Licence 许可。

## 介绍

该软件可以为 Windows 终端自动生成 `在此处打开终端` 的右键菜单。

## 安装

注意：脚本需要[新的 PowerShell](https://github.com/PowerShell/PowerShell)（版本至少为 6）。

首先，克隆这个仓库：

```powershell
git clone https://github.com/sharedwonder/windows-terminal-context-menu.git
```

- 或者下载[存档](https://github.com/sharedwonder/windows-terminal-context-menu/archive/main.zip)。

然后打开 PowerShell 会话运行 `install.ps1`。

*如果你的系统配置不允许执行未签名的 PowerShell 脚本，执行以下命令以允许当前 PowerShell 会话执行未签名脚本：*

```powershell
Set-ExecutionPolicy -Scope Process Bypass
```

## 卸载

运行 `uninstall.ps1`。

## 脚本参数

用于 `install.ps1` 的参数：

`-Layout`：菜单布局

- `Folded`：配置文件在二级菜单（默认）
- `Unfolded`：配置文件在一级菜单
- `Mini`：迷你布局，只有两项（标准权限和管理员权限，以默认配置文件运行）

`-Extended`：是否为扩展菜单

- `No`：否（默认）
- `Yes`：是
- `AdminOnly`：仅以管理员身份运行的条目

`-SelectedEdition`：指定版本

- `Inquiry`：当找到 2 个版本时询问用户（默认）
- `Stable`：正式版
- `Preview`：预览版

`-Language`：指定要在菜单中使用的语言，如果未指定，将使用系统语言

## 感谢

特别感谢 [lextm](https://github.com/lextm)，此项目修改于 lextm 的 [windowsterminal-shell](https://github.com/lextm/windowsterminal-shell)。
