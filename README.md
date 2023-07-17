[English](#windows-terminal-context-menu) | [简体中文](#windows-terminal-上下文菜单)

# Windows Terminal Context Menu

This software is licensed under MIT licence.

## 1.Introducing

This software can automatically generate and open the right-click menu of Windows Terminal here.

The script can automatically recognize the system language.

- Currently supported languages: see [translations file](translations.ini).

## 2.Install/Uninstall

Note: These scripts require new [PowerShell](https://github.com/PowerShell/PowerShell) (version at least 6).

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

## 3.Thanks

Special thanks to lextm. This project is modified on lextm's ["windowsterminal shell"](https://github.com/lextm/windowsterminal-shell).

## 4.Compare with "windowsterminal shell"

| | Windows Terminal Context Menu | windowsterminal shell |
| :- | :-: | :-: |
| Supported multiple languages | Yes | No |
| Supported Windows Terminal Preview | Yes | Yes |
| Not required administrator privileges | Yes | No |
| Supported multiple styles | No | Yes |
| Sort menus in the order of settings | Yes | No |
| Supported shortcut keys | Yes | No |

---

# Windows Terminal 上下文菜单

此软件遵循麻省理工学院许可证（MIT）。

## 1.介绍

本软件可在此自动生成在此处打开Windows Terminal的右键菜单。

脚本会自动识别系统语言。

- 当前支持的语言：请查看[翻译文件](translations.ini)。

## 2.安装/卸载

注意：这些脚本需要新的[PowerShell](https://github.com/PowerShell/PowerShell) （版本至少为6）。

首先，克隆这个仓库：

```powershell
git clone https://github.com/sharedwonder/windows-terminal-context-menu.git
```

- 或者下载存根。

如果你的系统不允许运行未签名的PowerShell脚本，执行以下命令以允许当前PowerShell会话执行未签名脚本：

```powershell
Set-ExecutionPolicy -Scope Process Bypass
```

- 安装：运行install.ps1
- 卸载：运行uninstall.ps1

## 3.感谢

特别感谢lextm，此项目修改于lextm的[“windowsterminal shell”](https://github.com/lextm/windowsterminal-shell)。

## 4.与"windowsterminal shell"比较

| | Windows Terminal 上下文菜单 | windowsterminal shell |
| :- | :-: | :-: |
| 支持多种语言 | 是 | 否 |
| 支持Windows Terminal预览版 | 是 | 是 |
| 不需要管理员权限 | 是 | 否 |
| 支持多种样式 | 否 | 是 |
| 按设置的顺序排序菜单 | 是 | 否 |
| 支持快捷键 | 是 | 否 |
