# PowerShell Web Installer for Nerd Fonts

[![Platform Support](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-blue)](#)
[![PowerShell Gallery](https://img.shields.io/powershellgallery/v/Invoke-NerdFontInstaller)](https://www.powershellgallery.com/packages/Invoke-NerdFontInstaller)
[![PowerShell Gallery Downloads](https://img.shields.io/powershellgallery/dt/Invoke-NerdFontInstaller)](https://www.powershellgallery.com/packages/Invoke-NerdFontInstaller)

An interactive installer for [Nerd Fonts](https://www.nerdfonts.com/) and [Cascadia Code](https://github.com/microsoft/cascadia-code) on Windows, macOS, or Linux.

> **TL;DR**: To quickly install Nerd Fonts using the interactive web installer, run the following command in your PowerShell terminal:
>
> ```powershell
> & ([scriptblock]::Create((iwr 'https://to.loredo.me/Install-NerdFont.ps1')))
> ```

![PowerShell Web Installer for Nerd Fonts](https://raw.githubusercontent.com/jpawlowski/nerd-fonts-installer-PS/main/images/nerd-fonts-installer.gif)

The script downloads the font archive from the GitHub release pages and installs the font files to
the user's font directory, or the system font directory when using the AllUsers scope with
elevated permissions.

## Prerequisites

The Web Installer for Nerd Fonts requires PowerShell to be available on your local machine.

- **PowerShell 7+** is highly recommended for the best user experience and latest features.
- The built-in **Windows PowerShell 5.1** is also supported but may lack some features available in PowerShell 7+.

### Why PowerShell 7+?

PowerShell 7+ can be installed on Windows, macOS, or Linux devices, providing a consistent experience across all platforms. It is highly recommended for Windows users as it offers the latest features and improvements over the older Windows PowerShell 5.1.

### Checking Your PowerShell Version

To check which version of PowerShell you have installed, run the following command in your PowerShell terminal:

```powershell
$PSVersionTable.PSVersion
```

### Installation Guides

To learn more about how to install PowerShell for your device, see one of these guides:

- [Installing PowerShell on **Windows**](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows)
- [Installing PowerShell on **macOS**](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-macos)
- [Installing PowerShell on **Linux**](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux)

## How to Use

### `Option 1`: Run the Installer directly from the Web _(Preferred)_

You can conveniently run this script directly from the web to install Nerd Fonts.

#### Run the Interactive Installer

```powershell
& ([scriptblock]::Create((iwr 'https://to.loredo.me/Install-NerdFont.ps1')))
```

Or alternatively without the shortened URL:

```powershell
& ([scriptblock]::Create((iwr 'https://raw.githubusercontent.com/jpawlowski/nerd-fonts-installer-PS/main/Invoke-NerdFontInstaller.ps1')))
```

> **IMPORTANT**: A code signature cannot be verified when running the script directly from the web.
> SSL transport layer encryption is used to protect the script during download from GitHub and during
> redirection from the URL shortener.

#### Script Parameters

You can pass parameters to the script just like any other PowerShell script. Here are some examples:

##### Install Fonts by Name

To install specific fonts by name, use the `-Name` parameter:

```powershell
& ([scriptblock]::Create((iwr 'https://to.loredo.me/Install-NerdFont.ps1'))) -Name hack, heavy-data
```

##### Install Fonts Unattended

To install fonts without any prompts, use the `-Confirm:$false` parameter:

```powershell
& ([scriptblock]::Create((iwr 'https://to.loredo.me/Install-NerdFont.ps1'))) -Confirm:$false -Name hack, heavy-data
```

##### List or Search Font Names

To list all available fonts or search for specific fonts, use the `-List` parameter:

```powershell
& ([scriptblock]::Create((iwr 'https://to.loredo.me/Install-NerdFont.ps1'))) -List All
& ([scriptblock]::Create((iwr 'https://to.loredo.me/Install-NerdFont.ps1'))) -List hack
& ([scriptblock]::Create((iwr 'https://to.loredo.me/Install-NerdFont.ps1'))) -List cas*
```

##### Get Help with Enhanced Parameter Options

To get help with the script and see enhanced parameter options, use the `-Help` parameter:

```powershell
& ([scriptblock]::Create((iwr 'https://to.loredo.me/Install-NerdFont.ps1'))) -Help ?
```

### `Option 2`: Install and Run Locally from PowerShell Gallery

The Nerd Font Web Installer can also be installed locally from the [PowerShell Gallery](https://www.powershellgallery.com/packages/Invoke-NerdFontInstaller):

```powershell
Install-PSResource Invoke-NerdFontInstaller
```

Running the script locally comes with some advantages:

- Use [tab completion](https://learn.microsoft.com/en-us/powershell/scripting/learn/shell/tab-completion) for parameters, including dynamic `-Name` parameter
- Ensuring script integrity by validating its [code signature](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_signing) _(Windows only)_
- Enhanced security by controlling the installation environment and dependencies
