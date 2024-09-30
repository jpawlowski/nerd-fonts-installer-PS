#!/usr/bin/env pwsh

<#PSScriptInfo

.VERSION 1.3.5

.GUID a3238c59-8a0e-4c11-a334-f071772d1255

.AUTHOR Julian Pawlowski

.COPYRIGHT © 2024 Julian Pawlowski.

.TAGS fonts nerdfonts cascadia-code cascadia-code-nerd-font cascadia-code-powerline-font cascadia-mono cascadia-mono-nerd-font cascadia-mono-powerline-font powershell powershell-script Windows MacOS Linux PSEdition_Core PSEdition_Desktop

.LICENSEURI https://github.com/jpawlowski/nerd-fonts-installer-PS/blob/main/LICENSE.txt

.PROJECTURI https://github.com/jpawlowski/nerd-fonts-installer-PS

.ICONURI https://raw.githubusercontent.com/jpawlowski/nerd-fonts-installer-PS/main/images/nerd-fonts-logo.png

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
    Version 1.3.5 (2024-09-27)
    - Fix font location for macOS
#>

<#
.SYNOPSIS
    Install Nerd Fonts on Windows, macOS, or Linux.

    You may also run this script directly from the web using the following command:

    ```powershell
    & ([scriptblock]::Create((iwr 'https://to.loredo.me/Install-NerdFont.ps1')))
    ```

    Parameters may be passed at the end just like any other PowerShell script.

.DESCRIPTION
    An interactive installer for Nerd Fonts and Cascadia Code on Windows, macOS, or Linux.

    The script downloads the font archive from the GitHub release pages and extracts the font files to
    the user's font directory, or the system font directory when using the AllUsers scope with
    elevated permissions.

    Besides installing the script locally, you may also run this script directly from the web
    using the following command:

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

    Parameters may be passed just like any other PowerShell script. For example:

    ```powershell
    & ([scriptblock]::Create((iwr 'https://to.loredo.me/Install-NerdFont.ps1'))) -Name cascadia-code, cascadia-mono
    ```

    To get a list of available Nerd Fonts, use the following command:

    ```powershell
    & ([scriptblock]::Create((iwr 'https://to.loredo.me/Install-NerdFont.ps1'))) -List All
    ```

.PARAMETER Help
    Show help content for the script or a specific parameter.

    The following values are supported:
    - Name: Shows help for the dynamic Name parameter (which is not possible using regular Get-Help)
    - Variant: Shows help for the Variant parameter.
    - Type: Shows help for the Type parameter.
    - All: Shows help for the All parameter.
    - List: Shows help for the List parameter.
    - Scope: Shows help for the Scope parameter.
    - Force: Shows help for the Force parameter.
    - Help: Shows help for the Help parameter.
    - Version: Shows help for the Version parameter.
    - Summary: Provides a summary of the help content. Note that the dynamic Name parameter is not included in the summary.
    - Detailed: Provides detailed help, including parameter descriptions and examples. Note that the dynamic Name parameter is not included in the detailed help.
    - Full: Provides full help, including detailed help, parameter descriptions, examples, and additional notes. Note that the dynamic Name parameter is not included in the full help.
    - Examples: Shows only the examples section of the help.

.PARAMETER Name
    The name of the Nerd Font to install.
    Multiple font names can be specified as an array of strings.

    If no font name is specified, the script provides an interactive menu to select the font to install
    (unless the All parameter is used).

    The menu is displayed only if the script is run in an interactive session.
    If the script is run in a non-interactive environment, the Name parameter is mandatory and must be specified.

    Possible font names are dynamically retrieved from the Nerd Font library on GitHub.
    To see a list of available fonts, use the parameter '-List All'.

.PARAMETER Variant
    Specify the font variant to install.
    The default value is 'Variable'.

    A variable font is a single font file that can contain multiple variations of a typeface while a static font
    is a traditional font file with a single style.
    For example, a variable font can contain multiple weights and styles in a single file.

    Most Nerd Fonts are only available as static fonts. The only exception today is Microsoft's Cascadia Code
    font where variable fonts are recommended to be used for all platforms.

    Setting this parameter to 'Static' will search for a folder with the name 'static' in the font archive.
    If that was found, files from that folder will be installed, otherwise, the script will install the font files
    from the root of the archive (or the folder for the font type, if any is found).

.PARAMETER Type
    Specify the order to search for font types. Only the first matching type will be installed.
    The default order is TTF, OTF, WOFF2.

    The script will search for folders with the specified type name in the font archive, and if found, install the fonts from that folder.
    If no folder is found, the script will install the font files from the root of the archive.

    The script will search for files with the specified type extension in any case.

.PARAMETER All
    Install all available Nerd Fonts.
    You will be prompted to confirm the installation for each font with the option to skip, cancel,
    or install all without further confirmation.

.PARAMETER List
    List available Nerd Fonts matching the specified pattern.
    Use '*' or 'All' to list all available Nerd Fonts.
    This parameter does not install any fonts.

.PARAMETER Scope
    Defined the scope in which the Nerd Font should be installed.
    The default value is CurrentUser.

    The AllUsers scope requires elevated permissions.
    The CurrentUser scope installs the font for the current user only.

.PARAMETER Force
    Overwrite existing font files instead of skipping them.

.PARAMETER Version
    Display the version of the script.

.EXAMPLE
    Invoke-NerdFontInstaller -Name cascadia-code
    Install the Cascadia Code Font Family from the Microsoft repository.

.EXAMPLE
    Invoke-NerdFontInstaller -Name cascadia-mono
    Install the Cascadia Mono Font Family from the Microsoft repository.

.EXAMPLE
    Invoke-NerdFontInstaller -Name cascadia-code, cascadia-mono
    Install the Cascadia Code and Cascadia Mono Font Families from the Microsoft repository.

.EXAMPLE
    Invoke-NerdFontInstaller -All -WhatIf
    Show what would happen if all fonts were installed.

.EXAMPLE
    Invoke-NerdFontInstaller -List cascadia*
    List all fonts with names starting with 'cascadia'.

.EXAMPLE
    Invoke-NerdFontInstaller -Help Name
    Get help for the dynamic Name parameter.

.EXAMPLE
    Invoke-NerdFontInstaller -Help ?
    Get explanation of the available help options.

.NOTES
    This script must be run on your local machine, not in a container.
#>

[CmdletBinding(DefaultParameterSetName = 'ByName', SupportsShouldProcess, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $false, ParameterSetName = 'ByAll', HelpMessage = 'Which Font variant do you prefer?')]
    [Parameter(Mandatory = $false, ParameterSetName = 'ByName', HelpMessage = 'Which Font variant do you prefer?')]
    [ArgumentCompleter({
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
            @('Variable', 'Static') | Where-Object { $_ -like "$wordToComplete*" } | ForEach-Object { [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', $_) }
        })]
    [ValidateSet('Variable', 'Static')]
    [string]$Variant = 'Variable',

    [Parameter(Mandatory = $false, ParameterSetName = 'ByAll', HelpMessage = 'Specify the order to search for font types. Only the first matching type will be installed.')]
    [Parameter(Mandatory = $false, ParameterSetName = 'ByName', HelpMessage = 'Specify the order to search for font types. Only the first matching type will be installed.')]
    [ArgumentCompleter({
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
            @('TTF', 'OTF', 'WOFF2') | Where-Object { $_ -like "$wordToComplete*" } | ForEach-Object { [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', $_) }
        })]
    [ValidateSet('TTF', 'OTF', 'WOFF2')]
    [string[]]$Type = @('TTF', 'OTF', 'WOFF2'),

    [Parameter(Mandatory = $true, ParameterSetName = 'ByAll')]
    [switch]$All,

    [Parameter(Mandatory = $false, ParameterSetName = 'ListOnly', HelpMessage = 'List available Nerd Fonts matching the specified pattern.')]
    [AllowNull()]
    [AllowEmptyString()]
    [ArgumentCompleter({
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
            @('All') | Where-Object { $_ -like "$wordToComplete*" } | ForEach-Object { [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', $_) }
        })]
    [string]$List,

    [Parameter(Mandatory = $false, ParameterSetName = 'ByAll', HelpMessage = 'In which scope do you want to install the Nerd Font, AllUsers or CurrentUser?')]
    [Parameter(Mandatory = $false, ParameterSetName = 'ByName', HelpMessage = 'In which scope do you want to install the Nerd Font, AllUsers or CurrentUser?')]
    [ValidateSet('AllUsers', 'CurrentUser')]
    [string]$Scope = 'CurrentUser',

    [Parameter(Mandatory = $false, ParameterSetName = 'ByAll')]
    [Parameter(Mandatory = $false, ParameterSetName = 'ByName')]
    [switch]$Force,

    [Parameter(Mandatory = $true, ParameterSetName = 'Help', HelpMessage = "What kind of help would you like to see?")]
    [AllowEmptyString()]
    [AllowNull()]
    [ArgumentCompleter({
            param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameter)
            $helpOptions = @{
                Name     = "Shows help for the dynamic Name parameter."
                Variant  = "Shows help for the Variant parameter."
                Type     = "Shows help for the Type parameter."
                All      = "Shows help for the All parameter."
                List     = "Shows help for the List parameter."
                Scope    = "Shows help for the Scope parameter."
                Force    = "Shows help for the Force parameter."
                Help     = "Shows help for the Help parameter."
                Version  = "Shows help for the Version parameter."
                Summary  = "Provides a summary of the help content. Note that the dynamic Name parameter is not included in the summary."
                Detailed = "Provides detailed help, including parameter descriptions and examples. Note that the dynamic Name parameter is not included in the detailed help."
                Full     = "Provides full help, including detailed help, parameter descriptions, examples, and additional notes. Note that the dynamic Name parameter is not included in the full help."
                Examples = "Shows only the examples section of the help."
            }
            $helpOptions.GetEnumerator() | Where-Object { $_.Key -like "$wordToComplete*" } | ForEach-Object {
                [System.Management.Automation.CompletionResult]::new($_.Key, $_.Key, 'ParameterValue', $_.Value)
            }
        })]
    [string]$Help = 'Help',

    [Parameter(Mandatory = $true, ParameterSetName = 'Version')]
    [switch]$Version
)

dynamicparam {
    # Define the URL and cache file path
    $url = 'https://raw.githubusercontent.com/ryanoasis/nerd-fonts/master/bin/scripts/lib/fonts.json'
    $cacheFilePath = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), 'github-nerd-fonts.json')
    $cacheDuration = [TimeSpan]::FromMinutes(2)

    #region Functions ==========================================================
    function Get-FontsListFromWeb {
        <#
        .SYNOPSIS
        Fetch fonts list from the web server.

        .DESCRIPTION
        This function fetches the fonts list from the specified web server URL.
        It also adds a release URL property to each font object.
        #>
        try {
            $fonts = (Invoke-RestMethod -Uri $url -ErrorAction Stop -Verbose:$false -Debug:$false).fonts
            $releaseUrl = "https://api.github.com/repos/ryanoasis/nerd-fonts/releases/latest"
            foreach ($font in $fonts) {
                $font.PSObject.Properties.Add([PSNoteProperty]::new("releaseUrl", $releaseUrl))
            }
            return $fonts
        }
        catch {
            $PSCmdlet.ThrowTerminatingError($_)
        }
    }

    function Get-FontsListFromCache {
        <#
        .SYNOPSIS
        Load fonts list from cache.

        .DESCRIPTION
        This function loads the fonts list from a cache file if it exists and is not expired.
        #>
        if ([System.IO.Directory]::Exists($cacheFilePath)) {
            $cacheFile = Get-Item $cacheFilePath
            if ((Get-Date) -lt $cacheFile.LastWriteTime.Add($cacheDuration)) {
                return Get-Content $cacheFilePath | ConvertFrom-Json
            }
        }
        return $null
    }

    function Save-FontsListToCache($fonts) {
        <#
        .SYNOPSIS
        Save fonts list to cache.

        .DESCRIPTION
        This function saves the fonts list to a cache file in JSON format.
        #>
        $fonts | ConvertTo-Json | Set-Content $cacheFilePath
    }

    function Add-CustomEntries($fonts) {
        <#
        .SYNOPSIS
        Add custom entries to the fonts list.

        .DESCRIPTION
        This function adds custom font entries to the provided fonts list and sorts them by folder name.
        #>
        $customEntries = @(
            [PSCustomObject]@{
                unpatchedName          = 'Cascadia Code Font Family'
                licenseId              = 'OFL-1.1-RFN'
                RFN                    = $true
                version                = 'latest'
                patchedName            = 'Cascadia Code Font Family'
                folderName             = 'CascadiaCode'
                imagePreviewFont       = 'Cascadia Code Font Family'
                imagePreviewFontSource = $null
                linkPreviewFont        = 'cascadia-code'
                caskName               = 'cascadia-code'
                repoRelease            = $false
                description            = 'The official Cascadia Code font by Microsoft with all variants, including Nerd Font and Powerline'
                releaseUrl             = 'https://api.github.com/repos/microsoft/cascadia-code/releases/latest'
            },
            [PSCustomObject]@{
                unpatchedName          = 'Cascadia Code NF'
                licenseId              = 'OFL-1.1-RFN'
                RFN                    = $true
                version                = 'latest'
                patchedName            = 'Cascadia Code NF'
                folderName             = 'CascadiaCodeNF'
                imagePreviewFont       = 'Cascadia Code Nerd Font'
                imagePreviewFontSource = $null
                linkPreviewFont        = 'cascadia-code'
                caskName               = 'cascadia-code-nerd-font'
                repoRelease            = $false
                description            = 'The official Cascadia Code font by Microsoft that is enabled with Nerd Font symbols'
                releaseUrl             = 'https://api.github.com/repos/microsoft/cascadia-code/releases/latest'
            },
            [PSCustomObject]@{
                unpatchedName          = 'Cascadia Code PL'
                licenseId              = 'OFL-1.1-RFN'
                RFN                    = $true
                version                = 'latest'
                patchedName            = 'Cascadia Code PL'
                folderName             = 'CascadiaCodePL'
                imagePreviewFont       = 'Cascadia Code Powerline Font'
                imagePreviewFontSource = $null
                linkPreviewFont        = 'cascadia-code'
                caskName               = 'cascadia-code-powerline-font'
                repoRelease            = $false
                description            = 'The official Cascadia Code font by Microsoft that is enabled with Powerline symbols'
                releaseUrl             = 'https://api.github.com/repos/microsoft/cascadia-code/releases/latest'
            },
            [PSCustomObject]@{
                unpatchedName          = 'Cascadia Mono Font Family'
                licenseId              = 'OFL-1.1-RFN'
                RFN                    = $true
                version                = 'latest'
                patchedName            = 'Cascadia Mono Font Family'
                folderName             = 'CascadiaMono'
                imagePreviewFont       = 'Cascadia Mono Font Family'
                imagePreviewFontSource = $null
                linkPreviewFont        = $null
                caskName               = 'cascadia-mono'
                repoRelease            = $false
                description            = 'The official Cascadia Mono font by Microsoft with all variants, including Nerd Font and Powerline'
                releaseUrl             = 'https://api.github.com/repos/microsoft/cascadia-code/releases/latest'
            },
            [PSCustomObject]@{
                unpatchedName          = 'Cascadia Mono NF'
                licenseId              = 'OFL-1.1-RFN'
                RFN                    = $true
                version                = 'latest'
                patchedName            = 'Cascadia Mono NF'
                folderName             = 'CascadiaMonoNF'
                imagePreviewFont       = 'Cascadia Mono Nerd Font'
                imagePreviewFontSource = $null
                linkPreviewFont        = $null
                caskName               = 'cascadia-mono-nerd-font'
                repoRelease            = $false
                description            = 'The official Cascadia Mono font by Microsoft that is enabled with Nerd Font symbols'
                releaseUrl             = 'https://api.github.com/repos/microsoft/cascadia-code/releases/latest'
            },
            [PSCustomObject]@{
                unpatchedName          = 'Cascadia Mono PL'
                licenseId              = 'OFL-1.1-RFN'
                RFN                    = $true
                version                = 'latest'
                patchedName            = 'Cascadia Mono PL'
                folderName             = 'CascadiaMonoPL'
                imagePreviewFont       = 'Cascadia Mono Powerline Font'
                imagePreviewFontSource = $null
                linkPreviewFont        = $null
                caskName               = 'cascadia-mono-powerline-font'
                repoRelease            = $false
                description            = 'The official Cascadia Mono font by Microsoft that is enabled with Powerline symbols'
                releaseUrl             = 'https://api.github.com/repos/microsoft/cascadia-code/releases/latest'
            }
        )

        # Combine the original fonts with custom entries and sort by folderName
        $allFonts = $fonts + $customEntries
        $sortedFonts = $allFonts | Sort-Object -Property caskName

        return $sortedFonts
    }
    #endregion Functions -------------------------------------------------------

    # Try to load fonts list from cache
    $allNerdFonts = Get-FontsListFromCache

    # If cache is not valid, fetch from web, add custom entries, and update cache
    if (-not $allNerdFonts) {
        $allNerdFonts = Get-FontsListFromWeb
        $allNerdFonts = Add-CustomEntries $allNerdFonts
        Save-FontsListToCache $allNerdFonts
    }

    # Extract caskName values for auto-completion
    $caskNames = [string[]]@($allNerdFonts | ForEach-Object { $_.caskName })

    # Define the name and type of the dynamic parameter
    $paramName = 'Name'
    $paramType = [string[]]

    # Create a collection to hold the attributes for the dynamic parameter
    $attributes = [System.Collections.ObjectModel.Collection[System.Attribute]]::new()

    # Convert the caskNames array to a string representation
    $caskNamesString = $caskNames -join "', '"
    $caskNamesString = "@('$caskNamesString')"

    # Create an ArgumentCompleter attribute using the caskName values for auto-completion and add it to the collection
    $argumentCompleterScript = [scriptblock]::Create(@"
param(`$commandName, `$parameterName, `$wordToComplete, `$commandAst, `$fakeBoundParameter)
# Static array of cask names for auto-completion
`$caskNames = $caskNamesString

# Filter and return matching cask names
`$caskNames | Where-Object { `$_ -like "`$wordToComplete*" } | ForEach-Object {
    [System.Management.Automation.CompletionResult]::new(`$_, `$_, 'ParameterValue', `$_)
}
"@)

    $argumentCompleterAttribute = [System.Management.Automation.ArgumentCompleterAttribute]::new($argumentCompleterScript)
    $attributes.Add($argumentCompleterAttribute)

    # Create a Parameter attribute and add it to the collection
    $paramAttribute = [System.Management.Automation.ParameterAttribute]::new()
    $paramAttribute.Mandatory = $(
        # Make the parameter mandatory if the script is not running interactively
        if (
            $null -ne ([System.Environment]::GetCommandLineArgs() | Where-Object { $_ -match '^-NonI.*' }) -or
            (
                $null -ne ($__PSProfileEnvCommandLineArgs | Where-Object { $_ -match '^-C.*' }) -and
                $null -eq ($__PSProfileEnvCommandLineArgs | Where-Object { $_ -match '^-NoE.*' })
            )
        ) {
            $true
        }
        elseif ($Host.UI.RawUI.KeyAvailable -or [System.Environment]::UserInteractive) {
            $false
        }
        else {
            $true
        }
    )
    $paramAttribute.Position = 0
    $paramAttribute.ParameterSetName = 'ByName'
    $paramAttribute.HelpMessage = 'Which Nerd Font do you want to install?' + "`n" + "Available values: $($caskNames -join ', ')"
    $paramAttribute.ValueFromPipeline = $true
    $paramAttribute.ValueFromPipelineByPropertyName = $true
    $attributes.Add($paramAttribute)

    # Create the dynamic parameter
    $runtimeParam = [System.Management.Automation.RuntimeDefinedParameter]::new($paramName, $paramType, $attributes)

    # Create a dictionary to hold the dynamic parameters
    $paramDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()
    $paramDictionary.Add($paramName, $runtimeParam)

    # Return the dictionary
    return $paramDictionary
}

begin {
    if (
        $PSBoundParameters.ContainsKey('Help') -or
        (
            $PSBoundParameters.Name.Count -eq 1 -and
            @('help', '--help', '?') -contains $PSBoundParameters.Name[0]
        )
    ) {
        try {
            if ($null -eq $PSCommandPath -or $PSCommandPath -eq '') {
                $scriptContent = $MyInvocation.MyCommand.ScriptBlock.ToString()
                $tempFilePath = [System.IO.Path]::GetTempFileName()
                $tempPs1FilePath = $tempFilePath + '.ps1'
                Write-Verbose "Creating temporary file: $tempPs1FilePath"
                Set-Content -Path $tempPs1FilePath -Value $scriptContent
            }
            else {
                $tempFilePath = $PSCommandPath
                $tempPs1FilePath = $tempFilePath
            }

            # Use Get-Help to render the help content
            $params = @{ Name = $tempPs1FilePath }
            if ([string]::IsNullOrEmpty($Help)) {
                $params.Parameter = 'Help'
            }
            elseif ($Help -ne 'Summary') {
                if (@('Detailed', 'Full', 'Examples') -contains $Help) {
                    $params.$Help = $true
                }
                elseif (@('Variant', 'Type', 'All', 'List', 'Scope', 'Force', 'Help', 'Version') -contains $Help) {
                    $params.Parameter = $Help
                }
                elseif ($Help -eq 'Name') {
                    $scriptContent = $MyInvocation.MyCommand.ScriptBlock.ToString()
                    $helpContent = @()
                    $inHelpBlock = $false
                    $helpBlockFound = $false
                    $inParameterNameSection = $false

                    foreach ($line in $scriptContent -split "`n") {
                        if ($line -match '^\s*<#' -and -not $helpBlockFound) {
                            if ($line -notmatch '^\s*<#PSScriptInfo') {
                                $inHelpBlock = $true
                            }
                        }
                        if ($inHelpBlock) {
                            if ($line -match '^\s*\.PARAMETER\s+Name\s*$') {
                                $inParameterNameSection = $true
                            }
                            elseif ($line -match '^\s*\.PARAMETER\s+' -and $inParameterNameSection) {
                                $inParameterNameSection = $false
                            }

                            if ($inParameterNameSection) {
                                $helpContent += $line
                            }
                        }
                        if ($line -match '#>\s*$' -and $inHelpBlock) {
                            $inHelpBlock = $false
                            $helpBlockFound = $true
                        }
                        if ($helpBlockFound -and -not $inHelpBlock) {
                            break
                        }
                    }

                    if ($helpContent) {
                        Write-Output ''
                        $helpContent[0] = '-Name <String[]>'
                        $helpText = $helpContent -join "`n"
                        Write-Output $helpText
                        Write-Output ''
                        Write-Output '    Required?                    true (unless running in an interactive session to display the selection menu)'
                        Write-Output '    Position?                    0'
                        Write-Output '    Default value'
                        Write-Output '    Accept pipeline input?       true (ByValue, ByPropertyName)'
                        Write-Output '    Accept wildcard characters?  false'
                        Write-Output ''
                        Write-Output ''
                        Write-Output ''
                    }
                    else {
                        Write-Output "No .PARAMETER Name content found."
                    }
                    return
                }
                else {
                    $params.Parameter = 'Help'
                }
            }
            Get-Help @params
        }
        finally {
            if ($null -eq $PSCommandPath -or $PSCommandPath -eq '') {
                Write-Verbose "Removing temporary files: $tempFilePath, $tempPs1FilePath"
                Remove-Item -Path $tempFilePath -Force
                Remove-Item -Path $tempPs1FilePath -Force
            }
        }
        return
    }

    if (
        $Version -or
        (
            $PSBoundParameters.Name.Count -eq 1 -and
            @('version', '--version', 'ver') -eq $PSBoundParameters.Name[0]
        )
    ) {
        $scriptContent = $MyInvocation.MyCommand.ScriptBlock.ToString()
        $versionNumber = $null
        $inHelpBlock = $false
        $helpBlockFound = $false

        foreach ($line in $scriptContent -split "`n") {
            if ($line -match '^\s*<#PSScriptInfo' -and -not $helpBlockFound) {
                $inHelpBlock = $true
            }
            if ($inHelpBlock) {
                if ($line -match '^\s*\.VERSION\s+(.+)$') {
                    $versionNumber = $Matches[1].Trim()
                    break
                }
            }
            if ($line -match '#>\s*$' -and $inHelpBlock) {
                $inHelpBlock = $false
                $helpBlockFound = $true
            }
        }

        if ($versionNumber) {
            Write-Output $versionNumber
        }
        else {
            Write-Output "No version information found."
        }
        return
    }

    if ($PSBoundParameters.ContainsKey('List')) {
        # Set default value if List is null or empty
        if ([string]::IsNullOrEmpty($List)) {
            $List = "*"
        }
        else {
            $List = $List.Trim()
        }

        # Handle special case for 'All'
        if ($List -eq 'All') {
            $List = "*"
        }
        elseif ($List -notmatch '\*') {
            # Ensure the List contains wildcard characters
            $List = "*$List*"
        }

        # Filter and format the output
        $allNerdFonts | Where-Object { $_.caskName -like $List } | ForEach-Object {
            [PSCustomObject]@{
                Name        = $_.caskName
                DisplayName = $_.imagePreviewFont
                Description = $_.description
                SourceUrl   = $_.releaseUrl -replace '^(https?://)(?:[^/]+\.)*([^/]+\.[^/]+)/repos/([^/]+)/([^/]+).*', '$1$2/$3/$4'
            }
        }
        return
    }

    if (
        $null -ne $env:REMOTE_CONTAINERS -or
        $null -ne $env:CODESPACES -or
        $null -ne $env:WSL_INTEROP
    ) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.Exception]::new('This script must be run on your local machine, not in a container.'),
                'NotLocalMachine',
                [System.Management.Automation.ErrorCategory]::InvalidOperation,
                $null
            )
        )
    }

    if (
        $Scope -eq 'AllUsers' -and
        (
            (
                $PSVersionTable.Platform -ne 'Unix' -and
                -not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')
            ) -or
            (
                $PSVersionTable.Platform -eq 'Unix' -and
                ($(id -u) -ne '0')
            )
        )
    ) {
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.Exception]::new('Elevated permissions are required to install fonts for all users. Alternatively, you can install fonts for the current user using the -Scope parameter with the CurrentUser value.'),
                'InsufficientPermissions',
                [System.Management.Automation.ErrorCategory]::InvalidOperation,
                $null
            )
        )
    }

    #region Functions ==========================================================
    function Show-Menu {
        <#
        .SYNOPSIS
        Displays a menu for selecting fonts.

        .DESCRIPTION
        This function clears the host and displays a menu with options for selecting fonts.
        It handles user input and terminal resizing to dynamically adjust the menu display.
        #>
        param (
            $Options
        )
        Clear-Host

        function Show-MenuOptions {
            <#
            .SYNOPSIS
            Draws the menu options.

            .DESCRIPTION
            This function prints the menu options in a formatted manner.
            It calculates the number of columns and rows based on the terminal width and displays the options accordingly.
            #>
            param (
                $Options,
                $terminalWidth
            )

            # Add the 'All Fonts' option at the top
            $Options = @([pscustomobject]@{ imagePreviewFont = 'All Fonts'; unpatchedName = 'All'; caskName = 'All' }) + $Options

            # Calculate the maximum width of each column
            $maxOptionLength = ($Options | ForEach-Object { $_.imagePreviewFont.Length } | Measure-Object -Maximum).Maximum + 1 # 1 for padding
            $maxIndexLength = ($Options.Length).ToString().Length
            $linkSymbolLength = 1
            $columnWidth = $maxIndexLength + $maxOptionLength + $linkSymbolLength + 3  # 3 for padding and ': '

            # Calculate the number of columns that can fit in the terminal width
            $numColumns = [math]::Floor($terminalWidth / $columnWidth)

            # Calculate the number of rows
            $numRows = [math]::Ceiling($Options.Length / $numColumns)

            # Print the centered and bold title
            if ($IsCoreCLR) {
                $title = "`u{1F913} $($PSStyle.Bold)`e]8;;https://www.nerdfonts.com/`e\Nerd Fonts`e]8;;`e\ Installation$($PSStyle.BoldOff)"
            }
            else {
                $title = 'Nerd Fonts Installation'
            }
            $totalWidth = $columnWidth * $numColumns
            $padding = [math]::Max(0, ($totalWidth - ($title.Length / 2)) / 2)
            Write-Host (' ' * $padding + $title) -ForegroundColor Cyan -NoNewline
            Write-Host -ForegroundColor Cyan
            Write-Host (('_' * $totalWidth) + "`n") -ForegroundColor Cyan

            # Print the options in rows
            for ($row = 0; $row -lt $numRows; $row++) {
                for ($col = 0; $col -lt $numColumns; $col++) {
                    $index = $row + $col * $numRows
                    if ($index -lt $Options.Length) {
                        $number = $index
                        $fontName = $Options[$index].imagePreviewFont
                        $numberText = ('{0,' + $maxIndexLength + '}') -f $number
                        $linkSymbol = "`u{2197}" # Up-Right Arrow

                        if ($index -eq 0) {
                            # Special formatting for 'All Fonts'
                            Write-Host -NoNewline -ForegroundColor Magenta $numberText
                            Write-Host -NoNewline -ForegroundColor Magenta ': '
                            Write-Host -NoNewline -ForegroundColor Magenta "$($PSStyle.Italic)$fontName$($PSStyle.ItalicOff)  "
                        }
                        else {
                            Write-Host -NoNewline -ForegroundColor DarkYellow $numberText
                            Write-Host -NoNewline -ForegroundColor Yellow ': '
                            if ($fontName -match '^(.+)(Font Family)(.*)$') {
                                if ($IsCoreCLR -and $Options[$index].linkPreviewFont -is [string] -and -not [string]::IsNullOrEmpty($Options[$index].linkPreviewFont)) {
                                    $link = $Options[$index].linkPreviewFont
                                    if ($link -notmatch '^https?://') {
                                        $link = "https://www.programmingfonts.org/#$link"
                                    }
                                    $clickableLinkSymbol = " `e]8;;$link`e\$linkSymbol`e]8;;`e\"
                                    Write-Host -NoNewline -ForegroundColor White "$($PSStyle.Bold)$($Matches[1])$($PSStyle.BoldOff)"
                                    Write-Host -NoNewline -ForegroundColor Gray "$($PSStyle.Italic)$($Matches[2])$($PSStyle.ItalicOff)"
                                    Write-Host -NoNewline -ForegroundColor White "$($Matches[3])"
                                    Write-Host -NoNewline -ForegroundColor DarkBlue "$clickableLinkSymbol"
                                }
                                else {
                                    Write-Host -NoNewline -ForegroundColor White "$($PSStyle.Bold)$($Matches[1])$($PSStyle.BoldOff)"
                                    Write-Host -NoNewline -ForegroundColor Gray "$($PSStyle.Italic)$($Matches[2])$($PSStyle.ItalicOff)"
                                    Write-Host -NoNewline -ForegroundColor White "$($Matches[3])  "
                                }
                            }
                            else {
                                if ($IsCoreCLR -and $Options[$index].linkPreviewFont -is [string] -and -not [string]::IsNullOrEmpty($Options[$index].linkPreviewFont)) {
                                    $link = $Options[$index].linkPreviewFont
                                    if ($link -notmatch '^https?://') {
                                        $link = "https://www.programmingfonts.org/#$link"
                                    }
                                    $clickableLinkSymbol = " `e]8;;$link`e\$linkSymbol`e]8;;`e\"
                                    Write-Host -NoNewline -ForegroundColor White "$($PSStyle.Bold)$fontName$($PSStyle.BoldOff)"
                                    Write-Host -NoNewline -ForegroundColor DarkBlue "$clickableLinkSymbol"
                                }
                                else {
                                    Write-Host -NoNewline -ForegroundColor White "$($PSStyle.Bold)$fontName$($PSStyle.BoldOff)  "
                                }
                            }
                        }
                        # Add padding to align columns
                        $paddingLength = $maxOptionLength - $fontName.Length
                        Write-Host -NoNewline (' ' * $paddingLength)
                    }
                }
                Write-Host
            }
        }

        # Initial terminal width
        $initialWidth = [console]::WindowWidth

        # Draw the initial menu
        Show-MenuOptions -Options $Options -terminalWidth $initialWidth

        Write-Host "`nEnter 'q' to quit." -ForegroundColor Cyan

        # Loop to handle user input and terminal resizing
        while ($true) {
            $currentWidth = [console]::WindowWidth
            if ($currentWidth -ne $initialWidth) {
                Clear-Host
                Show-MenuOptions -Options $Options -terminalWidth $currentWidth
                Write-Host "`nEnter 'q' to quit." -ForegroundColor Cyan
                $initialWidth = $currentWidth
            }

            $selection = Read-Host "`nSelect one or more numbers separated by commas"
            if ($selection -eq 'q') {
                return 'quit'
            }

            # Remove spaces and split the input by commas
            $selection = $selection -replace '\s', ''
            $numbers = $selection -split ',' | Select-Object -Unique

            # Validate each number
            $validSelections = @()
            $invalidSelections = @()
            foreach ($number in $numbers) {
                if ($number -match '^-?\d+$') {
                    $index = [int]$number - 1
                    if ($index -lt 0) {
                        return 'All'
                    }
                    elseif ($index -ge 0 -and $index -lt $Options.Count) {
                        $validSelections += $Options[$index]
                    }
                    else {
                        $invalidSelections += $number
                    }
                }
                else {
                    $invalidSelections += $number
                }
            }

            if ($invalidSelections.Count -eq 0) {
                return $validSelections.caskName
            }
            else {
                Write-Host "Invalid selection(s): $($invalidSelections -join ', '). Please enter valid numbers between 0 and $($Options.Length) or 'q' to quit." -ForegroundColor Red
            }
        }
    }

    function Invoke-GitHubApiRequest {
        <#
        .SYNOPSIS
        Makes anonymous requests to GitHub API and handles rate limiting.

        .DESCRIPTION
        This function sends a request to the specified GitHub API URI and handles rate limiting by retrying the request
        up to a maximum number of retries. It also converts JSON responses to PowerShell objects.
        #>
        param (
            [string]$Uri
        )
        $maxRetries = 5
        $retryCount = 0
        $baseWaitTime = 15

        while ($retryCount -lt $maxRetries) {
            try {
                $headers = @{}
                $parsedUri = [System.Uri]$Uri
                if ($parsedUri.Host -eq "api.github.com") {
                    $headers["Accept"] = "application/vnd.github.v3+json"
                }

                $response = Invoke-RestMethod -Uri $Uri -Headers $headers -ErrorAction Stop -Verbose:$false -Debug:$false

                return [PSCustomObject]@{
                    Headers = $response.PSObject.Properties["Headers"].Value
                    Content = $response
                }
            }
            catch {
                if ($_.Exception.Response.StatusCode -eq 403 -or $_.Exception.Response.StatusCode -eq 429) {
                    $retryAfter = $null
                    $rateLimitReset = $null
                    $waitTime = 0

                    if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers["Retry-After"]) {
                        $retryAfter = $_.Exception.Response.Headers["Retry-After"]
                    }
                    if ($_.Exception.Response.Headers -and $_.Exception.Response.Headers["X-RateLimit-Reset"]) {
                        $rateLimitReset = $_.Exception.Response.Headers["X-RateLimit-Reset"]
                    }

                    if ($retryAfter) {
                        $waitTime = [int]$retryAfter
                    }
                    elseif ($rateLimitReset) {
                        $resetTime = [DateTimeOffset]::FromUnixTimeSeconds([int]$rateLimitReset).LocalDateTime
                        $waitTime = ($resetTime - (Get-Date)).TotalSeconds
                    }

                    if ($waitTime -gt 0 -and $waitTime -le 60) {
                        Write-Host "Rate limit exceeded. Waiting for $waitTime seconds."
                        Start-Sleep -Seconds $waitTime
                    }
                    else {
                        $exponentialWait = $baseWaitTime * [math]::Pow(2, $retryCount)
                        Write-Host "Rate limit exceeded. Waiting for $exponentialWait seconds."
                        Start-Sleep -Seconds $exponentialWait
                    }
                    $retryCount++
                }
                else {
                    $PSCmdlet.ThrowTerminatingError($_)
                }
            }
        }
        $PSCmdlet.ThrowTerminatingError(
            [System.Management.Automation.ErrorRecord]::new(
                [System.Exception]::new('Max retries exceeded. Please try again later.'),
                'MaxRetriesExceeded',
                [System.Management.Automation.ErrorCategory]::ResourceUnavailable,
                $null
            )
        )
    }

    function Invoke-GitHubApiPaginatedRequest {
        <#
        .SYNOPSIS
        Fetches all pages of a paginated response if the host is api.github.com.

        .DESCRIPTION
        This function sends requests to the specified GitHub API URI and handles pagination by following the 'next' links
        in the response headers. It collects all pages of data and returns them as a single array.
        #>
        param (
            [string]$Uri
        )
        $allData = @()
        $parsedUri = [System.Uri]$Uri

        if ($parsedUri.Host -eq "api.github.com") {
            while ($true) {
                $response = Invoke-GitHubApiRequest -Uri $Uri
                if ($null -eq $response) {
                    break
                }
                $data = $response.Content
                $allData += $data
                $linkHeader = $null
                if ($response.Headers -and $response.Headers["Link"]) {
                    $linkHeader = $response.Headers["Link"]
                }
                if ($linkHeader -notmatch 'rel="next"') {
                    break
                }
                $nextLink = ($linkHeader -split ',') | Where-Object { $_ -match 'rel="next"' } | ForEach-Object { ($_ -split ';')[0].Trim('<> ') }
                $Uri = $nextLink
            }
        }
        else {
            $response = Invoke-GitHubApiRequest -Uri $Uri
            $allData = $response.Content
        }
        return $allData
    }
    function Test-TarSupportsFormat {
        param (
            [string]$format
        )
        $tarVersionOutput = & tar --version 2>&1
        $isGnuTar = $tarVersionOutput -match 'GNU tar'

        if ($isGnuTar) {
            # Extract GNU tar version
            $tarVersion = [regex]::Match($tarVersionOutput, 'tar \(GNU tar\) (\d+\.\d+(?:.\d+)?)').Groups[1].Value
            if ([string]::IsNullOrEmpty($tarVersion)) {
                return $false
            }

            switch ($format) {
                'xz' { return [version]$tarVersion -ge [version]'1.22' }
                'bzip2' { return [version]$tarVersion -ge [version]'1.15' }
                'gz' { return $true }  # Generally supported in all versions
                default { return $false }
            }
        }
        else {
            # Assume BSD tar
            switch ($format) {
                'xz' { return $tarVersionOutput -match 'liblzma' }
                'bzip2' { return $tarVersionOutput -match 'bz2lib' }
                'gz' { return $tarVersionOutput -match 'zlib' }
                default { return $false }
            }
        }
    }
    function Expand-FromArchiveType {
        param (
            [string]$SourceFile,
            [string]$DestinationFolder,
            [string]$FileExtension,
            [string]$Executable
        )

        # Define a mapping table for command templates
        $commandTemplates = @{
            'tar:xz'  = 'tar -xJf "{0}" -C "{1}"'
            'tar:bz2' = 'tar -xjf "{0}" -C "{1}"'
            'tar:gz'  = 'tar -xzf "{0}" -C "{1}"'
            'tar'     = 'tar -xf "{0}" -C "{1}"'
            'xz'      = 'xz -d "{0}" -C "{1}"'
            '7z'      = '7z x "{0}" -o"{1}"'
            'bzip2'   = 'bzip2 -d "{0}" -C "{1}"'
            'gzip'    = 'gzip -d "{0}" -C "{1}"'
        }

        if ($null -eq $Executable) {
            throw "Unsupported archive format: $FileExtension"
        }
        elseif ($Executable -eq 'powershell') {
            # Use .NET functions to extract zip files
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($SourceFile, $DestinationFolder)
            Write-Verbose "Extracted zip file using .NET functions."
        }
        else {
            $commandTemplate = $commandTemplates[$Executable]
            $command = $commandTemplate -f $SourceFile, $DestinationFolder
            Write-Verbose "Running command: $command"

            # Split the command into the executable and its arguments
            $commandParts = $command -split ' ', 2
            $arguments = if ($commandParts.Length -gt 1) { $commandParts[1] } else { "" }

            # Execute the external command
            Start-Process -FilePath $commandParts[0] -ArgumentList $arguments -NoNewWindow -Wait
        }
    }
    #endregion Functions -------------------------------------------------------

    # Provide interactive selection if no font name is specified
    if (-not $PSBoundParameters.Name -and -not $PSBoundParameters.All) {
        do {
            $Name = Show-Menu -Options $allNerdFonts
            if ($Name -eq 'quit') {
                Write-Host "Selection process canceled."
                return
            }
        } while (-not $Name)

        if ($Name) {
            if ($Name -eq 'All') {
                Write-Host "`nYou selected all fonts.`n" -ForegroundColor Yellow
                # Proceed with the installation of all fonts
            }
            else {
                Write-Host "`nYour selected font(s): $($Name -join ', ')`n" -ForegroundColor Yellow
                # Proceed with the installation of the selected font(s)
            }
        }
        else {
            return
        }
    }
    elseif ($PSBoundParameters.Name) {
        $Name = $PSBoundParameters.Name
    }

    $nerdFontsToInstall = if ($PSBoundParameters.All -or $Name -contains 'All') {
        $allNerdFonts
    }
    else {
        $allNerdFonts | Where-Object { $Name -contains $_.caskName }
    }

    if ($nerdFontsToInstall.Count -eq 0) {
        Write-Error "No matching fonts found."
        return
    }

    # Fetch releases for each unique URL
    $fontReleases = @{}
    foreach ($url in $nerdFontsToInstall.releaseUrl | Sort-Object -Unique) {
        Write-Verbose "Fetching release data for $url"
        $release = Invoke-GitHubApiPaginatedRequest -Uri $url
        $fontReleases[$url] = @{
            ReleaseData = $release
            Sha256Data  = @{}
        }

        # Check if the release contains a SHA-256.txt asset
        $shaAsset = $release.assets | Where-Object { $_.name -eq 'SHA-256.txt' }
        if ($shaAsset) {
            $shaUrl = $shaAsset.browser_download_url
            Write-Verbose "Fetching SHA-256.txt content from $shaUrl"
            $shaContent = Invoke-WebRequest -Uri $shaUrl -ErrorAction Stop -Verbose:$false -Debug:$false

            # Convert the binary content to a string
            $shaContentString = [System.Text.Encoding]::UTF8.GetString($shaContent.Content)

            # Parse the SHA-256.txt content
            $shaLines = $shaContentString -split "`n"
            foreach ($line in $shaLines) {
                if ($line -match '^\s*([a-fA-F0-9]{64})\s+(.+)$') {
                    $sha256 = $Matches[1]
                    $fileName = $Matches[2].Trim()
                    $fontReleases[$url].Sha256Data[$fileName] = $sha256
                    Write-Debug "SHA-256: $sha256, File: $fileName"
                }
            }
        }
    }

    # Determine the XDG_DATA_HOME directory
    $xdgDataHome = $env:XDG_DATA_HOME
    if (-not $xdgDataHome) {
        if ($IsMacOS -or $IsLinux) {
            $xdgDataHome = "${HOME}/.local/share"
        }
        else {
            $xdgDataHome = $env:LOCALAPPDATA
        }
    }

    # Determine the font destination folder path based on the platform and scope
    if ($IsMacOS) {
        if ($Scope -eq 'AllUsers') {
            $fontDestinationFolderPath = '/Library/Fonts'
        }
        else {
            $fontDestinationFolderPath = "${HOME}/Library/Fonts"
        }
    }
    elseif ($IsLinux) {
        if ($Scope -eq 'AllUsers') {
            $fontDestinationFolderPath = '/usr/share/fonts'
        }
        else {
            $fontDestinationFolderPath = "${xdgDataHome}/fonts"
        }
    }
    else {
        if ($Scope -eq 'AllUsers') {
            $fontDestinationFolderPath = "${env:windir}\Fonts"
        }
        else {
            $fontDestinationFolderPath = "${xdgDataHome}\Microsoft\Windows\Fonts"
        }
    }
    $null = [System.IO.Directory]::CreateDirectory($fontDestinationFolderPath)
    Write-Verbose "Font Destination directory: $fontDestinationFolderPath"

    # Determine the supported archive formats based on the local machine
    $supportedArchiveFormats = [System.Collections.Generic.List[PSCustomObject]]::new()
    $archivePreferenceOrder = @('tar.xz', 'xz', '7z', 'tar.bz2', 'tar.gz', 'zip', 'tar')

    # ZIP is natively supported in PowerShell
    [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'zip'; Executable = 'powershell' })

    if ($IsMacOS -or $IsLinux) {
        # Prefer tar if available
        if (Get-Command tar -ErrorAction Ignore) {
            if (Test-TarSupportsFormat 'xz') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.xz'; Executable = 'tar:xz' }) }
            if (Test-TarSupportsFormat 'bzip2') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.bz2'; Executable = 'tar:bz2' }) }
            if (Test-TarSupportsFormat 'gz') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.gz'; Executable = 'tar:gz' }) }
            [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar'; Executable = 'tar' })
        }
        # Check for individual tools
        if (Get-Command xz -ErrorAction Ignore) { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'xz'; Executable = 'xz' }) }
        if (Get-Command 7z -ErrorAction Ignore) { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = '7z'; Executable = '7z' }) }
        if (Get-Command bzip2 -ErrorAction Ignore) { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.bz2'; Executable = 'bzip2' }) }
        if (Get-Command gzip -ErrorAction Ignore) { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.gz'; Executable = 'gzip' }) }
    }
    else {
        if (Get-Command tar -ErrorAction Ignore) {
            if (Test-TarSupportsFormat 'xz') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.xz'; Executable = 'tar:xz' }) }
            if (Test-TarSupportsFormat 'bzip2') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.bz2'; Executable = 'tar:bz2' }) }
            if (Test-TarSupportsFormat 'gz') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.gz'; Executable = 'tar:gz' }) }
            [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar'; Executable = 'tar' })
        }
        if (Get-Command 7z -ErrorAction Ignore) { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = '7z'; Executable = '7z' }) }
    }

    # Sort the supportedArchiveFormats based on the preference order and remove duplicates
    $sortedSupportedArchiveFormats = [System.Collections.Generic.List[PSCustomObject]]::new()
    foreach ($FileExtension in $archivePreferenceOrder) {
        $item = $supportedArchiveFormats | Where-Object { $_.FileExtension -eq $FileExtension } | Select-Object -First 1
        if ($item) {
            [void]$sortedSupportedArchiveFormats.Add($item)
        }
    }

    $supportedArchiveFormats = $sortedSupportedArchiveFormats
    Write-Verbose "Supported Archive Formats: $($supportedArchiveFormats.FileExtension -join ', ')"

    # Generate a unique temporary directory to store the font files
    $tempFile = [System.IO.Path]::GetTempFileName()
    $tempPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($tempFile), [System.IO.Path]::GetFileNameWithoutExtension($tempFile))
    $null = [System.IO.Directory]::CreateDirectory($tempPath)
    [System.IO.File]::Delete($tempFile)
    Write-Verbose "Using temporary directory: $tempPath"
}

process {
    if ($nerdFontsToInstall.Count -eq 0) { return }

    try {
        Write-Verbose "Installing $($nerdFontsToInstall.Count) Nerd Fonts to $Scope scope."

        foreach ($nerdFont in $nerdFontsToInstall) {
            $sourceName = $nerdFont.releaseUrl -replace '^https?://(?:[^/]+\.)*([^/]+\.[^/]+)/repos/([^/]+)/([^/]+).*', '$1/$2/$3'

            Write-Verbose "Processing font: $($nerdFont.folderName) [$($nerdFont.caskName)] ($($nerdFont.imagePreviewFont)) from $sourceName"

            foreach ($archiveFormat in $supportedArchiveFormats) {
                if ($null -eq $nerdFont.imagePreviewFontSource) {
                    $assetUrl = $fontReleases[$nerdFont.releaseUrl].ReleaseData.assets | Where-Object { $_.name -match "\.$($archiveFormat.FileExtension)$" } | Select-Object -ExpandProperty browser_download_url
                }
                else {
                    $assetUrl = $fontReleases[$nerdFont.releaseUrl].ReleaseData.assets | Where-Object { $_.name -match "^$($nerdFont.folderName)\.$($archiveFormat.FileExtension)$" } | Select-Object -ExpandProperty browser_download_url
                }
                if (-not [string]::IsNullOrEmpty($assetUrl)) { break }
            }

            if ([string]::IsNullOrEmpty($assetUrl)) {
                if ($WhatIfPreference -eq $true) {
                    Write-Warning "Nerd Font '$($nerdFont.folderName)' not found."
                }
                else {
                    Write-Error "Nerd Font '$($nerdFont.folderName)' not found."
                }
                continue
            }

            Write-Verbose "Font archive URL: $assetUrl"
            Write-Verbose "Font archive format: $($archiveFormat.FileExtension)"
            Write-Verbose "Font archive extract executable: $($archiveFormat.Executable)"

            if (
                $PSCmdlet.ShouldProcess(
                    "Install '$($nerdFont.imagePreviewFont)' from $sourceName",
                    "Do you confirm to install '$($nerdFont.imagePreviewFont)' from $sourceName ?",
                    "Nerd Fonts Installation"
                )
            ) {

                # Download the archive file if not already downloaded
                $archiveFileName = [System.IO.Path]::GetFileName(([System.Uri]::new($assetUrl)).LocalPath)
                $archivePath = [System.IO.Path]::Combine($tempPath, $archiveFileName)
                if (Test-Path -Path $archivePath) {
                    Write-Verbose "Font archive already downloaded: $archivePath"
                }
                else {
                    Write-Verbose "Downloading font archive from $assetUrl to $archivePath"
                    Invoke-WebRequest -Uri $assetUrl -OutFile $archivePath -ErrorAction Stop -Verbose:$false -Debug:$false
                }

                # Verify the SHA-256 hash if available
                if ($fontReleases[$nerdFont.releaseUrl].Sha256Data.Count -gt 0) {
                    if (-not $fontReleases[$nerdFont.releaseUrl].Sha256Data.ContainsKey($archiveFileName)) {
                        Write-Warning "SHA-256 Hash not found for $archiveFileName. Skipping installation."
                        continue
                    }

                    $expectedSha256 = $fontReleases[$nerdFont.releaseUrl].Sha256Data[$archiveFileName]
                    $actualSha256 = Get-FileHash -Path $archivePath -Algorithm SHA256 | Select-Object -ExpandProperty Hash
                    if ($expectedSha256 -ne $actualSha256) {
                        Write-Error "SHA-256 Hash mismatch for $archiveFileName. Skipping installation."
                        continue
                    }
                    Write-Verbose "SHA-256 Hash verified for $archiveFileName"
                }

                # Extract the font files if not already extracted
                $extractPath = [System.IO.Path]::Combine([System.IO.Path]::GetDirectoryName($archivePath), [System.IO.Path]::GetFileNameWithoutExtension($archivePath))
                if (Test-Path -Path $extractPath) {
                    Write-Verbose "Font files already extracted to $extractPath"
                }
                else {
                    Write-Verbose "Extracting font files to $extractPath"
                    $null = [System.IO.Directory]::CreateDirectory($extractPath)
                    Expand-FromArchiveType -SourceFile $archivePath -DestinationFolder $extractPath -FileExtension $archiveFormat.FileExtension -Executable $archiveFormat.Executable
                }

                # Determine search paths for font files based in $Variant parameter
                if ($Variant -eq 'Static') {
                    $staticPath = [System.IO.Path]::Combine($extractPath, 'static')
                    if (Test-Path -PathType Container -Path $staticPath) {
                        Write-Verbose "Using static font files from $staticPath"
                        $extractPath = $staticPath
                    }
                }
                foreach ($t in $Type) {
                    $typePath = [System.IO.Path]::Combine($extractPath, $t.ToLower())
                    if (Test-Path -PathType Container -Path $typePath) {
                        if ($Variant -eq 'Static') {
                            $staticPath = [System.IO.Path]::Combine($typePath, 'static')
                            if (Test-Path -PathType Container -Path $staticPath) {
                                Write-Verbose "Using static font files from $staticPath"
                                $extractPath = $staticPath
                            }
                            else {
                                Write-Verbose "Using font files from $typePath"
                                $extractPath = $typePath
                            }
                        }
                        else {
                            Write-Verbose "Using font files from $typePath"
                            $extractPath = $typePath
                        }
                        break
                    }
                }

                # Search for font files in the extracted directory
                foreach ($t in $Type) {
                    $filter = "*.$($t.ToLower())"

                    # Special case for font archives with multiple fonts like 'Cascadia'
                    if ($null -eq $nerdFont.imagePreviewFontSource) {
                        $filter = "$($nerdFont.folderName)$filter"
                    }

                    # Get font files
                    $fontFiles = @( Get-ChildItem -Path $extractPath -Filter $filter )

                    # Check if any files were found
                    if ($fontFiles.Count -gt 0) {
                        break
                    }
                }

                if ($fontFiles.Count -eq 0) {
                    Write-Error "No font files found for $($nerdFont.folderName)."
                    continue
                }

                # Install the font files
                foreach ($fontFile in $fontFiles) {
                    try {
                        # Check if font file is already registered in user scope by another application like Windows Terminal
                        if ($IsWindows) {
                            $fontRegistryPath = 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts'
                            $fontRegistryKeys = Get-ChildItem -Path $fontRegistryPath -Recurse -ErrorAction Ignore
                            $fontOwnedByApp = $false

                            foreach ($fontRegistryKey in $fontRegistryKeys) {
                                $fontRegistryValues = $fontRegistryKey.GetValueNames() | ForEach-Object {
                                    $value = $fontRegistryKey.GetValue($_)
                                    $fileName = Split-Path -Path $value -Leaf
                                    [PSCustomObject]@{
                                        Name     = $_
                                        FileName = $fileName
                                    }
                                }

                                $fontRegistryValue = $fontRegistryValues | Where-Object { $_.FileName -eq $fontFile.Name }
                                if ($fontRegistryValue.Count -gt 0) {
                                    $fontOwnedByApp = $true
                                    Write-Verbose "Font file $($fontFile.Name) already registered by application: $(Split-Path -Path $fontRegistryKey.Name -Leaf)"
                                    continue
                                }
                            }

                            if ($fontOwnedByApp) { continue }
                        }

                        $fontFileDestinationPath = [System.IO.Path]::Combine($fontDestinationFolderPath, $fontFile.Name)
                        if (-not $Force -and (Test-Path -Path $fontFileDestinationPath)) {
                            if ($Force) {
                                Write-Verbose "Overwriting font file: $($fontFile.Name)"
                            }
                            Write-Verbose "Font file already exists: $($fontFile.Name)"
                            Write-Host -NoNewline "  `u{2713} " -ForegroundColor Green
                        }
                        else {
                            if ($Force) {
                                Write-Verbose "Overwriting font file: $($fontFile.Name)"
                            }
                            else {
                                Write-Verbose "Copying font file: $($fontFile.Name)"
                            }

                            $maxRetries = 10
                            $retryIntervalSeconds = 1
                            $retryCount = 0
                            $fileCopied = $false
                            do {
                                try {
                                    $null = $fontFile.CopyTo($fontFileDestinationPath, $Force)
                                    $fileCopied = $true
                                }
                                catch {
                                    $retryCount++
                                    if ($retryCount -eq $maxRetries) {
                                        Write-Verbose "Failed to copy font file: $($fontFile.Name). Maximum retries exceeded."
                                        break
                                    }
                                    Write-Verbose "Failed to copy font file: $($fontFile.Name). Retrying in $retryIntervalSeconds seconds ..."
                                    Start-Sleep -Seconds $retryIntervalSeconds
                                }
                            } while (-not $fileCopied -and $retryCount -lt $maxRetries)

                            if (-not $fileCopied) {
                                throw "Failed to copy font file: $($fontFile.Name)."
                            }

                            # Register font file on Windows
                            if ($IsWindows) {
                                $fontType = if ([System.IO.Path]::GetExtension($fontFile.FullName).TrimStart('.') -eq 'otf') { 'OpenType' } else { 'TrueType' }
                                $params = @{
                                    Name         = "$($fontFile.BaseName) ($fontType)"
                                    Path         = if ($Scope -eq 'AllUsers') { 'HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts' } else { 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts' }
                                    PropertyType = 'string'
                                    Value        = if ($Scope -eq 'AllUsers') { $fontFile.Name } else { $fontFileDestinationPath }
                                    Force        = $true
                                    ErrorAction  = 'Stop'
                                }
                                Write-Verbose "Registering font file as '$($params.Name)' in $($params.Path)"
                                $null = New-ItemProperty @params
                            }

                            Write-Host -NoNewline "  $($PSStyle.Bold)`u{2713}$($PSStyle.BoldOff) " -ForegroundColor Green
                        }
                        Write-Host $fontFile.Name
                    }
                    catch {
                        Write-Host -NoNewline "  `u{2717} " -ForegroundColor Red
                        Write-Host $fontFile.Name
                        throw $_
                    }
                }

                Write-Host "`n$($PSStyle.Bold)'$($nerdFont.imagePreviewFont)'$($PSStyle.BoldOff) installed successfully.`n" -ForegroundColor Green
            }
            elseif ($WhatIfPreference -eq $true) {
                Write-Verbose "Predicted installation: $($nerdFont.folderName) [$($nerdFont.caskName)] ($($nerdFont.imagePreviewFont))"
            }
            else {
                Write-Verbose "Skipping font: $($nerdFont.folderName) [$($nerdFont.caskName)] ($($nerdFont.imagePreviewFont))"
            }
        }
    }
    catch {
        if ([System.IO.Directory]::Exists($tempPath)) {
            Write-Verbose "Removing temporary directory: $tempPath"
            [System.IO.Directory]::Delete($tempPath, $true)
        }
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

end {
    if ($nerdFontsToInstall.Count -eq 0) { return }

    if ([System.IO.Directory]::Exists($tempPath)) {
        Write-Verbose "Removing temporary directory: $tempPath"
        [System.IO.Directory]::Delete($tempPath, $true)
    }

    if ($IsLinux -and (Get-Command -Name fc-cache -ErrorAction Ignore)) {
        if ($VerbosePreference -eq 'Continue') {
            Write-Verbose "Refreshing font cache"
            fc-cache -fv
        }
        else {
            fc-cache -f
        }
    }
}
# SIG # Begin signature block
# MIInkQYJKoZIhvcNAQcCoIIngjCCJ34CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDAhJbVCRVYHC9H
# ZJ3ytXJufmuxCby098VZUpkLfk9Go6CCIKcwggWNMIIEdaADAgECAhAOmxiO+dAt
# 5+/bUOIIQBhaMA0GCSqGSIb3DQEBDAUAMGUxCzAJBgNVBAYTAlVTMRUwEwYDVQQK
# EwxEaWdpQ2VydCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xJDAiBgNV
# BAMTG0RpZ2lDZXJ0IEFzc3VyZWQgSUQgUm9vdCBDQTAeFw0yMjA4MDEwMDAwMDBa
# Fw0zMTExMDkyMzU5NTlaMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2Vy
# dCBJbmMxGTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lD
# ZXJ0IFRydXN0ZWQgUm9vdCBHNDCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoC
# ggIBAL/mkHNo3rvkXUo8MCIwaTPswqclLskhPfKK2FnC4SmnPVirdprNrnsbhA3E
# MB/zG6Q4FutWxpdtHauyefLKEdLkX9YFPFIPUh/GnhWlfr6fqVcWWVVyr2iTcMKy
# unWZanMylNEQRBAu34LzB4TmdDttceItDBvuINXJIB1jKS3O7F5OyJP4IWGbNOsF
# xl7sWxq868nPzaw0QF+xembud8hIqGZXV59UWI4MK7dPpzDZVu7Ke13jrclPXuU1
# 5zHL2pNe3I6PgNq2kZhAkHnDeMe2scS1ahg4AxCN2NQ3pC4FfYj1gj4QkXCrVYJB
# MtfbBHMqbpEBfCFM1LyuGwN1XXhm2ToxRJozQL8I11pJpMLmqaBn3aQnvKFPObUR
# WBf3JFxGj2T3wWmIdph2PVldQnaHiZdpekjw4KISG2aadMreSx7nDmOu5tTvkpI6
# nj3cAORFJYm2mkQZK37AlLTSYW3rM9nF30sEAMx9HJXDj/chsrIRt7t/8tWMcCxB
# YKqxYxhElRp2Yn72gLD76GSmM9GJB+G9t+ZDpBi4pncB4Q+UDCEdslQpJYls5Q5S
# UUd0viastkF13nqsX40/ybzTQRESW+UQUOsxxcpyFiIJ33xMdT9j7CFfxCBRa2+x
# q4aLT8LWRV+dIPyhHsXAj6KxfgommfXkaS+YHS312amyHeUbAgMBAAGjggE6MIIB
# NjAPBgNVHRMBAf8EBTADAQH/MB0GA1UdDgQWBBTs1+OC0nFdZEzfLmc/57qYrhwP
# TzAfBgNVHSMEGDAWgBRF66Kv9JLLgjEtUYunpyGd823IDzAOBgNVHQ8BAf8EBAMC
# AYYweQYIKwYBBQUHAQEEbTBrMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdp
# Y2VydC5jb20wQwYIKwYBBQUHMAKGN2h0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNv
# bS9EaWdpQ2VydEFzc3VyZWRJRFJvb3RDQS5jcnQwRQYDVR0fBD4wPDA6oDigNoY0
# aHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENB
# LmNybDARBgNVHSAECjAIMAYGBFUdIAAwDQYJKoZIhvcNAQEMBQADggEBAHCgv0Nc
# Vec4X6CjdBs9thbX979XB72arKGHLOyFXqkauyL4hxppVCLtpIh3bb0aFPQTSnov
# Lbc47/T/gLn4offyct4kvFIDyE7QKt76LVbP+fT3rDB6mouyXtTP0UNEm0Mh65Zy
# oUi0mcudT6cGAxN3J0TU53/oWajwvy8LpunyNDzs9wPHh6jSTEAZNUZqaVSwuKFW
# juyk1T3osdz9HNj0d1pcVIxv76FQPfx2CWiEn2/K2yCNNWAcAgPLILCsWKAOQGPF
# mCLBsln1VWvPJ6tsds5vIy30fnFqI2si/xK4VC0nftg62fC2h5b9W9FcrBjDTZ9z
# twGpn1eqXijiuZQwggauMIIElqADAgECAhAHNje3JFR82Ees/ShmKl5bMA0GCSqG
# SIb3DQEBCwUAMGIxCzAJBgNVBAYTAlVTMRUwEwYDVQQKEwxEaWdpQ2VydCBJbmMx
# GTAXBgNVBAsTEHd3dy5kaWdpY2VydC5jb20xITAfBgNVBAMTGERpZ2lDZXJ0IFRy
# dXN0ZWQgUm9vdCBHNDAeFw0yMjAzMjMwMDAwMDBaFw0zNzAzMjIyMzU5NTlaMGMx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7MDkGA1UEAxMy
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBSU0E0MDk2IFNIQTI1NiBUaW1lU3RhbXBpbmcg
# Q0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDGhjUGSbPBPXJJUVXH
# JQPE8pE3qZdRodbSg9GeTKJtoLDMg/la9hGhRBVCX6SI82j6ffOciQt/nR+eDzMf
# UBMLJnOWbfhXqAJ9/UO0hNoR8XOxs+4rgISKIhjf69o9xBd/qxkrPkLcZ47qUT3w
# 1lbU5ygt69OxtXXnHwZljZQp09nsad/ZkIdGAHvbREGJ3HxqV3rwN3mfXazL6IRk
# tFLydkf3YYMZ3V+0VAshaG43IbtArF+y3kp9zvU5EmfvDqVjbOSmxR3NNg1c1eYb
# qMFkdECnwHLFuk4fsbVYTXn+149zk6wsOeKlSNbwsDETqVcplicu9Yemj052FVUm
# cJgmf6AaRyBD40NjgHt1biclkJg6OBGz9vae5jtb7IHeIhTZgirHkr+g3uM+onP6
# 5x9abJTyUpURK1h0QCirc0PO30qhHGs4xSnzyqqWc0Jon7ZGs506o9UD4L/wojzK
# QtwYSH8UNM/STKvvmz3+DrhkKvp1KCRB7UK/BZxmSVJQ9FHzNklNiyDSLFc1eSuo
# 80VgvCONWPfcYd6T/jnA+bIwpUzX6ZhKWD7TA4j+s4/TXkt2ElGTyYwMO1uKIqjB
# Jgj5FBASA31fI7tk42PgpuE+9sJ0sj8eCXbsq11GdeJgo1gJASgADoRU7s7pXche
# MBK9Rp6103a50g5rmQzSM7TNsQIDAQABo4IBXTCCAVkwEgYDVR0TAQH/BAgwBgEB
# /wIBADAdBgNVHQ4EFgQUuhbZbU2FL3MpdpovdYxqII+eyG8wHwYDVR0jBBgwFoAU
# 7NfjgtJxXWRM3y5nP+e6mK4cD08wDgYDVR0PAQH/BAQDAgGGMBMGA1UdJQQMMAoG
# CCsGAQUFBwMIMHcGCCsGAQUFBwEBBGswaTAkBggrBgEFBQcwAYYYaHR0cDovL29j
# c3AuZGlnaWNlcnQuY29tMEEGCCsGAQUFBzAChjVodHRwOi8vY2FjZXJ0cy5kaWdp
# Y2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9vdEc0LmNydDBDBgNVHR8EPDA6MDig
# NqA0hjJodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkUm9v
# dEc0LmNybDAgBgNVHSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwDQYJKoZI
# hvcNAQELBQADggIBAH1ZjsCTtm+YqUQiAX5m1tghQuGwGC4QTRPPMFPOvxj7x1Bd
# 4ksp+3CKDaopafxpwc8dB+k+YMjYC+VcW9dth/qEICU0MWfNthKWb8RQTGIdDAiC
# qBa9qVbPFXONASIlzpVpP0d3+3J0FNf/q0+KLHqrhc1DX+1gtqpPkWaeLJ7giqzl
# /Yy8ZCaHbJK9nXzQcAp876i8dU+6WvepELJd6f8oVInw1YpxdmXazPByoyP6wCeC
# RK6ZJxurJB4mwbfeKuv2nrF5mYGjVoarCkXJ38SNoOeY+/umnXKvxMfBwWpx2cYT
# gAnEtp/Nh4cku0+jSbl3ZpHxcpzpSwJSpzd+k1OsOx0ISQ+UzTl63f8lY5knLD0/
# a6fxZsNBzU+2QJshIUDQtxMkzdwdeDrknq3lNHGS1yZr5Dhzq6YBT70/O3itTK37
# xJV77QpfMzmHQXh6OOmc4d0j/R0o08f56PGYX/sr2H7yRp11LB4nLCbbbxV7HhmL
# NriT1ObyF5lZynDwN7+YAN8gFk8n+2BnFqFmut1VwDophrCYoCvtlUG3OtUVmDG0
# YgkPCr2B2RP+v6TR81fZvAT6gt4y3wSJ8ADNXcL50CN/AAvkdgIm2fBldkKmKYcJ
# RyvmfxqkhQ/8mJb2VVQrH4D6wPIOK+XW+6kvRBVK5xMOHds3OBqhK/bt1nz8MIIG
# uTCCBKGgAwIBAgIRAJmjgAomVTtlq9xuhKaz6jkwDQYJKoZIhvcNAQEMBQAwgYAx
# CzAJBgNVBAYTAlBMMSIwIAYDVQQKExlVbml6ZXRvIFRlY2hub2xvZ2llcyBTLkEu
# MScwJQYDVQQLEx5DZXJ0dW0gQ2VydGlmaWNhdGlvbiBBdXRob3JpdHkxJDAiBgNV
# BAMTG0NlcnR1bSBUcnVzdGVkIE5ldHdvcmsgQ0EgMjAeFw0yMTA1MTkwNTMyMTha
# Fw0zNjA1MTgwNTMyMThaMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhBc3NlY28g
# RGF0YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBDb2RlIFNpZ25pbmcg
# MjAyMSBDQTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAJ0jzwQwIzvB
# RiznM3M+Y116dbq+XE26vest+L7k5n5TeJkgH4Cyk74IL9uP61olRsxsU/WBAElT
# MNQI/HsE0uCJ3VPLO1UufnY0qDHG7yCnJOvoSNbIbMpT+Cci75scCx7UsKK1fcJo
# 4TXetu4du2vEXa09Tx/bndCBfp47zJNsamzUyD7J1rcNxOw5g6FJg0ImIv7nCeNn
# 3B6gZG28WAwe0mDqLrvU49chyKIc7gvCjan3GH+2eP4mYJASflBTQ3HOs6JGdriS
# MVoD1lzBJobtYDF4L/GhlLEXWgrVQ9m0pW37KuwYqpY42grp/kSYE4BUQrbLgBMN
# KRvfhQPskDfZ/5GbTCyvlqPN+0OEDmYGKlVkOMenDO/xtMrMINRJS5SY+jWCi8PR
# HAVxO0xdx8m2bWL4/ZQ1dp0/JhUpHEpABMc3eKax8GI1F03mSJVV6o/nmmKqDE6T
# K34eTAgDiBuZJzeEPyR7rq30yOVw2DvetlmWssewAhX+cnSaaBKMEj9O2GgYkPJ1
# 6Q5Da1APYO6n/6wpCm1qUOW6Ln1J6tVImDyAB5Xs3+JriasaiJ7P5KpXeiVV/HIs
# W3ej85A6cGaOEpQA2gotiUqZSkoQUjQ9+hPxDVb/Lqz0tMjp6RuLSKARsVQgETwo
# NQZ8jCeKwSQHDkpwFndfCceZ/OfCUqjxAgMBAAGjggFVMIIBUTAPBgNVHRMBAf8E
# BTADAQH/MB0GA1UdDgQWBBTddF1MANt7n6B0yrFu9zzAMsBwzTAfBgNVHSMEGDAW
# gBS2oVQ5AsOgP46KvPrU+Bym0ToO/TAOBgNVHQ8BAf8EBAMCAQYwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwMAYDVR0fBCkwJzAloCOgIYYfaHR0cDovL2NybC5jZXJ0dW0u
# cGwvY3RuY2EyLmNybDBsBggrBgEFBQcBAQRgMF4wKAYIKwYBBQUHMAGGHGh0dHA6
# Ly9zdWJjYS5vY3NwLWNlcnR1bS5jb20wMgYIKwYBBQUHMAKGJmh0dHA6Ly9yZXBv
# c2l0b3J5LmNlcnR1bS5wbC9jdG5jYTIuY2VyMDkGA1UdIAQyMDAwLgYEVR0gADAm
# MCQGCCsGAQUFBwIBFhhodHRwOi8vd3d3LmNlcnR1bS5wbC9DUFMwDQYJKoZIhvcN
# AQEMBQADggIBAHWIWA/lj1AomlOfEOxD/PQ7bcmahmJ9l0Q4SZC+j/v09CD2csX8
# Yl7pmJQETIMEcy0VErSZePdC/eAvSxhd7488x/Cat4ke+AUZZDtfCd8yHZgikGuS
# 8mePCHyAiU2VSXgoQ1MrkMuqxg8S1FALDtHqnizYS1bIMOv8znyJjZQESp9RT+6N
# H024/IqTRsRwSLrYkbFq4VjNn/KV3Xd8dpmyQiirZdrONoPSlCRxCIi54vQcqKiF
# LpeBm5S0IoDtLoIe21kSw5tAnWPazS6sgN2oXvFpcVVpMcq0C4x/CLSNe0XckmmG
# sl9z4UUguAJtf+5gE8GVsEg/ge3jHGTYaZ/MyfujE8hOmKBAUkVa7NMxRSB1EdPF
# pNIpEn/pSHuSL+kWN/2xQBJaDFPr1AX0qLgkXmcEi6PFnaw5T17UdIInA58rTu3m
# efNuzUtse4AgYmxEmJDodf8NbVcU6VdjWtz0e58WFZT7tST6EWQmx/OoHPelE77l
# ojq7lpsjhDCzhhp4kfsfszxf9g2hoCtltXhCX6NqsqwTT7xe8LgMkH4hVy8L1h2p
# qGLT2aNCx7h/F95/QvsTeGGjY7dssMzq/rSshFQKLZ8lPb8hFTmiGDJNyHga5hZ5
# 9IGynk08mHhBFM/0MLeBzlAQq1utNjQprztZ5vv/NJy8ua9AGbwkMWkOMIIGwjCC
# BKqgAwIBAgIQBUSv85SdCDmmv9s/X+VhFjANBgkqhkiG9w0BAQsFADBjMQswCQYD
# VQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNVBAMTMkRpZ2lD
# ZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1waW5nIENBMB4X
# DTIzMDcxNDAwMDAwMFoXDTM0MTAxMzIzNTk1OVowSDELMAkGA1UEBhMCVVMxFzAV
# BgNVBAoTDkRpZ2lDZXJ0LCBJbmMuMSAwHgYDVQQDExdEaWdpQ2VydCBUaW1lc3Rh
# bXAgMjAyMzCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAKNTRYcdg45b
# rD5UsyPgz5/X5dLnXaEOCdwvSKOXejsqnGfcYhVYwamTEafNqrJq3RApih5iY2nT
# WJw1cb86l+uUUI8cIOrHmjsvlmbjaedp/lvD1isgHMGXlLSlUIHyz8sHpjBoyoNC
# 2vx/CSSUpIIa2mq62DvKXd4ZGIX7ReoNYWyd/nFexAaaPPDFLnkPG2ZS48jWPl/a
# Q9OE9dDH9kgtXkV1lnX+3RChG4PBuOZSlbVH13gpOWvgeFmX40QrStWVzu8IF+qC
# ZE3/I+PKhu60pCFkcOvV5aDaY7Mu6QXuqvYk9R28mxyyt1/f8O52fTGZZUdVnUok
# L6wrl76f5P17cz4y7lI0+9S769SgLDSb495uZBkHNwGRDxy1Uc2qTGaDiGhiu7xB
# G3gZbeTZD+BYQfvYsSzhUa+0rRUGFOpiCBPTaR58ZE2dD9/O0V6MqqtQFcmzyrzX
# xDtoRKOlO0L9c33u3Qr/eTQQfqZcClhMAD6FaXXHg2TWdc2PEnZWpST618RrIbro
# HzSYLzrqawGw9/sqhux7UjipmAmhcbJsca8+uG+W1eEQE/5hRwqM/vC2x9XH3mwk
# 8L9CgsqgcT2ckpMEtGlwJw1Pt7U20clfCKRwo+wK8REuZODLIivK8SgTIUlRfgZm
# 0zu++uuRONhRB8qUt+JQofM604qDy0B7AgMBAAGjggGLMIIBhzAOBgNVHQ8BAf8E
# BAMCB4AwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAKBggrBgEFBQcDCDAgBgNV
# HSAEGTAXMAgGBmeBDAEEAjALBglghkgBhv1sBwEwHwYDVR0jBBgwFoAUuhbZbU2F
# L3MpdpovdYxqII+eyG8wHQYDVR0OBBYEFKW27xPn783QZKHVVqllMaPe1eNJMFoG
# A1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2Vy
# dFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcmwwgZAGCCsG
# AQUFBwEBBIGDMIGAMCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5j
# b20wWAYIKwYBBQUHMAKGTGh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdp
# Q2VydFRydXN0ZWRHNFJTQTQwOTZTSEEyNTZUaW1lU3RhbXBpbmdDQS5jcnQwDQYJ
# KoZIhvcNAQELBQADggIBAIEa1t6gqbWYF7xwjU+KPGic2CX/yyzkzepdIpLsjCIC
# qbjPgKjZ5+PF7SaCinEvGN1Ott5s1+FgnCvt7T1IjrhrunxdvcJhN2hJd6PrkKoS
# 1yeF844ektrCQDifXcigLiV4JZ0qBXqEKZi2V3mP2yZWK7Dzp703DNiYdk9WuVLC
# tp04qYHnbUFcjGnRuSvExnvPnPp44pMadqJpddNQ5EQSviANnqlE0PjlSXcIWiHF
# tM+YlRpUurm8wWkZus8W8oM3NG6wQSbd3lqXTzON1I13fXVFoaVYJmoDRd7ZULVQ
# jK9WvUzF4UbFKNOt50MAcN7MmJ4ZiQPq1JE3701S88lgIcRWR+3aEUuMMsOI5lji
# tts++V+wQtaP4xeR0arAVeOGv6wnLEHQmjNKqDbUuXKWfpd5OEhfysLcPTLfddY2
# Z1qJ+Panx+VPNTwAvb6cKmx5AdzaROY63jg7B145WPR8czFVoIARyxQMfq68/qTr
# eWWqaNYiyjvrmoI1VygWy2nyMpqy0tg6uLFGhmu6F/3Ed2wVbK6rr3M66ElGt9V/
# zLY4wNjsHPW2obhDLN9OTH0eaHDAdwrUAuBcYLso/zjlUlrWrBciI0707NMX+1Br
# /wd3H3GXREHJuEbTbDJ8WC9nR2XlG3O2mflrLAZG70Ee8PBf4NvZrZCARK+AEEGK
# MIIG3TCCBMWgAwIBAgIQJJHKH5ST1/a4Ov2c99Kv6zANBgkqhkiG9w0BAQsFADBW
# MQswCQYDVQQGEwJQTDEhMB8GA1UEChMYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEu
# MSQwIgYDVQQDExtDZXJ0dW0gQ29kZSBTaWduaW5nIDIwMjEgQ0EwHhcNMjQwOTAy
# MTA1NzI2WhcNMjUwOTAyMTA1NzI1WjCBgjELMAkGA1UEBhMCREUxEDAOBgNVBAgM
# B0JhdmFyaWExDzANBgNVBAcMBk11bmljaDEeMBwGA1UECgwVT3BlbiBTb3VyY2Ug
# RGV2ZWxvcGVyMTAwLgYDVQQDDCdPcGVuIFNvdXJjZSBEZXZlbG9wZXIsIEp1bGlh
# biBQYXdsb3dza2kwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCpFBRX
# Dt2Y6dui3ez8PJ3ql0CL0Whb4gs+3leX+q9nG1SWVabkIGvrPlEUQdzfJvPZM4PF
# 6m339OJKuS06IQwruBvnSGx7mcEgm82kwa6b0GRrodnbUm8WssnGUytY60YncjCu
# SgqfKI610e43oJMa+SVZCtXPHbiwciehhm93OJXyDOcevJhT/SXhh3zsEIq+hd1Y
# U5+XkKFXWUKD7ucxLB0BxXkAm2cj9wzkuTBCCi8RqTxW+m5b7VEJXcbjMk3HDDn7
# 5gom05FvYNmDTa2/fFDgh3BDUbTwsjbrvEoDdOySgAR3TIbkEVQaT4V/vxjy0xgr
# 6zpW4nP8XCqfwyPj4gbWTnEUrWd350IZcZbCcxwSfQkEqze2CxQiO9H706FLa7qD
# jP2iZH0OfkwhXntTbTNrYkNo2YmfSBN1LUPWV6w4PGRy1iaBcr3wOa1Re8nK8pon
# 9SaFvOfSoHhrzBFE8WfpBuMLRqN0W2C14b8N6qPWcpY2FRpj2KKdxhlnGMBjPh/p
# KM4Ltqg5ZZwGP1f5g4hodfVueaOuJXyNr7YnyEqD5YLcYCO9ycKCw4fgZ/cfxGmZ
# 2zmkSgo4SHZjoroGEs3QxOZVk2WN4as3YDeMqlgKNUOl2K8GoVonyYze8+yf1AYb
# 99q3p5tcObm4USL/hPqjxub5B52ga6uCoV2IeQIDAQABo4IBeDCCAXQwDAYDVR0T
# AQH/BAIwADA9BgNVHR8ENjA0MDKgMKAuhixodHRwOi8vY2NzY2EyMDIxLmNybC5j
# ZXJ0dW0ucGwvY2NzY2EyMDIxLmNybDBzBggrBgEFBQcBAQRnMGUwLAYIKwYBBQUH
# MAGGIGh0dHA6Ly9jY3NjYTIwMjEub2NzcC1jZXJ0dW0uY29tMDUGCCsGAQUFBzAC
# hilodHRwOi8vcmVwb3NpdG9yeS5jZXJ0dW0ucGwvY2NzY2EyMDIxLmNlcjAfBgNV
# HSMEGDAWgBTddF1MANt7n6B0yrFu9zzAMsBwzTAdBgNVHQ4EFgQUJ0fj+UjpDK75
# NKerDl0427zkmskwSwYDVR0gBEQwQjAIBgZngQwBBAEwNgYLKoRoAYb2dwIFAQQw
# JzAlBggrBgEFBQcCARYZaHR0cHM6Ly93d3cuY2VydHVtLnBsL0NQUzATBgNVHSUE
# DDAKBggrBgEFBQcDAzAOBgNVHQ8BAf8EBAMCB4AwDQYJKoZIhvcNAQELBQADggIB
# AJHuX55vDsuy3Ct90pCbzMtCQvuhA7RPPl+UC8JGheL23Dpx/2bFSoNpEPJK/bWe
# 0ah2gTlmMIp+92Tw45JesBY7jLUatkK8jIea+CwmGZSEV6lqY9nzX32nbpH0TLtk
# H72M6ns2A8pszloRdJ+GwhkyqIBXFWlGu5wx0rQQ258JarkF93syl4OXwQKeYhfP
# hSh0WL46+C3Nh90QH55T41/yb2RMl1PNqSv4n2Ev+e27mNIjlPUvvVqRqAYaM6OU
# LaJI+YOLQRPGPv7Np/oqnZ7Md5rMEH+v9TGERgEqDSRtuvfY1Te69HZhr2kFKDMm
# z4NSBy9YmONsgIMHSxP/PKZVKyxH7stwrzabbLyRdvSi+oHRVdG8KGilZ5ztsgOQ
# HJO88CQTaajLzzrshtf0eTEbkuXAB9RELw8eT2b4FD9/QTpx+63yPIz1O9jlY/cP
# /+LrdaysvDBtZOrAF9Gc8bgwG803w8Me3R3eN577NspJMxnoNlg2ZCmKnA+oq+YS
# NoYFNUEK8tdyD/JOv1HTpJRS61nevrNZ8MkH1vMFRx7aItRJSipIoYjaHcENm2Iq
# 3s5Sjvh+Qx9PibzZK/1ixwikthixF88uF7YxRWYbMqQO936+gAejVy3J+WetojTc
# zLRPmpC4mLDEKNOdFRq2/2VTQcizgn1/CUkbU5yUg57mMYIGQDCCBjwCAQEwajBW
# MQswCQYDVQQGEwJQTDEhMB8GA1UEChMYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEu
# MSQwIgYDVQQDExtDZXJ0dW0gQ29kZSBTaWduaW5nIDIwMjEgQ0ECECSRyh+Uk9f2
# uDr9nPfSr+swDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIBDDEKMAigAoAA
# oQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgtNx+kEXk6Ppgesf1pEDDxlP+
# +/h6lNF0Bex1OfiJrUYwDQYJKoZIhvcNAQEBBQAEggIAeUmvJaG8DxFW/O7jF5Kg
# F17LBDMyzaUhzlEgoyjYzEnp33GIw+MdP/UVj/iTdtQ5foUPxXpELxkzNAdKnHtg
# 1nNO5zcO/V//tqZblldsi6m0KOmt9aPD+KVa7ipp4IpFJXJy3AEMTCCwZPr7Xd3n
# OAuQ3Kc4nJc+ekPz48JqFj7ZuBICPlHJx+wL2jNOm4TsWmzjCDYUy4Wy/XPTvAiy
# 850ctoOR1FuRYk2ebI/A+KAbrL7sUBPuRS1i3/OZt4eqtrYpQ4/f3Fm59b34SSpM
# 78dNDsNsjaXQwVHz6cAQQEu869OmlJNBzL0QQ4BzNO/RCI6sFr+XJXznzd5iQL/0
# qGJhvn5KfkB3Xbdh2v3k6qXnfKTI3zqh1rKhVDWjgFSKvGB3TD4/QMnprbbJS71C
# lcJrMmZb3aaJweZn8UChWnsGAGZ20b2/XuigHTkAOON42vr5ktAvVcWO3pgoL5RQ
# 9N0i87Ag6gqAAjl00i8LgeMhBes6qDACiSITfAiRADhid5Y3lZ/yVeev1UDa9K7O
# d66AO1bJQGPdVmU3AbLU6tLpG1BV81A58UZC2bmwp/GxEP4hXrucHF8AhP4Ak1IV
# 2Q5QH80rimvEpTjtIIxmg+RSS6hgkfwOm7HKs+CRVD3pJOOaO1K9ltz47Q4qIWsb
# 5kwNDytVQhHFJRU3zW7c4gyhggMgMIIDHAYJKoZIhvcNAQkGMYIDDTCCAwkCAQEw
# dzBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4xOzA5BgNV
# BAMTMkRpZ2lDZXJ0IFRydXN0ZWQgRzQgUlNBNDA5NiBTSEEyNTYgVGltZVN0YW1w
# aW5nIENBAhAFRK/zlJ0IOaa/2z9f5WEWMA0GCWCGSAFlAwQCAQUAoGkwGAYJKoZI
# hvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcNMjQwOTMwMDg1MzU3
# WjAvBgkqhkiG9w0BCQQxIgQgWessoIjoI8AtWghCp1qrZbMgLeDFzPZH8vMGvThy
# mRgwDQYJKoZIhvcNAQEBBQAEggIAHKyzbaBjMzyHQb1bxhf7BbK+pWEPDDdviqq2
# vz1dMBzyxw7nO0CqWbDTKu0U2VXLMl1koGINLve6t2/yWVVu6ZkeWJ3avNoClZRU
# EkQ9QFD/pNQLUYJ1TrHwmuea9vuTRuXz9wWG/za0UVxKpXmyR/TmRtW3JI0dTNvT
# z15L0kWmXKwUPYq1VNTIP0GsARaopAVcP1gr2AoG475gYVHoSeg8agmkr+B+xhNb
# QIzso/r1D76OLR472IF16q0AD80GZ8byNOpsfc8z24fGssa6Np2k3RRyf6arkDsB
# qsCimi95vSMXe2r5iKOMjL7MYDllTCQyiAFAgRPiNBO0JI/dCU929L20TvJMIQJ8
# QwkShIXSKsGUj0XKda7EvK+k7Mq2yPL9Dw1+qMJrSOtrm6pqbUh7wRr0HFIejBBx
# br3aV7H5+48Q9mFql4JUi11oKgEhdDKRg4IXseDtKHrCabf3nQV5/Ra43+D+mn9l
# 7fnMZYBb7csrj423dVMfnjPcLM18cLoSR2YLe21dICBCnKkbGWfJAtY+mtCOfaZS
# 7m3fJh7zNG3+2z7el+PY+eCCU/PVD2KYaNMpWaWkIZP85FX93JGAbQGtj/rvsaYm
# p7n1pUCt1MOA+9ZuRS5mZCqou7DGq2puWDIxLaAIH3eKgGKHzfmsOLASlXZYa5bp
# jrZsScQ=
# SIG # End signature block
