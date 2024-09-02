#!/usr/bin/env pwsh

<#PSScriptInfo

.VERSION 1.3.0

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
    Version 1.3.0 (2024-09-02)
    - Move to separate repository.
    - Use personal shortlink service.
    - Add code signature.
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
                    $versionNumber = $matches[1].Trim()
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
        switch ($format) {
            'xz' { return $tarVersionOutput -match 'liblzma' }
            'bzip2' { return $tarVersionOutput -match 'bz2lib' }
            'gz' { return $tarVersionOutput -match 'zlib' }
            default { return $false }
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
            'tar.xz'  = '{0} -xJf "{1}" -C "{2}"'
            'tar.bz2' = '{0} -xjf "{1}" -C "{2}"'
            'tar.gz'  = '{0} -xzf "{1}" -C "{2}"'
            'tar'     = '{0} -xf "{1}" -C "{2}"'
            'xz'      = '{0} -d "{1}" -C "{2}"'
            '7z'      = '{0} x "{1}" -o"{2}"'
            'bzip2'   = '{0} -d "{1}" -C "{2}"'
            'gzip'    = '{0} -d "{1}" -C "{2}"'
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
            $commandTemplate = $commandTemplates[$FileExtension]
            $command = $commandTemplate -f $Executable, $SourceFile, $DestinationFolder
            Write-Verbose "Running command: $command"

            # Split the command into the executable and its arguments
            $commandParts = $command -split ' ', 2
            $arguments = if ($commandParts.Length -gt 1) { $commandParts[1] } else { "" }

            # Execute the external command
            Start-Process -FilePath $Executable -ArgumentList $arguments -NoNewWindow -Wait
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
                    $sha256 = $matches[1]
                    $fileName = $matches[2].Trim()
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
            $fontDestinationFolderPath = "${xdgDataHome}/fonts"
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
    $archivePreferenceOrder = @('tar.xz', '7z', 'tar.bz2', 'tar.gz', 'zip', 'tar')

    # ZIP is natively supported in PowerShell
    [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'zip'; Executable = 'powershell' })

    if ($IsMacOS -or $IsLinux) {
        # Prefer tar if available
        if (Get-Command tar -ErrorAction Ignore) {
            if (Test-TarSupportsFormat 'xz') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.xz'; Executable = 'tar' }) }
            if (Test-TarSupportsFormat 'bzip2') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.bz2'; Executable = 'tar' }) }
            if (Test-TarSupportsFormat 'gz') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.gz'; Executable = 'tar' }) }
            [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar'; Executable = 'tar' })
        }
        # Check for individual tools
        if (Get-Command xz -ErrorAction Ignore) { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.xz'; Executable = 'xz' }) }
        if (Get-Command 7z -ErrorAction Ignore) { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = '7z'; Executable = '7z' }) }
        if (Get-Command bzip2 -ErrorAction Ignore) { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.bz2'; Executable = 'bzip2' }) }
        if (Get-Command gzip -ErrorAction Ignore) { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.gz'; Executable = 'gzip' }) }
    }
    else {
        if (Get-Command tar -ErrorAction Ignore) {
            if (Test-TarSupportsFormat 'xz') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.xz'; Executable = 'tar' }) }
            if (Test-TarSupportsFormat 'bzip2') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.bz2'; Executable = 'tar' }) }
            if (Test-TarSupportsFormat 'gz') { [void]$supportedArchiveFormats.Add([pscustomobject]@{FileExtension = 'tar.gz'; Executable = 'tar' }) }
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
        if ($Verbose) {
            Write-Verbose "Refreshing font cache"
            fc-cache -fv
        }
        else {
            fc-cache -f
        }
    }
}
