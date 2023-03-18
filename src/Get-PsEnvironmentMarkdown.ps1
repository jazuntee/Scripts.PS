<#PSScriptInfo

.VERSION 1.0

.GUID c36c0886-c843-430c-b220-c2d833fed680

.AUTHOR Jason Thompson

.COMPANYNAME Microsoft Corporation

.COPYRIGHT (c) 2023 Jason Thompson. All rights reserved.

.TAGS Microsoft Windows PowerShell Environment Module Variable Setting PSEdition_Desktop PSEdition_Core Windows Linux MacOS

.LICENSEURI https://raw.githubusercontent.com/jasoth/Scripts.PS/main/LICENSE

.PROJECTURI https://github.com/jasoth/Scripts.PS

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

.PRIVATEDATA

#> 

<# 
.SYNOPSIS
    Format PowerShell Environment Settings, Modules, Variables, and Errors as Markdown.
.DESCRIPTION 
    Format PowerShell Environment Settings, Modules, Variables, and Errors as Markdown.
.EXAMPLE
    PS > Get-PsEnvironmentMarkdown.ps1

    Format PowerShell Environment Settings, Modules, Variables, and Errors as Markdown.

.EXAMPLE
    PS >iex (irm 'https://aka.ms/Get-PsEnvironmentMarkdown') | Set-Clipboard

    Invoke this script directly from GitHub using Invoke-Expression without parameters and copy to clipboard.

.EXAMPLE
    PS >& ([scriptblock]::Create((irm 'https://aka.ms/Get-PsEnvMd'))) -ModuleName 'MyModule' | Set-Clipboard

    Invoke this script directly from GitHub using Call operator with parameters and copy to clipboard.

#>
param (
    # Names of modules to include in the output
    [Parameter(Mandatory = $false)]
    [string[]] $ModuleName,
    # Names of environment variables to include in the output
    [Parameter(Mandatory = $false)]
    [string[]] $EnvironmentVariableName,
    # Names of variables to include in the output
    [Parameter(Mandatory = $false)]
    [string[]] $VariableName,
    # Include additional detail in the output
    [Parameter(Mandatory = $false)]
    [switch] $Full
)


#region Supporting Functions

<#
.SYNOPSIS
    Get object property value in a manner that satifies strict mode.

.EXAMPLE
    PS >$object = New-Object psobject -Property @{ title = 'title value' }
    PS >$object | Get-PropertyValue -Property 'title'

    Get value of object property named title.

.EXAMPLE
    PS >$object = New-Object psobject -Property @{ lvl1 = (New-Object psobject -Property @{ nextLevel = 'lvl2 data' }) }
    PS >Get-PropertyValue $object -Property 'lvl1', 'nextLevel'

    Get value of nested object property named nextLevel.

.INPUTS
    System.Collections.IDictionary
    System.Management.Automation.PSObject

.LINK
    https://github.com/jasoth/Utility.PS
#>
function Get-PropertyValue {
    [CmdletBinding()]
    [OutputType([psobject])]
    param (
        # Object containing property values
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [AllowNull()]
        [psobject] $InputObjects,
        # Name of property. Specify an array of property names to tranverse nested objects.
        [Parameter(Mandatory = $true, ValueFromRemainingArguments = $true)]
        [string[]] $Property
    )

    process {
        foreach ($InputObject in $InputObjects) {
            for ($iProperty = 0; $iProperty -lt $Property.Count; $iProperty++) {
                ## Get property value
                if ($InputObject -is [System.Collections.IDictionary]) {
                    if ($InputObject.Contains($Property[$iProperty])) {
                        $PropertyValue = $InputObject[$Property[$iProperty]]
                    }
                    else { $PropertyValue = $null }
                }
                else {
                    $PropertyValue = Select-Object -InputObject $InputObject -ExpandProperty $Property[$iProperty] -ErrorAction Ignore
                }
                ## Check for more nested properties
                if ($iProperty -lt $Property.Count - 1) {
                    $InputObject = $PropertyValue
                    if ($null -eq $InputObject) { break }
                }
                else {
                    Write-Output $PropertyValue
                }
            }
        }
    }
}

<#
.SYNOPSIS
    Converts an object to a markdown table.

.EXAMPLE
    PS >ConvertTo-MarkdownTable $PsVersionTable

    Converts the PsVersionTable variable object to markdown table.

.EXAMPLE
    PS >Get-PSHostProcessInfo | ConvertTo-MarkdownTable -Compact

    Converts PSHostProcessInfo objects to markdown table.

.INPUTS
    System.Object

.LINK
    https://github.com/jasoth/Utility.PS
#>
function ConvertTo-MarkdownTable {
    [CmdletBinding()]
    [OutputType([string])]
    param (
        # Objects to convert
        [Parameter(Mandatory = $true, Position = 0, ValueFromPipeline = $true)]
        [object[]] $InputObject,
        # Property names to include in the output.
        [Parameter(Mandatory = $false, Position = 1)]
        [string[]] $Property,
        # Output one row per input object or one keypair list table per input object
        [Parameter(Mandatory = $false)]
        [ValidateSet('Table', 'List')]
        [string] $As,
        # Do not include whitespace padding in table
        [Parameter(Mandatory = $false)]
        [switch] $Compact,
        # String to use as delimiter for array values
        [Parameter(Mandatory = $false)]
        [string] $ArrayDelimiter,
        # Format second level depth objects with the specified format
        [Parameter(Mandatory = $false)]
        [ValidateSet('ToString', 'PsFormat', 'Html')]
        [string] $ObjectFormat = 'PsFormat'
    )

    begin {
        ## Initalize variables
        $NewLineReplacement = '<br>'

        function FormatMarkdownTableHeaderRow ($ColumnWidths) {
            if ($ColumnWidths.Count -gt 0) {
                $InitialColumn = $true
                [string]$TableRow = '| '
                [string]$DelimiterRow = '| '
                foreach ($PropertyName in $ColumnWidths.Keys) {
                    if (!$InitialColumn) { $TableRow += ' | ' }
                    $TableRow += $PropertyName.PadRight($ColumnWidths[$PropertyName], ' ')

                    if (!$InitialColumn) { $DelimiterRow += ' | ' }
                    if ($ColumnWidths[$PropertyName] -gt 0) {
                        $DelimiterRow += '---'.PadRight($ColumnWidths[$PropertyName], '-')
                    }
                    else {
                        $DelimiterRow += '---' #.PadRight($PropertyName.Length, '-')
                    }

                    $InitialColumn = $false
                }
                $TableRow += ' |'
                $DelimiterRow += ' |'

                $TableRow
                $DelimiterRow
            }
        }

        function FormatMarkdownTableRow ($ColumnWidths, $InputObject) {
            $InitialColumn = $true
            [string]$TableRow = '| '
            foreach ($PropertyName in $ColumnWidths.Keys) {
                if (!$InitialColumn) { $TableRow += ' | ' }

                if ($InputObject) {
                    $StringValue = ''

                    $PropertyValue = Get-PropertyValue $InputObject $PropertyName
                    $StringValue = Transform $PropertyValue
                    
                    $TableRow += $StringValue.PadRight($ColumnWidths[$PropertyName], ' ')
                }

                $InitialColumn = $false
            }
            $TableRow += ' |'

            return $TableRow
        }

        function FormatMarkdownKeyPairRows ($ColumnWidths, $InputObject) {
            foreach ($Property in $InputObject.PSObject.Properties) {
                $InitialColumn = $true
                [string]$TableRow = '| '

                foreach ($PropertyName in $ColumnWidths.Keys) {
                    if (!$InitialColumn) { $TableRow += ' | ' }

                    if ($InputObject) {
                        if ($PropertyName -eq 'Name') {
                            $TableRow += $Property.Name.PadRight($ColumnWidths['Name'], ' ')
                        }
                        else {
                            $StringValue = ''

                            $PropertyValue = $Property.Value
                            $StringValue = Transform $PropertyValue
                    
                            $TableRow += $StringValue.PadRight($ColumnWidths['Value'], ' ')
                        }
                    }

                    $InitialColumn = $false
                }
                $TableRow += ' |'

                Write-Output $TableRow
            }
            
        }

        function Transform ($PropertyValue) {
            $StringValue = ''
            if ($null -ne $PropertyValue) {
                if ($ArrayDelimiter -ne '' -and $PropertyValue -is [System.Collections.IList]) {
                    [array]$ArrayObject = New-Object -TypeName object[] -ArgumentList $PropertyValue.Count  # ConstrainedLanguage safe
                    for ($i = 0; $i -lt $PropertyValue.Count; $i++) {
                        $ArrayObject[$i] = $PropertyValue[$i].ToString()
                        if (!$ArrayObject[$i]) { $ArrayObject[$i] = $PropertyValue[$i].psobject.TypeNames[0] }
                    }
                    $StringValue = ($ArrayObject -join $ArrayDelimiter)
                }
                elseif ($PropertyValue -is [System.Collections.IDictionary] -or $PropertyValue -is [psobject]) {
                    if ($PropertyValue -is [System.Collections.IDictionary]) {
                        $PropertyValue = New-Object -TypeName PSObject -Property $PropertyValue  # ConstrainedLanguage safe
                    }
                    
                    if ($ObjectFormat -eq 'PsFormat') {
                        $FormattedObject = $PropertyValue | Format-List | Out-String -Width 2147483647
                        $StringValue = $FormattedObject.Trim("`r", "`n")
                    }
                    elseif ($ObjectFormat -eq 'Html') {
                        $HtmlTable = $PropertyValue | ConvertTo-Html -Fragment -As List
                        $StringValue = $HtmlTable -join ''
                    }
                    else {
                        $StringValue = $PropertyValue.ToString()
                        if (!$StringValue) { $StringValue = $PropertyValue.psobject.TypeNames[0] }
                    }
                }
                else {
                    $StringValue = $PropertyValue.ToString()
                }
            }
            $StringValue = $StringValue.Replace('\', '\\').Replace('|', '\|') # Escape backslash and pipe characters
            $StringValue = $StringValue -replace '(?<=[>])[\r\n]+(?=[<])', '' # Remove newlines between html tags
            $StringValue = $StringValue -replace '[\r\n]+', $NewLineReplacement # Replace newlines

            return $StringValue
        }

        $TableObjects = @()
    }

    process {
        foreach ($_InputObject in $InputObject) {
            ## Convert dictionaries
            if ($_InputObject -is [System.Collections.IDictionary]) {
                $_InputObject = New-Object -TypeName PSObject -Property $_InputObject  # ConstrainedLanguage safe
            }
            
            if ($Property) {
                $OutputObject = Select-Object -InputObject $_InputObject -Property $Property
            }
            else {
                $OutputObject = Select-Object -InputObject $_InputObject -Property "*"
            }

            $TableObjects += $OutputObject
        }
    }

    end {
        
        if (!$As) {
            if ($TableObjects.Count -gt 1) { $As = 'Table' }
            else { $As = 'List' }
        }

        if ($As -eq 'List') {
            foreach ($ObjectTable in $TableObjects) {
                ## Get column names and widths
                $KeyPairWidths = [ordered]@{ Name = 0; Value = 0 }
                foreach ($objProperty in $ObjectTable.PSObject.Properties) {
                    if (!$Compact -and $KeyPairWidths['Name'] -lt $objProperty.Name.Length) {
                        $KeyPairWidths['Name'] = $objProperty.Name.Length
                    }

                    $PropertyValue = Transform $objProperty.Value
                    if (!$Compact -and $null -ne $PropertyValue) {
                        if ($KeyPairWidths['Value'] -lt $PropertyValue.Length) {
                            $KeyPairWidths['Value'] = $PropertyValue.Length
                        }
                    }
                }

                ## Output Header and Separator Rows
                FormatMarkdownTableHeaderRow $KeyPairWidths
                ## Output Object Rows
                FormatMarkdownKeyPairRows $KeyPairWidths $ObjectTable
                ''
            }
        }
        else {
            ## Get column names and widths
            $ColumnWidths = [ordered]@{}
            foreach ($ObjectRow in $TableObjects) {
                foreach ($objProperty in $ObjectRow.PSObject.Properties) {
                    if ($Compact) {
                        $ColumnWidths[$objProperty.Name] = 0
                    }
                    elseif ($null -eq $ColumnWidths[$objProperty.Name]) {
                        $ColumnWidths[$objProperty.Name] = $objProperty.Name.Length
                    }
                    
                    $PropertyValue = Transform $objProperty.Value
                    if (!$Compact -and $null -ne $PropertyValue) {
                        if ($ColumnWidths[$objProperty.Name] -lt $PropertyValue.Length) {
                            $ColumnWidths[$objProperty.Name] = $PropertyValue.Length
                        }
                    }
                }
            }

            ## Output Header and Separator Rows
            FormatMarkdownTableHeaderRow $ColumnWidths

            ## Output Object Rows
            foreach ($ObjectRow in $TableObjects) {
                FormatMarkdownTableRow $ColumnWidths $ObjectRow
            }
        }

    }
}

<#
.SYNOPSIS
    Get the strict mode version of the current session scope.
    
.DESCRIPTION
    Get the strict mode version of the current session scope.
    1.0
        Prohibits references to uninitialized variables, except for uninitialized variables in strings.
    2.0
        Prohibits references to uninitialized variables. This includes uninitialized variables in strings.
        Prohibits references to non-existent properties of an object.
        Prohibits function calls that use the syntax for calling methods.
    3.0
        Prohibits references to uninitialized variables. This includes uninitialized variables in strings.
        Prohibits references to non-existent properties of an object.
        Prohibits function calls that use the syntax for calling methods.
        Prohibit out of bounds or unresolvable array indexes.

.EXAMPLE
    PS > Get-StrictModeVersion

    Get the strict mode version of the current session scope.

.INPUTS
    None

.LINK
    https://github.com/jasoth/Utility.PS
#>
function Get-StrictModeVersion {
    [CmdletBinding()]
    [OutputType([version])]
    param ()

    try { $null = @()[0] }
    catch { return [version]'3.0' }

    try { $null = $null.NonExistentProperty }
    catch { return [version]'2.0' }

    try { $null = $UninitializedVariable }
    catch { return [version]'1.0' }

    return [version]'0.0'
}

<#
.SYNOPSIS
    Create new HTML details summary block.
#>
function New-HtmlSummary ($Name, $Body) {
    @"
<details><summary>$Name</summary>
$Body
</details>
"@
}

#endregion


## Collect Requested Module and Error Data
[array]$ImportedModules = Get-Module
[array]$ImportedModulesSelected = $null
[array]$RequestedModules = $null
$ModuleErrors = @()
$CommandHistory = @()
if ($ModuleName) {
    $ImportedModulesSelected = Get-Module $ModuleName
    if ($null -eq $ImportedModulesSelected) { $ImportedModulesSelected = @() }
    
    $AvailableModulesSelected = Get-Module $ModuleName -ListAvailable
    if ($null -eq $AvailableModulesSelected) { $AvailableModulesSelected = @() }
    
    $RequestedModules = Compare-Object $AvailableModulesSelected -DifferenceObject $ImportedModulesSelected -IncludeEqual -Property Name, ModuleBase -PassThru | Select-Object @{ Name = 'Imported'; Expression = { $_.SideIndicator -ne '<=' } }, * -ExcludeProperty SideIndicator | Sort-Object Name, @{ Expression = 'Version'; Descending = $true }, @{ Expression = { Get-PropertyValue $_ PrivateData PSData Prerelease }; Descending = $true }

    [string]$regexKeywords = $ModuleName -join '|'
    foreach ($Module in $ImportedModulesSelected) {
        if ($regexKeywords) { $regexKeywords += '|' }
        $regexKeywords += $Module.ExportedCommands.Keys -join '|'
    }

    if ($regexKeywords) {
        [array]$ErrorList = $Error
        if ($null -eq $ErrorList) { $ErrorList = @() }
        
        for ($i = 0; $i -lt $ErrorList.Count; $i++) {
            $ErrorRecord = $null
            if ($ErrorList[$i] -is [System.Management.Automation.ErrorRecord]) {
                $ErrorRecord = $ErrorList[$i]
            }
            elseif ($ErrorList[$i] -is [System.Management.Automation.CmdletInvocationException]) {
                $ErrorRecord = $ErrorList[$i].ErrorRecord
            }

            if ($ErrorRecord) {
                $ErrorInvocationModuleName = Get-PropertyValue $ErrorRecord InvocationInfo MyCommand ModuleName
                if ($ErrorInvocationModuleName -in $ModuleName -or $ErrorRecord.TargetObject -match $regexKeywords -or $ErrorRecord.ScriptStackTrace -match $regexKeywords) {
                    # if ($ErrorRecord.InvocationInfo.HistoryId -gt 0) {
                    #     $HistoryInfo = $CommandHistory | Where-Object Id -EQ $ErrorRecord.InvocationInfo.HistoryId
                    #     if (!$HistoryInfo) {
                    #         $HistoryInfo = Get-History $ErrorRecord.InvocationInfo.HistoryId | Add-Member Errors -MemberType NoteProperty -Value @() -PassThru
                    #         $CommandHistory += $HistoryInfo
                    #     }
                    #     $HistoryInfo.Errors += $ErrorRecord
                    # }0
                    $ModuleErrors += $ErrorRecord
                }
            }
        }
    }
}
[array]::Reverse($ModuleErrors)
[array]::Reverse($CommandHistory)

## Collect Requested Variable Data
[array]$VariablesSelected = @()
$VariablesSelected += $EnvironmentVariableName | ForEach-Object { if ($_) { "Env:$_" } } | Get-Item -ErrorAction Ignore | Select-Object @{ Name = 'Variable Name'; Expression = { '$Env:{0}' -f $_.Name } }, Value
$VariablesSelected += $VariableName | ForEach-Object { if ($_) { "Variable:$_" } } | Get-Item -ErrorAction Ignore | Select-Object @{ Name = 'Variable Name'; Expression = { '${0}' -f $_.Name } }, Value, @{ Name = 'Global Value'; Expression = { Get-Variable $_.Name -Scope Global -ValueOnly -ErrorAction Ignore } }


### Collect Additional Data and Structure Markdown Output into Objects
$Summary = [ordered]@{
    'Operating System'   = [System.Environment]::OSVersion.VersionString
    'PowerShell Version' = $PSVersionTable.PSVersion.ToString()
}

$AdditionalDetail = [ordered]@{
    PowerShellDistributionChannel = Get-Item 'Env:POWERSHELL_DISTRIBUTION_CHANNEL' -ErrorAction Ignore | Select-Object -ExpandProperty Value
    LanguageMode                  = $ExecutionContext.SessionState.LanguageMode
    ExecutionPolicy               = Get-ExecutionPolicy
    ProfileScripts                = ((Test-Path $Profile.AllUsersAllHosts) -and (Get-Content $Profile.AllUsersAllHosts -Raw)) -or ((Test-Path $Profile.AllUsersCurrentHost) -and (Get-Content $Profile.AllUsersCurrentHost -Raw)) -or ((Test-Path $Profile.CurrentUserAllHosts) -and (Get-Content $Profile.CurrentUserAllHosts -Raw)) -or ((Test-Path $Profile.CurrentUserCurrentHost) -and (Get-Content $Profile.CurrentUserCurrentHost -Raw))
    PSDefaultParameterValues      = (Get-Variable PSDefaultParameterValues -ValueOnly -ErrorAction Ignore).Keys | ForEach-Object { $_ }
}
if ($Full) {
    $AdditionalDetail += [ordered]@{
        StrictMode                    = Get-StrictModeVersion
        PSModuleAutoLoadingPreference = Get-Variable 'PSModuleAutoLoadingPreference' -ValueOnly -ErrorAction Ignore
        PSModulePath                  = Get-Item 'Env:PSModulePath' -ErrorAction Ignore | ForEach-Object { if ($_.Value -is [string]) { $_.Value.Split([System.IO.Path]::PathSeparator) } }
        Path                          = Get-Item 'Env:Path' -ErrorAction Ignore | ForEach-Object { if ($_.Value -is [string]) { $_.Value.Split([System.IO.Path]::PathSeparator) } }
        PathExt                       = Get-Item 'Env:PATHEXT' -ErrorAction Ignore | Select-Object -ExpandProperty Value
    }
}

$EnvironmentVariables = $null
if ($Full) {
    $EnvironmentVariables = Get-Item "Env:*" -ErrorAction Ignore
}

$PreferenceVariableName = @(
    'PSModuleAutoLoadingPreference'
    'PSNativeCommandArgumentPassing'
    'PSNativeCommandUseErrorActionPreference'
    'ErrorActionPreference'
    'WarningPreference'
    'InformationPreference'
    'VerbosePreference'
    'DebugPreference'
    'ProgressPreference'
    'ConfirmPreference'
    'WhatIfPreference'
)
$PreferenceVariables = $PreferenceVariableName | ForEach-Object {
    [ordered]@{
        'Variable Name' = $_
        'Value'         = Get-Variable $_ -ValueOnly -ErrorAction Ignore
        'Global Value'  = Get-Variable $_ -Scope Global -ValueOnly -ErrorAction Ignore
    }
}

$LoadedAssemblies = $null
if ($Full) {
    $LoadedAssemblies = try { [System.AppDomain]::CurrentDomain.GetAssemblies() } catch { }
}

### Produce Markdown/HTML String Output
$strSummary = $Summary | ConvertTo-MarkdownTable -As List | Out-String | ForEach-Object { "`r`n{0}" -f $_.Trim("`r", "`n") }
$strSummaryModules = $ImportedModulesSelected | Select-Object @{ Name = 'Module Name'; Expression = { $_.Name } }, @{ Name = 'Version'; Expression = { if (Get-PropertyValue $_ PrivateData PSData Prerelease) { '{0}-{1}' -f $_.Version, $_.PrivateData.PSData['Prerelease'] } else { $_.Version.ToString() } } }, @{ Name = 'PSGallery'; Expression = { $_.RepositorySourceLocation -eq 'https://www.powershellgallery.com/api/v2' } } | ConvertTo-MarkdownTable -As Table -ArrayDelimiter "; " | Out-String | ForEach-Object { "`r`n{0}" -f $_.Trim("`r", "`n") }
$strSummaryVariables = $VariablesSelected | ConvertTo-MarkdownTable -As Table -ArrayDelimiter "`r`n" -Compact | Out-String | ForEach-Object { "`r`n{0}" -f $_.Trim("`r", "`n") }
$strAdditionalDetail = $AdditionalDetail | ConvertTo-MarkdownTable -As List -ArrayDelimiter "`r`n" -Compact | Out-String | ForEach-Object { "`r`n{0}" -f $_.Trim("`r", "`n") }
$strEnvironmentVariables = $EnvironmentVariables | Select-Object Name, Value | ConvertTo-MarkdownTable -As Table -Compact | Out-String | ForEach-Object { "`r`n{0}" -f $_.Trim("`r", "`n") }
$strPreferenceVariables = $PreferenceVariables | ConvertTo-MarkdownTable -As Table | Out-String | ForEach-Object { "`r`n{0}" -f $_.Trim("`r", "`n") }
$strImportedModules = $ImportedModules | Select-Object Name, Guid, Version, @{ Name = 'Prerelease'; Expression = { Get-PropertyValue $_ PrivateData PSData Prerelease } }, Path, RepositorySourceLocation | Format-Table -AutoSize | Out-String -Width 2147483646 | ForEach-Object { "`r`n```````r`n{0}`r`n``````" -f $_.Trim("`r", "`n") }
$strAvailableModules = $RequestedModules | Select-Object Imported, Name, Guid, Version, @{ Name = 'Prerelease'; Expression = { Get-PropertyValue $_ PrivateData PSData Prerelease } }, Path, RepositorySourceLocation | Format-Table -AutoSize | Out-String -Width 2147483646 | ForEach-Object { "`r`n```````r`n{0}`r`n``````" -f $_.Trim("`r", "`n") }
$strLoadedAssemblies = $LoadedAssemblies | Select-Object @{ Name = 'GAC'; Expression = { $_.GlobalAssemblyCache } }, FullName, Location, ImageRuntimeVersion | Format-Table -AutoSize | Out-String -Width 2147483646 | ForEach-Object { "`r`n```````r`n{0}`r`n``````" -f $_.Trim("`r", "`n") }
$strCommandHistory = $CommandHistory | Select-Object Id, @{ Name = 'Duration'; Expression = { $_.Duration.TotalSeconds } }, CommandLine | ConvertTo-MarkdownTable -As Table | Out-String | ForEach-Object { "`r`n{0}" -f $_.Trim("`r", "`n") }

$PrevErrorView = $ErrorView
$ErrorView = 'NormalView'
$strErrorRecords = $ModuleErrors | ForEach-Object { if ($_.InvocationInfo.HistoryId -gt 0) { $History = Get-History $_.InvocationInfo.HistoryId; '({1:0.000}s) PS > {0}' -f $History.CommandLine, $History.Duration.TotalSeconds }; $_; $_.ScriptStackTrace; '' } | Out-String -Width 2147483646 | ForEach-Object { "`r`n```````r`n{0}`r`n``````" -f $_.Trim("`r", "`n") }
$ErrorView = $PrevErrorView

[array]$OutputString = $strSummary
if ($ImportedModulesSelected) { $OutputString += $strSummaryModules }
if ($VariablesSelected) { $OutputString += $strSummaryVariables }
if ($AdditionalDetail) { $OutputString += New-HtmlSummary 'Additional Detail' $strAdditionalDetail }
if ($EnvironmentVariables) { $OutputString += New-HtmlSummary 'Environment Variables' $strEnvironmentVariables }
if ($PreferenceVariables) { $OutputString += New-HtmlSummary 'Preference Variables' $strPreferenceVariables }
if ($ImportedModules) { $OutputString += New-HtmlSummary 'Imported Modules' $strImportedModules }
if ($RequestedModules) { $OutputString += New-HtmlSummary 'Available Modules' $strAvailableModules }
if ($LoadedAssemblies) { $OutputString += New-HtmlSummary 'Loaded Assemblies' $strLoadedAssemblies }
if ($CommandHistory) { $OutputString += New-HtmlSummary 'Executed Commands' $strCommandHistory }
if ($ModuleErrors) { $OutputString += New-HtmlSummary 'Error Records' $strErrorRecords }
$OutputString = $OutputString -join "`r`n`r`n"

## Write Output
$OutputString
