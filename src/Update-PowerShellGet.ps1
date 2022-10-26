
<#PSScriptInfo

.VERSION 1.0

.GUID b1c6267a-96ba-408d-84bd-5a106a9dd745

.AUTHOR Jason Thompson

.COMPANYNAME Microsoft Corporation

.COPYRIGHT (c) 2021 Jason Thompson. All rights reserved.

.TAGS Microsoft Windows PowerShell PSEdition_Desktop Windows

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
    Update PowerShellGet Module
.DESCRIPTION 
    Update PowerShellGet Module on Windows PowerShell
.EXAMPLE
    PS > Update-PowerShellGet.ps1

    Update PowerShellGet
.EXAMPLE
    PS > iex $(irm 'https://aka.ms/Update-PowerShellGet')

    Invoke this script directly from GitHub.
#> 


#region Supporting Functions
<#
.SYNOPSIS
    Test if current PowerShell process is elevated to local administrator privileges.
.DESCRIPTION
    Test if current PowerShell process is elevated to local administrator privileges.
.EXAMPLE
    PS C:\>Test-PsProcessElevated
    Test is current PowerShell process is elevated.
.LINK
    https://github.com/jasoth/Utility.PS
#>
function Test-PsProcessElevated {
    [CmdletBinding()]
    [OutputType([bool])]
    param()

    $WindowsIdentity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $WindowsPrincipal = New-Object 'System.Security.Principal.WindowsPrincipal' $WindowsIdentity
    $LocalAdministrator = [System.Security.Principal.WindowsBuiltInRole]::Administrator
    return $WindowsPrincipal.IsInRole($LocalAdministrator)
}
#endregion


if ($PSVersionTable.PSEdition -eq 'Desktop') {
    ## Force TLS 1.2 on old versions
    if ($PSVersionTable.PSVersion -lt [version]'5.1.17763.0') {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }

    ## Allow Scripts in current process
    Set-ExecutionPolicy RemoteSigned -Scope Process

    ## Update Nuget Package and PowerShellGet Module
    if (Test-PsProcessElevated) {
        Install-PackageProvider NuGet -Scope AllUsers -Force
        Install-Module PowerShellGet -Scope AllUsers -Force -AllowClobber
    }
    else {
        Install-PackageProvider NuGet -Scope CurrentUser -Force
        Install-Module PowerShellGet -Scope CurrentUser -Force -AllowClobber
    }

    ## Remove old modules from existing session
    Remove-Module PowerShellGet,PackageManagement -Force -ErrorAction Ignore

    ## Import updated module
    Import-Module PowerShellGet -MinimumVersion 2.0 -Force
    Import-PackageProvider PowerShellGet -MinimumVersion 2.0 -Force
}
else {
    Write-Warning 'This command is intended to update PowerShellGet on Windows PowerShell (<=v5.1) and does not do anything on PowerShell v6.0 and later.'
}
