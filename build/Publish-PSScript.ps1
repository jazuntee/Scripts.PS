#Requires -Version 7.0
param
(
    # Path to Script
    [Parameter(Mandatory = $true)]
    [string] $ScriptPath,
    # Repository for PowerShell Gallery
    [Parameter(Mandatory = $false)]
    [string] $RepositorySourceLocation = 'https://www.powershellgallery.com/api/v2',
    # API Key for PowerShell Gallery
    [Parameter(Mandatory = $true)]
    [securestring] $NuGetApiKey,
    # Unlist from PowerShell Gallery
    [Parameter(Mandatory = $false)]
    [switch] $Unlist
)

## Publish
$PSRepositoryAll = Get-PSRepository
$PSRepository = $PSRepositoryAll | Where-Object SourceLocation -like "$RepositorySourceLocation*"
if (!$PSRepository) {
    try {
        [string] $RepositoryName = New-Guid
        Register-PSRepository $RepositoryName -SourceLocation $RepositorySourceLocation
        $PSRepository = Get-PSRepository $RepositoryName
        Publish-Script -Path $ScriptPath -NuGetApiKey (ConvertFrom-SecureString $NuGetApiKey -AsPlainText) -Repository $PSRepository.Name
    }
    finally {
        Unregister-PSRepository $RepositoryName
    }
}
else {
    Write-Verbose ('Publishing Script [{0}]' -f $ScriptPath)
    Publish-Script -Path $ScriptPath -NuGetApiKey (ConvertFrom-SecureString $NuGetApiKey -AsPlainText) -Repository $PSRepository.Name
}

## Unlist the Package
if ($Unlist) {
    $ScriptInfo = Test-ScriptFileInfo $ScriptPath
    Invoke-RestMethod -Method Delete -Uri ("{0}/{1}/{2}" -f $PSRepository.PublishLocation, $ScriptInfo.Name, $ScriptInfo.Version) -Headers @{ 'X-NuGet-ApiKey' = ConvertFrom-SecureString $NuGetApiKey -AsPlainText }
}
