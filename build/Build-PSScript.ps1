param
(
    # Specifies the path to the script.
    [Parameter(Mandatory = $true)]
    [string] $ScriptPath,
    #
    [Parameter(Mandatory = $false)]
    [string] $OutputDirectory = ".\build\release\"
)

## Read Script Info
#$ScriptInfo = Test-ScriptFileInfo $ScriptPath
#[System.IO.DirectoryInfo] $ScriptOutputDirectoryInfo = Join-Path $OutputDirectory (Join-Path $ScriptInfo.Name $ScriptInfo.Version)

## Copy Source Script to Output Directory
Assert-DirectoryExists $OutputDirectory -ErrorAction Stop | Out-Null
$ScriptPathInfo = Copy-Item $ScriptPath -Destination $OutputDirectory -Force -PassThru
$ScriptPathInfo

## Sign Script
&$PSScriptRoot\Sign-PSScript.ps1 -ScriptPath $ScriptPathInfo | Format-Table Path, Status, StatusMessage
