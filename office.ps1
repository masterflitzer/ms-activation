#Requires -Version 5.1
Set-StrictMode -Version 3
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$repoBaseUrl = "https://github.com/masterflitzer/ms-activation.git"
$repoBaseUrl = $repoBaseUrl.Replace(".git", "")

function NormalizePath {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [string] $s
    )
    return [regex]::Replace($s.Replace('~', $HOME).Replace('/', '\'), "\\+", "\")
}

function GetFileDialog {
    Add-Type -AssemblyName "System.Windows.Forms"
    $fileDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $fileDialog.InitialDirectory = $HOME
    $fileDialog.Multiselect = $false
    $fileDialog.ShowDialog() > $null
    $file = $fileDialog.FileName
    return $file
}

Write-Output ""

$setup = "setup.exe"
$config = "office.xml"

$oldDir = Get-Location | NormalizePath
$newDir = [System.IO.Path]::Combine(
    "$HOME",
    "Downloads",
    "office-deployment-tool"
) | NormalizePath

Write-Output "if you run into problems, try deleting this folder:"
Write-Output "`tRemove-Item -Recurse `"$newDir`" -Force`n"

New-Item -ItemType Directory "$newDir" -Force -ErrorAction SilentlyContinue > $null
Set-Location "$newDir"

Write-Output "opening download page for the office deployment tool (odt)"
Write-Output "download the odt exe and select it in the following prompt"
Write-Output "if the prompt opened in the background try pressing alt + tab`n"

Start-Process -Wait "https://microsoft.com/download/details.aspx?id=49117"
$odt = GetFileDialog | NormalizePath

if (!$odt -or !$(Test-Path "$odt")) {
    $err = "invalid path to the odt exe: $odt"
    throw $err
}

Write-Output "extracting office deployment tool..."
Start-Process -Wait "$odt" -ArgumentList @(
    "/extract:`"$newDir`"",
    "/norestart",
    "/passive",
    "/quiet"
)

Write-Output "downloading office config file..."
curl.exe -Lso "$config" "$repoBaseUrl/raw/main/$config"

if (!$config -or !$(Test-Path "$config")) {
    $err = "invalid path to the office config file: $config"
    throw $err
}

Write-Output "downloading office..."
Start-Process -Verb RunAs -Wait "$setup" -ArgumentList @(
    "/download",
    "$config"
)

Write-Output "installing office..."
Start-Process -Verb RunAs -Wait "$setup" -ArgumentList @(
    "/configure",
    "$config"
)

Set-Location "$oldDir"

Write-Output "`ndone!`n"

Write-Output "press any key to continue..."
[System.Console]::ReadKey($true) > $null
