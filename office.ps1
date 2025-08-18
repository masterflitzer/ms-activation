#Requires -Version 5.1
Set-StrictMode -Version 3
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

filter NormalizePath {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string] $s
    )
    [regex]::Replace($s.Replace('~', "${HOME}").Replace('/', '\'), "\\+", "\")
}

filter GetFileDialog {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string] $dir
    )
    Add-Type -AssemblyName "System.Windows.Forms"
    $fileDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $fileDialog.InitialDirectory = $dir | NormalizePath | Resolve-Path
    $fileDialog.Multiselect = $false
    $fileDialog.ShowDialog() > $null
    $fileDialog.FileName | NormalizePath
}

$setup = "setup.exe"
$config = "config.xml"

$repoUrl = "https://github.com/masterflitzer/ms-activation.git".Replace(".git", "")
$configUrl = "${repoUrl}/raw/main/office.xml"

$downloadsDir = [System.IO.Path]::Combine(
    "${HOME}",
    "Downloads"
) | NormalizePath

$oldDir = Get-Location | NormalizePath
$newDir = [System.IO.Path]::Combine(
    "${downloadsDir}",
    "office-deployment-tool"
) | NormalizePath

Write-Output "if you run into problems, try deleting this folder:"
Write-Output "`tRemove-Item -Recurse `"${newDir}`" -Force"

New-Item -ItemType Directory "${newDir}" -Force -ErrorAction SilentlyContinue > $null
Set-Location "${newDir}"

Write-Output "`nthis script requires the office deployment tool (odt)"
Write-Output "opening the download page for the odt and a file selection dialog now..."
Write-Output "please download the odt *.exe file and select it in the file dialog"
Write-Output "`nnote: the dialog may open in the background,"
Write-Output "use the taskbar or alt+tab to bring it to the foreground`n"

Start-Process -Wait -WindowStyle Hidden -FilePath "https://microsoft.com/download/details.aspx?id=49117"
$odt = $downloadsDir | GetFileDialog

if (!$odt -or !$(Test-Path "${odt}")) {
    $err = "invalid path to the odt exe: ${odt}"
    throw $err
}

Write-Output "extracting office deployment tool..."
Start-Process -Verb RunAs -Wait -WindowStyle Hidden -FilePath "${odt}" -ArgumentList @(
    "/extract:`"${newDir}`"",
    "/norestart",
    "/passive",
    "/quiet"
)

Write-Output "downloading office config file..."
curl.exe -Lso "${config}" "${configUrl}"

if (!$config -or !$(Test-Path "${config}")) {
    $err = "invalid path to the office config file: ${config}"
    throw $err
}

Write-Output "downloading office..."
Start-Process -Verb RunAs -Wait -WindowStyle Hidden -FilePath "${setup}" -ArgumentList @(
    "/download",
    "${config}"
)

Write-Output "installing office..."
Start-Process -Verb RunAs -Wait -WindowStyle Hidden -FilePath "${setup}" -ArgumentList @(
    "/configure",
    "${config}"
)

Set-Location "${oldDir}"

Write-Output "`ndone!`n"

Write-Output "press any key to continue..."
[System.Console]::ReadKey($true) > $null
