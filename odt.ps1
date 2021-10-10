param (
    [Parameter(Mandatory = $false, Position = 0)]
    [string] $config = $null
)

$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$_config = $config
$_url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_14326-20404.exe"
$_home = [Environment]::GetFolderPath("UserProfile")

function IsNullOrWhiteSpace ([string]$s) { return [string]::IsNullOrWhiteSpace($s) }

function EliminateMultipleSlash ([string]$s) { return [Regex]::Replace($s, "//+", "/") }

function NormalizePath ([string]$s) { return EliminateMultipleSlash($s.Replace('~', $_home).Replace('\', '/')) }

function NormalizePathWin ([string]$s) { return NormalizePath($s).Replace('/', '\') }

function GetFilePath {
    param (
        [Parameter(Mandatory = $true)]
        [string]$initialDir
    )
    Add-Type -AssemblyName "system.windows.forms"
    $_file = New-Object System.Windows.Forms.OpenFileDialog
    $_file.InitialDirectory = $initialDir
    $_file.Multiselect = $false
    $_file.ShowDialog() | Out-Null
    return $_file.FileName
}

$_downloads = NormalizePath([System.IO.Path]::Combine($_home, "Downloads"))
$_dir = NormalizePath([System.IO.Path]::Combine($_downloads, "office-deployment-tool"))

if (IsNullOrWhiteSpace($_config)) { $_config = GetFilePath($_home) }
if (IsNullOrWhiteSpace($_config) -or !IsPathFullyQualified($_config)) { exit 1 }
$_config = NormalizePath($_config)

Write-Host "`nIf you run into problems, try deleting this folder:"
Write-Host "$_dir`n"

if (!(Test-Path "$_dir")) { New-Item -Force -ItemType Directory "$_dir" | Out-Null }
Set-Location "$_dir"

if (!(Test-Path "$_dir/odt.exe")) {
    Write-Host "Downloading Office Deployment Tool..."
    Invoke-WebRequest -Uri "$_url" -OutFile "$_dir/odt.exe"
}

Write-Host "Extracting Office Deployment Tool..."
Start-Process -Wait -FilePath "$_dir/odt.exe" -ArgumentList "/extract:`"$(NormalizePathWin("$_dir"))`" /quiet /passive /norestart"
Remove-Item -Force "$_dir/configuration-office*.xml"

Write-Host "Downloading Office..."
& "$_dir/setup.exe" /download "$_config"
Write-Host "Installing Office..."
& "$_dir/setup.exe" /configure "$_config"

Write-Host "Done!`n"
Write-Host -NoNewline "Press any key to continue..."
[System.Console]::ReadKey($true) | Out-Null
