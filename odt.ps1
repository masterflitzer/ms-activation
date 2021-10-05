param (
    [Parameter(Mandatory = $false, Position = 0)]
    [string] $config = $null
)

$ErrorActionPreference = "Stop"

function IsNullOrWhiteSpace ([string]$s) {
    return [string]::IsNullOrWhiteSpace($s)
}

function IsPathFullyQualified ([string]$path) {
    return [System.IO.Path]::IsPathFullyQualified($path)
}

function NormalizePath ([string]$path) {
    return $path.Replace('\', '/').ToLower()
}

function NormalizePathWin ([string]$path) {
    return $path.Replace('/', '\').ToLower()
}

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

$_config = $config
$_dir = NormalizePath([System.IO.Path]::Combine([Environment]::GetFolderPath('Personal'), "office-deployment-tool"))
$_home = NormalizePath([Environment]::GetFolderPath('UserProfile'))
$_url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_14326-20404.exe"

Write-Host "`nIf you run into problems, try deleting this folder:`n$_dir`n"

if (IsNullOrWhiteSpace($_config)) { $_config = GetFilePath($_home) }
if (IsNullOrWhiteSpace($_config) -or !IsPathFullyQualified($_config)) { exit 1 }

$_config = NormalizePath($_config)

#if (Test-Path $_dir) { Remove-Item -Recurse "$_dir" }
New-Item -Force -ItemType Directory "$_dir" | Out-Null
Invoke-WebRequest -Uri "$_url" -OutFile "$_dir/odt.exe"
Start-Process -Wait -FilePath "$_dir/odt.exe" -ArgumentList "/extract:`"$(NormalizePathWin("$_dir"))`" /quiet /passive /norestart"
Remove-Item -Force "$_dir/configuration-office*.xml"
& "$_dir/setup.exe" /download "$_config"
& "$_dir/setup.exe" /configure "$_config"

Write-Host -NoNewline "Press any key to continue..."
[System.Console]::ReadKey($true) | Out-Null
