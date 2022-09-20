[CmdletBinding()]
param (
    [string] $Config = $null
)

Set-StrictMode -Version 3
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$myHome = [Environment]::GetFolderPath("UserProfile")
$url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_14326-20404.exe"

function IsNullOrWhiteSpace ([string]$s) {
    return [string]::IsNullOrWhiteSpace($s)
}

function EliminateMultipleSlash ([string]$s) {
    return [Regex]::Replace($s, "//+", "/")
}

function NormalizePath ([string]$s) {
    return EliminateMultipleSlash $s.Replace('~', $myHome).Replace('\', '/')
}

function NormalizePathWin ([string]$s) {
    return $(NormalizePath $s).Replace('/', '\')
}

function ThrowInvalidPath ([string]$s) {
    if ($null -eq $s) {
        $x = "!"
    } else {
        $x = ": $s"
    }
    throw "Invalid path to config file$x"
}

function GetFilePath {
    param (
        [Parameter(Mandatory = $true)]
        [string]$InitialDir
    )
    Add-Type -AssemblyName "System.Windows.Forms"
    $fileDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $fileDialog.InitialDirectory = $InitialDir
    $fileDialog.Multiselect = $false
    $fileDialog.ShowDialog() > $null
    $filePath = $fileDialog.FileName
    if ($null -ne $filePath -or "" -ne $filePath -or $(Test-Path -PathType Leaf $filePath)) {
        return $filePath
    } else {
        return $null
    }
}

$dir = NormalizePath $([System.IO.Path]::Combine(
        $myHome,
        "Downloads",
        "office-deployment-tool"
    ))
$file = "$dir/odt.exe"
$setup = "$dir/setup.exe"

if (IsNullOrWhiteSpace $Config) {
    $path = GetFilePath -InitialDir $myHome
    if ($null -eq $path) {
        ThrowInvalidPath
    }
    $Config = $path
}

$Config = NormalizePath $Config

if (!(Test-Path -PathType Leaf $Config)) {
    ThrowInvalidPath $Config
}

Write-Output "`nIf you run into problems, try deleting this folder:`n"
Write-Output "`tRemove-Item -Recurse `"$dir`" -Force`n"

if (!(Test-Path -PathType Container $dir)) {
    New-Item -ItemType Directory $dir -Force > $null
}

Set-Location $dir

if (!(Test-Path -PathType Leaf $file)) {
    Write-Output "Downloading Office Deployment Tool..."
    Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $file
}

Write-Output "Extracting Office Deployment Tool..."
Start-Process -Wait -FilePath $file -ArgumentList "/extract:`"$(NormalizePathWin $dir)`" /quiet /passive /norestart"
Remove-Item "$dir/configuration-office*.xml" -Force

Write-Output "Downloading Office..."
& $setup /download $Config
Write-Output "Installing Office..."
& $setup /configure $Config

Write-Output "Done!`n"
Write-Output "Press any key to continue..."
[System.Console]::ReadKey($true) > $null
