param (
    [string] $config = $null,
    [string] $tool = $null
)

#Requires -Version 5.1
Set-StrictMode -Version 3
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$homeDir = [System.Environment]::GetFolderPath("UserProfile")

function EliminateMultipleSlash {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [string] $s
    )
    return [regex]::Replace($s, "//+", "/")
}

function NormalizePath {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [string] $s
    )
    return $s.Replace('~', $homeDir).Replace('\', '/') | EliminateMultipleSlash
}

function NormalizePathWin {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [string] $s
    )
    return ($s | NormalizePath).Replace('/', '\')
}

function GetFileDialog {
    Add-Type -AssemblyName "System.Windows.Forms"
    $fileDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $fileDialog.InitialDirectory = $homeDir
    $fileDialog.Multiselect = $false
    $fileDialog.ShowDialog() > $null
    $file = $fileDialog.FileName
    return $file
}

$oldDir = Get-Location | NormalizePath
$newDir = [System.IO.Path]::Combine(
    $homeDir,
    "Downloads",
    "office-deployment-tool"
) | NormalizePath
$setup = "$newDir/setup.exe"

if ($oldDir -eq $newDir) {
    Set-Location ..
    $oldDir = Get-Location | NormalizePath
}

if (!$config) {
    Write-Output "Please select the config file"
    $config = GetFileDialog | NormalizePath

    if (!$config) {
        throw "Invalid path to config file"
    }
}

if (!$tool) {
    Write-Output "Opening download page for the Office Deployment Tool (ODT)"
    Write-Output "Please download ODT and select the downloaded file in this prompt"
    Start-Process "https://microsoft.com/download/details.aspx?id=49117"
    $tool = GetFileDialog | NormalizePath
    
    if (!$tool) {
        throw "Invalid path to ODT file"
    }
}

Write-Output "`nIf you run into problems, try deleting this folder:`n"
Write-Output "`tRemove-Item -Recurse `"$newDir`" -Force`n"

New-Item -ItemType Directory $newDir -Force -ErrorAction SilentlyContinue > $null
Set-Location $newDir

Write-Output "Extracting Office Deployment Tool..."
Start-Process -Wait -FilePath $tool -ArgumentList "/extract:`"$($newDir | NormalizePathWin)`" /quiet /passive /norestart"
Remove-Item "$newDir/configuration-office*.xml" -Force

Write-Output "Downloading Office..."
& $setup /download $config
Write-Output "Installing Office..."
& $setup /configure $config

Set-Location $oldDir

Write-Output "Done!`n"
Write-Output "Press any key to continue..."
[System.Console]::ReadKey($true) > $null
