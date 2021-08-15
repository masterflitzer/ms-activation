set-psdebug -trace 1

$_url = "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_14131-20278.exe"
$_home = "$home".replace('\', '/').tolower()
$_dir1 = "$_home/downloads"
$_dir2 = $_dir1 + "/ms-activation"
$_config = "office-config" + ".xml"

Invoke-WebRequest -Uri "$_url" -OutFile "$_dir1/odt.exe"
& "$_dir1/odt.exe" "/quiet /passive /extract:$_dir2/"
Remove-Item -Force -Path "$_dir2/*.xml"
& "$_dir2/setup.exe" "/download $_dir1/$_config"
& "$_dir2/setup.exe" "/configure $_dir1/$_config"
