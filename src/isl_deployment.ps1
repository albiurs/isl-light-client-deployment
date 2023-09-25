########################################################
#
# ISL Light Client deployment script
#
# Script repository:
# https://github.com/albiurs/isl-light-client-deployment
#
########################################################


# Create target directory and download ISL client from the Digitag Website
mkdir "C:\PC-Service"
curl -o "C:\PC-Service\ISL Light Client.exe" "https://isl.digitag.ch/start/ISLLightClient"

# Create shortcut
$shell = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut("C:\PC-Service\Digitag_Fernwartung.lnk")
$shortcut.TargetPath = "C:\PC-Service\ISL Light Client.exe"
$shortcut.Save()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($shortcut) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null

# Copy shortcut to the current desktop
Copy-Item "C:\PC-Service\Digitag_Fernwartung.lnk" "$env:USERPROFILE\Desktop"
