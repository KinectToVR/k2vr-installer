# request admin so installing shit doesnt suck ass
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
 }
}
echo ""
$Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkCyan')
$Host.UI.RawUI.ForegroundColor = ($bckgrnd = 'White')
echo "  __ __ _                 __ ______    _    ______     "
echo "  / //_/(_)___  ___  _____/ //_  __/___| |  / / __ \   "
echo "  / ,<  / / __ \/ _ \/ ___/ __// / / __ \ | / / /_/ /  "
echo "  / /| |/ / / / /  __/ /__/ /_ / / / /_/ / |/ / _, _/  "
echo "  /_/ |_/_/_/ /_/\___/\___/\__//_/  \____/|___/_/ |_|  "
echo "                                                       "
echo ""
$Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
Start-Sleep -s 0.8

# preparation stage
# find out which vr headset the user has
# TODO: Pimax, Quest-VirtualDesktop and Quest-ALVR detection
# TODO: Differentiate between Vive wands and Index controllers on Vive/Index/Pimax
$HMDIndex = "rift-cv1","rift-s","quest","index","vive","vive-pro","vive-cosmos","windows-mr","quest-alvr","quest-vd","pimax","other"
$HMDIndexReadable = "Oculus Rift CV1","Oculus Rift S","Oculus Quest","Valve Index","HTC Vive","HTC Vive Pro","HTC Vive Cosmos","Windows Mixed Reality","Oculus Quest (ALVR)","Oculus Quest (VirtualDesktop)","Pimax","Other/Unknown"
$HMDStatus = 0
$SteamDIR = (Get-Item HKCU:\Software\Valve\Steam).GetValue("SteamPath")
$SteamVRSettings = Get-Content -Path "$SteamDIR\config\steamvr.vrsettings" -Raw
$NewSteamVRSettings
$SteamVRSettingsJSON = $SteamVRSettings | ConvertFrom-Json

if($SteamVRSettingsJson.LastKnown.HMDModel = "Oculus Rift CV1")      {$HMDStatus = 0}
elseif($SteamVRSettingsJson.LastKnown.HMDModel = "Oculus Rift S")    {$HMDStatus = 1}
elseif($SteamVRSettingsJson.LastKnown.HMDModel = "Oculus Quest")     {$HMDStatus = 2}
elseif($SteamVRSettingsJson.LastKnown.HMDModel = "Valve Index")      {$HMDStatus = 3}
elseif($SteamVRSettingsJson.LastKnown.HMDModel = "HTC Vive")         {$HMDStatus = 4}
elseif($SteamVRSettingsJson.LastKnown.HMDModel = "HTC Vive Pro")     {$HMDStatus = 5}
elseif($SteamVRSettingsJson.LastKnown.HMDModel = "HTC Vive Cosmos")  {$HMDStatus = 6}
elseif($SteamVRSettingsJson.LastKnown.HMDManufacturer = "WindowsMR") {$HMDStatus = 7}
else{$HMDStatus = 11}
$HMDReadable = $HMDIndexReadable[$HMDStatus]
echo "Current VR Headset: $HMDReadable"

Start-Sleep -s 0.7

# figure out what kinect model is plugged in and if it has drivers
$KinectStatus = 0 # 0 = 360 1 = one
$KinectDriverStatus = 0 # 0 = no drivers 1 = drivers
if (Test-Path "C:\Program Files\Microsoft SDKs\Kinect"){$KinectDriverStatus = 1}

echo "Checking for Kinect model..."
if (Get-PnpDevice -ErrorAction 'Ignore' -PresentOnly -FriendlyName 'Kinect for Windows Device'){
    echo "Xbox 360 Kinect (V1) Found!"
    $KinectStatus = 0
}elseif (Get-PnpDevice -ErrorAction 'Ignore' -PresentOnly -FriendlyName 'Xbox NUI Motor'){
    echo "Xbox 360 Kinect (V1) Found!"
    $KinectStatus = 0
}elseif (Get-PnpDevice -ErrorAction 'Ignore' -PresentOnly -FriendlyName 'Kinect USB Audio'){
    echo "Xbox 360 Kinect (V1) Found!"
    $KinectStatus = 0
}elseif (Get-PnpDevice -ErrorAction 'Ignore' -PresentOnly -FriendlyName 'WDF KinectSensor Interface 0'){
    echo "Xbox One Kinect (V2) Found!"
    $KinectStatus = 1
}elseif (Get-PnpDevice -ErrorAction 'Ignore' -PresentOnly -FriendlyName 'Xbox NUI Sensor'){
    echo "Xbox One Kinect (V2) Found!"
    $KinectStatus = 1
}else{
    echo "No Kinect found! Please plug in a Kinect and start again!"
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("No Kinect found! Please plug in a Kinect sensor and amke sure it's also connected to power and start again!  The installer will now exit.   If you're still having issues, join discord.gg/Mu28W4N", 0, "KinectToVR Installer",48)
    exit
}
Start-Sleep -s 0.7

# make new folder
$CurrentPath = Get-Location
if (!(Test-Path .\temp)){
    New-Item -ItemType directory -Path .\temp
    echo "Created temporary folder at: $CurrentPath\temp\"
}else{
    echo "A folder already exists at: $CurrentPath\temp\... using it"
}
Start-Sleep -s 1
# downloading things... TODO: verify downloads and add alternate mirrors

# before we start, set PS to allow any type of TLS, older versions only allow 1.0 by default
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# download 7zip cli zip
if (!(Test-Path .\temp\7zip\)){
    echo "Downloading 7-Zip CLI"
    Invoke-WebRequest https://www.7-zip.org/a/7za920.zip -OutFile .\temp\7za920.zip
    echo "Extracting..."
    Expand-Archive -Path .\temp\7za920.Zip -DestinationPath .\temp\7zip\ -Force
}else{
    echo "7-Zip CLI is already present! Skipping download and extract"
}

# download kinecttovr
Start-Sleep -s 0.5
if (!(Test-Path .\temp\k2vr-0.6.0r2.7z)){
    echo "Downloading KinectToVR 0.6.0r2"
    Invoke-WebRequest 'https://github.com/sharkyh20/KinectToVR/releases/download/a0.6.0/KinectToVR-a0.6.0.Prime-time.Test.R2.7z' -OutFile .\temp\k2vr-0.6.0r2.7z
}else{
    echo "KinectToVR 0.6.0r2 is already present! Skipping download"
}

# download redists
Start-Sleep -s 0.5
if (!(Test-Path .\temp\vcredist-2010-x64.exe)){
    echo "Downloading Visual C++ Redistribuable C++ 2010 x64"
    Invoke-WebRequest https://download.microsoft.com/download/3/2/2/3224B87F-CFA0-4E70-BDA3-3DE650EFEBA5/vcredist_x64.exe -OutFile .\temp\vcredist-2010-x64.exe
}else{
    echo "Visual C++ Redistribuable C++ 2010 x64 is already present! Skipping download"
}
Start-Sleep -s 0.5
if (!(Test-Path .\temp\vcredist-2017-x64.exe)){
    echo "Downloading Visual C++ Redistribuable C++ 2017 x64"
    Invoke-WebRequest https://download.visualstudio.microsoft.com/download/pr/11100230/15ccb3f02745c7b206ad10373cbca89b/VC_redist.x64.exe -OutFile .\temp\vcredist-2017-x64.exe
}else{
    echo "Visual C++ Redistribuable C++ 2017 x64 is already present! Skipping download"
}
Start-Sleep -s 0.7
# download openvr input emulator installer
if (!(Test-Path .\temp\ovrie-1.3.exe)){
    echo "Downloading OpenVR-InputEmulator 1.3"
    Invoke-WebRequest https://github.com/matzman666/OpenVR-InputEmulator/releases/download/v1.3/OpenVR-InputEmulator-v1.3.exe -OutFile .\temp\ovrie-1.3.exe
}else{
    echo "OpenVR-InputEmulator 1.3 is already present! Skipping download"
}
Start-Sleep -s 0.6
# download openvr input emulator dll fix
if (!(Test-Path .\temp\driver_00vrinputemulator.dll)){
    echo "Downloading the OpenVR-InputEmulator SteamVR Driver Fix"
    Invoke-WebRequest https://github.com/sharkyh20/OpenVR-InputEmulator/releases/download/SteamVR-Fix/driver_00vrinputemulator.dll -OutFile .\temp\driver_00vrinputemulator.dll
}else{
    echo "The OpenVR-InputEmulator SteamVR Driver Fix is already present! Skipping download"
}
Start-Sleep -s 0.7
# downloading kinect sdk
if ($KinectDriverStatus = 0){
    if ($KinectStatus = 0){
        echo "Downloading Kinect SDK 1.8 for Xbox 360 Kinect"
        Invoke-WebRequest https://download.microsoft.com/download/E/1/D/E1DEC243-0389-4A23-87BF-F47DE869FC1A/KinectSDK-v1.8-Setup.exe -OutFile .\temp\kinectv1-sdk-1.8.exe
    }elseif ($KinectStatus = 1){
        echo "Downloading Kinect SDK 2.0 for Xbox One Kinect"
        Invoke-WebRequest https://download.microsoft.com/download/F/2/D/F2D1012E-3BC6-49C5-B8B3-5ACFF58AF7B8/KinectSDK-v2.0_1409-Setup.exe -OutFile .\temp\kinectv2-sdk-2.0.exe
    }
}else{
    echo "Kinect drivers are already installed, skipping download"
}
Start-Sleep -s 2
# INSTALLING FOR REALS
if (!(Test-Path C:\KinectToVR\)){
    echo "Extracting K2VR to the C drive (C:\KinectToVR)"
    .\temp\7zip\7za.exe x .\temp\k2vr-0.6.0r2.7z -aoa -oC:\ | Out-Null
    Rename-Item -Path "c:\KinectToVR-a0.6.0 Prime-time Test R2" -NewName "KinectToVR"
}else{
    echo "KinectToVR is already present! Skipping extract"
}
Start-Sleep -s 3
echo "Installing Visual C++ Redistribuable 2010 x64"
Start-Process .\temp\vcredist-2010-x64.exe /q -NoNewWindow -Wait
Start-Sleep -s 1
echo "Downloading Visual C++ Redistribuable 2017 x64"
Start-Process .\temp\vcredist-2017-x64.exe /q -NoNewWindow -Wait
Start-Sleep -s 1
if (!(Test-Path "C:\Program Files\OpenVR-InputEmulator\")){
    echo "Installing OpenVR-InputEmulator 1.3"
    Start-Process .\temp\ovrie-1.3.exe /S -NoNewWindow -Wait
}else{
    echo "OVRIE is already installed, skipping install to avoid upgrade prompt"
}
Start-Sleep -s 0.6
echo "Copying the SteamVR DLL Fix to the right folder"
Copy-Item -Force .\temp\driver_00vrinputemulator.dll -Destination "$SteamDIR\steamapps\common\SteamVR\drivers\00vrinputemulator\bin\win64"
Start-Sleep -s 0.8
if (($KinectStatus = 0) -and ($KinectDriverStatus = 0)){
    echo "Running Kinect SDK 1.8 Installer for Xbox 360 Kinect"
    echo "Installation will continue when the SDK installer is done"
    Start-Process .\temp\kinectv1-sdk-1.8.exe -NoNewWindow -Wait
}elseif (($KinectStatus = 1) -and ($KinectDriverStatus = 0)){
    echo "Running Kinect SDK 2.0 Installer for Xbox One Kinect"
    echo "Installation will continue when the SDK installer is done"
    Start-Process .\temp\kinectv2-sdk-2.0.exe -NoNewWindow -Wait
}else{
    echo "Skipping Kinect SDK install since the drivers are already present"
}
Start-Sleep -s 2
# extra cleanup...
echo "Creating Start Menu entry for KinectToVR..."
echo "Downloading icon into the K2VR folder"
Invoke-WebRequest https://github.com/TripingPC/k2vr-installer/raw/master/k2vr.ico -OutFile C:\KinectToVR\k2vr.ico
New-Item -ItemType directory -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\KinectToVR"
Invoke-WebRequest https://github.com/TripingPC/k2vr-installer/raw/master/KinectToVR-V1.lnk -OutFile "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\KinectToVR\KinectToVR (Xbox 360).lnk"
echo "Created Start Menu entry for kinectv1process.exe!"
Invoke-WebRequest https://github.com/TripingPC/k2vr-installer/raw/master/KinectToVR-V2.lnk -OutFile "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\KinectToVR\KinectToVR (Xbox One).lnk"
echo "Created Start Menu entry for kinectv2process.exe!"
Start-Sleep -s 1
# force enable openvr input emu driver by abusing the steamvr settings file

# btw thanks valve for whatever JSON interpreter you use
# it literally cleans up and consolidates duplicate entries :D
# because PS' wrapper around Newtonsoft doesn't allow to change the tab length
# making this string replace hack kind of the only slightly sane method
echo "Force-enabling OpenVR-InputEmulator"
Start-Sleep -s 0.4
$NewSteamVRSettings = $SteamVRSettings.Replace(
"`"steamvr`" : {",
"`"driver_00vrinputemulator`" : {`n      `"blocked_by_safe_mode`" : false`n   },`n   `"steamvr`" : {")
echo "Safe mode blocked successfully!"
Set-Content -Path "$SteamDIR\config\steamvr.vrsettings" -Value $NewSteamVRSettings
Start-Sleep -s 0.7
echo "Saved to SteamVR settings"

# script end
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Installation completed! If you need more help, read the instructions on the website or join discord.gg/Mu28W4N", 0, "KinectToVR Installer",48)
