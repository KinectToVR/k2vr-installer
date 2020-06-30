# request admin so installing shit doesnt suck ass
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
 if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
  $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
  Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
  Exit
 }
}

$ProgressPreference = 'SilentlyContinue'

$version = "1.4.2"
$host.ui.RawUI.WindowTitle = "KinectToVR installer (Version $version)"

echo ""
$Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkGreen')
$Host.UI.RawUI.ForegroundColor = ($bckgrnd = 'White')
echo "   _  ___                 _  _____  __     ______    "
echo "  | |/ (_)_ __   ___  ___| ||_   _|_\ \   / /  _ \   "
echo "  | ' /| | '_ \ / _ \/ __| __|| |/ _ \ \ / /| |_) |  "
echo "  | . \| | | | |  __/ (__| |_ | | (_) \ V / |  _ <   "
echo "  |_|\_\_|_| |_|\___|\___|\__||_|\___/ \_/  |_| \_\  "
echo " Full-body tracking for Xbox 360 and Xbox One Kinect "
echo ""
$Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
Start-Sleep -s 0.2

# https://stackoverflow.com/a/28482050/
$steamVrMonitor = Get-Process vrmonitor -ErrorAction SilentlyContinue
if ($steamVrMonitor) {
  echo "SteamVR must be closed during the install process. Closing it now..."
  $steamVrMonitor.CloseMainWindow() | Out-Null
  Sleep 5
  if (!$steamVrMonitor.HasExited) {
    # When SteamVR is open with no headset detected,
    # CloseMainWindow will only close the "headset not found" popup
    # so we use Stop-Process to close it, if it's still open
    $steamVrMonitor | Stop-Process
    Sleep 3 # Give it time to actually close itself and vrserver
  }
}
Remove-Variable steamVrMonitor
# Apparently, SteamVR server can run without the monitor,
# so we close that, if it's open aswell (monitor will complain if you close server first)
$steamVrServer = Get-Process vrserver -ErrorAction SilentlyContinue
if ($steamVrServer) {
  # CloseMainWindow won't work here because it doesn't have a window
  $steamVrServer | Stop-Process
}
Remove-Variable steamVrServer

# BIG TODO
# ADD IMPROVED KINECT DETECTION AND PSMOVE DETECTION

# preparation stage
# find out which vr headset the user has
# TODO: Pimax, Quest-VirtualDesktop and Quest-ALVR detection
# TODO: Differentiate between Vive wands and Index controllers on Vive/Index/Pimax
$arg=$args[0]

md5 = ["2FAC454A90AE96021F4FFC607D4C00F8","A4BEFBE4A2D5BCA5C52F4542A15BEC89","C76CDA20DD2D939101E39B95DD30CF1F","479010273D25BEAB4868D57FFF907D50","09A95EA4BDB9A619828E6D13AC486962","2E742FDF0B00293CF55EEB9C5860EAC8","630D75210B325A280C3352F879297ED5","77C0F604585FB429C722BE111CA30C37"]
# in order: 7zip tools, ovrie dll, kinecttovr, sdk 1.8, sdk 2.0, ovrie installer, vcredist 2010, vcredist 2017

$HMDIndex = "rift-cv1","rift-s","quest","index","vive","vive-pro","vive-cosmos","windows-mr","quest-alvr","quest-vd","pimax","other"
$HMDIndexReadable = "Oculus Rift CV1","Oculus Rift S","Oculus Quest","Valve Index","HTC Vive","HTC Vive Pro","HTC Vive Cosmos","Windows Mixed Reality","Oculus Quest (ALVR)","Oculus Quest (VirtualDesktop)","Pimax","Other/Unknown"
$HMDStatus = 0
$openvrpath = "%appdata%/../LocaL/openvr/openvrpaths.vrpath"
$SteamDIR = (Get-Item HKLM:/SOFTWARE/WOW6432Node/Valve/Steam).GetValue("InstallPath")
$SteamVRSettings = Get-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Raw
$NewSteamVRSettings
$SteamVRSettingsJSON = $SteamVRSettings | ConvertFrom-Json

if($SteamVRSettingsJson.LastKnown.HMDModel -eq "Oculus Rift CV1")      {$HMDStatus = 0}
elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Oculus Rift S")    {$HMDStatus = 1}
elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Oculus Quest")     {$HMDStatus = 2}
elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Index")      {$HMDStatus = 3}
elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Vive. MV")         {$HMDStatus = 4}
elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Vive MV.")         {$HMDStatus = 4}
elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Vive MV")         {$HMDStatus = 4}
elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "VIVE_Pro MV")     {$HMDStatus = 5}
elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "HTC Vive Cosmos")  {$HMDStatus = 6}
elseif($SteamVRSettingsJson.LastKnown.HMDManufacturer -eq "WindowsMR") {$HMDStatus = 7}
else{$HMDStatus = 11}
$HMDReadable = $HMDIndexReadable[$HMDStatus]
echo "Current VR Headset: $HMDReadable"

Start-Sleep -s 0.2

# TODO
# usb controller checks
# prompt for OVRAS install
# FETCH UPDATED BINDINGS
# extract SDK MSIs for silent install https://github.com/Deledrius/KinectCam/blob/master/.github/workflows/build.yml#L35

# figure out what kinect model is plugged in and if it has drivers
$KinectStatus = 0 # 0 = 360 1 = one
$KinectDriverInstall = 0


# if (Get-PnpDevice -ErrorAction 'Ignore' -PresentOnly -FriendlyName 'Kinect for Windows Device'){
#     echo "Xbox 360 Kinect (V1) Found!"
#     $KinectStatus = 0
# }elseif (Get-PnpDevice -ErrorAction 'Ignore' -PresentOnly -FriendlyName 'Xbox NUI Motor'){
#     echo "Xbox 360 Kinect (V1) Found!"
#     $KinectStatus = 0
# }elseif (Get-PnpDevice -ErrorAction 'Ignore' -PresentOnly -FriendlyName 'Kinect USB Audio'){
#     echo "Xbox 360 Kinect (V1) Found!"
#     $KinectStatus = 0
# }elseif (Get-PnpDevice -ErrorAction 'Ignore' -PresentOnly -FriendlyName 'WDF KinectSensor Interface 0'){
#     echo "Xbox One Kinect (V2) Found!"
#     $KinectStatus = 1
# }elseif (Get-PnpDevice -ErrorAction 'Ignore' -PresentOnly -FriendlyName 'Xbox NUI Sensor'){
#     echo "Xbox One Kinect (V2) Found!"
#     $KinectStatus = 1
# }else{
#     echo "No device found! Please connect a Kinect sensor and start again!"
#     $wshell = New-Object -ComObject Wscript.Shell
#     $wshell.Popup("No device found! Please connect a Kinect sensor, verify it's connected to power and try again!  The installer will now exit.   If you're still having issues, join discord.gg/Mu28W4N", 0, "KinectToVR Installer",48)
#     exit
# }

# oh god new code
if (!($arg)) {
    echo "Checking for Kinect model..."
    $kinectv1_presence_main = (([regex]::Matches((gwmi Win32_USBControllerDevice |%{[wmi]($_.Dependent)} | Select -Property PNPDeviceID | Out-String), "02B0" )) | Select -Property Success)
    $kinectv1_presence_audio = (([regex]::Matches((gwmi Win32_USBControllerDevice |%{[wmi]($_.Dependent)} | Select -Property PNPDeviceID | Out-String), "02BB" )) | Select -Property Success)
    $kinectv1_presence_camera = (([regex]::Matches((gwmi Win32_USBControllerDevice |%{[wmi]($_.Dependent)} | Select -Property PNPDeviceID | Out-String), "02AE" )) | Select -Property Success)
    $kinectv2_presence = (([regex]::Matches((gwmi Win32_USBControllerDevice |%{[wmi]($_.Dependent)} | Select -Property PNPDeviceID | Out-String), "02C4" )) | Select -Property Success)
    if ($kinectv1_presence_main -or $kinectv1_presence_audio -or $kinectv1_presence_camera){
        echo "Xbox 360 Kinect (V1) Found!"
        $KinectStatus = 0
    }elseif ($kinectv2_presence) {
        echo "Xbox One Kinect (V2) Found!"
        $KinectStatus = 1
    }else {
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
        echo "No device found! Please connect a Kinect sensor and start again!"
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("No device found! Please connect a Kinect sensor, verify it's connected to power and try again!  The installer will now exit.   If you're still having issues, join discord.gg/YBQCRDG", 0, "KinectToVR Installer",48)
        $wshell.Popup("If you want to help with better detection, send us your USB device Instance IDs and which Kinect model you are using on the KinectToVR Discord`n`nAlso, you can force the installer by adding 'v1' or 'v2' as a launch argument for Xbox 360 and Xbox One Kinect respectively.", 0, "KinectToVR Installer",48)
        exit
    }
}elseif ($arg -eq 'v1') {
    $KinectStatus = 0
    echo "Xbox 360 Kinect (V1) Override Enabled"
}elseif ($arg -eq 'v2') {
    $KinectStatus = 1
    echo "Xbox One Kinect (V2) Override Enabled"
}elseif ($arg -eq 'psmove') {
    $KinectStatus = 2
    echo "PlayStation Move Override Enabled"
}
Start-Sleep -s 0.2

# make new folder
$CurrentPath = Get-Location
if (!(Test-Path ./temp)){
    New-Item -ItemType directory -Path ./temp
    echo "Created temporary folder at: $CurrentPath/temp/"
}else{
    echo "A folder already exists at: $CurrentPath/temp/... using it"
}
Start-Sleep -s 0.2
# downloading things... TODO: verify downloads and add alternate mirrors

# before we start, set PS to allow any type of TLS, older versions only allow 1.0 by default
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# download 7zip cli zip
if (!(Test-Path ./temp/7zip/)){
    echo "Downloading 7-Zip tools"
    Invoke-WebRequest https://www.7-zip.org/a/7za920.zip -OutFile ./temp/7za920.zip
    if (!(Test-Path ./temp/7za920.zip)){
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
        echo "Could not download! Check your firewall? If this error persists, join discord.gg/YBQCRDG for help."
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
        Pause
        Exit
    }
    echo "Extracting..."
    Expand-Archive -Path ./temp/7za920.Zip -DestinationPath ./temp/7zip/ -Force
}else{
    echo "7-Zip tools already present! Skipping download"
}

# download kinecttovr
Start-Sleep -s 0.2
if (!(Test-Path ./temp/k2vr-0.6.0r2.7z)){
    echo "Downloading KinectToVR 0.6.0r2"
    Invoke-WebRequest 'https://github.com/sharkyh20/KinectToVR/releases/download/a0.6.0/KinectToVR-a0.6.0.Prime-time.Test.R2.7z' -OutFile ./temp/k2vr-0.6.0r2.7z
    if (!(Test-Path ./temp/k2vr-0.6.0r2.7z)){
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
        echo "Could not download! Check your firewall? If this error persists, join discord.gg/YBQCRDG for help."
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
        Pause
        Exit
    }
}else{
    echo "KinectToVR 0.6.0r2 is already present! Skipping download"
}

# download redists
Start-Sleep -s 0.2
if (!(Test-Path ./temp/vcredist-2010-x64.exe)){
    echo "Downloading Visual C++ Redistribuable C++ 2010 x64"
    Invoke-WebRequest https://download.microsoft.com/download/3/2/2/3224B87F-CFA0-4E70-BDA3-3DE650EFEBA5/vcredist_x64.exe -OutFile ./temp/vcredist-2010-x64.exe
    if (!(Test-Path ./temp/vcredist-2010-x64.exe)){
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
        echo "Could not download! Check your firewall? If this error persists, join discord.gg/YBQCRDG for help."
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
        Pause
        Exit
    }
}else{
    echo "Visual C++ Redistribuable C++ 2010 x64 is already present! Skipping download"
}
Start-Sleep -s 0.2
if (!(Test-Path ./temp/vcredist-2017-x64.exe)){
    echo "Downloading Visual C++ Redistribuable C++ 2017 x64"
    Invoke-WebRequest https://download.visualstudio.microsoft.com/download/pr/11100230/15ccb3f02745c7b206ad10373cbca89b/VC_redist.x64.exe -OutFile ./temp/vcredist-2017-x64.exe
    if (!(Test-Path ./temp/vcredist-2017-x64.exe)){
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
        echo "Could not download! Check your firewall? If this error persists, join discord.gg/YBQCRDG for help."
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
        Pause
        Exit
    }
}else{
    echo "Visual C++ Redistribuable C++ 2017 x64 is already present! Skipping download"
}
Start-Sleep -s 0.2
# download openvr input emulator installer
if (!(Test-Path ./temp/ovrie-1.3.exe)){
    echo "Downloading OpenVR-InputEmulator 1.3"
    Invoke-WebRequest https://github.com/matzman666/OpenVR-InputEmulator/releases/download/v1.3/OpenVR-InputEmulator-v1.3.exe -OutFile ./temp/ovrie-1.3.exe
    if (!(Test-Path ./temp/ovrie-1.3.exe)){
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
        echo "Could not download! Check your firewall? If this error persists, join discord.gg/YBQCRDG for help."
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
        Pause
        Exit
    }
}else{
    echo "OpenVR-InputEmulator 1.3 is already present! Skipping download"
}
Start-Sleep -s 0.2
# download openvr input emulator dll fix
if (!(Test-Path ./temp/driver_00vrinputemulator.dll)){
    echo "Downloading the OpenVR-InputEmulator SteamVR Driver Fix"
    Invoke-WebRequest https://github.com/sharkyh20/OpenVR-InputEmulator/releases/download/SteamVR-Fix/driver_00vrinputemulator.dll -OutFile ./temp/driver_00vrinputemulator.dll
    if (!(Test-Path ./temp/driver_00vrinputemulator.dll)){
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
        echo "Could not download! Check your firewall? If this error persists, join discord.gg/YBQCRDG for help."
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
        Pause
        Exit
    }
}else{
    echo "The OpenVR-InputEmulator SteamVR Driver Fix is already present! Skipping download"
}
Start-Sleep -s 0.2
# downloading kinect sdk
if (!($arg) -or ($arg -eq 'v1') -or ($arg -eq 'v2')) {
    if ($KinectStatus -eq 0){ # xbox 360
        if (!(Test-Path "C:/Program Files/Microsoft SDKs/Kinect/v1.8")){ # and no driver
        echo "Downloading Kinect SDK 1.8 for Xbox 360 Kinect"
        Invoke-WebRequest https://download.microsoft.com/download/E/1/D/E1DEC243-0389-4A23-87BF-F47DE869FC1A/KinectSDK-v1.8-Setup.exe -OutFile ./temp/kinectv1-sdk-1.8.exe
        if (!(Test-Path ./temp/kinectv1-sdk-1.8.exe)){
            $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
            echo "Could not download! Check your firewall? If this error persists, join discord.gg/YBQCRDG for help."
            $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
            Pause
            Exit
        }
        $KinectDriverInstall = 1}
    }
    if ($KinectStatus -eq 1){ # xbox one
        if (!(Test-Path "C:/Program Files/Microsoft SDKs/Kinect/v2.0_1409")){ # and no driver
        echo "Downloading Kinect SDK 2.0 for Xbox One Kinect"
        Invoke-WebRequest https://download.microsoft.com/download/F/2/D/F2D1012E-3BC6-49C5-B8B3-5ACFF58AF7B8/KinectSDK-v2.0_1409-Setup.exe -OutFile ./temp/kinectv2-sdk-2.0.exe
        if (!(Test-Path ./temp/kinectv2-sdk-2.0.exe)){
            $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
            echo "Could not download! Check your firewall? If this error persists, join discord.gg/YBQCRDG for help."
            $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
            Pause
            Exit
        }
        $KinectDriverInstall = 1}
    }
    if ($KinectDriverInstall -eq 0){
        echo "Kinect drivers are already installed, skipping download"
    }
}

Start-Sleep -s 0.2
# INSTALLING FOR REALS
if (!(Test-Path C:/KinectToVR/)){
    echo "Extracting K2VR to the C drive (C:/KinectToVR)"
    ./temp/7zip/7za.exe x .\temp\k2vr-0.6.0r2.7z -aoa -oC:\ | Out-Null
    Rename-Item -Path "c:/KinectToVR-a0.6.0 Prime-time Test R2" -NewName "KinectToVR"
}else{
    echo "KinectToVR is already present! Skipping extract"
}
Start-Sleep -s 3
echo "Installing Visual C++ Redistribuable 2010 x64"
Start-Process ./temp/vcredist-2010-x64.exe /q -NoNewWindow -Wait
Start-Sleep -s 1
echo "Installing Visual C++ Redistribuable 2017 x64"
Start-Process ./temp/vcredist-2017-x64.exe /q -NoNewWindow -Wait
Start-Sleep -s 1
if (!(Test-Path "C:/Program Files/OpenVR-InputEmulator/OpenVR-InputEmulatorOverlay.exe")){
    echo "Installing OpenVR-InputEmulator 1.3"
    Start-Process ./temp/ovrie-1.3.exe /S -NoNewWindow -Wait
}else{
    echo "OpenVR-InputEmulator is already installed, skipping install to avoid upgrade prompt"
}
Start-Sleep -s 0.2
echo "Copying the SteamVR DLL Fix to the right folder"
Copy-Item -Force ./temp/driver_00vrinputemulator.dll -Destination "$SteamDIR/steamapps/common/SteamVR/drivers/00vrinputemulator/bin/win64"
Start-Sleep -s 0.2
if (!($arg) -or ($arg -eq 'v1') -or ($arg -eq 'v2')) {
    if ($KinectDriverInstall -ne 0){
        if ($KinectStatus -eq 0){
            echo "Running Kinect SDK 1.8 Installer for Xbox 360 Kinect"
            echo "Installation will continue when the SDK installer is done"
            Start-Process ./temp/kinectv1-sdk-1.8.exe -NoNewWindow -Wait
        }
        if ($KinectStatus -eq 1){
            echo "Running Kinect SDK 2.0 Installer for Xbox One Kinect"
            echo "Installation will continue when the SDK installer is done"
            Start-Process ./temp/kinectv2-sdk-2.0.exe -NoNewWindow -Wait
        }
    }
    if ($KinectDriverInstall -eq 0){
        echo "Skipping Kinect SDK install since the drivers are already present"
    }
}
Start-Sleep -s 0.2
# extra cleanup...
echo "Creating Start Menu entry for KinectToVR..."
echo "Downloading icon into the K2VR folder"
Invoke-WebRequest https://github.com/TripingPC/k2vr-installer/raw/master/k2vr.ico -OutFile C:/KinectToVR/k2vr.ico
New-Item -ItemType directory -Path "C:/ProgramData/Microsoft/Windows/Start Menu/Programs/KinectToVR"
Invoke-WebRequest https://github.com/TripingPC/k2vr-installer/raw/master/KinectToVR-V1.lnk -OutFile "C:/ProgramData/Microsoft/Windows/Start Menu/Programs/KinectToVR/KinectToVR (Xbox 360).lnk"
echo "Created Start Menu entry for kinectv1process.exe!"
Invoke-WebRequest https://github.com/TripingPC/k2vr-installer/raw/master/KinectToVR-V2.lnk -OutFile "C:/ProgramData/Microsoft/Windows/Start Menu/Programs/KinectToVR/KinectToVR (Xbox One).lnk"
echo "Created Start Menu entry for kinectv2process.exe!"
Start-Sleep -s 1
# force enable openvr input emu driver by abusing the steamvr settings file

# btw thanks valve for whatever JSON interpreter you use
# it literally cleans up and consolidates duplicate entries :D
# because PS' wrapper around Newtonsoft doesn't allow to change the tab length
# making this string replace hack kind of the only slightly sane method
echo "Force-enabling OpenVR-InputEmulator"
Start-Sleep -s 0.2
$NewSteamVRSettings = $SteamVRSettings.Replace(
"`"steamvr`" : {",
"`"driver_00vrinputemulator`" : {`n      `"blocked_by_safe_mode`" : false`n   },`n   `"steamvr`" : {")
echo "Safe mode blocked successfully for OpenVR-InputEmulator!"
Set-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Value $NewSteamVRSettings
# reload steamvr settings file
$SteamVRSettings = Get-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Raw
echo "Disabling SteamVR Home if not already done"
$NewSteamVRSettings = $SteamVRSettings.Replace(
"`"enableHomeApp`" : true",
"`"enableHomeApp`" : false")
Set-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Value $NewSteamVRSettings
# reload steamvr settings file
$SteamVRSettings = Get-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Raw
echo "Enabling SteamVR advanced settings (This is not OpenVR Advanced Settings!) if not already done"
$NewSteamVRSettings = $SteamVRSettings.Replace(
"`"showAdvancedSettings`" : false",
"`"showAdvancedSettings`" : true")
Set-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Value $NewSteamVRSettings
Start-Sleep -s 0.7
echo "Saved to SteamVR settings"

# TODO
# disable camera on vive/vivepro/viveroeye/index
if ($HMDStatus -in 3..5){
    # camera disable stuff
    $NewSteamVRSettings = $SteamVRSettings.Replace(
    "`"enableCamera`" : true",
    "`"enableCamera`" : false")
    Set-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Value $NewSteamVRSettings
    echo ""
    $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
    echo "Disabled camera for Vive/Vive Pro/Index!                   "
    echo "This is required because of a bug in OpenVR-InputEmulator. "
    echo "You can re-enable the camera in SteamVR settings           "
    echo "if you disable inputEmulator, but KinectToVR won't work.   "
    echo "You can find instructions to re-enable it on the website.  "
    $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
    echo ""
}
# script end
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Installation completed! You can find KinectToVR in the start menu.`n`nJoin discord.gg/Mu28W4N for help and updates.`n`nIf this is your first time, yoyu should watch the calibration tutorial on k2vr.tech", 0, "KinectToVR Installer",32)
$wshell.Popup("Rebooting your PC is recommended.", 0, "KinectToVR Installer",64)
