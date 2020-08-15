# request admin so installing shit doesnt suck ass
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
     $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
     Start-Process -FilePath PowerShell.exe -Verb Runas -ArgumentList $CommandLine
     Exit
    }
}

$version = "1.5.5"
$host.ui.RawUI.WindowTitle = "K2EX installer (Version $version)"

echo ""
$Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Magenta')
$Host.UI.RawUI.ForegroundColor = ($bckgrnd = 'White')
echo "  _  ___                 _ _________      _______    ________   __ "
echo " | |/ (_)               | |__   __\ \    / /  __ \  |  ____\ \ / / "
echo " | ' / _ _ __   ___  ___| |_ | | __\ \  / /| |__) | | |__   \ V /  "
echo " |  < | | '_ \ / _ \/ __| __|| |/ _ \ \/ / |  _  /  |  __|   > <   "
echo " | . \| | | | |  __/ (__| |_ | | (_) \  /  | | \ \  | |____ / . \  "
echo " |_|\_\_|_| |_|\___|\___|\__||_|\___/ \/   |_|  \_\ |______/_/ \_\ "
echo "Full-body tracking for Xbox 360, Xbox One Kinect & PlayStation Move"
# echo "   _  ___                 _  _____  __     ______    "
# echo "  | |/ (_)_ __   ___  ___| ||_   _|_\ \   / /  _ \   "
# echo "  | ' /| | '_ \ / _ \/ __| __|| |/ _ \ \ / /| |_) |  " 
# echo "  | . \| | | | |  __/ (__| |_ | | (_) \ V / |  _ <   "
# echo "  |_|\_\_|_| |_|\___|\___|\__||_|\___/ \_/  |_| \_\  "
# echo " Full-body tracking for Xbox 360 and Xbox One Kinect "
# echo ""
$Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
Start-Sleep -s 0.2

$steamVrMonitor = Get-Process vrmonitor -ErrorAction SilentlyContinue
$SteamVRDir
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

$HMDIndex = "rift-cv1","rift-s","quest","index","vive","vive-pro","vive-cosmos","windows-mr","quest-alvr","quest-vd","pimax","other"
$HMDIndexReadable = "Oculus Rift CV1","Oculus Rift S","Oculus Quest","Valve Index","HTC Vive","HTC Vive Pro","HTC Vive Cosmos","Windows Mixed Reality","Oculus Quest (ALVR)","Oculus Quest (VirtualDesktop)","Pimax","Other/Unknown"
$HMDStatus = 0
# if ($SteamDIR -eq $null){
#     $SteamDIR = (Get-Item HKCU:\Software\Valve\Steam).GetValue("SteamPath")
# }
# if (!(Test-Path "${Env:ProgramFiles(x86)}/Steam")){
#     echo "no fuck you"
#     exit
# }

# $steamProcess = Get-Process steam -ErrorAction SilentlyContinue
# if (!$steamProcess) {
    
# }
# $SteamDIR = "${Env:ProgramFiles(x86)}/Steam"
$SteamDIR = (Get-Item HKLM:/SOFTWARE/WOW6432Node/Valve/Steam).GetValue("InstallPath")
echo "Steam install found at $SteamDIR"
$SteamVRSettings = Get-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Raw
$LibraryFolders = Get-Content -Path "$SteamDIR/steamapps/libraryfolders.vdf" -Raw
$SteamVRDIR
if (Test-Path "$SteamDIR/steamapps/common/SteamVR"){
    $SteamVRDIR = "$SteamDIR/steamapps/common/SteamVR"
}
else{
    foreach ($line in $LibraryFolders.Split([Environment]::NewLine)[4..99]){
        $CurrentPath = ($line.substring(6).Replace("`"",""))
        if (Test-Path "$CurrentPath/steamapps/common/SteamVR"){
            $SteamVRDir = "$CurrentPath/steamapps/common/SteamVR"
        }
    }
}
echo "SteamVR install found at $SteamVRDir"
$SteamVRAppConfig = Get-Content -Path "$SteamDIR/config/appconfig.json" -Raw
$NewSteamVRSettings
$SteamVRSettingsJSON = $SteamVRSettings | ConvertFrom-Json

# if($SteamVRSettingsJson.LastKnown.HMDModel -eq "Oculus Rift CV1")      {$HMDStatus = 0}
# elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Oculus Rift S")    {$HMDStatus = 1}
# elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Oculus Quest")     {$HMDStatus = 2}
# elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Index")      {$HMDStatus = 3}
# elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Vive. MV")         {$HMDStatus = 4}
# elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Vive MV.")         {$HMDStatus = 4}
# elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "Vive MV")         {$HMDStatus = 4}
# elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "VIVE_Pro MV")     {$HMDStatus = 5}
# elseif($SteamVRSettingsJson.LastKnown.HMDModel -eq "HTC Vive Cosmos")  {$HMDStatus = 6}
# elseif($SteamVRSettingsJson.LastKnown.HMDManufacturer -eq "WindowsMR") {$HMDStatus = 7}
# else{$HMDStatus = 11}
# $HMDReadable = $HMDIndexReadable[$HMDStatus]
# echo "Current VR Headset: $HMDReadable"

Start-Sleep -s 0.2echo ""

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
# if (!(Test-Path ./temp/7zip/)){
#     echo "Downloading 7-Zip tools"
#     Invoke-WebRequest https://www.7-zip.org/a/7za920.zip -OutFile ./temp/7za920.zip
#     if (!(Test-Path ./temp/7za920.zip)){
#         $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
#         echo "Could not download! Check your firewall? If this error persists, join discord.gg/YBQCRDG for help."
#         $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
#         Pause
#         Exit
#     }
#     echo "Extracting..."
#     Expand-Archive -Path ./temp/7za920.Zip -DestinationPath ./temp/7zip/ -Force
# }else{
#     echo "7-Zip tools already present! Skipping download"
# }

# download kinecttovr
Start-Sleep -s 0.2
if (!(Test-Path ./temp/K2EX.zip)){
    echo "Downloading KinectToVR 0.7.1 EX (60MB file, may take a while)"
    Invoke-WebRequest 'https://github.com/KinectToVR/k2vr-website/releases/download/0.7.1ex/K2EX.zip' -OutFile ./temp/K2EX.zip
    if (!(Test-Path ./temp/K2EX.zip)){
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'DarkRed')
        echo "Could not download! Check your firewall? If this error persists, join discord.gg/YBQCRDG for help."
        $Host.UI.RawUI.BackgroundColor = ($bckgrnd = 'Black')
        Pause
        Exit
    }
}else{
    echo "KinectToVR 0.7.1 EX is already present! Skipping download"
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
if (!(Test-Path C:/K2EX/)){
    echo "Extracting K2EX to the C drive (C:/K2EX)"
    Expand-Archive -Path ./temp/K2EX.zip -DestinationPath C:/ -Force
    # ./temp/7zip/7za.exe x .\temp\K2EX.7z -aoa -oC:\ | Out-Null
    # Rename-Item -Path "c:/KinectToVR-a0.6.0 Prime-time Test R2" -NewName "KinectToVR"
}else{
    echo "K2EX is already present! Skipping extract"
}
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
echo "Creating Start Menu entry for K2EX..."
$StartMenu = [Environment]::GetFolderPath('CommonStartMenu')
New-Item -ItemType directory -Path "$StartMenu/K2EX"
Copy-Item -Path "C:/K2EX/*.lnk" -Destination "$StartMenu/K2EX"
taskkill /im StartMenuExperienceHost.exe /f
Start-Sleep -s 1
echo "Registering KinectToVR SteamVR driver..."
# C:/K2EX/driver_auto.cmd
echo "Attempting to run vrpathreg..."
& "$SteamVRDir/bin/win64/vrpathreg.exe" adddriver C:\K2EX\KinectToVR
$openvrpath = Get-Content -Path "$env:LOCALAPPDATA/../local/openvr/openvrpaths.vrpath" -Raw
if (!($openvrpath -Match "K2EX")){
    echo "vrpathreg failed! copying driver to folder"
    Copy-Item -Path C:/K2EX/KinectToVR -Destination "$SteamVRDIR/drivers"
}
# btw thanks valve for whatever JSON interpreter you use
# it literally cleans up and consolidates duplicate entries :D
# because PS' wrapper around Newtonsoft doesn't allow to change the tab length
# making this string replace hack kind of the only slightly sane method
echo "Disabling SteamVR Home if not already done"
$NewSteamVRSettings = $SteamVRSettings.Replace(
"`"enableHomeApp`" : true",
"`"enableHomeApp`" : false")
Set-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Value $NewSteamVRSettings
# reload steamvr settings file
$SteamVRSettings = Get-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Raw
echo "Enabling SteamVR advanced settings (This is not OVR Advanced Settings!) if not already done"
$NewSteamVRSettings = $SteamVRSettings.Replace(
"`"showAdvancedSettings`" : false",
"`"showAdvancedSettings`" : true")
Set-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Value $NewSteamVRSettings
Start-Sleep -s 0.7
echo "Saved to SteamVR settings"

if(!($SteamVRAppConfig -like "*Process*")){
    # bring kinect state back to xbox 360 if psmove override was enabled to enable v1process
    if($KinectStatus -eq 2){$KinectStatus = 0}
    echo "Adding K2EX as overlay app to SteamVR..."
    $NewSteamVRAppConfig = $SteamVRAppConfig.Replace(
    "   ]",
    ",`"C:`\`\K2EX`\`\KinectV$($KinectStatus+1)Process.vrmanifest`"`n   ]")
    Set-Content -Path "$SteamDIR/config/appconfig.json" -Value $NewSteamVRAppConfig
}else {
    echo "Already registered as overlay, skipping..."
}

# script end
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Installation complete! You can find K2EX in the start menu.`n`nTo enable autostart with SteamVR, enable it under Settings > Startup/Shutdown > Choose Startup Overlay Apps.", 0, "K2EX Installer",32)
$wshell.Popup("Please follow the calibration tutorial video in the #k2ex-info channel on Discord.", 0, "K2EX Installer",64)