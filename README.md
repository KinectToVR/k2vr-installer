# k2vr-installer
Installs KinectToVR and it's dependencies for people who can't read.
![main window](https://raytracing-benchmarks.are-really.cool/3trww15.png)

## About the project
I've been spending a concerning amount of time helping people set up KinectToVR mostly for VRChat.  And I've realized that most of the instructions I've been giving out could be turned into an install script.  That's what this is.  One big PS1 file compiled with PS2EXE that tries to do as clean an install of K2VR to your PC.

## What's left to do?
- Add support for launching automatically with SteamVR as an overlay.
(You can already kinda do it using the vrmanifest files included in the repo)
- More hand-holding for peculiar setups, notably Xbox One Kinect with Rift S or Lighthouse-based headsets.
- Headset detection for Pimax, ALVR and VirtualDesktop.
- ...*Maybe a GUI?*
- Better Kinect detection that isn't based on device "FriendlyNames"
- Localization

## How to contribute
Uhhh, just make a PR or an issue, fork or do whatever. I don't Git at all.
### How to build
Clone the project either manually or from the console.
Grab a copy of [PS2EXE](https://gallery.technet.microsoft.com/scriptcenter/PS2EXE-GUI-Convert-e7cb69d5) and in a Powershell window, run
```.\PS2EXE.ps1 -requireAdmin -noError -x64 -iconFile .\installer-icon.ico -InputFile .\installer.ps1 -OutputFile .\build\installer.exe```
During debugging, you should probably remove `-noError` though.

## Thanks
This project is possible because of

[KinectToVR](https://github.com/sharkyh20/KinectToVR) by Sharkyh20

[7-ZIP Command Line Tools](https://www.7-zip.org/download.html)
