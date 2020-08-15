$SteamDIR = (Get-Item HKLM:/SOFTWARE/WOW6432Node/Valve/Steam).GetValue("InstallPath")
$SteamVRSettings = Get-Content -Path "$SteamDIR/config/steamvr.vrsettings" -Raw
$LibraryFolders = Get-Content -Path "$SteamDIR/steamapps/libraryfolders.vdf" -Raw
$SteamVRDIR = $null
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
echo $SteamVRDir