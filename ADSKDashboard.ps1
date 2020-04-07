Clear-Host
#Initialize
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.ComponentModel') | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Data') | out-null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | out-null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | out-null

[System.Reflection.Assembly]::LoadWithPartialName('PresentationCore') | out-null
[System.Reflection.Assembly]::LoadFrom('res\assembly\MahApps.Metro.dll') | out-null
[System.Reflection.Assembly]::LoadFrom('res\assembly\System.Windows.Interactivity.dll') | out-null
[System.Reflection.Assembly]::LoadFrom('res\assembly\MahApps.Metro.IconPacks.dll') | out-null

# When compiled with PS2EXE the variable MyCommand contains no path anymore

if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    # Powershell script
    $PathShell = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
}
else {
    # PS2EXE compiled script
    $PathShell = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
}


#Some Settings
#On Top
#Maybe Define what year shows
#Dark Theme


#What can it check
#Installed Products depending on language?
#After that look for Languagepack


#===========================================================================
#                          Check for Installation
#===========================================================================

#AutoCAD
$RKeyACAD2017 = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R21.0\ACAD-0001:407"
$RKeyACAD2018 = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R22.0\ACAD-1001:407"
$RKeyACAD2019 = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R23.0\ACAD-2001:407"
$RKeyACAD2020 = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R23.1\ACAD-3001:407"

#AutoCAD Mechanical
$RKeyACADM2017 = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R21.0\ACAD-0005:407"
$RKeyACADM2018 = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R22.0\ACAD-1005:407"
$RKeyACADM2019 = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R23.0\ACAD-2005:407"
$RKeyACADM2020 = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R23.1\ACAD-3005:407"

#Inventor
$RKeyINV2017 = Test-Path "HKCU:\Software\Autodesk\Inventor\RegistryVersion21.0"
$RKeyINV2018 = Test-Path "HKCU:\Software\Autodesk\Inventor\RegistryVersion22.0"
$RKeyINV2019 = Test-Path "HKCU:\Software\Autodesk\Inventor\RegistryVersion23.0"
$RKeyINV2020 = Test-Path "HKCU:\Software\Autodesk\Inventor\RegistryVersion24.0"


#===========================================================================
#                          Check for English Language Packs
#===========================================================================

#AutoCAD English Languagepack
$RKeyACAD2017ENU = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R21.0\ACAD-0001:409"
$RKeyACAD2018ENU = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R22.0\ACAD-1001:409"
$RKeyACAD2019ENU = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R23.0\ACAD-2001:409"
$RKeyACAD2020ENU = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R23.1\ACAD-3001:409"

#AutoCAD Mechanical English Languagepack
$RKeyACADM2017ENU = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R21.0\ACAD-0005:409"
$RKeyACADM2018ENU = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R22.0\ACAD-1005:409"
$RKeyACADM2019ENU = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R23.0\ACAD-2005:409"
$RKeyACADM2020ENU = Test-Path "HKCU:\Software\Autodesk\AutoCAD\R23.1\ACAD-3005:409"

#Inventor English Languagepack
$RKeyINV2017ENU = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Autodesk Inventor Professional 2017 English Language Pack"
$RKeyINV2018ENU = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Autodesk Inventor Professional 2018 English Language Pack"
$RKeyINV2019ENU = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Autodesk Inventor Professional 2019 English Language Pack"
$RKeyINV2020ENU = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\Autodesk Inventor Professional 2020 English Language Pack"


Write-Host $RKeyINV2020ENU