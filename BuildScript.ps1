$pathExeCreater = "D:\OneDrive - MuM - OM\01_Develope\01_Powershell\Buildscript\PS2EXE-master\Module\ps2exe.ps1"
$version = "0.97"
$pathMainScript = "C:\Work\ADSKDashboardPS\ADSKDashboard.ps1"
$pathOutput = "C:\Work\ADSKDashboardPS\builds\ADSKDashboard_$($version).exe"
$pathIcon = "C:\Work\ADSKDashboardPS\res\ADSKDashboard_Icon.ico"
. $pathExeCreater 
Invoke-ps2exe -inputFile $pathMainScript -outputFile $pathOutput -iconFile $pathIcon -noConsole -noOutput -title "ADSK Dashboard" -copyright "Tobias Wiesendanger" -verbose -version $version