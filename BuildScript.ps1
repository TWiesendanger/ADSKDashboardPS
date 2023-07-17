$pathExeCreater = "D:\OneDrive - MuM - OM\TWIProgrammierung\Autodesk\CreateExe\ps2exe.ps1"
$version = "0.96"
$pathMainScript = "C:\MuM_Git\ADSKDashboardPS\ADSKDashboard.ps1"
$pathOutput = "C:\MuM_Git\ADSKDashboardPS\builds\ADSKDashboard_$($version).exe"
$pathIcon = "C:\MuM_Git\ADSKDashboardPS\res\ADSKDashboard_Icon.ico"
. $pathExeCreater -inputFile $pathMainScript -outputFile $pathOutput -iconFile $pathIcon -noConsole -noOutput -title "ADSK Dashboard" -copyright "Tobias Wiesendanger" -verbose -version $version