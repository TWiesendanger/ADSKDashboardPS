$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
$version = "0.97"
$pathMainScript = Join-Path $scriptPath "ADSKDashboard.ps1"
$pathOutput = Join-Path $scriptPath "builds\ADSKDashboard_$version.exe"
$pathIcon = Join-Path $scriptPath "res\ADSKDashboard_Icon.ico"

Invoke-ps2exe -inputFile $pathMainScript -outputFile $pathOutput -iconFile $pathIcon -noConsole -noOutput -title "ADSK Dashboard" -copyright "Tobias Wiesendanger" -verbose -version $version