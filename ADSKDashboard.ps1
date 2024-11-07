# get ADSKDashboard process
# Check if an instance is already running / if yes kill it and start new / didnt found a way to get the instance and maximize
$ADSKDashboard = Get-Process ADSKDashboard* -ErrorAction SilentlyContinue
Write-Host $ADSKDashboard.Id
if ($ADSKDashboard.Length -gt 1) {
    Write-Host "Already Running!"
    $ButtonType = [System.Windows.Forms.MessageBoxButtons]::OK
    $MessageIcon = [System.Windows.Forms.MessageBoxIcon]::Information
    $MessageBody = "Already Running!"
    $MessageTitle = "Error"
    $Result = [System.Windows.Forms.MessageBox]::Show($MessageBody, $MessageTitle, $ButtonType, $MessageIcon)
    exit
}

# When compiled with PS2EXE the variable MyCommand contains no path anymore
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript") {
    # Powershell script
    $PathShell = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
}
else {
    # PS2EXE compiled script
    $PathShell = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0])
}

#Initialize
[System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('WindowsFormsIntegration') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('System.ComponentModel') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('System.Data') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | Out-Null

[System.Reflection.Assembly]::LoadWithPartialName('PresentationCore') | Out-Null
[System.Reflection.Assembly]::LoadFrom("$($PathShell)\res\assembly\MahApps.Metro.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom("$($PathShell)\res\assembly\System.Windows.Interactivity.dll") | Out-Null
[System.Reflection.Assembly]::LoadFrom("$($PathShell)\res\assembly\MahApps.Metro.IconPacks.dll") | Out-Null



##############################################################
#                                                              Config                                                         #
##############################################################
function New-Config {
    #Checks for target directory and creates if non-existent 
    if (!(Test-Path -Path "$PathShell\res\settings")) {
        New-Item -ItemType Directory -Path "$PathShell\res\settings"
    }
	
    #Setup default preferences	
    #Creates hash table and .clixml config file
    $Config = @{
        'ActiveYear'    = "2025"
        'Always'        = $false
        'ThemeProperty' = "LightTheme"
        'AutoClose'     = $true
    }
    $Config | Export-Clixml -Path "$PathShell\res\settings\options.config"
    Import-Config
}

function Import-Config {
    #If a config file exists for the current user in the expected location, it is imported
    #and values from the config file are placed into global variables
    if (Test-Path -Path "$PathShell\res\settings\options.config") {
        try {
            #Imports the config file and saves it to variable $Config
            $Config = Import-Clixml -Path "$PathShell\res\settings\options.config"
			
            #Creates global variables for each config property and sets their values
            $global:ActiveYear = $Config.ActiveYear
            $global:Always = $Config.Always
            $global:ThemeProperty = $Config.ThemeProperty
            $global:AutoClose = $Config.AutoClose
            
            #Set Properties depending on Optionfile
            if ($global:ThemeProperty -eq "DarkTheme") {
                $WPFThemeSwitch.IsChecked = $true
                SetTheme("DarkTheme")
            }
            else {
                $WPFThemeSwitch.IsChecked = $false
                SetTheme("LightTheme")
            }

            if ($global:Always -eq $true) {
                $Form.Topmost = $true
                $WPFOnTop.IsChecked = $true
            }
            else {
                $Form.Topmost = $false
                $WPFOnTop.IsChecked = $false
            }
            
            if ($global:AutoClose -eq $true) {
                $WPFAutoClose.IsChecked = $true
            }
            else {
                $WPFAutoClose.IsChecked = $false
            }

        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("An error occurred importing your Config file. A new Config file will be generated for you. $_", 'Import Config Error', 'OK', 'Error')
            New-Config
        }
    }
    else {
        New-Config
    }
}

function Update-Config {
    #Creates a new Config hash table with the current preferences set by the user
    $Config = @{
        'ActiveYear'    = $global:ActiveYear
        'ThemeProperty' = $global:ThemeProperty
        'Always'        = $global:Always
        'AutoClose'     = $global:AutoClose
    }

    $Config | Export-Clixml -Path "$PathShell\res\settings\options.config"
}

#===========================================================================
#                                                                              Check for Installation                                                                         #
#===========================================================================

#AutoCAD German
$RKeyACADDE2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-3001-0407-1102-CF3F3A09B77D}"
$RKeyACADDE2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-4101-0407-1102-CF3F3A09B77D}"
$RKeyACADDE2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-5101-0407-1102-CF3F3A09B77D}"
$RKeyACADDE2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-6101-0407-2102-CF3F3A09B77D}"
$RKeyACADDE2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-7101-0407-2102-CF3F3A09B77D}"
$RKeyACADDE2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-8101-0407-2102-CF3F3A09B77D}"

#AutoCAD English
$RKeyACADENU2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-3001-0409-1102-CF3F3A09B77D}"
$RKeyACADENU2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-4101-0409-1102-CF3F3A09B77D}"
$RKeyACADENU2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-5101-0409-1102-CF3F3A09B77D}"
$RKeyACADENU2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-6101-0409-2102-CF3F3A09B77D}"
$RKeyACADENU2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-7101-0409-2102-CF3F3A09B77D}"
$RKeyACADENU2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-8101-0409-2102-CF3F3A09B77D}"

#AutoCAD Mechanical German
$RKeyACADMDE2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-3005-0407-1102-CF3F3A09B77D}"
$RKeyACADMDE2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-4105-0407-1102-CF3F3A09B77D}"
$RKeyACADMDE2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-5105-0407-2102-CF3F3A09B77D}"
$RKeyACADMDE2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-6105-0407-2102-CF3F3A09B77D}"
$RKeyACADMDE2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-7105-0407-2102-CF3F3A09B77D}"
$RKeyACADMDE2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-8105-0407-2102-CF3F3A09B77D}"

#AutoCAD Mechanical English 
$RKeyACADMENU2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-3005-0409-1102-CF3F3A09B77D}"
$RKeyACADMENU2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-4105-0409-1102-CF3F3A09B77D}"
$RKeyACADMENU2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-5105-0409-2102-CF3F3A09B77D}"
$RKeyACADMENU2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-6105-0409-2102-CF3F3A09B77D}"
$RKeyACADMENU2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-7105-0409-2102-CF3F3A09B77D}"
$RKeyACADMENU2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-8105-0409-2102-CF3F3A09B77D}"

#Inventor German
$RKeyINVDE2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2464-0001-1031-7107D70F3DB4}"
$RKeyINVDE2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2564-0001-1031-7107D70F3DB4}"
$RKeyINVDE2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2664-0001-1031-7107D70F3DB4}"
$RKeyINVDE2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2764-0001-1031-7107D70F3DB4}"
$RKeyINVDE2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2864-0001-1031-7107D70F3DB4}"
$RKeyINVDE2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2964-0001-1031-7107D70F3DB4}"

#Inventor English
$RKeyINVENU2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2464-0001-1033-7107D70F3DB4}"
$RKeyINVENU2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2564-0001-1033-7107D70F3DB4}"
$RKeyINVENU2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2664-0001-1033-7107D70F3DB4}"
$RKeyINVENU2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2764-0001-1033-7107D70F3DB4}"
$RKeyINVENU2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2864-0001-1033-7107D70F3DB4}"
$RKeyINVENU2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2964-0001-1033-7107D70F3DB4}"

# Revit (no languages)
$RKeyREV2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{F9013D08-6F9F-3F9B-8360-93C40ABE4C1B}"
$RKeyREV2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{686CE2A3-7C33-3AD5-806A-75A6E648117F}"

# Navisworks Manage (no languages)
$RKeyNAV2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{1BB907BC-4F14-3ED2-950C-39A3D99D2EFE}"
$RKeyNAV2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{8B16656F-649F-0000-99BD-374235D3813D}"

# Recap (no languages)
$RKeyRECAPYYYY = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{EA239CF5-0000-1033-0102-A32E3BC9F58A}"

Function HideShow($Year) {
    #Set Active Year
    $YearButton = (Get-Variable -Name "WPFY$($Year)Tab").Value
    $YearButton.IsChecked = $true

    $RKeyACADDE = (Get-Variable -Name "RKeyACADDE$($Year)").Value
    $RKeyACADENU = (Get-Variable -Name "RKeyACADENU$($Year)").Value
    $RKeyACADMDE = (Get-Variable -Name "RKeyACADMDE$($Year)").Value
    $RKeyACADMENU = (Get-Variable -Name "RKeyACADMENU$($Year)").Value
    $RKeyINVDE = (Get-Variable -Name "RKeyINVDE$($Year)").Value
    $RKeyINVENU = (Get-Variable -Name "RKeyINVENU$($Year)").Value
    if ($Year -ge 2024) {
        $RKeyREV = (Get-Variable -Name "RKeyREV$($Year)").Value
    }
    if ($Year -ge 2024) {
        $RKeyNAV = (Get-Variable -Name "RKeyNAV$($Year)").Value
    }
    $RKeyRECAP = (Get-Variable -Name "RKeyRECAPYYYY").Value


    if ($RKeyACADDE) { $WPFACADDE_BT.Visibility = "visible" }
    else { $WPFACADDE_BT.Visibility = "hidden" }
    
    if ($RKeyACADENU) { $WPFACADENU_BT.Visibility = "visible" }
    else { $WPFACADENU_BT.Visibility = "hidden" }

    #AutoCAD Mechanical
    if ($RKeyACADMDE) { $WPFACADMDE_BT.Visibility = "visible" }
    else { $WPFACADMDE_BT.Visibility = "hidden" }
        
    if ($RKeyACADMENU) { $WPFACADMENU_BT.Visibility = "visible" }
    else { $WPFACADMENU_BT.Visibility = "hidden" }

    #Inventor / Inventor Read Only
    if ($RKeyINVDE) { 
        $WPFINVDE_BT.Visibility = "visible"
        if ($Year -gt 2020) {
            $WPFINVRODE_BT.Visibility = "visible" 
        }
        else {
            $WPFINVRODE_BT.Visibility = "hidden"  
        }
    }
    else { 
        $WPFINVDE_BT.Visibility = "hidden"
        $WPFINVRODE_BT.Visibility = "hidden" 
    }
            
    if ($RKeyINVENU) {
        $WPFINVENU_BT.Visibility = "visible"
        if ($Year -gt 2020) {
            $WPFINVROENU_BT.Visibility = "visible"
        }
        else {
            $WPFINVROENU_BT.Visibility = "hidden"
        }
    }
    else {
        $WPFINVENU_BT.Visibility = "hidden" 
        $WPFINVROENU_BT.Visibility = "hidden"
    }

    if ($RKeyREV) {
        $WPFREVITDE_BT.Visibility = "visible"
        $WPFREVITENG_BT.Visibility = "visible"
    }
    else {
        $WPFREVITDE_BT.Visibility = "hidden"
        $WPFREVITENG_BT.Visibility = "hidden"
    }

    if ($RKeyNav) {
        $WPFNAVDE_BT.Visibility = "visible"
        $WPFNAVENG_BT.Visibility = "visible"
    }
    else {
        $WPFNAVDE_BT.Visibility = "hidden"
        $WPFNAVENG_BT.Visibility = "hidden"
    }

    if ($RKeyRECAPYYYY) {
        $WPFRECAP_BT.Visibility = "visible"
    }
    else {
        $WPFRECAP_BT.Visibility = "hidden"
    }
    
}

Function OpenSoftware($SoftwareProduct, $language) {

    # INVENTOR and INVENTOR READ ONLY
    if ($SoftwareProduct -eq "Inventor" -and $language -eq "Deutsch") {
        $exe = "Inventor.exe"
        $arguments = "/language=DEU"
        $sFolder = "\Bin\"
        $noYear = $false
    }
    elseif ($SoftwareProduct -eq "Inventor Read Only" -and $language -eq "Deutsch") {
        $SoftwareProduct = "Inventor"
        $Addition = " Read Only"
        $exe = "InvRO.exe"
        $arguments = "/language=DEU"
        $sFolder = "\Bin\"
        $noYear = $false
    }
    elseif ($SoftwareProduct -eq "Inventor" -and $language -eq "English") {
        $exe = "Inventor.exe"
        $arguments = "/language=ENU"
        $sFolder = "\Bin\"
        $noYear = $false
    }
    elseif ($SoftwareProduct -eq "Inventor Read Only" -and $language -eq "English") {
        $SoftwareProduct = "Inventor"
        $Addition = " Read Only"
        $exe = "InvRO.exe"
        $arguments = "/language=ENU"
        $sFolder = "\Bin\"
        $noYear = $false
    }
    # AUTOCAD
    elseif ($SoftwareProduct -eq "AutoCAD" -and $language -eq "Deutsch") {
        $exe = "acad.exe"
        $arguments = "/product ACAD /language de-DE"
        $sFolder = "\"
        $noYear = $false
    }
    elseif ($SoftwareProduct -eq "AutoCAD" -and $language -eq "English") {
        $exe = "acad.exe"
        $arguments = "/product ACAD /language en-US"
        $sFolder = "\"
        $noYear = $false
    }
    # AUTOCAD MECHANICAL
    elseif ($SoftwareProduct -eq "AutoCAD Mechanical" -and $language -eq "Deutsch") {
        $SoftwareProduct = "AutoCAD"
        $Addition = " Mechanical"
        $exe = "acad.exe"
        $arguments = "/product ACADM /language de-DE"
        $sFolder = "\"
        $noYear = $false
    }
    elseif ($SoftwareProduct -eq "AutoCAD Mechanical" -and $language -eq "English") {
        $SoftwareProduct = "AutoCAD"
        $Addition = " Mechanical"
        $exe = "acad.exe"
        $arguments = "/product ACADM /language en-US"
        $sFolder = "\"
        $noYear = $false
    }
    elseif ($SoftwareProduct -eq "Revit" -and $language -eq "Deutsch") {
        $SoftwareProduct = "Revit"
        $exe = "Revit.exe"
        $arguments = "/language DEU"
        $sFolder = "\"
        $noYear = $false
    }
    elseif ($SoftwareProduct -eq "Revit" -and $language -eq "English") {
        $SoftwareProduct = "Revit"
        $exe = "Revit.exe"
        $arguments = "/language ENU"
        $sFolder = "\"
        $noYear = $false
    }
    elseif ($SoftwareProduct -eq "Navisworks Manage" -and $language -eq "Deutsch") {
        $SoftwareProduct = "Navisworks Manage"
        $exe = "Roamer.exe"
        $arguments = " /licensing AdLM /lang de-DE"
        $sFolder = "\"
        $noYear = $false
    }
    elseif ($SoftwareProduct -eq "Navisworks Manage" -and $language -eq "English") {
        $SoftwareProduct = "Navisworks Manage"
        $exe = "Roamer.exe"
        $arguments = " /licensing AdLM /lang en-US"
        $sFolder = "\"
        $noYear = $false
    }
    elseif ($SoftwareProduct -eq "Autodesk Recap") {
        $SoftwareProduct = "Autodesk Recap"
        $exe = "ReCap.exe"
        $arguments = ""
        $sFolder = "\"
        $noYear = $true
    }

    #endregion
    
    # Invoke Process
    if ($noYear -eq $true) {
        $pHelp = "C:\Program Files\Autodesk\" + $SoftwareProduct + $sFolder + $exe 
        $WPFInfoText.Text = ("{0}{1} {2} - {3}" -f $SoftwareProduct, $Addition, $language, "is loading." )
    }
    else {
        $pHelp = "C:\Program Files\Autodesk\" + $SoftwareProduct + " " + $global:ActiveYear + $sFolder + $exe 
        $WPFInfoText.Text = ("{0}{1} {2} - {3} {4}" -f $SoftwareProduct, $Addition, $global:ActiveYear, $language, "is loading." )
    }
    Write-Host "Starting Process: " + $pHelp + "with the following arguments: " + $arguments

    if ([string]::IsNullOrWhiteSpace($arguments)) {
        Start-Process -FilePath $pHelp
    }
    else {
        Start-Process -FilePath $pHelp -ArgumentList $arguments
    }
    # Hide the form after opening of product
    if ($WPFAutoClose.IsChecked -eq $true) {
        $Form.Hide()
    }

    Write-Host $WPFInfoText.Text
    $WPFInfoDialog.IsOpen = $true  
}

function SetTheme($Themestr) {
    $Theme = [MahApps.Metro.ThemeManager]::DetectAppStyle($Form)
    if ($Themestr -eq "DarkTheme") {
        $global:ThemeProperty = "DarkTheme"
        [MahApps.Metro.ThemeManager]::ChangeAppStyle($Form, $Theme.Item2, [MahApps.Metro.ThemeManager]::GetAppTheme("BaseDark"))
        $Theme = [MahApps.Metro.ThemeManager]::DetectAppStyle($Form)
        [MahApps.Metro.ThemeManager]::ChangeAppStyle($Form, [MahApps.Metro.ThemeManager]::GetAccent("Steel"), $Theme.Item1)
    }
    else {
        $global:ThemeProperty = "LightTheme"
        [MahApps.Metro.ThemeManager]::ChangeAppStyle($Form, $Theme.Item2, [MahApps.Metro.ThemeManager]::GetAppTheme("BaseLight")) 
        $Theme = [MahApps.Metro.ThemeManager]::DetectAppStyle($Form)
        [MahApps.Metro.ThemeManager]::ChangeAppStyle($Form, [MahApps.Metro.ThemeManager]::GetAccent("Cobalt"), $Theme.Item1)
    }
}

function InitializeAll {
    # Icon and Image Source
    $global:Form.Icon = $PathShell + "\res\ADSKDashboard_Icon.png"
    $global:WPFY2025Tab_IMG.Source = $PathShell + "\res\YButton2025.png"
    $global:WPFY2024Tab_IMG.Source = $PathShell + "\res\YButton2024.png"
    $global:WPFY2023Tab_IMG.Source = $PathShell + "\res\YButton2023.png"
    $global:WPFY2022Tab_IMG.Source = $PathShell + "\res\YButton2022.png"
    $global:WPFY2021Tab_IMG.Source = $PathShell + "\res\YButton2021.png"
    $global:WPFY2020Tab_IMG.Source = $PathShell + "\res\YButton2020.png"

    $global:WPFINVDE_IMG.Source = $PathShell + "\res\Inventor_DE.png"
    $global:WPFINVENU_IMG.Source = $PathShell + "\res\Inventor_ENU.png"
    $global:WPFINVRODE_IMG.Source = $PathShell + "\res\InventorRO_DE.png"
    $global:WPFINVROENU_IMG.Source = $PathShell + "\res\InventorRO_ENU.png"

    $global:WPFACADDE_IMG.Source = $PathShell + "\res\AutoCAD_DE.png"
    $global:WPFACADENU_IMG.Source = $PathShell + "\res\AutoCAD_ENU.png"

    $global:WPFACADMDE_IMG.Source = $PathShell + "\res\AutoCAD_DE.png"
    $global:WPFACADMENU_IMG.Source = $PathShell + "\res\AutoCAD_ENU.png"

    $global:WPFREVITDE_IMG.Source = $PathShell + "\res\Revit_DE.png"
    $global:WPFREVITENG_IMG.Source = $PathShell + "\res\Revit_ENU.png"
    
    $global:WPFNAVDE_IMG.Source = $PathShell + "\res\Navisworks_DE.png"
    $global:WPFNAVENG_IMG.Source = $PathShell + "\res\Navisworks_ENU.png"

    $global:WPFRECAP_IMG.Source = $PathShell + "\res\Recap.png"

    #Hide all at start
    foreach ($item in $wpfElement) {
        if ($item.name -like "*_BT") {
            $item.Visibility = "hidden"
        }
    }

    Import-Config

    if ($global:ActiveYear -eq "2025")
    { HideShow 2025 }
    elseif ($global:ActiveYear -eq "2024")
    { HideShow 2024 }
    elseif ($global:ActiveYear -eq "2023")
    { HideShow 2023 }
    elseif ($global:ActiveYear -eq "2022")
    { HideShow 2022 }
    elseif ($global:ActiveYear -eq "2021")
    { HideShow 2021 }
    elseif ($global:ActiveYear -eq "2020")
    { HideShow 2020 }
}

#===========================================================================
#                                                                                    Systray Region
#===========================================================================

# Add the systray icon 
$Main_Tool_Icon = New-Object System.Windows.Forms.NotifyIcon
$Main_Tool_Icon.Text = "ADSK Dashboard"
$Main_Tool_Icon.Icon = $PathShell + "\res\ADSKDashboard_Icon.ico"
$Main_Tool_Icon.Visible = $true

[System.GC]::Collect()

$Menu_Exit = New-Object System.Windows.Forms.MenuItem
$Menu_Exit.Text = "Exit"

# Add all menus as context menus
$contextmenu = New-Object System.Windows.Forms.ContextMenu
$Main_Tool_Icon.ContextMenu = $contextmenu
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Exit)

#===========================================================================
#                            XAML Reader Region
#===========================================================================


Function ReadXAML {
    $xamlFile = $PathShell + "\res\MainWindow.xaml"
    #$xamlFile = "H:\Dropbox (Data)\TWIProgrammierung\Autodesk\ADSKDashboard\ADSKDashboard\MainWindow.xaml"
    $inputXML = Get-Content $xamlFile -Raw
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML
 
    $reader = (New-Object System.Xml.XmlNodeReader $xaml)
    try {
        $global:Form = [Windows.Markup.XamlReader]::Load( $reader )
    }
    catch {
        Write-Warning "Unable to parse XML, with error: $($Error[0])`n Ensure that there are NO SelectionChanged or TextChanged properties in your textboxes (PowerShell cannot process them)"
        throw
    }
    #===========================================================================
    # Load XAML Objects In PowerShell
    #===========================================================================
    $global:wpfElement = @()
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object { "trying item $($_.Name)"
        try {
            Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop -Scope global
            $global:wpfElement += (Get-Variable -Name "WPF$($_.Name)").Value
            Write-Host "WPF$($_.Name)"
        }
        catch { throw }
    }

    Function Get-FormVariables {
        if ($global:ReadmeDisplay -ne $true) { Write-Host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow; $global:ReadmeDisplay = $true }
        Write-Host "Found the following interactable elements from our form" -ForegroundColor Cyan
        Get-Variable $global:WPF*
    }
    Get-FormVariables
}

ReadXAML
InitializeAll

#===========================================================================
#                                                            Button Click Events
#===========================================================================
#region Tabs
$WPFY2025Tab.Add_click( {
        Write-Host "2025"
        HideShow 2025
        $global:ActiveYear = "2025"
        Update-Config
        $WPFY2024Tab.IsChecked = $false
        $WPFY2023Tab.IsChecked = $false
        $WPFY2022Tab.IsChecked = $false
        $WPFY2021Tab.IsChecked = $false
        $WPFY2020Tab.IsChecked = $false
    })
$WPFY2024Tab.Add_click( {
        Write-Host "2024"
        HideShow 2024
        $global:ActiveYear = "2024"
        Update-Config
        $WPFY2025Tab.IsChecked = $false
        $WPFY2023Tab.IsChecked = $false
        $WPFY2022Tab.IsChecked = $false
        $WPFY2021Tab.IsChecked = $false
        $WPFY2020Tab.IsChecked = $false
    })
$WPFY2023Tab.Add_click( {
        Write-Host "2023"
        HideShow 2023
        $global:ActiveYear = "2023"
        Update-Config
        $WPFY2025Tab.IsChecked = $false
        $WPFY2024Tab.IsChecked = $false
        $WPFY2022Tab.IsChecked = $false
        $WPFY2021Tab.IsChecked = $false
        $WPFY2020Tab.IsChecked = $false
    })
$WPFY2022Tab.Add_click( {
        Write-Host "2022"
        HideShow 2022
        $global:ActiveYear = "2022"
        Update-Config
        $WPFY2025Tab.IsChecked = $false
        $WPFY2024Tab.IsChecked = $false
        $WPFY2023Tab.IsChecked = $false
        $WPFY2021Tab.IsChecked = $false
        $WPFY2020Tab.IsChecked = $false
    })
$WPFY2021Tab.Add_click( {
        Write-Host "2021"
        HideShow 2021
        $global:ActiveYear = "2021"
        Update-Config
        $WPFY2025Tab.IsChecked = $false
        $WPFY2024Tab.IsChecked = $false
        $WPFY2023Tab.IsChecked = $false
        $WPFY2022Tab.IsChecked = $false
        $WPFY2020Tab.IsChecked = $false
    })

$WPFY2020Tab.Add_click( {
        Write-Host "2020"
        HideShow 2020
        $global:ActiveYear = "2020"
        Update-Config
        $WPFY2025Tab.IsChecked = $false
        $WPFY2024Tab.IsChecked = $false
        $WPFY2023Tab.IsChecked = $false
        $WPFY2022Tab.IsChecked = $false
        $WPFY2021Tab.IsChecked = $false
    })
#endregion

#region Software Buttons
#INVENTOR
$WPFINVDE_BT.Add_click( {
        OpenSoftware "Inventor" "Deutsch"
    })
$WPFINVENU_BT.Add_click( {
        OpenSoftware "Inventor" "English"
    })

#INVENTOR READ ONLY
$WPFINVRODE_BT.Add_click( {
        OpenSoftware "Inventor Read Only" "Deutsch"
    })
$WPFINVROENU_BT.Add_click( {
        OpenSoftware "Inventor Read Only" "English"
    })

#AUTOCAD
$WPFACADDE_BT.Add_click( {
        OpenSoftware "AutoCAD" "Deutsch"
    })
$WPFACADENU_BT.Add_click( {
        OpenSoftware "AutoCAD" "English"
    })

#AUTOCAD MECHANICAL 
$WPFACADMDE_BT.Add_click( {
        OpenSoftware "AutoCAD Mechanical" "Deutsch"
    })
$WPFACADMENU_BT.Add_click( {
        OpenSoftware "AutoCAD Mechanical" "English"
    })

#REVIT
$WPFREVITDE_BT.Add_click( {
        OpenSoftware "Revit" "Deutsch"
    })
$WPFREVITENG_BT.Add_click( {
        OpenSoftware "Revit" "English"
    })

#NAVISWORKS MANAGE
$WPFNAVDE_BT.Add_click( {
        OpenSoftware "Navisworks Manage" "Deutsch"
    })
$WPFNAVENG_BT.Add_click( {
        OpenSoftware "Navisworks Manage" "English"
    })

#Recap
$WPFRECAP_BT.Add_click( {
        OpenSoftware "Autodesk Recap"
    })

#endregion

#===========================================================================
#                                                                               Settings Click Events                                                                          #
#===========================================================================

$WPFSettingsButton.Add_click( {
        $WPFFlyOutContent.IsOpen = $true
    })

$WPFHelpButton.Add_click( {
        start "https://github.com/TWiesendanger/ADSKDashboardPS"
    })

$WPFOnTop.Add_click( {
        if ($WPFOnTop.IsChecked -eq $true) {
            $Form.Topmost = $true
            $global:Always = $true
            Update-Config
        }
        else {
            $Form.Topmost = $false
            $global:Always = $false
            Update-Config
        }
    })

$WPFAutoClose.Add_click( {
        if ($WPFAutoClose.IsChecked -eq $true) {
            $global:AutoClose = $true
            Update-Config
        }
        else {
            $global:AutoClose = $false
            Update-Config
        }
    })

$WPFThemeSwitch.Add_Click( {
        if ($WPFThemeSwitch.IsChecked -eq $true) {
            #BaseDark
            SetTheme("DarkTheme")
            Update-Config
        }
        else {
            SetTheme("LightTheme")
            Update-Config
        }
    })

$Form.Add_closing( {
        Write-Host "CLOSE"
        $Form.Hide()
    })

#===========================================================================
#                                                                                        Show GUI            
#===========================================================================

$WPFClose.Add_Click( {
        $Form.Hide()
    })

$Menu_Exit.Add_click( {
        $Main_Tool_Icon.Visible = $false
        $Form.Close()
        Stop-Process $pid
    })

# Action after a click on the main systray icon
$Main_Tool_Icon.Add_Click( {     
        $Form.ShowDialog()
        $Form.Activate() 
    })

# Close the Form GUI if it loses focus
$Form.Add_Deactivated( {

        if ($global:AutoClose -eq $true) {
            $Form.Hide()
        }
    })

$Form.ShowDialog()
try {
    $Form.Activate()
}
catch { }

#Keep it alive while closed to systray
$appContext = New-Object System.Windows.Forms.ApplicationContext 
[void][System.Windows.Forms.Application]::Run($appContext)