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
#                Config                                      #
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
} #end function New-Config

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
    } #end if config file exists
    else {
        New-Config
    }
} #end function Import-Config

function Update-Config {
    #Creates a new Config hash table with the current preferences set by the user
    $Config = @{
        'ActiveYear'    = $global:ActiveYear
        'ThemeProperty' = $global:ThemeProperty
        'Always'        = $global:Always
        'AutoClose'     = $global:AutoClose
    }
    #Export the updated config
    $Config | Export-Clixml -Path "$PathShell\res\settings\options.config"
} #end function Update-Config

#===========================================================================
#                          Check for Installation
#===========================================================================

#AutoCAD German
# $RKeyACADDE2017 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-0001-0407-1102-CF3F3A09B77D}"
# $RKeyACADDE2018 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-1001-0407-1102-CF3F3A09B77D}"
# $RKeyACADDE2019 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-2001-0407-1102-CF3F3A09B77D}"
$RKeyACADDE2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-3001-0407-1102-CF3F3A09B77D}"
$RKeyACADDE2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-4101-0407-1102-CF3F3A09B77D}"
$RKeyACADDE2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-5101-0407-1102-CF3F3A09B77D}"
$RKeyACADDE2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-6101-0407-2102-CF3F3A09B77D}"
$RKeyACADDE2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-7101-0407-2102-CF3F3A09B77D}"
$RKeyACADDE2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-8101-0407-2102-CF3F3A09B77D}"

#AutoCAD English
# $RKeyACADENU2017 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-0001-0409-1102-CF3F3A09B77D}"
# $RKeyACADENU2018 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-1001-0409-1102-CF3F3A09B77D}"
# $RKeyACADENU2019 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-2001-0409-1102-CF3F3A09B77D}" 
$RKeyACADENU2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-3001-0409-1102-CF3F3A09B77D}"
$RKeyACADENU2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-4101-0409-1102-CF3F3A09B77D}"
$RKeyACADENU2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-5101-0409-1102-CF3F3A09B77D}"
$RKeyACADENU2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-6101-0409-2102-CF3F3A09B77D}"
$RKeyACADENU2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-7101-0409-2102-CF3F3A09B77D}"
$RKeyACADENU2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-8101-0409-2102-CF3F3A09B77D}"

#AutoCAD Mechanical German
# $RKeyACADMDE2017 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-0005-0407-1102-CF3F3A09B77D}"
# $RKeyACADMDE2018 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-1005-0407-1102-CF3F3A09B77D}"
# $RKeyACADMDE2019 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-2005-0407-1102-CF3F3A09B77D}" 
$RKeyACADMDE2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-3005-0407-1102-CF3F3A09B77D}"
$RKeyACADMDE2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-4105-0407-1102-CF3F3A09B77D}"
$RKeyACADMDE2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-5105-0407-2102-CF3F3A09B77D}"
$RKeyACADMDE2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-6105-0407-2102-CF3F3A09B77D}"
$RKeyACADMDE2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-7105-0407-2102-CF3F3A09B77D}"
$RKeyACADMDE2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-8105-0407-2102-CF3F3A09B77D}"

#AutoCAD Mechanical English 
# $RKeyACADMENU2017 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-0005-0409-1102-CF3F3A09B77D}"
# $RKeyACADMENU2018 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-1005-0409-1102-CF3F3A09B77D}"
# $RKeyACADMENU2019 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-2005-0409-1102-CF3F3A09B77D}"
$RKeyACADMENU2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-3005-0409-1102-CF3F3A09B77D}"
$RKeyACADMENU2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-4105-0409-1102-CF3F3A09B77D}"
$RKeyACADMENU2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-5105-0409-2102-CF3F3A09B77D}"
$RKeyACADMENU2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-6105-0409-2102-CF3F3A09B77D}"
$RKeyACADMENU2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-7105-0409-2102-CF3F3A09B77D}"
$RKeyACADMENU2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{28B89EEF-8105-0409-2102-CF3F3A09B77D}"

#Inventor German
# $RKeyINVDE2017 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2164-0001-1031-7107D70F3DB4}"
# $RKeyINVDE2018 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2264-0001-1031-7107D70F3DB4}"
# $RKeyINVDE2019 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2364-0001-1031-7107D70F3DB4}"
$RKeyINVDE2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2464-0001-1031-7107D70F3DB4}"
$RKeyINVDE2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2564-0001-1031-7107D70F3DB4}"
$RKeyINVDE2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2664-0001-1031-7107D70F3DB4}"
$RKeyINVDE2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2764-0001-1031-7107D70F3DB4}"
$RKeyINVDE2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2864-0001-1031-7107D70F3DB4}"
$RKeyINVDE2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2964-0001-1031-7107D70F3DB4}"

#Inventor English
# $RKeyINVENU2017 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2164-0001-1033-7107D70F3DB4}"
# $RKeyINVENU2018 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2264-0001-1033-7107D70F3DB4}"
# $RKeyINVENU2019 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2364-0001-1033-7107D70F3DB4}"
$RKeyINVENU2020 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2464-0001-1033-7107D70F3DB4}"
$RKeyINVENU2021 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2564-0001-1033-7107D70F3DB4}"
$RKeyINVENU2022 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2664-0001-1033-7107D70F3DB4}"
$RKeyINVENU2023 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2764-0001-1033-7107D70F3DB4}"
$RKeyINVENU2024 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2864-0001-1033-7107D70F3DB4}"
$RKeyINVENU2025 = Test-Path "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7F4DD591-2964-0001-1033-7107D70F3DB4}"

Function HideShow($Year) {
    #Set Active Year
    $YearButton = (Get-Variable -Name "WPFY$($Year)Tab").Value
    $YearButton.IsChecked = $true

    #AutoCAD
    $RKeyACADDE = (Get-Variable -Name "RKeyACADDE$($Year)").Value
    $RKeyACADENU = (Get-Variable -Name "RKeyACADENU$($Year)").Value
    $RKeyACADMDE = (Get-Variable -Name "RKeyACADMDE$($Year)").Value
    $RKeyACADMENU = (Get-Variable -Name "RKeyACADMENU$($Year)").Value
    $RKeyINVDE = (Get-Variable -Name "RKeyINVDE$($Year)").Value
    $RKeyINVENU = (Get-Variable -Name "RKeyINVENU$($Year)").Value

    # if ($year -eq "2022" -and $RKeyINVENU -eq $false ) {
    #     # there are servicepack languagepacks that need to be detected
    #     # so if there isnt a languagepack found for the base installation look for the patches
    #     $RKeyINVENU = (Get-Variable -Name "RKeyINVENU$($Year)_2").Value
    # }
    # if ($year -eq "2022" -and $RKeyINVDE -eq $false ) {
    #     # there are servicepack languagepacks that need to be detected
    #     # so if there isnt a languagepack found for the base installation look for the patches
    #     $RKeyINVDE = (Get-Variable -Name "RKeyINVDE$($Year)_2").Value
    # }

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
}

Function OpenSoftware($SoftwareProduct, $language) {

    #region set parameter for string

    # INVENTOR and INVENTOR READ ONLY
    if ($SoftwareProduct -eq "Inventor" -and $language -eq "Deutsch") {
        $exe = "Inventor.exe"
        $arguments = "/language=DEU"
        $sFolder = "\Bin\"
    }
    elseif ($SoftwareProduct -eq "Inventor Read Only" -and $language -eq "Deutsch") {
        $SoftwareProduct = "Inventor"
        $Addition = " Read Only"
        $exe = "InvRO.exe"
        $arguments = "/language=DEU"
        $sFolder = "\Bin\"
    }
    elseif ($SoftwareProduct -eq "Inventor" -and $language -eq "English") {
        $exe = "Inventor.exe"
        $arguments = "/language=ENU"
        $sFolder = "\Bin\"
    }
    elseif ($SoftwareProduct -eq "Inventor Read Only" -and $language -eq "English") {
        $SoftwareProduct = "Inventor"
        $Addition = " Read Only"
        $exe = "InvRO.exe"
        $arguments = "/language=ENU"
        $sFolder = "\Bin\"
    }
    # AUTOCAD
    elseif ($SoftwareProduct -eq "AutoCAD" -and $language -eq "Deutsch") {
        $exe = "acad.exe"
        $arguments = "/product ACAD /language de-DE"
        $sFolder = "\"
    }
    elseif ($SoftwareProduct -eq "AutoCAD" -and $language -eq "English") {
        $exe = "acad.exe"
        $arguments = "/product ACAD /language en-US"
        $sFolder = "\"
    }
    # AUTOCAD MECHANICAL
    elseif ($SoftwareProduct -eq "AutoCAD Mechanical" -and $language -eq "Deutsch") {
        $SoftwareProduct = "AutoCAD"
        $Addition = " Mechanical"
        $exe = "acad.exe"
        $arguments = "/product ACADM /language de-DE"
        $sFolder = "\"
    }
    elseif ($SoftwareProduct -eq "AutoCAD Mechanical" -and $language -eq "English") {
        $SoftwareProduct = "AutoCAD"
        $Addition = " Mechanical"
        $exe = "acad.exe"
        $arguments = "/product ACADM /language en-US"
        $sFolder = "\"
    }
    #endregion
    
    # Invoke Process
    $pHelp = "C:\Program Files\Autodesk\" + $SoftwareProduct + " " + $global:ActiveYear + $sFolder + $exe 
    Write-Host "Starting Process: " + $pHelp + "with the following arguments: " + $arguments

    Start-Process -FilePath $pHelp -ArgumentList $arguments
    # Hide the form after opening of product
    if ($WPFAutoClose.IsChecked -eq $true) {
        $Form.Hide()
    }

    $WPFInfoText.Text = ("{0}{1} {2} - {3} {4}" -f $SoftwareProduct, $Addition, $global:ActiveYear, $language, "is loading." )
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
    # $global:WPFY2019Tab_IMG.Source = $PathShell + "\res\YButton2019.png"
    # $global:WPFY2018Tab_IMG.Source = $PathShell + "\res\YButton2018.png"
    # $global:WPFY2017Tab_IMG.Source = $PathShell + "\res\YButton2017.png"

    $global:WPFINVDE_IMG.Source = $PathShell + "\res\Inventor_DE.png"
    $global:WPFINVENU_IMG.Source = $PathShell + "\res\Inventor_ENU.png"
    $global:WPFINVRODE_IMG.Source = $PathShell + "\res\InventorRO_DE.png"
    $global:WPFINVROENU_IMG.Source = $PathShell + "\res\InventorRO_ENU.png"

    $global:WPFACADDE_IMG.Source = $PathShell + "\res\AutoCAD_DE.png"
    $global:WPFACADENU_IMG.Source = $PathShell + "\res\AutoCAD_ENU.png"

    $global:WPFACADMDE_IMG.Source = $PathShell + "\res\AutoCAD_DE.png"
    $global:WPFACADMENU_IMG.Source = $PathShell + "\res\AutoCAD_ENU.png"

    #Hide all at start
    foreach ($item in $wpfElement) {
        if ($item.name -like "*_BT") {
            $item.Visibility = "hidden"
        }
    }

    #Initialize / Import config
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
    # elseif ($global:ActiveYear -eq "2019")
    # { HideShow 2019 }
    # elseif ($global:ActiveYear -eq "2018")
    # { HideShow 2018 }
    # elseif ($global:ActiveYear -eq "2017") 
    # { HideShow 2017 }
}

#===========================================================================
#                            Systray Region
#===========================================================================

# Add the systray icon 
$Main_Tool_Icon = New-Object System.Windows.Forms.NotifyIcon
$Main_Tool_Icon.Text = "ADSK Dashboard"
$Main_Tool_Icon.Icon = $PathShell + "\res\ADSKDashboard_Icon.ico"
$Main_Tool_Icon.Visible = $true

# Garbage Collection
[System.GC]::Collect()

# Add menu exit
$Menu_Exit = New-Object System.Windows.Forms.MenuItem
$Menu_Exit.Text = "Exit"

# Add all menus as context menus
$contextmenu = New-Object System.Windows.Forms.ContextMenu
$Main_Tool_Icon.ContextMenu = $contextmenu
$Main_Tool_Icon.contextMenu.MenuItems.AddRange($Menu_Exit)

#===========================================================================
#                            XAML Reader Region
#===========================================================================

#region XAML Reader
Function ReadXAML {
    # where is the XAML file?
    $xamlFile = $PathShell + "\res\MainWindow.xaml"
    #$xamlFile = "H:\Dropbox (Data)\TWIProgrammierung\Autodesk\ADSKDashboard\ADSKDashboard\MainWindow.xaml"
    $inputXML = Get-Content $xamlFile -Raw
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
    [xml]$XAML = $inputXML
    #Read XAML
 
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
    #Print foun variables 
    Get-FormVariables
    #endregion XAML Reader
}

#Read XAML File
ReadXAML
#Get-Variable -Scope global

#Load some things
InitializeAll

#===========================================================================
#                            Button Click Events
#===========================================================================
#region Tabs
$WPFY2024Tab.Add_click( {
        Write-Host "2025"
        HideShow 2025
        $global:ActiveYear = "2025"
        Update-Config
        $WPFY2024Tab.IsChecked = $false
        $WPFY2023Tab.IsChecked = $false
        $WPFY2022Tab.IsChecked = $false
        $WPFY2021Tab.IsChecked = $false
        $WPFY2020Tab.IsChecked = $false
        # $WPFY2019Tab.IsChecked = $false
        # $WPFY2018Tab.IsChecked = $false
        # $WPFY2017Tab.IsChecked = $false

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
        # $WPFY2019Tab.IsChecked = $false
        # $WPFY2018Tab.IsChecked = $false
        # $WPFY2017Tab.IsChecked = $false

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
        # $WPFY2019Tab.IsChecked = $false
        # $WPFY2018Tab.IsChecked = $false
        # $WPFY2017Tab.IsChecked = $false

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
        # $WPFY2019Tab.IsChecked = $false
        # $WPFY2018Tab.IsChecked = $false
        # $WPFY2017Tab.IsChecked = $false

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
        # $WPFY2019Tab.IsChecked = $false
        # $WPFY2018Tab.IsChecked = $false
        # $WPFY2017Tab.IsChecked = $false

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
        # $WPFY2019Tab.IsChecked = $false
        # $WPFY2018Tab.IsChecked = $false
        # $WPFY2017Tab.IsChecked = $false
    })

# $WPFY2019Tab.Add_click( {
#         Write-Host "2019"
#         HideShow 2019
#         $global:ActiveYear = "2019"
#         Update-Config
#         $WPFY2024Tab.IsChecked = $false
#         $WPFY2023Tab.IsChecked = $false
#         $WPFY2022Tab.IsChecked = $false
#         $WPFY2021Tab.IsChecked = $false
#         $WPFY2020Tab.IsChecked = $false
#         # $WPFY2018Tab.IsChecked = $false
#         # $WPFY2017Tab.IsChecked = $false
#     })

# $WPFY2018Tab.Add_click( {
#         Write-Host "2018"
#         HideShow 2018
#         $global:ActiveYear = "2018"
#         Update-Config
#         $WPFY2023Tab.IsChecked = $false
#         $WPFY2022Tab.IsChecked = $false
#         $WPFY2021Tab.IsChecked = $false
#         $WPFY2020Tab.IsChecked = $false
#         $WPFY2019Tab.IsChecked = $false
#         # $WPFY2017Tab.IsChecked = $false
#     })

# $WPFY2017Tab.Add_click( {
#         Write-Host "2017"
#         HideShow 2017
#         $global:ActiveYear = "2017"
#         Update-Config
#         $WPFY2023Tab.IsChecked = $false
#         $WPFY2022Tab.IsChecked = $false
#         $WPFY2021Tab.IsChecked = $false
#         $WPFY2020Tab.IsChecked = $false
#         $WPFY2019Tab.IsChecked = $false
#         # $WPFY2018Tab.IsChecked = $false
#     })
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

#endregion

#===========================================================================
#                            Settings Click Events
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
#                                 Show GUI
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
