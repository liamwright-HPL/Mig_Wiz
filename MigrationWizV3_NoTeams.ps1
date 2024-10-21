#########################################################################
#
#        Name:   MigrationWiz.ps1
#    Modified:   2024-10-14
#      Author:   Liam Wright
#     Version:	 1.0.0
# Description:   Creates UI for user to run through reconfiguring OneDrive, Outlook and MFA following tenant migration.
#
#########################################################################

<# Icon attribution
No start - <a href="https://www.flaticon.com/free-icons/no-entry" title="no entry icons">No entry icons created by Freepik - Flaticon</a>
Loading - <a href="https://www.flaticon.com/free-icons/loading" title="loading icons">Loading icons created by Krystsina Mikhailouskaya - Flaticon</a>
Completed - <a href="https://www.flaticon.com/free-icons/check-box" title="check box icons">Check box icons created by Hogr - Flaticon</a>
#>

# =======================================
# FUNCTIONS

# Closes OneDrive and clears cache. Also includes UI updates.

function Configure-OneDrive { 

# Close OneDrive processes
function Close-OneDrive {
    If ((Get-Process "OneDrive" -ea SilentlyContinue)) { 
            Stop-Process -Name "OneDrive" -Force -ea SilentlyContinue
                    }
            Start-Sleep -Seconds 1
}

#Clear OneDrive Cache
function Clear-OneDriveCache {
    If (Test-Path "$env:LOCALAPPDATA\Microsoft\OneDrive\settings\Business1") {
                        Remove-Item "$env:LOCALAPPDATA\Microsoft\OneDrive\settings\Business1" -Recurse -Force
                    }
                        
                    # Löschung OneDrive Folder
                    If (Test-Path "$env:USERPROFILE\OneDrive - Migros") {
                        Remove-Item "$env:USERPROFILE\OneDrive - Migros" -Recurse -Force
                    }

                    # Löschung Account + Settings in Registry
                    If (Test-Path -Path HKCU:\Software\Microsoft\OneDrive\Accounts\Business1) {
                        Remove-Item -Path HKCU:\Software\Microsoft\OneDrive\Accounts\Business1 -Recurse -Force
                    }

                    # Löschung OneDrive NameSpace in Registry
                    If (Test-Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace) {
                        Get-ChildItem -Path HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace | ForEach-Object {
                            If ((Get-ItemProperty Registry::$_).'(default)' -eq "OneDrive - Migros") {
                                Remove-Item -Path Registry::$($_.Name) -Force
                            }
                        }
                    }
                # Trigger OneDrive Reset
                If (Test-Path "$env:LOCALAPPDATA\Microsoft\OneDrive\onedrive.exe") {
                    Start-Process -FilePath "$env:LOCALAPPDATA\Microsoft\OneDrive\onedrive.exe" -ArgumentList "/reset" -Wait
                }
}

                $btnNext.IsEnabled = $false
                $txtDescription.Text = "Updating OneDrive, please wait...`n`n"
                $txtOneDrive_Status.Text = "Running..."
                $imgRunOneDrive.Visibility = "Visible"
                $imgNSOneDrive.Visibility = "Hidden"
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},"Render") #Forces UI to update $txtDescription.Text without having to wait for the functions to complete
                Close-OneDrive
                Clear-OneDriveCache
                $txtDescription.Inlines.Add("OneDrive has been updated.`n`n")
                $txtDescription.Inlines.Add("Click Next to sign into OneDrive. Once completed, return here to continue the migration.`n`n")
                $txtDescription.Inlines.Add("NOTE: Please see this Jira guide ")
                $txtDescription.Inlines.Add($hyperlinkJiraOneDrive)
                $txtDescription.Inlines.Add(" for further assistance.")
                $btnNext.IsEnabled = $True
                $UI.Topmost = $false # Disables window from being on top

}

#Closes Outlook and creates new profile. Also includes UI updates
function Configure-Outlook {
# Close Outlook processes
function Close-Outlook {
    If ((Get-Process "Outlook" -ea SilentlyContinue)) { 
            Stop-Process -Name "Outlook" -Force -ea SilentlyContinue
                    } 
            Start-Sleep 1          
}

#Create new Outlook profile
function New-OutlookProfile {
    New-Item -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles\TenantMigration" -Force
    New-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\Outlook" -Name DefaultProfile -Value "TenantMigration"  -Force
}

                $btnNext.IsEnabled = $false
                $txtDescription.Text = "Updating Outlook, please wait...`n`n"
                $txtOutlook_Status.Text = "Running..."
                $imgRunOutlook.Visibility = "Visible"
                $imgNSOutlook.Visibility = "Hidden"
                [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([Action]{},"Render") #Forces UI to update $txtDescription.Text without having to wait for the functions to complete
                Close-Outlook
                New-OutlookProfile
                $txtDescription.Inlines.Add("Outlook has been updated.`n`n")
                $txtDescription.Inlines.Add("Click Next to sign into Outlook. Once completed, return here to continue the migration.`n`n")
                $txtDescription.Inlines.Add("NOTE: Please see this Jira guide ")
                $txtDescription.Inlines.Add($hyperlinkJiraOutlook)
                $txtDescription.Inlines.Add(" for further assistance.")
                $btnNext.IsEnabled = $True
                $UI.Topmost = $false # Disables window from being on top

}

function Configure-MFA {
                $btnNext.IsEnabled = $false
                $txtMFA_Status.Text = "Running..."
                $imgRunMFA.Visibility = "Visible"
                $imgNSMFA.Visibility = "Hidden"
                $txtDescription.Text = "Click Next to setup Multi-Factor Authentication. Once completed, return here to continue the migration.`n`n"
                $txtDescription.Inlines.Add("NOTE: Please see this Jira guide ")
                $txtDescription.Inlines.Add($hyperlinkJiraMFA)
                $txtDescription.Inlines.Add(" for further assistance.")
                $btnNext.IsEnabled = $True
                $UI.Topmost = $false # Disables window from being on top
}
# =======================================
# MIGRATIONWIZ FUNCTION
function Run-MigrationWiz {
    # Load the Windows Presentation Format assemblies
    [System.Reflection.Assembly]::LoadWithPartialName('PresentationCore') | out-null
    [System.Reflection.Assembly]::LoadWithPartialName('PresentationFramework') | out-null
    
    # Embed xaml code in script (Window)
    [string]$XAML_Main = @"
<Window x:Name="Hotelplan_Migration_Wiz" x:Class="MigrationWiz.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:MigrationWiz"
        mc:Ignorable="d"
        Title="HotelplanMigrationWiz" Height="450" Width="800" ResizeMode="NoResize" Topmost="True" WindowStyle="None" >
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="74*"/>
            <ColumnDefinition Width="206*"/>
            <ColumnDefinition Width="423*"/>
            <ColumnDefinition Width="97*"/>
        </Grid.ColumnDefinitions>
        <Frame x:Name="frame" Content="" NavigationUIVisibility="Hidden" Grid.ColumnSpan="2"/>
        <Rectangle
        Width="418" VerticalAlignment="Top" Stroke="Black" Margin="92,200,0,0" Height="175" HorizontalAlignment="Left" Grid.ColumnSpan="2" Grid.Column="2"/>
        <Button x:Name="btnExit" Content="Exit" HorizontalAlignment="Left" Margin="92,380,0,0" VerticalAlignment="Top" Height="60" Width="200" Background="White" BorderThickness="2,2,2,2" FontSize="20" Grid.Column="2"/>
        <Button x:Name="btnNext" Content="Next" HorizontalAlignment="Left" Margin="310,380,0,0" VerticalAlignment="Top" Height="60" Width="200" Background="#FF979797" BorderThickness="2,2,2,2" FontSize="20" Grid.ColumnSpan="2" Grid.Column="2"/>
        <TextBlock x:Name="txtDescription" HorizontalAlignment="Left" Margin="102,200,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Height="175" Width="397" FontSize="14" Grid.ColumnSpan="2" Grid.Column="2"/>
        <TextBlock x:Name="txtOneDrive" HorizontalAlignment="Left" Margin="10,203,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Height="50" Width="228" FontSize="20" Grid.ColumnSpan="2"/>
        <TextBlock x:Name="txtOutlook" HorizontalAlignment="Left" Margin="10,262,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Height="50" Width="228" FontSize="20" Grid.ColumnSpan="2"/>
        <TextBlock x:Name="txtMFA" HorizontalAlignment="Left" Margin="10,325,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Height="50" Width="228" FontSize="20" Grid.ColumnSpan="2"/>
        <Image x:Name="imgHPL_logo" HorizontalAlignment="Left" Height="100" Margin="10,10,0,0" VerticalAlignment="Top" Width="390" Source="https://www.hotelplan.co.uk/assets/imgs/logo.png" Grid.ColumnSpan="3"/>
        <Rectangle x:Name="aesTopLine_Copy" HorizontalAlignment="Left" Height="1" Margin="10,182,0,0" Stroke="Black" VerticalAlignment="Top" Width="780" Grid.ColumnSpan="4"/>
        <TextBlock x:Name="txtTitle" HorizontalAlignment="Left" Margin="10,137,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Height="50" Width="554" FontSize="20" Grid.ColumnSpan="3"/>
        <Button x:Name="btnGuides" Content="Jira Guides" HorizontalAlignment="Left" Margin="310,137,0,0" VerticalAlignment="Top" Height="30" Width="92" Background="White" BorderThickness="2,2,2,2" FontSize="12" Grid.Column="2"/>
        <Rectangle x:Name="aesTopLine_Copy1" HorizontalAlignment="Left" Height="1" Margin="10,182,0,0" Stroke="Black" VerticalAlignment="Top" Width="780" Grid.ColumnSpan="4"/>
        <TextBlock x:Name="txtOneDrive_Status" HorizontalAlignment="Left" Margin="170,207,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Height="50" Width="123" FontSize="16" Grid.ColumnSpan="2" Grid.Column="1"/>
        <TextBlock x:Name="txtOutlook_Status" HorizontalAlignment="Left" Margin="170,266,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Height="50" Width="123" FontSize="16" Grid.ColumnSpan="2" Grid.Column="1"/>
        <TextBlock x:Name="txtMFA_Status" HorizontalAlignment="Left" Margin="170,329,0,0" TextWrapping="Wrap" Text="TextBlock" VerticalAlignment="Top" Height="50" Width="123" FontSize="16" Grid.ColumnSpan="2" Grid.Column="1"/>
        <Image x:Name="imgNSOneDrive" HorizontalAlignment="Left" Height="26" Margin="131,206,0,0" VerticalAlignment="Top" Width="25" Source="https://cdn-icons-png.flaticon.com/512/4207/4207834.png" Grid.Column="1"/>
        <Image x:Name="imgNSOutlook" HorizontalAlignment="Left" Height="25" Margin="131,266,0,0" VerticalAlignment="Top" Width="25" Source="https://cdn-icons-png.flaticon.com/512/4207/4207834.png" Grid.Column="1"/>
        <Image x:Name="imgNSMFA" HorizontalAlignment="Left" Height="25" Margin="131,329,0,0" VerticalAlignment="Top" Width="25" Source="https://cdn-icons-png.flaticon.com/512/4207/4207834.png" Visibility="Visible" Grid.Column="1"/>
        <Image x:Name="imgCptOneDrive" HorizontalAlignment="Left" Height="25" Margin="131,206,0,0" VerticalAlignment="Top" Width="25" Source="https://cdn-icons-png.flaticon.com/512/845/845646.png" Grid.Column="1" Visibility="Hidden"/>
        <Image x:Name="imgCptOutlook" HorizontalAlignment="Left" Height="25" Margin="131,266,0,0" VerticalAlignment="Top" Width="25" Source="https://cdn-icons-png.flaticon.com/512/845/845646.png" Grid.Column="1" Visibility="Hidden"/>
        <Image x:Name="imgCptMFA" HorizontalAlignment="Left" Height="25" Margin="131,329,0,0" VerticalAlignment="Top" Width="25" Source="https://cdn-icons-png.flaticon.com/512/845/845646.png" Grid.Column="1" Visibility="Hidden"/>
        <Image x:Name="imgRunOneDrive" HorizontalAlignment="Left" Height="25" Margin="131,206,0,0" VerticalAlignment="Top" Width="25" Source="https://cdn-icons-png.flaticon.com/512/6356/6356587.png" Grid.Column="1" Visibility="Hidden"/>
        <Image x:Name="imgRunOutlook" HorizontalAlignment="Left" Height="25" Margin="131,266,0,0" VerticalAlignment="Top" Width="25" Source="https://cdn-icons-png.flaticon.com/512/6356/6356587.png" Grid.Column="1" Visibility="Hidden"/>
        <Image x:Name="imgRunMFA" HorizontalAlignment="Left" Height="25" Margin="131,329,0,0" VerticalAlignment="Top" Width="25" Source="https://cdn-icons-png.flaticon.com/512/6356/6356587.png" Grid.Column="1" Visibility="Hidden"/>
        <CheckBox x:Name="chkbxOneDrive" Grid.Column="2" Content="" HorizontalAlignment="Left" Margin="62,211,0,0" VerticalAlignment="Top" IsEnabled="False" Visibility="Hidden" IsChecked="False"/>
        <CheckBox x:Name="chkbxOutlook" Grid.Column="2" Content="" HorizontalAlignment="Left" Margin="62,271,0,0" VerticalAlignment="Top" IsEnabled="False" Visibility="Hidden" IsChecked="False"/>
        <CheckBox x:Name="chkbxMFA" Grid.Column="2" Content="" HorizontalAlignment="Left" Margin="62,334,0,0" VerticalAlignment="Top" IsEnabled="False" Visibility="Hidden" IsChecked="False"/>
        <Button x:Name="btnTroubleshooter" Content="Troubleshooter" HorizontalAlignment="Left" Margin="418,137,0,0" VerticalAlignment="Top" Height="30" Width="92" Background="White" BorderThickness="2,2,2,2" FontSize="12" Grid.Column="2" Grid.ColumnSpan="2"/>
    </Grid>
</Window>
"@

    # Replace some default attributes to support PowerShell's xml node reader
    [string]$XAML_Main = $XAML_Main -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
  
    # Convert to XML
    [xml]$UIXML_Main = $XAML_Main
   
    # Read the XML
    $reader_Main = (New-Object System.Xml.XmlNodeReader $UIXML_Main)

    # Load the xml reader into a form as a UI
try {
    $UI = [Windows.Markup.XamlReader]::Load($reader_Main)
} catch {
    Write-Error "Error loading XAML: $_"
    throw $_  # rethrow error after logging
}

    # Take the UI elements and make them variables (scope set to Script so variables can be called after function Run-MigrationWiz has run)
    $UIXML_Main.SelectNodes("//*[@Name]") | %{ Set-Variable -Scope Script -Name "$($_.Name)" -Value $UI.FindName($_.Name) }

    # =======================================
    # SCRIPT

    #Clickable Hyperlinks to Jira Guides

    $hyperlinkJiraOneDriveURL = "https://bbc.co.uk"
    $hyperlinkJiraOneDrive = New-Object System.Windows.Documents.Hyperlink
    $hyperlinkJiraOneDrive.NavigateUri = [System.Uri]::new($hyperlinkJiraOneDriveURL)
    $hyperlinkJiraOneDrive.Add_Click({
                    Start-Process -FilePath $hyperlinkJiraOneDriveURL
                })
    $hyperlinkJiraOneDrive.Inlines.Add("here")

    $hyperlinkJiraOutlookURL = "https://uk.yahoo.com/"
    $hyperlinkJiraOutlook = New-Object System.Windows.Documents.Hyperlink
    $hyperlinkJiraOutlook.NavigateUri = [System.Uri]::new($hyperlinkJiraOutlookURL)
    $hyperlinkJiraOutlook.Add_Click({
                    Start-Process -FilePath $hyperlinkJiraOutlookURL
                })
    $hyperlinkJiraOutlook.Inlines.Add("here")    

    $hyperlinkJiraMFAURL = "https://hotelplan.co.uk"
    $hyperlinkJiraMFA = New-Object System.Windows.Documents.Hyperlink
    $hyperlinkJiraMFA.NavigateUri = [System.Uri]::new($hyperlinkJiraMFAURL)
    $hyperlinkJiraMFA.Add_Click({
        Start-Process -FilePath $hyperlinkJiraMFAURL
    })
    $hyperlinkJiraMFA.Inlines.Add("here") 

    $hyperlinkMFASetupURL = "https://aka.ms/mfasetup"
    $hyperlinkMFASetup = New-Object System.Windows.Documents.Hyperlink
    $hyperlinkMFASetup.NavigateUri = [System.Uri]::new($hyperlinkMFASetupURL)

    $hyperlinkJiraGuidesURL = "https://hotelplanuk.atlassian.net/servicedesk/customer/portal/7/article/2397929474"
    $hyperlinkJiraGuides = New-Object System.Windows.Documents.Hyperlink
    $hyperlinkJiraGuides.NavigateUri = [System.Uri]::new($hyperlinkJiraGuidesURL)
   
    #User Lock file
    $UserLockPath = $env:LOCALAPPDATA
    $UserLockName = "user.migration.lock"
    $UserLockFile = "$UserLockPath\$UserLockName"
    
# =======================================
    # =======================================
    # EVENTS AND INTERACTIVE ELEMENTS

    # Disables the window X
    $UI.Add_Closing({$_.Cancel = $true})

    # Exit action for button
    $btnExit.Add_Click({
        $UI.Add_Closing({$_.Cancel = $false})
        $UI.Close()
    })

    # Set script states, states include default, migrateOneDrive, completeOneDrive, migrateOutlook, completed, troubleshooter.
    $script:currentstate = "default" # alt state is troubleshooting

    # Troubleshooter action for button 
    $btnTroubleshooter.Add_Click({
        Run-Troubleshooter
        
    })

    $btnGuides.Add_Click({
        $UI.Topmost = $false # Disables window from being on top
        Start-Process -FilePath $hyperlinkJiraGuidesURL
    })

    # Add the ability to drag windowsless screen around
    $eventHandler_LeftButtonDown = [Windows.Input.MouseButtonEventHandler] { $this.DragMove() }
    $UI.Add_MouseLeftButtonDown($eventHandler_LeftButtonDown)

    # Name TextBlocks
    $txtMFA.Text = "Multi-Factor Authentication"
    $txtOneDrive.Text = "Microsoft OneDrive"
    $txtOutlook.Text = "Microsoft Outlook"
        If (Test-Path -Path $UserLockPath\$UserLockName) { $btnNext.IsEnabled = $False
                                                            $txtDescription.Text = "Migration Wizard has already been completed for this user. `n`nClick Troubleshooter for more options or Exit to close."}
                                                            else { $txtDescription.Text = "Please click Next to start and follow the instructions in this window."}
    $txtOneDrive_Status.Text = "Not started"
    $txtOutlook_Status.Text = "Not started"
    $txtMFA_Status.Text = "Not started"
    $txtTitle.Text = "Welcome to the Hotelplan Migration Wiz"

    # Next button, switches states on click to run functions
        $btnNext.Add_Click({
        switch ($script:currentstate){
            "default" { $txtDescription.Text = "Click Next to start OneDrive migration."
                        $script:currentstate = "migrateOneDrive"
                        }
            "migrateOneDrive" { Configure-OneDrive
                                $script:currentstate = "completeOneDrive"
                                }
            "completeOneDrive" { Start-Process "C:\Program Files\Microsoft OneDrive\OneDrive.exe"
                                    $txtDescription.Text = "Click Next to start OneDrive migration."
                                    $txtOneDrive_Status.Text = "Completed"
                                    $imgRunOneDrive.Visibility = "Hidden"
                                    $imgCptOneDrive.Visibility = "Visible"
                                    $script:currentstate = "migrateOutlook"
                                }
            "migrateOutlook" { Configure-Outlook
                                $script:currentstate = "completeOutlook"
                                }
            "completeOutlook" { Start-Process outlook.exe
                                New-Item -Path $UserLockPath -Name $UserLockName -ItemType "file" -Force | Out-Null # Creates lock file to prevent from running again
                                $txtDescription.Text = "Outlook migration completed. `n`n Click Next to setup Multi-Factor Authentication (MFA)."
                                $txtOutlook_Status.Text = "Completed"
                                $imgRunOulook.Visibility = "Hidden"
                                $imgCptOutlook.Visibility = "Visible"
                                $script:currentstate = "migrateMFA"
                              }
            "migrateMFA"    {   Configure-MFA
                                $script:currentstate = "runMFA"
                            }
            "runMFA"        {   Start-Process -FilePath $hyperlinkMFASetupURL
                                $script:currentstate = "completeMFA"
                            }
            "completeMFA"   {   $txtDescription.Text = "Multi-Factor Authentication (MFA) setup completed. Please click Next to continue."
                                $txtMFA_Status.Text = "Completed"
                                $imgRunMFA.Visibility = "Hidden"
                                $imgCptMFA.Visibility = "Visible"
                                $script:currentstate = "completed"
                            }
            "completed" { 
                            $txtDescription.Text = "Migration Wiz completed. Please click Exit to close."
                            $btnNext.IsEnabled = $false
                            }
            "troubleshooter" { 
                            }
            "troubleshootermigrateOneDrive" {Configure-OneDrive
                                            $script:currentstate = "troubleshootercompleteOneDrive"}

            "troubleshootercompleteOneDrive" { Start-Process "C:\Program Files\Microsoft OneDrive\OneDrive.exe"
                                    $txtDescription.Text = "Migration completed. `n`nPlease click Exit to close or Troubleshooter for more options."
                                    $txtOneDrive_Status.Text = "Completed"
                                    $imgRunOneDrive.Visibility = "Hidden"
                                    $imgCptOneDrive.Visibility = "Visible"
                                    $btnNext.IsEnabled = $false
                                            }
            "troubleshootermigrateOutlook" {Configure-Outlook
                                            $script:currentstate = "troubleshootercompleteOutlook"}

            "troubleshootercompleteOutlook" { Start-Process outlook.exe
                                    $txtDescription.Text = "Migration completed. `n`nPlease click Exit to close or Troubleshooter for more options."
                            $txtOutlook_Status.Text = "Completed"
                            $imgRunOutlook.Visibility = "Hidden"
                            $imgCptOutlook.Visibility = "Visible"
                            $btnNext.IsEnabled = $false
                                            }
                }

    })

    # Function to add Troubleshooter options and add alternative button switches. 
    function Run-Troubleshooter {
        $script:currentstate = "troubleshooter"

        # Adds check boxes, update description.
        $txtDescription.Text = "Please tick the app you wish to configure and click Next to start."
        $chkbxOneDrive.Visibility = "Visible"
        $chkbxOneDrive.IsEnabled = $True
        $chkbxOneDrive.IsChecked = $False
        $chkbxOutlook.Visibility = "Visible"
        $chkbxOutlook.IsEnabled = $True
        $chkbxOutlook.IsChecked = $False
        $chkbxMFA.Visibility = "Visible"
        $chkbxMFA.IsEnabled = $True
        $chkbxMFA.IsChecked = $False
        
        #Reset status
        $imgNSOneDrive.Visibility = "Visible"
        $imgNSOutlook.Visibility = "Visible"
        $imgNSMFA.Visibility = "Visible"
        $imgRunOneDrive.Visibility = "Hidden"
        $imgRunOutlook.Visibility = "Hidden"
        $imgRunMFA.Visibility = "Hidden"
        $imgCptOneDrive.Visibility = "Hidden"
        $imgCptOutlook.Visibility = "Hidden"
        $imgCptMFA.Visibility = "Hidden"
        $txtOneDrive_Status.Text = "Not started"
        $txtOutlook_Status.Text = "Not started"
        $txtMFA_Status.Text = "Not started"
        
        # Disables unselected boxes when one is ticked, updates script state.

        $chkbxOneDrive.Add_Checked({
            $btnNext.IsEnabled = $True
            $txtDescription.Text = "OneDrive selected. Please click Next to start and follow the instructions in this window."
            $chkbxOutlook.IsEnabled = $False
            $chkbxMFA.IsEnabled = $False
            $script:currentstate = "troubleshootermigrateOneDrive"
            })

        $chkbxOutlook.Add_Checked({
            $btnNext.IsEnabled = $True
            $txtDescription.Text = "Outlook selected. Please click Next to start and follow the instructions in this window."
            $chkbxOneDrive.IsEnabled = $False
            $chkbxMFA.IsEnabled = $False
            $script:currentstate = "troubleshootermigrateOutlook"
            })

        $chkbxMFA.Add_Checked({
            $btnNext.IsEnabled = $True
            $txtDescription.Text = "Multi-Factor Authentication selected. Please click Next to start and follow the instructions in this window."
            $chkbxMFA.IsEnabled = $False
            $chkbxMFA.IsEnabled = $False
            $script:currentstate = "migrateMFA"
            })

        $chkbxOneDrive.Add_Unchecked({
            $btnNext.IsEnabled = $False
            Run-Troubleshooter
            })

        $chkbxOutlook.Add_Unchecked({
            $btnNext.IsEnabled = $False
            Run-Troubleshooter
            })

        $chkbxMFA.Add_Unchecked({
            $btnNext.IsEnabled = $False
            Run-Troubleshooter
            })
    }
    
    # PRESENT UI
    $UI.ShowDialog()  
}


# Call the function
Run-MigrationWiz
