<#
.SYNOPSIS
Interactively select and remove Appx packages, or supply a list file for automation.

.DESCRIPTION
Interactive (default):
  Displays all installed and provisioned Appx packages in a checkbox GUI.
  Packages are color-coded: green (bloat), yellow (optional), red (essential), grey (protected).
  Installed packages are removed for the current user by default. Check the
  all-users toggle to remove for every user profile on the machine.
  Provisioned removal prevents the package from auto-installing for new accounts.

Automated:
  .\AppxManager.ps1 -ListFile .\targets.txt

Compile to exe (double-clickable, auto-elevates):
  Install-Module ps2exe -Scope CurrentUser
  Invoke-PS2EXE -InputFile .\AppxManager.ps1 -OutputFile .\AppxManager.exe -RequireAdmin

.PARAMETER ListFile
  Optional path to a text file with one target per line for unattended removal.

.PARAMETER WhatIf
  Preview what would be removed without actually removing anything.

.PARAMETER LogFile
  Override the default log path (Desktop\AppxRemove.log).
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$ListFile,
    [string]$LogFile = (Join-Path $env:USERPROFILE 'Desktop\AppxRemove.log')
)

# -- Self-elevate if not admin -------------------------------------------------
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',$MyInvocation.MyCommand.Path -Verb RunAs
    exit
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -- Package classification ----------------------------------------------------
$script:Classification = @{
    'Clipchamp.Clipchamp'                        = 'Bloat'
    'Microsoft.549981C3F5F10'                     = 'Bloat'
    'Microsoft.BingNews'                          = 'Bloat'
    'Microsoft.BingWeather'                       = 'Bloat'
    'Microsoft.BingFinance'                       = 'Bloat'
    'Microsoft.BingSports'                        = 'Bloat'
    'Microsoft.BingTranslator'                    = 'Bloat'
    'Microsoft.BingFoodAndDrink'                  = 'Bloat'
    'Microsoft.BingHealthAndFitness'              = 'Bloat'
    'Microsoft.BingTravel'                        = 'Bloat'
    'Microsoft.GamingApp'                         = 'Bloat'
    'Microsoft.GetHelp'                           = 'Bloat'
    'Microsoft.Getstarted'                        = 'Bloat'
    'Microsoft.Microsoft3DViewer'                 = 'Bloat'
    'Microsoft.MicrosoftOfficeHub'                = 'Bloat'
    'Microsoft.MicrosoftSolitaireCollection'      = 'Bloat'
    'Microsoft.MicrosoftStickyNotes'              = 'Bloat'
    'Microsoft.MixedReality.Portal'               = 'Bloat'
    'Microsoft.MSPaint'                           = 'Bloat'
    'Microsoft.OneConnect'                        = 'Bloat'
    'Microsoft.People'                            = 'Bloat'
    'Microsoft.PowerAutomateDesktop'              = 'Bloat'
    'Microsoft.Print3D'                           = 'Bloat'
    'Microsoft.SkypeApp'                          = 'Bloat'
    'Microsoft.Todos'                             = 'Bloat'
    'Microsoft.WindowsCommunicationsApps'         = 'Bloat'
    'Microsoft.WindowsFeedbackHub'                = 'Bloat'
    'Microsoft.WindowsMaps'                       = 'Bloat'
    'Microsoft.Xbox.TCUI'                         = 'Bloat'
    'Microsoft.XboxApp'                           = 'Bloat'
    'Microsoft.XboxGameOverlay'                   = 'Bloat'
    'Microsoft.XboxGamingOverlay'                 = 'Bloat'
    'Microsoft.XboxSpeechToTextOverlay'           = 'Bloat'
    'Microsoft.YourPhone'                         = 'Bloat'
    'Microsoft.ZuneMusic'                         = 'Bloat'
    'Microsoft.ZuneVideo'                         = 'Bloat'
    'MicrosoftCorporationII.QuickAssist'          = 'Bloat'
    'MicrosoftTeams'                              = 'Bloat'
    'Microsoft.OutlookForWindows'                 = 'Bloat'
    'Microsoft.MicrosoftJournal'                  = 'Bloat'
    'Microsoft.Whiteboard'                        = 'Bloat'
    'Microsoft.RemoteDesktop'                     = 'Bloat'
    'Microsoft.NetworkSpeedTest'                  = 'Bloat'
    'Microsoft.WindowsAlarms'                     = 'Bloat'
    'Microsoft.WindowsSoundRecorder'              = 'Bloat'
    'Microsoft.StartExperiencesApp'               = 'Bloat'
    'Microsoft.LinkedInforWindows'                = 'Bloat'
    'SpotifyAB.SpotifyMusic'                      = 'Bloat'
    'Disney.37853FC22B2CE'                        = 'Bloat'
    'BytedancePte.Ltd.TikTok'                     = 'Bloat'
    'Amazon.com.Amazon'                           = 'Bloat'
    'Facebook.Facebook'                           = 'Bloat'
    'Facebook.Instagram'                          = 'Bloat'
    'Facebook.FacebookMessenger'                  = 'Bloat'
    'king.com.CandyCrushSaga'                     = 'Bloat'
    'king.com.CandyCrushSodaSaga'                 = 'Bloat'
    'AmazonVideo.PrimeVideo'                      = 'Bloat'
    'Microsoft.ScreenSketch'                      = 'Optional'
    'Microsoft.Windows.Photos'                    = 'Optional'
    'Microsoft.WindowsCalculator'                 = 'Optional'
    'Microsoft.WindowsCamera'                     = 'Optional'
    'Microsoft.WindowsNotepad'                    = 'Optional'
    'Microsoft.WindowsTerminal'                   = 'Optional'
    'Microsoft.WindowsStore'                      = 'Optional'
    'Microsoft.StorePurchaseApp'                  = 'Optional'
    'Microsoft.Paint'                             = 'Optional'
    'Microsoft.MicrosoftEdge.Stable'              = 'Optional'
    'Microsoft.OneDrive'                          = 'Optional'
    'Microsoft.Office.OneNote'                    = 'Optional'
    'Microsoft.WindowsTerminalPreview'            = 'Optional'
    'Microsoft.DevHome'                           = 'Optional'
    'Microsoft.Copilot'                           = 'Optional'
    'MicrosoftWindows.CrossDevice'                = 'Optional'
    'Microsoft.DesktopAppInstaller'               = 'Essential'
    'Microsoft.WindowsAppRuntime.*'               = 'Essential'
    'Microsoft.VCLibs.*'                          = 'Essential'
    'Microsoft.UI.Xaml.*'                         = 'Essential'
    'Microsoft.NET.Native.*'                      = 'Essential'
    'Microsoft.Services.Store.Engagement'         = 'Essential'
    'Microsoft.HEIFImageExtension'                = 'Essential'
    'Microsoft.HEVCVideoExtension'                = 'Essential'
    'Microsoft.VP9VideoExtensions'                = 'Essential'
    'Microsoft.WebMediaExtensions'                = 'Essential'
    'Microsoft.WebpImageExtension'                = 'Essential'
    'Microsoft.RawImageExtension'                 = 'Essential'
    'Microsoft.AV1VideoExtension'                 = 'Essential'
    'Microsoft.MPEG2VideoExtension'               = 'Essential'
    'Microsoft.XboxIdentityProvider'              = 'Essential'
    'Microsoft.WindowsAppRuntime.Main'            = 'Essential'
    'Microsoft.AAD.BrokerPlugin'                  = 'Protected'
    'Microsoft.AccountsControl'                   = 'Protected'
    'Microsoft.Windows.CloudExperienceHost'       = 'Protected'
    'Microsoft.Windows.ContentDeliveryManager'    = 'Protected'
    'Microsoft.Windows.OOBENetworkCaptivePortal'  = 'Protected'
    'Microsoft.Windows.OOBENetworkConnectionFlow' = 'Protected'
    'Microsoft.Windows.PeopleExperienceHost'      = 'Protected'
    'Microsoft.Windows.SecHealthUI'               = 'Protected'
    'Microsoft.Windows.ShellExperienceHost'       = 'Protected'
    'Microsoft.Windows.StartMenuExperienceHost'   = 'Protected'
    'Microsoft.Windows.XGpuEjectDialog'           = 'Protected'
    'Microsoft.Windows.ParentalControls'          = 'Protected'
    'Microsoft.Windows.Search'                    = 'Protected'
    'Microsoft.Windows.Apprep.ChxApp'             = 'Protected'
    'Microsoft.Windows.AssignedAccessLockApp'     = 'Protected'
    'Microsoft.Windows.CapturePicker'             = 'Protected'
    'Microsoft.Windows.NarratorQuickStart'        = 'Protected'
    'Microsoft.Windows.PrintQueueActionCenter'    = 'Protected'
    'Microsoft.BioEnrollment'                     = 'Protected'
    'Microsoft.LockApp'                           = 'Protected'
    'Microsoft.MicrosoftEdge'                     = 'Protected'
    'Microsoft.Win32WebViewHost'                  = 'Protected'
    'MicrosoftWindows.Client.Core'                = 'Protected'
    'MicrosoftWindows.Client.CBS'                 = 'Protected'
    'MicrosoftWindows.UndockedDevKit'             = 'Protected'
    'windows.immersivecontrolpanel'               = 'Protected'
    'Windows.PrintDialog'                         = 'Protected'
    'NcsiUwpApp'                                  = 'Protected'
}

function Get-Tier {
    param([string]$Name)
    if ($script:Classification.ContainsKey($Name)) {
        return $script:Classification[$Name]
    }
    foreach ($key in $script:Classification.Keys) {
        if ($key.Contains('*') -and $Name -like $key) {
            return $script:Classification[$key]
        }
    }
    if ($Name -match '\.NET\.' -or $Name -match 'VCLibs' -or $Name -match 'UI\.Xaml' -or $Name -match 'AppRuntime') {
        return 'Essential'
    }
    if ($Name -match '^Microsoft\.Windows\.' -or $Name -match '^MicrosoftWindows\.') {
        return 'Protected'
    }
    return 'Optional'
}

# -- Helpers -------------------------------------------------------------------
function Log {
    param(
        [string]$Msg,
        [ValidateSet('INFO','WARN','ERROR')][string]$Lvl = 'INFO'
    )
    $colors = @{ INFO = 'Gray'; WARN = 'Yellow'; ERROR = 'Red' }
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format s), $Lvl, $Msg
    Write-Host $line -ForegroundColor $colors[$Lvl]
    Add-Content -LiteralPath $LogFile -Value $line
}

function Parse-ListFile {
    param([string]$Path)
    (Get-Content -LiteralPath $Path) |
        ForEach-Object { ($_ -replace '\s*#.*$','').Trim() } |
        Where-Object { $_ -ne '' }
}

function Resolve-Pattern {
    param([string]$Name)
    if ($Name -match '[\*\?]') { return $Name }
    return ($Name + '*')
}

function Remove-Installed {
    param([array]$Packages, [bool]$AllUsers = $false)
    foreach ($pkg in $Packages) {
        $scope = if ($AllUsers) { 'all users' } else { 'current user' }
        $label = '{0} ({1}) [{2}]' -f $pkg.Name, $pkg.PackageFullName, $scope
        if ($PSCmdlet.ShouldProcess($label, 'Remove installed package')) {
            try {
                if ($AllUsers) {
                    Remove-AppxPackage -Package $pkg.PackageFullName -AllUsers -ErrorAction Stop
                }
                else {
                    Remove-AppxPackage -Package $pkg.PackageFullName -ErrorAction Stop
                }
                Log ('Removed: ' + $label)
            }
            catch {
                Log ('FAILED: ' + $label + ' -- ' + $_.Exception.Message) 'ERROR'
            }
        }
    }
}

function Remove-Provisioned {
    param([array]$Packages)
    foreach ($pp in $Packages) {
        $label = '{0} ({1})' -f $pp.DisplayName, $pp.PackageName
        if ($PSCmdlet.ShouldProcess($label, 'Remove provisioned package')) {
            try {
                Remove-AppxProvisionedPackage -Online -PackageName $pp.PackageName -ErrorAction Stop | Out-Null
                Log ('De-provisioned: ' + $label)
            }
            catch {
                if ($_.Exception.Message -match 'cannot find the path') {
                    Log ('Already gone (staging path missing): ' + $label) 'WARN'
                }
                else {
                    Log ('FAILED: ' + $label + ' -- ' + $_.Exception.Message) 'ERROR'
                }
            }
        }
    }
}

# -- WPF Checkbox GUI ----------------------------------------------------------
function Show-PackageSelector {
    param([array]$PackageList)

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    [xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Appx Package Manager" Height="720" Width="800"
    WindowStartupLocation="CenterScreen"
    Background="#1e1e1e">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Select packages to remove"
                   Foreground="#cccccc" FontSize="18" FontWeight="SemiBold"
                   Margin="0,0,0,4"/>

        <StackPanel Grid.Row="1" Margin="0,0,0,8">
            <StackPanel Orientation="Horizontal" Margin="0,2,0,2">
                <Ellipse Width="10" Height="10" Fill="#4ec969" Margin="0,0,6,0" VerticalAlignment="Center"/>
                <TextBlock Foreground="#aaa" FontSize="11" VerticalAlignment="Center">
                    <Run FontWeight="SemiBold">Bloat</Run>
                    <Run Foreground="#888"> - safe to remove, no OS impact</Run>
                </TextBlock>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2,0,2">
                <Ellipse Width="10" Height="10" Fill="#e8b634" Margin="0,0,6,0" VerticalAlignment="Center"/>
                <TextBlock Foreground="#aaa" FontSize="11" VerticalAlignment="Center">
                    <Run FontWeight="SemiBold">Optional</Run>
                    <Run Foreground="#888"> - useful but not critical; remove if you don't need it</Run>
                </TextBlock>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2,0,2">
                <Ellipse Width="10" Height="10" Fill="#e05454" Margin="0,0,6,0" VerticalAlignment="Center"/>
                <TextBlock Foreground="#aaa" FontSize="11" VerticalAlignment="Center">
                    <Run FontWeight="SemiBold">Essential</Run>
                    <Run Foreground="#888"> - removing may break the OS or other apps</Run>
                </TextBlock>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2,0,2">
                <Ellipse Width="10" Height="10" Fill="#777777" Margin="0,0,6,0" VerticalAlignment="Center"/>
                <TextBlock Foreground="#aaa" FontSize="11" VerticalAlignment="Center">
                    <Run FontWeight="SemiBold">Protected</Run>
                    <Run Foreground="#888"> - Windows blocks removal of these</Run>
                </TextBlock>
            </StackPanel>
            <TextBlock Foreground="#666" FontSize="11" FontStyle="Italic" TextWrapping="Wrap" Margin="0,4,0,0">
                Installed = currently installed on this PC. Removed for the current user by default, or for every user account if the toggle below is checked.
            </TextBlock>
            <TextBlock Foreground="#666" FontSize="11" FontStyle="Italic" TextWrapping="Wrap" Margin="0,2,0,0">
                Provisioned = staged in the system image. Removing prevents the package from auto-installing for new user accounts. Not affected by the all-users toggle.
            </TextBlock>
        </StackPanel>

        <Grid Grid.Row="2" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox Grid.Column="0" Name="SearchBox"
                     Background="#2d2d2d" Foreground="#cccccc" BorderBrush="#444"
                     Padding="6,4" FontSize="13" VerticalContentAlignment="Center"/>
            <Button Grid.Column="1" Name="FilterBloatBtn" Content="Bloat only"
                    Margin="8,0,0,0" Padding="10,4"
                    Background="#333" Foreground="#cccccc" BorderBrush="#555"
                    FontSize="12"/>
            <Button Grid.Column="2" Name="FilterAllBtn" Content="Show all"
                    Margin="8,0,0,0" Padding="10,4"
                    Background="#333" Foreground="#cccccc" BorderBrush="#555"
                    FontSize="12"/>
            <Button Grid.Column="3" Name="DeselectAllBtn" Content="Deselect All"
                    Margin="8,0,0,0" Padding="10,4"
                    Background="#333" Foreground="#cccccc" BorderBrush="#555"
                    FontSize="12"/>
        </Grid>

        <CheckBox Grid.Row="3" Name="AllUsersCheck" Margin="0,0,0,8"
                  Foreground="#cccccc" FontSize="12" IsChecked="False">
            <TextBlock Foreground="#cccccc" FontSize="12">
                <Run>Remove installed packages for ALL users on this PC</Run>
                <Run Foreground="#888"> (default: current user only)</Run>
            </TextBlock>
        </CheckBox>

        <ListView Grid.Row="4" Name="PackageListView"
                  Background="#252526" BorderBrush="#444" Foreground="#cccccc"
                  ScrollViewer.HorizontalScrollBarVisibility="Disabled">
            <ListView.ItemTemplate>
                <DataTemplate>
                    <Grid Margin="2,2,2,2">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <CheckBox Grid.Column="0" IsChecked="{Binding IsChecked, Mode=TwoWay}"
                                  Margin="0,0,6,0" VerticalAlignment="Center"/>
                        <Ellipse Grid.Column="1" Width="10" Height="10"
                                 Fill="{Binding TierColor}" Margin="0,0,8,0"
                                 VerticalAlignment="Center"/>
                        <TextBlock Grid.Column="2" Text="{Binding Name}" FontSize="13"
                                   Foreground="#cccccc" VerticalAlignment="Center"/>
                        <TextBlock Grid.Column="3" VerticalAlignment="Center" Margin="8,0,0,0">
                            <Run Text="{Binding Tier, Mode=OneWay}" FontSize="11"
                                 Foreground="{Binding TierColor}"/>
                            <Run Text=" | " Foreground="#444" FontSize="11"/>
                            <Run Text="{Binding Source, Mode=OneWay}" FontSize="11"
                                 Foreground="#888"/>
                        </TextBlock>
                    </Grid>
                </DataTemplate>
            </ListView.ItemTemplate>
        </ListView>

        <TextBlock Grid.Row="5" Name="StatusText"
                   Foreground="#888" FontSize="12" Margin="0,8,0,4"
                   Text="0 selected"/>

        <StackPanel Grid.Row="6" Orientation="Horizontal"
                    HorizontalAlignment="Right" Margin="0,4,0,0">
            <Button Name="CancelBtn" Content="Cancel" Width="90" Height="30"
                    Margin="0,0,8,0"
                    Background="#333" Foreground="#cccccc" BorderBrush="#555"
                    FontSize="13"/>
            <Button Name="RemoveBtn" Content="Remove Selected" Width="130" Height="30"
                    Background="#c42b1c" Foreground="White" BorderBrush="#a02010"
                    FontSize="13" FontWeight="SemiBold"/>
        </StackPanel>
    </Grid>
</Window>
'@

    $reader = [System.Xml.XmlNodeReader]::new($xaml)
    $window = [System.Windows.Markup.XamlReader]::Load($reader)

    $searchBox     = $window.FindName('SearchBox')
    $listView      = $window.FindName('PackageListView')
    $statusText    = $window.FindName('StatusText')
    $removeBtn     = $window.FindName('RemoveBtn')
    $cancelBtn     = $window.FindName('CancelBtn')
    $deselectAll   = $window.FindName('DeselectAllBtn')
    $filterBloat   = $window.FindName('FilterBloatBtn')
    $filterAll     = $window.FindName('FilterAllBtn')
    $allUsersCheck = $window.FindName('AllUsersCheck')

    $tierOrder  = @{ Bloat = 0; Optional = 1; Essential = 2; Protected = 3 }
    $tierColors = @{ Bloat = '#4ec969'; Optional = '#e8b634'; Essential = '#e05454'; Protected = '#777777' }

    $allItems = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
    $sorted = $PackageList | Sort-Object {
        $t = Get-Tier -Name $_.Name
        $tierOrder[$t]
    }, Name

    foreach ($p in $sorted) {
        $tier = Get-Tier -Name $p.Name
        $item = [PSCustomObject]@{
            IsChecked = $false
            Name      = $p.Name
            Source    = $p.Source
            Id        = $p.Id
            Tier      = $tier
            TierColor = $tierColors[$tier]
        }
        $allItems.Add($item)
    }

    $listView.ItemsSource = $allItems

    $script:showBloatOnly = $false

    $applyFilter = {
        $term = $searchBox.Text.Trim().ToLower()
        $filtered = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
        foreach ($item in $allItems) {
            $nameMatch = ($term -eq '') -or ($item.Name.ToLower().Contains($term))
            $tierMatch = (-not $script:showBloatOnly) -or ($item.Tier -eq 'Bloat')
            if ($nameMatch -and $tierMatch) {
                $filtered.Add($item)
            }
        }
        $listView.ItemsSource = $filtered
    }

    $updateStatus = {
        $count = @($allItems | Where-Object { $_.IsChecked }).Count
        $statusText.Text = ('{0} selected' -f $count)
    }

    $searchBox.Add_TextChanged({
        & $applyFilter
        & $updateStatus
    })

    $filterBloat.Add_Click({
        $script:showBloatOnly = $true
        & $applyFilter
        & $updateStatus
    })

    $filterAll.Add_Click({
        $script:showBloatOnly = $false
        & $applyFilter
        & $updateStatus
    })

    $deselectAll.Add_Click({
        foreach ($item in $listView.ItemsSource) { $item.IsChecked = $false }
        $listView.Items.Refresh()
        & $updateStatus
    })

    $listView.Add_PreviewMouseUp({ & $updateStatus })
    $listView.Add_KeyUp({ & $updateStatus })

    $cancelBtn.Add_Click({ $window.DialogResult = $false; $window.Close() })

    $removeBtn.Add_Click({
        $checked = @($allItems | Where-Object { $_.IsChecked }).Count
        if ($checked -eq 0) {
            [System.Windows.MessageBox]::Show(
                'No packages selected.',
                'Nothing to do',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
            return
        }

        $hasEssential = @($allItems | Where-Object { $_.IsChecked -and $_.Tier -eq 'Essential' }).Count
        $hasProtected = @($allItems | Where-Object { $_.IsChecked -and $_.Tier -eq 'Protected' }).Count
        $warning = 'Remove ' + $checked + ' package(s)? This cannot be undone.'
        if ($hasProtected -gt 0) {
            $nl = [Environment]::NewLine
            $warning = 'NOTE: ' + $hasProtected + ' protected package(s) selected. Windows will likely block their removal.' + $nl + $nl + $warning
        }
        if ($hasEssential -gt 0) {
            $nl = [Environment]::NewLine
            $warning = 'WARNING: ' + $hasEssential + ' essential package(s) selected! Removing these may break your system.' + $nl + $nl + $warning
        }

        $answer = [System.Windows.MessageBox]::Show(
            $warning,
            'Confirm Removal',
            [System.Windows.MessageBoxButton]::YesNo,
            [System.Windows.MessageBoxImage]::Warning
        )

        if ($answer -eq [System.Windows.MessageBoxResult]::Yes) {
            $window.DialogResult = $true
            $window.Close()
        }
    })

    $result = $window.ShowDialog()

    if ($result -eq $true) {
        return [PSCustomObject]@{
            Selected = ($allItems | Where-Object { $_.IsChecked })
            AllUsers = $allUsersCheck.IsChecked
        }
    }
    return $null
}

# -- Collect packages ----------------------------------------------------------
$installed   = Get-AppxPackage -AllUsers
$provisioned = Get-AppxProvisionedPackage -Online

# -- Interactive mode (default) ------------------------------------------------
if (-not $ListFile) {

    $gridInstalled = $installed | Select-Object `
        @{N='Name';E={$_.Name}}, `
        @{N='Source';E={'Installed'}}, `
        @{N='Id';E={$_.PackageFullName}}

    $gridProvisioned = $provisioned | Select-Object `
        @{N='Name';E={$_.DisplayName}}, `
        @{N='Source';E={'Provisioned'}}, `
        @{N='Id';E={$_.PackageName}}

    $grid = @($gridInstalled) + @($gridProvisioned) | Sort-Object Name, Source -Unique

    $result = Show-PackageSelector -PackageList $grid

    if (-not $result) {
        Write-Host 'Nothing selected - exiting.' -ForegroundColor Yellow
        exit 0
    }

    $selected = @($result.Selected)
    $removeAllUsers = $result.AllUsers

    if ($selected.Count -eq 0) {
        Write-Host 'Nothing selected - exiting.' -ForegroundColor Yellow
        exit 0
    }

    $selInstalled   = @($selected | Where-Object { $_.Source -eq 'Installed' })
    $selProvisioned = @($selected | Where-Object { $_.Source -eq 'Provisioned' })

    if ($selInstalled.Count -gt 0) {
        $pkgs = @($installed | Where-Object { $_.PackageFullName -in @($selInstalled | ForEach-Object { $_.Id }) })
        Remove-Installed -Packages $pkgs -AllUsers $removeAllUsers
    }
    if ($selProvisioned.Count -gt 0) {
        $provNow = Get-AppxProvisionedPackage -Online
        $provIds = @($selProvisioned | ForEach-Object { $_.Id })
        $pkgs = @($provNow | Where-Object { $_.PackageName -in $provIds })

        $foundIds = @($pkgs | ForEach-Object { $_.PackageName })
        foreach ($id in $provIds) {
            if ($id -notin $foundIds) {
                Log ('Provisioned entry no longer present (may have been missing before run): ' + $id) 'WARN'
            }
        }

        if ($pkgs.Count -gt 0) {
            Remove-Provisioned -Packages $pkgs
        }
    }

    Log 'Interactive removal complete.'
    exit 0
}

# -- List-file mode (automation) -----------------------------------------------
if (-not (Test-Path -LiteralPath $ListFile)) {
    Write-Error ('List file not found: ' + $ListFile)
    exit 2
}

$targets = @(Parse-ListFile -Path $ListFile)
if ($targets.Count -eq 0) {
    Log 'List file is empty.' 'WARN'
    exit 0
}

Log ('List-file mode -- targets: ' + ($targets -join ', '))

foreach ($t in $targets) {
    $pat = Resolve-Pattern -Name $t

    $hitI = @($installed   | Where-Object { $_.Name -like $pat })
    $hitP = @($provisioned | Where-Object { $_.DisplayName -like $pat })

    if ($hitI.Count -eq 0 -and $hitP.Count -eq 0) {
        Log ('No matches for: ' + $t) 'WARN'
        continue
    }

    if ($hitI.Count -gt 0) {
        $deduped = $hitI | Group-Object PackageFullName | ForEach-Object { $_.Group[0] }
        Remove-Installed -Packages $deduped -AllUsers $true
    }
    if ($hitP.Count -gt 0) {
        Remove-Provisioned -Packages $hitP
    }
}

Log 'List-file removal complete.'
exit 0
