<#
.SYNOPSIS
Interactively select and remove local user profiles via a checkbox GUI.

.DESCRIPTION
Displays all user profiles on the machine in a WPF GUI with checkboxes.
Each profile shows the username, profile size, and last login date.
The currently logged-in user is always protected and cannot be selected.
Built-in system profiles (Default, systemprofile, etc.) are also protected.

Self-elevates to admin if needed via UAC prompt.

Compile to exe (double-clickable, auto-elevates):
  Install-Module ps2exe -Scope CurrentUser
  Invoke-PS2EXE -InputFile .\ProfileManager.ps1 -OutputFile .\ProfileManager.exe -RequireAdmin

.PARAMETER LogFile
  Override the default log path (Desktop\ProfileRemove.log).
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$LogFile = (Join-Path $env:USERPROFILE 'Desktop\ProfileRemove.log')
)

# -- Self-elevate if not admin -------------------------------------------------
$principal = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Start-Process powershell -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',$MyInvocation.MyCommand.Path -Verb RunAs
    exit
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

function Get-FolderSizeMB {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return 0 }
    try {
        $bytes = (Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum
        if ($null -eq $bytes) { return 0 }
        return [math]::Round($bytes / 1MB, 1)
    }
    catch { return 0 }
}

function Format-Size {
    param([double]$MB)
    if ($MB -ge 1024) {
        return ('{0:N1} GB' -f ($MB / 1024))
    }
    return ('{0:N0} MB' -f $MB)
}

# -- Collect profiles ----------------------------------------------------------
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent().User.Value
$systemPaths = @('systemprofile', 'LocalService', 'NetworkService', 'Default', 'Default User', 'Public')

$profiles = Get-CimInstance -Class Win32_UserProfile | Where-Object {
    -not $_.Special -and $_.LocalPath -and $_.LocalPath -ne ''
}

$profileData = foreach ($p in $profiles) {
    $username = Split-Path -Leaf $p.LocalPath
    $isSystem = $username -in $systemPaths
    $isCurrent = $p.SID -eq $currentUser

    $lastUse = $p.LastUseTime
    $lastUseStr = if ($null -ne $lastUse) { $lastUse.ToString('yyyy-MM-dd HH:mm') } else { 'Unknown' }

    $tier = 'User'
    if ($isCurrent) { $tier = 'Current' }
    elseif ($isSystem) { $tier = 'System' }

    [PSCustomObject]@{
        Username  = $username
        Path      = $p.LocalPath
        SID       = $p.SID
        LastUsed  = $lastUseStr
        Size      = 'Calculating...'
        SizeMB    = -1
        Tier      = $tier
        Protected = ($isCurrent -or $isSystem)
    }
}

$profileData = @($profileData | Sort-Object @{E={
    switch ($_.Tier) { 'User' { 0 } 'Current' { 1 } 'System' { 2 } }
}}, Username)

# -- WPF GUI -------------------------------------------------------------------
function Show-ProfileSelector {
    param([array]$Profiles)

    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase

    [xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="User Profile Manager" Height="580" Width="750"
    WindowStartupLocation="CenterScreen"
    Background="#1e1e1e">
    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBlock Grid.Row="0" Text="Select user profiles to remove"
                   Foreground="#cccccc" FontSize="18" FontWeight="SemiBold"
                   Margin="0,0,0,4"/>

        <StackPanel Grid.Row="1" Margin="0,0,0,8">
            <StackPanel Orientation="Horizontal" Margin="0,2,0,2">
                <Ellipse Width="10" Height="10" Fill="#4ec969" Margin="0,0,6,0" VerticalAlignment="Center"/>
                <TextBlock Foreground="#aaa" FontSize="11" VerticalAlignment="Center">
                    <Run FontWeight="SemiBold">User</Run>
                    <Run Foreground="#888"> - removable user profile</Run>
                </TextBlock>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2,0,2">
                <Ellipse Width="10" Height="10" Fill="#5b9bd5" Margin="0,0,6,0" VerticalAlignment="Center"/>
                <TextBlock Foreground="#aaa" FontSize="11" VerticalAlignment="Center">
                    <Run FontWeight="SemiBold">Current</Run>
                    <Run Foreground="#888"> - you are logged in as this user (protected)</Run>
                </TextBlock>
            </StackPanel>
            <StackPanel Orientation="Horizontal" Margin="0,2,0,2">
                <Ellipse Width="10" Height="10" Fill="#777777" Margin="0,0,6,0" VerticalAlignment="Center"/>
                <TextBlock Foreground="#aaa" FontSize="11" VerticalAlignment="Center">
                    <Run FontWeight="SemiBold">System</Run>
                    <Run Foreground="#888"> - built-in system profile (protected)</Run>
                </TextBlock>
            </StackPanel>
            <TextBlock Foreground="#666" FontSize="11" FontStyle="Italic" TextWrapping="Wrap" Margin="0,4,0,0">
                Removing a profile deletes the user folder and all its contents (documents, desktop, downloads, etc.). This cannot be undone.
            </TextBlock>
        </StackPanel>

        <Grid Grid.Row="2" Margin="0,0,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBox Grid.Column="0" Name="SearchBox"
                     Background="#2d2d2d" Foreground="#cccccc" BorderBrush="#444"
                     Padding="6,4" FontSize="13" VerticalContentAlignment="Center"/>
            <Button Grid.Column="1" Name="DeselectAllBtn" Content="Deselect All"
                    Margin="8,0,0,0" Padding="10,4"
                    Background="#333" Foreground="#cccccc" BorderBrush="#555"
                    FontSize="12"/>
        </Grid>

        <ListView Grid.Row="3" Name="ProfileListView"
                  Background="#252526" BorderBrush="#444" Foreground="#cccccc"
                  ScrollViewer.HorizontalScrollBarVisibility="Disabled">
            <ListView.ItemTemplate>
                <DataTemplate>
                    <Grid Margin="2,3,2,3">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <CheckBox Grid.Column="0"
                                  IsChecked="{Binding IsChecked, Mode=TwoWay}"
                                  IsEnabled="{Binding CanSelect}"
                                  Margin="0,0,6,0" VerticalAlignment="Center"/>
                        <Ellipse Grid.Column="1" Width="10" Height="10"
                                 Fill="{Binding TierColor}" Margin="0,0,8,0"
                                 VerticalAlignment="Center"/>
                        <TextBlock Grid.Column="2" Text="{Binding Username}" FontSize="13"
                                   Foreground="{Binding TextColor}" VerticalAlignment="Center"/>
                        <TextBlock Grid.Column="3" Text="{Binding Size}" FontSize="11"
                                   Foreground="#888" VerticalAlignment="Center" Margin="12,0,0,0"/>
                        <TextBlock Grid.Column="4" Text="{Binding LastUsed}" FontSize="11"
                                   Foreground="#888" VerticalAlignment="Center" Margin="12,0,0,0"/>
                        <TextBlock Grid.Column="5" Text="{Binding Tier}" FontSize="11"
                                   Foreground="{Binding TierColor}" VerticalAlignment="Center"
                                   Margin="12,0,0,0" Width="55"/>
                    </Grid>
                </DataTemplate>
            </ListView.ItemTemplate>
        </ListView>

        <TextBlock Grid.Row="4" Name="StatusText"
                   Foreground="#888" FontSize="12" Margin="0,8,0,4"
                   Text="0 selected"/>

        <StackPanel Grid.Row="5" Orientation="Horizontal"
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

    $searchBox   = $window.FindName('SearchBox')
    $listView    = $window.FindName('ProfileListView')
    $statusText  = $window.FindName('StatusText')
    $removeBtn   = $window.FindName('RemoveBtn')
    $cancelBtn   = $window.FindName('CancelBtn')
    $deselectAll = $window.FindName('DeselectAllBtn')

    $tierColors = @{ User = '#4ec969'; Current = '#5b9bd5'; System = '#777777' }

    $allItems = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
    foreach ($p in $Profiles) {
        $item = [PSCustomObject]@{
            IsChecked = $false
            CanSelect = (-not $p.Protected)
            Username  = $p.Username
            Path      = $p.Path
            SID       = $p.SID
            LastUsed  = $p.LastUsed
            Size      = $p.Size
            Tier      = $p.Tier
            TierColor = $tierColors[$p.Tier]
            TextColor = if ($p.Protected) { '#666666' } else { '#cccccc' }
        }
        $allItems.Add($item)
    }

    $listView.ItemsSource = $allItems

    $applyFilter = {
        $term = $searchBox.Text.Trim().ToLower()
        if ($term -eq '') {
            $listView.ItemsSource = $allItems
        }
        else {
            $filtered = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
            foreach ($item in $allItems) {
                if ($item.Username.ToLower().Contains($term)) {
                    $filtered.Add($item)
                }
            }
            $listView.ItemsSource = $filtered
        }
    }

    $updateStatus = {
        $count = @($allItems | Where-Object { $_.IsChecked }).Count
        $statusText.Text = ('{0} selected' -f $count)
    }

    $searchBox.Add_TextChanged({
        & $applyFilter
    })

    $deselectAll.Add_Click({
        foreach ($item in $listView.ItemsSource) {
            if ($item.CanSelect) { $item.IsChecked = $false }
        }
        $listView.Items.Refresh()
        & $updateStatus
    })

    $listView.Add_PreviewMouseUp({ & $updateStatus })
    $listView.Add_KeyUp({ & $updateStatus })

    # Async size calculation - one profile per tick so UI stays responsive
    $script:sizeIndex = 0
    $sizeTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $sizeTimer.Interval = [TimeSpan]::FromMilliseconds(50)
    $sizeTimer.Add_Tick({
        if ($script:sizeIndex -ge $allItems.Count) {
            $sizeTimer.Stop()
            return
        }
        $item = $allItems[$script:sizeIndex]
        $mb = Get-FolderSizeMB -Path $item.Path
        $item.Size = Format-Size -MB $mb
        $listView.Items.Refresh()
        $script:sizeIndex++
    })
    $sizeTimer.Start()

    $cancelBtn.Add_Click({ $window.DialogResult = $false; $window.Close() })

    $removeBtn.Add_Click({
        $checked = @($allItems | Where-Object { $_.IsChecked }).Count
        if ($checked -eq 0) {
            [System.Windows.MessageBox]::Show(
                'No profiles selected.',
                'Nothing to do',
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            ) | Out-Null
            return
        }

        $names = ($allItems | Where-Object { $_.IsChecked } | ForEach-Object { $_.Username }) -join ', '
        $warning = 'Permanently remove ' + $checked + ' user profile(s)?' + [Environment]::NewLine + [Environment]::NewLine + $names + [Environment]::NewLine + [Environment]::NewLine + 'All user data (documents, desktop, downloads, etc.) will be deleted. This cannot be undone.'

        $answer = [System.Windows.MessageBox]::Show(
            $warning,
            'Confirm Profile Removal',
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
        return @($allItems | Where-Object { $_.IsChecked })
    }
    return $null
}

# -- Run GUI -------------------------------------------------------------------
$selected = @(Show-ProfileSelector -Profiles $profileData)

if ($selected.Count -eq 0) {
    Write-Host 'Nothing selected - exiting.' -ForegroundColor Yellow
    exit 0
}

# -- Remove selected profiles --------------------------------------------------
foreach ($sel in $selected) {
    $profile = Get-CimInstance -Class Win32_UserProfile | Where-Object { $_.SID -eq $sel.SID }
    if (-not $profile) {
        Log ('Profile not found for: ' + $sel.Username + ' (SID: ' + $sel.SID + ')') 'WARN'
        continue
    }

    $label = '{0} ({1})' -f $sel.Username, $sel.Path
    try {
        Remove-CimInstance -InputObject $profile -ErrorAction Stop
        Log ('Removed profile: ' + $label)
    }
    catch {
        Log ('FAILED to remove profile: ' + $label + ' -- ' + $_.Exception.Message) 'ERROR'
    }
}

Log 'Profile removal complete.'
exit 0
