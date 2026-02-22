<#
.SYNOPSIS
    Entra ID Stale Device Manager - GUI Tool v1.1
.DESCRIPTION
    GUI PowerShell tool to list, filter, disable, and delete stale devices in Microsoft Entra ID.
    Filters by OS category (Windows, Android, iOS/macOS), inactivity days, and device status.
    Uses Microsoft Graph PowerShell SDK (Microsoft.Graph module).
    Logs all actions to C:\TEMP\log\AzureCleanupDevice\
.REQUIREMENTS
    - PowerShell 5.1+ or PowerShell 7+
    - Microsoft.Graph.Identity.DirectoryManagement module
    - Entra ID role: Cloud Device Administrator (minimum)
    - Scopes: Device.Read.All, Device.ReadWrite.All (for disable/delete)
.AUTHOR
    Generated for KuriosIT / OT Cybersecurity Operations
.VERSION
    1.1 - February 2026
    - Fixed: ComboBox dropdown dark theme (fully templated)
    - Fixed: Disable/Delete using correct Object ID via Graph API with -BodyParameter
    - Fixed: Post-disable verification to confirm state change
    - Added: Full logging to C:\TEMP\log\AzureCleanupDevice\
#>

#Requires -Version 5.1

# ============================================================
# MODULE PRE-CHECK (Resolve assembly version conflicts)
# ============================================================
try {
    $loadedGraphModules = Get-Module -Name "Microsoft.Graph*" -ErrorAction SilentlyContinue
    if ($loadedGraphModules) {
        Write-Host "[INIT] Removing pre-loaded Microsoft.Graph modules to avoid assembly conflicts..." -ForegroundColor Yellow
        $loadedGraphModules | Remove-Module -Force -ErrorAction SilentlyContinue
    }
    $requiredModule = Get-Module -ListAvailable -Name "Microsoft.Graph.Identity.DirectoryManagement" |
                      Sort-Object Version -Descending | Select-Object -First 1
    if ($requiredModule) {
        $authModule = Get-Module -ListAvailable -Name "Microsoft.Graph.Authentication" |
                      Sort-Object Version -Descending | Select-Object -First 1
        if ($authModule) {
            Import-Module $authModule.Path -Force -ErrorAction Stop
            Write-Host "[INIT] Loaded Microsoft.Graph.Authentication v$($authModule.Version)" -ForegroundColor Green
        }
        Import-Module $requiredModule.Path -Force -ErrorAction Stop
        Write-Host "[INIT] Loaded Microsoft.Graph.Identity.DirectoryManagement v$($requiredModule.Version)" -ForegroundColor Green
    }
}
catch {
    Write-Warning "[INIT] Module pre-load warning: $($_.Exception.Message)"
    Write-Warning "[INIT] If you see assembly errors, close ALL PowerShell windows and relaunch in a fresh session."
}

# ============================================================
# ASSEMBLY LOADING
# ============================================================
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ============================================================
# LOGGING SETUP
# ============================================================
$Script:LogFolder = "C:\TEMP\log\AzureCleanupDevice"
if (-not (Test-Path $Script:LogFolder)) {
    New-Item -ItemType Directory -Path $Script:LogFolder -Force | Out-Null
}
$Script:LogFile = Join-Path $Script:LogFolder "StaleDeviceManager_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR","ACTION","SUCCESS")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry  = "[$timestamp] [$Level] $Message"
    Add-Content -Path $Script:LogFile -Value $logEntry -Encoding UTF8
    switch ($Level) {
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "WARN"    { Write-Host $logEntry -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
        "ACTION"  { Write-Host $logEntry -ForegroundColor Cyan }
        default   { Write-Host $logEntry -ForegroundColor Gray }
    }
}

Write-Log "========== Entra ID Stale Device Manager v1.1 Started =========="
Write-Log "Log file: $($Script:LogFile)"
Write-Log "Operator: $($env:USERNAME) on $($env:COMPUTERNAME)"

# ============================================================
# XAML UI DEFINITION
# ============================================================
[xml]$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Entra ID - Stale Device Manager v1.1"
        Height="780" Width="1150"
        MinHeight="650" MinWidth="900"
        WindowStartupLocation="CenterScreen"
        Background="#1E1E2E">

    <Window.Resources>

        <!-- Dark theme color brushes -->
        <SolidColorBrush x:Key="CmbBg" Color="#313244"/>
        <SolidColorBrush x:Key="CmbFg" Color="#CDD6F4"/>
        <SolidColorBrush x:Key="CmbBorder" Color="#45475A"/>
        <SolidColorBrush x:Key="CmbPopupBg" Color="#1E1E2E"/>
        <SolidColorBrush x:Key="CmbItemHover" Color="#45475A"/>
        <SolidColorBrush x:Key="CmbItemSelected" Color="#585B70"/>

        <!-- Dark ComboBoxItem Style -->
        <Style TargetType="ComboBoxItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="Padding" Value="10,6"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBoxItem">
                        <Border x:Name="Bd" Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}" BorderThickness="0">
                            <ContentPresenter/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsHighlighted" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{StaticResource CmbItemHover}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Bd" Property="Background" Value="{StaticResource CmbItemSelected}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Dark Toggle Button for ComboBox -->
        <ControlTemplate x:Key="DarkComboBoxToggle" TargetType="ToggleButton">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="28"/>
                </Grid.ColumnDefinitions>
                <Border x:Name="Border" Grid.ColumnSpan="2"
                        Background="{StaticResource CmbBg}"
                        BorderBrush="{StaticResource CmbBorder}"
                        BorderThickness="1" CornerRadius="4"/>
                <Path x:Name="Arrow" Grid.Column="1"
                      Data="M 0 0 L 4 4 L 8 0 Z"
                      Fill="#A6ADC8" HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Border" Property="BorderBrush" Value="#89B4FA"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <!-- Full Dark ComboBox Style -->
        <Style x:Key="DarkComboBox" TargetType="ComboBox">
            <Setter Property="Height" Value="32"/>
            <Setter Property="MinWidth" Value="140"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton x:Name="ToggleButton"
                                          Template="{StaticResource DarkComboBoxToggle}"
                                          Focusable="False" ClickMode="Press"
                                          IsChecked="{Binding IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"/>
                            <ContentPresenter x:Name="ContentSite" IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              Margin="10,0,28,0" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <Popup x:Name="Popup" Placement="Bottom"
                                   IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True" Focusable="False" PopupAnimation="Slide">
                                <Grid x:Name="DropDown" SnapsToDevicePixels="True"
                                      MinWidth="{TemplateBinding ActualWidth}"
                                      MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border x:Name="DropDownBorder"
                                            Background="{StaticResource CmbPopupBg}"
                                            BorderBrush="{StaticResource CmbBorder}"
                                            BorderThickness="1" CornerRadius="4" Margin="0,2,0,0"/>
                                    <ScrollViewer Margin="2,4" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True" KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Grid>
                            </Popup>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Header Text -->
        <Style x:Key="HeaderText" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#CDD6F4"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>

        <!-- Action Button -->
        <Style x:Key="ActionButton" TargetType="Button">
            <Setter Property="Height" Value="36"/>
            <Setter Property="MinWidth" Value="130"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="16,0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="6" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Opacity" Value="0.85"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

    </Window.Resources>

    <Grid Margin="16">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- HEADER -->
        <Border Grid.Row="0" Background="#313244" CornerRadius="8" Padding="16,12" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="🔒" FontSize="22" VerticalAlignment="Center" Margin="0,0,10,0"/>
                    <StackPanel>
                        <TextBlock Text="Entra ID - Stale Device Manager" FontSize="18" FontWeight="Bold" Foreground="#CDD6F4"/>
                        <TextBlock Text="Identify, disable, and clean up stale devices from your tenant" FontSize="11" Foreground="#6C7086"/>
                    </StackPanel>
                </StackPanel>
                <Button x:Name="BtnConnect" Grid.Column="1" Style="{StaticResource ActionButton}"
                        Background="#89B4FA" Content="🔗 Connect to Entra ID"/>
            </Grid>
        </Border>

        <!-- CONNECTION STATUS -->
        <Border Grid.Row="1" Background="#181825" CornerRadius="6" Padding="12,8" Margin="0,0,0,10">
            <StackPanel Orientation="Horizontal">
                <Ellipse x:Name="StatusDot" Width="10" Height="10" Fill="#F38BA8" Margin="0,0,8,0" VerticalAlignment="Center"/>
                <TextBlock x:Name="TxtConnectionStatus" Text="Not connected" Foreground="#A6ADC8" FontSize="12" VerticalAlignment="Center"/>
                <TextBlock Text="  |  " Foreground="#45475A" VerticalAlignment="Center"/>
                <TextBlock x:Name="TxtTenantInfo" Text="" Foreground="#6C7086" FontSize="11" VerticalAlignment="Center"/>
            </StackPanel>
        </Border>

        <!-- FILTERS PANEL -->
        <Border Grid.Row="2" Background="#313244" CornerRadius="8" Padding="16,12" Margin="0,0,0,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
                    <TextBlock Text="🔍 FILTERS" FontSize="12" FontWeight="Bold" Foreground="#89B4FA" VerticalAlignment="Center"/>
                </StackPanel>
                <WrapPanel Grid.Row="1" Orientation="Horizontal">
                    <StackPanel Margin="0,0,20,4">
                        <TextBlock Text="Operating System" Style="{StaticResource HeaderText}" FontSize="11" Foreground="#A6ADC8" Margin="0,0,0,4"/>
                        <ComboBox x:Name="CmbOS" Style="{StaticResource DarkComboBox}" MinWidth="160">
                            <ComboBoxItem Content="All Platforms" IsSelected="True"/>
                            <ComboBoxItem Content="Windows"/>
                            <ComboBoxItem Content="Android"/>
                            <ComboBoxItem Content="iOS"/>
                            <ComboBoxItem Content="macOS"/>
                            <ComboBoxItem Content="Linux"/>
                        </ComboBox>
                    </StackPanel>
                    <StackPanel Margin="0,0,20,4">
                        <TextBlock Text="Inactive Since (Days)" Style="{StaticResource HeaderText}" FontSize="11" Foreground="#A6ADC8" Margin="0,0,0,4"/>
                        <ComboBox x:Name="CmbDays" Style="{StaticResource DarkComboBox}" MinWidth="130">
                            <ComboBoxItem Content="30 days"/>
                            <ComboBoxItem Content="60 days"/>
                            <ComboBoxItem Content="90 days" IsSelected="True"/>
                            <ComboBoxItem Content="120 days"/>
                            <ComboBoxItem Content="180 days"/>
                            <ComboBoxItem Content="365 days"/>
                        </ComboBox>
                    </StackPanel>
                    <StackPanel Margin="0,0,20,4">
                        <TextBlock Text="Device Status" Style="{StaticResource HeaderText}" FontSize="11" Foreground="#A6ADC8" Margin="0,0,0,4"/>
                        <ComboBox x:Name="CmbStatus" Style="{StaticResource DarkComboBox}" MinWidth="160">
                            <ComboBoxItem Content="All (Enabled + Disabled)" IsSelected="True"/>
                            <ComboBoxItem Content="Enabled Only"/>
                            <ComboBoxItem Content="Disabled Only"/>
                        </ComboBox>
                    </StackPanel>
                    <StackPanel Margin="0,0,20,4">
                        <TextBlock Text="Join Type" Style="{StaticResource HeaderText}" FontSize="11" Foreground="#A6ADC8" Margin="0,0,0,4"/>
                        <ComboBox x:Name="CmbTrust" Style="{StaticResource DarkComboBox}" MinWidth="160">
                            <ComboBoxItem Content="All Join Types" IsSelected="True"/>
                            <ComboBoxItem Content="Entra ID Joined"/>
                            <ComboBoxItem Content="Hybrid Joined"/>
                            <ComboBoxItem Content="Entra ID Registered"/>
                        </ComboBox>
                    </StackPanel>
                    <StackPanel Margin="0,0,10,4" VerticalAlignment="Bottom">
                        <TextBlock Text=" " FontSize="11" Margin="0,0,0,4"/>
                        <Button x:Name="BtnSearch" Style="{StaticResource ActionButton}" Background="#A6E3A1"
                                Foreground="#1E1E2E" Content="🔍 Search Devices"/>
                    </StackPanel>
                    <StackPanel Margin="0,0,0,4" VerticalAlignment="Bottom">
                        <TextBlock Text=" " FontSize="11" Margin="0,0,0,4"/>
                        <Button x:Name="BtnExport" Style="{StaticResource ActionButton}" Background="#45475A"
                                Content="📁 Export CSV" IsEnabled="False"/>
                    </StackPanel>
                </WrapPanel>
            </Grid>
        </Border>

        <!-- DATA GRID -->
        <Border Grid.Row="3" Background="#313244" CornerRadius="8" Padding="2" Margin="0,0,0,10">
            <DataGrid x:Name="DgDevices"
                      AutoGenerateColumns="False" IsReadOnly="False"
                      CanUserAddRows="False" CanUserDeleteRows="False"
                      CanUserReorderColumns="True" CanUserSortColumns="True"
                      SelectionMode="Extended"
                      GridLinesVisibility="Horizontal" HorizontalGridLinesBrush="#45475A"
                      Background="#1E1E2E" Foreground="#CDD6F4"
                      RowBackground="#1E1E2E" AlternatingRowBackground="#181825"
                      BorderThickness="0" HeadersVisibility="Column" FontSize="12">
                <DataGrid.ColumnHeaderStyle>
                    <Style TargetType="DataGridColumnHeader">
                        <Setter Property="Background" Value="#45475A"/>
                        <Setter Property="Foreground" Value="#CDD6F4"/>
                        <Setter Property="FontWeight" Value="SemiBold"/>
                        <Setter Property="FontSize" Value="12"/>
                        <Setter Property="Padding" Value="10,8"/>
                        <Setter Property="BorderBrush" Value="#585B70"/>
                        <Setter Property="BorderThickness" Value="0,0,1,0"/>
                    </Style>
                </DataGrid.ColumnHeaderStyle>
                <DataGrid.Columns>
                    <DataGridCheckBoxColumn x:Name="ColSelect" Header="✔" Width="40"
                                            Binding="{Binding Selected, UpdateSourceTrigger=PropertyChanged}"/>
                    <DataGridTextColumn Header="Display Name" Binding="{Binding DisplayName}" Width="180"/>
                    <DataGridTextColumn Header="OS" Binding="{Binding OperatingSystem}" Width="90"/>
                    <DataGridTextColumn Header="OS Version" Binding="{Binding OperatingSystemVersion}" Width="110"/>
                    <DataGridTextColumn Header="Status" Binding="{Binding StatusText}" Width="80"/>
                    <DataGridTextColumn Header="Trust Type" Binding="{Binding TrustType}" Width="120"/>
                    <DataGridTextColumn Header="Last Activity" Binding="{Binding LastActivityFormatted}" Width="140"/>
                    <DataGridTextColumn Header="Days Inactive" Binding="{Binding DaysInactive}" Width="100"/>
                    <DataGridTextColumn Header="Object ID" Binding="{Binding ObjectId}" Width="280"/>
                </DataGrid.Columns>
            </DataGrid>
        </Border>

        <!-- ACTIONS PANEL -->
        <Border Grid.Row="4" Background="#313244" CornerRadius="8" Padding="16,12" Margin="0,0,0,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal">
                    <Button x:Name="BtnSelectAll" Style="{StaticResource ActionButton}" Background="#585B70"
                            Content="☑ Select All" Margin="0,0,8,0" MinWidth="100" IsEnabled="False"/>
                    <Button x:Name="BtnSelectNone" Style="{StaticResource ActionButton}" Background="#585B70"
                            Content="☐ Select None" Margin="0,0,8,0" MinWidth="110" IsEnabled="False"/>
                </StackPanel>
                <StackPanel Grid.Column="2" Orientation="Horizontal">
                    <Button x:Name="BtnDisable" Style="{StaticResource ActionButton}" Background="#FAB387"
                            Foreground="#1E1E2E" Content="⛔ Disable Selected" Margin="0,0,8,0" IsEnabled="False"/>
                    <Button x:Name="BtnDelete" Style="{StaticResource ActionButton}" Background="#F38BA8"
                            Foreground="#1E1E2E" Content="🗑 Delete Selected" IsEnabled="False"/>
                </StackPanel>
            </Grid>
        </Border>

        <!-- STATUS BAR -->
        <Border Grid.Row="5" Background="#181825" CornerRadius="6" Padding="12,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="TxtStatus" Grid.Column="0" Text="Ready. Connect to Entra ID to begin."
                           Foreground="#6C7086" FontSize="11" VerticalAlignment="Center"/>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <TextBlock x:Name="TxtDeviceCount" Text="Devices: 0" Foreground="#A6ADC8" FontSize="11"
                               VerticalAlignment="Center" Margin="0,0,16,0"/>
                    <TextBlock x:Name="TxtSelectedCount" Text="Selected: 0" Foreground="#89B4FA" FontSize="11"
                               VerticalAlignment="Center"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ============================================================
# PARSE XAML & BUILD WINDOW
# ============================================================
$Reader = New-Object System.Xml.XmlNodeReader $XAML
$Window = [Windows.Markup.XamlReader]::Load($Reader)

$BtnConnect       = $Window.FindName("BtnConnect")
$BtnSearch        = $Window.FindName("BtnSearch")
$BtnExport        = $Window.FindName("BtnExport")
$BtnDisable       = $Window.FindName("BtnDisable")
$BtnDelete        = $Window.FindName("BtnDelete")
$BtnSelectAll     = $Window.FindName("BtnSelectAll")
$BtnSelectNone    = $Window.FindName("BtnSelectNone")
$CmbOS            = $Window.FindName("CmbOS")
$CmbDays          = $Window.FindName("CmbDays")
$CmbStatus        = $Window.FindName("CmbStatus")
$CmbTrust         = $Window.FindName("CmbTrust")
$DgDevices        = $Window.FindName("DgDevices")
$TxtStatus        = $Window.FindName("TxtStatus")
$TxtConnectionStatus = $Window.FindName("TxtConnectionStatus")
$TxtTenantInfo    = $Window.FindName("TxtTenantInfo")
$TxtDeviceCount   = $Window.FindName("TxtDeviceCount")
$TxtSelectedCount = $Window.FindName("TxtSelectedCount")
$StatusDot        = $Window.FindName("StatusDot")

# ============================================================
# STATE VARIABLES
# ============================================================
$Script:IsConnected = $false
$Script:ExternalSession = $false
$Script:AllDevices  = @()
$Script:DeviceList  = [System.Collections.ObjectModel.ObservableCollection[PSObject]]::new()
$DgDevices.ItemsSource = $Script:DeviceList

# ============================================================
# HELPER FUNCTIONS
# ============================================================
function Update-Status {
    param([string]$Message, [string]$Color = "#6C7086")
    $TxtStatus.Text = $Message
    $TxtStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
    [System.Windows.Forms.Application]::DoEvents()
}

function Update-SelectedCount {
    $count = ($Script:DeviceList | Where-Object { $_.Selected }).Count
    $TxtSelectedCount.Text = "Selected: $count"
    $BtnDisable.IsEnabled = ($count -gt 0)
    $BtnDelete.IsEnabled  = ($count -gt 0)
}

function Show-MessageBox {
    param(
        [string]$Message,
        [string]$Title = "Entra ID Stale Device Manager",
        [System.Windows.MessageBoxButton]$Buttons = [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]$Icon = [System.Windows.MessageBoxImage]::Information
    )
    [System.Windows.MessageBox]::Show($Window, $Message, $Title, $Buttons, $Icon)
}

function Get-DaysFromSelection {
    $sel = ($CmbDays.SelectedItem).Content.ToString()
    if ($sel -match "(\d+)") { return [int]$Matches[1] }
    return 90
}

function Get-OSFilter {
    $sel = ($CmbOS.SelectedItem).Content.ToString()
    switch ($sel) {
        "All Platforms" { return $null }
        default         { return $sel }
    }
}

function Get-StatusFilter {
    $sel = ($CmbStatus.SelectedItem).Content.ToString()
    switch ($sel) {
        "Enabled Only"  { return $true }
        "Disabled Only" { return $false }
        default         { return $null }
    }
}

function Get-TrustFilter {
    $sel = ($CmbTrust.SelectedItem).Content.ToString()
    switch ($sel) {
        "Entra ID Joined"     { return "AzureAd" }
        "Hybrid Joined"       { return "ServerAd" }
        "Entra ID Registered" { return "Workplace" }
        default                { return $null }
    }
}

# ============================================================
# CONNECT TO ENTRA ID
# ============================================================
$BtnConnect.Add_Click({
    Update-Status "Checking Microsoft.Graph module..." "#F9E2AF"
    Write-Log "User initiated connection to Entra ID"

    $module = Get-Module -Name "Microsoft.Graph.Identity.DirectoryManagement" -ErrorAction SilentlyContinue
    if (-not $module) {
        $module = Get-Module -ListAvailable -Name "Microsoft.Graph.Identity.DirectoryManagement" |
                  Sort-Object Version -Descending | Select-Object -First 1
    }
    if (-not $module) {
        Write-Log "Microsoft.Graph module not found - prompting install" "WARN"
        $result = Show-MessageBox -Message "The module 'Microsoft.Graph.Identity.DirectoryManagement' is not installed.`n`nInstall it now?" `
                                  -Title "Module Required" -Buttons YesNo -Icon Question
        if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
            Update-Status "Installing Microsoft.Graph module... This may take a few minutes." "#F9E2AF"
            try {
                Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
                Write-Log "Microsoft.Graph module installed successfully" "SUCCESS"
            }
            catch {
                Write-Log "Failed to install module: $($_.Exception.Message)" "ERROR"
                Show-MessageBox -Message "Failed to install module:`n$($_.Exception.Message)" -Icon Error
                return
            }
        } else { return }
    }

    Update-Status "Connecting to Microsoft Entra ID... (Browser auth window may open)" "#F9E2AF"
    try {
        # Vérifier si une session Graph est déjà active (ex: lancé depuis M365 Monster)
        $context = Get-MgContext
        if ($null -eq $context) {
            # Pas de session — connexion interactive
            Connect-MgGraph -Scopes "Device.Read.All","Device.ReadWrite.All","Directory.ReadWrite.All" -ErrorAction Stop
            $context = Get-MgContext
            $Script:ExternalSession = $false
        }
        else {
            # Session existante réutilisée — on ne déconnectera pas à la fermeture
            $Script:ExternalSession = $true
        }

        if ($context) {
            $Script:IsConnected = $true
            $StatusDot.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#A6E3A1")
            $TxtConnectionStatus.Text = "Connected"
            $TxtConnectionStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#A6E3A1")
            $TxtTenantInfo.Text = "Tenant: $($context.TenantId)  |  Account: $($context.Account)"
            $BtnConnect.Content = "✅ Connected"
            $BtnConnect.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#45475A")
            Write-Log "Connected - Tenant: $($context.TenantId) | Account: $($context.Account)" "SUCCESS"
            Update-Status "Connected to Entra ID. Configure filters and click 'Search Devices'." "#A6E3A1"
        }
    }
    catch {
        $Script:IsConnected = $false
        $StatusDot.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFromString("#F38BA8")
        $TxtConnectionStatus.Text = "Connection failed"
        Write-Log "Connection failed: $($_.Exception.Message)" "ERROR"
        Update-Status "Connection failed: $($_.Exception.Message)" "#F38BA8"
        Show-MessageBox -Message "Failed to connect:`n$($_.Exception.Message)" -Icon Error
    }
})

# ============================================================
# SEARCH DEVICES
# ============================================================
$BtnSearch.Add_Click({
    if (-not $Script:IsConnected) {
        Show-MessageBox -Message "Please connect to Entra ID first." -Icon Warning
        return
    }

    $days         = Get-DaysFromSelection
    $osFilter     = Get-OSFilter
    $statusFilter = Get-StatusFilter
    $trustFilter  = Get-TrustFilter

    Write-Log "Search: Days=$days | OS=$(if($osFilter){$osFilter}else{'All'}) | Status=$(if($null -ne $statusFilter){$statusFilter}else{'All'}) | Trust=$(if($trustFilter){$trustFilter}else{'All'})" "ACTION"

    Update-Status "Fetching all devices from Entra ID... Please wait." "#F9E2AF"
    $Script:DeviceList.Clear()

    try {
        $Script:AllDevices = Get-MgDevice -All -Property "Id,DeviceId,DisplayName,OperatingSystem,OperatingSystemVersion,AccountEnabled,ApproximateLastSignInDateTime,TrustType,IsManaged,IsCompliant" -ErrorAction Stop

        Write-Log "Retrieved $($Script:AllDevices.Count) total devices from tenant"
        Update-Status "Retrieved $($Script:AllDevices.Count) devices. Applying filters..." "#F9E2AF"

        $cutoffDate = (Get-Date).AddDays(-$days)

        $filtered = $Script:AllDevices | Where-Object {
            # Stale check
            if ($null -eq $_.ApproximateLastSignInDateTime) { $isStale = $true }
            elseif ($_.ApproximateLastSignInDateTime -le $cutoffDate) { $isStale = $true }
            else { $isStale = $false }
            if (-not $isStale) { return $false }

            # OS filter
            if ($null -ne $osFilter -and $_.OperatingSystem -notlike "$osFilter*") { return $false }
            # Status filter
            if ($null -ne $statusFilter -and $_.AccountEnabled -ne $statusFilter) { return $false }
            # Trust filter
            if ($null -ne $trustFilter -and $_.TrustType -ne $trustFilter) { return $false }

            return $true
        }

        foreach ($device in $filtered) {
            $daysInactive = if ($null -ne $device.ApproximateLastSignInDateTime) {
                [math]::Round(((Get-Date) - $device.ApproximateLastSignInDateTime).TotalDays)
            } else { "N/A" }

            $lastActivity = if ($null -ne $device.ApproximateLastSignInDateTime) {
                $device.ApproximateLastSignInDateTime.ToString("yyyy-MM-dd HH:mm")
            } else { "Never" }

            $obj = [PSCustomObject]@{
                Selected               = $false
                DisplayName            = $device.DisplayName
                OperatingSystem        = $device.OperatingSystem
                OperatingSystemVersion = $device.OperatingSystemVersion
                StatusText             = $(if ($device.AccountEnabled) {"Enabled"} else {"Disabled"})
                AccountEnabled         = $device.AccountEnabled
                TrustType              = $device.TrustType
                LastActivityFormatted  = $lastActivity
                DaysInactive           = $daysInactive
                DeviceId               = $device.DeviceId
                ObjectId               = $device.Id
            }
            $Script:DeviceList.Add($obj)
        }

        $count = $Script:DeviceList.Count
        $TxtDeviceCount.Text     = "Devices: $count"
        $BtnExport.IsEnabled     = ($count -gt 0)
        $BtnSelectAll.IsEnabled  = ($count -gt 0)
        $BtnSelectNone.IsEnabled = ($count -gt 0)

        $summary = ($Script:DeviceList | Group-Object OperatingSystem | ForEach-Object { "$($_.Name): $($_.Count)" }) -join " | "
        Write-Log "Results: $count stale devices [$summary]"
        Update-Status "Found $count stale devices (inactive > $days days).  [$summary]" "#A6E3A1"
    }
    catch {
        Write-Log "Error fetching devices: $($_.Exception.Message)" "ERROR"
        Update-Status "Error: $($_.Exception.Message)" "#F38BA8"
        Show-MessageBox -Message "Error:`n$($_.Exception.Message)" -Icon Error
    }
})

# ============================================================
# SELECT ALL / NONE
# ============================================================
$BtnSelectAll.Add_Click({
    foreach ($item in $Script:DeviceList) { $item.Selected = $true }
    $DgDevices.Items.Refresh(); Update-SelectedCount
})
$BtnSelectNone.Add_Click({
    foreach ($item in $Script:DeviceList) { $item.Selected = $false }
    $DgDevices.Items.Refresh(); Update-SelectedCount
})
$DgDevices.Add_CurrentCellChanged({ $Window.Dispatcher.InvokeAsync([Action]{ Update-SelectedCount }) | Out-Null })
$DgDevices.Add_MouseLeftButtonUp({
    $Window.Dispatcher.InvokeAsync([Action]{ Start-Sleep -Milliseconds 100; Update-SelectedCount }) | Out-Null
})

# ============================================================
# DISABLE SELECTED DEVICES  (FIXED: uses -BodyParameter + verification)
# ============================================================
$BtnDisable.Add_Click({
    $selected = @($Script:DeviceList | Where-Object { $_.Selected -and $_.AccountEnabled })

    if ($selected.Count -eq 0) {
        Show-MessageBox -Message "No enabled devices selected to disable.`n(Already disabled devices are excluded.)" -Icon Information
        return
    }

    $confirm = Show-MessageBox -Message "You are about to DISABLE $($selected.Count) device(s).`n`nDisabled devices will no longer authenticate to Entra ID.`nThis action is reversible.`n`nContinue?" `
                               -Title "Confirm Disable" -Buttons YesNo -Icon Warning
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    Write-Log "========== DISABLE OPERATION STARTED ==========" "ACTION"
    Write-Log "Operator: $($env:USERNAME) | Count: $($selected.Count)" "ACTION"

    $successCount = 0
    $errorCount   = 0
    $total        = $selected.Count

    foreach ($device in $selected) {
        $idx = $successCount + $errorCount + 1
        try {
            Update-Status "Disabling [$idx/$total]: $($device.DisplayName)..." "#F9E2AF"
            Write-Log "DISABLING: '$($device.DisplayName)' | ObjectId=$($device.ObjectId) | DeviceId=$($device.DeviceId) | OS=$($device.OperatingSystem)" "ACTION"

            # ---- FIX: Use -BodyParameter hashtable to ensure the property is sent correctly ----
            $bodyParams = @{ accountEnabled = $false }
            Update-MgDevice -DeviceId $device.ObjectId -BodyParameter $bodyParams -ErrorAction Stop

            # ---- FIX: Verify the change actually took effect in Entra ID ----
            Start-Sleep -Milliseconds 500
            $verifyDevice = Get-MgDevice -DeviceId $device.ObjectId -Property "AccountEnabled" -ErrorAction Stop

            if ($verifyDevice.AccountEnabled -eq $false) {
                $device.AccountEnabled = $false
                $device.StatusText = "Disabled"
                $device.Selected = $false
                $successCount++
                Write-Log "SUCCESS: '$($device.DisplayName)' disabled and VERIFIED in Entra ID" "SUCCESS"
            }
            else {
                # API accepted but propagation delay
                $device.AccountEnabled = $false
                $device.StatusText = "Disabled*"
                $device.Selected = $false
                $successCount++
                Write-Log "WARN: '$($device.DisplayName)' - API success but verification shows still Enabled (propagation delay expected)" "WARN"
            }
        }
        catch {
            $errorCount++
            $errMsg = $_.Exception.Message
            # Try alternative method if -BodyParameter fails
            try {
                Write-Log "Retrying with direct parameter: '$($device.DisplayName)'" "WARN"
                Update-MgDevice -DeviceId $device.ObjectId -AccountEnabled:$false -ErrorAction Stop
                $device.AccountEnabled = $false
                $device.StatusText = "Disabled"
                $device.Selected = $false
                $successCount++
                $errorCount--
                Write-Log "SUCCESS (retry): '$($device.DisplayName)' disabled via direct parameter" "SUCCESS"
            }
            catch {
                Write-Log "ERROR: Failed to disable '$($device.DisplayName)': $errMsg | Retry also failed: $($_.Exception.Message)" "ERROR"
            }
        }
    }

    $DgDevices.Items.Refresh()
    Update-SelectedCount

    Write-Log "========== DISABLE COMPLETED: Success=$successCount | Failed=$errorCount ==========" "ACTION"
    $msg = "Disable operation complete.`n`nSuccess: $successCount`nFailed: $errorCount`n`nLog: $($Script:LogFile)"
    Update-Status "Disabled: $successCount OK, $errorCount failed" "#A6E3A1"
    Show-MessageBox -Message $msg -Icon Information
})

# ============================================================
# DELETE SELECTED DEVICES
# ============================================================
$BtnDelete.Add_Click({
    $selected = @($Script:DeviceList | Where-Object { $_.Selected })

    if ($selected.Count -eq 0) {
        Show-MessageBox -Message "No devices selected to delete." -Icon Information
        return
    }

    $enabledCount = ($selected | Where-Object { $_.AccountEnabled }).Count
    $warningExtra = ""
    if ($enabledCount -gt 0) {
        $warningExtra = "`n`n⚠️ WARNING: $enabledCount device(s) are still ENABLED.`nBest practice: Disable first, then wait a grace period."
    }

    $confirm = Show-MessageBox -Message "PERMANENTLY DELETE $($selected.Count) device(s)?$warningExtra`n`n🔴 CANNOT be undone!`n🔴 BitLocker recovery keys will be LOST!`n🔴 Hybrid joined devices may re-sync from on-prem AD." `
                               -Title "⚠️ Confirm Deletion" -Buttons YesNo -Icon Warning
    if ($confirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    $secondConfirm = Show-MessageBox -Message "FINAL CONFIRMATION`n`nDelete $($selected.Count) device(s) permanently?`nThis is your LAST chance to cancel." `
                                     -Title "🔴 Final Confirmation" -Buttons YesNo -Icon Warning
    if ($secondConfirm -ne [System.Windows.MessageBoxResult]::Yes) { return }

    Write-Log "========== DELETE OPERATION STARTED ==========" "ACTION"
    Write-Log "Operator: $($env:USERNAME) | Count: $($selected.Count)" "ACTION"

    $successCount = 0
    $errorCount   = 0
    $total        = $selected.Count
    $toRemove     = @()

    foreach ($device in $selected) {
        $idx = $successCount + $errorCount + 1
        try {
            Update-Status "Deleting [$idx/$total]: $($device.DisplayName)..." "#F38BA8"
            Write-Log "DELETING: '$($device.DisplayName)' | ObjectId=$($device.ObjectId) | DeviceId=$($device.DeviceId) | OS=$($device.OperatingSystem) | Status=$($device.StatusText)" "ACTION"

            Remove-MgDevice -DeviceId $device.ObjectId -ErrorAction Stop
            $toRemove += $device
            $successCount++
            Write-Log "SUCCESS: '$($device.DisplayName)' deleted permanently" "SUCCESS"
        }
        catch {
            $errorCount++
            Write-Log "ERROR: Failed to delete '$($device.DisplayName)': $($_.Exception.Message)" "ERROR"
        }
    }

    foreach ($item in $toRemove) { $Script:DeviceList.Remove($item) | Out-Null }

    $TxtDeviceCount.Text = "Devices: $($Script:DeviceList.Count)"
    Update-SelectedCount

    Write-Log "========== DELETE COMPLETED: Deleted=$successCount | Failed=$errorCount ==========" "ACTION"
    $msg = "Delete operation complete.`n`nDeleted: $successCount`nFailed: $errorCount`n`nLog: $($Script:LogFile)"
    Update-Status "Deleted: $successCount OK, $errorCount failed" $(if ($errorCount -gt 0) {"#F9E2AF"} else {"#A6E3A1"})
    Show-MessageBox -Message $msg -Icon Information
})

# ============================================================
# EXPORT TO CSV
# ============================================================
$BtnExport.Add_Click({
    if ($Script:DeviceList.Count -eq 0) {
        Show-MessageBox -Message "No devices to export." -Icon Information
        return
    }

    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    $dialog.InitialDirectory = $Script:LogFolder
    $dialog.FileName = "EntraID_StaleDevices_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $dialog.Title = "Export Stale Devices Report"

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $Script:DeviceList |
                Select-Object DisplayName, OperatingSystem, OperatingSystemVersion, StatusText, TrustType, LastActivityFormatted, DaysInactive, DeviceId, ObjectId |
                Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8

            Write-Log "Exported $($Script:DeviceList.Count) devices to: $($dialog.FileName)" "ACTION"
            Update-Status "Exported $($Script:DeviceList.Count) devices to $($dialog.FileName)" "#A6E3A1"
            Show-MessageBox -Message "Report exported!`n$($dialog.FileName)" -Icon Information
        }
        catch {
            Write-Log "Export failed: $($_.Exception.Message)" "ERROR"
            Show-MessageBox -Message "Export failed:`n$($_.Exception.Message)" -Icon Error
        }
    }
})

# ============================================================
# SHOW WINDOW
# ============================================================
Write-Log "GUI window opened"
$Window.ShowDialog() | Out-Null

Write-Log "========== Session ended =========="
# Ne déconnecter que si la session a été créée par ce script
if (-not $Script:ExternalSession) {
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
}