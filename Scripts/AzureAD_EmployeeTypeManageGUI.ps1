<#
.SYNOPSIS
    Entra ID Employee Type Manager - GUI Tool
.DESCRIPTION
    PowerShell WPF GUI application to manage the 'employeeType' attribute
    in Microsoft Entra ID (Azure AD). Allows categorizing all user accounts
    with bulk assignment, CSV import, and real-time dashboard statistics.
.NOTES
    Requires: Microsoft.Graph.Users PowerShell module
    Permissions: User.ReadWrite.All (delegated or application)
    Install: Install-Module Microsoft.Graph.Users -Scope CurrentUser
    Author: Leo - KuriosIT / Valto
#>

#Requires -Modules Microsoft.Graph.Users

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ============================================================
# XAML - WPF Interface Definition
# ============================================================
[xml]$XAML = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Entra ID - Employee Type Manager"
    Width="1280" Height="820"
    MinWidth="1100" MinHeight="700"
    WindowStartupLocation="CenterScreen"
    Background="#1b1f2b"
    FontFamily="Segoe UI">

    <Window.Resources>
        <!-- Color Palette -->
        <SolidColorBrush x:Key="BgDark" Color="#1b1f2b"/>
        <SolidColorBrush x:Key="BgCard" Color="#242938"/>
        <SolidColorBrush x:Key="BgCardHover" Color="#2d3348"/>
        <SolidColorBrush x:Key="BorderSubtle" Color="#333a4f"/>
        <SolidColorBrush x:Key="AccentBlue" Color="#4f8cff"/>
        <SolidColorBrush x:Key="AccentGreen" Color="#34d399"/>
        <SolidColorBrush x:Key="AccentOrange" Color="#fb923c"/>
        <SolidColorBrush x:Key="AccentRed" Color="#f87171"/>
        <SolidColorBrush x:Key="AccentPurple" Color="#a78bfa"/>
        <SolidColorBrush x:Key="TextPrimary" Color="#e2e8f0"/>
        <SolidColorBrush x:Key="TextSecondary" Color="#94a3b8"/>
        <SolidColorBrush x:Key="TextMuted" Color="#64748b"/>

        <!-- Button Style -->
        <Style x:Key="ModernButton" TargetType="Button">
            <Setter Property="Background" Value="#4f8cff"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="20,10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="border"
                                Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#6ba0ff"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#3a72e0"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="border" Property="Background" Value="#3a4260"/>
                                <Setter Property="Foreground" Value="#64748b"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#333a4f"/>
        </Style>

        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#dc2626"/>
        </Style>

        <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource ModernButton}">
            <Setter Property="Background" Value="#059669"/>
        </Style>

        <!-- TextBox Style -->
        <Style x:Key="ModernTextBox" TargetType="TextBox">
            <Setter Property="Background" Value="#1b1f2b"/>
            <Setter Property="Foreground" Value="#e2e8f0"/>
            <Setter Property="BorderBrush" Value="#333a4f"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="CaretBrush" Value="#4f8cff"/>
        </Style>

        <!-- ComboBox Dark Theme - Full Template Override -->
        <ControlTemplate x:Key="ComboBoxToggleButton" TargetType="ToggleButton">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition/>
                    <ColumnDefinition Width="30"/>
                </Grid.ColumnDefinitions>
                <Border x:Name="Border" Grid.ColumnSpan="2" CornerRadius="4"
                        Background="#1b1f2b" BorderBrush="#333a4f" BorderThickness="1"/>
                <Border Grid.Column="0" CornerRadius="4,0,0,4" Margin="1"
                        Background="Transparent"/>
                <Path x:Name="Arrow" Grid.Column="1" HorizontalAlignment="Center" VerticalAlignment="Center"
                      Data="M 0 0 L 4 4 L 8 0 Z" Fill="#94a3b8"/>
            </Grid>
            <ControlTemplate.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter TargetName="Border" Property="BorderBrush" Value="#4f8cff"/>
                </Trigger>
            </ControlTemplate.Triggers>
        </ControlTemplate>

        <ControlTemplate x:Key="ComboBoxTextBox" TargetType="TextBox">
            <Border x:Name="PART_ContentHost" Focusable="False"
                    Background="Transparent"/>
        </ControlTemplate>

        <Style x:Key="ModernComboBox" TargetType="ComboBox">
            <Setter Property="Foreground" Value="#e2e8f0"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="SnapsToDevicePixels" Value="True"/>
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Setter Property="ScrollViewer.HorizontalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.VerticalScrollBarVisibility" Value="Auto"/>
            <Setter Property="ScrollViewer.CanContentScroll" Value="True"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ComboBox">
                        <Grid>
                            <ToggleButton Name="ToggleButton"
                                          Template="{StaticResource ComboBoxToggleButton}"
                                          Grid.Column="2" Focusable="False"
                                          IsChecked="{Binding Path=IsDropDownOpen, Mode=TwoWay, RelativeSource={RelativeSource TemplatedParent}}"
                                          ClickMode="Press"/>
                            <ContentPresenter Name="ContentSite" IsHitTestVisible="False"
                                              Content="{TemplateBinding SelectionBoxItem}"
                                              ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"
                                              ContentTemplateSelector="{TemplateBinding ItemTemplateSelector}"
                                              Margin="10,8,30,8" VerticalAlignment="Center"
                                              HorizontalAlignment="Left"/>
                            <TextBox x:Name="PART_EditableTextBox"
                                     Style="{x:Null}"
                                     Template="{StaticResource ComboBoxTextBox}"
                                     HorizontalAlignment="Left" VerticalAlignment="Center"
                                     Margin="10,8,30,8" Focusable="True"
                                     Background="Transparent" Foreground="#e2e8f0"
                                     CaretBrush="#4f8cff"
                                     Visibility="Hidden" IsReadOnly="{TemplateBinding IsReadOnly}"/>
                            <Popup Name="Popup" Placement="Bottom"
                                   IsOpen="{TemplateBinding IsDropDownOpen}"
                                   AllowsTransparency="True" Focusable="False"
                                   PopupAnimation="Slide">
                                <Grid Name="DropDown" SnapsToDevicePixels="True"
                                      MinWidth="{TemplateBinding ActualWidth}"
                                      MaxHeight="{TemplateBinding MaxDropDownHeight}">
                                    <Border x:Name="DropDownBorder"
                                            Background="#242938" BorderBrush="#333a4f"
                                            BorderThickness="1" CornerRadius="4"
                                            Margin="0,2,0,0"/>
                                    <ScrollViewer Margin="4,6,4,6" SnapsToDevicePixels="True">
                                        <StackPanel IsItemsHost="True"
                                                    KeyboardNavigation.DirectionalNavigation="Contained"/>
                                    </ScrollViewer>
                                </Grid>
                            </Popup>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="HasItems" Value="False">
                                <Setter TargetName="DropDownBorder" Property="MinHeight" Value="95"/>
                            </Trigger>
                            <Trigger Property="IsEditable" Value="True">
                                <Setter Property="IsTabStop" Value="False"/>
                                <Setter TargetName="PART_EditableTextBox" Property="Visibility" Value="Visible"/>
                                <Setter TargetName="ContentSite" Property="Visibility" Value="Hidden"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Setter Property="ItemContainerStyle">
                <Setter.Value>
                    <Style TargetType="ComboBoxItem">
                        <Setter Property="SnapsToDevicePixels" Value="True"/>
                        <Setter Property="OverridesDefaultStyle" Value="True"/>
                        <Setter Property="Foreground" Value="#e2e8f0"/>
                        <Setter Property="Padding" Value="10,8"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="ComboBoxItem">
                                    <Border Name="Border" Padding="{TemplateBinding Padding}"
                                            SnapsToDevicePixels="True" Background="Transparent"
                                            CornerRadius="3">
                                        <ContentPresenter/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsHighlighted" Value="True">
                                            <Setter TargetName="Border" Property="Background" Value="#4f8cff"/>
                                            <Setter Property="Foreground" Value="#ffffff"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="0">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- ==================== HEADER ==================== -->
        <Border Grid.Row="0" Background="#242938" BorderBrush="#333a4f" BorderThickness="0,0,0,1" Padding="24,16">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>

                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="&#xE7EE;" FontFamily="Segoe MDL2 Assets" FontSize="28"
                               Foreground="#4f8cff" VerticalAlignment="Center" Margin="0,0,14,0"/>
                    <StackPanel>
                        <TextBlock Text="Entra ID - Employee Type Manager" FontSize="20" FontWeight="Bold"
                                   Foreground="#e2e8f0"/>
                        <TextBlock Text="Catégorisation des comptes Microsoft Entra ID"
                                   FontSize="12" Foreground="#64748b" Margin="0,2,0,0"/>
                    </StackPanel>
                </StackPanel>

                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="btnConnect" Style="{StaticResource SuccessButton}"
                            Padding="16,10" Margin="0,0,10,0">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="&#xE77B;" FontFamily="Segoe MDL2 Assets" FontSize="14"
                                       VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <TextBlock Text="Connexion Graph" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Button>
                    <Button x:Name="btnRefresh" Style="{StaticResource SecondaryButton}"
                            Margin="0,0,10,0" Padding="16,10" IsEnabled="False">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="&#xE72C;" FontFamily="Segoe MDL2 Assets" FontSize="14"
                                       VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <TextBlock Text="Actualiser" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Button>
                    <Button x:Name="btnExportReport" Style="{StaticResource SecondaryButton}"
                            Padding="16,10" IsEnabled="False">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="&#xE78C;" FontFamily="Segoe MDL2 Assets" FontSize="14"
                                       VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <TextBlock Text="Exporter Rapport" VerticalAlignment="Center"/>
                        </StackPanel>
                    </Button>
                </StackPanel>
            </Grid>
        </Border>

        <!-- ==================== MAIN CONTENT ==================== -->
        <Grid Grid.Row="1" Margin="24,20,24,10">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="280"/>
                <ColumnDefinition Width="20"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- ========== LEFT PANEL - Dashboard & Actions ========== -->
            <ScrollViewer Grid.Column="0" VerticalScrollBarVisibility="Auto">
                <StackPanel>

                    <!-- Dashboard Card - Uncategorized -->
                    <Border Background="#242938" CornerRadius="10" Padding="20" Margin="0,0,0,14"
                            BorderBrush="#333a4f" BorderThickness="1">
                        <StackPanel>
                            <TextBlock Text="COMPTES NON CATÉGORISÉS" FontSize="10" FontWeight="Bold"
                                       Foreground="#64748b"/>
                            <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                                <TextBlock x:Name="txtUncategorized" Text="--" FontSize="42" FontWeight="Bold"
                                           Foreground="#fb923c"/>
                                <TextBlock x:Name="txtTotalUsers" Text="/--" FontSize="16"
                                           Foreground="#64748b" VerticalAlignment="Bottom" Margin="4,0,0,8"/>
                            </StackPanel>
                            <!-- Progress Bar -->
                            <Grid Margin="0,12,0,0">
                                <Border Background="#333a4f" CornerRadius="4" Height="8"/>
                                <Border x:Name="progressBar" Background="#34d399" CornerRadius="4"
                                        Height="8" HorizontalAlignment="Left" Width="0"/>
                            </Grid>
                            <TextBlock x:Name="txtProgressPercent" Text="--% catégorisés"
                                       FontSize="11" Foreground="#64748b" Margin="0,6,0,0"/>
                        </StackPanel>
                    </Border>

                    <!-- Stats by Category -->
                    <Border Background="#242938" CornerRadius="10" Padding="20" Margin="0,0,0,14"
                            BorderBrush="#333a4f" BorderThickness="1">
                        <StackPanel>
                            <TextBlock Text="RÉPARTITION PAR TYPE" FontSize="10" FontWeight="Bold"
                                       Foreground="#64748b" Margin="0,0,0,12"/>
                            <StackPanel x:Name="panelStats"/>
                        </StackPanel>
                    </Border>

                    <!-- Filter Section -->
                    <Border Background="#242938" CornerRadius="10" Padding="20" Margin="0,0,0,14"
                            BorderBrush="#333a4f" BorderThickness="1">
                        <StackPanel>
                            <TextBlock Text="FILTRER PAR TYPE" FontSize="10" FontWeight="Bold"
                                       Foreground="#64748b" Margin="0,0,0,10"/>
                            <ComboBox x:Name="cboFilter" Style="{StaticResource ModernComboBox}" Margin="0,0,0,12"/>

                            <TextBlock Text="NOM COMPLET" FontSize="10" FontWeight="Bold"
                                       Foreground="#64748b" Margin="0,0,0,6"/>
                            <TextBox x:Name="txtFilterName" Style="{StaticResource ModernTextBox}"
                                     ToolTip="Filtrer par nom..." Margin="0,0,0,10"/>

                            <TextBlock Text="DÉPARTEMENT" FontSize="10" FontWeight="Bold"
                                       Foreground="#64748b" Margin="0,0,0,6"/>
                            <TextBox x:Name="txtFilterDept" Style="{StaticResource ModernTextBox}"
                                     ToolTip="Filtrer par département..." Margin="0,0,0,10"/>

                            <TextBlock Text="TITRE" FontSize="10" FontWeight="Bold"
                                       Foreground="#64748b" Margin="0,0,0,6"/>
                            <TextBox x:Name="txtFilterTitle" Style="{StaticResource ModernTextBox}"
                                     ToolTip="Filtrer par titre..." Margin="0,0,0,12"/>

                            <TextBlock Text="TYPE DE COMPTE" FontSize="10" FontWeight="Bold"
                                       Foreground="#64748b" Margin="0,0,0,10"/>
                            <ComboBox x:Name="cboUserType" Style="{StaticResource ModernComboBox}"/>

                            <TextBlock Text="STATUT DU COMPTE" FontSize="10" FontWeight="Bold"
                                       Foreground="#64748b" Margin="0,12,0,10"/>
                            <ComboBox x:Name="cboAccountStatus" Style="{StaticResource ModernComboBox}"/>
                        </StackPanel>
                    </Border>

                    <!-- Manage Categories -->
                    <Border Background="#242938" CornerRadius="10" Padding="20" Margin="0,0,0,14"
                            BorderBrush="#333a4f" BorderThickness="1">
                        <StackPanel>
                            <TextBlock Text="GÉRER LES CATÉGORIES" FontSize="10" FontWeight="Bold"
                                       Foreground="#64748b" Margin="0,0,0,10"/>
                            <Grid Margin="0,0,0,8">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="txtNewCategory" Grid.Column="0" Style="{StaticResource ModernTextBox}"
                                         Margin="0,0,8,0" ToolTip="Nouvelle catégorie..."/>
                                <Button x:Name="btnAddCategory" Grid.Column="1" Content="+"
                                        Style="{StaticResource SuccessButton}" Padding="12,8"
                                        FontSize="16" FontWeight="Bold"/>
                            </Grid>
                            <ItemsControl x:Name="listCategories"/>
                        </StackPanel>
                    </Border>

                </StackPanel>
            </ScrollViewer>

            <!-- ========== RIGHT PANEL - User List & Actions ========== -->
            <Grid Grid.Column="2">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Action Bar -->
                <Border Grid.Row="0" Background="#242938" CornerRadius="10" Padding="16" Margin="0,0,0,14"
                        BorderBrush="#333a4f" BorderThickness="1">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBlock Text="Assigner à :" Foreground="#94a3b8" FontSize="13"
                                       VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <ComboBox x:Name="cboAssignType" Style="{StaticResource ModernComboBox}"
                                      Width="200" Margin="0,0,10,0"/>
                        </StackPanel>

                        <Button Grid.Column="2" x:Name="btnAssignSelected"
                                Style="{StaticResource ModernButton}" Margin="0,0,8,0" Padding="16,10" IsEnabled="False">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="&#xE73E;" FontFamily="Segoe MDL2 Assets" FontSize="13"
                                           VerticalAlignment="Center" Margin="0,0,6,0"/>
                                <TextBlock Text="Assigner la sélection" VerticalAlignment="Center"/>
                            </StackPanel>
                        </Button>
                        <Button Grid.Column="3" x:Name="btnImportCSV"
                                Style="{StaticResource SecondaryButton}" Margin="0,0,8,0" Padding="16,10" IsEnabled="False">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="&#xE8B5;" FontFamily="Segoe MDL2 Assets" FontSize="13"
                                           VerticalAlignment="Center" Margin="0,0,6,0"/>
                                <TextBlock Text="Importer CSV" VerticalAlignment="Center"/>
                            </StackPanel>
                        </Button>
                        <Button Grid.Column="4" x:Name="btnClearType"
                                Style="{StaticResource DangerButton}" Padding="16,10" IsEnabled="False">
                            <StackPanel Orientation="Horizontal">
                                <TextBlock Text="&#xE74D;" FontFamily="Segoe MDL2 Assets" FontSize="13"
                                           VerticalAlignment="Center" Margin="0,0,6,0"/>
                                <TextBlock Text="Vider le type" VerticalAlignment="Center"/>
                            </StackPanel>
                        </Button>
                    </Grid>
                </Border>

                <!-- DataGrid -->
                <Border Grid.Row="1" Background="#242938" CornerRadius="10" Padding="2"
                        BorderBrush="#333a4f" BorderThickness="1">
                    <DataGrid x:Name="dgUsers"
                              AutoGenerateColumns="False"
                              IsReadOnly="True"
                              SelectionMode="Extended"
                              HeadersVisibility="Column"
                              GridLinesVisibility="Horizontal"
                              HorizontalGridLinesBrush="#2d3348"
                              Background="#242938"
                              Foreground="#e2e8f0"
                              BorderThickness="0"
                              RowBackground="#242938"
                              AlternatingRowBackground="#272d3e"
                              FontSize="12.5"
                              CanUserSortColumns="True"
                              CanUserResizeColumns="True"
                              ColumnHeaderHeight="40"
                              RowHeight="36"
                              SelectionUnit="FullRow">
                        <DataGrid.ColumnHeaderStyle>
                            <Style TargetType="DataGridColumnHeader">
                                <Setter Property="Background" Value="#1b1f2b"/>
                                <Setter Property="Foreground" Value="#94a3b8"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                                <Setter Property="FontSize" Value="11"/>
                                <Setter Property="Padding" Value="12,8"/>
                                <Setter Property="BorderBrush" Value="#333a4f"/>
                                <Setter Property="BorderThickness" Value="0,0,0,1"/>
                            </Style>
                        </DataGrid.ColumnHeaderStyle>
                        <DataGrid.RowStyle>
                            <Style TargetType="DataGridRow">
                                <Setter Property="BorderBrush" Value="Transparent"/>
                                <Setter Property="BorderThickness" Value="3,0,0,0"/>
                                <Style.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter Property="Background" Value="#1e3a5f"/>
                                        <Setter Property="BorderBrush" Value="#4f8cff"/>
                                        <Setter Property="Foreground" Value="#ffffff"/>
                                    </Trigger>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter Property="Background" Value="#2d3348"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.RowStyle>
                        <DataGrid.CellStyle>
                            <Style TargetType="DataGridCell">
                                <Setter Property="BorderThickness" Value="0"/>
                                <Setter Property="Padding" Value="12,6"/>
                                <Setter Property="Foreground" Value="#e2e8f0"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="DataGridCell">
                                            <Border Padding="{TemplateBinding Padding}"
                                                    Background="{TemplateBinding Background}">
                                                <ContentPresenter VerticalAlignment="Center"/>
                                            </Border>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                                <Style.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter Property="Foreground" Value="#ffffff"/>
                                        <Setter Property="Background" Value="Transparent"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.CellStyle>
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="NOM COMPLET" Binding="{Binding DisplayName}" Width="200"/>
                            <DataGridTextColumn Header="UPN" Binding="{Binding UPN}" Width="250"/>
                            <DataGridTextColumn Header="DÉPARTEMENT" Binding="{Binding Department}" Width="140"/>
                            <DataGridTextColumn Header="TITRE" Binding="{Binding JobTitle}" Width="140"/>
                            <DataGridTextColumn Header="EMPLOYEE TYPE" Binding="{Binding EmployeeType}" Width="140"/>
                            <DataGridTextColumn Header="TYPE COMPTE" Binding="{Binding UserType}" Width="100"/>
                            <DataGridTextColumn Header="ACTIVÉ" Binding="{Binding AccountEnabled}" Width="80"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Border>

                <!-- Selection Info Bar -->
                <Border Grid.Row="2" Background="#242938" CornerRadius="10" Padding="14" Margin="0,10,0,0"
                        BorderBrush="#333a4f" BorderThickness="1">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock x:Name="txtSelectionInfo" Grid.Column="0"
                                   Text="0 utilisateur(s) sélectionné(s)  |  0 utilisateur(s) affiché(s)"
                                   Foreground="#64748b" FontSize="12" VerticalAlignment="Center"/>
                        <StackPanel Grid.Column="1" Orientation="Horizontal">
                            <Button x:Name="btnSelectAll" Content="Tout sélectionner"
                                    Style="{StaticResource SecondaryButton}" Padding="12,6" FontSize="11" Margin="0,0,8,0"/>
                            <Button x:Name="btnSelectNone" Content="Désélectionner tout"
                                    Style="{StaticResource SecondaryButton}" Padding="12,6" FontSize="11"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </Grid>
        </Grid>

        <!-- ==================== STATUS BAR ==================== -->
        <Border Grid.Row="2" Background="#242938" BorderBrush="#333a4f" BorderThickness="0,1,0,0" Padding="24,8">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="txtStatus" Grid.Column="0" Text="Non connecté. Cliquez sur 'Connexion Graph' pour démarrer."
                           Foreground="#64748b" FontSize="11" VerticalAlignment="Center"/>
                <TextBlock x:Name="txtTenant" Grid.Column="1" Text=""
                           Foreground="#4f8cff" FontSize="11" VerticalAlignment="Center"/>
            </Grid>
        </Border>

    </Grid>
</Window>
"@

# ============================================================
# LOAD XAML & BUILD WINDOW
# ============================================================
$reader = (New-Object System.Xml.XmlNodeReader $XAML)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Map all named controls
$XAML.SelectNodes("//*[@*[contains(translate(name(),'x','X'),'Name')]]") | ForEach-Object {
    $name = $_.Name
    Set-Variable -Name $name -Value $window.FindName($name) -Scope Script
}

# ============================================================
# GLOBAL STATE
# ============================================================
$script:AllUsers = @()
$script:FilteredUsers = @()
$script:IsConnected = $false
$script:ExternalSession = $false
$script:Categories = [System.Collections.ObjectModel.ObservableCollection[string]]::new()

# Default categories
@("Permanent", "Consultant", "Syndiqué", "Compte_Service", "VIP") | ForEach-Object {
    $script:Categories.Add($_)
}

# Category color mapping
$script:CategoryColors = @{
    "Permanent"      = "#4f8cff"
    "Consultant"     = "#a78bfa"
    "Syndiqué"       = "#34d399"
    "Compte_Service" = "#fb923c"
    "VIP"            = "#f472b6"
}

$script:DefaultColor = "#64748b"

function Get-CategoryColor([string]$cat) {
    if ($script:CategoryColors.ContainsKey($cat)) { return $script:CategoryColors[$cat] }
    $hash = 0
    $cat.ToCharArray() | ForEach-Object { $hash = $hash * 31 + [int]$_ }
    $colors = @("#38bdf8", "#facc15", "#f87171", "#2dd4bf", "#c084fc", "#fb7185", "#a3e635")
    return $colors[[Math]::Abs($hash) % $colors.Count]
}

# ============================================================
# MICROSOFT GRAPH CONNECTION
# ============================================================
function Connect-ToGraph {
    $txtStatus.Text = "Connexion à Microsoft Graph en cours..."
    $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Render)

    try {
        # Vérifier si une session Graph est déjà active (ex: lancé depuis M365 Monster)
        $context = Get-MgContext
        if ($null -eq $context) {
            # Pas de session — connexion interactive
            Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.Read.All" -ErrorAction Stop
            $context = Get-MgContext
            $script:ExternalSession = $false
        }
        else {
            # Session existante réutilisée — on ne déconnectera pas à la fermeture
            $script:ExternalSession = $true
        }

        $tenantId = $context.TenantId

        # Try to get tenant display name
        $tenantName = $tenantId
        try {
            $org = Get-MgOrganization -ErrorAction Stop
            if ($org.DisplayName) { $tenantName = $org.DisplayName }
        } catch {
            # Fallback to tenant ID if org read fails
        }

        $txtTenant.Text = "Tenant : $tenantName"
        $script:IsConnected = $true

        # Enable buttons
        $btnRefresh.IsEnabled = $true
        $btnExportReport.IsEnabled = $true
        $btnAssignSelected.IsEnabled = $true
        $btnImportCSV.IsEnabled = $true
        $btnClearType.IsEnabled = $true

        # Update connect button
        $btnConnect.Content = "✓  Connecté"
        $btnConnect.IsEnabled = $false

        $txtStatus.Text = "Connecté à Microsoft Graph - Tenant : $tenantName"
        return $true
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Erreur de connexion à Microsoft Graph :`n$($_.Exception.Message)`n`nVérifiez que :`n- Le module Microsoft.Graph.Users est installé`n- Vous avez les permissions User.ReadWrite.All`n- Vous êtes authentifié avec un compte administrateur",
            "Erreur de connexion", "OK", "Error")
        $txtStatus.Text = "Erreur de connexion à Microsoft Graph."
        return $false
    }
}

# ============================================================
# ENTRA ID DATA RETRIEVAL
# ============================================================
function Get-EntraUserData {
    if (-not $script:IsConnected) {
        [System.Windows.MessageBox]::Show("Veuillez d'abord vous connecter à Microsoft Graph.", "Attention", "OK", "Warning")
        return
    }

    $txtStatus.Text = "Chargement des utilisateurs Entra ID (pagination automatique)..."
    $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Render)

    try {
        # Retrieve all users with pagination via -All
        # Properties: Id, DisplayName, UserPrincipalName, Department, JobTitle, EmployeeType, AccountEnabled, UserType
        $mgUsers = Get-MgUser -All -Property Id, DisplayName, UserPrincipalName, Mail, Department, JobTitle, EmployeeType, AccountEnabled, UserType -ConsistencyLevel eventual -CountVariable totalCount -ErrorAction Stop

        $entraUsers = $mgUsers | Select-Object `
            @{N='EntraId';E={$_.Id}},
            @{N='DisplayName';E={if($_.DisplayName){$_.DisplayName}else{'(Sans nom)'}}},
            @{N='UPN';E={$_.UserPrincipalName}},
            @{N='Mail';E={$_.Mail}},
            Department,
            JobTitle,
            EmployeeType,
            @{N='AccountEnabled';E={if($_.AccountEnabled){'Oui'}else{'Non'}}},
            @{N='UserType';E={$_.UserType}}

        $script:AllUsers = @($entraUsers)

        # Discover existing EmployeeType values from Entra
        $existingTypes = $script:AllUsers |
            Where-Object { $_.EmployeeType -and $_.EmployeeType -ne '' } |
            Select-Object -ExpandProperty EmployeeType -Unique
        foreach ($type in $existingTypes) {
            if (-not $script:Categories.Contains($type)) {
                $script:Categories.Add($type)
            }
        }

        $txtStatus.Text = "Chargé : $($script:AllUsers.Count) utilisateur(s) depuis Entra ID"
    }
    catch {
        [System.Windows.MessageBox]::Show(
            "Erreur lors de la récupération des utilisateurs :`n$($_.Exception.Message)`n`nVérifiez vos permissions Graph API.",
            "Erreur Entra ID", "OK", "Error")
        $txtStatus.Text = "Erreur de chargement des utilisateurs."
    }
}

# ============================================================
# UI UPDATE FUNCTIONS
# ============================================================
function Update-Dashboard {
    # Dashboard uses scope filters (UserType + AccountStatus) to define the working set
    # It always shows the full category breakdown within that scope
    $userTypeFilter = $cboUserType.SelectedItem
    $accountStatusFilter = $cboAccountStatus.SelectedItem

    $scopeUsers = $script:AllUsers

    if ($userTypeFilter -eq "Member") {
        $scopeUsers = @($scopeUsers | Where-Object { $_.UserType -eq 'Member' })
    } elseif ($userTypeFilter -eq "Guest") {
        $scopeUsers = @($scopeUsers | Where-Object { $_.UserType -eq 'Guest' })
    }

    if ($accountStatusFilter -eq "Actif") {
        $scopeUsers = @($scopeUsers | Where-Object { $_.AccountEnabled -eq 'Oui' })
    } elseif ($accountStatusFilter -eq "Inactif") {
        $scopeUsers = @($scopeUsers | Where-Object { $_.AccountEnabled -eq 'Non' })
    }

    $total = $scopeUsers.Count
    $uncategorized = @($scopeUsers | Where-Object { -not $_.EmployeeType -or $_.EmployeeType -eq '' }).Count
    $categorized = $total - $uncategorized
    $percent = if ($total -gt 0) { [math]::Round(($categorized / $total) * 100, 1) } else { 0 }

    $txtUncategorized.Text = $uncategorized.ToString()
    $txtTotalUsers.Text = "/ $total"
    $txtProgressPercent.Text = "$percent% catégorisés"

    # Update progress bar
    $maxWidth = 240
    $progressBar.Width = if ($total -gt 0) { [math]::Round(($categorized / $total) * $maxWidth) } else { 0 }

    # Color the uncategorized count
    if ($uncategorized -eq 0) {
        $txtUncategorized.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#34d399")
        $progressBar.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#34d399")
    } elseif ($percent -ge 80) {
        $txtUncategorized.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#facc15")
        $progressBar.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#facc15")
    } else {
        $txtUncategorized.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#fb923c")
        $progressBar.Background = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#fb923c")
    }

    # Stats by category
    $panelStats.Children.Clear()

    # Uncategorized stat
    $statBlock = New-Object System.Windows.Controls.Grid
    $col1 = New-Object System.Windows.Controls.ColumnDefinition; $col1.Width = "Auto"
    $col2 = New-Object System.Windows.Controls.ColumnDefinition; $col2.Width = "*"
    $col3 = New-Object System.Windows.Controls.ColumnDefinition; $col3.Width = "Auto"
    $statBlock.ColumnDefinitions.Add($col1)
    $statBlock.ColumnDefinitions.Add($col2)
    $statBlock.ColumnDefinitions.Add($col3)
    $statBlock.Margin = [System.Windows.Thickness]::new(0,0,0,8)

    $dot = New-Object System.Windows.Shapes.Ellipse
    $dot.Width = 10; $dot.Height = 10
    $dot.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#fb923c")
    $dot.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    $dot.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($dot, 0)
    $statBlock.Children.Add($dot) | Out-Null

    $label = New-Object System.Windows.Controls.TextBlock
    $label.Text = "(Non catégorisé)"
    $label.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#94a3b8")
    $label.FontSize = 12; $label.FontStyle = "Italic"
    $label.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($label, 1)
    $statBlock.Children.Add($label) | Out-Null

    $count = New-Object System.Windows.Controls.TextBlock
    $count.Text = $uncategorized.ToString()
    $count.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#e2e8f0")
    $count.FontSize = 13; $count.FontWeight = "SemiBold"
    $count.VerticalAlignment = "Center"
    [System.Windows.Controls.Grid]::SetColumn($count, 2)
    $statBlock.Children.Add($count) | Out-Null

    $panelStats.Children.Add($statBlock) | Out-Null

    # Each category stat
    foreach ($cat in $script:Categories) {
        $catCount = @($scopeUsers | Where-Object { $_.EmployeeType -eq $cat }).Count

        $statBlock = New-Object System.Windows.Controls.Grid
        $col1 = New-Object System.Windows.Controls.ColumnDefinition; $col1.Width = "Auto"
        $col2 = New-Object System.Windows.Controls.ColumnDefinition; $col2.Width = "*"
        $col3 = New-Object System.Windows.Controls.ColumnDefinition; $col3.Width = "Auto"
        $statBlock.ColumnDefinitions.Add($col1)
        $statBlock.ColumnDefinitions.Add($col2)
        $statBlock.ColumnDefinitions.Add($col3)
        $statBlock.Margin = [System.Windows.Thickness]::new(0,0,0,6)

        $dot = New-Object System.Windows.Shapes.Ellipse
        $dot.Width = 10; $dot.Height = 10
        $dot.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFrom((Get-CategoryColor $cat))
        $dot.Margin = [System.Windows.Thickness]::new(0,0,8,0)
        $dot.VerticalAlignment = "Center"
        [System.Windows.Controls.Grid]::SetColumn($dot, 0)
        $statBlock.Children.Add($dot) | Out-Null

        $label = New-Object System.Windows.Controls.TextBlock
        $label.Text = $cat
        $label.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#94a3b8")
        $label.FontSize = 12
        $label.VerticalAlignment = "Center"
        [System.Windows.Controls.Grid]::SetColumn($label, 1)
        $statBlock.Children.Add($label) | Out-Null

        $countTxt = New-Object System.Windows.Controls.TextBlock
        $countTxt.Text = $catCount.ToString()
        $countTxt.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#e2e8f0")
        $countTxt.FontSize = 13; $countTxt.FontWeight = "SemiBold"
        $countTxt.VerticalAlignment = "Center"
        [System.Windows.Controls.Grid]::SetColumn($countTxt, 2)
        $statBlock.Children.Add($countTxt) | Out-Null

        $panelStats.Children.Add($statBlock) | Out-Null
    }
}

function Update-CategoryLists {
    # Filter combobox
    $selectedFilter = $cboFilter.SelectedItem
    $cboFilter.Items.Clear()
    $cboFilter.Items.Add("-- Tous les utilisateurs --") | Out-Null
    $cboFilter.Items.Add("-- Non catégorisés --") | Out-Null
    foreach ($cat in $script:Categories) {
        $cboFilter.Items.Add($cat) | Out-Null
    }
    if ($selectedFilter -and $cboFilter.Items.Contains($selectedFilter)) {
        $cboFilter.SelectedItem = $selectedFilter
    } else {
        $cboFilter.SelectedIndex = 0
    }

    # Assign combobox
    $selectedAssign = $cboAssignType.SelectedItem
    $cboAssignType.Items.Clear()
    foreach ($cat in $script:Categories) {
        $cboAssignType.Items.Add($cat) | Out-Null
    }
    if ($selectedAssign -and $cboAssignType.Items.Contains($selectedAssign)) {
        $cboAssignType.SelectedItem = $selectedAssign
    } elseif ($cboAssignType.Items.Count -gt 0) {
        $cboAssignType.SelectedIndex = 0
    }

    # Category management list
    Update-CategoryManagementList
}

function Update-CategoryManagementList {
    $listCategories.Items.Clear()

    foreach ($cat in $script:Categories) {
        $sp = New-Object System.Windows.Controls.Grid
        $c1 = New-Object System.Windows.Controls.ColumnDefinition; $c1.Width = "Auto"
        $c2 = New-Object System.Windows.Controls.ColumnDefinition; $c2.Width = "*"
        $c3 = New-Object System.Windows.Controls.ColumnDefinition; $c3.Width = "Auto"
        $sp.ColumnDefinitions.Add($c1)
        $sp.ColumnDefinitions.Add($c2)
        $sp.ColumnDefinitions.Add($c3)
        $sp.Margin = [System.Windows.Thickness]::new(0,3,0,3)

        $dot = New-Object System.Windows.Shapes.Ellipse
        $dot.Width = 8; $dot.Height = 8
        $dot.Fill = [System.Windows.Media.BrushConverter]::new().ConvertFrom((Get-CategoryColor $cat))
        $dot.Margin = [System.Windows.Thickness]::new(0,0,8,0)
        $dot.VerticalAlignment = "Center"
        [System.Windows.Controls.Grid]::SetColumn($dot, 0)
        $sp.Children.Add($dot) | Out-Null

        $lbl = New-Object System.Windows.Controls.TextBlock
        $lbl.Text = $cat
        $lbl.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#e2e8f0")
        $lbl.FontSize = 12
        $lbl.VerticalAlignment = "Center"
        [System.Windows.Controls.Grid]::SetColumn($lbl, 1)
        $sp.Children.Add($lbl) | Out-Null

        $btnDel = New-Object System.Windows.Controls.Button
        $btnDel.Content = "✕"
        $btnDel.FontSize = 10
        $btnDel.Padding = [System.Windows.Thickness]::new(6,2,6,2)
        $btnDel.Background = [System.Windows.Media.Brushes]::Transparent
        $btnDel.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#64748b")
        $btnDel.BorderThickness = [System.Windows.Thickness]::new(0)
        $btnDel.Cursor = [System.Windows.Input.Cursors]::Hand
        $btnDel.Tag = $cat
        $btnDel.ToolTip = "Supprimer la catégorie '$cat'"
        [System.Windows.Controls.Grid]::SetColumn($btnDel, 2)
        $btnDel.Add_Click({
            param($clickSender, $e)
            $catToRemove = $clickSender.Tag
            $usersWithCat = @($script:AllUsers | Where-Object { $_.EmployeeType -eq $catToRemove }).Count
            $msg = "Supprimer la catégorie '$catToRemove' ?"
            if ($usersWithCat -gt 0) {
                $msg += "`n`nAttention : $usersWithCat utilisateur(s) ont actuellement ce type.`nLeur champ EmployeeType ne sera PAS modifié dans Entra ID."
            }
            $result = [System.Windows.MessageBox]::Show($msg, "Confirmer", "YesNo", "Question")
            if ($result -eq "Yes") {
                $script:Categories.Remove($catToRemove) | Out-Null
                if ($script:CategoryColors.ContainsKey($catToRemove)) {
                    $script:CategoryColors.Remove($catToRemove)
                }
                Update-CategoryLists
                Update-UserGrid
                Update-Dashboard
                $txtStatus.Text = "Catégorie '$catToRemove' supprimée de la liste."
            }
        })
        $sp.Children.Add($btnDel) | Out-Null

        $listCategories.Items.Add($sp) | Out-Null
    }
}

function Update-UserGrid {
    $filterType = $cboFilter.SelectedItem
    $filterName = $txtFilterName.Text.Trim().ToLower()
    $filterDept = $txtFilterDept.Text.Trim().ToLower()
    $filterTitle = $txtFilterTitle.Text.Trim().ToLower()
    $userTypeFilter = $cboUserType.SelectedItem
    $accountStatusFilter = $cboAccountStatus.SelectedItem

    $filtered = $script:AllUsers

    # Apply employee type filter
    if ($filterType -eq "-- Non catégorisés --") {
        $filtered = @($filtered | Where-Object { -not $_.EmployeeType -or $_.EmployeeType -eq '' })
    }
    elseif ($filterType -and $filterType -ne "-- Tous les utilisateurs --") {
        $filtered = @($filtered | Where-Object { $_.EmployeeType -eq $filterType })
    }

    # Apply UserType filter (Member / Guest)
    if ($userTypeFilter -eq "Member") {
        $filtered = @($filtered | Where-Object { $_.UserType -eq 'Member' })
    }
    elseif ($userTypeFilter -eq "Guest") {
        $filtered = @($filtered | Where-Object { $_.UserType -eq 'Guest' })
    }

    # Apply Account Status filter (Actif / Inactif)
    if ($accountStatusFilter -eq "Actif") {
        $filtered = @($filtered | Where-Object { $_.AccountEnabled -eq 'Oui' })
    }
    elseif ($accountStatusFilter -eq "Inactif") {
        $filtered = @($filtered | Where-Object { $_.AccountEnabled -eq 'Non' })
    }

    # Apply column text filters
    if ($filterName) {
        $filtered = @($filtered | Where-Object {
            ($_.DisplayName -and $_.DisplayName.ToLower().Contains($filterName)) -or
            ($_.UPN -and $_.UPN.ToLower().Contains($filterName)) -or
            ($_.Mail -and $_.Mail.ToLower().Contains($filterName))
        })
    }
    if ($filterDept) {
        $filtered = @($filtered | Where-Object {
            $_.Department -and $_.Department.ToLower().Contains($filterDept)
        })
    }
    if ($filterTitle) {
        $filtered = @($filtered | Where-Object {
            $_.JobTitle -and $_.JobTitle.ToLower().Contains($filterTitle)
        })
    }

    $script:FilteredUsers = $filtered
    $dgUsers.ItemsSource = $filtered
    Update-SelectionInfo
}

function Update-SelectionInfo {
    $selected = $dgUsers.SelectedItems.Count
    $displayed = $script:FilteredUsers.Count
    $txtSelectionInfo.Text = "$selected utilisateur(s) sélectionné(s)  |  $displayed utilisateur(s) affiché(s)"
}

# ============================================================
# ENTRA ID WRITE OPERATIONS (Microsoft Graph)
# ============================================================
function Set-EmployeeTypes {
    param(
        [array]$Users,
        [string]$EmployeeType
    )

    $successCount = 0
    $errorCount = 0
    $errors = @()

    foreach ($user in $Users) {
        try {
            $userId = $user.EntraId

            if ($EmployeeType -eq '') {
                # Clear the EmployeeType by setting to null
                Update-MgUser -UserId $userId -EmployeeType $null -ErrorAction Stop
            } else {
                Update-MgUser -UserId $userId -EmployeeType $EmployeeType -ErrorAction Stop
            }
            $successCount++

            # Update local cache
            $idx = [Array]::FindIndex($script:AllUsers, [Predicate[object]]{ param($u) $u.EntraId -eq $user.EntraId })
            if ($idx -ge 0) {
                $script:AllUsers[$idx].EmployeeType = if ($EmployeeType -eq '') { $null } else { $EmployeeType }
            }
        }
        catch {
            $errorCount++
            $errors += "$($user.UPN): $($_.Exception.Message)"
        }
    }

    return @{
        Success = $successCount
        Errors  = $errorCount
        Details = $errors
    }
}

# ============================================================
# EVENT HANDLERS
# ============================================================

# Connect to Graph
$btnConnect.Add_Click({
    $connected = Connect-ToGraph
    if ($connected) {
        Get-EntraUserData
        Update-CategoryLists
        Update-UserGrid
        Update-Dashboard
    }
})

# Refresh
$btnRefresh.Add_Click({
    Get-EntraUserData
    Update-CategoryLists
    Update-UserGrid
    Update-Dashboard
})

# Filter change
$cboFilter.Add_SelectionChanged({
    Update-UserGrid
})

# UserType filter change
$cboUserType.Add_SelectionChanged({
    Update-Dashboard
    Update-UserGrid
})

# AccountStatus filter change
$cboAccountStatus.Add_SelectionChanged({
    Update-Dashboard
    Update-UserGrid
})

# Column text filters
$txtFilterName.Add_TextChanged({
    Update-UserGrid
})

$txtFilterDept.Add_TextChanged({
    Update-UserGrid
})

$txtFilterTitle.Add_TextChanged({
    Update-UserGrid
})

# Selection changed in DataGrid
$dgUsers.Add_SelectionChanged({
    Update-SelectionInfo
})

# Select All
$btnSelectAll.Add_Click({
    $dgUsers.SelectAll()
    Update-SelectionInfo
})

# Select None
$btnSelectNone.Add_Click({
    $dgUsers.UnselectAll()
    Update-SelectionInfo
})

# Add Category
$btnAddCategory.Add_Click({
    $newCat = $txtNewCategory.Text.Trim()
    if (-not $newCat) {
        [System.Windows.MessageBox]::Show("Veuillez saisir un nom de catégorie.", "Attention", "OK", "Warning")
        return
    }
    if ($script:Categories.Contains($newCat)) {
        [System.Windows.MessageBox]::Show("La catégorie '$newCat' existe déjà.", "Attention", "OK", "Warning")
        return
    }
    $script:Categories.Add($newCat)
    $txtNewCategory.Text = ""
    Update-CategoryLists
    Update-UserGrid
    Update-Dashboard
    $txtStatus.Text = "Catégorie '$newCat' ajoutée."
})

# Enter key in new category textbox
$txtNewCategory.Add_KeyDown({
    param($clickSender, $e)
    if ($e.Key -eq "Return") {
        $btnAddCategory.RaiseEvent((New-Object System.Windows.RoutedEventArgs([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)))
    }
})

# Assign Selected
$btnAssignSelected.Add_Click({
    $selectedUsers = @($dgUsers.SelectedItems)
    $targetType = $cboAssignType.SelectedItem

    if ($selectedUsers.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Veuillez sélectionner au moins un utilisateur.", "Attention", "OK", "Warning")
        return
    }
    if (-not $targetType) {
        [System.Windows.MessageBox]::Show("Veuillez sélectionner un type à assigner.", "Attention", "OK", "Warning")
        return
    }

    $confirm = [System.Windows.MessageBox]::Show(
        "Assigner le type '$targetType' à $($selectedUsers.Count) utilisateur(s) ?",
        "Confirmer l'assignation", "YesNo", "Question")

    if ($confirm -eq "Yes") {
        $txtStatus.Text = "Assignation en cours via Microsoft Graph..."
        $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Render)

        $result = Set-EmployeeTypes -Users $selectedUsers -EmployeeType $targetType

        Update-UserGrid
        Update-Dashboard

        $msg = "Terminé :`n- $($result.Success) succès`n- $($result.Errors) erreur(s)"
        if ($result.Errors -gt 0) {
            $msg += "`n`nDétails des erreurs :`n$($result.Details -join "`n")"
        }
        [System.Windows.MessageBox]::Show($msg, "Résultat", "OK", $(if($result.Errors -gt 0){"Warning"}else{"Information"}))
        $txtStatus.Text = "Assignation terminée : $($result.Success) succès, $($result.Errors) erreur(s)"
    }
})

# Clear Type
$btnClearType.Add_Click({
    $selectedUsers = @($dgUsers.SelectedItems)
    if ($selectedUsers.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Veuillez sélectionner au moins un utilisateur.", "Attention", "OK", "Warning")
        return
    }

    $confirm = [System.Windows.MessageBox]::Show(
        "Vider le champ Employee Type pour $($selectedUsers.Count) utilisateur(s) ?`n`nCette action remettra le champ à vide dans Entra ID.",
        "Confirmer", "YesNo", "Warning")

    if ($confirm -eq "Yes") {
        $txtStatus.Text = "Suppression du type en cours via Microsoft Graph..."
        $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Render)

        $result = Set-EmployeeTypes -Users $selectedUsers -EmployeeType ''

        Update-UserGrid
        Update-Dashboard

        $msg = "Terminé :`n- $($result.Success) succès`n- $($result.Errors) erreur(s)"
        [System.Windows.MessageBox]::Show($msg, "Résultat", "OK", "Information")
        $txtStatus.Text = "Champ vidé pour $($result.Success) utilisateur(s)"
    }
})

# Import CSV
$btnImportCSV.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "Importer un fichier CSV"
    $dialog.Filter = "Fichiers CSV (*.csv)|*.csv|Tous les fichiers (*.*)|*.*"
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($dialog.ShowDialog() -eq "OK") {
        $txtStatus.Text = "Import CSV en cours..."
        $window.Dispatcher.Invoke([action]{}, [System.Windows.Threading.DispatcherPriority]::Render)

        try {
            $csvData = Import-Csv -Path $dialog.FileName -Encoding UTF8

            # Detect column names (flexible) - supports UPN or ObjectId as identifier
            $idCol = $null
            $typeCol = $null

            $headers = $csvData[0].PSObject.Properties.Name
            foreach ($h in $headers) {
                $hLower = $h.ToLower().Trim()
                if ($hLower -in @('userprincipalname','upn','email','mail','username','login','identifiant','utilisateur','objectid','id','entraid')) { $idCol = $h }
                if ($hLower -in @('employeetype','employee_type','type','categorie','catégorie','category')) { $typeCol = $h }
            }

            if (-not $idCol -or -not $typeCol) {
                $colList = $headers -join ", "
                [System.Windows.MessageBox]::Show(
                    "Colonnes requises introuvables dans le CSV.`n`nColonnes détectées : $colList`n`nLe CSV doit contenir :`n- Une colonne identifiant (ex: UserPrincipalName, UPN, Email, ObjectId)`n- Une colonne pour le type (ex: EmployeeType, Type, Categorie)",
                    "Erreur CSV", "OK", "Error")
                $txtStatus.Text = "Import CSV échoué : colonnes introuvables."
                return
            }

            # Preview
            $preview = $csvData | Select-Object -First 5 | ForEach-Object {
                "  $($_.$idCol) → $($_.$typeCol)"
            }
            $previewText = $preview -join "`n"
            $totalLines = $csvData.Count

            $confirm = [System.Windows.MessageBox]::Show(
                "Fichier : $($dialog.FileName)`nColonne identifiant : $idCol`nColonne type : $typeCol`nLignes : $totalLines`n`nAperçu (5 premiers) :`n$previewText`n`nProcéder à l'import ?",
                "Confirmer l'import CSV", "YesNo", "Question")

            if ($confirm -eq "Yes") {
                $successCount = 0
                $errorCount = 0
                $notFound = @()
                $errors = @()

                foreach ($row in $csvData) {
                    $identifier = $row.$idCol.Trim()
                    $type = $row.$typeCol.Trim()

                    if (-not $identifier) { continue }

                    # Add category if new
                    if ($type -and -not $script:Categories.Contains($type)) {
                        $script:Categories.Add($type)
                    }

                    # Find user in local cache by UPN, Mail, or EntraId
                    $matchedUser = $script:AllUsers | Where-Object {
                        $_.UPN -eq $identifier -or
                        $_.Mail -eq $identifier -or
                        $_.EntraId -eq $identifier
                    } | Select-Object -First 1

                    if (-not $matchedUser) {
                        $notFound += $identifier
                        $errorCount++
                        continue
                    }

                    try {
                        if ($type) {
                            Update-MgUser -UserId $matchedUser.EntraId -EmployeeType $type -ErrorAction Stop
                        } else {
                            Update-MgUser -UserId $matchedUser.EntraId -EmployeeType $null -ErrorAction Stop
                        }
                        $successCount++

                        # Update local cache
                        $idx = [Array]::FindIndex($script:AllUsers, [Predicate[object]]{ param($u) $u.EntraId -eq $matchedUser.EntraId })
                        if ($idx -ge 0) {
                            $script:AllUsers[$idx].EmployeeType = if ($type) { $type } else { $null }
                        }
                    }
                    catch {
                        $errorCount++
                        $errors += "$identifier : $($_.Exception.Message)"
                    }
                }

                Update-CategoryLists
                Update-UserGrid
                Update-Dashboard

                $msg = "Import terminé :`n- $successCount succès`n- $errorCount erreur(s)"
                if ($notFound.Count -gt 0) {
                    $msg += "`n`nUtilisateurs introuvables ($($notFound.Count)) :`n$($notFound[0..([Math]::Min(9, $notFound.Count - 1))] -join "`n")"
                    if ($notFound.Count -gt 10) { $msg += "`n... et $($notFound.Count - 10) autres" }
                }
                if ($errors.Count -gt 0) {
                    $msg += "`n`nErreurs Graph API :`n$($errors[0..([Math]::Min(4, $errors.Count - 1))] -join "`n")"
                }

                [System.Windows.MessageBox]::Show($msg, "Résultat Import", "OK",
                    $(if($errorCount -gt 0){"Warning"}else{"Information"}))

                $txtStatus.Text = "Import CSV : $successCount succès, $errorCount erreur(s)"
            }
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Erreur lors de la lecture du CSV :`n$($_.Exception.Message)",
                "Erreur", "OK", "Error")
            $txtStatus.Text = "Erreur de lecture du CSV."
        }
    }
})

# Export Report
$btnExportReport.Add_Click({
    $dialog = New-Object System.Windows.Forms.SaveFileDialog
    $dialog.Title = "Exporter le rapport"
    $dialog.Filter = "Fichiers CSV (*.csv)|*.csv"
    $dialog.FileName = "EntraID_EmployeeType_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $dialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

    if ($dialog.ShowDialog() -eq "OK") {
        try {
            $script:AllUsers | Select-Object EntraId, DisplayName, UPN, Mail, Department, JobTitle, EmployeeType, AccountEnabled, UserType |
                Export-Csv -Path $dialog.FileName -NoTypeInformation -Encoding UTF8

            $total = $script:AllUsers.Count
            $uncategorized = @($script:AllUsers | Where-Object { -not $_.EmployeeType -or $_.EmployeeType -eq '' }).Count

            # Generate summary section
            $summary = @()
            $summary += ""
            $summary += "# ====== RAPPORT EMPLOYEE TYPE - ENTRA ID ======"
            $summary += "# Date : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            $summary += "# Tenant : $($txtTenant.Text)"
            $summary += "# Total utilisateurs : $total"
            $summary += "# Non catégorisés : $uncategorized"
            $summary += "# Catégorisés : $($total - $uncategorized)"
            $summary += "#"
            foreach ($cat in $script:Categories) {
                $catCount = @($script:AllUsers | Where-Object { $_.EmployeeType -eq $cat }).Count
                $summary += "# $cat : $catCount"
            }
            $summary += "# ================================================"

            $summary | Out-File -Append -FilePath $dialog.FileName -Encoding UTF8

            [System.Windows.MessageBox]::Show(
                "Rapport exporté avec succès :`n$($dialog.FileName)",
                "Export", "OK", "Information")

            $txtStatus.Text = "Rapport exporté : $($dialog.FileName)"
        }
        catch {
            [System.Windows.MessageBox]::Show(
                "Erreur lors de l'export :`n$($_.Exception.Message)",
                "Erreur", "OK", "Error")
        }
    }
})

# ============================================================
# INITIALIZE
# ============================================================
$window.Add_Loaded({
    # Populate UserType filter
    $cboUserType.Items.Add("-- Tous --") | Out-Null
    $cboUserType.Items.Add("Member") | Out-Null
    $cboUserType.Items.Add("Guest") | Out-Null
    $cboUserType.SelectedIndex = 0

    # Populate Account Status filter
    $cboAccountStatus.Items.Add("-- Tous --") | Out-Null
    $cboAccountStatus.Items.Add("Actif") | Out-Null
    $cboAccountStatus.Items.Add("Inactif") | Out-Null
    $cboAccountStatus.SelectedIndex = 0

    Update-CategoryLists
})

# Cleanup on close - ne déconnecter que si la session a été créée par ce script
$window.Add_Closing({
    if ($script:IsConnected -and -not $script:ExternalSession) {
        try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
    }
})

# Show the window
$window.ShowDialog() | Out-Null