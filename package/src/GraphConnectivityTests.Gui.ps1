<#
.SYNOPSIS
    Graph Connectivity Tests - graphical interface (dot-sourced by GraphConnectivityTests.psm1).

.DESCRIPTION
    WPF window, same design as the other tools of the author: Fluent theme of Windows 11 with .NET 9 or
    later (PowerShell 7.5+), light or dark like Windows, the crimson accent of the other tools of the author; classic WPF
    controls with the same colours on PowerShell 7.4.

    Layout: a header; the test picker (one card per test of Get-GctTestCatalog, with its description);
    on the left the mailboxes (source, targets, on-premises host, domain, cloud), the sign-in (context,
    client, tenant, certificate, on-premises authentication, sign-in mode) and the options (changes,
    administrator checks, free/busy days); on the right a notice and the progress, with the counts of
    the last run, in two views: the timeline (one node per stage on a rail, its checks under it - a stage
    without problem folds, a click unfolds it - and, for each check, its diagnostics: the details and
    every HTTP exchange, a click showing the request and the response) and the diagnostic log (the same
    in time order, as text, to copy); the Diagnostics switch unfolds everything. At the bottom the actions. The fields a test does not use are
    disabled. No secret is ever typed: the sign-in is interactive, an application uses a certificate.

    It runs exactly the same engine as the command line (Invoke-GctConnectivityTest, Export-GctReport).
    The settings are checked with Test-GctConfiguration first and every problem is listed. The run
    happens on the window thread, kept responsive while waiting (Invoke-GctUiPump), and can be
    cancelled. Close is the cancel button of the window (Esc too); a Closing handler is attached only
    while a test runs. Ctrl+C is ignored in the console while the window is open.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
#>

$script:Gui = $null
$script:GuiIconFont = 'Segoe Fluent Icons, Segoe MDL2 Assets'
$script:GuiTestGlyphs = @{ FreeBusy = 0xE787; MailboxServices = 0xE715; Provisioning = 0xE77B; DomainDns = 0xE774; All = 0xE9D5 }
$script:GuiContexts = @(
    [pscustomobject]@{ Name = 'User'; Text = 'User (source user)' }
    [pscustomobject]@{ Name = 'Application'; Text = 'Application' }
)
$script:GuiClients = @(
    [pscustomobject]@{ Name = 'MicrosoftOffice'; Text = 'Microsoft Office - no application to register' }
    [pscustomobject]@{ Name = 'GraphCommandLineTools'; Text = 'Microsoft Graph Command Line Tools - consent' }
    [pscustomobject]@{ Name = 'Custom'; Text = 'Your application (application ID)' }
)
$script:GuiSignInModes = @(
    [pscustomobject]@{ Name = 'Auto'; Text = 'Auto' }
    [pscustomobject]@{ Name = 'Window'; Text = 'Sign-in window' }
    [pscustomobject]@{ Name = 'DeviceCode'; Text = 'Device code' }
)
$script:GuiOnPremisesAuthentication = @(
    [pscustomobject]@{ Name = 'OAuth'; Text = 'OAuth - hybrid modern authentication' }
    [pscustomobject]@{ Name = 'Windows'; Text = 'Windows - current Windows account' }
    [pscustomobject]@{ Name = 'Credential'; Text = 'Credential - account and password' }
)
$script:GuiCredentialSchemes = @(
    [pscustomobject]@{ Name = 'Auto'; Text = 'Auto - the best scheme EWS offers' }
    [pscustomobject]@{ Name = 'NTLM'; Text = 'NTLM' }
    [pscustomobject]@{ Name = 'Negotiate'; Text = 'Negotiate (Kerberos, else NTLM)' }
    [pscustomobject]@{ Name = 'Basic'; Text = 'Basic' }
)

function Get-GctGuiXaml {
    <# The window. Colours come from the Fluent theme resources (or the classic fallback of Set-GctGuiTheme). #>
    @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Width="1400" Height="940" MinWidth="1120" MinHeight="660" WindowStartupLocation="CenterScreen"
        FontFamily="Segoe UI Variable Text, Segoe UI" FontSize="14" UseLayoutRounding="True">
  <Window.Resources>
    <Style x:Key="GctCard" TargetType="Border">
      <Setter Property="Background" Value="{DynamicResource CardBackgroundFillColorDefaultBrush}"/>
      <Setter Property="BorderBrush" Value="{DynamicResource CardStrokeColorDefaultBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="8"/>
      <Setter Property="Padding" Value="18,14,18,16"/>
      <Setter Property="Margin" Value="0,0,0,12"/>
    </Style>
    <Style x:Key="GctCardTitle" TargetType="TextBlock">
      <Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Margin" Value="0,0,0,4"/>
      <Setter Property="Foreground" Value="{DynamicResource TextFillColorPrimaryBrush}"/>
    </Style>
    <Style x:Key="GctLabel" TargetType="TextBlock">
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="Foreground" Value="{DynamicResource TextFillColorSecondaryBrush}"/>
      <Setter Property="Margin" Value="0,8,0,4"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="GctHint" TargetType="TextBlock">
      <Setter Property="FontSize" Value="11.5"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
      <Setter Property="Foreground" Value="{DynamicResource TextFillColorTertiaryBrush}"/>
      <Setter Property="Margin" Value="0,4,0,0"/>
    </Style>
    <Style x:Key="GctIcon" TargetType="TextBlock">
      <Setter Property="FontFamily" Value="Segoe Fluent Icons, Segoe MDL2 Assets"/>
    </Style>
    <Style x:Key="GctChoice" TargetType="ListBoxItem">
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ListBoxItem">
            <Border x:Name="Card" Margin="0,0,8,0" Padding="10,8,8,10" CornerRadius="6" BorderThickness="2"
                    Background="{DynamicResource ControlFillColorDefaultBrush}" BorderBrush="{DynamicResource ControlStrokeColorDefaultBrush}">
              <ContentPresenter/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="Card" Property="Background" Value="{DynamicResource ControlFillColorSecondaryBrush}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="Card" Property="BorderBrush" Value="{DynamicResource AccentFillColorDefaultBrush}"/>
                <Setter TargetName="Card" Property="Background" Value="{DynamicResource GctAccentSoft}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Grid x:Name="Root" Background="{DynamicResource ApplicationBackgroundBrush}">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Grid x:Name="Header" Margin="24,16,24,12">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="Auto"/>
      </Grid.ColumnDefinitions>
      <Border Width="46" Height="46" CornerRadius="10" Background="{DynamicResource GctBrand}" VerticalAlignment="Center">
        <TextBlock Style="{StaticResource GctIcon}" Text="&#xE787;" FontSize="22" Foreground="White" HorizontalAlignment="Center" VerticalAlignment="Center"/>
      </Border>
      <StackPanel Grid.Column="1" Margin="14,0,0,0" VerticalAlignment="Center">
        <TextBlock Text="MICROSOFT GRAPH CONNECTIVITY TESTS - USER CONTEXT" FontSize="11" FontWeight="SemiBold" Foreground="{DynamicResource GctBrandText}"/>
        <TextBlock Text="Graph Connectivity Tests" FontSize="24" FontWeight="SemiBold" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
        <TextBlock FontSize="13" Foreground="{DynamicResource TextFillColorSecondaryBrush}" TextTrimming="CharacterEllipsis"
                   Text="Choose a test, the mailboxes and how to sign in, then run it: same checks, same HTTP trace and same report as the command line."/>
      </StackPanel>
      <TextBlock x:Name="Version" Grid.Column="2" FontSize="12" Foreground="{DynamicResource TextFillColorSecondaryBrush}" VerticalAlignment="Top"/>
    </Grid>


    <Grid Grid.Row="1" Margin="24,0,24,0">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="3*" MinWidth="700"/>
        <ColumnDefinition Width="16"/>
        <ColumnDefinition Width="2*" MinWidth="360"/>
      </Grid.ColumnDefinitions>
      <ScrollViewer x:Name="SettingsScroll" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled" Padding="0,0,4,0">
        <StackPanel>
          <Border x:Name="TestCard" Style="{StaticResource GctCard}" Padding="18,12,10,14">
            <StackPanel>
              <TextBlock Text="Test" Style="{StaticResource GctCardTitle}" Margin="0,0,0,8"/>
              <ListBox x:Name="TestPicker" ItemContainerStyle="{StaticResource GctChoice}" BorderThickness="0" Background="Transparent" Padding="0"
                       ScrollViewer.HorizontalScrollBarVisibility="Disabled" ScrollViewer.VerticalScrollBarVisibility="Disabled">
                <ListBox.ItemsPanel>
                  <ItemsPanelTemplate><UniformGrid Rows="1"/></ItemsPanelTemplate>
                </ListBox.ItemsPanel>
              </ListBox>
            </StackPanel>
          </Border>
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/>
              <ColumnDefinition Width="12"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <StackPanel>
            <Border x:Name="MailboxCard" Style="{StaticResource GctCard}">
              <StackPanel>
                <TextBlock Text="Mailboxes" Style="{StaticResource GctCardTitle}"/>
                <TextBlock Text="Source mailbox: signs in and asks for the free/busy" Style="{StaticResource GctLabel}"/>
                <TextBox x:Name="SourceMailbox"/>
                <TextBlock x:Name="TargetLabel" Text="Target mailboxes: one per line or a comma list" Style="{StaticResource GctLabel}"/>
                <TextBox x:Name="TargetMailboxes" AcceptsReturn="True" TextWrapping="Wrap" Height="76" VerticalScrollBarVisibility="Auto" VerticalContentAlignment="Top"/>
                <TextBlock x:Name="TargetHint" Style="{StaticResource GctHint}"/>
                <TextBlock x:Name="OnPremisesLabel" Text="On-premises host (optional, e.g. mail.contoso.com)" Style="{StaticResource GctLabel}" ToolTip="FQDN of the on-premises Availability service - empty: the host Autodiscover gives"/>
                <TextBox x:Name="OnPremisesHost"/>
                <CheckBox x:Name="SkipOnPremisesChecks" Margin="0,6,0,0" Content="Skip the on-premises checks"/>
                <TextBlock x:Name="DomainLabel" Text="Domain (empty = the domain of the source)" Style="{StaticResource GctLabel}"/>
                <TextBox x:Name="Domain"/>
                <TextBlock Text="Cloud" Style="{StaticResource GctLabel}"/>
                <ComboBox x:Name="Cloud" HorizontalAlignment="Stretch"/>
              </StackPanel>
            </Border>
              <Border x:Name="OptionsCard" Style="{StaticResource GctCard}">
                <StackPanel>
                  <TextBlock Text="Options" Style="{StaticResource GctCardTitle}"/>
                  <CheckBox x:Name="AllowWrite" Margin="0,8,0,0">
                    <TextBlock Text="Allow changes to the source mailbox (automatic replies)" TextWrapping="Wrap"/>
                  </CheckBox>
                  <TextBlock x:Name="AllowWriteHint" Style="{StaticResource GctHint}" Margin="28,0,0,0"/>
                  <CheckBox x:Name="AdminChecks" Margin="0,10,0,0">
                    <TextBlock Text="Read the hybrid configuration as an Exchange administrator" TextWrapping="Wrap"/>
                  </CheckBox>
                  <TextBlock x:Name="AdminChecksHint" Style="{StaticResource GctHint}" Margin="28,0,0,0"
                             Text="One more sign-in, with an Exchange administrator account: organization relationships, IntraOrganizationConnector, availability address spaces."/>
                  <TextBlock x:Name="FreeBusyDaysLabel" Text="Free/busy period: days from today (1 to 62)" Style="{StaticResource GctLabel}" Margin="0,12,0,4"/>
                  <TextBox x:Name="FreeBusyDays" Width="80" HorizontalAlignment="Left"/>
                </StackPanel>
              </Border>
            </StackPanel>
            <StackPanel Grid.Column="2">
              <Border x:Name="SignInCard" Style="{StaticResource GctCard}">
                <StackPanel>
                  <TextBlock Text="Sign-in" Style="{StaticResource GctCardTitle}"/>
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="12"/>
                      <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel>
                      <TextBlock Text="Context" Style="{StaticResource GctLabel}"/>
                      <ComboBox x:Name="Context" HorizontalAlignment="Stretch"/>
                    </StackPanel>
                    <StackPanel Grid.Column="2">
                      <TextBlock x:Name="SignInLabel" Text="Sign-in mode" Style="{StaticResource GctLabel}"/>
                      <ComboBox x:Name="SignIn" HorizontalAlignment="Stretch"/>
                    </StackPanel>
                  </Grid>
                  <TextBlock x:Name="ClientLabel" Text="Client" Style="{StaticResource GctLabel}"/>
                  <ComboBox x:Name="Client" HorizontalAlignment="Stretch"/>
                  <StackPanel x:Name="CustomPanel">
                    <TextBlock Text="Application ID (GUID)" Style="{StaticResource GctLabel}"/>
                    <TextBox x:Name="AppClientId" FontFamily="Cascadia Mono, Consolas"/>
                    <StackPanel x:Name="RedirectPanel">
                      <TextBlock Text="Redirect URI (empty = native client page)" Style="{StaticResource GctLabel}"/>
                      <TextBox x:Name="RedirectUri"/>
                    </StackPanel>
                  </StackPanel>
                  <TextBlock x:Name="TenantLabel" Text="Tenant (optional: ID or domain)" Style="{StaticResource GctLabel}"/>
                  <TextBox x:Name="TenantId"/>
                  <StackPanel x:Name="CertificatePanel">
                    <TextBlock Text="Certificate thumbprint (certificate store)" Style="{StaticResource GctLabel}" ToolTip="Certificate of the application in Cert:\CurrentUser\My or Cert:\LocalMachine\My (40 hexadecimal characters)"/>
                    <TextBox x:Name="CertificateThumbprint" FontFamily="Cascadia Mono, Consolas"/>
                  </StackPanel>
                  <TextBlock x:Name="OnPremisesAuthLabel" Text="On-premises authentication (source on-premises)" Style="{StaticResource GctLabel}"/>
                  <ComboBox x:Name="OnPremisesAuthentication" HorizontalAlignment="Stretch"/>
                  <StackPanel x:Name="CredentialPanel">
                    <TextBlock Text="Scheme" Style="{StaticResource GctLabel}"/>
                    <ComboBox x:Name="OnPremisesCredentialScheme" HorizontalAlignment="Stretch"/>
                    <TextBlock Text="On-premises account (DOMAIN\user or user@domain)" Style="{StaticResource GctLabel}"/>
                    <TextBox x:Name="OnPremisesUserName"/>
                    <TextBlock Text="Password (kept in memory for this run only)" Style="{StaticResource GctLabel}"/>
                    <PasswordBox x:Name="OnPremisesPassword"/>
                  </StackPanel>
                  <TextBlock x:Name="SignInHint" Style="{StaticResource GctHint}" Margin="0,10,0,0"/>
                </StackPanel>
              </Border>

            </StackPanel>
          </Grid>
        </StackPanel>
      </ScrollViewer>

      <Grid Grid.Column="2">
        <Grid.RowDefinitions>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*"/>
        </Grid.RowDefinitions>
        <Border x:Name="AboutCard" Style="{StaticResource GctCard}">
          <StackPanel>
            <TextBlock x:Name="DescriptionTitle" Style="{StaticResource GctCardTitle}" TextWrapping="Wrap"/>
            <TextBlock x:Name="Description" TextWrapping="Wrap" FontSize="12" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
          </StackPanel>
        </Border>
        <Border x:Name="WarningBox" Grid.Row="1" Margin="0,0,0,12" Padding="14,10" CornerRadius="8" BorderThickness="1">
          <Grid>
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="Auto"/>
              <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>
            <TextBlock x:Name="WarningIcon" Style="{StaticResource GctIcon}" FontSize="16" Margin="0,1,10,0" VerticalAlignment="Top"/>
            <TextBlock x:Name="Warning" Grid.Column="1" TextWrapping="Wrap" FontSize="12.5" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
          </Grid>
        </Border>
        <Border x:Name="Progress" Grid.Row="2" Style="{StaticResource GctCard}">
          <Grid>
            <Grid.RowDefinitions>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="Auto"/>
              <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid Margin="0,0,0,8">
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Text="Progress" Style="{StaticResource GctCardTitle}"/>
              <Border x:Name="StatusPill" Grid.Column="1" CornerRadius="10" Padding="10,2" VerticalAlignment="Top">
                <TextBlock x:Name="Status" FontSize="12" FontWeight="SemiBold"/>
              </Border>
            </Grid>
            <StackPanel x:Name="ResultPanel" Grid.Row="1" Visibility="Collapsed" Margin="0,0,0,12">
              <WrapPanel x:Name="Counts"/>
              <TextBlock x:Name="ResultText" TextWrapping="Wrap" FontSize="12" Margin="0,6,0,0" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
            </StackPanel>
            <Border x:Name="CodeBanner" Grid.Row="2" Visibility="Collapsed" Margin="0,0,0,12" Padding="14,10" CornerRadius="6" BorderThickness="1"
                    Background="{DynamicResource GctAccentSoft}" BorderBrush="{DynamicResource AccentFillColorDefaultBrush}">
              <StackPanel>
                <TextBlock Text="Device code: type it on the sign-in page, from any device, with the source account" FontSize="12" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
                <TextBlock x:Name="CodeText" FontSize="28" FontWeight="SemiBold" FontFamily="Cascadia Mono, Consolas" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
                <TextBlock x:Name="CodePage" FontSize="12" TextWrapping="Wrap" Foreground="{DynamicResource AccentTextFillColorPrimaryBrush}"/>
                <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                  <Button x:Name="CopyCode" Margin="0,0,8,0">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Style="{StaticResource GctIcon}" Text="&#xE8C8;" Margin="0,2,8,0"/>
                      <TextBlock Text="Copy the code"/>
                    </StackPanel>
                  </Button>
                  <Button x:Name="OpenPage" Content="Open the page"/>
                </StackPanel>
              </StackPanel>
            </Border>
            <Grid Grid.Row="3">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Grid Margin="0,0,0,10">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Horizontal">
                  <ToggleButton x:Name="ViewTimeline" Content="Timeline" IsChecked="True" Padding="12,4" Margin="0,0,4,0"
                                ToolTip="The stages in order, their checks and, on demand, the diagnostics of each check"/>
                  <ToggleButton x:Name="ViewLog" Content="Diagnostic log" Padding="12,4"
                                ToolTip="Everything in time order, as text: stages, HTTP exchanges, checks"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                  <CheckBox x:Name="ShowDiagnostics" Content="Diagnostics" VerticalAlignment="Center" Margin="0,0,8,0" MinWidth="0"
                            ToolTip="Unfolds the diagnostics of every check (details, HTTP requests and responses) and adds them to the diagnostic log"/>
                  <Button x:Name="CopyLog" Padding="10,4" ToolTip="Copies the whole diagnostic log, with the requests and the responses">
                    <StackPanel Orientation="Horizontal">
                      <TextBlock Style="{StaticResource GctIcon}" Text="&#xE8C8;" Margin="0,2,8,0"/>
                      <TextBlock Text="Copy the log"/>
                    </StackPanel>
                  </Button>
                </StackPanel>
              </Grid>
              <TextBlock x:Name="LogEmpty" Grid.Row="1" FontSize="13" Foreground="{DynamicResource TextFillColorTertiaryBrush}" TextWrapping="Wrap"
                         Text="The stages and their checks appear here as they run. Diagnostics shows, for each check, its details and the HTTP requests and responses."/>
              <ScrollViewer x:Name="LogScroll" Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                <StackPanel x:Name="Timeline" Margin="0,0,12,0"/>
              </ScrollViewer>
              <TextBox x:Name="DiagLog" Grid.Row="1" Visibility="Collapsed" IsReadOnly="True" TextWrapping="NoWrap" AcceptsReturn="True"
                       FontFamily="Cascadia Mono, Consolas" FontSize="11.5" Padding="8,6"
                       VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
            </Grid>
          </Grid>
        </Border>
      </Grid>
    </Grid>

    <Border x:Name="Actions" Grid.Row="2" Padding="24,12" BorderThickness="0,1,0,0"
            BorderBrush="{DynamicResource DividerStrokeColorDefaultBrush}" Background="{DynamicResource LayerFillColorDefaultBrush}">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal">
          <Button x:Name="Run" MinWidth="150" Padding="16,6" Margin="0,0,8,0">
            <StackPanel Orientation="Horizontal">
              <TextBlock Style="{StaticResource GctIcon}" Text="&#xE768;" Margin="0,2,8,0"/>
              <TextBlock Text="Run the test"/>
            </StackPanel>
          </Button>
          <Button x:Name="Cancel" Content="Cancel" MinWidth="96" IsEnabled="False"/>
        </StackPanel>
        <TextBlock x:Name="Footer" Grid.Column="1" Margin="16,0" VerticalAlignment="Center" TextTrimming="CharacterEllipsis" FontSize="12"
                   Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
        <StackPanel Grid.Column="2" Orientation="Horizontal">
          <Button x:Name="OpenReport" Content="Open the report" Margin="0,0,8,0" IsEnabled="False"/>
          <Button x:Name="OpenFolder" Content="Open the folder" Margin="0,0,8,0" IsEnabled="False"/>
          <Button x:Name="Close" Content="Close" MinWidth="96" IsCancel="True"/>
        </StackPanel>
      </Grid>
    </Border>
  </Grid>
</Window>
'@
}

function New-GctTestForm {
    <#
    .SYNOPSIS
        Builds the window (without showing it). Used by Show-GctTestGui, the tests and the documentation tool.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][hashtable]$Configuration, [ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')

    $look = Initialize-GctGuiTheme -Theme $Theme
    $catalog = @(Get-GctTestCatalog)
    $xaml = Get-GctGuiXaml
    $window = [Windows.Markup.XamlReader]::Parse($xaml)
    $window.Title = "Graph Connectivity Tests $($script:ToolVersion)"
    Set-GctGuiTheme -Window $window -Theme $look
    $controls = @{}
    foreach ($m in [regex]::Matches($xaml, 'x:Name="(\w+)"')) { $controls[$m.Groups[1].Value] = $window.FindName($m.Groups[1].Value) }
    if ($look.Fluent) { $controls.Run.SetResourceReference([Windows.FrameworkElement]::StyleProperty, 'AccentButtonStyle') }
    else {
        $controls.Run.SetResourceReference([Windows.Controls.Control]::BackgroundProperty, 'AccentFillColorDefaultBrush')
        $controls.Run.SetResourceReference([Windows.Controls.Control]::ForegroundProperty, 'AccentButtonForeground')
    }
    $controls.Version.Text = "v$($script:ToolVersion)  " + [char]0x00B7 + '  Nicolas Fabert'

    # Test cards.
    foreach ($test in $catalog) { [void]$controls.TestPicker.Items.Add((New-GctGuiTestCard -Test $test)) }
    $wanted = if ($catalog.Name -contains [string]$Configuration.TestType) { [string]$Configuration.TestType } else { 'FreeBusy' }
    $controls.TestPicker.SelectedIndex = [Math]::Max(0, [Array]::IndexOf([string[]]@($catalog.Name), $wanted))

    # Mailboxes.
    $controls.SourceMailbox.Text = [string]$Configuration.SourceMailbox
    $controls.TargetMailboxes.Text = (@($Configuration.TargetMailboxes | Where-Object { $_ }) -join [Environment]::NewLine)
    $controls.OnPremisesHost.Text = [string]$Configuration.OnPremisesHost
    $controls.SkipOnPremisesChecks.IsChecked = [bool]$Configuration.SkipOnPremisesChecks
    $controls.Domain.Text = [string]$Configuration.Domain
    $clouds = foreach ($key in $script:Clouds.Keys) { [pscustomobject]@{ Name = $key; Text = "$key - $($script:Clouds[$key].Name)" } }
    Set-GctGuiChoices -ComboBox $controls.Cloud -Choices @($clouds) -Selected ([string]$Configuration.Cloud)

    # Sign-in. The client: a known name (or its ID), else your application.
    Set-GctGuiChoices -ComboBox $controls.Context -Choices $script:GuiContexts -Selected ([string]$Configuration.Context)
    $client = 'Custom'
    foreach ($key in 'MicrosoftOffice', 'GraphCommandLineTools') {
        if ([string]$Configuration.ClientId -ieq $key -or [string]$Configuration.ClientId -ieq $script:KnownClients[$key].Id) { $client = $key }
    }
    if ([string]$Configuration.ClientId -eq '' -and [string]$Configuration.Context -ne 'Application') { $client = 'MicrosoftOffice' }
    Set-GctGuiChoices -ComboBox $controls.Client -Choices $script:GuiClients -Selected $client
    $controls.AppClientId.Text = if ($client -eq 'Custom') { [string]$Configuration.ClientId } else { '' }
    $controls.RedirectUri.Text = if ($client -eq 'Custom') { [string]$Configuration.RedirectUri } else { '' }
    $controls.TenantId.Text = [string]$Configuration.TenantId
    $controls.CertificateThumbprint.Text = [string]$Configuration.CertificateThumbprint
    Set-GctGuiChoices -ComboBox $controls.OnPremisesAuthentication -Choices $script:GuiOnPremisesAuthentication -Selected ([string]$Configuration.OnPremisesAuthentication)
    Set-GctGuiChoices -ComboBox $controls.OnPremisesCredentialScheme -Choices $script:GuiCredentialSchemes -Selected ([string]$Configuration.OnPremisesCredentialScheme)
    $controls.OnPremisesUserName.Text = [string]$Configuration.OnPremisesUserName
    Set-GctGuiChoices -ComboBox $controls.SignIn -Choices $script:GuiSignInModes -Selected ([string]$Configuration.SignIn)

    # Options.
    $controls.AllowWrite.IsChecked = [bool]$Configuration.AllowWrite
    $controls.AdminChecks.IsChecked = [bool]$Configuration.AdminChecks
    $controls.FreeBusyDays.Text = [string]$Configuration.FreeBusyDays

    $script:Gui = @{
        Form = $window; Controls = $controls; Configuration = $Configuration.Clone(); Catalog = $catalog; Theme = $look
        Running = $false; LastReport = $null; LastFolder = $null
        Lines = [Collections.Generic.List[string]]::new(); CodePage = $null
        Stages = [Collections.Generic.List[object]]::new(); Checks = [Collections.Generic.List[object]]::new(); Events = [Collections.Generic.List[object]]::new()
        CurrentStage = $null; Diagnostics = $false
        ClosingGuard = [ComponentModel.CancelEventHandler] {
            param($source, $e)
            $e.Cancel = $true
            if ($script:Ui) { $script:Ui.Cancel = $true }
            Add-GctGuiLine 'Warn' 'A test is running: it stops at the next check, then the window can be closed.'
        }
    }
    Set-GctGuiStatus 'Ready' 'Ready'

    foreach ($name in 'TestPicker', 'Context', 'Client', 'SignIn', 'OnPremisesAuthentication', 'OnPremisesCredentialScheme', 'Cloud') { $controls[$name].Add_SelectionChanged({ Update-GctGuiForm }) }
    foreach ($name in 'AllowWrite', 'AdminChecks', 'SkipOnPremisesChecks') { $controls[$name].Add_Click({ Update-GctGuiForm }) }
    foreach ($name in 'SourceMailbox', 'TargetMailboxes', 'AppClientId', 'TenantId', 'CertificateThumbprint') { $controls[$name].Add_TextChanged({ Update-GctGuiForm }) }
    $controls.Run.Add_Click({ Invoke-GctGuiRun })
    $controls.Cancel.Add_Click({
            if ($script:Ui) { $script:Ui.Cancel = $true; Add-GctGuiLine 'Warn' 'Cancellation requested: the run stops at the next check.' }
        })
    $controls.OpenReport.Add_Click({ if ($script:Gui.LastReport) { Start-Process -FilePath $script:Gui.LastReport } })
    $controls.OpenFolder.Add_Click({ if ($script:Gui.LastFolder) { Start-Process -FilePath 'explorer.exe' -ArgumentList "`"$($script:Gui.LastFolder)`"" } })
    $controls.CopyCode.Add_Click({
            try { [Windows.Clipboard]::SetText([string]$script:Gui.Controls.CodeText.Text); Add-GctGuiLine 'Info' 'Device code copied to the clipboard.' }
            catch { Add-GctGuiLine 'Warn' "The clipboard is not available: $($_.Exception.Message)" }
        })
    $controls.OpenPage.Add_Click({ if ($script:Gui.CodePage) { Start-Process -FilePath $script:Gui.CodePage } })
    $controls.ViewTimeline.Add_Click({ Set-GctGuiView 'Timeline' })
    $controls.ViewLog.Add_Click({ Set-GctGuiView 'Log' })
    $controls.ShowDiagnostics.Add_Click({ Set-GctGuiDiagnosticMode -On ([bool]$script:Gui.Controls.ShowDiagnostics.IsChecked) })
    $controls.CopyLog.Add_Click({
            $all = [Text.StringBuilder]::new()
            foreach ($ev in $script:Gui.Events) { [void]$all.Append((Format-GctGuiEvent -Event $ev -Full)) }
            Copy-GctGuiText -Text $all.ToString().TrimStart() -What 'The diagnostic log'
        })

    $area = [Windows.SystemParameters]::WorkArea
    $window.Width = [Math]::Min($window.Width, $area.Width)
    $window.Height = [Math]::Min($window.Height, $area.Height)
    $window.MinWidth = [Math]::Min($window.MinWidth, $area.Width)
    $window.MinHeight = [Math]::Min($window.MinHeight, $area.Height)

    Update-GctGuiForm
    [pscustomobject]@{ Form = $window; Controls = $controls; Lines = $script:Gui.Lines; Gui = $script:Gui }
}

function Set-GctGuiChoices {
    <# Fills a combo box with Name/Text choices (the Name in the Tag) and selects one by name (else the first). #>
    param([Parameter(Mandatory = $true)][Windows.Controls.ComboBox]$ComboBox, [Parameter(Mandatory = $true)][object[]]$Choices, [string]$Selected)

    $ComboBox.Items.Clear()
    $index = 0
    for ($i = 0; $i -lt $Choices.Count; $i++) {
        $item = [Windows.Controls.ComboBoxItem]::new()
        $item.Content = $Choices[$i].Text
        $item.Tag = $Choices[$i].Name
        [void]$ComboBox.Items.Add($item)
        if ($Choices[$i].Name -ieq $Selected) { $index = $i }
    }
    $ComboBox.SelectedIndex = $index
}

function Get-GctGuiChoice {
    <# Name of the choice selected in a combo box or in the test picker. #>
    param([Parameter(Mandatory = $true)][object]$Control)
    $item = $Control.SelectedItem
    if ($null -eq $item) { return $null }
    return [string]$item.Tag
}

function Split-GctGuiList {
    <# Addresses typed one per line, or separated by commas, semicolons or spaces. #>
    param([AllowEmptyString()][string]$Text)
    @(([string]$Text) -split '[,;\s]+' | Where-Object { $_ } | Select-Object -Unique)
}

function New-GctGuiTestCard {
    <# Content of one test card: icon, name, title. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Test)

    $panel = [Windows.Controls.StackPanel]::new()
    $top = [Windows.Controls.StackPanel]::new()
    $top.Orientation = [Windows.Controls.Orientation]::Horizontal
    $top.Margin = [Windows.Thickness]::new(0, 0, 0, 4)
    $icon = [Windows.Controls.TextBlock]::new()
    $icon.Text = [string][char]$(if ($script:GuiTestGlyphs.ContainsKey($Test.Name)) { $script:GuiTestGlyphs[$Test.Name] } else { 0xE9D5 })
    $icon.FontFamily = [Windows.Media.FontFamily]::new($script:GuiIconFont)
    $icon.FontSize = 16
    $icon.Margin = [Windows.Thickness]::new(0, 1, 8, 0)
    $icon.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'GctBrandText')
    $name = [Windows.Controls.TextBlock]::new()
    $name.Text = $Test.Name
    $name.FontSize = 13
    $name.FontWeight = [Windows.FontWeights]::SemiBold
    $name.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'TextFillColorPrimaryBrush')
    [void]$top.Children.Add($icon)
    [void]$top.Children.Add($name)
    $title = [Windows.Controls.TextBlock]::new()
    $title.Text = $Test.DisplayName
    $title.FontSize = 11.5
    $title.TextWrapping = [Windows.TextWrapping]::Wrap
    $title.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'TextFillColorSecondaryBrush')
    [void]$panel.Children.Add($top)
    [void]$panel.Children.Add($title)
    $item = [Windows.Controls.ListBoxItem]::new()
    $item.Content = $panel
    $item.Tag = $Test.Name
    $item.ToolTip = $Test.DisplayName
    return $item
}

function Set-GctGuiField {
    <# Enables or disables a field and dims its label. #>
    param([Parameter(Mandatory = $true)][object]$Control, [AllowNull()][object]$Label, [bool]$Enabled)
    $Control.IsEnabled = $Enabled
    if ($Label) { $Label.Opacity = if ($Enabled) { 1.0 } else { 0.45 } }
}

function Get-GctGuiTest {
    <# The test chosen in the picker (an entry of Get-GctTestCatalog). #>
    $g = $script:Gui
    $name = Get-GctGuiChoice -Control $g.Controls.TestPicker
    $g.Catalog | Where-Object Name -eq $name | Select-Object -First 1
}

function Update-GctGuiForm {
    <# Enables the fields the chosen test uses, shows the fields of the context and the notice. #>
    $g = $script:Gui
    if (-not $g) { return }
    $c = $g.Controls
    $test = Get-GctGuiTest
    if (-not $test) { return }
    $stages = @($test.Stages)
    $app = (Get-GctGuiChoice $c.Context) -eq 'Application'
    $client = Get-GctGuiChoice $c.Client
    $show = { param($Control, [bool]$Visible) $Control.Visibility = if ($Visible) { 'Visible' } else { 'Collapsed' } }

    $c.DescriptionTitle.Text = $test.DisplayName
    $c.Description.Text = $test.Description

    # Mailboxes: which fields the test uses.
    $targets = @(Split-GctGuiList $c.TargetMailboxes.Text)
    $usesTargets = $stages -contains 'Target' -or $stages -contains 'Provisioning'
    Set-GctGuiField $c.TargetMailboxes $c.TargetLabel $usesTargets
    $c.TargetHint.Opacity = if ($usesTargets) { 1.0 } else { 0.6 }
    $c.TargetHint.Text = switch ($test.Name) {
        'FreeBusy' { "Required: the mailboxes whose free/busy the source asks for (attendees). $($targets.Count) of 20." }
        'Provisioning' { "The addresses checked; empty = the source mailbox. $($targets.Count) of 20." }
        'All' { "Free/busy and provisioning of these mailboxes; empty = no free/busy of others, provisioning of the source. $($targets.Count) of 20." }
        default { 'Not used by this test.' }
    }
    if ($usesTargets -and $targets.Count -gt 20) { $c.TargetHint.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'GctCritical') }
    else { $c.TargetHint.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, 'TextFillColorTertiaryBrush') }
    $onPremises = $stages -contains 'OnPremises'
    Set-GctGuiField $c.OnPremisesHost $c.OnPremisesLabel $onPremises
    Set-GctGuiField $c.SkipOnPremisesChecks $null $onPremises
    Set-GctGuiField $c.Domain $c.DomainLabel ($stages -contains 'Dns')

    # Sign-in: the fields of the context.
    & $show $c.Client (-not $app)
    & $show $c.ClientLabel (-not $app)
    & $show $c.CustomPanel ($app -or $client -eq 'Custom')
    & $show $c.RedirectPanel (-not $app)
    & $show $c.CertificatePanel $app
    Set-GctGuiField $c.SignIn $c.SignInLabel (-not $app)
    $c.TenantLabel.Text = if ($app) { 'Tenant (required: ID or domain)' } else { 'Tenant (optional: ID or domain)' }
    Set-GctGuiField $c.OnPremisesAuthentication $c.OnPremisesAuthLabel ($stages -contains 'FreeBusy')
    & $show $c.CredentialPanel ($stages -contains 'FreeBusy' -and (Get-GctGuiChoice $c.OnPremisesAuthentication) -eq 'Credential')
    $c.SignInHint.Text = if ($app) {
        'Your application signs in alone with the certificate (client credentials) and reads /users/{source}: application permissions of Microsoft Graph (Calendars.Read, User.Read.All...) with an administrator consent. No secret is typed here.'
    }
    else {
        switch ($client) {
            'MicrosoftOffice' { 'The source user signs in with Microsoft Office, a client of Microsoft present in every tenant and pre-authorized for Microsoft Graph: no application to register, no consent. Password, MFA and Conditional Access as usual.' }
            'GraphCommandLineTools' { 'The source user signs in with the client of Connect-MgGraph: the permissions are asked by name and need the consent of the user or of an administrator.' }
            default { 'Your own public client (multi-tenant if the users come from several tenants), with a redirect URI of type Mobile and desktop applications.' }
        }
    }

    # Options.
    $canWrite = $stages -contains 'AutoReply' -and -not $app
    Set-GctGuiField $c.AllowWrite $null $canWrite
    Set-GctGuiField $c.AllowWriteHint $null $canWrite
    $writing = $canWrite -and [bool]$c.AllowWrite.IsChecked
    $c.AllowWriteHint.Text = if (-not ($stages -contains 'AutoReply')) { 'Only for MailboxServices and All: this test changes nothing.' }
    elseif ($app) { 'Only in the User context: the application only reads the automatic replies.' }
    else { 'Schedules the automatic replies 60 days ahead, reads them back and restores them - only when they are disabled. Use a test mailbox.' }
    $c.AllowWriteHint.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, $(if ($writing) { 'GctCaution' } else { 'TextFillColorTertiaryBrush' }))
    $admin = $stages -contains 'Hybrid'
    Set-GctGuiField $c.AdminChecks $null $admin
    Set-GctGuiField $c.AdminChecksHint $null $admin
    $days = $stages -contains 'FreeBusy' -or $stages -contains 'Availability'
    Set-GctGuiField $c.FreeBusyDays $c.FreeBusyDaysLabel $days

    # The notice.
    $source = if ($c.SourceMailbox.Text.Trim()) { $c.SourceMailbox.Text.Trim() } else { 'the source mailbox' }
    if ($writing) {
        Set-GctGuiNotice "Changes allowed: the automatic replies of $source are scheduled 60 days ahead, read back and restored (only when they are disabled, so a real absence is never touched). Use a test mailbox." Caution
    }
    elseif ($test.Name -eq 'FreeBusy' -and -not $targets.Count) {
        Set-GctGuiNotice 'The free/busy test needs at least one target mailbox: the mailboxes whose free/busy the source asks for.' Caution
    }
    elseif ($app) {
        Set-GctGuiNotice 'Application context: your application signs in with the certificate of the certificate store (no secret typed here). Read-only: an application never changes the automatic replies.' Info
    }
    elseif ($test.Writes -or $stages -contains 'AutoReply') {
        Set-GctGuiNotice "Read-only: the automatic replies of $source are only read. Check `"Allow changes`" to also schedule them, read them back and restore them." Info
    }
    else {
        Set-GctGuiNotice "Read-only: nothing is changed in any mailbox. $source signs in once; every request and response is in the report." Info
    }
    $g.Controls.Footer.Text = try { Get-GctSignInText -Configuration (Get-GctGuiConfiguration) } catch { "Reports: $($g.Configuration.OutputPath)" }
}

function Get-GctGuiConfiguration {
    <#
        The settings of a run from the window: the configuration it was opened with, changed by the fields.
        A field the test does not use is not passed (empty, or the value of the configuration).
    #>
    $g = $script:Gui
    $c = $g.Controls
    $cfg = $g.Configuration.Clone()
    $test = Get-GctGuiTest
    $cfg.TestType = [string]$test.Name
    $cfg.SourceMailbox = $c.SourceMailbox.Text.Trim()
    $cfg.TargetMailboxes = @(if ($c.TargetMailboxes.IsEnabled) { Split-GctGuiList $c.TargetMailboxes.Text })
    $cfg.OnPremisesHost = if ($c.OnPremisesHost.IsEnabled) { $c.OnPremisesHost.Text.Trim() } else { '' }
    $cfg.SkipOnPremisesChecks = $c.SkipOnPremisesChecks.IsEnabled -and [bool]$c.SkipOnPremisesChecks.IsChecked
    $cfg.Domain = if ($c.Domain.IsEnabled) { $c.Domain.Text.Trim() } else { '' }
    $cfg.Cloud = Get-GctGuiChoice $c.Cloud
    $cfg.Context = Get-GctGuiChoice $c.Context
    $client = Get-GctGuiChoice $c.Client
    if ($cfg.Context -eq 'Application') {
        $cfg.ClientId = $c.AppClientId.Text.Trim()
        $cfg.RedirectUri = ''
        # Copied from the certificate console, a thumbprint often has spaces and an invisible mark.
        $cfg.CertificateThumbprint = $c.CertificateThumbprint.Text -replace '[\s‎‏]', ''
    }
    else {
        $cfg.ClientId = if ($client -eq 'Custom') { $c.AppClientId.Text.Trim() } else { $client }
        $cfg.RedirectUri = if ($client -eq 'Custom') { $c.RedirectUri.Text.Trim() } else { '' }
        $cfg.CertificateThumbprint = ''
        $cfg.SignIn = Get-GctGuiChoice $c.SignIn
    }
    $cfg.TenantId = $c.TenantId.Text.Trim()
    if ($c.OnPremisesAuthentication.IsEnabled) { $cfg.OnPremisesAuthentication = Get-GctGuiChoice $c.OnPremisesAuthentication }
    if ($cfg.OnPremisesAuthentication -eq 'Credential') {
        $cfg.OnPremisesCredentialScheme = Get-GctGuiChoice $c.OnPremisesCredentialScheme
        $cfg.OnPremisesUserName = $c.OnPremisesUserName.Text.Trim()
    }
    $cfg.AllowWrite = $c.AllowWrite.IsEnabled -and [bool]$c.AllowWrite.IsChecked
    $cfg.AdminChecks = $c.AdminChecks.IsEnabled -and [bool]$c.AdminChecks.IsChecked
    if ($c.FreeBusyDays.IsEnabled) {
        $days = 0
        $cfg.FreeBusyDays = if ([int]::TryParse($c.FreeBusyDays.Text.Trim(), [ref]$days)) { $days } else { $c.FreeBusyDays.Text.Trim() }
    }
    return $cfg
}

function Invoke-GctGuiRun {
    <# Checks the settings (every problem listed), runs the test on the window thread, writes the report. #>
    $g = $script:Gui
    $c = $g.Controls
    $cfg = Get-GctGuiConfiguration
    Clear-GctGuiProgress
    $check = Test-GctConfiguration -Configuration $cfg
    $credential = $null
    if ($cfg.OnPremisesAuthentication -eq 'Credential' -and $c.OnPremisesAuthentication.IsEnabled) {
        $problems = [Collections.Generic.List[string]]::new(); foreach ($p in @($check.Problems)) { $problems.Add($p) }
        if (-not $cfg.OnPremisesUserName) { $problems.Add('On-premises authentication Credential: type the on-premises account (DOMAIN\user or user@domain).') }
        if ($c.OnPremisesPassword.SecurePassword.Length -eq 0) { $problems.Add('On-premises authentication Credential: type the password of the on-premises account.') }
        if ($problems.Count) { $check = [pscustomobject]@{ IsValid = $false; Problems = @($problems) } }
        else { $credential = [pscredential]::new($cfg.OnPremisesUserName, $c.OnPremisesPassword.SecurePassword.Copy()) }
    }
    if (-not $check.IsValid) {
        Add-GctGuiLine 'Step' 'Fix these values, then run the test again'
        foreach ($problem in $check.Problems) { Add-GctGuiLine 'Fail' $problem }
        Set-GctGuiStatus "$(@($check.Problems).Count) value(s) to fix" 'Failed'
        return
    }

    $script:Ui = @{
        Sink       = { param($Status, $Text) Add-GctGuiLine $Status $Text }
        StepAdded  = { param($Step, $Exchanges) Add-GctGuiCheck -Step $Step -Exchanges $Exchanges }
        Pump       = { Invoke-GctGuiPump }
        DeviceCode = { param($Code, $Page) Show-GctGuiDeviceCode -Code $Code -Page $Page }
        Cancel     = $false
    }
    try {
        $g.Running = $true
        $g.Form.add_Closing($g.ClosingGuard)
        foreach ($b in 'Run', 'OpenReport', 'OpenFolder', 'Close', 'SettingsScroll', 'TestPicker') { $c[$b].IsEnabled = $false }
        $c.Cancel.IsEnabled = $true
        Set-GctGuiStatus "Running $($cfg.TestType)..." 'Running'
        $c.Footer.Text = Get-GctSignInText -Configuration $cfg
        Write-GctLog 'STEP' "GUI run: $($cfg.TestType) ($($cfg.Context), $($cfg.ClientId)) for $($cfg.SourceMailbox)"
        $result = Invoke-GctConnectivityTest -Configuration $cfg -TestType $cfg.TestType -OnPremisesCredential $credential -Quiet
        $report = Export-GctReport -Result $result -OutputPath $cfg.OutputPath -Prefix $cfg.ReportPrefix -Formats $cfg.ReportFormats -Delimiter $cfg.CsvDelimiter
        $g.LastFolder = $report.Directory
        $g.LastReport = Get-GctField $report.Files 'Html'
        $c.Footer.Text = "Report: $($report.Directory)"
        Show-GctGuiResult -Result $result
    }
    catch {
        Add-GctGuiLine 'Fail' $_.Exception.Message
        Set-GctGuiStatus 'Failed - see the progress' 'Failed'
    }
    finally {
        Complete-GctGuiStage
        $g.Form.remove_Closing($g.ClosingGuard)
        $script:Ui = $null
        $g.Running = $false
        foreach ($b in 'Run', 'Close', 'SettingsScroll', 'TestPicker') { $c[$b].IsEnabled = $true }
        $c.Cancel.IsEnabled = $false
        $c.OpenReport.IsEnabled = [bool]$g.LastReport
        $c.OpenFolder.IsEnabled = [bool]$g.LastFolder
    }
}

function Show-GctGuiResult {
    <# After a run: the status pill, one chip per count, the free/busy answers, the first issue and what to do. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Result)

    $c = $script:Gui.Controls
    Complete-GctGuiStage
    $n = $Result.Counts
    $dot = [char]0x00B7
    Set-GctGuiStatus ("{0}  {1}  {2}/{3} checks passed  {1}  {4}" -f $Result.Status, $dot, $n.Passed, @($Result.Steps).Count, (Format-GctDuration $Result.DurationSeconds)) $Result.Status
    $c.Counts.Children.Clear()
    $chips = [ordered]@{ Passed = 'passed'; Warning = 'warning(s)'; Blocked = 'blocked'; Failed = 'failed'; Skipped = 'skipped' }
    foreach ($key in $chips.Keys) {
        if (-not $n[$key] -and $key -ne 'Passed') { continue }
        [void]$c.Counts.Children.Add((New-GctGuiChip -Text "$($n[$key]) $($chips[$key])" -Status $key))
    }
    $lines = [Collections.Generic.List[string]]::new()
    $fb = @($Result.FreeBusy)
    if ($fb.Count) { $lines.Add(('Free/busy: {0} of {1} mailbox(es) answered.' -f @($fb | Where-Object Status -ne 'Failed').Count, $fb.Count)) }
    # The first problem by name only: its message, details and HTTP exchanges are in the timeline, scrolled to it.
    $first = @($Result.Steps | Where-Object Status -eq 'Failed') + @($Result.Steps | Where-Object Status -in 'Blocked', 'Warning') | Select-Object -First 1
    if ($first) { $lines.Add("First issue: $($first.Name) ($($Result.StageTitles[$first.Stage])), shown below.") }
    $lines.Add($(switch ($Result.Status) {
                'Passed' { 'Nothing to do. The report shows every request and response.' }
                'Blocked' { 'A writing step needs "Allow changes" (test mailbox only).' }
                'Warning' { 'Read the warnings in the report: each one gives the cause and what to check.' }
                default { 'Open the report: each failed step gives the error, its cause and what to check.' }
            }))
    $c.ResultText.Text = $lines -join [Environment]::NewLine
    $c.ResultPanel.Visibility = 'Visible'
    $c.CodeBanner.Visibility = 'Collapsed'
    if ($first) {
        $check = $script:Gui.Checks | Where-Object { [object]::ReferenceEquals($_.Step, $first) } | Select-Object -First 1
        if ($check) {
            Set-GctGuiStageExpanded -Stage $check.Stage -Expanded $true
            Invoke-GctGuiPump
            $c.LogScroll.ScrollToVerticalOffset([Math]::Max(0, $check.Root.TranslatePoint([Windows.Point]::new(0, 0), $c.Timeline).Y - 40))
        }
    }
}

function New-GctGuiChip {
    <# A small coloured count: 12 passed, 1 warning... #>
    param([Parameter(Mandatory = $true)][string]$Text, [Parameter(Mandatory = $true)][string]$Status)

    $pair = Get-GctGuiStatusBrushes -Status $Status
    $border = [Windows.Controls.Border]::new()
    $border.CornerRadius = [Windows.CornerRadius]::new(10)
    $border.Padding = [Windows.Thickness]::new(10, 2, 10, 3)
    $border.Margin = [Windows.Thickness]::new(0, 0, 8, 4)
    $border.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, $pair[1])
    $label = [Windows.Controls.TextBlock]::new()
    $label.Text = $Text
    $label.FontSize = 12
    $label.FontWeight = [Windows.FontWeights]::SemiBold
    $label.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, $pair[0])
    $border.Child = $label
    return $border
}

function Show-GctGuiDeviceCode {
    <# The device code, big, with Copy and Open buttons (called by Invoke-GctDeviceCodeSignIn through $script:Ui.DeviceCode). #>
    param([Parameter(Mandatory = $true)][string]$Code, [string]$Page)

    $g = $script:Gui
    if (-not $g) { return }
    $g.Controls.CodeText.Text = $Code
    $g.Controls.CodePage.Text = $Page
    $g.CodePage = $Page
    $g.Controls.CodeBanner.Visibility = 'Visible'
    Invoke-GctGuiPump
}

#region Theme and window helpers ----------------------------------------------------------------------

function New-GctGuiBrush {
    param([Parameter(Mandatory = $true)][string]$Color)
    $brush = [Windows.Media.SolidColorBrush]::new([Windows.Media.ColorConverter]::ConvertFromString($Color))
    $brush.Freeze()
    return $brush
}

function Initialize-GctGuiTheme {
    <#
        Loads WPF and applies the theme to the application: Fluent (.NET 9+), light or dark as Windows
        (System), or Light / Dark for the documentation images. Returns Fluent and Dark.
    #>
    param([ValidateSet('System', 'Light', 'Dark')][string]$Theme = 'System')

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml
    $dark = $Theme -eq 'Dark'
    if ($Theme -eq 'System') { $dark = Test-GctGuiDarkMode }
    # One application per process: created once, never shut down by a closed window.
    $app = [Windows.Application]::Current
    if (-not $app) {
        $app = [Windows.Application]::new()
        $app.ShutdownMode = [Windows.ShutdownMode]::OnExplicitShutdown
    }
    $fluent = $null -ne [Windows.Application].GetProperty('ThemeMode')
    if ($fluent) {
        # ThemeMode is the Fluent theme of WPF (.NET 9 and later); its name is experimental in .NET 9.
        # Light or Dark, never System: without the setting (Windows Server 2016) WPF would pick dark and the
        # colours of the window (light) would not match.
        $app.ThemeMode = [Windows.ThemeMode]::new($(if ($dark) { 'Dark' } else { 'Light' }))
    }
    [pscustomobject]@{ Fluent = $fluent; Dark = $dark; Application = $app }
}

function Test-GctGuiDarkMode {
    <# Windows shows the applications in dark mode (AppsUseLightTheme = 0); light when the setting is missing. #>
    $personalize = Get-ItemProperty -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -ErrorAction SilentlyContinue
    return [string](Get-GctField $personalize 'AppsUseLightTheme') -eq '0'
}

function Set-GctGuiTheme {
    <#
        Colours of the window: the crimson accent of the other tools of the author on the Fluent accent resources, the status
        colours, and with the classic theme (.NET 8) the Fluent resources the window uses.
    #>
    param([Parameter(Mandatory = $true)][Windows.Window]$Window, [Parameter(Mandatory = $true)][pscustomobject]$Theme)

    $dark = $Theme.Dark
    $r = $Window.Resources
    $set = { param([string[]]$Keys, [string]$LightColor, [string]$DarkColor) $b = New-GctGuiBrush $(if ($dark) { $DarkColor } else { $LightColor }); foreach ($k in $Keys) { $r[$k] = $b } }
    if (-not $Theme.Fluent) {
        & $set 'ApplicationBackgroundBrush' '#F7F4EF' '#202020'
        & $set 'CardBackgroundFillColorDefaultBrush' '#FFFFFF' '#2B2B2B'
        & $set 'CardStrokeColorDefaultBrush', 'ControlStrokeColorDefaultBrush' '#DEDEDE' '#3D3D3D'
        & $set 'ControlFillColorDefaultBrush' '#FFFFFF' '#2D2D2D'
        & $set 'ControlFillColorSecondaryBrush' '#F5F5F5' '#323232'
        & $set 'DividerStrokeColorDefaultBrush' '#DEDEDE' '#3D3D3D'
        & $set 'LayerFillColorDefaultBrush' '#FCFBF8' '#262626'
        & $set 'TextFillColorPrimaryBrush' '#242424' '#FFFFFF'
        & $set 'TextFillColorSecondaryBrush' '#5C5C5C' '#C5C5C5'
        & $set 'TextFillColorTertiaryBrush' '#8A8A8A' '#9A9A9A'
    }
    # The crimson accent of the other tools of the author instead of the accent colour of Windows.
    & $set 'AccentFillColorDefaultBrush', 'AccentButtonBackground', 'AccentButtonBorderBrush' '#B11F4B' '#FD8EA1'
    & $set 'AccentFillColorSecondaryBrush', 'AccentButtonBackgroundPointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'AccentFillColorTertiaryBrush', 'AccentButtonBackgroundPressed' '#CCB11F4B' '#CCFD8EA1'
    & $set 'AccentTextFillColorPrimaryBrush' '#9A1A41' '#FD8EA1'
    & $set 'AccentButtonForeground', 'AccentButtonForegroundPointerOver', 'AccentButtonForegroundPressed', 'TextOnAccentFillColorPrimaryBrush' '#FFFFFF' '#1A1A1A'
    & $set 'CheckBoxCheckBackgroundFillChecked', 'CheckBoxCheckBackgroundStrokeChecked', 'TextControlFocusedBorderBrush', 'ComboBoxBorderBrushFocused' '#B11F4B' '#FD8EA1'
    & $set 'CheckBoxCheckBackgroundFillCheckedPointerOver', 'CheckBoxCheckBackgroundStrokeCheckedPointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'CheckBoxCheckBackgroundFillCheckedPressed', 'CheckBoxCheckBackgroundStrokeCheckedPressed' '#CCB11F4B' '#CCFD8EA1'
    & $set 'AccentFillColorSelectedTextBackgroundBrush' '#B11F4B' '#B11F4B'
    # The selected view (Timeline, Diagnostic log): the accent of the tool, not the accent colour of Windows.
    & $set 'ToggleButtonBackgroundChecked', 'ToggleButtonBorderBrushChecked' '#B11F4B' '#FD8EA1'
    & $set 'ToggleButtonBackgroundCheckedPointerOver', 'ToggleButtonBorderBrushCheckedPointerOver' '#E6B11F4B' '#E6FD8EA1'
    & $set 'ToggleButtonBackgroundCheckedPressed', 'ToggleButtonBorderBrushCheckedPressed' '#CCB11F4B' '#CCFD8EA1'
    & $set 'ToggleButtonForegroundChecked', 'ToggleButtonForegroundCheckedPointerOver', 'ToggleButtonForegroundCheckedPressed' '#FFFFFF' '#1A1A1A'
    & $set 'GctBrand' '#B11F4B' '#B11F4B'
    & $set 'GctBrandText' '#B11F4B' '#FD8EA1'
    & $set 'GctAccentSoft' '#14B11F4B' '#33FD8EA1'
    & $set 'GctSuccess' '#16A34A' '#4ADE80'
    & $set 'GctCaution' '#D97706' '#FBBF24'
    & $set 'GctCritical' '#DC2626' '#F87171'
    & $set 'GctInfoBackground' '#F3F3F3' '#2E2E2E'
    & $set 'GctInfoBorder' '#E0E0E0' '#3D3D3D'
    & $set 'GctCautionBackground' '#FFF7E8' '#33FBBF24'
    & $set 'GctCautionBorder' '#F5D7A1' '#66FBBF24'
    & $set 'GctCriticalBackground' '#FDECEC' '#33F87171'
    & $set 'GctCriticalBorder' '#F4B4B4' '#66F87171'
    & $set 'GctSuccessBackground' '#EAF7EE' '#334ADE80'
    & $set 'GctNeutralBackground' '#F0F0F0' '#333333'
    & $set 'GctHover' '#0A000000' '#14FFFFFF'
    & $set 'GctDiagBackground' '#F7F6F4' '#2A2A2A'
}

function Invoke-GctGuiPump {
    <# Lets the window repaint and handle clicks during a run (the WPF equivalent of DoEvents). #>
    $frame = [Windows.Threading.DispatcherFrame]::new()
    [void][Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::Background,
        [Windows.Threading.DispatcherOperationCallback] { param($f) $f.Continue = $false; $null }, $frame)
    [Windows.Threading.Dispatcher]::PushFrame($frame)
}

function Set-GctGuiNotice {
    <# The coloured box above the progress: Info, Caution or Critical. #>
    param([Parameter(Mandatory = $true)][string]$Text, [ValidateSet('Info', 'Caution', 'Critical')][string]$Level = 'Info')

    $c = $script:Gui.Controls
    $c.Warning.Text = $Text
    $c.WarningIcon.Text = [string][char]$(switch ($Level) { 'Caution' { 0xE7BA } 'Critical' { 0xEA39 } default { 0xE946 } })
    $c.WarningIcon.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, $(switch ($Level) { 'Caution' { 'GctCaution' } 'Critical' { 'GctCritical' } default { 'TextFillColorSecondaryBrush' } }))
    $c.WarningBox.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, $(switch ($Level) { 'Caution' { 'GctCautionBackground' } 'Critical' { 'GctCriticalBackground' } default { 'GctInfoBackground' } }))
    $c.WarningBox.SetResourceReference([Windows.Controls.Border]::BorderBrushProperty, $(switch ($Level) { 'Caution' { 'GctCautionBorder' } 'Critical' { 'GctCriticalBorder' } default { 'GctInfoBorder' } }))
}

function Get-GctGuiStatusBrushes {
    <# Foreground and background resources of a status. #>
    param([string]$Status)
    switch ($Status) {
        'Passed' { 'GctSuccess', 'GctSuccessBackground' }
        'Failed' { 'GctCritical', 'GctCriticalBackground' }
        { $_ -in 'Warning', 'Blocked' } { 'GctCaution', 'GctCautionBackground' }
        'Running' { 'GctBrandText', 'GctAccentSoft' }
        default { 'TextFillColorSecondaryBrush', 'GctNeutralBackground' }
    }
}

function Set-GctGuiStatus {
    <# The status pill of the progress: Ready, Running, Passed, Warning, Blocked or Failed. #>
    param([string]$Text, [string]$Status)

    $c = $script:Gui.Controls
    $c.Status.Text = $Text
    $pair = Get-GctGuiStatusBrushes -Status $Status
    $c.Status.SetResourceReference([Windows.Controls.TextBlock]::ForegroundProperty, $pair[0])
    $c.StatusPill.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, $pair[1])
}

#endregion

#region Progress: timeline of the stages, diagnostics of each check, diagnostic log -------------------

$script:GuiStatusGlyphs = @{ Passed = 0xE73E; Warning = 0xE7BA; Failed = 0xEA39; Blocked = 0xE733; Skipped = 0xE72A; Info = 0xE946 }
$script:GuiStatusRank = @{ Skipped = 0; Passed = 1; Warning = 2; Blocked = 3; Failed = 4 }
$script:GuiLogLevels = @{ Passed = 'OK'; Warning = 'WARN'; Failed = 'FAIL'; Blocked = 'BLOCK'; Skipped = 'SKIP' }
$script:GuiMono = 'Cascadia Mono, Consolas'

function New-GctGuiElement {
    <# A piece of XAML built in code (the root element gets the WPF namespaces; its x:Name are found with FindName). #>
    param([Parameter(Mandatory = $true)][string]$Xaml)
    $m = [regex]::Match($Xaml, '<\w+')
    $full = $Xaml.Substring(0, $m.Index + $m.Length) + ' xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"' + $Xaml.Substring($m.Index + $m.Length)
    return [Windows.Markup.XamlReader]::Parse($full)
}

function Format-GctGuiDuration {
    param([double]$Milliseconds)
    $inv = [Globalization.CultureInfo]::InvariantCulture
    if ($Milliseconds -lt 1000) { return [string]::Format($inv, '{0:0} ms', $Milliseconds) }
    return Format-GctDuration ($Milliseconds / 1000)
}

function Get-GctGuiLocalTime {
    <# HH:mm:ss.fff (local) of an ISO UTC time of a step or an exchange, else now. #>
    param([AllowEmptyString()][string]$Utc)
    $t = [DateTimeOffset]::Now
    if ($Utc) { try { $t = [DateTimeOffset]::Parse($Utc, [Globalization.CultureInfo]::InvariantCulture).ToLocalTime() } catch { $t = [DateTimeOffset]::Now } }
    return $t.ToString('HH:mm:ss.fff')
}

function Set-GctGuiResource {
    <# Colour of an element from a resource of the theme (it follows a change of theme). #>
    param([Parameter(Mandatory = $true)][Windows.DependencyObject]$Element, [Parameter(Mandatory = $true)][Windows.DependencyProperty]$Property, [Parameter(Mandatory = $true)][string]$Key)
    $Element.SetResourceReference($Property, $Key)
}

function Add-GctGuiStage {
    <# A stage of the timeline: its node on the rail, its title, the summary of its checks; a click folds it. #>
    param([Parameter(Mandatory = $true)][string]$Title, [string]$Number)

    $g = $script:Gui
    Complete-GctGuiStage
    $e = New-GctGuiElement @"
<Grid Margin="0,0,0,2">
  <Grid.ColumnDefinitions><ColumnDefinition Width="30"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
  <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
  <Rectangle x:Name="Rail" Grid.RowSpan="2" Width="2" HorizontalAlignment="Center" Margin="0,26,0,0" Fill="{DynamicResource DividerStrokeColorDefaultBrush}"/>
  <Border x:Name="Node" Width="24" Height="24" CornerRadius="12" VerticalAlignment="Top" HorizontalAlignment="Center" Margin="0,3,0,0">
    <TextBlock x:Name="NodeText" FontSize="11" FontWeight="SemiBold" HorizontalAlignment="Center" VerticalAlignment="Center"/>
  </Border>
  <Border x:Name="Header" Grid.Column="1" Background="Transparent" Cursor="Hand" Padding="8,4,6,5" CornerRadius="6">
    <Grid>
      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
      <StackPanel>
        <TextBlock x:Name="Title" FontSize="14" FontWeight="SemiBold" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
        <TextBlock x:Name="Summary" FontSize="12" Margin="0,1,0,0" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
      </StackPanel>
      <TextBlock x:Name="Time" Grid.Column="1" FontSize="11" Margin="10,3,10,0" VerticalAlignment="Top" Foreground="{DynamicResource TextFillColorTertiaryBrush}"/>
      <TextBlock x:Name="Chevron" Grid.Column="2" FontFamily="$($script:GuiIconFont)" FontSize="10" Margin="0,5,0,0" VerticalAlignment="Top" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
    </Grid>
  </Border>
  <StackPanel x:Name="Checks" Grid.Column="1" Grid.Row="1" Margin="8,2,0,12"/>
</Grid>
"@
    $stage = [pscustomobject]@{
        Root = $e; Title = $Title; Number = $Number; Header = $e.FindName('Header'); Node = $e.FindName('Node'); NodeText = $e.FindName('NodeText')
        Summary = $e.FindName('Summary'); Time = $e.FindName('Time'); Chevron = $e.FindName('Chevron'); Panel = $e.FindName('Checks')
        Checks = [Collections.Generic.List[object]]::new(); Expanded = $true; Running = $true; Worst = $null; DurationMs = 0.0
    }
    $e.FindName('Title').Text = $Title
    $stage.Time.Text = (Get-Date).ToString('HH:mm:ss')
    $stage.Header.Tag = $stage
    $stage.Header.Add_MouseLeftButtonUp({ param($source, $e) Set-GctGuiStageExpanded -Stage $source.Tag -Expanded (-not $source.Tag.Expanded) })
    $stage.Header.Add_MouseEnter({ param($source, $e) $source.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, 'GctHover') })
    $stage.Header.Add_MouseLeave({ param($source, $e) $source.Background = [Windows.Media.Brushes]::Transparent })
    Set-GctGuiStageNode -Stage $stage
    Set-GctGuiStageExpanded -Stage $stage -Expanded $true
    Update-GctGuiStageSummary -Stage $stage
    [void]$g.Controls.Timeline.Children.Add($e)
    $g.Stages.Add($stage)
    $g.CurrentStage = $stage
    Add-GctGuiEvent ([pscustomobject]@{ Kind = 'Stage'; Time = Get-GctGuiLocalTime; Text = $(if ($Number) { "$Number  $Title" } else { $Title }) })
    return $stage
}

function Set-GctGuiStageNode {
    <# The node of a stage: its number on the accent colour while it runs, then the icon of its worst status. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Stage)
    $t = $Stage.NodeText
    if ($Stage.Running) {
        $t.FontFamily = [Windows.Media.FontFamily]::new('Segoe UI Variable Text, Segoe UI')
        $t.FontSize = 11
        $t.Text = if ($Stage.Number) { ($Stage.Number -split '/')[0] } else { [string][char]0x2022 }
        Set-GctGuiResource $Stage.Node ([Windows.Controls.Border]::BackgroundProperty) 'GctBrand'
        Set-GctGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) 'TextOnAccentFillColorPrimaryBrush'
        return
    }
    $status = if ($Stage.Worst) { $Stage.Worst } else { 'Skipped' }
    $pair = Get-GctGuiStatusBrushes -Status $status
    $t.FontFamily = [Windows.Media.FontFamily]::new($script:GuiIconFont)
    $t.FontSize = 11
    $t.Text = [string][char]$script:GuiStatusGlyphs[$status]
    Set-GctGuiResource $Stage.Node ([Windows.Controls.Border]::BackgroundProperty) $pair[1]
    Set-GctGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) $pair[0]
}

function Set-GctGuiStageExpanded {
    param([Parameter(Mandatory = $true)][pscustomobject]$Stage, [bool]$Expanded)
    $Stage.Expanded = $Expanded
    $Stage.Panel.Visibility = if ($Expanded) { 'Visible' } else { 'Collapsed' }
    $Stage.Chevron.Text = [string][char]$(if ($Expanded) { 0xE70E } else { 0xE70D })
}

function Update-GctGuiStageSummary {
    <# 2 failed, 1 warning, 3 passed - 1.4 s (or Running... while it runs). #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Stage)
    $parts = foreach ($s in 'Failed', 'Blocked', 'Warning', 'Passed', 'Skipped') {
        $n = @($Stage.Checks | Where-Object { $_.Step.Status -eq $s }).Count
        if ($n) { '{0} {1}' -f $n, $(switch ($s) { 'Warning' { if ($n -gt 1) { 'warnings' } else { 'warning' } } default { $s.ToLowerInvariant() } }) }
    }
    $count = $Stage.Checks.Count
    $text = if ($count) { "$(@($parts) -join ', ') $([char]0x00B7) $count check$(if ($count -gt 1) { 's' }) $([char]0x00B7) $(Format-GctGuiDuration $Stage.DurationMs)" } else { '' }
    if ($Stage.Running) { $text = if ($text) { "Running $([char]0x00B7) $text" } else { 'Running...' } }
    elseif (-not $count) { $text = 'Nothing recorded' }
    $Stage.Summary.Text = $text
}

function Complete-GctGuiStage {
    <# The current stage is over: its node shows its worst status; a stage without problem folds (unless the diagnostics are shown). #>
    $g = $script:Gui
    $stage = $g.CurrentStage
    if (-not $stage -or -not $stage.Running) { return }
    $stage.Running = $false
    Set-GctGuiStageNode -Stage $stage
    Update-GctGuiStageSummary -Stage $stage
    if ($stage.Checks.Count -and $stage.Worst -in 'Passed', 'Skipped' -and -not $g.Diagnostics) { Set-GctGuiStageExpanded -Stage $stage -Expanded $false }
}

function Add-GctGuiCheck {
    <# One check in the current stage: icon, name, message, time and duration; its diagnostics fold out under it. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Step, [object[]]$Exchanges = @())

    $g = $script:Gui
    if (-not $g) { return }
    $stage = $g.CurrentStage
    if (-not $stage) { $stage = Add-GctGuiStage -Title ([string]$Step.Stage) }
    $e = New-GctGuiElement @"
<StackPanel Margin="0,4,0,4">
  <Grid>
    <Grid.ColumnDefinitions><ColumnDefinition Width="22"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
    <TextBlock x:Name="Glyph" FontFamily="$($script:GuiIconFont)" FontSize="13" Margin="0,3,0,0" VerticalAlignment="Top"/>
    <StackPanel Grid.Column="1">
      <TextBlock x:Name="Name" FontSize="13" FontWeight="SemiBold" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
      <TextBlock x:Name="Message" FontSize="12.5" Margin="0,1,0,0" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
      <TextBlock x:Name="Link" FontSize="12" Margin="0,3,0,0" Cursor="Hand" Foreground="{DynamicResource AccentTextFillColorPrimaryBrush}"/>
    </StackPanel>
    <TextBlock x:Name="Time" Grid.Column="2" FontSize="11" Margin="10,2,0,0" TextAlignment="Right" Foreground="{DynamicResource TextFillColorTertiaryBrush}"/>
  </Grid>
  <Border x:Name="Diag" Visibility="Collapsed" Margin="22,6,0,4" Padding="12,10,12,10" CornerRadius="6" BorderThickness="1"
          Background="{DynamicResource GctDiagBackground}" BorderBrush="{DynamicResource CardStrokeColorDefaultBrush}">
    <StackPanel x:Name="DiagBody"/>
  </Border>
</StackPanel>
"@
    $status = [string]$Step.Status
    $pair = Get-GctGuiStatusBrushes -Status $status
    $glyph = $e.FindName('Glyph')
    $glyph.Text = [string][char]$script:GuiStatusGlyphs[$status]
    Set-GctGuiResource $glyph ([Windows.Controls.TextBlock]::ForegroundProperty) $(if ($status -eq 'Skipped') { 'TextFillColorTertiaryBrush' } else { $pair[0] })
    $e.FindName('Name').Text = [string]$Step.Name
    $e.FindName('Message').Text = [string]$Step.Message
    $e.FindName('Time').Text = "$((Get-GctGuiLocalTime ([string]$Step.TimestampUtc)).Substring(0, 8))`n$(Format-GctGuiDuration ([double]$Step.DurationMs))"
    $details = @(if ($Step.Details) { $Step.Details.Keys })
    $check = [pscustomobject]@{
        Root = $e; Step = $Step; Exchanges = @($Exchanges); Diag = $e.FindName('Diag'); Body = $e.FindName('DiagBody'); Link = $e.FindName('Link')
        Built = $false; Open = $false; Stage = $stage; HasDiagnostics = [bool]($details.Count -or @($Exchanges).Count)
    }
    if ($check.HasDiagnostics) {
        $check.Link.Tag = $check
        $check.Link.Add_MouseLeftButtonUp({ param($source, $e) Set-GctGuiCheckOpen -Check $source.Tag -Open (-not $source.Tag.Open) })
    }
    else { $check.Link.Visibility = 'Collapsed' }
    Set-GctGuiCheckOpen -Check $check -Open $false
    [void]$stage.Panel.Children.Add($e)
    $stage.Checks.Add($check)
    $stage.DurationMs += [double]$Step.DurationMs
    if (-not $stage.Worst -or $script:GuiStatusRank[$status] -gt $script:GuiStatusRank[$stage.Worst]) { $stage.Worst = $status }
    Update-GctGuiStageSummary -Stage $stage
    $g.Checks.Add($check)
    if ($g.Diagnostics -and $check.HasDiagnostics) { Set-GctGuiCheckOpen -Check $check -Open $true }
    Add-GctGuiEvent ([pscustomobject]@{ Kind = 'Check'; Time = Get-GctGuiLocalTime ([string]$Step.TimestampUtc); Step = $Step; Exchanges = @($Exchanges) })
    $g.Lines.Add("[$status] $($Step.Name): $($Step.Message)")
    $g.Controls.LogEmpty.Visibility = 'Collapsed'
    if ($g.Running) { $g.Controls.LogScroll.ScrollToEnd() }
    Invoke-GctGuiPump
}

function Set-GctGuiCheckOpen {
    <# Shows or hides the diagnostics of a check (built the first time). #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Check, [bool]$Open)
    if (-not $Check.HasDiagnostics) { return }
    if ($Open -and -not $Check.Built) { New-GctGuiDiagnosticPanel -Check $Check; $Check.Built = $true }
    $Check.Open = $Open
    $Check.Diag.Visibility = if ($Open) { 'Visible' } else { 'Collapsed' }
    $n = $Check.Exchanges.Count
    $what = @(
        if ($n) { "$n HTTP exchange$(if ($n -gt 1) { 's' })" }
        $k = @($Check.Step.Details.Keys).Count; if ($k) { "$k detail$(if ($k -gt 1) { 's' })" }
    ) -join ', '
    $Check.Link.Text = "$(if ($Open) { [char]0x25BE } else { [char]0x25B8 }) Diagnostics ($what)"
}

function Get-GctGuiExchangeLine {
    <# One HTTP exchange on one line: #12 POST https://... -> 401 Unauthorized (230 ms) [operation] label, server. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Exchange)
    $code = if ($null -ne $Exchange.StatusCode) { "$($Exchange.StatusCode) $($Exchange.Reason)".TrimEnd() } else { 'no response' }
    $more = @(
        if ($Exchange.Operation) { "[$($Exchange.Operation)]" }
        if ($Exchange.Label) { $Exchange.Label }
        if ($Exchange.Server) { $Exchange.Server }
        if ([int]$Exchange.Repeated -gt 1) { "sent $($Exchange.Repeated) times" }
    ) -join ' - '
    return ('#{0} {1} {2} -> {3} ({4}){5}' -f $Exchange.Sequence, $Exchange.Method, $Exchange.Url, $code, (Format-GctGuiDuration ([double]$Exchange.DurationMs)), $(if ($more) { " $more" } else { '' }))
}

function Format-GctGuiCheckText {
    <# The diagnostics of a check as text (Copy). #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Check)
    $s = $Check.Step
    $b = [Text.StringBuilder]::new()
    [void]$b.AppendLine("$($s.Status) - $($s.Stage) - $($s.Name)")
    [void]$b.AppendLine([string]$s.Message)
    foreach ($k in @($s.Details.Keys)) { $v = $s.Details[$k]; if ($v -is [array]) { $v = $v -join ', ' }; [void]$b.AppendLine("  $k = $v") }
    foreach ($x in $Check.Exchanges) {
        [void]$b.AppendLine('')
        [void]$b.AppendLine((Get-GctGuiExchangeLine -Exchange $x))
        if ($x.Note) { [void]$b.AppendLine("Note: $($x.Note)") }
        [void]$b.AppendLine('--- request ---'); [void]$b.AppendLine([string]$x.Request)
        [void]$b.AppendLine('--- response ---'); [void]$b.AppendLine([string]$x.Response)
    }
    return $b.ToString()
}

function New-GctGuiDiagnosticPanel {
    <# The diagnostics of a check: its details (key and value), then each HTTP exchange; a click shows the request and the response. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Check)

    $body = $Check.Body
    $step = $Check.Step
    $title = {
        param([string]$Text)
        $t = [Windows.Controls.TextBlock]::new()
        $t.Text = $Text
        $t.FontSize = 11; $t.FontWeight = [Windows.FontWeights]::SemiBold
        $t.Margin = [Windows.Thickness]::new(0, 0, 0, 4)
        Set-GctGuiResource $t ([Windows.Controls.TextBlock]::ForegroundProperty) 'TextFillColorTertiaryBrush'
        $t
    }
    $keys = @($step.Details.Keys)
    if ($keys.Count) {
        [void]$body.Children.Add((& $title 'DETAILS'))
        $grid = [Windows.Controls.Grid]::new()
        $c0 = [Windows.Controls.ColumnDefinition]::new(); $c0.Width = [Windows.GridLength]::Auto
        $c1 = [Windows.Controls.ColumnDefinition]::new(); $c1.Width = [Windows.GridLength]::new(1, 'Star')
        $grid.ColumnDefinitions.Add($c0); $grid.ColumnDefinitions.Add($c1)
        $row = 0
        foreach ($k in $keys) {
            $v = $step.Details[$k]; if ($v -is [array]) { $v = $v -join ', ' }
            $grid.RowDefinitions.Add([Windows.Controls.RowDefinition]::new())
            $kt = [Windows.Controls.TextBlock]::new(); $kt.Text = [string]$k; $kt.FontSize = 12; $kt.Margin = [Windows.Thickness]::new(0, 1, 14, 1)
            Set-GctGuiResource $kt ([Windows.Controls.TextBlock]::ForegroundProperty) 'TextFillColorSecondaryBrush'
            # A read-only box without border: the value can be selected and copied.
            $vt = [Windows.Controls.TextBox]::new(); $vt.Text = [string]$v; $vt.IsReadOnly = $true; $vt.TextWrapping = 'Wrap'; $vt.BorderThickness = [Windows.Thickness]::new(0)
            $vt.Background = [Windows.Media.Brushes]::Transparent; $vt.Padding = [Windows.Thickness]::new(0); $vt.Margin = [Windows.Thickness]::new(0, 1, 0, 1); $vt.MinHeight = 0
            $vt.FontFamily = [Windows.Media.FontFamily]::new($script:GuiMono); $vt.FontSize = 11.5
            Set-GctGuiResource $vt ([Windows.Controls.Control]::ForegroundProperty) 'TextFillColorPrimaryBrush'
            [Windows.Controls.Grid]::SetRow($kt, $row); [Windows.Controls.Grid]::SetRow($vt, $row); [Windows.Controls.Grid]::SetColumn($vt, 1)
            [void]$grid.Children.Add($kt); [void]$grid.Children.Add($vt)
            $row++
        }
        [void]$body.Children.Add($grid)
    }
    if ($Check.Exchanges.Count) {
        $h = & $title "HTTP EXCHANGES ($($Check.Exchanges.Count)) - click one to see the request and the response"
        if ($keys.Count) { $h.Margin = [Windows.Thickness]::new(0, 12, 0, 4) }
        [void]$body.Children.Add($h)
        foreach ($x in $Check.Exchanges) {
            $e = New-GctGuiElement @"
<StackPanel Margin="0,2,0,2">
  <Border x:Name="Head" Cursor="Hand" Background="Transparent" Padding="4,3,4,3" CornerRadius="4">
    <Grid>
      <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
      <Border x:Name="Code" CornerRadius="4" Padding="6,0,6,1" Margin="0,0,8,0" VerticalAlignment="Top" MinWidth="38">
        <TextBlock x:Name="CodeText" FontSize="11.5" FontWeight="SemiBold" HorizontalAlignment="Center"/>
      </Border>
      <TextBlock x:Name="Method" Grid.Column="1" FontFamily="$($script:GuiMono)" FontSize="11.5" Margin="0,1,8,0" Foreground="{DynamicResource TextFillColorSecondaryBrush}"/>
      <StackPanel Grid.Column="2">
        <TextBlock x:Name="Url" FontFamily="$($script:GuiMono)" FontSize="11.5" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorPrimaryBrush}"/>
        <TextBlock x:Name="Info" FontSize="11.5" TextWrapping="Wrap" Foreground="{DynamicResource TextFillColorTertiaryBrush}"/>
      </StackPanel>
      <TextBlock x:Name="Ms" Grid.Column="3" FontSize="11" Margin="8,1,0,0" Foreground="{DynamicResource TextFillColorTertiaryBrush}"/>
    </Grid>
  </Border>
  <TextBox x:Name="Raw" Visibility="Collapsed" IsReadOnly="True" TextWrapping="Wrap" FontFamily="$($script:GuiMono)" FontSize="11"
           MaxHeight="380" VerticalScrollBarVisibility="Auto" Margin="4,4,0,8" Padding="8,6"/>
</StackPanel>
"@
            $code = if ($null -ne $x.StatusCode) { [int]$x.StatusCode } else { 0 }
            $e.FindName('CodeText').Text = if ($code) { [string]$code } else { '---' }
            $kind = if ($code -ge 200 -and $code -lt 300) { 'Passed' } elseif ($code -ge 300 -and $code -lt 400) { 'Neutral' } elseif ($code -ge 400 -and $code -lt 500) { 'Warning' } else { 'Failed' }
            $pair = Get-GctGuiStatusBrushes -Status $kind
            Set-GctGuiResource $e.FindName('Code') ([Windows.Controls.Border]::BackgroundProperty) $pair[1]
            Set-GctGuiResource $e.FindName('CodeText') ([Windows.Controls.TextBlock]::ForegroundProperty) $pair[0]
            $e.FindName('Method').Text = [string]$x.Method
            $e.FindName('Url').Text = [string]$x.Url
            $info = @(
                if ($null -ne $x.StatusCode -and $x.Reason) { [string]$x.Reason } elseif ($null -eq $x.StatusCode) { 'no response' }
                if ($x.Operation) { [string]$x.Operation }
                if ($x.Label) { [string]$x.Label }
                if ($x.Server) { [string]$x.Server }
                if ($x.Note) { [string]$x.Note }
            ) -join " $([char]0x00B7) "
            $e.FindName('Info').Text = $info
            $e.FindName('Ms').Text = Format-GctGuiDuration ([double]$x.DurationMs)
            $raw = $e.FindName('Raw')
            $raw.Text = "REQUEST`n$([string]$x.Request)`n`nRESPONSE`n$([string]$x.Response)"
            $head = $e.FindName('Head')
            $head.Tag = $raw
            $head.Add_MouseLeftButtonUp({ param($source, $e) $source.Tag.Visibility = if ($source.Tag.Visibility -eq 'Visible') { 'Collapsed' } else { 'Visible' } })
            $head.Add_MouseEnter({ param($source, $e) $source.SetResourceReference([Windows.Controls.Border]::BackgroundProperty, 'GctHover') })
            $head.Add_MouseLeave({ param($source, $e) $source.Background = [Windows.Media.Brushes]::Transparent })
            [void]$body.Children.Add($e)
        }
    }
    $copy = [Windows.Controls.Button]::new()
    $copy.Content = 'Copy these diagnostics'
    $copy.HorizontalAlignment = 'Left'
    $copy.Margin = [Windows.Thickness]::new(0, 10, 0, 0)
    $copy.Padding = [Windows.Thickness]::new(10, 3, 10, 4)
    $copy.FontSize = 12
    $copy.Tag = $Check
    $copy.Add_Click({ param($source, $e) Copy-GctGuiText -Text (Format-GctGuiCheckText -Check $source.Tag) -What 'The diagnostics of the check' })
    [void]$body.Children.Add($copy)
}

function Copy-GctGuiText {
    param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text, [string]$What = 'The text')
    try { [Windows.Clipboard]::SetText($Text); Set-GctGuiStatus "$What copied to the clipboard" 'Ready' }
    catch { Add-GctGuiLine 'Warn' "The clipboard is not available: $($_.Exception.Message)" }
}

function Format-GctGuiEvent {
    <# One event of the diagnostic log; -Full adds the details of each check and the request and response of each exchange. #>
    param([Parameter(Mandatory = $true)][pscustomobject]$Event, [switch]$Full)
    $b = [Text.StringBuilder]::new()
    $pad = ' ' * 21
    switch ($Event.Kind) {
        'Stage' { [void]$b.AppendLine(''); [void]$b.AppendLine(('{0}  {1} {2} {1}{1}{1}' -f $Event.Time, [string][char]0x2500, $Event.Text)) }
        'Note' { [void]$b.AppendLine(('{0}  {1,-5}  {2}' -f $Event.Time, $Event.Level, $Event.Text)) }
        'Check' {
            foreach ($x in $Event.Exchanges) {
                [void]$b.AppendLine(('{0}  {1,-5}  {2}' -f (Get-GctGuiLocalTime ([string]$x.TimestampUtc)), 'HTTP', (Get-GctGuiExchangeLine -Exchange $x)))
                if ($Full) {
                    if ($x.Note) { [void]$b.AppendLine("$pad  Note: $($x.Note)") }
                    foreach ($part in @(@('>', [string]$x.Request), @('<', [string]$x.Response))) {
                        foreach ($line in ($part[1] -split "`r?`n")) { [void]$b.AppendLine("$pad$($part[0]) $line") }
                    }
                }
            }
            $s = $Event.Step
            [void]$b.AppendLine(('{0}  {1,-5}  {2}: {3}' -f $Event.Time, $script:GuiLogLevels[[string]$s.Status], $s.Name, $s.Message))
            if ($Full) { foreach ($k in @($s.Details.Keys)) { $v = $s.Details[$k]; if ($v -is [array]) { $v = $v -join ', ' }; [void]$b.AppendLine("$pad  $k = $v") } }
        }
    }
    return $b.ToString()
}

function Add-GctGuiEvent {
    param([Parameter(Mandatory = $true)][pscustomobject]$Event)
    $g = $script:Gui
    $g.Events.Add($Event)
    $g.Controls.DiagLog.AppendText((Format-GctGuiEvent -Event $Event -Full:$g.Diagnostics))
    if ($g.Controls.DiagLog.Visibility -eq 'Visible') { $g.Controls.DiagLog.ScrollToEnd() }
}

function Update-GctGuiLog {
    <# Rebuilds the diagnostic log (after the Diagnostics switch changed). #>
    $g = $script:Gui
    $b = [Text.StringBuilder]::new()
    foreach ($ev in $g.Events) { [void]$b.Append((Format-GctGuiEvent -Event $ev -Full:$g.Diagnostics)) }
    $g.Controls.DiagLog.Text = $b.ToString()
}

function Set-GctGuiDiagnosticMode {
    <# The Diagnostics switch: every check shows its diagnostics, every stage unfolds, the log shows the requests and responses. #>
    param([bool]$On)
    $g = $script:Gui
    $g.Diagnostics = $On
    foreach ($check in $g.Checks) { Set-GctGuiCheckOpen -Check $check -Open $On }
    foreach ($stage in $g.Stages) {
        if ($On) { Set-GctGuiStageExpanded -Stage $stage -Expanded $true }
        elseif (-not $stage.Running -and $stage.Checks.Count -and $stage.Worst -in 'Passed', 'Skipped') { Set-GctGuiStageExpanded -Stage $stage -Expanded $false }
    }
    Update-GctGuiLog
}

function Set-GctGuiView {
    <# Timeline or Diagnostic log. #>
    param([ValidateSet('Timeline', 'Log')][string]$View)
    $c = $script:Gui.Controls
    $c.ViewTimeline.IsChecked = $View -eq 'Timeline'
    $c.ViewLog.IsChecked = $View -eq 'Log'
    $c.LogScroll.Visibility = if ($View -eq 'Timeline') { 'Visible' } else { 'Collapsed' }
    $c.DiagLog.Visibility = if ($View -eq 'Log') { 'Visible' } else { 'Collapsed' }
    if ($View -eq 'Log') { $c.LogEmpty.Visibility = 'Collapsed'; $c.DiagLog.ScrollToEnd() }
    elseif (-not $script:Gui.Events.Count) { $c.LogEmpty.Visibility = 'Visible' }
}

function Add-GctGuiLine {
    <#
        A line sent by the engine: a stage header ([n/t] Title) starts a stage of the timeline; any other line is a
        note in the current stage (sign-in in progress, cancellation, a problem of the settings...). The checks
        themselves arrive through Add-GctGuiCheck.
    #>
    param([string]$Status, [string]$Text)

    $g = $script:Gui
    if (-not $g) { return }
    $g.Lines.Add("[$Status] $Text")
    if ($Status -eq 'Step') {
        $m = [regex]::Match($Text, '^\[(\d+/\d+)\]\s*(.*)$')
        if ($m.Success) { [void](Add-GctGuiStage -Title $m.Groups[2].Value -Number $m.Groups[1].Value) } else { [void](Add-GctGuiStage -Title $Text) }
    }
    else {
        $stage = $g.CurrentStage
        if (-not $stage) { $stage = Add-GctGuiStage -Title 'Run' }
        $map = @{ Ok = 'Passed'; Warn = 'Warning'; Fail = 'Failed'; Skip = 'Skipped'; Block = 'Blocked' }
        $status = if ($map.ContainsKey($Status)) { $map[$Status] } else { 'Info' }
        $row = [Windows.Controls.Grid]::new()
        $row.Margin = [Windows.Thickness]::new(0, 3, 0, 3)
        $c0 = [Windows.Controls.ColumnDefinition]::new(); $c0.Width = [Windows.GridLength]::new(22)
        $c1 = [Windows.Controls.ColumnDefinition]::new(); $c1.Width = [Windows.GridLength]::new(1, 'Star')
        $c2 = [Windows.Controls.ColumnDefinition]::new(); $c2.Width = [Windows.GridLength]::Auto
        $row.ColumnDefinitions.Add($c0); $row.ColumnDefinitions.Add($c1); $row.ColumnDefinitions.Add($c2)
        $glyph = [Windows.Controls.TextBlock]::new()
        $glyph.FontFamily = [Windows.Media.FontFamily]::new($script:GuiIconFont); $glyph.FontSize = 12; $glyph.Margin = [Windows.Thickness]::new(0, 3, 0, 0)
        $glyph.Text = [string][char]$script:GuiStatusGlyphs[$status]
        Set-GctGuiResource $glyph ([Windows.Controls.TextBlock]::ForegroundProperty) $(if ($status -eq 'Info') { 'TextFillColorTertiaryBrush' } else { (Get-GctGuiStatusBrushes -Status $status)[0] })
        $label = [Windows.Controls.TextBlock]::new()
        $label.Text = $Text; $label.TextWrapping = 'Wrap'; $label.FontSize = 12.5; $label.FontStyle = if ($status -eq 'Info') { [Windows.FontStyles]::Italic } else { [Windows.FontStyles]::Normal }
        Set-GctGuiResource $label ([Windows.Controls.TextBlock]::ForegroundProperty) $(if ($status -eq 'Info') { 'TextFillColorSecondaryBrush' } else { 'TextFillColorPrimaryBrush' })
        $time = [Windows.Controls.TextBlock]::new()
        $time.Text = (Get-Date).ToString('HH:mm:ss'); $time.FontSize = 11; $time.Margin = [Windows.Thickness]::new(10, 2, 0, 0)
        Set-GctGuiResource $time ([Windows.Controls.TextBlock]::ForegroundProperty) 'TextFillColorTertiaryBrush'
        [Windows.Controls.Grid]::SetColumn($label, 1); [Windows.Controls.Grid]::SetColumn($time, 2)
        [void]$row.Children.Add($glyph); [void]$row.Children.Add($label); [void]$row.Children.Add($time)
        [void]$stage.Panel.Children.Add($row)
        if ($status -in 'Failed', 'Warning', 'Blocked') { Set-GctGuiStageExpanded -Stage $stage -Expanded $true }
        $level = @{ Passed = 'OK'; Warning = 'WARN'; Failed = 'FAIL'; Skipped = 'SKIP'; Blocked = 'BLOCK'; Info = 'INFO' }[$status]
        Add-GctGuiEvent ([pscustomobject]@{ Kind = 'Note'; Time = Get-GctGuiLocalTime; Level = $level; Text = $Text })
    }
    $g.Controls.LogEmpty.Visibility = 'Collapsed'
    if ($g.Running -or $Status -ne 'Step') { $g.Controls.LogScroll.ScrollToEnd() }
    Invoke-GctGuiPump
}

function Clear-GctGuiProgress {
    $g = $script:Gui
    $g.Controls.Timeline.Children.Clear()
    $g.Stages.Clear()
    $g.Checks.Clear()
    $g.Events.Clear()
    $g.Lines.Clear()
    $g.CurrentStage = $null
    $g.Controls.DiagLog.Clear()
    $g.CodePage = $null
    $g.Controls.CodeBanner.Visibility = 'Collapsed'
    $g.Controls.ResultPanel.Visibility = 'Collapsed'
    $g.Controls.LogEmpty.Visibility = if ($g.Controls.LogScroll.Visibility -eq 'Visible') { 'Visible' } else { 'Collapsed' }
}

#endregion

#region Window -------------------------------------------------------------------------------------------

function Show-GctTestGui {
    <#
    .SYNOPSIS
        Opens the window. Default configuration: config\GraphConnectivityTests.config.psd1 of the tool folder.
    #>
    [CmdletBinding()]
    param([hashtable]$Configuration)

    if (-not $Configuration) { $Configuration = Import-GctConfiguration }
    $window = New-GctTestForm -Configuration $Configuration
    # Ctrl+C in the console would stop the command that owns the window: the window then could not
    # run any of its PowerShell handlers. Ctrl+C is ignored while the window is open.
    $previousCtrlC = $null
    try { if (-not [Console]::IsInputRedirected) { $previousCtrlC = [Console]::TreatControlCAsInput; [Console]::TreatControlCAsInput = $true } } catch { $previousCtrlC = $null }
    try {
        [void]$window.Form.ShowDialog()
    }
    finally {
        if ($null -ne $previousCtrlC) { try { [Console]::TreatControlCAsInput = $previousCtrlC } catch { } }
        $script:Gui = $null
    }
}

#endregion

# SIG # Begin signature block
# MIImKQYJKoZIhvcNAQcCoIImGjCCJhYCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBbPcNxvUPeVvJY
# MwTjAJdyNenH9mOHw09l3XjdZUZ5aqCCDW8wggZyMIIEWqADAgECAghkM1HTxzif
# CDANBgkqhkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMx
# EDAOBgNVBAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8G
# A1UEAwwoU1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAe
# Fw0xNjA2MjQyMDQ0MzBaFw0zMTA2MjQyMDQ0MzBaMHgxCzAJBgNVBAYTAlVTMQ4w
# DAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENv
# cnAxNDAyBgNVBAMMK1NTTC5jb20gQ29kZSBTaWduaW5nIEludGVybWVkaWF0ZSBD
# QSBSU0EgUjEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCfgxNzqrDG
# bSHL24t6h3TQcdyOl3Ka5LuINLTdgAPGL0WkdJq/Hg9Q6p5tePOf+lEmqT2d0bKU
# Vz77OYkbkStW72fL5gvjDjmMxjX0jD3dJekBrBdCfVgWQNz51ShEHZVkMGE6ZPKX
# 13NMfXsjAm3zdetVPW+qLcSvvnSsXf5qtvzqXHnpD0OctVIFD+8+sbGP0EmtpuNC
# GVQ/8y8Ooct8/hP5IznaJRy4PgBKOm8yMDdkHseudQfYVdIYyQ6KvKNc8HwKp4WB
# wg6vj5lc02AlvINaaRwlE81y9eucgJvcLGfE3ckJmNVz68Qho+Uyjj4vUpjGYDdk
# jLJvSlRyGMwnh/rNdaJjIUy1PWT9K6abVa8mTGC0uVz+q0O9rdATZlAfC9KJpv/X
# gAbxwxECMzNhF/dWH44vO2jnFfF3VkopngPawismYTJboFblSSmNNqf1x1KiVgMg
# Lzh4gL32Bq5BNMuURb2bx4kYHwu6/6muakCZE93vUN8BuvIE1tAx3zQ4XldbyDge
# VtSsSKbt//m4wTvtwiS+RGCnd83VPZhZtEPqqmB9zcLlL/Hr9dQg1Zc0bl0EawUR
# 0tOSjAknRO1PNTFGfnQZBWLsiePqI3CY5NEv1IoTGEaTZeVYc9NMPSd6Ij/D+KNV
# t/nmh4LsRR7Fbjp8sU65q2j3m2PVkUG8qQIDAQABo4H7MIH4MA8GA1UdEwEB/wQF
# MAMBAf8wHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKSle44gCUNplkwMAYIKwYBBQUH
# AQEEJDAiMCAGCCsGAQUFBzABhhRodHRwOi8vb2NzcHMuc3NsLmNvbTARBgNVHSAE
# CjAIMAYGBFUdIAAwEwYDVR0lBAwwCgYIKwYBBQUHAwMwOwYDVR0fBDQwMjAwoC6g
# LIYqaHR0cDovL2NybHMuc3NsLmNvbS9zc2wuY29tLXJzYS1Sb290Q0EuY3JsMB0G
# A1UdDgQWBBRUwv4QlQCTzWr158DX2bJLuI8M4zAOBgNVHQ8BAf8EBAMCAYYwDQYJ
# KoZIhvcNAQELBQADggIBAPUPJodwr5miyvXWyfCNZj05gtOII9iCv49UhCe204MH
# 154niU2EjlTRIO5gQ9tXQjzHsJX2vszqoz2OTwbGK1mGf+tzG8rlQCbgPW/M9r1x
# xs19DiBAOdYF0q+UCL9/wlG3K7V7gyHwY9rlnOFpLnUdTsthHvWlM98CnRXZ7WmT
# V7pGRS6AvGW+5xI+3kf/kJwQrfZWsqTU+tb8LryXIbN2g9KR+gZQ0bGAKID+260P
# Z+34fdzZcFt6umi1s0pmF4/n8OdX3Wn+vF7h1YyfE7uVmhX7eSuF1W0+Z0duGwdc
# +1RFDxYRLhHDsLy1bhwzV5Qe/kI0Ro4xUE7bM1eV+jjk5hLbq1guRbfZIsr0WkdJ
# LCjoT4xCPGRo6eZDrBmRqccTgl/8cQo3t51Qezxd96JSgjXktefTCm9r/o35pNfV
# HUvnfWII+NnXrJlJ27WEQRQu9i5gl1NLmv7xiHp0up516eDap8nMLDt7TAp4z5T3
# NmC2gzyKVMtODWgqlBF1JhTqIDfM63kXdlV4cW3iSTgzN9vkbFnHI2LmvM4uVEv9
# XgMqyN0eS3FE0HU+MWJliymm7STheh2ENH+kF3y0rH0/NVjLw78a3Z9UVm1F5VPz
# iIorMaPKPlDRADTsJwjDZ8Zc6Gi/zy4WZbg8Zv87spWrmo2dzJTw7XhQf+xkR6Od
# MIIG9TCCBN2gAwIBAgIQBof4ljXihJpmCqqt2mhAWzANBgkqhkiG9w0BAQsFADB4
# MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24x
# ETAPBgNVBAoMCFNTTCBDb3JwMTQwMgYDVQQDDCtTU0wuY29tIENvZGUgU2lnbmlu
# ZyBJbnRlcm1lZGlhdGUgQ0EgUlNBIFIxMB4XDTI2MTAxMzAyNDgwMloXDTI3MTAx
# MzAyNDgwMlowXDELMAkGA1UEBhMCRlIxGzAZBgNVBAcMEkNoYW50ZWxvdXAtRW4t
# QnJpZTEXMBUGA1UECgwOTmljb2xhcyBGYWJlcnQxFzAVBgNVBAMMDk5pY29sYXMg
# RmFiZXJ0MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsg6CT0mFQ4jB
# V2plIS4uWFSoVqRHd82Y7XYZ4CUTe/7nPBRUlXXdKTZMMHzGgiG8dEXNE8K1Xe90
# Dgl/FnorzAFlrNWSwf6t27UAUw9xkqtRDVVpVgpixpWkAwjcQeLG+YMqBkGbFi35
# TGvS23lPFWvLp4DxARrv15MtZ/NjNTCAAGUlMP2hFVBknEmrnYjEGyeJ93sl20RS
# WXQQLNqlenBdkGCrQW4XjHtRjJAEqAC68hN6nWevopSQ8CLUV7A1RFfHFarDrzrL
# zr0/NjIdSyIfM9hShln31NDrktB6Ups2XIyrtPDNB7fjkW4jpLlOqQGfYKsU9ErE
# 4SnFsh0yOOjitegyh1RvZ9wH++3DF4CmJfTli5l/gEIm8xCCryxI+K5EogW3wu1y
# u8NT2hcvHjdh7OpRyPyFAuLZ8MfX7kPAJpnuCeOuktzY/20eVD8bDglj+4zgqzsn
# cR/t6T/7w1/HxpRyrVoKvEyRx+mphYh5jMdEa0BjxOL0vAJpdGZSfXGmPm/yPqRn
# 5STTkMtzkVQXPbcvgKM0luzFcQKPHk3qdHLEXa8gf93ktYDbLE5Wd2UHrXz3SZxN
# URHPdIXCr1qgrqQve7041kMpzrkDI5DnG2lvRG4108gp02bHG0muKyZlDdCOra7h
# Zonu2nO79sFallruzzNi0o994PoWksUCAwEAAaOCAZUwggGRMAwGA1UdEwEB/wQC
# MAAwHwYDVR0jBBgwFoAUVML+EJUAk81q9efA19myS7iPDOMwegYIKwYBBQUHAQEE
# bjBsMEgGCCsGAQUFBzAChjxodHRwOi8vY2VydC5zc2wuY29tL1NTTGNvbS1TdWJD
# QS1Db2RlU2lnbmluZy1SU0EtNDA5Ni1SMS5jZXIwIAYIKwYBBQUHMAGGFGh0dHA6
# Ly9vY3Nwcy5zc2wuY29tMFEGA1UdIARKMEgwCAYGZ4EMAQQBMDwGDCsGAQQBgqkw
# AQMDATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3JlcG9zaXRv
# cnkwEwYDVR0lBAwwCgYIKwYBBQUHAwMwTQYDVR0fBEYwRDBCoECgPoY8aHR0cDov
# L2NybHMuc3NsLmNvbS9TU0xjb20tU3ViQ0EtQ29kZVNpZ25pbmctUlNBLTQwOTYt
# UjEuY3JsMB0GA1UdDgQWBBRoUx3630l9wA1fUpkKkw2GUIxl/jAOBgNVHQ8BAf8E
# BAMCB4AwDQYJKoZIhvcNAQELBQADggIBAE3oj6cNuwlDToi1BwjgsA5djfRN4DIi
# BYDFZIRYaIRe/rEdnItz3aqQQdFuwXnJA/MQEmi2WAGek8GeZF+R41s8FbcDxBUE
# BoSt98TuBE+0xgxQWDiPhhXaX1Aw9V0GF4ZnlfCSw1eu2KiASaKn7Bo0PLbEAXff
# vftRXof/kyqtw8iID6I4Rhrf5DI3WVZQodakUUmWnbnziWvEZOWbNOfzwP+7BwEt
# 7jQANfbmEfb4p75rfy5dQVwniTDSvdo6xkrbe1Cr93GS7cJta18IG5CoIVUmdK+s
# W38B1sLuy8Mi0MstMXV3004yaRuzLdNaAM0f9WAFerxemiQC0kdmA44/lWV3lVkM
# WJSwkHBkkSg8Eh0hxcGzTH2xmtDhLJHNNLRwkI9u2Ur+HbLCt50oBgSV8BcCs/1i
# qi884iC/yuX+HIv9YgTUNZ8c4lLaq4VuliJpXLu2buHIsTkO0v0+5oCM91zjTVZm
# ULeH38N0Vr83psiJVuBSxQyt5nRVrpYz7FACq5gtXELkyViq/rGOlJKaRrgQ6WWq
# eoIXqKhVFAdP/jAEm2M8ezoVYsUgPyKY2OcuecV9coulAcxUIUydx6qYiOfsGj5u
# Dy7l4aYjvB771wkUMzUQwxzcDYXJtMMTaWRe6n9EemR67jc73u5CJyFg3PfA5WR1
# wZislcdCZuS6MYIYEDCCGAwCAQEwgYwweDELMAkGA1UEBhMCVVMxDjAMBgNVBAgM
# BVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMREwDwYDVQQKDAhTU0wgQ29ycDE0MDIG
# A1UEAwwrU1NMLmNvbSBDb2RlIFNpZ25pbmcgSW50ZXJtZWRpYXRlIENBIFJTQSBS
# MQIQBof4ljXihJpmCqqt2mhAWzANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3
# AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisG
# AQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCcQXxWsa0j
# xw4SPa6CqYsw7Kb9lZMukdma/g+wd5+k2jANBgkqhkiG9w0BAQEFAASCAgAOXQPJ
# 41fNYyPSoeIluIAWG6AzgCMq2/BFOFLQ40vz82tQFUssrAkuAyepsYO1U3LoHdgd
# ZUhnshN5uKML7NJI1YkODWrj5z31OqpNqibC60N76eGRZNr4xkMqkQoJ2Jvjcuj/
# 110UrXthfi7EiQ7LhCnD1p/Z7RkzYMoJoUAegelicOnQKmdFFJwKyFb7Q/T0vPgL
# DCHZQNUGx5W/x9t1VJQPCXjA5KDFFkISNJQ25+e8sP+RhUHqdwEKhFlYCx8Y05nl
# zwBShQiwSGuGHLTER6d+sJx7uvfoP70R4bbb0ic7XxlVSaUsNce2EpWOcsL7Pfk6
# tRzWn1uBShYtlRCusf53f8kO+KO5+3r2axemgu+AijARHdbP6A4Gl5US5y1gMIos
# ysGzem6iSp8vry4Qg589XifpQkjOVTGrt1UUpbowUZfCikUwycFGWqqry+/yHtiX
# rAT2x/kz2vgiOa2LvZ60U1Y+Fr5wBtKcQdKuxaYifcJp7vyCUagXat/fIiiMDnKA
# pp8mzC+QQ3PoBoSoOd1Ea1jWxyZGmt5JArKsIEg5o2kZTuitdekyCsYXYKDykvB3
# ODFsvMutLy9/JC6ZDdigvK6X0BQJ1yJyFQtgSK9a1DTGbgstkKmNyMczAv8h7X5s
# G0bSyP7L0bedYXqhwxjASC9KdRb3eQQKKhOxlKGCFM0wghTJBgorBgEEAYI3AwMB
# MYIUuTCCFLUGCSqGSIb3DQEHAqCCFKYwghSiAgEDMQ0wCwYJYIZIAWUDBAIBMIGJ
# BgsqhkiG9w0BCRABBKB6BHgwdgIBAQYMKwYBBAGCqTABAwYBMDEwDQYJYIZIAWUD
# BAIBBQAEILaaZJcNRm7KnTY4IqABddZm/fOQXNemhex2NTIJn1clAggQbPrrR97u
# vxgPMjAyNjExMTAxMDIzNDlaMAMCAQECECdnN+7rEDzMhS7CvVOy+KOgghGnMIIE
# 8zCCAtugAwIBAgIQVEP4N71a7hgBhHK4XNqO4zANBgkqhkiG9w0BAQsFADByMQsw
# CQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0b24xETAP
# BgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFtcGluZyBJ
# c3N1aW5nIFJTQSBDQSAyMB4XDTI2MDYxMjE3MjI0NloXDTM2MDYwOTE3MjI0NVow
# bjELMAkGA1UEBhMCVVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9u
# MREwDwYDVQQKDAhTU0wgQ29ycDEqMCgGA1UEAwwhU1NMLmNvbSBUaW1lc3RhbXBp
# bmcgVW5pdCAyMDI2IEUxMFkwEwYHKoZIzj0CAQYIKoZIzj0DAQcDQgAE6CWLuGYT
# LkiSOPMLNskOBthXsIhRSA46feOQwvSLcDaNh6vxUzBKEx1KdFBRFs7OHrnH/K0V
# A6MbDQIFpUg0PaOCAVIwggFOMB8GA1UdIwQYMBaAFHCaZcFDgKP86bFkiW7xMJwJ
# Icl8ME0GCCsGAQUFBwEBBEEwPzA9BggrBgEFBQcwAoYxaHR0cDovL2NlcnQuc3Ns
# LmNvbS9TU0wuY29tLXRpbWVTdGFtcGluZy1JLVIyLmNlcjBRBgNVHSAESjBIMDwG
# DCsGAQQBgqkwAQMGATAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29t
# L3JlcG9zaXRvcnkwCAYGZ4EMAQQCMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMEIG
# A1UdHwQ7MDkwN6A1oDOGMWh0dHA6Ly9jcmxzLnNzbC5jb20vU1NMLmNvbS10aW1l
# U3RhbXBpbmctSS1SMi5jcmwwHQYDVR0OBBYEFCesA1a6KGq6FzuDPUB/eagk7MM8
# MA4GA1UdDwEB/wQEAwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAT3ZuzvwDnsIZTXBn
# +hjmuBXgzB12fbYxT9uYe6ZvhvpbZWKcyDQgacETkPw8k6yeqnYCYyFZ063lMPEq
# AhAD2nMG614Zh7PtHW9DWqfsgivxbPYBeN+t7UxgkYT3expbWMdPEqBtWVyHOvLc
# KvMxgJWfJjANwVkwVDfky/j7cf28+iimN8PKc7S/u8/LenRRnV2W2YfncRnkz9Cf
# s5nRDzaZguiHrrp2GxwgnNQSH4GBjHOl3PzLsnVbGoZKi4ZEjgXups4mGGOK7+GB
# 2HAlQn+OS944cHqlkH0gElt2WezoCNRpELph9qaxuBnjSEqXWWfgE6yGBwO0uxKN
# qpWVTHmAwr6SvBPofxubXIuEn9ivR4CLOuAxUSrJ41ZBg/qqG0oJjF6GkRKbxq7c
# yWSkEZYJhmKCg2Hg5fXHagOxiwbmn5JPZ1vhWCrbk3v/dTx2gvTK5+erC5H9/rTc
# ljUI776yd2rfhjfkpvyAsbtXs4D7nrnw8i9wor5+fdtRSJBVxqzLmykFftoL0Up2
# Q++tyKCUdribdcnrB4HKS3s+t3AjdcbA06DCL8GifQi9fL0yRBRNlA4o6BP+Sr75
# 5MaCNjURYZDvDyxJHBBT7cqjZIwsiR8s+6860noFbBBtdge2Iw7CPC/1Q31I27hV
# olS7B0yao3v8X2MtWy3msZlXTcEwggXdMIIDxaADAgECAgh7LJvTFoAymTANBgkq
# hkiG9w0BAQsFADB8MQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNV
# BAcMB0hvdXN0b24xGDAWBgNVBAoMD1NTTCBDb3Jwb3JhdGlvbjExMC8GA1UEAwwo
# U1NMLmNvbSBSb290IENlcnRpZmljYXRpb24gQXV0aG9yaXR5IFJTQTAeFw0xNjAy
# MTIxNzM5MzlaFw00MTAyMTIxNzM5MzlaMHwxCzAJBgNVBAYTAlVTMQ4wDAYDVQQI
# DAVUZXhhczEQMA4GA1UEBwwHSG91c3RvbjEYMBYGA1UECgwPU1NMIENvcnBvcmF0
# aW9uMTEwLwYDVQQDDChTU0wuY29tIFJvb3QgQ2VydGlmaWNhdGlvbiBBdXRob3Jp
# dHkgUlNBMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA+Q/doyt9y9Aq
# /uxnhabnLhu6d+Hj9a+k7PpKXZHEV0drGHdrdvL9k+Q9D8IWngtmw1aUnheDhc5W
# 7/IW/QBi9SIJVOhlF05BueBPRpeqG8i4bmJeabFf2yoCfvxsyvNB2O3Q6Pw/YUjt
# sAMUHRAOSxngu07shmX/NvNeZwILnYZVYf16OO3+4hkAt2+hUGJ1dDyg+sglkrRu
# eiLH+B6h47LdkTGrKx0E/6VKBDfphaQzK/3i1lU0fBmkSmjHsqjTt8qhk4jrwZe8
# jPkd2SKEJHTHBD1qqSmTzOu4W+H+XyWqNFjIwSNUnRuYEcM4nH49hmylD0CGfAL0
# XAJPKMuucZ8POsgz/hElNer8usVgPdl8GNWyqdN1eANyIso6wx/vLOUuqfqeLLZR
# Rv2vA9bqYGjqhRY2a4XpHsCz3cQk3IAqgUFtlD7I4MmBQQCeXr9/xQiYohgsQkCz
# +W84J0tOgPQ9gUfgiHzqHM61dVxRLhwrfxpyKOcAtdF0xtfkn60Hk7ZTNTX8N+TD
# 9l0WviFz3pIK+KBjaryWkmo++LxlVZve9Q2JJgT8JRqmJWnLwm3KfOJZX5es6+8u
# yLzXG1k8K8zyGciTaydjGc/86Sb4ynGbf5P+NGeETpnr/LN4CTNwumamdu0bc+sa
# pQ3EIhMglFYKTixsTrH9z5wJuqIz7YcCAwEAAaNjMGEwHQYDVR0OBBYEFN0ECQei
# 9Xp9UlMSkpXuOIAlDaZZMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU3QQJ
# B6L1en1SUxKSle44gCUNplkwDgYDVR0PAQH/BAQDAgGGMA0GCSqGSIb3DQEBCwUA
# A4ICAQAgGBGUKfsmnRweHnBh8ZVyk3EkrWiTWI4yrxuzcAP8JSt0hZA9eGr0uYul
# lzu1GJG7Hqf5QFuR+VWZrx4R0Fwdp2bjsZQHDDI5puobsHnYHZxwROOK3cT5lR+K
# OEM/AYWlR6c9RrK85SJo93uc2Cw+CiHILTOsv8WBmTF0wXVxxb6x8CNF9J1r/Blj
# naO8BMYYCyW7U4kPs4BQ3kXuRH+rlHhkmNP2KN2H2HBldPsOuRPrpw9hqTKWzN67
# 7WNMGLupQPegVG4giHF1GOp6tDRy4CMnd1y2kOqGJUCr7zMPy5+CvqIg+/a1LRrm
# woWxdA/7yGUCpFIBR91JIsG/2OtrrH7e7GMzFbcjCI/GD41BWt2OxbmP5UU/eNu6
# 0htAsf5xTT/ggaK6XrTsFeCT3QgffuFVmQsh3pOeCvvmo0m9NjD+53eyoHWXtS2B
# iBdlIPfakACfyVLMMso1fPU9D9gr1/UmbMkGNJYW6nBZGjJ5eQu2iH8PUkg9v2zY
# okQu0U63cljTiROV/kSr+NeLG26cvCygW9VqAK9fN+HV+hALmJyG5yaPzvDsbopX
# C4DjTrLAoGNhkLpVaDd0araS25+hhiK2ZScO7LafQmDkZ8K12kELxNOLYRu8+h+R
# K9dEB166KazZxenvU0ha64DxKFghzbAGVfsnP1OQcKkEHlcnuTCCBsswggSzoAMC
# AQICEFKdu8K4iYf1g4DgsOuiBjYwDQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMC
# VVMxDjAMBgNVBAgMBVRleGFzMRAwDgYDVQQHDAdIb3VzdG9uMRgwFgYDVQQKDA9T
# U0wgQ29ycG9yYXRpb24xMTAvBgNVBAMMKFNTTC5jb20gUm9vdCBDZXJ0aWZpY2F0
# aW9uIEF1dGhvcml0eSBSU0EwHhcNMjYwNjEyMTY0MjE1WhcNNDEwMjEyMTczOTM5
# WjByMQswCQYDVQQGEwJVUzEOMAwGA1UECAwFVGV4YXMxEDAOBgNVBAcMB0hvdXN0
# b24xETAPBgNVBAoMCFNTTCBDb3JwMS4wLAYDVQQDDCVTU0wuY29tIFRpbWVzdGFt
# cGluZyBJc3N1aW5nIFJTQSBDQSAyMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAny+r5nNNkLwxRps2B3laUGNU9y7vWiJ5013yf0iXz65bQdwa9bjj8l/S
# keXqHGqfSh2Foae+80E8mDEf93oIwjsy2CX2cXmKxqtF8HJGYB4f6yRhKED50wkb
# OysL02IXf+/EZna44jffhhVz4OU/k52xQ/PnIg/uFhc2Wk+Sc/CHqAr+lYVkPz7L
# y5RqB9f1RKGr9JC37W1LgVdHC50dNIdsa3QXPhltYZubR3npE4Bbfi0joUW5n5O0
# ayweN45R/YVLxA0YGlp1OHVfVnsqXq8bhnw1Lc4kjJ6XOgXlOlw/a8ze57TvbduZ
# l3ehewTmvO0Aky/nqOWDS9IPPxdxdG6vprBbSuHlxNANBAbHTiFgkfZmf0dZYlT/
# AnZIoFxAq0HILPolf4OgWgET76SiWSlm2XYRIXwVjpeqG0plvXwRIQBzSiEJj7Qe
# QFnCvHx4gQMINvN3+yG/1rxhEbajGve0tTP3lX4PPOHVAE0QJp4ZzPI1YYJkhVRr
# iVVxVRU4Fsg4BPwP6I+YLmHb1tZ4RLwOFNLh/gJYDgPlChCoxOogFxXgwPZZrDyE
# RYB0aeOexW0OKacN4RRu9cqF2KKu/Ab5AL7eYDLVGUPrEEyYbIMApVKqpebc0/jD
# QniC+3rkauRAjgVc8m1ip/CQgG1Dw4Z9ulVRAY9VR6Za26LsTbUCAwEAAaOCAVEw
# ggFNMBIGA1UdEwEB/wQIMAYBAf8CAQAwHwYDVR0jBBgwFoAU3QQJB6L1en1SUxKS
# le44gCUNplkwTQYIKwYBBQUHAQEEQTA/MD0GCCsGAQUFBzAChjFodHRwOi8vY2Vy
# dC5zc2wuY29tL1NTTGNvbS1Sb290Q0EtUlNBLTQwOTYtUjEuY2VyMD8GA1UdIAQ4
# MDYwNAYEVR0gADAsMCoGCCsGAQUFBwIBFh5odHRwczovL3d3dy5zc2wuY29tL3Jl
# cG9zaXRvcnkwEwYDVR0lBAwwCgYIKwYBBQUHAwgwQgYDVR0fBDswOTA3oDWgM4Yx
# aHR0cDovL2NybHMuc3NsLmNvbS9TU0xjb20tUm9vdENBLVJTQS00MDk2LVIxLmNy
# bDAdBgNVHQ4EFgQUcJplwUOAo/zpsWSJbvEwnAkhyXwwDgYDVR0PAQH/BAQDAgGG
# MA0GCSqGSIb3DQEBCwUAA4ICAQBlFFvNMemA6iBiBYo56IKQjWhWuQaYKuo4qpf0
# KJNERAt1gz+HsA2SL3xA/V+EA8SRPNP9SvX0is+5mDsppYvpKyYGqYP36YFW0mbF
# 1F0ZtXxumCUUO/ZV5mBNKDf5Q4PY7Qb0eWAepUc0rNJr/Mvn5ZiXfJPxDb2xmcg8
# aRD652/nr+Ork2UBP0E43iJdI4LF7Vs29R3SGRLYZOxM15DI4QebsKRl0KDlf4Om
# YGCKc7umTxibkZtN1C1KWgnf9jadU/JwnySaYEBgx6/v+FVC6SGONziY7TcvI9MV
# +qYxHUPlOm6AaQP2dS4bjWXnm86ldCbUkxtOACumSfT97077JNiVa7PZ1IBkd/2P
# E1AC8gs7U+/Yk1BN9XRPIi15k/uA1xF9oeee5KGjyiGKkqIT7W8QkgUT5zeCinFK
# kcKafZMcROSVYpXtmtQqGGR2g4uPYqJ0K1hms6gJXrHt/34p8M9PrV8oNWe0Zl1E
# rDMtXndUg3qhJL4pA3ZUlo7w5lM5ml+A7ijr/3Gk2v4QRoDeE2ShmTXdHgtqo8L/
# fT119mYcETycE2NYGRXsBl3N0rzDsHRr8P3TrprqhRrX04iS0P7eQ+lDxgS3yhGn
# 3OzWGhDt9Gf2x1qe0/qdhFAVOkXxzMZz4Uat6xsAfzX0+/m8GEKW5Ojw/HY+kLvV
# wzSPOzGCAlUwggJRAgEBMIGGMHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhh
# czEQMA4GA1UEBwwHSG91c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMM
# JVNTTC5jb20gVGltZXN0YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4Y
# AYRyuFzajuMwCwYJYIZIAWUDBAIBoIIBYDAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
# AQkQAQQwHAYJKoZIhvcNAQkFMQ8XDTI2MTExMDEwMjM0OVowKAYJKoZIhvcNAQk0
# MRswGTALBglghkgBZQMEAgGhCgYIKoZIzj0EAwIwLwYJKoZIhvcNAQkEMSIEIEJF
# dYqyAeMOZbojctPRgbRo/O0+nrouIozwVpsg1G1KMIHIBgsqhkiG9w0BCRACLzGB
# uDCBtTCBsjCBrwQgV0/qF6r6XnxqK5sIyZixc15FCzhzykTmokygkw5j87AwgYow
# dqR0MHIxCzAJBgNVBAYTAlVTMQ4wDAYDVQQIDAVUZXhhczEQMA4GA1UEBwwHSG91
# c3RvbjERMA8GA1UECgwIU1NMIENvcnAxLjAsBgNVBAMMJVNTTC5jb20gVGltZXN0
# YW1waW5nIElzc3VpbmcgUlNBIENBIDICEFRD+De9Wu4YAYRyuFzajuMwCgYIKoZI
# zj0EAwIERjBEAiALjIunJi/MTG8oyWizhQVOEpuKSW3K8rzgCVTTjSIXMwIgG4oq
# 6EGo3PezMXTWcOXhxDhKes8j51iOXhkx34KXVgU=
# SIG # End signature block
