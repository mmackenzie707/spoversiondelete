<#
.SYNOPSIS
    SharePoint Online Version Trimmer - With GUI Device Code Display
.DESCRIPTION
    Shows the device login code and URL directly in the GUI, no need to check console.
#>

#requires -Version 5.1
#requires -Modules PnP.PowerShell

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Web  # For UrlEncode

# XAML GUI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SPO Version Trimmer - Device Login" Height="900" Width="950"
        WindowStartupLocation="CenterScreen" Background="#FFF5F5F5">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="20,8"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Background" Value="#FF0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#FF005A9E"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Background" Value="#FFCCCCCC"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="BorderBrush" Value="#FFCCCCCC"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style x:Key="CodeDisplayStyle" TargetType="TextBox">
            <Setter Property="FontSize" Value="32"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="FontFamily" Value="Consolas"/>
            <Setter Property="Foreground" Value="#FF107C10"/>
            <Setter Property="Background" Value="#FFF0F0F0"/>
            <Setter Property="HorizontalContentAlignment" Value="Center"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Height" Value="60"/>
            <Setter Property="IsReadOnly" Value="True"/>
        </Style>
    </Window.Resources>
    
    <Grid Margin="15">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Header -->
        <Border Grid.Row="0" Background="#FF0078D4" CornerRadius="4" Margin="0,0,0,15">
            <StackPanel Margin="20,15">
                <TextBlock Text="SharePoint Online Version Trimmer" Foreground="White" FontSize="22" FontWeight="Bold"/>
                <TextBlock Text="Device Login Edition - Code displayed in GUI" Foreground="#FFE0E0E0" FontSize="12" Margin="0,5,0,0"/>
            </StackPanel>
        </Border>
        
        <!-- Connection Panel -->
        <GroupBox Grid.Row="1" Header="Step 1: Connection Settings" Margin="0,0,0,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="130"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                
                <Label Grid.Row="0" Grid.Column="0" Content="Site URL:" VerticalAlignment="Center"/>
                <TextBox x:Name="txtSiteUrl" Grid.Row="0" Grid.Column="1"/>
                
                <Label Grid.Row="1" Grid.Column="0" Content="Client ID:" VerticalAlignment="Center"/>
                <TextBox x:Name="txtClientId" Grid.Row="1" Grid.Column="1"/>
                
                <Label Grid.Row="2" Grid.Column="0" Content="Tenant ID:" VerticalAlignment="Center"/>
                <TextBox x:Name="txtTenantId" Grid.Row="2" Grid.Column="1"/>
                
                <Button x:Name="btnConnect" Grid.Row="3" Grid.Column="1" Content="Connect to SharePoint" 
                        HorizontalAlignment="Left" Width="200" Margin="5,10,5,5"/>
            </Grid>
        </GroupBox>
        
        <!-- Device Code Display Panel (Initially Hidden) -->
        <Border x:Name="brdDeviceCode" Grid.Row="2" Background="#FFFFF3CD" BorderBrush="#FFE0A800" 
                BorderThickness="1" CornerRadius="4" Margin="0,0,0,10" Visibility="Collapsed" Padding="15">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <TextBlock Grid.Row="0" Text="Authentication Required" FontSize="18" FontWeight="Bold" 
                          Foreground="#FF856404" HorizontalAlignment="Center"/>
                
                <TextBlock Grid.Row="1" Text="Enter this code at microsoft.com/devicelogin:" 
                          HorizontalAlignment="Center" Margin="0,10,0,5" FontSize="14"/>
                
                <TextBox x:Name="txtDeviceCode" Grid.Row="2" Style="{StaticResource CodeDisplayStyle}" 
                        Text="XXXX-XXXX" Margin="50,5"/>
                
                <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Center" Margin="0,10,0,0">
                    <Button x:Name="btnOpenBrowser" Content="Open Browser" Width="150" Margin="5"/>
                    <Button x:Name="btnCopyCode" Content="Copy Code" Width="150" Margin="5" Background="#FF6C757D"/>
                </StackPanel>
                
                <TextBlock x:Name="txtAuthStatus" Grid.Row="4" Text="Waiting for authentication..." 
                          HorizontalAlignment="Center" Margin="0,10,0,0" FontStyle="Italic" Foreground="#FF856404"/>
            </Grid>
        </Border>
        
        <!-- Configuration -->
        <GroupBox Grid.Row="3" Header="Step 2: Trim Configuration" IsEnabled="False" x:Name="grpConfig" Margin="0,0,0,10">
            <Grid>
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>
                
                <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,5">
                    <Label Content="Library:"/>
                    <ComboBox x:Name="cmbLibraries" Width="350" Height="32" Margin="5" VerticalContentAlignment="Center"/>
                    <Button x:Name="btnRefresh" Content="Refresh" Width="100" Margin="5"/>
                </StackPanel>
                
                <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,5">
                    <Label Content="Keep Versions:"/>
                    <TextBox x:Name="txtKeepVersions" Text="10" Width="60" TextAlignment="Center"/>
                    
                    <Label Content="Batch Size:" Margin="20,0,0,0"/>
                    <TextBox x:Name="txtBatchSize" Text="100" Width="60" TextAlignment="Center"/>
                    
                    <CheckBox x:Name="chkWhatIf" Content="SIMULATION ONLY (What-If)" IsChecked="True" 
                              VerticalAlignment="Center" Margin="30,0,0,0" Foreground="#FFCC0000" FontWeight="Bold"/>
                </StackPanel>
            </Grid>
        </GroupBox>
        
        <!-- Action -->
        <Button x:Name="btnTrim" Grid.Row="4" Content="START VERSION TRIM" VerticalAlignment="Top"
                Width="280" Height="50" FontSize="16" Background="#FF107C10" 
                IsEnabled="False" HorizontalAlignment="Center" Margin="0,10"/>
        
        <!-- Log -->
        <GroupBox Grid.Row="5" Header="Activity Log" Height="200">
            <TextBox x:Name="txtLog" IsReadOnly="True" FontFamily="Consolas" FontSize="11"
                    VerticalScrollBarVisibility="Auto" TextWrapping="Wrap" 
                    Background="#FF1E1E1E" Foreground="#FF00FF00"/>
        </GroupBox>
    </Grid>
</Window>
"@

$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Get Controls
$txtSiteUrl = $window.FindName("txtSiteUrl")
$txtClientId = $window.FindName("txtClientId")
$txtTenantId = $window.FindName("txtTenantId")
$txtKeepVersions = $window.FindName("txtKeepVersions")
$txtBatchSize = $window.FindName("txtBatchSize")
$cmbLibraries = $window.FindName("cmbLibraries")
$chkWhatIf = $window.FindName("chkWhatIf")
$btnConnect = $window.FindName("btnConnect")
$btnRefresh = $window.FindName("btnRefresh")
$btnTrim = $window.FindName("btnTrim")
$grpConfig = $window.FindName("grpConfig")
$txtLog = $window.FindName("txtLog")
$brdDeviceCode = $window.FindName("brdDeviceCode")
$txtDeviceCode = $window.FindName("txtDeviceCode")
$btnOpenBrowser = $window.FindName("btnOpenBrowser")
$btnCopyCode = $window.FindName("btnCopyCode")
$txtAuthStatus = $window.FindName("txtAuthStatus")

# Globals
$script:Connection = $null
$script:AuthContext = @{}

# Logging
function Write-Log {
    param([string]$Message, [switch]$IsError)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $window.Dispatcher.Invoke([action]{
        if ($IsError) {
            $txtLog.AppendText("[$timestamp] ERROR: $Message`r`n")
        } else {
            $txtLog.AppendText("[$timestamp] $Message`r`n")
        }
        $txtLog.ScrollToEnd()
    })
}

# Device Login Flow (Manual OAuth2 to capture code)
function Start-DeviceLogin {
    param($SiteUrl, $ClientId, $TenantId)
    
    $uri = [System.Uri]$SiteUrl
    $resource = "$($uri.Scheme)://$($uri.Host)"
    
    $body = @{
        client_id = $ClientId
        scope = "$resource/.default offline_access"
    }
    
    try {
        Write-Log "Requesting device login code..."
        $response = Invoke-RestMethod -Method POST `
            -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/devicecode" `
            -Body $body `
            -ErrorAction Stop
        
        # Store context - Note: expires_in is usually 900 seconds (15 mins)
        $script:AuthContext = @{
            DeviceCode = $response.device_code
            UserCode = $response.user_code
            VerificationUri = "https://microsoft.com/devicelogin"
            Interval = $response.interval  # Usually 5 seconds
            ExpiresAt = (Get-Date).AddSeconds($response.expires_in)
            ClientId = $ClientId
            TenantId = $TenantId
            Resource = $resource
            SiteUrl = $SiteUrl
        }
        
        # Show in GUI
        $window.Dispatcher.Invoke([action]{
            $txtDeviceCode.Text = $response.user_code
            $brdDeviceCode.Visibility = "Visible"
            $txtAuthStatus.Text = "Enter this code in your browser (you have 15 minutes)"
            $txtAuthStatus.Foreground = "#FF856404"
            $btnConnect.IsEnabled = $false
        })
        
        Write-Log "DEVICE CODE: $($response.user_code)"
        Write-Log "Go to: https://microsoft.com/devicelogin"
        Write-Log "You have 15 minutes to complete authentication"
        
        # Open browser
        try {
            Start-Process "https://microsoft.com/devicelogin"
        } catch {
            Write-Log "Please manually open: https://microsoft.com/devicelogin" -IsWarning
        }
        
        # Start polling
        Start-TokenPolling
        
    } catch {
        Write-Log "Failed to start device login: $($_.Exception.Message)" -IsError
        $window.Dispatcher.Invoke([action]{
            $btnConnect.IsEnabled = $true
        })
    }
}

function Start-TokenPolling {
    # Stop any existing timer
    if ($script:PollTimer) {
        $script:PollTimer.Stop()
        $script:PollTimer = $null
    }
    
    $script:PollTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:PollTimer.Interval = [TimeSpan]::FromSeconds($script:AuthContext.Interval)
    
    $script:PollTimer.Add_Tick({
        # Check if expired
        if ((Get-Date) -gt $script:AuthContext.ExpiresAt) {
            $script:PollTimer.Stop()
            $script:PollTimer = $null
            Write-Log "Authentication window expired (15 minutes elapsed)" -IsError
            $window.Dispatcher.Invoke([action]{
                $brdDeviceCode.Visibility = "Collapsed"
                $btnConnect.IsEnabled = $true
                $btnConnect.Content = "Connect to SharePoint"
                [System.Windows.MessageBox]::Show("The 15-minute authentication window expired. Please click Connect to try again.", "Timeout", "OK", "Warning")
            })
            return
        }
        
        try {
            $tokenBody = @{
                grant_type = "urn:ietf:params:oauth:grant-type:device_code"
                client_id = $script:AuthContext.ClientId
                device_code = $script:AuthContext.DeviceCode
            }
            
            $response = Invoke-RestMethod -Method POST `
                -Uri "https://login.microsoftonline.com/$($script:AuthContext.TenantId)/oauth2/v2.0/token" `
                -Body $tokenBody `
                -ErrorAction Stop
            
            # SUCCESS!
            Write-Log "Authentication successful!"
            $script:PollTimer.Stop()
            $script:PollTimer = $null
            
            $window.Dispatcher.Invoke([action]{
                $txtAuthStatus.Text = "Authenticated! Connecting..."
                $txtAuthStatus.Foreground = "Green"
            })
            
            Connect-WithToken -AccessToken $response.access_token
            
        } catch {
            $errorBody = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
            $errorCode = $errorBody.error
            
            if ($errorCode -eq "authorization_pending") {
                # Normal - user hasn't done it yet
                $window.Dispatcher.Invoke([action]{
                    $timeLeft = [math]::Round(($script:AuthContext.ExpiresAt - (Get-Date)).TotalMinutes)
                    $txtAuthStatus.Text = "Waiting... ($timeLeft mins left) - Enter code in browser if you haven't already"
                })
            } 
            elseif ($errorCode -eq "slow_down") {
                Write-Log "Server asked to slow down, waiting longer..."
                $script:PollTimer.Interval = [TimeSpan]::FromSeconds(10)
            }
            elseif ($errorCode -eq "expired_token") {
                $script:PollTimer.Stop()
                $script:PollTimer = $null
                Write-Log "Code expired" -IsError
                $window.Dispatcher.Invoke([action]{
                    $brdDeviceCode.Visibility = "Collapsed"
                    $btnConnect.IsEnabled = $true
                    [System.Windows.MessageBox]::Show("The code expired. Please click Connect to get a new code.", "Expired", "OK", "Warning")
                })
            }
            elseif ($errorCode -eq "invalid_grant") {
                $script:PollTimer.Stop()
                $script:PollTimer = $null
                Write-Log "Invalid code or already used" -IsError
                $window.Dispatcher.Invoke([action]{
                    $brdDeviceCode.Visibility = "Collapsed"
                    $btnConnect.IsEnabled = $true
                    [System.Windows.MessageBox]::Show("Authentication failed. The code may have been entered incorrectly.", "Failed", "OK", "Error")
                })
            }
            else {
                # Unknown error, log but keep trying
                Write-Log "Polling error: $errorCode" -IsError
            }
        }
    })
    
    $script:PollTimer.Start()
}

function Get-DeviceToken {
    $maxAttempts = [math]::Floor($script:AuthContext.ExpiresIn / $script:AuthContext.Interval) + 10
    $script:PollAttempts = 0
    $script:AuthComplete = $false
    
    Write-Log "Starting token polling (max $maxAttempts attempts, interval $($script:AuthContext.Interval)s)"
    
    if ($script:TokenTimer) {
        $script:TokenTimer.Stop()
        $script:TokenTimer = $null
    }
    
    $script:TokenTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:TokenTimer.Interval = [TimeSpan]::FromSeconds($script:AuthContext.Interval)
    
    $script:TokenTimer.Add_Tick({
        if ($script:AuthComplete) { return } # Prevent re-entry
        
        $script:PollAttempts++
        $window.Dispatcher.Invoke([action]{
            $txtAuthStatus.Text = "Checking authentication status... (Attempt $script:PollAttempts)"
        })
        
        try {
            $tokenBody = @{
                grant_type = "urn:ietf:params:oauth:grant-type:device_code"
                client_id = $script:AuthContext.ClientId
                device_code = $script:AuthContext.DeviceCode
            }
            
            $response = Invoke-RestMethod -Method POST `
                -Uri "https://login.microsoftonline.com/$($script:AuthContext.TenantId)/oauth2/v2.0/token" `
                -Body $tokenBody `
                -ErrorAction Stop
            
            # Success! We got a token
            Write-Log "Success! Token received."
            $script:AuthComplete = $true
            
            if ($script:TokenTimer) {
                $script:TokenTimer.Stop()
                $script:TokenTimer = $null
            }
            
            $window.Dispatcher.Invoke([action]{
                $txtAuthStatus.Text = "Authenticated! Connecting to SharePoint..."
                $txtAuthStatus.Foreground = "Green"
            })
            
            # Use the access token
            Connect-WithToken -AccessToken $response.access_token
            
        } catch {
            $errorResponse = $_.Exception.Response
            $errorMessage = $_.Exception.Message
            
            # Try to parse the error from the response body
            $errorDetail = "unknown_error"
            if ($errorResponse) {
                try {
                    $reader = New-Object System.IO.StreamReader($errorResponse.GetResponseStream())
                    $errorBody = $reader.ReadToEnd()
                    $reader.Close()
                    $errorJson = $errorBody | ConvertFrom-Json
                    $errorDetail = $errorJson.error
                    Write-Log "OAuth error detail: $errorDetail" -IsError
                } catch {
                    $errorDetail = $errorMessage
                }
            }
            
            # Handle specific OAuth errors
            switch -Wildcard ($errorDetail) {
                "*authorization_pending*" {
                    # Normal - user hasn't completed auth yet
                    Write-Log "Still waiting for you to enter code in browser..."
                    $window.Dispatcher.Invoke([action]{
                        $txtAuthStatus.Text = "Enter code in browser if not done... (Check #$script:PollAttempts)"
                    })
                }
                "*slow_down*" {
                    # Increase interval
                    Write-Log "Server asked to slow down, increasing wait time"
                    $script:TokenTimer.Interval = [TimeSpan]::FromSeconds($script:AuthContext.Interval + 5)
                }
                "*expired_token*" {
                    Write-Log "Device code expired. Please try again." -IsError
                    $script:AuthComplete = $true
                    if ($script:TokenTimer) { $script:TokenTimer.Stop(); $script:TokenTimer = $null }
                    $window.Dispatcher.Invoke([action]{
                        $brdDeviceCode.Visibility = "Collapsed"
                        $btnConnect.IsEnabled = $true
                        [System.Windows.MessageBox]::Show("The login code expired. Please click Connect to try again.", "Code Expired", "OK", "Warning")
                    })
                }
                "*invalid_grant*" {
                    Write-Log "Invalid grant. Code may be wrong or already used." -IsError
                    $script:AuthComplete = $true
                    if ($script:TokenTimer) { $script:TokenTimer.Stop(); $script:TokenTimer = $null }
                    $window.Dispatcher.Invoke([action]{
                        $brdDeviceCode.Visibility = "Collapsed"
                        $btnConnect.IsEnabled = $true
                    })
                }
                default {
                    # Unknown error, log but keep trying unless max attempts reached
                    Write-Log "Error checking status: $errorDetail" -IsError
                    if ($script:PollAttempts -ge $maxAttempts) {
                        Write-Log "Max attempts reached, giving up." -IsError
                        $script:AuthComplete = $true
                        if ($script:TokenTimer) { $script:TokenTimer.Stop(); $script:TokenTimer = $null }
                        $window.Dispatcher.Invoke([action]{
                            $brdDeviceCode.Visibility = "Collapsed"
                            $btnConnect.IsEnabled = $true
                            [System.Windows.MessageBox]::Show("Authentication timed out. Please try again.", "Timeout", "OK", "Warning")
                        })
                    }
                }
            }
        }
    })
    
    $script:TokenTimer.Start()
}

function Connect-WithToken {
    param($AccessToken)
    
    try {
        Write-Log "Connecting to SharePoint..."
        $script:Connection = Connect-PnPOnline -Url $script:AuthContext.SiteUrl `
            -AccessToken $AccessToken `
            -ReturnConnection `
            -ErrorAction Stop
        
        $web = Get-PnPWeb -Connection $script:Connection
        Write-Log "Connected to: $($web.Title)"
        
        $window.Dispatcher.Invoke([action]{
            $brdDeviceCode.Visibility = "Collapsed"
            $grpConfig.IsEnabled = $true
            $btnConnect.IsEnabled = $false
            $btnConnect.Content = "CONNECTED ✓"
            $btnConnect.Background = "#FF107C10"
            
            $txtSiteUrl.IsReadOnly = $true
            $txtClientId.IsReadOnly = $true
            $txtTenantId.IsReadOnly = $true
            
            Load-Libraries
        })
        
    } catch {
        Write-Log "Connection failed: $($_.Exception.Message)" -IsError
        $window.Dispatcher.Invoke([action]{
            $brdDeviceCode.Visibility = "Collapsed"
            $btnConnect.IsEnabled = $true
            $btnConnect.Content = "Connect to SharePoint"
            [System.Windows.MessageBox]::Show("Failed to connect to SharePoint: $($_.Exception.Message)", "Error", "OK", "Error")
        })
    }
}

function Load-Libraries {
    try {
        Write-Log "Loading document libraries..."
        $libs = Get-PnPList -Connection $script:Connection | 
                Where-Object { $_.BaseTemplate -eq 101 -and -not $_.Hidden } |
                Select-Object Title
        
        $cmbLibraries.Items.Clear()
        foreach ($lib in $libs) {
            $cmbLibraries.Items.Add($lib.Title)
        }
        
        if ($cmbLibraries.Items.Count -gt 0) {
            $cmbLibraries.SelectedIndex = 0
            $btnTrim.IsEnabled = $true
            Write-Log "Found $($libs.Count) libraries. Ready!"
        } else {
            Write-Log "Warning: No document libraries found" -IsWarning
        }
    } catch {
        Write-Log "Error loading libraries: $($_.Exception.Message)" -IsError
    }
}

# Event Handlers
$btnConnect.Add_Click({
    $btnConnect.IsEnabled = $false
    Start-DeviceLogin -SiteUrl $txtSiteUrl.Text -ClientId $txtClientId.Text -TenantId $txtTenantId.Text
})

$btnOpenBrowser.Add_Click({
    try {
        Start-Process "https://microsoft.com/devicelogin"
        Write-Log "Opened browser to microsoft.com/devicelogin"
    } catch {
        [System.Windows.Clipboard]::SetText("https://microsoft.com/devicelogin")
        [System.Windows.MessageBox]::Show("Could not open browser. URL copied to clipboard: https://microsoft.com/devicelogin", "Browser", "OK", "Information")
    }
})

$btnCopyCode.Add_Click({
    [System.Windows.Clipboard]::SetText($txtDeviceCode.Text)
    [System.Windows.MessageBox]::Show("Code copied to clipboard!", "Copied", "OK", "Information")
})

$btnRefresh.Add_Click({
    Load-Libraries
})

$btnTrim.Add_Click({
    $libraryName = $cmbLibraries.SelectedItem
    if (-not $libraryName) {
        [System.Windows.MessageBox]::Show("Please select a document library first.", "No Library Selected", "OK", "Warning")
        return
    }
    
    $keep = [int]$txtKeepVersions.Text
    $batch = [int]$txtBatchSize.Text
    $whatIf = $chkWhatIf.IsChecked
    
    if ($keep -lt 1) {
        [System.Windows.MessageBox]::Show("Keep Versions must be at least 1", "Invalid Input", "OK", "Warning")
        return
    }
    
    # Confirmation
    $msgIcon = if ($whatIf) { "Information" } else { "Warning" }
    $confirmText = if ($whatIf) { 
        "SIMULATION MODE - No changes will be made.`n`nLibrary: $libraryName`nKeep Versions: $keep`nBatch Size: $batch" 
    } else { 
        "WARNING: This will PERMANENTLY DELETE versions!`n`nLibrary: $libraryName`nKeeping last $keep versions only`nBatch Size: $batch" 
    }
    
    $result = [System.Windows.MessageBox]::Show($confirmText, "Confirm Operation", "OKCancel", $msgIcon)
    if ($result -ne "OK") { return }
    
    # Disable UI during operation
    $btnTrim.IsEnabled = $false
    $btnTrim.Content = "RUNNING..."
    $grpConfig.IsEnabled = $false
    
    Write-Log "======================================="
    Write-Log "STARTING VERSION TRIM"
    Write-Log "Library: $libraryName"
    Write-Log "Keep Versions: $keep | What-If Mode: $whatIf"
    Write-Log "======================================="
    
    # Run trim in background job
    $trimJob = Start-Job -ScriptBlock {
        param($Connection, $LibraryName, $KeepCount, $PageSize, $Simulate)
        
        $results = @{
            Log = @()
            Processed = 0
            Skipped = 0
            Failed = 0
            Deleted = 0
        }
        
        function Add-JobLog($Message, $IsError = $false) {
            $ts = Get-Date -Format "HH:mm:ss"
            $prefix = if ($IsError) { "ERROR" } else { "INFO" }
            $results.Log += "[$ts] [$prefix] $Message"
        }
        
        try {
            # Get the list
            Add-JobLog "Connecting to library: $LibraryName"
            $list = Get-PnPList -Identity $LibraryName -Connection $Connection
            
            # Get all items with files
            Add-JobLog "Retrieving files (batch size: $PageSize)..."
            $items = Get-PnPListItem -List $list -Connection $Connection -PageSize $PageSize -Fields "FileLeafRef"
            
            $total = $items.Count
            Add-JobLog "Found $total items with files"
            
            $counter = 0
            foreach ($item in $items) {
                $counter++
                
                try {
                    # Get file object
                    $file = Get-PnPProperty -ClientObject $item -Property File -Connection $Connection
                    
                    if ($file -and $file.Name) {
                        # Get versions
                        $versions = Get-PnPProperty -ClientObject $file -Property Versions -Connection $Connection
                        
                        if ($versions -and $versions.Count -gt $KeepCount) {
                            $toDelete = $versions.Count - $KeepCount
                            $percent = [math]::Round(($counter / $total) * 100, 1)
                            
                            Add-JobLog "[$counter/$total] [$percent%] $($file.Name): $($versions.Count) versions → keeping $KeepCount, removing $toDelete"
                            
                            if (-not $Simulate) {
                                try {
                                    # Delete oldest versions (keep newest $KeepCount)
                                    # Versions are ordered oldest first (index 0 = oldest)
                                    for ($i = 0; $i -lt $toDelete; $i++) {
                                        if ($versions.Count -gt 0) {
                                            $versions[0].DeleteObject()
                                        }
                                    }
                                    Invoke-PnPQuery -Connection $Connection
                                    $results.Deleted += $toDelete
                                    $results.Processed++
                                } catch {
                                    Add-JobLog "Failed to delete versions for $($file.Name): $($_.Exception.Message)" $true
                                    $results.Failed++
                                }
                            } else {
                                # Simulation mode - just count
                                $results.Deleted += $toDelete
                                $results.Processed++
                            }
                        } else {
                            $results.Skipped++
                        }
                    }
                } catch {
                    Add-JobLog "Error processing item #$counter : $($_.Exception.Message)" $true
                    $results.Failed++
                }
            }
            
            Add-JobLog "---------------------------------------"
            Add-JobLog "TRIM COMPLETE"
            Add-JobLog "Files Processed: $($results.Processed)"
            Add-JobLog "Files Skipped (no excess versions): $($results.Skipped)"
            Add-JobLog "Files Failed: $($results.Failed)"
            Add-JobLog "Total Versions Removed: $($results.Deleted)"
            
            return @{
                Success = $true
                Log = $results.Log
                Stats = $results
                WhatIf = $Simulate
            }
            
        } catch {
            Add-JobLog "Critical error: $($_.Exception.Message)" $true
            return @{
                Success = $false
                Log = $results.Log
                Error = $_.Exception.Message
            }
        }
    } -ArgumentList $script:Connection, $libraryName, $keep, $batch, $whatIf
    
    # Monitor the job with a timer
    $monitorTimer = New-Object System.Windows.Threading.DispatcherTimer
    $monitorTimer.Interval = [TimeSpan]::FromMilliseconds(1000)
    
    $monitorTimer.Add_Tick({
        if ($trimJob.State -eq "Completed") {
            $monitorTimer.Stop()
            $jobResult = Receive-Job -Job $trimJob
            Remove-Job -Job $trimJob
            
            # Display logs
            foreach ($line in $jobResult.Log) {
                if ($line -like "*ERROR*") {
                    Write-Log $line -IsError
                } else {
                    Write-Log $line
                }
            }
            
            if ($jobResult.Success) {
                $title = if ($jobResult.WhatIf) { "SIMULATION COMPLETE" } else { "TRIM COMPLETE" }
                $msg = "$title`n`nFiles Processed: $($jobResult.Stats.Processed)`nFiles Skipped: $($jobResult.Stats.Skipped)`nFiles Failed: $($jobResult.Stats.Failed)`nVersions Removed: $($jobResult.Stats.Deleted)"
                [System.Windows.MessageBox]::Show($msg, "Operation Complete", "OK", "Information")
            } else {
                [System.Windows.MessageBox]::Show("Error: $($jobResult.Error)", "Operation Failed", "OK", "Error")
            }
            
            # Reset UI
            $btnTrim.Content = "START VERSION TRIM"
            $btnTrim.IsEnabled = $true
            $grpConfig.IsEnabled = $true
            
        } elseif ($trimJob.State -eq "Failed") {
            $monitorTimer.Stop()
            $err = $trimJob.ChildJobs[0].JobStateInfo.Reason.Message
            Remove-Job -Job $trimJob
            Write-Log "Job failed: $err" -IsError
            [System.Windows.MessageBox]::Show("Processing failed: $err", "Error", "OK", "Error")
            
            $btnTrim.Content = "START VERSION TRIM"
            $btnTrim.IsEnabled = $true
            $grpConfig.IsEnabled = $true
        }
        # If still running, continue polling (logs will show when complete)
    })
    
    $monitorTimer.Start()
})

# Show Window
Write-Log "Ready. Enter connection details and click Connect."
$window.ShowDialog() | Out-Null


$cmbLibraries.Add_SelectionChanged({
    if ($cmbLibraries.SelectedItem) {
        Write-Log "Selected library: $($cmbLibraries.SelectedItem)"
        $btnTrim.IsEnabled = $true
    }
})

$btnTrim.Add_Click({
    $libraryName = $cmbLibraries.SelectedItem
    if (-not $libraryName) {
        [System.Windows.MessageBox]::Show("Please select a document library first.", "No Library Selected", "OK", "Warning")
        return
    }
    
    $keep = [int]$txtKeepVersions.Text
    $batch = [int]$txtBatchSize.Text
    $whatIf = $chkWhatIf.IsChecked
    
    if ($keep -lt 1) {
        [System.Windows.MessageBox]::Show("Keep Versions must be at least 1", "Invalid Input", "OK", "Warning")
        return
    }
    
    # Confirmation dialog
    $msgIcon = if ($whatIf) { "Information" } else { "Warning" }
    $confirmText = if ($whatIf) { 
        "SIMULATION MODE - No changes will be made.`n`nLibrary: $libraryName`nKeep Versions: $keep`nBatch Size: $batch" 
    } else { 
        "WARNING: This will PERMANENTLY DELETE versions!`n`nLibrary: $libraryName`nKeeping last $keep versions`nBatch Size: $batch" 
    }
    
    $result = [System.Windows.MessageBox]::Show($confirmText, "Confirm Operation", "OKCancel", $msgIcon)
    if ($result -ne "OK") { return }
    
    # UI Updates
    $btnTrim.IsEnabled = $false
    $btnTrim.Content = "PROCESSING..."
    Write-Log "======================================="
    Write-Log "STARTING VERSION TRIM"
    Write-Log "Library: $libraryName"
    Write-Log "Keep Versions: $keep | What-If: $whatIf"
    Write-Log "---------------------------------------"
    
    # Run in background
    $trimJob = Start-Job -ScriptBlock {
        param($Connection, $LibraryName, $KeepCount, $BatchSize, $WhatIf)
        
        $results = @{
            Log = @()
            Processed = 0
            Skipped = 0
            Failed = 0
            Deleted = 0
        }
        
        function Add-ResultLog($Message, $IsError = $false) {
            $ts = Get-Date -Format "HH:mm:ss"
            $results.Log += "[$ts] $Message"
        }
        
        try {
            $list = Get-PnPList -Identity $LibraryName -Connection $Connection
            Add-ResultLog "Scanning library: $LibraryName"
            
            # Get items with files
            $items = Get-PnPListItem -List $list -Connection $Connection -PageSize $BatchSize -Fields "FileRef"
            $total = $items.Count
            Add-ResultLog "Found $total items to process"
            
            $counter = 0
            foreach ($item in $items) {
                $counter++
                
                try {
                    $file = Get-PnPProperty -ClientObject $item -Property File -Connection $Connection
                    
                    if ($file -and $file.Name) {
                        $versions = Get-PnPProperty -ClientObject $file -Property Versions -Connection $Connection
                        
                        if ($versions -and $versions.Count -gt $KeepCount) {
                            $toDelete = $versions.Count - $KeepCount
                            $percent = [math]::Round(($counter / $total) * 100, 1)
                            Add-ResultLog "[$counter/$total][$percent%] $($file.Name) - v$($versions.Count) → keeping $KeepCount, removing $toDelete"
                            
                            if (-not $WhatIf) {
                                try {
                                    # Delete oldest versions (index 0 is oldest)
                                    for ($i = 0; $i -lt $toDelete; $i++) {
                                        if ($versions.Count -gt 0) {
                                            $versions[0].DeleteObject()
                                            Invoke-PnPQuery -Connection $Connection
                                        }
                                    }
                                    $results.Deleted += $toDelete
                                } catch {
                                    Add-ResultLog "ERROR deleting: $($_.Exception.Message)" $true
                                    $results.Failed++
                                }
                            } else {
                                $results.Deleted += $toDelete
                            }
                            $results.Processed++
                        } else {
                            $results.Skipped++
                        }
                    }
                } catch {
                    Add-ResultLog "ERROR processing item: $($_.Exception.Message)" $true
                    $results.Failed++
                }
            }
            
            Add-ResultLog "---------------------------------------"
            Add-ResultLog "COMPLETED"
            Add-ResultLog "Processed: $($results.Processed) | Skipped: $($results.Skipped) | Failed: $($results.Failed)"
            Add-ResultLog "Versions removed: $($results.Deleted)"
            
            return @{
                Success = $true
                Log = $results.Log
                Stats = $results
                WhatIf = $WhatIf
            }
            
        } catch {
            Add-ResultLog "CRITICAL ERROR: $($_.Exception.Message)" $true
            return @{
                Success = $false
                Log = $results.Log
                Error = $_.Exception.Message
            }
        }
    } -ArgumentList $script:Connection, $libraryName, $keep, $batch, $whatIf
    
    # Monitor job
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(800)
    
    $timer.Add_Tick({
        if ($trimJob.State -eq "Completed") {
            $timer.Stop()
            $jobResult = Receive-Job -Job $trimJob
            Remove-Job -Job $trimJob
            
            # Display logs
            foreach ($line in $jobResult.Log) {
                if ($line -like "*ERROR*") {
                    Write-Log $line -IsError
                } else {
                    Write-Log $line
                }
            }
            
            if ($jobResult.Success) {
                $title = if ($jobResult.WhatIf) { "SIMULATION COMPLETE" } else { "TRIM COMPLETE" }
                $msg = "$title`n`nProcessed: $($jobResult.Stats.Processed)`nSkipped: $($jobResult.Stats.Skipped)`nFailed: $($jobResult.Stats.Failed)`nVersions Removed: $($jobResult.Stats.Deleted)"
                [System.Windows.MessageBox]::Show($msg, "Operation Complete", "OK", "Information")
            } else {
                [System.Windows.MessageBox]::Show("Error: $($jobResult.Error)", "Operation Failed", "OK", "Error")
            }
            
            $btnTrim.Content = "START VERSION TRIM"
            $btnTrim.IsEnabled = $true
            
        } elseif ($trimJob.State -eq "Failed") {
            $timer.Stop()
            $err = $trimJob.ChildJobs[0].JobStateInfo.Reason.Message
            Remove-Job -Job $trimJob
            Write-Log "Job failed: $err" -IsError
            $btnTrim.Content = "START VERSION TRIM"
            $btnTrim.IsEnabled = $true
        }
    })
    
    $timer.Start()
})

# Show Window
Write-Log "SPO Version Trimmer Ready"
Write-Log "Enter Site URL, Client ID, and Tenant ID, then click Connect"
Write-Log "Note: Minimize this console window - all output shown in GUI"
$window.ShowDialog() | Out-Null

# Cleanup
if ($script:Connection) {
    Disconnect-PnPOnline -Connection $script:Connection -ErrorAction SilentlyContinue
    Write-Host "Disconnected from SharePoint" -ForegroundColor Green
}