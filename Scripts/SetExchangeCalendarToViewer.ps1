# Ensure Exchange Online module is loaded
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "ExchangeOnlineManagement module is required. Installing..."
    Install-Module ExchangeOnlineManagement -Force -Scope CurrentUser
}

# Connect to Exchange Online if not already
if (-not (Get-ConnectionInformation | Where-Object { $_.State -eq "Connected" })) {
    Connect-ExchangeOnline -UserPrincipalName (Read-Host "Enter your Exchange admin UPN")
}

Add-Type -AssemblyName PresentationFramework

# Define the WPF XAML UI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Calendar Permission Updater" Height="500" Width="800">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <TextBox Name="SearchBox" Grid.Row="0" Height="30" Margin="0,0,0,5" />

        <DataGrid Name="CalendarGrid" Grid.Row="1" AutoGenerateColumns="True" IsReadOnly="True" Margin="0,0,0,5" />

        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Right">
            <ProgressBar Name="ProgressBar" Width="200" Height="20" Margin="0,0,10,0"/>
            <Button Name="RunButton" Width="120" Height="30" Content="Set Permissions"/>
        </StackPanel>
    </Grid>
</Window>
"@

# Load WPF
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Link named controls
$SearchBox   = $window.FindName("SearchBox")
$CalendarGrid = $window.FindName("CalendarGrid")
$RunButton    = $window.FindName("RunButton")
$ProgressBar  = $window.FindName("ProgressBar")

if (-not ($SearchBox -and $CalendarGrid -and $RunButton -and $ProgressBar)) {
    Write-Error "Failed to initialize WPF controls. Ensure all control names match."
    return
}

# Fetch calendar permissions
Write-Host "Fetching calendar folders..."
try {
    $global:calendars = Get-Mailbox -RecipientTypeDetails UserMailbox |
        Get-MailboxFolderStatistics -FolderScope Calendar |
        Where-Object { $_.FolderType -eq "Calendar" } |
        Select-Object @{Name="Identity"; Expression={ $_.Identity.ToString().Replace("\", ":\") }}

    $CalendarGrid.ItemsSource = $global:calendars
} catch {
    [System.Windows.MessageBox]::Show("Error fetching calendars: $_","Error",[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error)
    return
}

# Search filter
$SearchBox.Add_TextChanged({
    $text = $SearchBox.Text.ToLower()
    $filtered = $global:calendars | Where-Object { $_.Identity.ToLower() -like "*$text*" }
    $CalendarGrid.ItemsSource = $filtered
})

# Run button click event
$RunButton.Add_Click({
    $RunButton.IsEnabled = $false
    $selectedItems = $CalendarGrid.ItemsSource
    $ProgressBar.Maximum = $selectedItems.Count
    $ProgressBar.Value = 0

    foreach ($entry in $selectedItems) {
        try {
            Set-MailboxFolderPermission -Identity $entry.Identity -User Default -AccessRights LimitedDetails -ErrorAction Stop
        } catch {
            Write-Warning "Failed to update $($entry.Identity): $_"
        }
        $ProgressBar.Value += 1
    }

    [System.Windows.MessageBox]::Show("Permissions applied successfully.","Done",[System.Windows.MessageBoxButton]::OK)
    $RunButton.IsEnabled = $true
})

# Show the GUI
$window.ShowDialog() | Out-Null
