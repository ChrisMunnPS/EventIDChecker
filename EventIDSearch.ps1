# Load assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

# Define categories, event IDs, and meanings
$categories = @{
    "AccountActivity" = @{ 
        Log = "Security"; 
        IDs = @(
            @{ ID = 4624; Meaning = "Successful logon" },
            @{ ID = 4625; Meaning = "Failed logon attempt" },
            @{ ID = 4634; Meaning = "Logoff" },
            @{ ID = 4647; Meaning = "User-initiated logoff" },
            @{ ID = 4648; Meaning = "Explicit credentials used" },
            @{ ID = 4672; Meaning = "Security ID assigned to user" },
            @{ ID = 4768; Meaning = "Kerberos TGT request" },
            @{ ID = 4769; Meaning = "Kerberos service ticket request" },
            @{ ID = 4771; Meaning = "Kerberos pre-authentication failed" },
            @{ ID = 4776; Meaning = "NTLM authentication attempt" }
        )
    }
    "ADAccountChanges" = @{ 
        Log = "Security"; 
        IDs = @(
            @{ ID = 4720; Meaning = "User account created" },
            @{ ID = 4722; Meaning = "User account enabled" },
            @{ ID = 4723; Meaning = "Password change attempt" },
            @{ ID = 4724; Meaning = "Password reset attempt" },
            @{ ID = 4725; Meaning = "User account disabled" },
            @{ ID = 4726; Meaning = "User account deleted" },
            @{ ID = 4732; Meaning = "User added to security group" },
            @{ ID = 4735; Meaning = "Security group modified" },
            @{ ID = 4738; Meaning = "User account modified" },
            @{ ID = 4740; Meaning = "Account locked out" },
            @{ ID = 4756; Meaning = "Security group membership change" },
            @{ ID = 4767; Meaning = "Account unlocked" }
        )
    }
    "SecurityThreatIndicators" = @{ 
        Log = "Security"; 
        IDs = @(
            @{ ID = 1102; Meaning = "Audit log cleared" },
            @{ ID = 2886; Meaning = "LDAP unsigned/simple bind detected" },
            @{ ID = 2887; Meaning = "Count of unsigned/simple bind attempts" },
            @{ ID = 2889; Meaning = "Source of unsigned/simple bind" },
            @{ ID = 1644; Meaning = "Expensive LDAP query detected" },
            @{ ID = 4627; Meaning = "Group membership information" },
            @{ ID = 4663; Meaning = "Access to an object" }
        )
    }
    "ServerHealthReliability" = @{ 
        Log = "System"; 
        IDs = @(
            @{ ID = 41; Meaning = "Kernel-Power: unexpected restart/shutdown" },
            @{ ID = 55; Meaning = "NTFS file system corruption detected" },
            @{ ID = 6005; Meaning = "Event log service started" },
            @{ ID = 6006; Meaning = "Event log service stopped" },
            @{ ID = 6008; Meaning = "Unexpected shutdown" },
            @{ ID = 6009; Meaning = "System startup information" },
            @{ ID = 1074; Meaning = "System shutdown/restart initiated" },
            @{ ID = 1014; Meaning = "DNS name resolution failure" },
            @{ ID = 1058; Meaning = "Group Policy failure to read from DC" },
            @{ ID = 5719; Meaning = "Netlogon: no DC available" }
        )
    }
    "ApplicationLevelIssues" = @{ 
        Log = "Application"; 
        IDs = @(
            @{ ID = 1000; Meaning = "Application error (crash)" },
            @{ ID = 1001; Meaning = "Application hang or bugcheck info" },
            @{ ID = 1002; Meaning = "Application hang" },
            @{ ID = 1309; Meaning = "ASP.NET application error (IIS)" },
            @{ ID = 11707; Meaning = "Application installation" }
        )
    }
}

# Global variables
$global:allEvents = @()
$global:autoRefreshTimer = $null
$global:currentPage = 1
$global:pageSize = 100

# XAML for UI
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Enhanced Event Log Viewer" Height="700" Width="1000">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <!-- Category Selection -->
        <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="10">
            <RadioButton x:Name="rbAccountActivity" Content="Account Activity" GroupName="Category" IsChecked="True" Margin="5"/>
            <RadioButton x:Name="rbADAccountChanges" Content="AD Account Changes" GroupName="Category" Margin="5"/>
            <RadioButton x:Name="rbSecurityThreat" Content="Security Threat Indicators" GroupName="Category" Margin="5"/>
            <RadioButton x:Name="rbServerHealth" Content="Server Health and Reliability" GroupName="Category" Margin="5"/>
            <RadioButton x:Name="rbApplicationIssues" Content="Application-Level Issues" GroupName="Category" Margin="5"/>
        </StackPanel>
        
        <!-- Date and Computer Selection -->
        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="10">
            <Label Content="Computer:" Margin="5"/>
            <TextBox x:Name="txtComputer" Width="150" Margin="5" Text="localhost"/>
            <Label Content="Start Date:" Margin="5"/>
            <DatePicker x:Name="dpStart" Margin="5"/>
            <Label Content="End Date:" Margin="5"/>
            <DatePicker x:Name="dpEnd" Margin="5"/>
        </StackPanel>
        
        <!-- Event ID and Text Filter -->
        <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="10">
            <Label Content="Event ID:" Margin="5"/>
            <ComboBox x:Name="cbEventID" Width="200" Margin="5"/>
            <Label Content="Search Text:" Margin="5"/>
            <TextBox x:Name="txtSearchText" Width="200" Margin="5"/>
        </StackPanel>
        
        <!-- Action Buttons -->
        <StackPanel Grid.Row="3" Orientation="Horizontal" Margin="10">
            <Button x:Name="btnSearch" Content="Search" Width="80" Height="30" Margin="5"/>
            <Button x:Name="btnClear" Content="Clear" Width="80" Height="30" Margin="5"/>
            <Button x:Name="btnExportCSV" Content="Export CSV" Width="80" Height="30" Margin="5"/>
            <Button x:Name="btnExportExcel" Content="Export Excel" Width="80" Height="30" Margin="5"/>
            <Button x:Name="btnWebSearch" Content="Web Search" Width="85" Height="30" Margin="5" ToolTip="Search web for selected event (Event ID + Source)"/>
            <CheckBox x:Name="chkAutoRefresh" Content="Auto-Refresh (30s)" Margin="15,5,5,5" VerticalAlignment="Center"/>
            <Label x:Name="lblStatus" Content="" Margin="15,5,5,5" VerticalAlignment="Center" Foreground="Blue"/>
        </StackPanel>
        
        <!-- Progress Bar -->
        <ProgressBar x:Name="pbProgress" Grid.Row="4" Height="20" Margin="10,0,10,5" Visibility="Collapsed"/>
        
        <!-- Results Grid -->
        <DataGrid x:Name="dgResults" Grid.Row="5" AutoGenerateColumns="False" Margin="10" IsReadOnly="True" 
                  AlternatingRowBackground="#F0F0F0" RowHeight="30" GridLinesVisibility="Horizontal"
                  HorizontalGridLinesBrush="#E0E0E0">
            <DataGrid.RowStyle>
                <Style TargetType="DataGridRow">
                    <Setter Property="BorderThickness" Value="0,0,0,1"/>
                    <Setter Property="BorderBrush" Value="#D0D0D0"/>
                    <Style.Triggers>
                        <Trigger Property="IsMouseOver" Value="True">
                            <Setter Property="Background" Value="#E3F2FD"/>
                        </Trigger>
                        <Trigger Property="IsSelected" Value="True">
                            <Setter Property="Background" Value="#BBDEFB"/>
                        </Trigger>
                    </Style.Triggers>
                </Style>
            </DataGrid.RowStyle>
            <DataGrid.Columns>
                <DataGridTextColumn Header="Time Created" Binding="{Binding TimeCreated}" Width="140">
                    <DataGridTextColumn.ElementStyle>
                        <Style TargetType="TextBlock">
                            <Setter Property="Padding" Value="5"/>
                            <Setter Property="VerticalAlignment" Value="Center"/>
                        </Style>
                    </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Log Name" Binding="{Binding LogName}" Width="100">
                    <DataGridTextColumn.ElementStyle>
                        <Style TargetType="TextBlock">
                            <Setter Property="Padding" Value="5"/>
                            <Setter Property="VerticalAlignment" Value="Center"/>
                            <Setter Property="FontWeight" Value="SemiBold"/>
                        </Style>
                    </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="150">
                    <DataGridTextColumn.ElementStyle>
                        <Style TargetType="TextBlock">
                            <Setter Property="Padding" Value="5"/>
                            <Setter Property="VerticalAlignment" Value="Center"/>
                        </Style>
                    </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Event ID" Binding="{Binding Id}" Width="80">
                    <DataGridTextColumn.ElementStyle>
                        <Style TargetType="TextBlock">
                            <Setter Property="Padding" Value="5"/>
                            <Setter Property="VerticalAlignment" Value="Center"/>
                            <Setter Property="FontWeight" Value="Bold"/>
                        </Style>
                    </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Meaning" Binding="{Binding Meaning}" Width="200">
                    <DataGridTextColumn.ElementStyle>
                        <Style TargetType="TextBlock">
                            <Setter Property="Padding" Value="5"/>
                            <Setter Property="VerticalAlignment" Value="Center"/>
                            <Setter Property="FontStyle" Value="Italic"/>
                        </Style>
                    </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Level" Binding="{Binding LevelDisplayName}" Width="80">
                    <DataGridTextColumn.ElementStyle>
                        <Style TargetType="TextBlock">
                            <Setter Property="Padding" Value="5"/>
                            <Setter Property="VerticalAlignment" Value="Center"/>
                        </Style>
                    </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Message" Binding="{Binding Message}" Width="*">
                    <DataGridTextColumn.ElementStyle>
                        <Style TargetType="TextBlock">
                            <Setter Property="Padding" Value="5"/>
                            <Setter Property="VerticalAlignment" Value="Center"/>
                            <Setter Property="TextWrapping" Value="Wrap"/>
                        </Style>
                    </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
            </DataGrid.Columns>
        </DataGrid>
        
        <!-- Pagination Controls -->
        <StackPanel Grid.Row="6" Orientation="Horizontal" Margin="10" HorizontalAlignment="Center">
            <Button x:Name="btnFirstPage" Content="⏮ First" Width="70" Height="25" Margin="5"/>
            <Button x:Name="btnPrevPage" Content="⏪ Previous" Width="80" Height="25" Margin="5"/>
            <Label x:Name="lblPageInfo" Content="Page 1" Margin="10,0" VerticalAlignment="Center"/>
            <Button x:Name="btnNextPage" Content="Next ⏩" Width="80" Height="25" Margin="5"/>
            <Button x:Name="btnLastPage" Content="Last ⏭" Width="70" Height="25" Margin="5"/>
            <Label Content="Page Size:" Margin="15,0,5,0" VerticalAlignment="Center"/>
            <ComboBox x:Name="cbPageSize" Width="70" Margin="5" SelectedIndex="1">
                <ComboBoxItem Content="50"/>
                <ComboBoxItem Content="100"/>
                <ComboBoxItem Content="250"/>
                <ComboBoxItem Content="500"/>
            </ComboBox>
        </StackPanel>
        
        <!-- Status Bar -->
        <StatusBar Grid.Row="7">
            <StatusBarItem>
                <TextBlock x:Name="txtStatusBar" Text="Ready"/>
            </StatusBarItem>
        </StatusBar>
    </Grid>
</Window>
"@

# Load XAML
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Find controls
$rbAccountActivity = $window.FindName("rbAccountActivity")
$rbADAccountChanges = $window.FindName("rbADAccountChanges")
$rbSecurityThreat = $window.FindName("rbSecurityThreat")
$rbServerHealth = $window.FindName("rbServerHealth")
$rbApplicationIssues = $window.FindName("rbApplicationIssues")
$txtComputer = $window.FindName("txtComputer")
$dpStart = $window.FindName("dpStart")
$dpEnd = $window.FindName("dpEnd")
$cbEventID = $window.FindName("cbEventID")
$txtSearchText = $window.FindName("txtSearchText")
$btnSearch = $window.FindName("btnSearch")
$btnClear = $window.FindName("btnClear")
$btnExportCSV = $window.FindName("btnExportCSV")
$btnExportExcel = $window.FindName("btnExportExcel")
$btnWebSearch = $window.FindName("btnWebSearch")
$chkAutoRefresh = $window.FindName("chkAutoRefresh")
$lblStatus = $window.FindName("lblStatus")
$pbProgress = $window.FindName("pbProgress")
$dgResults = $window.FindName("dgResults")
$btnFirstPage = $window.FindName("btnFirstPage")
$btnPrevPage = $window.FindName("btnPrevPage")
$lblPageInfo = $window.FindName("lblPageInfo")
$btnNextPage = $window.FindName("btnNextPage")
$btnLastPage = $window.FindName("btnLastPage")
$cbPageSize = $window.FindName("cbPageSize")
$txtStatusBar = $window.FindName("txtStatusBar")

# Set default dates
$dpStart.SelectedDate = (Get-Date).AddDays(-3)
$dpEnd.SelectedDate = Get-Date

# Function to update ComboBox with Event IDs based on selected category
function Update-EventIDComboBox {
    $selectedCategory = ""
    if ($rbAccountActivity.IsChecked) { $selectedCategory = "AccountActivity" }
    elseif ($rbADAccountChanges.IsChecked) { $selectedCategory = "ADAccountChanges" }
    elseif ($rbSecurityThreat.IsChecked) { $selectedCategory = "SecurityThreatIndicators" }
    elseif ($rbServerHealth.IsChecked) { $selectedCategory = "ServerHealthReliability" }
    elseif ($rbApplicationIssues.IsChecked) { $selectedCategory = "ApplicationLevelIssues" }

    if ($selectedCategory -ne "") {
        $cbEventID.Items.Clear()
        $cbEventID.Items.Add("(All)") | Out-Null
        $categories[$selectedCategory].IDs | ForEach-Object {
            $cbEventID.Items.Add("$($_.ID) - $($_.Meaning)") | Out-Null
        }
        $cbEventID.SelectedIndex = 0
    }
}

# Function to update pagination display
function Update-PaginationDisplay {
    $totalEvents = $global:allEvents.Count
    $totalPages = [Math]::Ceiling($totalEvents / $global:pageSize)
    
    if ($totalPages -eq 0) { $totalPages = 1 }
    if ($global:currentPage -gt $totalPages) { $global:currentPage = $totalPages }
    if ($global:currentPage -lt 1) { $global:currentPage = 1 }
    
    $lblPageInfo.Content = "Page $($global:currentPage) of $totalPages"
    $txtStatusBar.Text = "Total Events: $totalEvents | Displaying: $(($global:currentPage - 1) * $global:pageSize + 1) - $([Math]::Min($global:currentPage * $global:pageSize, $totalEvents))"
    
    $btnFirstPage.IsEnabled = $global:currentPage -gt 1
    $btnPrevPage.IsEnabled = $global:currentPage -gt 1
    $btnNextPage.IsEnabled = $global:currentPage -lt $totalPages
    $btnLastPage.IsEnabled = $global:currentPage -lt $totalPages
    
    # Display current page
    $startIndex = ($global:currentPage - 1) * $global:pageSize
    $endIndex = [Math]::Min($startIndex + $global:pageSize, $totalEvents)
    $pageData = $global:allEvents[$startIndex..($endIndex - 1)]
    $dgResults.ItemsSource = $pageData
}

# Function to perform search
function Invoke-EventSearch {
    $selectedCategory = ""
    if ($rbAccountActivity.IsChecked) { $selectedCategory = "AccountActivity" }
    elseif ($rbADAccountChanges.IsChecked) { $selectedCategory = "ADAccountChanges" }
    elseif ($rbSecurityThreat.IsChecked) { $selectedCategory = "SecurityThreatIndicators" }
    elseif ($rbServerHealth.IsChecked) { $selectedCategory = "ServerHealthReliability" }
    elseif ($rbApplicationIssues.IsChecked) { $selectedCategory = "ApplicationLevelIssues" }

    if ($selectedCategory -ne "") {
        $cat = $categories[$selectedCategory]
        $start = $dpStart.SelectedDate
        $end = $dpEnd.SelectedDate.AddDays(1)
        $computer = $txtComputer.Text
        
        if ([string]::IsNullOrWhiteSpace($computer)) { $computer = "localhost" }

        # Determine Event IDs to filter
        $eventIDs = if ($cbEventID.SelectedItem -eq "(All)") {
            $cat.IDs.ID
        } else {
            $selectedID = $cbEventID.SelectedItem -split " - " | Select-Object -First 1
            @([int]$selectedID)
        }

        # Show progress
        $pbProgress.Visibility = "Visible"
        $pbProgress.IsIndeterminate = $true
        $lblStatus.Content = "Searching..."
        $txtStatusBar.Text = "Querying event logs..."

        try {
            $filter = @{
                LogName = $cat.Log
                ID = $eventIDs
                StartTime = $start
                EndTime = $end
            }
            
            # Add computer name if not localhost
            if ($computer -ne "localhost") {
                $events = Get-WinEvent -FilterHashtable $filter -ComputerName $computer -ErrorAction Stop
            } else {
                $events = Get-WinEvent -FilterHashtable $filter -ErrorAction Stop
            }
            
            $lblStatus.Content = "Processing results..."
            
            $eventData = $events | ForEach-Object {
                $id = $_.Id
                $meaning = ($cat.IDs | Where-Object { $_.ID -eq $id }).Meaning
                [PSCustomObject]@{
                    TimeCreated = $_.TimeCreated
                    LogName = $_.LogName
                    Source = $_.ProviderName
                    Id = $id
                    Meaning = $meaning
                    LevelDisplayName = $_.LevelDisplayName
                    Message = $_.Message
                    FullEvent = $_
                }
            }
            
            # Apply text filter if specified
            $searchText = $txtSearchText.Text
            if (-not [string]::IsNullOrWhiteSpace($searchText)) {
                $eventData = $eventData | Where-Object { $_.Message -like "*$searchText*" }
            }
            
            $global:allEvents = $eventData
            $global:currentPage = 1
            Update-PaginationDisplay
            
            $lblStatus.Content = "Search complete"
            $txtStatusBar.Text = "Found $($global:allEvents.Count) events"
            
        } catch {
            [System.Windows.MessageBox]::Show("Error retrieving events: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
            $lblStatus.Content = "Error occurred"
            $txtStatusBar.Text = "Error: $($_.Exception.Message)"
        } finally {
            $pbProgress.Visibility = "Collapsed"
            $pbProgress.IsIndeterminate = $false
        }
    }
}

# Attach event handlers to radio buttons
$rbAccountActivity.Add_Checked({ Update-EventIDComboBox })
$rbADAccountChanges.Add_Checked({ Update-EventIDComboBox })
$rbSecurityThreat.Add_Checked({ Update-EventIDComboBox })
$rbServerHealth.Add_Checked({ Update-EventIDComboBox })
$rbApplicationIssues.Add_Checked({ Update-EventIDComboBox })

# Initialize ComboBox
Update-EventIDComboBox

# Search button click event
$btnSearch.Add_Click({
    Invoke-EventSearch
})

# Clear button click event
$btnClear.Add_Click({
    $global:allEvents = @()
    $global:currentPage = 1
    $dgResults.ItemsSource = $null
    $lblStatus.Content = ""
    $txtStatusBar.Text = "Ready"
    Update-PaginationDisplay
})

# Export CSV button
$btnExportCSV.Add_Click({
    if ($global:allEvents.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No data to export", "Warning", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "CSV files (*.csv)|*.csv"
    $saveDialog.FileName = "EventLog_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            $global:allEvents | Select-Object TimeCreated, LogName, Source, Id, Meaning, LevelDisplayName, Message | Export-Csv -Path $saveDialog.FileName -NoTypeInformation
            [System.Windows.MessageBox]::Show("Data exported successfully to $($saveDialog.FileName)", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            [System.Windows.MessageBox]::Show("Error exporting data: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Export Excel button
$btnExportExcel.Add_Click({
    if ($global:allEvents.Count -eq 0) {
        [System.Windows.MessageBox]::Show("No data to export", "Warning", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
    $saveDialog.FileName = "EventLog_Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
    
    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            # Create Excel COM object
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $workbook = $excel.Workbooks.Add()
            $worksheet = $workbook.Worksheets.Item(1)
            
            # Add headers
            $worksheet.Cells.Item(1, 1) = "Time Created"
            $worksheet.Cells.Item(1, 2) = "Log Name"
            $worksheet.Cells.Item(1, 3) = "Source"
            $worksheet.Cells.Item(1, 4) = "Event ID"
            $worksheet.Cells.Item(1, 5) = "Meaning"
            $worksheet.Cells.Item(1, 6) = "Level"
            $worksheet.Cells.Item(1, 7) = "Message"
            
            # Format headers
            $headerRange = $worksheet.Range("A1", "G1")
            $headerRange.Font.Bold = $true
            $headerRange.Interior.ColorIndex = 15
            
            # Add data
            $row = 2
            foreach ($evt in $global:allEvents) {
                $worksheet.Cells.Item($row, 1) = $evt.TimeCreated.ToString()
                $worksheet.Cells.Item($row, 2) = $evt.LogName
                $worksheet.Cells.Item($row, 3) = $evt.Source
                $worksheet.Cells.Item($row, 4) = $evt.Id
                $worksheet.Cells.Item($row, 5) = $evt.Meaning
                $worksheet.Cells.Item($row, 6) = $evt.LevelDisplayName
                $worksheet.Cells.Item($row, 7) = $evt.Message
                $row++
            }
            
            # Auto-fit columns
            $worksheet.UsedRange.Columns.AutoFit() | Out-Null
            
            # Save and close
            $workbook.SaveAs($saveDialog.FileName)
            $workbook.Close()
            $excel.Quit()
            
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
            
            [System.Windows.MessageBox]::Show("Data exported successfully to $($saveDialog.FileName)", "Success", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        } catch {
            [System.Windows.MessageBox]::Show("Error exporting to Excel: $($_.Exception.Message)`n`nMake sure Excel is installed.", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        }
    }
})

# Web Search button
$btnWebSearch.Add_Click({
    $selectedEvent = $dgResults.SelectedItem
    if ($selectedEvent -eq $null) {
        [System.Windows.MessageBox]::Show("Please select an event to search for.", "No Selection", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        return
    }
    
    # Create search query with Event ID and Source
    $eventId = $selectedEvent.Id
    $source = $selectedEvent.Source
    $logName = $selectedEvent.LogName
    
    # Build search query - using Event ID + Source for most specific results
    $searchQuery = "Event ID $eventId $source Windows"
    
    # URL encode the search query
    $encodedQuery = [System.Uri]::EscapeDataString($searchQuery)
    
    # Build search URLs for multiple engines
    $googleUrl = "https://www.google.com/search?q=$encodedQuery"
    $duckduckgoUrl = "https://duckduckgo.com/?q=$encodedQuery"
    
    # Create custom dialog for search engine selection
    [xml]$searchDialogXaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Select Search Engine" Height="200" Width="400" WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        
        <TextBlock Grid.Row="0" TextWrapping="Wrap" Margin="0,0,0,15">
            <Run FontWeight="Bold">Search for:</Run>
            <LineBreak/>
            <Run Text="Event ID $eventId - $source"/>
        </TextBlock>
        
        <StackPanel Grid.Row="1" VerticalAlignment="Center">
            <TextBlock Text="Choose a search engine:" Margin="0,0,0,10"/>
        </StackPanel>
        
        <StackPanel Grid.Row="2" Orientation="Horizontal" HorizontalAlignment="Center">
            <Button x:Name="btnGoogle" Content="Google" Width="100" Height="35" Margin="5"/>
            <Button x:Name="btnDuckDuckGo" Content="DuckDuckGo" Width="100" Height="35" Margin="5"/>
            <Button x:Name="btnCancel" Content="Cancel" Width="100" Height="35" Margin="5"/>
        </StackPanel>
    </Grid>
</Window>
"@
    
    try {
        $searchDialogReader = New-Object System.Xml.XmlNodeReader $searchDialogXaml
        $searchDialog = [Windows.Markup.XamlReader]::Load($searchDialogReader)
        
        $btnGoogle = $searchDialog.FindName("btnGoogle")
        $btnDuckDuckGo = $searchDialog.FindName("btnDuckDuckGo")
        $btnCancel = $searchDialog.FindName("btnCancel")
        
        $searchDialog.Tag = $null
        
        $btnGoogle.Add_Click({
            $searchDialog.Tag = "Google"
            $searchDialog.Close()
        })
        
        $btnDuckDuckGo.Add_Click({
            $searchDialog.Tag = "DuckDuckGo"
            $searchDialog.Close()
        })
        
        $btnCancel.Add_Click({
            $searchDialog.Tag = "Cancel"
            $searchDialog.Close()
        })
        
        $searchDialog.ShowDialog() | Out-Null
        
        if ($searchDialog.Tag -eq "Google") {
            Start-Process $googleUrl
        } elseif ($searchDialog.Tag -eq "DuckDuckGo") {
            Start-Process $duckduckgoUrl
        }
        
    } catch {
        [System.Windows.MessageBox]::Show("Error opening browser: $($_.Exception.Message)", "Error", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    }
})

# Auto-refresh checkbox
$chkAutoRefresh.Add_Checked({
    $global:autoRefreshTimer = New-Object System.Windows.Threading.DispatcherTimer
    $global:autoRefreshTimer.Interval = [TimeSpan]::FromSeconds(30)
    $global:autoRefreshTimer.Add_Tick({
        Invoke-EventSearch
    })
    $global:autoRefreshTimer.Start()
    $lblStatus.Content = "Auto-refresh enabled"
})

$chkAutoRefresh.Add_Unchecked({
    if ($global:autoRefreshTimer -ne $null) {
        $global:autoRefreshTimer.Stop()
        $global:autoRefreshTimer = $null
        $lblStatus.Content = "Auto-refresh disabled"
    }
})

# Pagination button events
$btnFirstPage.Add_Click({
    $global:currentPage = 1
    Update-PaginationDisplay
})

$btnPrevPage.Add_Click({
    if ($global:currentPage -gt 1) {
        $global:currentPage--
        Update-PaginationDisplay
    }
})

$btnNextPage.Add_Click({
    $totalPages = [Math]::Ceiling($global:allEvents.Count / $global:pageSize)
    if ($global:currentPage -lt $totalPages) {
        $global:currentPage++
        Update-PaginationDisplay
    }
})

$btnLastPage.Add_Click({
    $totalPages = [Math]::Ceiling($global:allEvents.Count / $global:pageSize)
    $global:currentPage = $totalPages
    Update-PaginationDisplay
})

# Page size changed
$cbPageSize.Add_SelectionChanged({
    $global:pageSize = [int]$cbPageSize.SelectedItem.Content
    $global:currentPage = 1
    Update-PaginationDisplay
})

# Double-click to view event details
$dgResults.Add_MouseDoubleClick({
    $selectedEvent = $dgResults.SelectedItem
    if ($selectedEvent -ne $null) {
        $details = @"
Event Details
=============
Time: $($selectedEvent.TimeCreated)
Log Name: $($selectedEvent.LogName)
Source: $($selectedEvent.Source)
Event ID: $($selectedEvent.Id)
Meaning: $($selectedEvent.Meaning)
Level: $($selectedEvent.LevelDisplayName)

Message:
$($selectedEvent.Message)

Full Event Properties:
$(if ($selectedEvent.FullEvent) { $selectedEvent.FullEvent | Format-List * | Out-String } else { "Not available" })
"@
        [System.Windows.MessageBox]::Show($details, "Event Details", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
    }
})

# Window closing event - cleanup auto-refresh timer
$window.Add_Closing({
    if ($global:autoRefreshTimer -ne $null) {
        $global:autoRefreshTimer.Stop()
        $global:autoRefreshTimer = $null
    }
})

# Show window
$window.ShowDialog() | Out-Null
