# Path to the log file
$logFilePath = ".\Log.txt"

# Function to write log messages to the log file
function Write-Log($message) {
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $message"
    $logMessage | Out-File -FilePath $logFilePath -Append
}

# Add a trap to handle script termination
$ErrorActionPreference = 'Stop'
trap {
    Write-Log "Script terminated with error: $($_.Exception.Message)"
    Exit 1
}

try {

Write-Log "Log Started : Loading the HTML Agility Pack assembly..."
# Load the HTML Agility Pack assembly
Add-Type -Path "HtmlAgilityPack.dll"

Write-Log "Importing the required ADO.NET assembly..."
# Import the required ADO.NET assembly
Add-Type -Path "System.Data.SqlClient.dll"

# Access the COMPUTERNAME environment variable
$computerName = $env:COMPUTERNAME
Write-Log "Computer Name: $computerName"

Write-Log "HTML Generation is in process..."
Write-Host "HTML Generation is in process..."
######--------------------HTML Generation ---------------------------#################

# Current date and time
$currentTime = Get-Date

# Create an HTML report
$htmlReport = @"
<!DOCTYPE html>
<html>
<head>
    <title>System Information Report</title>
    <style>
        body { font-family: Arial, sans-serif; }
        h1 { text-align: center; }
        table { border-collapse: collapse; width: 100%; }
        th, td { border: 1px solid black; padding: 8px; text-align: left; }
        th { background-color: #f2f2f2; }
    </style>
</head>
<body>
    <h1>System Information Report - $currentTime</h1>
"@
####################

########### System Information ##############
Write-Log "HTML report generating for : System Information"
Write-Host "HTML report generating for : System Information"
$systemInfo = Get-WmiObject Win32_ComputerSystem | Select-Object Name, Manufacturer, Model, UserName, @{Name='LastLogon'; Expression={Get-LastLogonDate $_.UserName}}

# Get the system IP address
$ipAddress = (Get-NetIPAddress | Where-Object { $_.InterfaceAlias -eq 'Wi-Fi' -and $_.AddressFamily -eq 'IPv4' }).IPAddress

$htmlReport += @"
    <table align="center" style="table-layout: fixed;WORD-BREAK:BREAK-ALL;margin:10px"><tr><td colspan="2" align="center"><b>System Information</b></td></tr><tr><td align="center"><b>SysInfo_Name</b></td><td align="center"><b>SysInfo_Details</b></td></tr>


"@
$systemInfo | ForEach-Object {
    $computerName = $_.Name
    $manufacturer = $_.Manufacturer
    $model = $_.Model
    $userName = $_.UserName
    $lastLogon = $_.LastLogon
    $htmlReport += "<tr><td>Computer Name</td><td>$computerName</td></tr>
	<tr><td>Manufacturer</td><td>$manufacturer</td></tr>
	<tr><td>Model</td><td>$model</td></tr>
	<tr><td>User Name</td><td>$userName</td></tr>
	<tr><td>IP Address</td><td>$ipAddress</td></tr>"
}

$htmlReport += @"
    </table>
"@
##############################################

##########Operating System Information########
Write-Log "HTML report generating for : Operating System Information"
Write-Host "HTML report generating for : Operating System Information"
$osInfo = Get-WmiObject Win32_OperatingSystem | Select-Object Caption, Version

$htmlReport += @"
    <table align="center" style="table-layout: fixed;WORD-BREAK:BREAK-ALL;margin:10px"><tr><td colspan="2" align="center"><b>Operating System Information</b></td></tr><tr><td align="center"><b>OS_Details</b></td><td align="center"><b>OS_Values</b></td></tr>
"@
$osInfo | ForEach-Object {
    $caption = $_.Caption
    $version = $_.Version
    $htmlReport += "<tr><td>Caption</td><td>$caption</td></tr>
    <tr><td>Version</td><td>$version</td></tr>"
}

$htmlReport += @"
    </table>
"@

##############################################

################CPU Information###############
Write-Log "HTML report generating for : CPU Information"
Write-Host "HTML report generating for : CPU Information"
$cpuInfo = Get-WmiObject Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors

$htmlReport += @"
    <table align="center" style="table-layout: fixed;WORD-BREAK:BREAK-ALL;margin:10px"><tr><td colspan="2" align="center"><b>CPU Information</b></td></tr><tr><td align="center"><b>CPU_Details</b></td><td align="center"><b>CPU_Values</b></td></tr>
        
"@
$cpuInfo | ForEach-Object {
    $cpuName = $_.Name
    $cores = $_.NumberOfCores
    $logicalProcessors = $_.NumberOfLogicalProcessors
    $htmlReport += "<tr><td>Name</td><td>$cpuName</td></tr>
    <tr><td>Number of Cores</td><td>$cores</td></tr>
    <tr><td>Number of Logical Processors</td><td>$logicalProcessors</td></tr>"
}

$htmlReport += @"
    </table>
"@
##############################################

################Memory Information############
Write-Log "HTML report generating for : Memory Information"
Write-Host "HTML report generating for : Memory Information"
$memoryInfo = Get-WmiObject Win32_ComputerSystem | Select-Object TotalPhysicalMemory

$htmlReport += @"
    <table align="center" style="table-layout: fixed;WORD-BREAK:BREAK-ALL;margin:10px"><tr><td colspan="2" align="center"><b>Memory Information</b></td></tr><tr><td align="center"><b>Memory_Details</b></td><td align="center"><b>Memory_Values</b></td></tr>
        
"@
$memoryInfo | ForEach-Object {
    $totalMemory = [math]::Round($_.TotalPhysicalMemory / 1GB, 2)
    $htmlReport += "<tr><td>Total Physical Memory</td><td>$totalMemory GB</td></tr>"
}

$htmlReport += @"
    </table>
"@
##############################################

################Disk Information##############
Write-Log "HTML report generating for : Disk Information"
Write-Host "HTML report generating for : Disk Information"
# Function to convert bytes to a human-readable format
function ConvertTo-HumanReadableSize {
    param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [double]$SizeInBytes
    )
    process {
        $units = "B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"
        $size = $SizeInBytes
        $unitIndex = 0

        while ($size -ge 1024 -and $unitIndex -lt ($units.Length - 1)) {
            $size /= 1024
            $unitIndex++
        }

        [pscustomobject]@{
            Size = [math]::Round($size, 2)
            Unit = $units[$unitIndex]
        }
    }
}

# Get disk drive information
$drives = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }

$htmlReport += @"
    <table align="center" style="table-layout: fixed;WORD-BREAK:BREAK-ALL;margin:10px"><tr><td colspan="6" align="center"><b>Disk Drive Information</b></td></tr>
        <tr>
            <td><b>Drive Letter</b></td>
            <td><b>Label</b></td>
            <td><b>Disk Number</b></td>
            <td><b>Total Space</b></td>
            <td><b>Used Space</b></td>
            <td><b>Free Space</b></td>
        </tr>
"@

foreach ($drive in $drives) {
    $diskNumber = $drive.DeviceID -replace '\\', ''
    $totalSpace = ConvertTo-HumanReadableSize $drive.Size
    $usedSpace = ConvertTo-HumanReadableSize ($drive.Size - $drive.FreeSpace)
    $freeSpace = ConvertTo-HumanReadableSize $drive.FreeSpace

    $htmlReport += @"
        <tr>
            <td>$($drive.DeviceID)</td>
            <td>$($drive.VolumeName)</td>
            <td>$diskNumber</td>
            <td>$($totalSpace.Size) $($totalSpace.Unit)</td>
            <td>$($usedSpace.Size) $($usedSpace.Unit)</td>
            <td>$($freeSpace.Size) $($freeSpace.Unit)</td>
        </tr>
"@
}

$htmlReport += @"
    </table>
"@
##############################################

##########User Account Information############
Write-Log "HTML report generating for : User Account Information"
Write-Host "HTML report generating for : User Account Information"
$localUsers = Get-WmiObject Win32_UserAccount | Where-Object { $_.LocalAccount -eq $true }

$htmlReport += @"
<table align="center" style="table-layout: fixed;WORD-BREAK:BREAK-ALL;margin:10px"><tr><td colspan="3" align="center"><b>Local Users Information</b></td></tr>
    <tr>
        <td><b>Username</b></td>
        <td><b>Last Login</b></td>
        <td><b>Password Expiry</b></td>
    </tr>
"@

foreach ($user in $localUsers) {
    $username = $user.Name
    $lastLogin = $null
    $passwordExpiry = $null

    # Get last login date
    try {
        $lastLoginEvent = Get-WinEvent -FilterHashtable @{
            LogName = 'Security'
            ID = 4624
            UserID = $user.SID
        } -MaxEvents 1 -ErrorAction Stop

        $lastLogin = $lastLoginEvent.TimeCreated
    }
    catch {
        # No matching events found for this user
        $lastLogin = "No recent login events found."
    }

    # Get password expiry date
    $userObj = Get-LocalUser -Name $username
    $passwordExpiry = $userObj.PasswordExpires

    # Append user information to the HTML content
    $htmlReport += @"
    <tr>
        <td>$username</td>
        <td>$lastLogin</td>
        <td>$passwordExpiry</td>
    </tr>
"@
}
$htmlReport += @"
    </table>
"@

##############################################

########System Environment Variable###########
Write-Log "HTML report generating for : System Environment Variable"
Write-Host "HTML report generating for : System Environment Variable"

# Function to generate the system environment information as an HTML table
function Generate-EnvironmentTable {
    param (
        [hashtable]$Data
    )

    $content = @"
    <table align="center" style="table-layout: fixed;WORD-BREAK:BREAK-ALL;margin:10px"><tr><td colspan="2" align="center"><b>System Environment Variable</b></td></tr>
        <tr>
            <td><b>Environment Variable</b></td>
            <td><b>Value</b></td>
        </tr>
"@

    foreach ($entry in $Data.GetEnumerator()) {
        $content += @"
        <tr>
            <td>$($entry.Key)</td>
            <td>$($entry.Value)</td>
        </tr>
"@
    }

    $content += @"
    </table>
"@

    return $content
}

# Get system environment information
$environmentVariables = Get-ChildItem Env:

# Convert environment variables to a hashtable
$environment = @{}
foreach ($var in $environmentVariables) {
    $environment[$var.Name] = $var.Value
}

#$htmlReport += @"
    #<table align="center" style="width: 900px;WORD-BREAK:BREAK-ALL"><tr><td colspan="3" align="center"><b>System Environment Variable</b></td></tr>>
    
#"@

# Append the system environment information to the existing HTML report template
$htmlReport += Generate-EnvironmentTable -Data $environment

$htmlReport += @"
    </table>
"@

##############################################

############Running Processes#################
Write-Log "HTML report generating for : Running Processes"
Write-Host "HTML report generating for : Running Processes"
# Get the list of running processes
$processes = Get-Process | Select-Object Id, Name, CPU, Memory, Path, Description

$htmlReport += @"
 
    <table align="center" style="table-layout: fixed;WORD-BREAK:BREAK-ALL;margin:10px"><tr><td colspan="3" align="center"><b>Running Processes</b></td></tr>
        <tr>
            <td><b>PID</b></td>
            <td><b>Name</b></td>
            <td><b>Path</b></td>
        </tr>
"@

foreach ($process in $processes) {
    $htmlReport += @"
        <tr>
            <td>$($process.Id)</td>
            <td>$($process.Name)</td>
            <td>$($process.Path)</td>
        </tr>
"@
}
$htmlReport += @"
    </table>
"@
##############################################

##########Open Network Connection#############
Write-Log "HTML report generating for : Open Network Connection"
Write-Host "HTML report generating for : Open Network Connection"
function Get-NetworkStatistics
{
    $properties = 'PID', 'Process Name', 'Protocol', 'Local Address', 'Local Port', 'Remote Address', 'Remote Port', 'State'
	
    netstat -ano | Select-String -Pattern '\s+(TCP|UDP)' | ForEach-Object {
        $item = $_.Line.Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)

        if ($item[1] -notmatch '^\[::')
        {
            if (($la = $item[1] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6')
            {
                $localAddress = $la.IPAddressToString
                $localPort = $item[1].Split('\]:')[-1]
            }
            else
            {
                $localAddress = $item[1].Split(':')[0]
                $localPort = $item[1].Split(':')[-1]
            }

            if (($ra = $item[2] -as [ipaddress]).AddressFamily -eq 'InterNetworkV6')
            {
                $remoteAddress = $ra.IPAddressToString
                $remotePort = $item[2].Split('\]:')[-1]
            }
            else
            {
                $remoteAddress = $item[2].Split(':')[0]
                $remotePort = $item[2].Split(':')[-1]
            }

            $processId = $item[-1]
            $processName = (Get-Process -Id $processId -ErrorAction SilentlyContinue).Name

            $connectionInfo = @{
                'PID'           = $processId
				'Process Name'   = $processName	
				'Protocol'      = $item[0]
                'Local Address'  = $localAddress
                'Local Port'     = $localPort
                'Remote Address' = $remoteAddress
                'Remote Port'    = $remotePort
                'State'         = if ($item[0] -eq 'tcp') { $item[3] } else { $null }		

            }

            New-Object PSObject -Property $connectionInfo | Select-Object -Property $properties
        }
    }
}

$htmlReport += @"
    <table align="center" style="table-layout: fixed;WORD-BREAK:BREAK-ALL"><tr><td colspan="8" align="center"><b>Open Network Connection</b></td></tr>

"@

# Get network statistics and convert to HTML table
$networkStatistics = Get-NetworkStatistics
$htmlTable = $networkStatistics | ConvertTo-Html -Property $networkStatistics[0].PSObject.Properties.Name

# Append the network statistics HTML table to the report
$htmlReport += $htmlTable

$htmlReport += @"
    </table>
"@
##############################################

##############Services Information############
Write-Log "HTML report generating for : Services Information"
Write-Host "HTML report generating for : Services Information"
$servicesInfo = Get-Service | Select-Object Name, DisplayName, Status

$htmlReport += @"
    <table align="center" style="table-layout: fixed;WORD-BREAK:BREAK-ALL;margin:10px"><tr><td colspan="3" align="center"><b>Services Information</b></td></tr>
        <tr><td><b>Name</b></td><td><b>Display Name</b></td><td><b>Status</b></td></tr>
"@
$servicesInfo | ForEach-Object {
    $serviceName = $_.Name
    $displayName = $_.DisplayName
    $status = $_.Status
    $htmlReport += "<tr><td>$serviceName</td><td>$displayName</td><td>$status</td></tr>"
}

$htmlReport += @"
    </table>
"@

##############################################

#############Installed Application############
Write-Log "HTML report generating for : Installed Application"
Write-Host "HTML report generating for : Installed Application"
$installedApps = Get-WmiObject Win32_Product | Select-Object Name, Version

$htmlReport += @"
    <table align="center" style="table-layout: fixed;WORD-BREAK:BREAK-ALL;margin:10px"><tr><td colspan="2" align="center"><b>Installed Applications</b></td></tr>
        <tr><td><b>Name</b></td><td><b>Version</b></td></tr>
"@
$installedApps | ForEach-Object {
    $appName = $_.Name
    $appVersion = $_.Version
    $htmlReport += "<tr><td>$appName</td><td>$appVersion</td></tr>"
}

$htmlReport += @"
    </table>
"@

##############################################

###########Generating HTML Report#############

$htmlReport += @"
</body>
</html>
"@

# Save the HTML report to a file
$reportPath = "systeminfo-html-report_$computerName.html"
$htmlReport | Out-File -FilePath $reportPath -Encoding UTF8

Write-Host "HTML report saved to: $reportPath"

Write-Log "HTML report saved to: $reportPath"
Write-Host "HTML report saved to: $reportPath"




# Check if XAMPP service is running
$xamppService = $services | Where-Object { $_.Name -eq "XAMPP" -and $_.Status -eq "Running" }

# Check if IIS service is running
$iisService = $services | Where-Object { $_.Name -eq "W3SVC" -and $_.Status -eq "Running" }

$SqlServerServiceName = "MSSQLSERVER"  # Add more service names as needed

# Specify the path to the HTML file
$htmlFilePath = ".\systeminfo-html-report_$computerName.html"
$html = Get-Content -Path $htmlFilePath -Raw

  $iisStatus = ""
  $xamppService = ""
  $databaseServerRunning = ""
  $LocaldatabaseName = "Sql Server"
  $xamppServiceNew = "NO"

# Function to get the general server details
function GetServerGeneralInfo($sysid) {
    try {
        
            Write-Log "Getting General information of server"
            $sqlServerService = Get-Service -Name $SqlServerServiceName -ErrorAction SilentlyContinue

            if ($sqlServerService -eq $null) {
                $databaseServerRunning = "Not Installed"
                #Write-Host "SQL Server is not installed or the service name is incorrect."
                }
            elseif ($sqlServerService.Status -eq "Running") {
                $databaseServerRunning = "Running"
                #Write-Host "SQL Server is running."
                }
           else {
                    $databaseServerRunning = "Not Running"
                    #Write-Host "SQL Server is installed but not running."
                }

        # Check if XAMPP service is running
        $xamppService = $services | Where-Object { $_.Name -eq "XAMPP" -and $_.Status -eq "Running" }
        if ($xamppService)
        {
            $xamppServiceNew = "YES"
        }
        # Get IIS status
        $iisStatus = (Get-Service -Name W3SVC).Status

        # Insert command for ServerGenralInfo table
$InsertGeneralSystemdataSql = @"
                INSERT INTO ServerGenralInfo (sys_id, IsDatabaseServerAvailable, DatabaseName, IsIISActive, IsXamppActive)
                VALUES ('$sysid', '$databaseServerRunning', '$LocaldatabaseName', '$iisStatus', '$xamppServiceNew' );
"@
        # Insert data to table
        InsertData $InsertGeneralSystemdataSql
    }
    catch {
        Write-Log "Error While getting Server GenralInfo"
        Write-Log $_
        Exit 1
    }
}

# Database name
$databaseName = "SectionInformation_db"

# Connection string for the Local SQL Server
$connectionString = "Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=$databaseName;Integrated Security=True;Connect Timeout=30;Encrypt=False;Application Name=FinalHTMLTODB"


# Connection string for the Azure SQL Server
#$connectionString = "Server=tcp:getintoit1dbserver.database.windows.net,1433;Initial Catalog=SectionInformation_db;Persist Security Info=False;User ID=AshwaniGusain;Password=*****;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;"

# Create a new connection for database operations
try {
    $useDatabaseConnection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $useDatabaseConnection.Open()
    Write-Log "Connected to the database."
}
catch {
    Write-Log "Error connecting to the database: $_"
    Exit 1
}

# Function to execute a SQL query and return a single value
function GetSingleValueFromQuery($query) {
    try {
        $command = $useDatabaseConnection.CreateCommand()
        $command.CommandText = $query
        $result = $command.ExecuteScalar()
        $command.Dispose()
        return $result
    }
    catch {
        Write-Log "Error executing query: $query"
        Write-Log $_
        Exit 1
    }
}

# Function to execute a SQL insert query and return the last inserted ID
function InsertAndGetLastID($query) {
    try {
        #Write-Host "$query"
        # Start a new transaction
        $transaction = $useDatabaseConnection.BeginTransaction()

        $command = $useDatabaseConnection.CreateCommand()
        $command.Transaction = $transaction
        $command.CommandText = $query
        $command.ExecuteNonQuery()
        $command.Dispose()

        # Fetch the last inserted ID using SCOPE_IDENTITY()
        $identityQuery = "SELECT SCOPE_IDENTITY() AS LastInsertedID"
        $identityCommand = $useDatabaseConnection.CreateCommand()
        $identityCommand.Transaction = $transaction
        $identityCommand.CommandText = $identityQuery
        $lastInsertedID = $identityCommand.ExecuteScalar()
        $identityCommand.Dispose()

        # Commit the transaction
        $transaction.Commit()

        #Write-Host $lastInsertedID
        return $lastInsertedID
    } catch {
        # Handle any exceptions
        Write-Host "Error: $_"
        # Rollback the transaction in case of an error
        $transaction.Rollback()
        return $null
    } finally {
        # Always dispose of the transaction
        $transaction.Dispose()
    }
}



# Function to insert data into the database
function InsertData($query) {
    try {
        $command = $useDatabaseConnection.CreateCommand()
        $command.CommandText = $query
        $command.ExecuteNonQuery()
        $command.Dispose()
        }
    catch {
        Write-Log "Error Inserting the data: $query"
        Write-Log $_
        Exit 1
    }
}

# Create an HTML document object
$htmlDocument = New-Object HtmlAgilityPack.HtmlDocument
$htmlDocument.LoadHtml($html)

$tableHeadingHTML = "System Information"
#$tableUserId = "System Information"

# Fetch desiredUserDetails and desiredIPValue
$desiredUserDetails = ""
$desiredIPValue = ""


$tables = $htmlDocument.DocumentNode.SelectNodes("//table[@style='table-layout: fixed;WORD-BREAK:BREAK-ALL;margin:10px']")
Write-Host "$tables"
foreach ($table in $tables) {
    # Extract the table heading
    $tableHeading = $table.SelectSingleNode("tr/td").InnerText
    Write-Host "Hi I'm here"
    if ($tableHeading -eq $tableHeadingHTML) {
        $tableRows = $table.SelectNodes("tr[position() > 1]")

        foreach ($row in $tableRows[1..($tableRows.Count - 1)]) {
            $columnValues = $row.SelectNodes("td") | ForEach-Object { $_.InnerText }
            Write-Host "$columnValues"
            if ($columnValues -like "User Name") {
                $desiredUserDetails = $columnValues[1]
                
            }
            if ($columnValues -like "IP Address") {
                $desiredIPValue = $columnValues[1]
                break
            }
        }

    }
}

$currentTime = Get-Date
$dateTimeObject = [DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss")

# Fetch the latest iteration number for the specified IP address
$sysiteration = GetSingleValueFromQuery("SELECT ISNULL(MAX(IterationNumber), 0) FROM systems WHERE sys_IP LIKE '%$desiredIPValue%'")

# Increment the iteration number
$sysiteration++

# Insert or update data in the systems table
$InsertSysIpSql = @"
    IF NOT EXISTS (SELECT 1 FROM systems WHERE sys_IP = '$desiredIPValue' AND IterationNumber = '$sysiteration')
    BEGIN
        INSERT INTO systems (sys_IP, sys_name, IterationNumber, sys_time)
        VALUES ('$desiredIPValue', '$desiredUserDetails', '$sysiteration', '$dateTimeObject')
    END
    ELSE
    BEGIN
        UPDATE systems
        SET sys_name = '$desiredUserDetails', sys_time = '$dateTimeObject'
        WHERE sys_IP = '$desiredIPValue' AND IterationNumber = '$sysiteration'
    END
"@
InsertData $InsertSysIpSql

# Fetch the sysid from systems table
$sysid = GetSingleValueFromQuery("SELECT Id FROM systems WHERE sys_IP = '$desiredIPValue' AND IterationNumber = '$sysiteration'")


# call to get initial system information
GetServerGeneralInfo $sysid


# Fetch the sectioninfo IDs and sys_headings from systemsInformation table
$systemsInformationQuery = "SELECT Id, sys_heading FROM sectioninfo"
$systemsInformationCommand = $useDatabaseConnection.CreateCommand()
$systemsInformationCommand.CommandText = $systemsInformationQuery
$systemsInformationReader = $systemsInformationCommand.ExecuteReader()

Write-Log "This is IterationNumber : $sysiteration processing on  : $desiredIPValue"
# Create a DataTable to store the results from the SQL query
$dataTable = New-Object System.Data.DataTable

# Close the DataReader before filling the DataTable
$systemsInformationReader.Close()
$systemsInformationReader.Dispose()

# Use a SqlDataAdapter to fill the DataTable with data from the SQL query
$adapter = New-Object System.Data.SqlClient.SqlDataAdapter
$adapter.SelectCommand = $systemsInformationCommand
$adapter.Fill($dataTable)

$systemsInformationCommand.Dispose()

foreach ($row in $dataTable.Rows) {
    $id = $row["Id"]
    $sysHeading = $row["sys_heading"]

    foreach ($table in $tables) {
        # Extract the table heading
        $tableHeading = $table.SelectSingleNode("tr/td").InnerText
        $trimmedHeading = ($tableHeading -split '-')[0].Trim()
        #Write-Host "$trimmedHeading" 
        # Compare the sys_heading value with the table heading
        if ($sysHeading -eq $trimmedHeading) {
            Write-Host "Table heading matches sys_heading in sectioninfo table."
            $tableRows = $table.SelectNodes("tr[position() > 1]")

            $headerCells = $table.SelectSingleNode(".//tr[2]").SelectNodes("td[b]")


            # Extract and output the text of each header cell
            $columnNames = $headerCells | ForEach-Object { $_.InnerText }


              #Write-Host "$columnNames"
             # Exit 1
            # Extract column names from the first row
            #$columnNames = $tableRows[1].SelectNodes("td") | ForEach-Object { $_.InnerText }
                 $wordCount = $columnNames.Split(' ').Count
                 
            for ($i = 0; $i -lt $columnNames.Length; $i++) {

            #Write-Host "$wordCount"

        if ($wordCount -gt 1) {
                    
                    $coln = $columnNames[$i]
                }
                else {
        $coln = $columnNames
        }
            #$columnIndex = $headerCells.FindIndex({ $_.InnerText -eq $coln })
            # Insert column names into sectioninfo_columns table
            $columnsSql = @"
                INSERT INTO sectioninfo_columns (sys_id, section_id, column_name, IterationNumber)
                VALUES ('$sysid', '$id', '$coln', '$sysiteration');
"@
         #write-host "$columnsSql" 
        $lastInsertedColumnID = InsertAndGetLastID $columnsSql
        $lastInsertedColumnID_trim = ($lastInsertedColumnID -split '-')[1].Trim()
        #write-host "$lastInsertedColumnID_trim" 
        #Exit 1
        
        # Iterate through each row (excluding the header row)
    $rows = $table.SelectNodes(".//tr[position() > 2]")
    foreach ($row in $rows) {
        $cells = $row.SelectNodes("td")
        #Write-Host $cells[0].InnerText
        # Iterate through each cell (column) and extract the value
            #$columnName = $columnNames[$i]
            if ( $coln -eq $columnNames[$i])
            {
            $columnValue = $cells[$i].InnerText

    # Generate the insert SQL statement
    $insertDataSql = @"
        INSERT INTO sectioninfo_rowsdata (system_id, sectioninfo_id, Column_id, rowdata, IterationNumber)
        VALUES ('$sysid', '$id', '$lastInsertedColumnID_trim', '$columnValue', '$sysiteration');
"@
    InsertData $insertDataSql
    }
    }
   }
 }
}
}

$systemsInformationReader.Close()
$systemsInformationReader.Dispose()
$systemsInformationCommand.Dispose()

# Close and dispose of the database connection
$useDatabaseConnection.Close()
$useDatabaseConnection.Dispose()
Write-Log "Log Ended :Script completed successfully."
}
catch {
    Write-Log "Error in the script: $_"
    Exit 1
}


finally {
    # Close and dispose of the database connection
    if ($useDatabaseConnection -ne $null) {
        if ($useDatabaseConnection.State -eq 'Open') {
            try {
                $useDatabaseConnection.Close()
                Write-Log "May be user cancelled the operation."
            }
            catch {
                Write-Log "Error closing the database connection: $_"
            }
        }
        $useDatabaseConnection.Dispose()
    }
}