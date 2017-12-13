##Decalring Worker Functions
function LogData{
    #Write to specified Log file.
    if(!(Test-Path "Log.txt")){
        ##If file does not exist, create it on the specified backup drive
        $Text = "Backup Creation Logs for $env:USERNAME on $env:COMPUTERNAME"
        $Text | Out-File "Log.txt"
        Get-WmiObject -Class Win32_ComputerSystem | Out-File "Log.txt"
    }    
    $TimeStamp = Get-Date
    $LogLine = "$TimeStamp `t $args"
    $LogLine | Out-File "Log.txt" -Append
}
function TestProcess{
    #Test if a process is open and prompt user to close it.
    param( [string] $Name )
    if(get-process -ErrorAction SilentlyContinue $Name | where {$_.ProcessName -eq “$Name”}){
        LogData "$StartDate `t $Name is Open"
        Write-Host "$Name is Open, Please close $Name before continuing." -ForegroundColor Red
        write-Host "Failure to close the program listed above may result in corrupt or missing data."
        Read-Host 'Press ENTER to continue ' | Out-Null
        if(get-process -ErrorAction SilentlyContinue $Name | where {$_.ProcessName -eq “$Name”}){
            #Retest to verify the use closed the application.
            Logdata "• ï‹‰User did not close $Name before continuing!"
        } else {
            Logdata "• User closed $Name as prompted."
        }
    } else {
        Write-Host "$Name is Closed" -ForegroundColor Green
    }
}
function SendEmailNotification{ 
    $Timestamp = Get-Date
    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItem(0)
    $Mail.To = "helpdesk@atlantaforklifts.com"
    $Mail.Subject = "PS Backup Notification"
    $Mail.HTMLBody ="
    <h1>::This is an automatically generated Held Desk Ticket::</h1>


    <p>The following error has occurred on $Timestamp</P>
    <p><b> $args </b></p>

    <p>USER: <b> $env:Username </b><br>
    DEVICE: <b>$env:COMPUTERNAME </b><br></p>   

    <p>Folders included: <b> $Dirs </b><br><br>
    Network Connections <br> $MYConnections</p>
"

    $mail.save()
    $inspector = $mail.GetInspector
    $inspector.Display()
}

# Import Assemblies
[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")

# Setup variables
$StartDate = Get-Date
$StartTime = $StartDate.ToShortTimeString()
Write-Host "Backup start: $StartDate"
LogData "***** Backup started *****"

### User profile directories to be backed up. ###
$Dirs = @("Desktop","Pictures","Favorites","Documents","Downloads")

### Applications that need to be verified closed before continuing. ###
$Applications = @('Outlook', 'Excel', 'Word')

###
$SendNotificationTo = "jshofstahl@atlantaforklifts.com"

$MYConnections = Get-NetIPAddress | Select-Object ifIndex, IPAddress, PrefixLength, PrefixOrigin, SuffixOrigin, AddressState | Where-Object {$_.SuffixOrigin -eq "Dhcp" -or $_.SuffixOrigin -eq "Manual" } |  ConvertTo-HTML -head $a | Out-String

$DeviceLabel = 'BackupDrive'
$BKUPDrive = Get-WMIObject Win32_Volume | ? { $_.Label -eq $DeviceLabel } 
$BKUP = Get-WMIObject Win32_Volume | ? { $_.Label -eq $DeviceLabel } | %{$_.deviceid}
$BKUPLetter = Get-WMIObject Win32_Volume | ? { $_.Label -eq $DeviceLabel } | %{$_.DriveLetter}
$yncFolder = "TFA_SyncFolder"


# Loop through target applications to make sure they are closed.
foreach($App in $Applications){
    TestProcess -Name $App
}
# Sleep to display results
Start-Sleep -s 3

if(!($BKUP)) {
    #This error will only show if the volume name is not set poperly -or- not inserted
    LogData "ERROR: Backupdrive Not Found"
    [void][System.Windows.Forms.MessageBox]::Show("                             BKUP DRIVE NOT FOUND!
    If you are receiving this error please contact your IT Department.")
    SendEmailNotification "<br>ERROR: <b><font color=red>Testing Email Notification.</b></font> "

} else {
    Write-Host  "BKUP Drive Found!"
    $bytes = Get-WMIObject Win32_Volume | ? { $_.Label -eq $DeviceLabel } | %{$_.Capacity}
    [int]$Capacity = $bytes / 1GB
    Write-Host "Your Backup Drive has a $Capacity GB Capacity"
    
    Write-Host "Building Source Directory Array"
    $TotalSize = 0;    
    
    $DirTableName = "Directories"
    $DirTable = New-Object System.Data.DataTable "$DirTableName"
    $col1 = New-Object system.Data.DataColumn Path,([string])
    $col2 = New-Object system.Data.DataColumn Size,([string])
    $col3 = New-Object system.Data.DataColumn Action,([string])
    $DirTable.Columns.Add($col1)
    $DirTable.Columns.Add($col2)
    $DirTable.Columns.Add($col3)




    foreach ($Dir in $Dirs) {
       $DirTableRow = $DirTable.NewRow()
       $DirTableRow.Path = "$ENV:USERPROFILE\$Dir\"
       
       [int]$Objects = Get-ChildItem -Path "$ENV:USERPROFILE\$dir\" -Recurse | Measure-Object | %{$_.Count}
       
       if($Objects -eq 0) {
          Write-Host  "$ENV:USERPROFILE\$Dir\`t`t has $Objects Objects **SKIPPING DIRECTORY**" -ForegroundColor DarkRed
          LogData "NOTICE: $ENV:USERPROFILE\$Dir\`t`t has $Objects Objects. Skipping Null folder."
          $DirTableRow.Size = "Null"
          $DirTableRow.Action = "Skipping"
       } else {
          $Size = Get-ChildItem -Recurse "$ENV:USERPROFILE\$Dir" | Measure-Object -property length -sum
          $TotalSize += $Size.sum
          [int] $SizeinKB = $Size.sum /1KB
          [int] $SizeinMB = $Size.sum / 1MB
          [int] $SizeinGB = $Size.sum / 1GB
          
          if($SizeinGB -gt 1) {
            #Write-Host  "$ENV:USERPROFILE\$Dir\`t`t $SizeinGB GB " -ForegroundColor Green
            $DirSize = "$SizeinGB`tGB"
          } elseif($SizeinMB -gt 1) {
            #Write-Host  "$ENV:USERPROFILE\$Dir\`t`t $SizeinMB MB" -ForegroundColor Green
            $DirSize = "$SizeinMB`tMB"
          } else {
            #Write-Host  "$ENV:USERPROFILE\$Dir\`t`t $SizeinBytes Bytes" -ForegroundColor Green
            $DirSize = "$SizeinKB`tKB"
          }
          $DirTableRow.Size = $DirSize
          $DirTableRow.Action = "Backup"
       }
       $DirTable.Rows.Add($DirTableRow)
    }

    $DirTable | Format-Table
    [Int]$TotalSizeinGB = $TotalSize / 1GB
    [int]$Percent =  (($TotalSizeinGB / $Capacity) * 100 )
    if($Percent -gt 100){
        # Verify there is enough space on the disk.
        write-host "Your documents exceeds the available  disk space. Please contat the Toyota Forklifts of Atlanta IT Department." -ForegroundColor Red
        LogData "ERROR: User Profile exceeds the avalible disk space. BACKUP TERMINATED"
        SendEmailNotification "<br>ERROR: <b><font color=red>User Profile Exceeds the disk space available  on drive.</b></font> <br><br> Backup Size: $TotalSizeinGB GB <br> Backup Drive Capacity $Capacity GB <br>"
        write-host "A HelpDesk ticket has been automatically generated for this error. Pelase end it from the outlook window." -ForegroundColor DarkRed
        Read-Host 'Press ENTER to continue' | Out-Null

    } else {
        write-host "Total Size = $TotalSizeinGB GB    [$Percent % of Disk Capacity]" -ForegroundColor Green
        #Discover or Create Target Directory
        $PathExist = Test-Path "$BKUPLetter\$yncFolder"
        if($PathExist){
            Write-Host "$BKUPLetter\$yncFolder folder FOUND..."
        } else {
            Write-Host "$BKUPLetter\$yncFolder folder NOT FOUND!"
            new-item "$BKUPLetter\$yncFolder" -itemtype directory
            Write-Host "$BKUPLetter\$yncFolder has been created!" -ForegroundColor DarkGreen
        }
        $PreviousSize = Get-ChildItem -Recurse "$BKUPLetter\$yncFolder" | Measure-Object -property length -sum

        ###Start Robocopy
        $MBCompleted = 0;
        foreach($Dir in $Dirs){
            Robocopy "$ENV:USERPROFILE\$Dir" "$BKUPLetter\$yncFolder\$Dir" /MIR /NJH /R:1 /W:1 /NP /NC /TEE /Log+:"$BKUPLetter\$yncFolder\_SyncLog-$Dir.txt"
        }

    
        ## Determain remaining space on disk.
        [int] $PSize = $PSize.sum / 1GB
        If($Percent -lt 50){
            Write-Host -NoNewline Total User Profile $TotalSizeinGB GB. -ForegroundColor DarkGreen
            Write-Host " Your drive is at $Percent% capacity." -ForegroundColor DarkGreen
        } elseif($Percent -lt 75) {
            Write-Host -NoNewline Total User Profile $TotalSizeinGB GB. -ForegroundColor DarkGreen
            Write-Host " Your drive is at $Percent% capacity." -ForegroundColor DarkYellow
        } elseif($Percent -lt 85) {
            Write-Host -NoNewline Total User Profile $TotalSizeinGB GB. -ForegroundColor DarkGreen
            Write-Host " Your drive is at $Percent% capacity." -ForegroundColor DarkMagenta
        } elseif($Percent -lt 95) {
            Write-Host -NoNewline Total User Profile $TotalSizeinGB GB. -ForegroundColor DarkRed
            Write-Host " Your drive is at $Percent% capacity." -ForegroundColor DarkRed        
        } else {
            Write-Host -NoNewline Total User Profile $TotalSizeinGB GB. -ForegroundColor Red
            Write-Host " Your drive is at $Percent% capacity." -ForegroundColor Red
            Write-Host "" 
            Write-Host Contact the Atlana Forklifts Inc. IT Department for an increased capacity drive. -ForegroundColor Red
        }
        $ReminingSpace = $Capacity - $TotalSizeinGB
        if($ReminingSpace -lt 1){
            SendEmailNotification "<br><b><font color=red>NOTICE:</font></b>Users BackupDrive is at <b><font color=red>$Percent% capacity, and has less than 1GB remaining.</b></font><br><br>Backup Size: $TotalSizeinGB GB<br>Backup Drive Capacity $Capacity GB<br>"
        }
   
    
        #Set End time
        $EndDate = Get-Date
        $EndTime = $EndDate.ToShortTimeString()
        Write-Host "Completed on $EndDate" -ForegroundColor DarkGreen
        LogData "----- Backup Complete ----- `n"

        $TS = NEW-TIMESPAN –Start $StartDate –End $EndDate
        $TShours = $TS.Hours
        $TSmin = $TS.Minutes
        $TSsec = $TS.Seconds
    
      
        Write-Host Time Elapsed: $TShours hours $TSmin minutes $TSsec seconds -ForegroundColor DarkGreen 
        
    }
}
Read-Host 'Press ENTER to Exit' | Out-Null

