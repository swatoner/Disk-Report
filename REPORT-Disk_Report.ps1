<#
.SYNOPSIS
    Check disk space and send an HTML report with email.  
.DESCRIPTION
    Generates a report for disks of Servers or Computers.
    Additional filters can be applied. 
.NOTES
    Author         : Nick Menegatos 
    Requires       : PowerShell V5.1 
    Date Created   : 2019-04-06 18:56:03 
    Script Version : 2.1  
.EXAMPLE
    
#>

if (get-module -listavailable -Name 'ActiveDirectory' ) {
    import-module activedirectory
}
else {
try {
    Import-Module ServerManager
    Add-WindowsFeature RSAT-AD-PowerShell
    import-module activedirectory
    }
    catch { Write-Output "Could not load Active Directory module"; break }
}

if ($psise) {$runningpath = Split-Path $psise.CurrentFile.FullPath}
else {$runningpath = $pwd | Select-Object -ExpandProperty Path}


<#
if (!(Find-Module -Name VMware.PowerCLI)) {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-Module -Name VMware.PowerCLI -Scope CurrentUser -Confirm:$false
}
else {
    # Error out if loading fails  
    Write-Error "`nERROR: Cannot load the VMware Snapin or Module. Is the PowerCLI installed?"  
}  #>

    Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -Confirm:$false
    Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false
    
# Logging Process
$log_file = "$runningpath\Disk_Report.log"
if (Test-Path $log_file) { }
else { New-Item $log_file -Force }
$log_date = Get-Date
Write-Output "Config Backup Schedule initiated on $log_date" | Out-File $log_file


# REPORT PROPERTIES
    # Report Status
        $report = 'Full' ### Full / Warning
        $filter_ora = 'Yes' ### Yes / No   # Select Yes if you want to filter the Data Disks
   	# Send report via e-mail or no?
        $sendemail = 'Yes' ### Yes / No
        # When it should send the e-mail report.
        $email_report = 'Warning' ### Full / Warning
	# Path for the report file
        $reportPath = "C:\IT\Scripts\Reports\"
    # Report name
	    $reportName = "DiskSpaceReport_$(get-date -format yyyyddMMHHmm).html";
    # Path and Report name together
        $Report_Path = $reportPath + $reportName
    # Set your warning and critical thresholds
        $percentWarning = 10;
        $percentCritcal = 5;

# Continue even if there are errors
    #$ErrorActionPreference = "Continue";

#ENVIRONMENT Parameters    
    $Domain = get-addomain
    $Domain_Name = $domain.Name
    $Domain_DNS = $domain.DNSRoot
    $DC = (Get-ADDomainController).Name
    $title = "Environment DiskSpace Report"
    #$servers_list = invoke-command -ComputerName $dc -scriptblock { $domain_dname = (Get-ADDomain).DistinguishedName; (get-adcomputer -filter {(Enabled -eq $true)} -SearchBase "OU=Servers,OU=SERVICES,$domain_dname").Name } | Sort-Object
    $domain_dname = (Get-ADDomain).DistinguishedName
    $servers_list = (Get-ADComputer -filter {(Enabled -eq $true)} -SearchBase "OU=Servers,$domain_dname").Name | Sort-Object
    #$runningpath = "C:\IT\Scripts"
    $Destination_Path = "C:\IT\Scripts\Disk_Report\"
    $config_file = "\\$Domain_DNS\NETLOGON\Tools\GUConfig.xml"

#vCenter Parameters
    if (($Env:ComputerName -like "GLVS*") -or ($Env:ComputerName -like "GPVS*")) {
        [array]$vCenters = "vcenter.gunion.gr", "vcenter-gl.gunion.gr"
    }
    else { 
        $vCenters = "vcenter." + $Domain_DNS
    }

# CONFIGURATION FILE
    $required_config_version  = "0.1"
    if (!(Test-Path $config_file)) { Write-Output "Configuration file $config_file not found. EXITING"; Pause; Break }
    elseif (Test-Path $config_file) { [xml]$Config = Get-Content -path $config_file }
    # Check Compatibility with Configuration File
    if ($config.CONFIG.VERSION -lt $required_config_version) { Write-Host "The configuration file is older and not supported."; Break }
# Check folders
    if (!(Test-Path $ReportPath)) { New-Item -Path $ReportPath -ItemType Directory }
    if (!(Test-Path $Destination_Path)) { 
        try { New-Item -Path $Destination_Path -ItemType Directory }
        catch { Write-Host "Could not create folder"; exit }
    }

# EMAIL PROPERTIES
    # Set the recipients of the report. You can set multiple e-mails
    $email = $config.CONFIG.EMAIL
    $email_users = $email.MAILTO
    $email_from = $email.SENDER
    $email_smtp = $email.SERVER
    $email_login = $email.USERNAME
    $email_port = $email.PORT
    $email_password = $email.PASSWORD
    $email_password = $email_password | ConvertTo-SecureString -AsPlainText -Force
    $email_credentials = New-Object System.Management.Automation.Pscredential -Argumentlist $email_login, $email_password

# Prerequeseties for reports
    if ($report -ne 'Full' -and $report -ne 'Warning') { 
        $message = 'Incorrert parameters set for value $Report'
        Write-Host $message
#        if ($debug -eq 1) { Break } else { AIMS-ErrorLog $message; Exit } 
    }
    if ($sendemail -ne 'Yes' -and $sendemail -ne 'No') {
        $message = 'Incorrert parameters set for value $sendemail'
        Write-Host $message
#         #New-EventLog -LogName AIMS -Source IT -MessageResourceFile "$script $title"
#        if ($debug -eq 1) { Break } else { AIMS-ErrorLog $message; Exit }
    }
    if (!(Test-Path $servers_list) -or $servers_list -eq $null) {
        $message = "Cannot get list of Computers"
        Write-Host $message
#        if ($debug -eq 1) { Break } else { AIMS-ErrorLog $message; Exit }  
    }

# Count if any computers have low disk space.  Do not send report if less than 1.
    $global:i = 0;

# Get computer list to check disk space
    if ($debug -eq 1) { 
        if ($servers_list -is [array]) { $computers = $servers_list }
        #else {$computer = $servers_list }
    }
    elseif ($servers_list -is [array]) { $computers = $servers_list }
    else { $computers = Get-Content $servers_list }
#Date
    #$datetime = Get-Date -Format "yyyy-MM-dd_HHmmss";
# Test if the report path exists or else create it
    if (!(Test-Path $ReportPath)) { New-Item -Path $ReportPath -ItemType Directory }
# Remove the report if it has already been run today so it does not append to the existing report
    If (Test-Path $Report_Path) { Remove-Item $Report_Path }
# Cleanup old files..
    $Daysback = "-365"
    $CurrentDate = Get-Date;
    $DateToDelete = $CurrentDate.AddDays($Daysback);
    Get-ChildItem $Destination_Path | Where-Object { $_.LastWriteTime -lt $DatetoDelete } | Remove-Item -Force;
# HTML
    #Set colors for table cell backgrounds
    $redColor = "#FF0000"
    $orangeColor = "#FBB917"
    $whiteColor = "#FFFFFF"
    $blackColor = "000000"
    $greyColor = "808080"
    $titleDate = get-date -uformat "%d-%m-%Y @ %R"
    $HTMLtitle = "$title for $titledate"
    $HTML_Header = $null
    $HTML_tableHeader = $null
    $HTML_dataRow = $null
    $HTML_tableDescription = $null
    $HTML_Body = $null
    
# HTML Header
    $HTML_Header = "
		<html>
		<head>
		<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
		<title>DiskSpace Report</title>
		<STYLE TYPE='text/css'>
		<!--
		td {
			font-family: Tahoma;
			font-size: 11px;
			border-top: 1px solid #999999;
			border-right: 1px solid #999999;
			border-bottom: 1px solid #999999;
			border-left: 1px solid #999999;
			padding-top: 0px;
			padding-right: 0px;
			padding-bottom: 0px;
			padding-left: 0px;
		}
		body {
			margin-left: 5px;
			margin-top: 5px;
			margin-right: 0px;
			margin-bottom: 10px;
			table {
			border: thin solid #000000;
		}
		-->
		</style>
		</head>
		<body>
		<table width='100%'>
		<tr bgcolor='#CCCCCC'>
		<td colspan='7' height='25' align='center'>
		<font face='tahoma' color='#003399' size='4'><strong>$HTMLtitle</strong></font>
		</td>
		</tr>
		</table>"

    # Create and write Table header for report
        $HTML_tableHeader = "
         <table width='100%'><tbody>
	        <tr bgcolor=#CCCCCC>
            <td width='10%' align='center'>Server</td>
	        <td width='5%' align='center'>Drive</td>
	        <td width='15%' align='center'>Drive Label</td>
	        <td width='10%' align='center'>Total Capacity(GB)</td>
	        <td width='10%' align='center'>Used Capacity(GB)</td>
	        <td width='10%' align='center'>Free Space(GB)</td>
	        <td width='5%' align='center'>Freespace %</td>
	        </tr>"

# PROCESS
    # Start processing disk space reports against a list of servers
        foreach($computer in $computers) {
            if (Test-Connection $computer -Count 1 -Quiet) { 
                $disks = Get-WmiObject -ComputerName $computer -Class Win32_LogicalDisk -Filter "DriveType = 3"
                if ($filter_ora -eq 'Yes') { 
                    $disks = $disks | Where-Object {$_.VolumeName -notlike "ORA_Data*" -and $_.VolumeName -notlike "ORAData*"}
                }
                #$disks = $disks | Where-Object {$_.VolumeName -notlike "Data*"}
	            #$computer = $computer.toupper()
                foreach($disk in $disks){
                    $deviceID = $disk.DeviceID;
                    $volName = $disk.VolumeName;
	                [float]$size = $disk.Size;
		            [float]$freespace = $disk.FreeSpace; 
		            $percentFree = [Math]::Round(($freespace / $size) * 100, 2);
		            $sizeGB = [Math]::Round($size / 1073741824, 2);
		            $freeSpaceGB = [Math]::Round($freespace / 1073741824, 2);
                    $usedSpaceGB = [Math]::Round($sizeGB - $freeSpaceGB, 2);
                    $color = $whiteColor;

                    # Set background color to Orange if just a warning
	                if($percentFree -lt $percentWarning) { 
                        $color = $orangeColor
                        $global:i++
                    }

                    # Set background color to Orange if space is Critical
                    if($percentFree -lt $percentCritcal) { 
                        $color = $redColor
                        $global:i++
                    }
 
                    # Create table data rows 
                    $HTML_dataRow = "
		                <tr>
                        <td width='10%' bgcolor=`'$color`'>$computer</td>
		                <td width='5%' bgcolor=`'$color`' align='center'>$deviceID</td>
		                <td width='15%' bgcolor=`'$color`'>$volName</td>
		                <td width='10%' bgcolor=`'$color`' align='center'>$sizeGB</td>
		                <td width='10%' bgcolor=`'$color`' align='center'>$usedSpaceGB</td>
		                <td width='10%' bgcolor=`'$color`' align='center'>$freeSpaceGB</td>
		                <td width='5%' bgcolor=`'$color`' align='center'>$percentFree</td>
		                </tr>
                "
                if(($percentFree -lt $percentWarning -or $percentFree -lt $percentCritcal) -and $report -eq 'Warning') { 
                    $HTML_Body = $HTML_Body + $HTML_dataRow
                    Write-Host -ForegroundColor Red -BackgroundColor Black "$computer $deviceID percentage free space = $percentFree";
                } 
                elseif ( $report -eq 'Full') { 
                    $HTML_Body = $HTML_Body + $HTML_dataRow
                    Write-Output "$computer $deviceID percentage free space = $percentFree";
                }
                else { Write-Output "$computer $deviceID percentage free space = $percentFree"; }
                }  
            }
            else { 
                # Create table data rows 
                $color = $greyColor
                $HTML_dataRow = "
                    <tr>
                    <td width='10%' bgcolor=`'$color`'>$computer</td>
                    <td width='5%' bgcolor=`'$color`' align='center'>OFFLINE</td>
                    <td width='15%' bgcolor=`'$color`'>OFFLINE</td>
                    <td width='10%' bgcolor=`'$color`' align='center'>OFFLINE</td>
                    <td width='10%' bgcolor=`'$color`' align='center'>OFFLINE</td>
                    <td width='10%'bgcolor=`'$color`'  align='center'>OFFLINE</td>
                    <td width='5%' bgcolor=`'$color`' align='center'>OFFLINE</td>
                    </tr>
                "
                if($report -eq 'Full' ) { 
                    $HTML_Body = $HTML_Body + $HTML_dataRow
                    Write-Host -ForegroundColor Red -BackgroundColor Black "$computer is Offline";
                } 
                else { Write-Host -ForegroundColor DarkYellow "$computer is Offline"; }
            }
        }

foreach ($vCenter in $vCenters) {
    $VIVolumes = $null
    # Get vCenter Information
    Connect-ViServer -server $vcenter
    get-cluster
    $VIVolumes = Get-Datastore | sort -Property Name
    foreach ($VIVolume in $VIVolumes) {
                    $deviceID = $null;
                    $volName = $VIVolume.Name;
	                [float]$size = $VIVolume.CapacityGB;
		            [float]$freespace = $VIVolume.FreeSpaceGB; 
		            $percentFree = [Math]::Round(($freespace / $size) * 100, 2);
		            $sizeGB = [Math]::Round(($VIVolume.CapacityGB), 2);
		            $freeSpaceGB = [Math]::Round(($freespace), 2);
                    $usedSpaceGB = [Math]::Round(($size - $freespace), 2);
                    $color = $whiteColor;

                    # Set background color to Orange if just a warning
	                if($percentFree -lt $percentWarning) { 
                        $color = $orangeColor
                        $global:i++
                    }

                    # Set background color to Orange if space is Critical
                    if($percentFree -lt $percentCritcal) { 
                        $color = $redColor
                        $global:i++
                    }
                    # Create table data rows 
                    $HTML_dataRow = "
		                <tr>
                        <td width='10%' bgcolor=`'$color`'>$vcenter</td>
		                <td width='5%' bgcolor=`'$color`' align='center'>$deviceID</td>
		                <td width='15%' bgcolor=`'$color`'>$volName</td>
		                <td width='10%' bgcolor=`'$color`' align='center'>$sizeGB</td>
		                <td width='10%' bgcolor=`'$color`' align='center'>$usedSpaceGB</td>
		                <td width='10%' bgcolor=`'$color`' align='center'>$freeSpaceGB</td>
		                <td width='5%' bgcolor=`'$color`' align='center'>$percentFree</td>
		                </tr>
                "
                if(($percentFree -lt $percentWarning -or $percentFree -lt $percentCritcal) -and $report -eq 'Warning') { 
                    $HTML_Body = $HTML_Body + $HTML_dataRow
                    Write-Host -ForegroundColor Red -BackgroundColor Black "$computer $deviceID percentage free space = $percentFree";
                } 
                elseif ( $report -eq 'Full') { 
                    $HTML_Body = $HTML_Body + $HTML_dataRow
                    Write-Output "$computer $deviceID percentage free space = $percentFree";
                }
                else { Write-Output "$computer $deviceID percentage free space = $percentFree"; }
    } 
    Disconnect-VIServer -server * -Confirm:$false
}


    # Create table at end of report showing legend of colors for the critical and warning
     $HTML_tableDescription = "
     </table><br><table width='20%'>
	    <tr bgcolor='White'>
        <td width='20%' align='center' bgcolor='#FBB917'>Warning less than $percentWarning% free space</td>
	    <td width='20%' align='center' bgcolor='#FF0000'>Critical less than $percentCritcal% free space</td>
	    </tr>
    </table><br><table width='100%'>
        <tr bgcolor='White'>
        <td width='100%' align='center'>This script was executed on $evv:ComputerName and monitors the servers contained in the Servers OU from Active Directory</td>
        </tr>
        <tr>
        <td width='100%' align='center'>Send e-mail: $sendemail / Filter Oracle Data Disks: $filter_ora / Report Level: $report / Script Execution: $Env:ComputerName </td>
        </tr>
        </body></html>"
# Create the Report File
    Add-Content $Report_Path $HTML_Header
    Add-Content $Report_Path $HTML_tableHeader
    Add-Content $Report_Path $HTML_Body
    Add-Content $Report_Path $HTML_tableDescription

# Send Notification if alert $i is greater then 0
if ($sendemail -eq 'Yes') {
    $email_Subject = "[REPORT][$Domain_DNS] $title Report for $titledate"
    $email_body = Get-Content $Report_Path -Raw
    if ($email_report -eq 'Warning') {
        if ($global:i -gt 0) {
            foreach ($user in $email_users) {
                Write-Output "Sending e-mail to $user"
                Send-MailMessage -To $user -From $email_from -Subject $email_subject -SmtpServer $email_smtp -Port $email_port -UseSsl -Credential $email_credentials -Body $email_body -BodyAsHtml -Attachments $Report_Path
            }
        }
    }
    elseif ($email_report -eq 'Full') {
        foreach ($user in $email_users) {
            Write-Output "Sending e-mail to $user"
            Send-MailMessage -To $user -From $email_from -Subject $email_subject -SmtpServer $email_smtp -Port $email_port -UseSsl -Credential $email_credentials -Body $email_body -BodyAsHtml -Attachments $Report_Path
        }
    }

}
# Copy Report file to Destination Path
Copy-Item $Report_Path $Destination_Path -Force
Remove-Item $Report_Path -Force

if ($Error) { $Error | Out-File $log_file -Append } 