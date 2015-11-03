## Script to get and create a CSV report of all HDDs in a list
## Check if store is online
## if yes then get HDD space and free space and put it in a list
## export this list in an CSV file
#
#clear the screen
clear

#Declare the files we will be using
$computers = "C:\hdd_report\list_of_pc_to_be_queried.csv"
$reportfile = "C:\hdd_report\HDD_Report_$((Get-Date).ToString('yyyy-MM-dd')).csv"
          

##Import Store list to be queried
$computerlist = get-Content $computers  # Replace it with your TXT file which contain Name of Computers 

## Define the local file to be checked if store is online
## explorer.exe seems a good example
$WantFile = "\c$\Windows\explorer.exe" 

## Create report
## Push info where possible and "Offline" Value when Store is offline
foreach ($computer in $computerlist)
    {$POSonline = Test-Path "\\$computer$WantFile"
     trap { continue; }
        if ($POSonline -eq $true) 
        { 
            Get-WMIObject -ComputerName $computer Win32_LogicalDisk `
            | select `
                SystemName,`
                DriveType, `
                VolumeName, `
                Name, `
                @{n='Size (Gb)' ;e={"{0:n2}" -f ($_.size/1gb)}}, `
                @{n='FreeSpace (Gb)';e={"{0:n2}" -f ($_.freespace/1gb)}}, `
                @{n='PercentFree';e={"{0:n2}" -f ($_.freespace/$_.size*100)}} `
            | Where-Object {$_.DriveType -eq 3} `
            | Export-CSV -Append $reportfile -NoTypeInformation `
            | Sort-Object -Property freespace
        }
        else 
        {
        # Padd data for offline Computers
        Add-Content -Path $reportfile `
        -Value "`"$computer`",`"Offline`",`"Offline`",`"Offline`",`"Offline`",`"Offline`",`"Offline`""
        }  
    } 
	
	
# Send Email
$EmailFrom = "from@email.com"
$EmailTo1 = "to1@email.com"
$EmailTo2 = "to2@email.com"
$Subject = "HDD Space report" 
$Body = "Please find in the attachment the report for HDD Free space" 
$SMTPServer = "your.smtp.server.com" 
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 25) 
#optional SSL
#$SMTPClient.EnableSsl = $true 
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential("from@email.com", "your-password"); 


$emailMessage = New-Object System.Net.Mail.MailMessage
$emailMessage.From = $EmailFrom
$emailMessage.To.Add($EmailTo1)
$emailMessage.To.Add($EmailTo2)
$emailMessage.Subject = $Subject
$emailMessage.Body = $Body
$emailMessage.Attachments.Add("$reportfile")
$SMTPClient.Send($emailMessage)
