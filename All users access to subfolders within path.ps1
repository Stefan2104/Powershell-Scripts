#Create Excel application
$excel = New-Object -ComObject excel.application
#Makes Excel visible
$excel.Application.Visible = $true
$excel.DisplayAlerts = $false
#Create Excel workBook
$Book = $excel.Workbooks.Add()
#Add Worksheets
#Gets the work sheet and Names it
$sheet = $book.Worksheets.Item(1)
$sheet.name = 'Computer information'
#Select a worksheet
$sheet.Activate() | Out-Null
#Create a row and set it to Row 1
$row = 1
$column = 1
#Set domain name
$domain = "global.centralorg.net"
$initialcolumn= 1
# Create initial row
$initialRow = 1


## Line 24-61: find all ATEA users in AD
# Set Excel sheet text information
$sheet.Cells.Item($row,$column) = 'ATEA users in AD'
$sheet.Cells.Item($row,$column).Font.Name = "Copper Black"
$sheet.Cells.Item($row,$column).Font.Size = 13
$sheet.Cells.Item($row,$column).Font.ColorIndex = 16
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 2
$sheet.Cells.Item($row,$column).HorizontalAlignment = -4108
$sheet.Cells.Item($row,$column).Font.Bold = $true
# Move to the next row
$row++
#Create headers
$sheet.Cells.Item($row,$column) = "Username"
$sheet.Cells.Item($row,$column).Font.Size = 12
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true
$column++
$sheet.Cells.Item($row,$column) = "Full Name"
$sheet.Cells.Item($row,$column).Font.Size = 12
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true
$column++
$sheet.Cells.Item($row,$column) = "State:"
$sheet.Cells.Item($row,$column).Font.Size = 12
$sheet.Cells.Item($row,$column).Font.ColorIndex = 1
$sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
$sheet.Cells.Item($row,$column).Font.Bold = $true
#Headers are done, we down 1 row and back to column 1
$row++

#command used to get information
$users = get-aduser -Server $domain -Filter "(Title -like '*ATEA*')" -Properties samaccountname,displayname
ForEach($user in $users) {
    $column = $initialcolumn
    #Headers are done, we down row and back to column 1
    $sheet.Cells.Item($row,$column) = $user.samaccountname
    $column++
    $sheet.Cells.Item($row,$column) = $user.displayname
    $column++
        Switch($user.enabled){
          True{$USEnabled = "Enabled"; $sheet.Cells.Item($row,$column).Interior.ColorIndex = 4}
          False{$USEnabled = "Disabled"; $sheet.Cells.Item($row,$column).Interior.ColorIndex = 3}
    }
    $sheet.Cells.Item($row,$column) = $USEnabled
    $row++
    $column = $initialcolumn
    
}    

# Moving to row first row and other column for different 
$row = 1
$initialcolumn += 4
$groups = "Arla-all-switch-RW-Atea","PA-USER-Atea-RW","PA-USER-Atea-R","PA-GLOBAL-ATEA-SERVER-SUPPORT"
ForEach($group in $groups) {
    $column = $initialcolumn
    # Set Excel sheet text information
    $sheet.Cells.Item($row,$column) = $group
    $sheet.Cells.Item($row,$column).Font.Name = "Copper Black"
    $sheet.Cells.Item($row,$column).Font.Size = 13
    $sheet.Cells.Item($row,$column).Font.ColorIndex = 16
    $sheet.Cells.Item($row,$column).Interior.ColorIndex = 2
    $sheet.Cells.Item($row,$column).HorizontalAlignment = -4108
    $sheet.Cells.Item($row,$column).Font.Bold = $true
    # Merge the cells

    # Move to the next row
    $row++

    #Create headers
    $sheet.Cells.Item($row,$column) = "Username // Initials"
    $sheet.Cells.Item($row,$column).Font.Size = 12
    $sheet.Cells.Item($row,$column).Font.ColorIndex = 1
    $sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
    $sheet.Cells.Item($row,$column).Font.Bold = $true
    $column++
    $sheet.Cells.Item($row,$column) = "Full Name"
    $sheet.Cells.Item($row,$column).Font.Size = 12
    $sheet.Cells.Item($row,$column).Font.ColorIndex = 1
    $sheet.Cells.Item($row,$column).Interior.ColorIndex = 48
    $sheet.Cells.Item($row,$column).Font.Bold = $true
    #Headers are done, we down row and back to column 1
    $row++
    $column = $initialcolumn
    #command used to get information
    $users = $users = Get-ADGroupMember -Server $domain -Identity $group | Where-Object { $_.objectClass -eq 'user' } | Get-ADUser -Properties displayname
    ForEach ($user in $users) { 
        $sheet.Cells.Item($row,$column) = $user.samaccountname
        $column++ 
        $sheet.Cells.Item($row,$column) = $user.displayname
        $row++
        $column = $initialcolumn
     }
     $row = 1
     $initialcolumn += 3
}

# Merge the cells
$sheet.Range('A1:B1').Merge() | Out-Null
$sheet.Range('E1:F1').Merge() | Out-Null
$sheet.Range('H1:i1').Merge() | Out-Null
$sheet.Range('K1:L1').Merge() | Out-Null
$sheet.Range('N1:O1').Merge() | Out-Null

# Fits cell to size
$UsedRange = $sheet.UsedRange
$UsedRange.entireColumn.autofit() | Out-Null

# Get date to append filename with
$CurrentDate = Get-Date
$CurrentDate = $CurrentDate.ToString('MM-dd-yyyy_hh-mm-ss')
# Set filename and path where to save
$FileName = "C:\Users\sflas.GLOBAL\Desktop\test folder\Scheduled task_$CurrentDate.csv"
# Save the file
$Book.SaveAs($FileName)
## $Book.Close()
# $Book.Close($false)
# $excel.Quit()
# $body = "Hi Niels. `nThe customized scheduled job to find ATEA users and extract from AD groups ran today `nAttached you will find the group members.`n`nBe adviced that this is an automated e-mail! `n`nIf any questions, write to me on Skype. Username: SFLAS`n`nHave a Wonderful day."
# Send-MailMessage -From "Stefan Larsen <sflas@arlafoods.com>" -To "Poul-Erik-cc-Niris <sflas@arlafoods.com>" -Subject "Monthly Extract of ATEA users - Scheduled task" -Body $body -Attachments $FileName -DeliveryNotificationOption OnFailure, OnSuccess -SmtpServer "smtp.arla.net"




 


