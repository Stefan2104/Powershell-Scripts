import-module activedirectory
#Set domain
$domain = "global.centralorg.net"
#File containing users to remove
$users = import-csv C:\users\sflas.GLOBAL\Desktop\SMS_Passcode_cleaup.csv
#Create array for not removed users
$not_removed = @()

#Prompt for input to ensure execute
$run = Read-host "Are you sure you want to run this action? (y/n)"
if ($run -eq 'y'){
    #Loop through all usernames (Pre-W2k Name) in file
    foreach ($user in $users."Pre-W2K Name"){

        #Try and remove user
        try{
            #Cmdlet to remove user
            Remove-ADGroupMember -Server $domain -Identity WW_APP_SMSPasscode_Users -Members $user
	        Write-Host -NoNewline "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] -> Removing $user : "
	        Write-Host -ForegroundColor Green "OK"
        }

        #Catches error, and adds user to array, if user couldn't be removed
        catch{
	        Write-Host -NoNewline "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] -> Removing $user : "
	        Write-Host -ForegroundColor Red "Failed"
            $not_removed += $user
        }
    }
    #Writes all not-removed users to file
    $not_removed | Out-File "C:\users\sflas.GLOBAL\Desktop\not_removed.csv"
}

#Exits if input not-equal to 'y'
else {
    write-host "You didn't accept to run the action"
    exit
}