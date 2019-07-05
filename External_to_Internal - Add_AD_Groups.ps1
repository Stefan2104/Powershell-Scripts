# The purpose of the script is to automatically
# add all groups of an external user to an internal user
# when the user has been hired as an internal.
#
# Instead of adding these groups manually,
# by running this script only the username of the external and internal user
# needs to be typed in when prompted.
#
# Created by Stefan Larsen (SFLAS) from Arla IT Service Desk.




import-module activedirectory
#set domain server
$domain = "global.centralorg.net"
#Create empty array for groups not added to user
$not_added = @()
#Create empty array for groups not added to user after second attempt
$not_added_verify = @()

#Function to read external and internal username from user-input
Function GetUsers{
    #Get ext_username from input
    $ext_user = Read-host "Username of the external user?"
    #get internal username from input
    $int_user = Read-host "Username of the internal user?"

    #Output user-input
    Write-Host -NoNewline "External user: "
    Write-Host -ForegroundColor Green "$ext_user"
    Write-Host -NoNewline "Interal user: "
    Write-Host -ForegroundColor Green "$int_user"

    #Ask user to confirm input
    $confirm = Read-host "Confirm above users (y/n)"

    #Check if user confirms, if yes, continues
    if ($confirm -eq 'y'){ 
        AddGroups
    }

    #If no, go to TryAgain function 
    else {
        TryAgain

    }   
}

# Main function, adds all groups the external is member of
# to the internal user defined in the beginning
# Thereafter, it verifies if the user is in fact added
Function AddGroups{
    #Get all groups the external user is a member off
    $ext_user_groups = Get-ADPrincipalGroupMembership -Server $domain -Identity $ext_user | select -ExpandProperty name
    #Loop through all groups ext_user is member off 
    Foreach ($group in $ext_user_groups){
        #Try to add internal user to the group
        try{
            Add-ADGroupMember -Server $domain -Identity $group -Members $int_user
            #Displays success message
            Write-Host -NoNewline "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] -> Added $int_user to: $group"
	        Write-Host -ForegroundColor Green "OK"
        }
        #Catch error if user couldn't be added
        catch{
            #Displays failed message
            Write-Host -NoNewline "[$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")] -> Added $int_user to: $group"
	        Write-Host -ForegroundColor Red "Failed"
            #Adding failed group to array defined in line 18
            $not_added += $group
        }
    }
    #After main loop is finished, run 'Verify' function 
    #To verify all the groups has been added
    Verify
}

Function TryAgain{
    #Prompt user to try again (y/n)
    $Try_again = Read-host "Do you want to try again? (y/n)"
    #If 'y' is entered, go to restart script (GetUsers function)
    if ($Try_again -eq 'y') { GetUsers }
    #If not 'y' is entered, abort script (Abort function)
    else { Abort }   
}

Function Verify{
    #Check if any failing added groups
    #By subtracting number of failed groups with number of groups
    $groups_added = $ext_user_groups.Length - $not_added.Length
    #Displays how many groups is added out of how many groups should be added
    Write-host "$groups_added groups added out of $ext_users_groups groups"
    #If some groups not added, prompt user to list not-added groups
    if ($groups_added -lt $ext_users_groups) {
        $show_not_added_groups = read-host "List groups failed to add user to? (y/n)"
        #If 'y' is entered, listing groups
        if ($show_not_added_groups -eq 'y') {
            write-host $groups_added
            #Ask user to check if user is in all groups either way
            $confirm_not_added = Read-host "Check if user is in the failed groups? (y/n)"
            if ($confirm_not_added -eq 'y'){
                #Check For each failed group, checking if user is in it
                ForEach ($failed_group in $not_added){
                     $members = Get-adgroupmember -Server $domain -Identity $failed_group | select -ExpandProperty samaccountname
                     #If memberlist of group contains user, display message
                     if ($members -contains $int_user){
                        Write-Host -NoNewline "$int_user in group $failed_group -> "
	                    Write-Host -ForegroundColor Green "Yes"
                     }

                     #If memberlist of group doesn't contain user, display message
                     #and add to array defined in line 20
                     else{
                        Write-Host -NoNewline "$int_user in group $failed_group -> "
	                    Write-Host -ForegroundColor Red "No" 
                        $not_added_verify += $failed_group
                     }
                }

                #If array containing failed group is not empty
                #Displays list of groups user can't be added to
                if ($not_added_verify -gt 0){
                    Write-Host "Could not add $int_user to below groups:"
                    Write-host $not_added_verify
                }

                #If array containing failed group is empty
                #Displays success message
                else{
                    Write-host "$int_user has been added to all groups"
                    exit
                }
            }
            else{
                Abort
            }
        }
        else{
            Abort
        }
    }

    #This is only run if no groups failed to add
    #I.e. line 95, $groups_added equals $ext_users_groups
    else{
        #Ask users if he/she wants to list all the groups user was added to
        $list_groups = Read-host "User was added to all groups. Want to list the groups? (y/n)"
        if ($list_groups -eq 'y') { write-host $ext_user_groups }
        else { Abort }
    }
    
}

#Function to abort script if user-input not equal 'y'
Function Abort{
    write-host "Action interrupted by user"
    exit  
}  

GetUsers


