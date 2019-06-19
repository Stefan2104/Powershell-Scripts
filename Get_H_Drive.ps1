import-module activedirectory
clear

Function Check_Path{
    #Set Server to look in
    $domain = "global.centralorg.net"
    #Prompts for username 
    $read_user = Read-Host -Prompt 'Username of user you want to grant access to:'
    #Get AD user object, including H-drive and store it in $user
    $user = get-aduser -Server $domain -Identity $read_user -Properties homedirectory
    #Get userID (initials) and store in $username
    $username = $user.SamAccountName
    #Get user's H-drive from Profile in AD
    $drive_path = $user.homedirectory

    write-host "[+] Checking for user's H-drive file path in AD"
    #Checking if there's path to H-drive in AD
    if ($drive_path) {
        write-host "[+] Path found in AD`n"

        # Path found, granting user permission
        # to the H-drive path in AD profile.
        # Catches error if not possible

        Write-host "[+] Trying to grant user permission to $drive_path`n"
        $result = ICACLS ("$drive_path") /grant ("global\$username" + ':(OI)(CI)F') /T
        if ($result -eq 'Successfully processed 1 files; Failed processing 0 files'){
            # User has been granted access. Exitting
            clear
            write-host "[+] Access has been granted. Click 'enter' to exit"
            pause
            exit
        }else {
            clear
            write-host "[-] Couldn't grant the access for unknown reason`n[-] Checking if H-drive folder for user exist elsewhere`n"
            # Couldn't grant access for unknown reason, 
            # Jumping to Folder_exist function
            Folder_exist
                
        }


        
    }
    # Path in AD is empty. Jumping to Folder_exist function
    # To check if path in AD is wrong or folder non-existent
    else {
        write-host "[-] File path not found`n[-] Checking if user's H-drive folder exist elsewhere"
        Folder_exist
    }
}


function Folder_exist {
    # For-loop to loop through all ~\volxx\ folders
    # to check if H-drive is other paths than in AD
    For ($i=1; $i -lt 21; $i++) {
        # Loop through vol01-10
        # Using $i to set as index for checking all vol-folders
        if ($i -lt 10) {
            $path = "\\global.centralorg.net\ww-data\home\vol0$i\$username\"
            # Checking if folder exist
            if([System.IO.File]::Exists($path)){
                # Folder found, storing path in $drive_path
                # and jumping to 'Empty_folder' function
                $drive_path = "\\global.centralorg.net\ww-data\home\vol0$i\$username\"
                write-host "[+] H-drive folder exist on $drive_path.`nChecking if folder is empty"
                Empty_folder
             }
        }
        # Loop through vol10-20
        elseif ($i -le 20) {
            $path = "\\global.centralorg.net\ww-data\home\vol$i\$username\"
            # Checking if folder exist
            if([System.IO.File]::Exists($path)){
                # Folder found, storing path in $drive_path
                # and jumping to 'Empty_folder' function
                $drive_path = "\\global.centralorg.net\ww-data\home\vol$i\$username\"
                write-host "[+] H-drive folder exist on $drive_path.`nChecking if folder is empty"
                Empty_folder
            } 
        }
    }
    # No H-drive found for user in any folder.
    write-host "[-] No H-drive found on any vol-folder.`n[+] Trying to create new folder on vol20"
    # Jumping to Create-homedirectory function to create one
    Create-homedirectory

}  


function Empty_folder {
    # Measuring users H-drive
    $directoryinfo = Get-ChildItem $drive_path | Measure-Object
    # Checking if zero files with $_.count method
    if ($directoryinfo.Count -eq 0) {
        # H-drive is empty
        write-host "[+] H-drive is empty. Deleting folder, clearing homedirectory in AD and creating H-drive on ~\vol20\`n"
        # Removing folder to create on ~\vol20\ instead
        rmdir $drive_path
        # Jumping to Create-homedirectory function to create a
        Create-homedirectory
    }
    else { 
        # H-drive is not empty, aborting script
        write-host "[-] H-drive is not empty. Permissions can't be granted and folder is not being deleted`nAssign ticket to NNIT-IAM-Tech`n".
        exit
    }
}


function Create-homedirectory{
    # Storing new path to h-drive vol20 folder on variable
    $new_path = "\\global.centralorg.net\ww-data\home\vol20\$username"
    # Replacing path in user's AD profile
    set-aduser -Server $domain -Identity $user -HomeDirectory $new_path -Homedrive H;

    try {
        # Trying to createthe folder for user on vol20
        mkdir $new_path
        write-host "[+] H-drive created. Granting user access`n"
        try {
            # Trying to grant the user access to the new folder
            $result = ICACLS ("$new_path") /grant ("global\$username" + ':(OI)(CI)F') /T
            if ($result -eq 'Successfully processed 1 files; Failed processing 0 files'){
                write-host "[+] Access granted!`nUse command: 'net use h: $new_path' on client PC to add the H-drive manually.`n"
                # Display message if succeed
                pause
                exit
            }
        }
        catch {
            # Can't grant access, displays message and aborts
            write-host "[-]Access Denied. Assign to NNIT-IAM-Tech or contact SFLAS`n"
            pause
            exit
        }
    }
    catch {
        # Can't create folder, displays message and aborts
        write-host "[-] Access denied to create H-drive on $new_path.`nAssign to NNIT-IAM-Tech or contact SFLAS`n"
        pause
        exit
    }     
}


Check_Path




