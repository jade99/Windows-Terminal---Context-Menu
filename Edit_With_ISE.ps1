# ----==== User Config Section ====----

# path to save registry file to
$OutFilePath = "$env:USERPROFILE\Desktop\wt_context_menu.reg"    


# Name shown in context menu
$ContextMenuName = "Open Windows Terminal here"   

# Only show entry when holding shift               
$ContexMenuShiftOpen = $true    

# Show profile icon if configured
$ContextMenuShowIcon = $true                                


# can cause permission issues
# change yourself to owner of key [HKCR/Directory/Background/shell/Powershell] and give yourself write permission
$HidePowershell = $true
$HideWSL = $true

# ---- Uninstall Section ----

# only reverts Windows Terminal key
$uninstall = $false

# ---- Debug ----
$verbose = $true

# ----==== End of User Config Section ====----
#    DO NOT CHANGE ANYTHING BELOW THIS LINE

<# Gets Windows Terminal settings.json and reads out configured profiles #>
Function GetProfiles {

    <# Reading Profile #>
    Write-Verbose "`treading profile..." -Verbose:$verbose
    $profile = (Get-Item "$env:LOCALAPPDATA\Packages\Microsoft.WindowsTerminal_*\LocalState\settings.json")
    $dirtyProfile = Get-Content $profile


    <# Remove comments #>
    Write-Verbose "`tremoving comments..." -Verbose:$verbose
    $jsonProfile = ''
    $dirtyProfile | ForEach-Object {
        $jsonProfile += ($_ -replace "^\s*\/\/.*$", '') + "`n"
    }


    <# Convert to PSObject and get profile list #>
    Write-Verbose "`tconverting to PSObject..." -Verbose:$verbose
    $objProfile = ConvertFrom-Json $jsonProfile
    $profileList = $objProfile.profiles.list

    <# Prompt user for wanted profiles #>
    Write-Verbose "`tuser prompt..." -Verbose:$verbose
    $selectedProfiles = $profileList | Where-Object {$_.Hidden -ne $true } | Select-Object -Property Name, commandline, guid, icon | Out-GridView -OutputMode Multiple -Title "Select profiles you want to add to context menu"
    return $selectedProfiles
}

<# Write Registry File #>
Function WriteRegFile {
    <# See https://support.microsoft.com/en-us/help/310516/how-to-add-modify-or-delete-registry-subkeys-and-values-by-using-a-reg #>

    <# Registry Version Number #>
    Write-Verbose "creating registry file..." -Verbose:$verbose
    "Windows Registry Editor Version 5.00`n" | Out-File $OutFilePath

    if (-not $uninstall) {
        <# Install Branch #>
        <# sanity information multiple escaped string sequences used #>
        
        # Get WT Profiles
        Write-Verbose "requesting profiles from user ..." -Verbose:$verbose
        $profiles = GetProfiles

        
        # Create Parent Key
        Write-Verbose "creating parent key..." -Verbose:$verbose
        "[HKEY_CLASSES_ROOT\Directory\Background\shell\Windows Terminal]" | Out-File $OutFilePath -Append
        '@=""' | Out-File $OutFilePath -Append
        "`"MUIVerb`"=`"$ContextMenuName`"" | Out-File $OutFilePath -Append
        "`"SubCommands`"=`"$($profiles.guid | ForEach-Object { "$_;" })`"" | Out-File $OutFilePath -Append
    
        # Enable Shift Only if configured
        if ($ContexMenuShiftOpen) {
            Write-Verbose "enabling shift only mode" -Verbose:$verbose
            '"Extended"=""' | Out-File $OutFilePath -Append
            '"NoWorkingDirectory"=""' | Out-File $OutFilePath -Append
            '"ShowBasedOnVelocityId"=dword:00639bc8' | Out-File $OutFilePath -Append
        }
        '' | Out-File $OutFilePath -Append

        # Iterate through each selected profile
        Write-Verbose "iterating through profiles..." -Verbose:$verbose
        $profiles | ForEach-Object {
            # Create sub key for profile
            Write-Verbose "creating key for profile `"$($_.name)`"" -Verbose:$verbose
            "[HKEY_CLASSES_ROOT\Directory\Background\shell\Windows Terminal\shell\$($_.guid)]" | Out-File $OutFilePath -Append
            "@=`"$($_.name)`"" | Out-File $OutFilePath -Append

            # if allowed and profile has icon show it in context menu
            if ($_.icon -and $ContextMenuShowIcon) {
                Write-Verbose "`tprofile has icon, configuring..." -Verbose:$verbose
                "`"Icon`"=`"$($_.icon -replace '\\', '\\')`"`n" | Out-File $OutFilePath -Append
            } else {
                '' | Out-File $OutFilePath -Append
            }

            # Add key for command to execute
            Write-Verbose "`tconfiguring command..." -Verbose:$verbose
            "[HKEY_CLASSES_ROOT\Directory\Background\shell\Windows Terminal\shell\$($_.guid)\command]" | Out-File $OutFilePath -Append
            "@=`"wt -p $($_.guid) -d \`"%V \`"`"`n" | Out-File $OutFilePath -Append
        }


        # Hide powershell if configured - see userconfig for permission info
        if ($HidePowershell) {
            Write-Verbose "hiding powershell..." -Verbose:$verbose
            "[HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShell]" | Out-File $OutFilePath -Append
            '"ShowBasedOnVelocityId"=-' | Out-File $OutFilePath -Append
            "`"HideBasedOnVelocityid`"=dword:00639bc8`n" | Out-File $OutFilePath -Append
        }

        # Hide WSL if configured - see userconfig for permission info
        if ($HideWSL) {
            Write-Verbose "hiding WSL..." -Verbose:$verbose
            "[HKEY_CLASSES_ROOT\Directory\Background\shell\WSL]" | Out-File $OutFilePath -Append
            '"ShowBasedOnVelocityId"=-' | Out-File $OutFilePath -Append
            "`"HideBasedOnVelocityid`"=dword:00639bc8`n" | Out-File $OutFilePath -Append
        }
    } else {
        <# Uninstall branch #>

        # Remove "Windows Terminal" parent key and all sub-keys
        Write-Verbose "removing parent key..." -Verbose:$verbose 
        "[-HKEY_CLASSES_ROOT\Directory\Background\shell\Windows Terminal]`n" | Out-File $OutFilePath -Append

        # Re-enable powershell - see userconfig for permission info
        if (-not $HidePowershell) {
            Write-Verbose "re-enabling powershell..." -Verbose:$verbose
            "[HKEY_CLASSES_ROOT\Directory\Background\shell\PowerShell]" | Out-File $OutFilePath -Append
            "`"ShowBasedOnVelocityId`"=dword:00639bc8" | Out-File $OutFilePath -Append
            "`"HideBasedOnVelocityid`"=-`n" | Out-File $OutFilePath -Append
        }

        # Re-enable WSL - see userconfig for permission info
        if (-not $HideWSL) {
            Write-Verbose "re-enabling WSL..." -Verbose:$verbose
            "[HKEY_CLASSES_ROOT\Directory\Background\shell\WSL]" | Out-File $OutFilePath -Append
            "`"ShowBasedOnVelocityId`"=dword:00639bc8" | Out-File $OutFilePath -Append
            "`"HideBasedOnVelocityid`"=-`n" | Out-File $OutFilePath -Append
        }
    }
}

WriteRegFile