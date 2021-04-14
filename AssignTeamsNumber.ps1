<#
.SYNOPSIS
    This script is designed to assign a user a license, relevant policies and number in Skype for Business Online
.DESCRIPTION
    The script contains logging that is exported to the path of the $Logfile variable. Initially connects to SFBOnline, confrims has an active Teams license and doesn't already have a number assigned to them. 
    Then various calling policies are assigned and the number inputed will be assigned to the user. You will be prompted for caller ID and international calling preferences which will be set at the end of the script. 
.NOTES
    Generated On: 29/01/2021
    Update On: 24/03/2021
    Author: Matthew Heuer
    NOTE: Confirm that version 1.1.6 of the MicrosoftTeams module (run Get-Module to confirm) is installed, if you have any other version then run the following before proceeding:
        Uninstall-Module -Name MicrosoftTeams
        Install-ModuleMicrosoftTeams -RequiredVersion1.1.6
#>

Import-Module MicrosoftTeams
Import-Module ActiveDirectory

Function Add-LogEntry {
    Param([ValidateSet("Error", "Info", "Warning")][String]$LogLevel, [String]$LogEntry)
    $TimeStamp = Get-Date -Format "dd/MM/yyyy HH:mm:ss"
    $StreamWriter = New-Object System.IO.StreamWriter -ArgumentList ([IO.File]::Open($LogFile, "Append"))
    $StreamWriter.WriteLine("$TimeStamp - $LogLevel - $LogEntry")
    if ($LogLevel -eq 'Error') {
        Write-Host "$TimeStamp - $LogLevel - $LogEntry" -ForegroundColor Red
    } elseif ($LogLevel -eq 'Warning') {
        Write-Host "$TimeStamp - $LogLevel - $LogEntry" -ForegroundColor Yellow 
    } elseif ($LogLevel -eq 'Info') {
        Write-Host "$TimeStamp - $LogLevel - $LogEntry" -ForegroundColor Green
    }
    $StreamWriter.Close()
}

$Today = Get-Date -Format "ddMMyyyy"
$Logfile = "#LogFilePath#\ProvisionTeamsNumberLog_$Today.txt"

# Checks for an active session and then connects to Skype for Business Online
$getsessions = Get-PSSession | Select-Object -Property State, Name
$isconnected = (@($getsessions) -like '@{State=Opened; Name=SfBPowerShellSessionViaTeamsModule*').Count -gt 0
Add-LogEntry -LogLevel Info -LogEntry "Checking for active SFBOnline connections..."
If ($isconnected -ne "True") {
    Try { 
    $sfbSession = New-CsOnlineSession
    Import-PSSession $sfbSession -AllowClobber
    Add-LogEntry -LogLevel Info -LogEntry "No active connection found, establishing new connection to SFBOnline"
} Catch {
    Add-LogEntry -LogLevel Error -LogEntry "SFBOnline failed to connect, exiting script"
    Exit
}}
Add-LogEntry -LogLevel Info -LogEntry "Connection established to SFBOnline"

$UPN = Read-Host "Enter the affected users UPN"

# Queries the account to determine if it already has an active Teams license, then checks to see if they already have a number assgined. If both of these checks are passed the script will continue otherwise it will exit.
Add-LogEntry -LogLevel Info -LogEntry "Checking $UPN in SFBOnline..."
$Number = Get-CsOnlineUser $UPN | Select-Object -ExpandProperty LineURI
$License = ((Get-ADUser -Filter "UserPrincipalName -eq '$UPN'" -Properties *).memberof -like "#Name of the group that provisions Teams licenses#*")
if (!$License) {
    Add-LogEntry -LogLevel Error -LogEntry "$UPN doesn't have a Teams license, exiting script"
    Exit
} elseif ($License -eq '#Name of the group that provisions Teams licenses#' -and $number -gt 0) {
    Add-LogEntry -LogLevel Warning -LogEntry "$UPN already has $Number assigned to them in SFBOnline!"
    Exit
} elseif ($License -eq '#Name of the group that provisions Teams licenses#' -and $number -lt 1) {
    Add-LogEntry -LogLevel Info -LogEntry "$UPN has a license but no number assigned"
}

# Enter the next available number within the allocated range and select Y or N to enable no caller ID and international calling. 
$NextAvail = Read-Host -Prompt "Enter the next available number you want to assign to the user"
Add-LogEntry -LogLevel Info -LogEntry "Checking $NextAvail is available in SFBOnline..."
if (Get-CsOnlineUser | Where-Object {$_.LineURI -eq "tel:+$NextAvail"}) {
    Add-LogEntry -LogLevel Error -LogEntry "$NextAvail is already assigned, please try again"
    Exit
} Else {
    Add-LogEntry -LogLevel Info -LogEntry "$NextAvail is available to assign"
}
$NoCallerID = Read-Host -Prompt "Does the user require their Caller ID to be blocked? [Y/N]"
$InternationalPolicy = Read-Host -Prompt "Does the user require international calls enabled? [Y/N]"

# Assigns number and relevant policies to user
Try {
    Set-CsUser -Identity $UPN -EnterpriseVoiceEnabled $True -HostedVoicemail $True -OnPremLineURI "tel:+$NextAvail" -ErrorAction stop
    Grant-CsTenantDialPlan -PolicyName tag:DoCDialPlan -Identity $UPN
    Grant-CsOnlineVOiceRoutingPolicy -PolicyName "AU-WA-National" -Identity $UPN
    Grant-CsTeamsCallingPolicy -PolicyName AllowCalling -Identity $UPN
    Add-LogEntry -LogLevel Info -LogEntry "$NextAvail has been assigned to $UPN"
} Catch {
    Add-LogEntry -LogLevel Error -LogEntry "Unable to assign number to user, please review inputs and try again"
    Exit
}

# Sets the CallingLineIdentity to Anonymous
if ($NoCallerID -like 'y') {
    Grant-CsCallingLineIdentity -Identity $UPN -PolicyName Anonymous -ErrorAction stop
    Add-LogEntry -LogLevel Info -LogEntry "Set CallingLineIdentity policy to Anonymous"
} elseif ($NoCallerID -like 'n') {
    Continue
    Add-LogEntry -LogLevel Info -LogEntry "CallingLineIdentity policy set to default"
} else {
    Add-LogEntry -LogLevel Error -LogEntry "Input not valid, CallingLineIdentity policy set to default"
}

# Sets the VoiceRoutingPolicy to TeamsUnRestrictedVoiceRoutingPolicy to enable international calls
if ($InternationalPolicy -like 'y') {
    Grant-CsOnlineVoiceRoutingPolicy -Identity $UPN -PolicyName TeamsUnRestrictedVoiceRoutingPolicy -ErrorAction stop
    Add-LogEntry -LogLevel Info -LogEntry "Set VoiceRoutingPolicy policy to TeamsUnRestrictedVoiceRoutingPolicy"
} elseif ($InternationalPolicy -like 'n') {
    Continue
    Add-LogEntry -LogLevel Info -LogEntry "VoiceRoutingPolicy policy set to default"
} else {
    Add-LogEntry -LogLevel Error -LogEntry "Input not valid, VoiceRoutingPolicy policy set to default"
}

Add-LogEntry -LogLevel Info -LogEntry "Script completed for $UPN"
