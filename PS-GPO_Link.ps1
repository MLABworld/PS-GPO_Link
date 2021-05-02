# Run like this:
# powershell.exe -executionpolicy bypass -file "C:\Temp\PS Automation\GPOLinkUpdate-vX.ps1"

#####################################################################
# Supply values for $NewGroupPolicy and $OldGroupPolicy             #
#####################################################################
$NewGroupPolicy = "SamplePolicy_2"
$OldGroupPolicy = "SamplePolicy_1"
#####################################################################
# Supply values for the Organizational Units where linking and      #
# un-linking should occur. Include the linking order if applicable  #
# specify '0' if link order doesn't matter (the new link will land  #
# at the bottom)                                                    #
#####################################################################
$OUsForLinking = @(
                ('OU=Region1,DC=mlabeng,DC=lab,DC=local', '1'),
                ('OU=Region2,DC=mlabeng,DC=lab,DC=local', '3'),
                ('OU=Region3,DC=mlabeng,DC=lab,DC=local', '0')
#                ('OU=Region1,DC=mlabeng,DC=lab,DC=local', '2')
                )
#########################################################
#########################################################
#########################################################

# Set up logging
$LogFile = $PSCommandPath + "Log" + (Get-Date -Format "MM-dd-yy_HHmmss") + ".txt"
Add-Content -Path $LogFile -Value "Group Policy Link Automation - Powershell Script"
Add-Content -Path $LogFile -Value "Path of running script: $PSCommandPath"
Add-Content -Path $LogFile -Value $("Started: " + (Get-Date).ToString())

# Backup location
$BackupShare = "\\SERV-2019\GPBackups"
$BackupPath = $Backupshare + "\" + $OldGroupPolicy
# Backup existing Group Policy Object
Add-Content -Path $LogFile -Value "`r`n**************** BackUp Phase ****************"
if (Test-Path $BackupPath) 
    {
    # If the backup folder already exists, skip the backup
    Add-Content -Path $LogFile -Value "Path '$BackupPath' already exists. Skipping backup of '$OldGroupPolicy'."
    }
    else
    {
    # If there is no backup folder for the existing (old) Group Policy Object, create one
    New-Item -ItemType Directory -Path $BackupPath
    Backup-GPO -Name $OldGroupPolicy -Path $BackupPath
    Add-Content -Path $LogFile -Value "Backed up the Group Policy '$OldGroupPolicy' to '$BackupPath'."
    }

# Linking New Group Policy Object
Add-Content -Path $LogFile -Value "`r`n**************** Linking Phase ****************"
Write-Host "*** Linking Phase ***"
Write-host "OUs for linking" 
Write-Host $OUsForLinking
ForEach ($OUDN in $OUsForLinking)
    {    
        $OUToLink = $OUDN[0]
        $OUOrder = $OUDN[1]
        Write-host "OU to Link " $OUToLink
        Write-Host "OU Order " $OUOrder
        #Check of DN of OU is valid
       $X = [ADSI]"LDAP://$OUToLink"
       If ($X.Name)
        {
         # If the OU is valid, link the new GPO to it and un-link the existing GPO
                try { 
                    Write-host "OU " $OUToLink "is valid, linking."
                    If ($OUOrder -eq 0) 
                        {
                        New-GPLink -Name $NewGroupPolicy -Target $OUToLink -ErrorAction Stop
                        Add-Content -Path $LogFile -Value "Linked '$NewGroupPolicy' To '$OUToLink'"
                        Remove-GPLink -Name $OldGroupPolicy -Target $OUToLink -ErrorAction Stop
                        Add-Content -Path $LogFile -Value "Un-Linked '$OldGroupPolicy' From '$OUToLink'"
                        }
                    Else                 
                        {
                        New-GPLink -Name $NewGroupPolicy -Target $OUToLink -Order $OUOrder -ErrorAction Stop 
                        Add-Content -Path $LogFile -Value "Linked '$NewGroupPolicy' To '$OUToLink' LinkOrder '$OUOrder'"
                        Remove-GPLink -Name $OldGroupPolicy -Target $OUToLink -ErrorAction Stop
                        Add-Content -Path $LogFile -Value "Un-Linked '$OldGroupPolicy' From '$OUToLink'"
                        }
                    }
                catch { 
                        Add-Content -Path $LogFile -Value "    ***Warning*** Problem detected linking or un-linking on '$OUToLink'"
                      }
        }
       Else
        {
         # Skip this OU and write an alert in the logfile
         Add-Content -Path $LogFile -Value "    ***Warning*** OU not found: $OUToLink"
        }
    Add-Content -Path $LogFile -Value "`n"
    }
