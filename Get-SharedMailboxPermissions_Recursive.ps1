#########################################################################################
# LEGAL DISCLAIMER
# This Sample Code is provided for the purpose of illustration only and is not
# intended to be used in a production environment.  THIS SAMPLE CODE AND ANY
# RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER
# EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF
# MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a
# nonexclusive, royalty-free right to use and modify the Sample Code and to
# reproduce and distribute the object code form of the Sample Code, provided
# that You agree: (i) to not use Our name, logo, or trademarks to market Your
# software product in which the Sample Code is embedded; (ii) to include a valid
# copyright notice on Your software product in which the Sample Code is embedded;
# and (iii) to indemnify, hold harmless, and defend Us and Our suppliers from and
# against any claims or lawsuits, including attorneys’ fees, that arise or result
# from the use or distribution of the Sample Code.
# 
# This posting is provided "AS IS" with no warranties, and confers no rights. Use
# of included script samples are subject to the terms specified at 
# https://www.microsoft.com/en-us/legal/intellectualproperty/copyright/default.aspx.
#
# Recursively gathers and reports who has access to all shared mailboxes and their access to it
# Get-SharedMailboxPermissions_Recursive.ps1
#  
# Created by: Kevin Bloom 2/4/2021 Kevin.Bloom@Microsoft.com 
#
#########################################################################################

##Define variables and constants
#Array to collect and gather all of the results
$Global:Records = @()
#Gets the date and is used for the output file
$DateTicks = (Get-Date).Ticks
#Output file
$OutputFile = "C:\Temp\SharedMailboxPermissionsRecursiv_$DateTicks.csv"

##Functions
#Function to enumerate the sharedmailbox permissions
Function Enumerate-SharedmailboxAccess
{
    Param ($Item)
    Write-Host $Item.user" -Enumerate SharedmailboxAccess function" -ForegroundColor Cyan
    #Gets recipient type, will be used to determine if the object is a user or group
    $IDRecipientType = (Get-Recipient -Identity $($Item.User)).RecipientType
    #If entry is not a group, add the values to a hash table and add the record to the $Records array
    If ($IDRecipientType -notlike "*group*")
    {
        $Record = "" | select Identity,User,AccessRights
        $Record.Identity = $Item.Identity
        $Record.User = $Item.User
        $Record.AccessRights = $Item.AccessRights
        $global:Records += $Record
    }
    #If entry is  a group, send the group to the Enumerate-Group function
    Elseif ($IDRecipientType -like "*group*")
    {
        Enumerate-Group ($Item)
    }
}

#Function to enumerate group memberships including nested groups
Function Enumerate-Group
{
    Param ($Group)
    Write-Host $Group.user" -Enumerate Group function" -ForegroundColor Cyan
    #Gets the members of the group
    $GroupMembers = Get-DistributionGroupMember -ResultSize Unlimited -Identity $($Group.user)
    Foreach ($GroupMember in $GroupMembers)
    {
        #If entry is not a group, add the values to a hash table and add the record to the $Records array
        If ($GroupMember.RecipientTypeDetails -notlike "*group*")
        {
            $Record = "" | select Identity,User,AccessRights
            $Record.Identity = $Group.Identity
            $Record.User = $GroupMember.PrimarySmtpAddress.Address
            $Record.AccessRights = $Group.AccessRights
            $global:Records += $Record
        }
        #If entry is  a group, send the group to the Enumerate-Group function
        Elseif ($GroupMember.RecipientTypeDetails -like "*group*")
        {
            $SubGroup = "" | select Identity,User,AccessRights
            $SubGroup.Identity = $Group.Identity
            $SubGroup.User = $GroupMember.PrimarySmtpAddress.Address
            $SubGroup.AccessRights = $Group.AccessRights
            #Calls itself so the nested group can be enumerated
            Enumerate-Group ($SubGroup)
        }
    }
}

##Primary script
#Gathers all shared mailboxes in an environment
$SharedMailboxes = Get-Mailbox -ResultSize unlimited -RecipientTypeDetails sharedmailbox 
#Loops through all shared mailboxes
Foreach ($SharedMailbox in $SharedMailboxes)
{
    #Retreives the root shared mailbox permissions and filters out non-users and non-groups
    $SharedmailboxPermissions = Get-MailboxPermission -ResultSize unlimited -Identity $SharedMailbox 
    $SharedmailboxPermissions = $SharedmailboxPermissions | Where-Object {$_.user.tostring() -ne "NT AUTHORITY\SELF" -and $_.IsInherited -eq $false}
    #Loops through ever entry in the shared mailbox permissions
    Foreach ($Item in $SharedmailboxPermissions)
    {
        #Creates a hash table used to pass objects to the functions
        $Object = "" | select Identity,User,AccessRights
        $Object.Identity = $Item.Identity.tostring()
        $Object.User = $Item.User.tostring()
        $Object.AccessRights = $Item.AccessRights
        Write-Host $Object.user" -main script"  -ForegroundColor Cyan
        Enumerate-SharedmailboxAccess $Object
    }
}

#Exports the results:
$Global:Records | select Identity,user,@{name='AccessRights';Expression={[string]::join(";",($_.accessrights))}} |Export-Csv $OutputFile