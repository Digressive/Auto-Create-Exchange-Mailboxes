<#PSScriptInfo

.VERSION 1.7

.GUID 2905a44a-0932-41e9-9d09-b6339a9f0143

.AUTHOR Mike Galvin twitter.com/digressive

.COMPANYNAME

.COPYRIGHT (C) Mike Galvin. All rights reserved.

.TAGS Exchange Mailbox Active Directory Syncronization

.LICENSEURI

.PROJECTURI https://gal.vin/2017/06/07/powershell-create-mailboxes

.ICONURI

.EXTERNALMODULEDEPENDENCIES Exchange Management PowerShell module. Active Directory Management PowerShell module.

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES

#>

<#
    .SYNOPSIS
    Create Exchange mailboxes for users with no mailbox contained within an OU tree.

    .DESCRIPTION
    Create Exchange mailboxes for users with no mailbox contained within an OU tree.

    This script will:
    
    Create mailboxes for users contained witin an OU tree.
    You can configure the database and retention policy to use.
    Output and e-mail a log file.

    This script is designed to be run locally on an Exchange Server.

    Please note: to send a log file using ssl and an SMTP password you must generate an encrypted
    password file. The password file is unique to both the user and machine.
    
    The command is as follows:

    $creds = Get-Credential
    $creds.Password | ConvertFrom-SecureString | Set-Content c:\foo\ps-script-pwd.txt
    
    .PARAMETER Ou
    The AD Organisational Unit (including child OUs) that contains the users to create Exchange Mailboxes for.

    .PARAMETER Datab
    The Exchange database to create the mailboxes in. If you do not configure a Database, the smallest database will be used.

    .PARAMETER Rp
    The retention policy that should be applied to the users.
    
    .PARAMETER Compat
    Use this switch if you are using Exchange 2010.

    .PARAMETER L
    The path to output the log file to.
    The file name will be "Create-Mailboxes.log"
    
    .PARAMETER SendTo
    The e-mail address the log should be sent to.

    .PARAMETER From
    The from address the log should be sent from.

    .PARAMETER Smtp
    The DNS name or IP address of the SMTP server.

    .PARAMETER User
    The user account to connect to the SMTP server.

    .PARAMETER Pwd
    The txt file containing the encrypted password for the user account.

    .PARAMETER UseSsl
    Connect to the SMTP server using SSL.

    .EXAMPLE
    Create-Mailboxes.ps1 -Ou "OU=NewUsers,OU=Dept,DC=contoso,DC=com" -Datab "Mail DB 2" -Rp "1-Month-Deleted-Items" -L E:\scripts -Sendto me@contoso.com -From Exch01@contoso.com -Smtp smtp.live.com -User Exch01@contoso.com -Pwd P@ssw0rd -UseSsl

    This will create mailboxes for users that do not already have one in the OU NewUsers and all child OUs.
    It will create the mailbox using Mail DB 2 and apply the retention policy "1-Month-Deleted-Items".
    If you do not configure a Database, the smallest database will be used.
    A log will be output to E:\scripts and e-mail using a secure connection.

    The powershell code to get the smallest database is by Jason Sherry: https://blog.jasonsherry.net/2012/03/25/script_smallest_db/.
#>

[CmdletBinding()]
Param(
    [parameter(Mandatory=$True)]
    [alias("Ou")]
    $OrganisationalUnit,
    [alias("Datab")]
    $Database,
    [alias("Rp")]
    $Retention,
    [alias("L")]
    $LogPath,
    [alias("SendTo")]
    $MailTo,
    [alias("From")]
    $MailFrom,
    [alias("Smtp")]
    $SmtpServer,
    [alias("User")]
    $SmtpUser,
    [alias("Pwd")]
    $SmtpPwd,
    [switch]$UseSsl,
    [switch]$Compat)

## If compat is configured load the old Exchange PS Module, if not load the current one.
If ($Compat)
{
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
}

Else
{
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}

## Count number of users in ou specified without a mailbox
$UsersNo = Get-ADUser -SearchBase $OrganisationalUnit -Filter * -Properties mail | Where-Object {$_.mail -eq $null} | Measure-Object

## If users exist without mailboxes run the script
If ($UsersNo.count -ne 0)
{

    ## If logging is configured, start log
    If ($LogPath)
    {
        $LogFile = "Create-Mailboxes.log"
        $Log = "$LogPath\$LogFile"

        ## If the log file already exists, clear it
        $LogT = Test-Path -Path $Log
        If ($LogT)
        {
            Clear-Content -Path $Log
        }

        Add-Content -Path $Log -Value "****************************************"
        Add-Content -Path $Log -Value "$(Get-Date -Format G) Log started"
        Add-Content -Path $Log -Value ""
    }

    ## Find the users in the OU specified
    $Users = Get-ADUser -SearchBase $OrganisationalUnit -Filter *

    ## If a database is configure, create a mailbox for each user that does not have one, and set the retention policy if it has been specified
    If ($Database)
    {
        ForEach ($User in $Users)
        {
            If ($(Get-ADUser $User -Properties mail).mail -eq $null)
            {
                Enable-Mailbox -Identity $User.SamAccountName -Database $Database -RetentionPolicy $Retention

                ## Logging
                If ($LogPath)
                {
                    Add-Content -Path $Log -Value "$(Get-Date -Format G) Creating mailbox for User: $($User.SamAccountName) in Database: $Database with Retention Policy: $Retention"
                }
            }
        }
    }

    ## If a database is not configured, find the smallest database
    Else
    {
        $MBXDbs = Get-MailboxDatabase

        ForEach ($MBXDB in $MBXDbs)
        {
            $TotalItemSize = Get-MailboxStatistics -Database $MBXDB | %{$_.TotalItemSize.Value.ToMB()} | Measure-Object -sum
            $TotalDeletedItemSize = Get-MailboxStatistics -Database $MBXDB.DistinguishedName | %{$_.TotalDeletedItemSize.Value.ToMB()} | Measure-Object -sum
     
            $TotalDBSize = $TotalItemSize.Sum + $TotalDeletedItemSize.Sum

            If (($TotalDBSize -lt $SmallestDBsize) -or ($SmallestDBsize -eq $null))
            {
                $SmallestDBsize = $TotalDBSize
                $SmallestDB = $MBXDB
            }
        }

        ## For each user that does not have a mailbox, create one and set the retention policy if it has been specified
        ForEach ($User in $Users)
        {
            If ($(Get-ADUser $User -Properties mail).mail -eq $null)
            {
                Enable-Mailbox -Identity $User.SamAccountName -Database $SmallestDB -RetentionPolicy $Retention

                ## Logging
                If ($LogPath)
                {
                    Add-Content -Path $Log -Value "$(Get-Date -Format G) Creating mailbox for User: $($User.SamAccountName) in Database: $SmallestDB with Retention Policy: $Retention"
                }
            }
        }
    }

    ## If log was configured stop the log
    If ($LogPath)
    {
        Add-Content -Path $Log -Value ""
        Add-Content -Path $Log -Value "$(Get-Date -Format G) Log finished"
        Add-Content -Path $Log -Value "****************************************"

        ## If email was configured, set the variables for the email subject and body
        If ($SmtpServer)
        {
            $MailSubject = "Create Mailboxes Log"
            $MailBody = Get-Content -Path $Log | Out-String

            ## If an email password was configured, create a variable with the username and password
            If ($SmtpPwd)
            {
                $SmtpPwdEncrypt = Get-Content $SmtpPwd | ConvertTo-SecureString
                $SmtpCreds = New-Object System.Management.Automation.PSCredential -ArgumentList ($SmtpUser, $SmtpPwdEncrypt)

                ## If ssl was configured, send the email with ssl
                If ($UseSsl)
                {
                    Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer -UseSsl -Credential $SmtpCreds
                }

                ## If ssl wasn't configured, send the email without ssl
                Else
                {
                    Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer -Credential $SmtpCreds
                }
            }

            ## If an email username and password were not configured, send the email without authentication
            Else
            {
                Send-MailMessage -To $MailTo -From $MailFrom -Subject $MailSubject -Body $MailBody -SmtpServer $SmtpServer
            }
        }
    }
}

## End
