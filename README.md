# Auto Create Exchange Mailboxes

PowerShell script to create Exchange Mailboxes for users in an OU structure.

For full instructions and documentation, [visit my blog post](https://gal.vin/2017/06/07/powershell-create-mailboxes)

Please consider donating to support my work:

* You can support me on a monthly basis [using Patreon.](https://www.patreon.com/mikegalvin)
* You can support me with a one-time payment [using PayPal](https://www.paypal.me/digressive) or by [using Kofi.](https://ko-fi.com/mikegalvin)

Auto Create Exchange Mailboxes can also be downloaded from:

* [The Microsoft TechNet Gallery](https://gallery.technet.microsoft.com/scriptcenter/Create-Exchange-Mailboxes-3b54f12c?redir=0)
* [The PowerShell Gallery](https://www.powershellgallery.com/packages/Create-Mailboxes)

Tweet me if you have questions: [@mikegalvin_](https://twitter.com/mikegalvin_)

-Mike

## Features and Requirements

* The script will run the WSUS server cleanup process, which will delete obsolete updates, as well as declining expired and superseded updates.
* The script can optionally create a log file and e-mail the log file to an address of your choice.
* The script can be run locally on a WSUS server, or on a remote sever.
* The script requires that the WSUS management tools be installed.
* The script has been tested on Windows 10 and Windows Server 2016.

### Generating A Password File

The password used for SMTP server authentication must be in an encrypted text file. To generate the password file, run the following command in PowerShell, on the computer that is going to run the script and logged in with the user that will be running the script. When you run the command you will be prompted for a username and password. Enter the username and password you want to use to authenticate to your SMTP server.

Please note: This is only required if you need to authenticate to the SMTP server when send the log via e-mail.

``` powershell
$creds = Get-Credential
$creds.Password | ConvertFrom-SecureString | Set-Content c:\scripts\ps-script-pwd.txt
```

After running the commands, you will have a text file containing the encrypted password. When configuring the -Pwd switch enter the path and file name of this file.

### Configuration

Here’s a list of all the command line switches and example configurations.

``` txt
-OU
```

The AD Organisational Unit (including child OUs) that contains the users to create Exchange Mailboxes for.

``` txt
-Datab
```

The Exchange database to create the mailboxes in. If you do not configure a Database, the smallest database will be used.

``` txt
-RP
```

The retention policy that should be applied to the users.

``` txt
-Compat
```

Use this switch if you are using Exchange 2010.

``` txt
-L
```

The path to output the log file to. The file name will be "Create-Mailboxes.log"

``` txt
-Subject
```

The email subject that the email should have. Encapulate with single or double quotes.

``` txt
-SendTo
```

The e-mail address the log should be sent to.

``` txt
-From
```

The from address the log should be sent from.

``` txt
-Smtp
```

The DNS name or IP address of the SMTP server.

``` txt
-User
```

The user account to connect to the SMTP server.

``` txt
-Pwd
```

The txt file containing the encrypted password for the user account.

``` txt
-UseSsl
```

Connect to the SMTP server using SSL.

### Example

``` txt
Create-Mailboxes.ps1 -Ou "OU=NewUsers,OU=Dept,DC=contoso,DC=com" -Datab "Mail DB 2" -Rp "1-Month-Deleted-Items" -L C:\scripts\logs -Subject 'Server: Created Mailboxes' -Sendto me@contoso.com -From Exch01@contoso.com -Smtp smtp.live.com -User Exch01@contoso.com -Pwd P@ssw0rd -UseSsl
```

This will create mailboxes for users that do not already have one in the OU NewUsers and all child OUs. It will create the mailbox using Mail DB 2 and apply the retention policy "1-Month-Deleted-Items". If you do not configure a database, the smallest database will be used. A log will be output to C:\scripts\logs and e-mailed with a custom subject line, using a secure connection.
