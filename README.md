# Remove-NspUsers.ps1

Remove users from NoSpamProxy users database that do not exist in Active Directory

## Description

This script deletes user from the NoSpamProxy database table [Usermanagement].[User] table that have not been removed by the Active Directory synchronization job.

## Parameters

No inputs required, however you should modify the Settings.xml file to suit your environment.

### Delete

Switch to finally DELETE users that exist in NoSpamProxy user table only. Without using this switch, found users information will be written to the log file only.

### Detailed

Switch to log existing Active Directory users as well

### SqlServerInstance

SQL Server instance hosting NoSpamProxyAddressSynchronization database 

## Examples

``` PowerShell
.\Remove-NspUsers.ps1
```

Check for Active Directory existance of all users stored in NoSpamProxy database. Do NOT delete any users from the database.

``` PowerShell
.\Remove-NspUsers.ps1 -Delete -SqlServerInstance MYNSPSERVER\SQLEXPRESS
```

Delete users from NoSpamProxy database hosted on SQL instance MYNSPSERVER\SQLEXPRESS that do NOT exist in Active Directory.

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)