<#
    .SYNOPSIS
    Remove users from NoSpamProxyAddressSynchronization database that have been deleted from Active Directory
   
    Author: Thomas Stensitzki
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.0, 2017-08-02

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    http://scripts.Granikos.eu
	
    .DESCRIPTION
	
    This script deletes user from the NoSpamProxy database table [Usermanagement].[User]
    table that have not been removed by the Active Directory synchronization job. 
    

    .NOTES 
    Requirements 
    - Windows Server 2012 R2 or Windows Server 2016
    - Utilites global function library found here: http://scripts.granikos.eu
    - NoSpamProxy PowerShell module, script requires to run on a server having NoSpamProxy installed
    - ActiveDirectory PowerShell Module (Install-WindowsFeature RSAT-AD-PowerShell)

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
	
    .PARAMETER Delete
    Switch to finally DELETE users that exist in NoSpamProxy user table only. Without using this switch, found users information will be written to the log file only.

    .PARAMETER Detailed
    Switch to log existing Active Directory users as well

    .PARAMETER SqlServerInstance
    SQL Server instance hosting NoSpamProxyAddressSynchronization database 
   
    .EXAMPLE
    Check for Active Directory existance of all users stored in NoSpamProxy database. Do NOT delete any users from the database.
    .\Remove-NspUsers.ps1 

    .EXAMPLE
    Delete users from NoSpamProxy database hosted on SQL instance MYNSPSERVER\SQLEXPRESS that do NOT exist in Active Directory.
    .\Remove-NspUsers.ps1 -Delete -SqlServerInstance MYNSPSERVER\SQLEXPRESS

#>

[CmdletBinding()]
Param(
  [switch]$Delete,
  [switch]$Detailed,
  [string]$SqlServerInstance = 'MYNSPSERVER\SQLEXPRESS'
)

# import modules
Import-Module -Name ActiveDirectory
Import-Module -Name NoSpamProxy
# Import GlobalFunctions
if($null -ne (Get-Module -Name GlobalFunctions -ListAvailable).Version) {
  Import-Module -Name GlobalFunctions
}
else {
  Write-Warning -Message 'Unable to load GlobalFunctions PowerShell module.'
  Write-Warning -Message 'Open an administrative PowerShell session and run Import-Module GlobalFunctions'
  Write-Warning -Message 'Please check http://bit.ly/GlobalFunctions for further instructions'
  exit
}
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$ScriptName = $MyInvocation.MyCommand.Name
$logger = New-Logger -ScriptRoot $ScriptDir -ScriptName $ScriptName -LogFileRetention 30
$logger.Write(('Script started | DELETE = {0}' -f ($Delete)))

function Get-DomainController {
  $SiteName = (Get-ADDomainController -Discover).Site
  $GC = Get-ADDomainController -Discover -Service ADWS -SiteName $SiteName
  if ($GC -eq $null) {
    $GC = Get-ADDomainController -Discover -Service ADWS -NextClosestSite
  }
  $LocalGC = ('{0}:3268' -f $GC.HostName)

  return $LocalGC
}


# Fetch all users from NoSpamProxy database
Write-Host 'Fetching all user object from NoSpamProxy database...'

$AllNspUsers = Get-NspUser
$NspUserCount = ($AllNspUsers | Measure-Object).Count

Write-Verbose -Message ('Fetched {0} user objects from NoSpamProxy database' -f ($NspUserCount))
$logger.Write(('Fetched {0} user objects from NoSpamProxy database' -f ($NspUserCount)))

$UserCount = 1
$Existing = 0
$Missing = 0

$DomainController = Get-DomainController

Write-Host $DomainController

#break

# check all fetched users for existence in Active Directory
# we will check for each user email address currently stored in the NoSpamProxy database
foreach($NspUser in $AllNspUsers) {
    
  # write som nice progress bar
  Write-Progress -Activity ('Working on object ({1}/{2}) | [{0}] ' -f $NspUser.DisplayName, $UserCount, $NspUserCount) -Status 'Checking NoSpamProxy Users' -PercentComplete(($UserCount/$NspUserCount)*100)

  $found = $false
    
  # loop through each email address
  foreach($MailAddress in $NspUser.MailAddresses) {
        
    $Mail = $MailAddress.ToString()
    $ProxyAddress = ('smtp:{0}' -f ($MailAddress))

    # fetch AD object
    $AdUser = Get-ADObject -Properties mail, proxyAddresses, DisplayName -Filter {(mail -eq $Mail) -or (proxyAddresses -eq $ProxyAddress)} -Server $DomainController -ErrorAction SilentlyContinue

    if($AdUser -ne $null) {
      #Write-Host "  $($AdUser.mail)"
      if ($Detailed) {
        # log existing user AND email address
        $logger.Write(('AD exist  : {0} ({1})' -f $AdUser.DisplayName, $Mail))
      }

      $found = $true
    }
    else {
      # log missing users AND email address
      $logger.Write(('AD MISSING: {0} ({1})' -f $NspUser.DisplayName, $Mail))
    }
  }

  # increment user count
  $UserCount++

  if($found) {
    # User found in AD, so nothing to do. Just increment for statistics
    $Existing++
  }
  else {
    # Oops, user NOT found in AD, so we need to remove the user from NoSpamproxy database and increment for statistics
    $Missing++

    if($Delete) {
      # we will finally delete the user object identified by NoSpamProxy user id
      $logger.Write(('DELETE    : {0} | Id: {1}' -f $NspUser.DisplayName, $NspUser.Id))
      $cmd = "DELETE FROM [NoSpamProxyAddressSynchronization].[Usermanagement].[User] WHERE Id = '$($NspUser.Id)'"
            
      # Debug only
      # $Invoke = "Invoke-Sqlcmd -Query '$($cmd)' -ServerInstance $($SqlServerInstance)" 
      # $logger.Write(('INVOKE    : {0} ' -f $Invoke))

      Invoke-Sqlcmd -Query $cmd -ServerInstance $SqlServerInstance
    }
  }
}

# Write stats 
Write-Verbose -Message "Existing: $($Existing) | Missing: $($Missing)"
$logger.Write(('Existing: {0} | Missing: {1}' -f ($Existing), ($Missing)))

# Done
$logger.Write('Script finished')